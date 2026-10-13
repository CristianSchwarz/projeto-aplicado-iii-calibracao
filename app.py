from pathlib import Path
from uuid import uuid4
from datetime import datetime
import io
import json
import zipfile

import pandas as pd
from flask import Flask, render_template, request, send_file, abort
from werkzeug.utils import secure_filename

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle


app = Flask(__name__)

# Diretórios do projeto
BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)

# Configurações de upload
ALLOWED_EXTENSIONS = {".xlsx", ".xlsm"}
MAX_FILE_SIZE = 10 * 1024 * 1024
app.config["MAX_CONTENT_LENGTH"] = MAX_FILE_SIZE

COLUNAS_IMPORTACAO = {
    "dimension", "nominal", "+tol", "-tol", "meas 1", "meas 2"
}
COLUNAS_RELATORIO = {
    "characteristic", "actual", "nominal", "upper tol", "lower tol"
}


def extensao_permitida(nome):
    """Verifica se a extensão do arquivo é permitida."""
    return Path(nome).suffix.lower() in ALLOWED_EXTENSIONS


def normalizar_excel_strict(caminho):
    """
    Normaliza os namespaces Strict OOXML para leitura pelo openpyxl.
    O arquivo original não é alterado; o conteúdo convertido fica em memória.
    """
    buffer = io.BytesIO()

    substituicoes = {
        "http://purl.oclc.org/ooxml/spreadsheetml/main":
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "http://purl.oclc.org/ooxml/officeDocument/relationships":
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "http://purl.oclc.org/ooxml/package/relationships":
            "http://schemas.openxmlformats.org/package/2006/relationships",
    }

    with zipfile.ZipFile(caminho, "r") as origem:
        with zipfile.ZipFile(
            buffer, "w", compression=zipfile.ZIP_DEFLATED
        ) as destino:
            for item in origem.infolist():
                conteudo = origem.read(item.filename)

                if item.filename.endswith((".xml", ".rels")):
                    texto = conteudo.decode("utf-8")
                    for antigo, novo in substituicoes.items():
                        texto = texto.replace(antigo, novo)
                    conteudo = texto.encode("utf-8")

                destino.writestr(item, conteudo)

    buffer.seek(0)
    return buffer


def abrir_excel_compativel(caminho):
    """
    Primeiro tenta abrir normalmente. Se falhar, tenta normalizar o XLSX
    Strict OOXML e abre a versão temporária em memória.
    """
    try:
        return pd.ExcelFile(caminho, engine="openpyxl")
    except Exception:
        buffer = normalizar_excel_strict(caminho)
        try:
            excel = pd.ExcelFile(buffer, engine="openpyxl")
            # Mantém a referência ao buffer enquanto o ExcelFile estiver aberto.
            excel._strict_buffer = buffer
            return excel
        except Exception as exc:
            buffer.close()
            raise ValueError(
                "Não foi possível ler o Excel. O arquivo pode estar "
                "corrompido ou usar um formato não compatível."
            ) from exc


def localizar_cabecalho(planilha, colunas_necessarias):
    """Procura uma linha que contenha todas as colunas necessárias."""
    necessarias = {str(c).strip().casefold() for c in colunas_necessarias}

    for indice, linha in planilha.iterrows():
        valores = {
            str(valor).strip().casefold()
            for valor in linha
            if pd.notna(valor)
        }
        if necessarias.issubset(valores):
            return indice

    return None


def limpar_registros(df, coluna_caracteristica):
    """Remove linhas sem nome de característica."""
    df = df.dropna(subset=[coluna_caracteristica]).copy()
    df = df[
        df[coluna_caracteristica].astype(str).str.strip().ne("")
    ].copy()
    return df


def calcular_lea(tol_superior, tol_inferior):
    """
    REGRA PROVISÓRIA:
    usa o maior valor absoluto entre +TOL e -TOL como LEA.
    Confirmar a regra oficial com o professor/critério metrológico.
    """
    return max(abs(float(tol_superior)), abs(float(tol_inferior)))


def montar_resumo(df):
    resumo = {
        "total": len(df),
        "dentro": int(
            (df["Classificação"] == "DENTRO_DA_TOLERANCIA").sum()
        ),
        "fora": int(
            (df["Classificação"] == "FORA_DA_TOLERANCIA").sum()
        ),
        "invalidos": int(
            (df["Classificação"] == "DADOS_INVALIDOS").sum()
        ),
    }

    if resumo["invalidos"] > 0:
        parecer = "PENDENTE DE ANÁLISE"
    elif resumo["fora"] > 0:
        parecer = "REPROVADO"
    else:
        parecer = "APROVADO"

    return parecer, resumo


def processar_importacao(df):
    """
    Processa a aba Importação com duas medições.

    A validação segue a Aula 02: compara a média das leituras com os
    limites calculados a partir do nominal e dos módulos das tolerâncias.
    A margem 1e-7 é a tolerância numérica ilustrada na apresentação.
    Isso não substitui a confirmação do critério metrológico oficial.
    """
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    mapa = {str(c).strip().casefold(): c for c in df.columns}

    faltantes = sorted(COLUNAS_IMPORTACAO - set(mapa))
    if faltantes:
        raise ValueError(
            "A aba Importação não contém as colunas necessárias: "
            + ", ".join(faltantes)
        )

    dimension = mapa["dimension"]
    nominal = mapa["nominal"]
    tol_sup = mapa["+tol"]
    tol_inf = mapa["-tol"]
    meas1 = mapa["meas 1"]
    meas2 = mapa["meas 2"]

    df = limpar_registros(df, dimension)

    for coluna in [nominal, tol_sup, tol_inf, meas1, meas2]:
        df[coluna] = pd.to_numeric(df[coluna], errors="coerce")

    campos_validos = df[
        [nominal, tol_sup, tol_inf, meas1, meas2]
    ].notna().all(axis=1)

    df["Média das leituras"] = (df[meas1] + df[meas2]) / 2
    df["Erro sistemático"] = df["Média das leituras"] - df[nominal]
    df["Erro absoluto"] = df["Erro sistemático"].abs()

    # Limites conforme a Aula 02: nominal - |Tol.-| e nominal + |Tol.+|.
    df["Limite inferior"] = df[nominal] - df[tol_inf].abs()
    df["Limite superior"] = df[nominal] + df[tol_sup].abs()

    # Mantém a coluna LEA para rastreabilidade, mas não a usa como critério
    # de classificação nesta regra baseada nos limites inferior/superior.
    df["LEA"] = df[[tol_sup, tol_inf]].abs().max(axis=1)

    tolerancias_validas = (df[tol_sup] >= 0) & (df[tol_inf] <= 0)
    validos = campos_validos & tolerancias_validas
    margem = 1e-7

    df["Classificação"] = "DADOS_INVALIDOS"
    dentro = (
        validos
        & (df["Média das leituras"] >= df["Limite inferior"] - margem)
        & (df["Média das leituras"] <= df["Limite superior"] + margem)
    )
    df.loc[dentro, "Classificação"] = "DENTRO_DA_TOLERANCIA"
    df.loc[validos & ~dentro, "Classificação"] = "FORA_DA_TOLERANCIA"

    parecer, resumo = montar_resumo(df)
    return df, parecer, resumo


def processar_relatorio_mmc(df):
    """
    Mantém a compatibilidade com a aba Report que contém valores consolidados.
    Não calcula média porque o relatório já apresenta Actual como resultado.
    """
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    mapa = {str(c).strip().casefold(): c for c in df.columns}

    faltantes = sorted(COLUNAS_RELATORIO - set(mapa))
    if faltantes:
        raise ValueError(
            "O cabeçalho Report não contém as colunas necessárias: "
            + ", ".join(faltantes)
        )

    caracteristica = mapa["characteristic"]
    actual = mapa["actual"]
    nominal = mapa["nominal"]
    tol_sup = mapa["upper tol"]
    tol_inf = mapa["lower tol"]

    df = limpar_registros(df, caracteristica)

    for coluna in [actual, nominal, tol_sup, tol_inf]:
        df[coluna] = pd.to_numeric(df[coluna], errors="coerce")

    campos_validos = df[[actual, nominal, tol_sup, tol_inf]].notna().all(axis=1)
    tolerancias_validas = (df[tol_sup] >= 0) & (df[tol_inf] <= 0)
    validos = campos_validos & tolerancias_validas

    df["Erro calculado"] = df[actual] - df[nominal]
    df["Erro absoluto"] = df["Erro calculado"].abs()
    df["Limite inferior"] = df[nominal] - df[tol_inf].abs()
    df["Limite superior"] = df[nominal] + df[tol_sup].abs()
    df["Classificação"] = "DADOS_INVALIDOS"
    margem = 1e-7

    dentro = (
        validos
        & (df[actual] >= df["Limite inferior"] - margem)
        & (df[actual] <= df["Limite superior"] + margem)
    )
    df.loc[dentro, "Classificação"] = "DENTRO_DA_TOLERANCIA"
    df.loc[validos & ~dentro, "Classificação"] = "FORA_DA_TOLERANCIA"

    parecer, resumo = montar_resumo(df)
    return df, parecer, resumo


def salvar_resultado(df, nome, aba, tipo, parecer, resumo):
    job_id = uuid4().hex
    caminho_resultado = UPLOAD_DIR / f"{job_id}.json"

    registros = json.loads(
        df.to_json(orient="records", force_ascii=False)
    )

    dados = {
        "arquivo_origem": nome,
        "aba": aba,
        "tipo": tipo,
        "data_geracao": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "parecer": parecer,
        "resumo": resumo,
        "colunas": [str(c) for c in df.columns],
        "registros": registros,
    }

    with caminho_resultado.open("w", encoding="utf-8") as arquivo_json:
        json.dump(
            dados, arquivo_json, ensure_ascii=False, allow_nan=False
        )

    return job_id


@app.route("/", methods=["GET", "POST"])
def index():
    erro = None
    arquivo = None
    abas = []
    aba_selecionada = None
    colunas = []
    linhas = []
    total_linhas = 0
    parecer = None
    resumo = None
    job_id = None
    tipo_processamento = None

    if request.method == "POST":
        enviado = request.files.get("arquivo")

        if not enviado or not enviado.filename:
            erro = "Selecione uma planilha para continuar."
        elif not extensao_permitida(enviado.filename):
            erro = "Envie um arquivo .xlsx ou .xlsm."
        else:
            nome = secure_filename(enviado.filename)

            if not nome:
                erro = "O nome do arquivo é inválido."
            else:
                # Nome único evita que dois uploads com o mesmo nome se sobrescrevam.
                caminho = UPLOAD_DIR / f"{uuid4().hex}_{nome}"

                try:
                    enviado.save(caminho)

                    with abrir_excel_compativel(caminho) as excel:
                        abas = excel.sheet_names

                        if not abas:
                            raise ValueError(
                                "Nenhuma aba foi encontrada no arquivo."
                            )

                        # Prioriza Importação; se não existir, usa Report; senão, primeira aba.
                        preferida = next(
                            (
                                aba for aba in abas
                                if aba.strip().casefold() == "importação"
                            ),
                            next(
                                (
                                    aba for aba in abas
                                    if aba.strip().casefold() == "report"
                                ),
                                abas[0],
                            ),
                        )

                        aba_solicitada = request.form.get("aba")
                        aba_selecionada = (
                            aba_solicitada
                            if aba_solicitada in abas
                            else preferida
                        )

                        bruto = pd.read_excel(
                            excel,
                            sheet_name=aba_selecionada,
                            header=None,
                        )

                        cabecalho_importacao = localizar_cabecalho(
                            bruto, COLUNAS_IMPORTACAO
                        )
                        cabecalho_report = localizar_cabecalho(
                            bruto, COLUNAS_RELATORIO
                        )

                        if cabecalho_importacao is not None:
                            df = pd.read_excel(
                                excel,
                                sheet_name=aba_selecionada,
                                header=cabecalho_importacao,
                            )
                            df.columns = [str(c).strip() for c in df.columns]
                            df, parecer, resumo = processar_importacao(df)
                            tipo_processamento = (
                                "Importação — média, erro sistemático e LEA"
                            )

                        elif cabecalho_report is not None:
                            df = pd.read_excel(
                                excel,
                                sheet_name=aba_selecionada,
                                header=cabecalho_report,
                            )
                            df.columns = [str(c).strip() for c in df.columns]
                            df, parecer, resumo = processar_relatorio_mmc(df)
                            tipo_processamento = (
                                "Report — valores consolidados da MMC"
                            )
                        else:
                            raise ValueError(
                                "Não encontrei o cabeçalho esperado nesta aba. "
                                "A aba Importação deve conter DIMENSION, NOMINAL, "
                                "+TOL, -TOL, MEAS 1 e MEAS 2; a aba Report deve "
                                "conter Characteristic, Actual, Nominal, Upper Tol "
                                "e Lower Tol."
                            )

                        total_linhas = len(df)
                        job_id = salvar_resultado(
                            df, nome, aba_selecionada, tipo_processamento,
                            parecer, resumo
                        )

                        # A tela mostra até 20 linhas; resumo/PDF usam todas.
                        visualizacao = df.head(20).copy()
                        visualizacao = visualizacao.astype(object).where(
                            pd.notna(visualizacao), None
                        )
                        colunas = [str(c) for c in visualizacao.columns]
                        linhas = visualizacao.values.tolist()
                        arquivo = nome

                except Exception as exc:
                    app.logger.exception("Falha ao importar planilha")
                    erro = (
                        "Não foi possível processar a planilha. "
                        f"Detalhe técnico: {exc}"
                    )
                finally:
                    try:
                        if caminho.exists():
                            caminho.unlink()
                    except Exception:
                        pass

    return render_template(
        "index.html",
        erro=erro,
        arquivo=arquivo,
        abas=abas,
        aba_selecionada=aba_selecionada,
        colunas=colunas,
        linhas=linhas,
        total_linhas=total_linhas,
        parecer=parecer,
        resumo=resumo,
        job_id=job_id,
        tipo_processamento=tipo_processamento,
    )


@app.route("/gerar-certificado", methods=["POST"])
def gerar_certificado():
    job_id = request.form.get("job_id", "")

    if len(job_id) != 32 or any(
        caractere not in "0123456789abcdef" for caractere in job_id
    ):
        abort(400, "Identificador de processamento inválido.")

    caminho_json = UPLOAD_DIR / f"{job_id}.json"
    if not caminho_json.is_file():
        abort(404, "Resultado não encontrado. Importe a planilha novamente.")

    try:
        with caminho_json.open("r", encoding="utf-8") as arquivo_json:
            dados = json.load(arquivo_json)

        caminho_pdf = UPLOAD_DIR / f"relatorio_{job_id}.pdf"

        estilos = getSampleStyleSheet()

        titulo = ParagraphStyle(
            "TituloCertificado",
            parent=estilos["Title"],
            alignment=TA_CENTER,
            fontName="Helvetica-Bold",
            fontSize=16,
            leading=20,
            textColor=colors.HexColor("#172B4D"),
            spaceAfter=8,
        )
        aviso = ParagraphStyle(
            "AvisoCertificado",
            parent=estilos["Normal"],
            alignment=TA_CENTER,
            fontName="Helvetica-Bold",
            fontSize=9,
            leading=12,
            textColor=colors.HexColor("#B00020"),
        )
        pequeno = ParagraphStyle(
            "TextoPequeno",
            parent=estilos["Normal"],
            fontName="Helvetica",
            fontSize=6.4,
            leading=7.4,
            textColor=colors.HexColor("#1F2937"),
            wordWrap="CJK",
        )
        estilo_cabecalho = ParagraphStyle(
            "CabecalhoTabela",
            parent=estilos["Normal"],
            fontName="Helvetica-Bold",
            fontSize=6.9,
            leading=8.0,
            textColor=colors.white,
            wordWrap="CJK",
        )

        documento = SimpleDocTemplate(
            str(caminho_pdf),
            pagesize=landscape(A4),
            rightMargin=10 * mm,
            leftMargin=10 * mm,
            topMargin=12 * mm,
            bottomMargin=15 * mm,
            title="Relatório de análise de calibração",
            author="Projeto Aplicado III - Schulz S.A.",
        )

        conteudo = [
            Paragraph("RELATÓRIO DE ANÁLISE DE CALIBRAÇÃO", titulo),
            Paragraph(
                "DOCUMENTO DE DEMONSTRAÇÃO — NÃO É UM CERTIFICADO METROLÓGICO OFICIAL",
                aviso,
            ),
            Spacer(1, 8),
            Paragraph(
                f"<b>Arquivo de origem:</b> {dados['arquivo_origem']}",
                estilos["Normal"],
            ),
            Paragraph(
                f"<b>Aba:</b> {dados['aba']} &nbsp;&nbsp; "
                f"<b>Tipo:</b> {dados['tipo']} &nbsp;&nbsp; "
                f"<b>Data de geração:</b> {dados['data_geracao']}",
                estilos["Normal"],
            ),
            Spacer(1, 8),
            Paragraph(
                f"Parecer calculado: {dados['parecer']}",
                estilos["Heading2"],
            ),
        ]

        resumo = dados["resumo"]
        conteudo.append(
            Paragraph(
                f"Total de pontos: {resumo['total']} &nbsp;&nbsp; "
                f"Dentro da tolerância: {resumo['dentro']} &nbsp;&nbsp; "
                f"Fora da tolerância: {resumo['fora']} &nbsp;&nbsp; "
                f"Dados inválidos: {resumo['invalidos']}",
                estilos["Normal"],
            )
        )
        conteudo.append(Spacer(1, 10))

        colunas_pdf = dados["colunas"]
        tabela_dados = [[
            Paragraph(
                str(coluna).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"),
                estilo_cabecalho,
            )
            for coluna in colunas_pdf
        ]]

        for registro in dados["registros"]:
            linha = []
            for coluna in colunas_pdf:
                valor = registro.get(coluna)

                if valor is None:
                    texto = "—"
                elif isinstance(valor, (int, float)) and not isinstance(valor, bool):
                    texto = f"{valor:.6f}"
                else:
                    texto = str(valor)

                # Deixa a classificação mais legível no PDF, sem alterar os dados salvos.
                if coluna.casefold() == "classificação":
                    texto = texto.replace("DENTRO_DA_TOLERANCIA", "DENTRO DA TOLERÂNCIA")
                    texto = texto.replace("FORA_DA_TOLERANCIA", "FORA DA TOLERÂNCIA")
                    texto = texto.replace("DADOS_INVALIDOS", "DADOS INVÁLIDOS")

                texto = (
                    texto.replace("&", "&amp;")
                    .replace("<", "&lt;")
                    .replace(">", "&gt;")
                )
                linha.append(Paragraph(texto, pequeno))

            tabela_dados.append(linha)

        largura_util = landscape(A4)[0] - 20 * mm
        quantidade_colunas = max(len(colunas_pdf), 1)

        # Distribui a largura conforme o conteúdo: identificação e classificação
        # recebem mais espaço; colunas numéricas ficam compactas.
        pesos = []
        for coluna in colunas_pdf:
            nome_coluna = coluna.casefold()
            if nome_coluna in {"dimension", "characteristic"}:
                pesos.append(2.15)
            elif nome_coluna in {"título", "titulo"}:
                pesos.append(0.8)
            elif nome_coluna == "classificação":
                pesos.append(1.75)
            elif nome_coluna in {
                "média das leituras", "erro sistemático", "erro calculado",
                "limite inferior", "limite superior", "erro absoluto",
            }:
                pesos.append(1.25)
            elif nome_coluna in {"nominal", "+tol", "-tol", "upper tol", "lower tol", "meas 1", "meas 2", "lea", "actual"}:
                pesos.append(0.9)
            else:
                pesos.append(1.0)

        soma_pesos = sum(pesos) or 1
        larguras = [largura_util * peso / soma_pesos for peso in pesos]

        tabela = Table(
            tabela_dados,
            colWidths=larguras,
            repeatRows=1,
            hAlign="LEFT",
        )
        tabela.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#243447")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#D1D5DB")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [
                colors.white, colors.HexColor("#F2F5F8")
            ]),
        ]))

        conteudo.append(tabela)
        conteudo.append(Spacer(1, 10))
        conteudo.append(
            Paragraph(
                "Documento demonstrativo do MVP. A classificação da aba Importação "
                "utiliza os limites inferior e superior calculados a partir do nominal "
                "e das tolerâncias, conforme a abordagem apresentada na Aula 02. "
                "O campo LEA exibido permanece provisório e o critério deve ser "
                "confirmado antes de qualquer uso metrológico operacional.",
                pequeno,
            )
        )

        def rodape(canvas, doc):
            canvas.saveState()
            canvas.setFont("Helvetica", 8)
            canvas.setFillColor(colors.grey)
            canvas.drawString(
                10 * mm,
                7 * mm,
                "Projeto Aplicado III — Schulz S.A. — Documento de demonstração",
            )
            canvas.drawRightString(
                landscape(A4)[0] - 10 * mm,
                7 * mm,
                f"Página {doc.page}",
            )
            canvas.restoreState()

        documento.build(
            conteudo,
            onFirstPage=rodape,
            onLaterPages=rodape,
        )

        return send_file(
            caminho_pdf,
            as_attachment=True,
            download_name=f"relatorio_calibracao_{job_id[:8]}.pdf",
            mimetype="application/pdf",
        )

    except Exception:
        app.logger.exception("Erro ao gerar relatório PDF")
        return (
            "Não foi possível gerar o PDF. Confira o terminal do VS Code.",
            500,
        )


@app.errorhandler(413)
def arquivo_muito_grande(_erro):
    return (
        "O arquivo ultrapassa o limite de 10 MB. Selecione uma planilha menor.",
        413,
    )


if __name__ == "__main__":
    app.run(debug=True)
