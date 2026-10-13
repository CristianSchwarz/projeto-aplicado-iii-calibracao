# Sistema de Calibração de Instrumentos

Projeto desenvolvido para o Projeto Aplicado III, com o objetivo de facilitar a importação e a análise de planilhas de calibração de instrumentos.

## Funcionalidades

- Importação de planilhas Excel.
- Processamento dos dados de calibração.
- Cálculo dos limites inferior e superior.
- Classificação dos resultados.
- Geração de relatório em PDF.

## Tecnologias utilizadas

- Python
- Flask
- HTML e CSS
- Pandas
- OpenPyXL
- ReportLab

## Como executar o projeto

1. Clone este repositório:
   ```bash
   git clone https://github.com/CristianSchwarz/projeto-aplicado-iii-calibracao.git
   ```

2. Entre na pasta do projeto:
   ```bash
   cd projeto-aplicado-iii-calibracao
   ```

3. Crie e ative um ambiente virtual:
   ```bash
   python -m venv venv
   ```

4. Instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```

5. Execute a aplicação:
   ```bash
   python app.py
   ```

6. Acesse o endereço local informado no terminal pelo Flask.

## Observação

Este projeto tem finalidade acadêmica e demonstrativa. Os resultados devem ser validados conforme os procedimentos metrológicos aplicáveis antes de qualquer utilização oficial.
