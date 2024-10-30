# Server Data Verifier

## Descrição
Este projeto consiste em um script Python que analisa uma planilha Excel contendo dados de servidores e verifica se as informações foram criadas, modificadas ou excluídas a partir de uma data específica (13/08/2024). O script utiliza a biblioteca `openpyxl` para manipular a planilha e aplicar formatação de cores para facilitar a identificação das alterações.

## Funcionalidades
- **Destacar Linhas**: Altera a cor das linhas com base nas seguintes regras:
  - Verde: Dados criados a partir de 13/08/2024.
  - Amarelo: Dados modificados após 13/08/2024, mas criados antes.
  - Vermelho: Dados excluídos após 13/08/2024.

## Requisitos
- Python 3.x
- Bibliotecas: `openpyxl`
