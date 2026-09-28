# Demonstração — Conversão de TXT de notas para importação na Domínio

> Projeto de portfólio de **Victória Pedrosa**. **Demonstração** de conversão de TXT de notas para importação na Domínio — versão com dados fictícios (nomes, CNPJs, e-mails e IDs internos substituídos).

## Problema de negócio
Arquivos TXT de notas de saída, entrada e serviços tomados precisavam ser convertidos para o leiaute de importação da Domínio.

## Antes x depois
| | Antes | Depois |
|---|---|---|
| Como é feito | Conversão manual e digitação de dados. | Scripts leem os TXT (com apoio do Gemini) e geram os arquivos no leiaute de importação. |

## Ganho
- Importação na Domínio sem redigitação.

## Tecnologias
APIs REST, Gemini API, Google Apps Script, Google Drive, Google Sheets

## Arquivos
- `LerTxtEntradas.gs`
- `LerTxtSaidas.gs`
- `LerTxtServicosTomados.gs`

## Como usar
Crie um projeto no Google Apps Script, copie os arquivos `.gs`/`.html` e configure as Propriedades do script indicadas no código.

## Autora
Victória Pedrosa — Product Owner do Time de IA, automação de processos contábeis e fiscais.
