# Shipsmart Fluxo de Caixa API

API que recebe dois CSVs e gera o fluxo de caixa em .xlsx.

## Endpoint

POST /gerar-fluxo
- csv_pagar: arquivo CSV de contas a pagar
- csv_receber: arquivo CSV de contas a receber

Retorna: arquivo .xlsx
