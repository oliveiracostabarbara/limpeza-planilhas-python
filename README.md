# Consolidador e limpador de planilhas (Python)

## Problema
Uma loja fictícia de produtos pet recebe pedidos por 3 canais (caixa, site e WhatsApp). Cada canal exporta num formato diferente: separadores, datas e valores variados, nomes e telefones sem padrão, linhas vazias e pedidos repetidos. Juntar e limpar isso à mão leva cerca de 2 horas (estimado) e deixa erros passarem.

## Solução
Script em Python (pandas) que lê as 3 planilhas, padroniza as colunas, limpa nomes, telefones, valores e datas, separa as linhas inválidas (com o motivo) e remove pedidos duplicados. Gera 1 Excel com as abas "Pedidos" e "Erros".

## Resultado
| | Antes (manual) | Depois (script) |
|---|---|---|
| Tempo | cerca de 2 horas (estimado) | cerca de 1 minuto |

- 276 linhas lidas nas 3 planilhas
- 14 linhas inválidas (5 sem valor, 3 sem data, 3 sem telefone, 3 telefone inválido), listadas na aba "Erros"
- 12 pedidos duplicados removidos
- 250 pedidos únicos, com receita total conferida de R$ 11.649,00

## Como rodar
1. Instale o Python 3 e as bibliotecas: `pip install pandas openpyxl`
2. Coloque as planilhas na pasta `dados/`
3. Rode: `python limpar-planilhas.py`
4. O resultado sai em `saida/pedidos_consolidados.xlsx`

## Regras de limpeza
- Cliente identificado pelo telefone (11 dígitos com DDD)
- Linha inválida: sem valor, sem data, sem telefone ou telefone incompleto
- Pedido duplicado: mesmo telefone, data, produto e valor

## Dados
100% fictícios, criados para este projeto.