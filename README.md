# Sistema de rateio para Faturas Vivo

## Objetivo
O sistema consiste em um programa em Python que realiza a leitura da fatura em PDF, localiza os números pertencentes ao plano contratado, cria uma planilha para o Excel e cruza as informações das linhas com outra planilha.
Isso permite que seja realizado o rateio da conta entre as filiais.

## Como usar
- Baixe e execute o programa [nesse link](https://github.com/BRGustavo/Rateio/releases/tag/v0.1).
- Selecione a planilha utilizada como base para comparação dos locais estão as linhas.
- Por fim, selecione o arquivo PDF com a fatura e deixe o programa trabalhar.

### Padrão tabela fonte.
É necessário que a planilha usada como base para extração dos números possua a seguinte estrutura.

| Nº  | Local | Agência |
| ------------- | ------------- | ------------- |
| 439XXXXXXXX | 000000 | AG .... |
