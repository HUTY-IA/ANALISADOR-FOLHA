# Validação de Folha de Pagamento

Ferramenta interna — Huty Contabilidade Ltda.

Cruza os arquivos Excel recebidos dos clientes contra o extrato mensal do sistema Domínio.

## Instalação

```bash
pip install pdfplumber pandas openpyxl xlrd
```

## Uso

```bash
# Opção 1 — pasta com todos os arquivos (auto-detecção)
python validador.py --pasta ./06-2026

# Opção 2 — arquivos individuais
python validador.py \
  --extrato      Aponatmentos__-_Sistema.pdf \
  --apontamentos APONTAMENTOS_06-2026.xls \
  --emprestimos  EMPRESTIMOS_06-2026.xlsx \
  --coop         DESCONTOS_DE_COOPARTICIPAC_A_O_06-2026.xls
```

## O que é validado

| Módulo | Rubricas / Verificações |
|---|---|
| Apontamentos | 227, 230, 231, 224, 8111, 202, 203, 204, 48, HE 50%/100% |
| Empréstimos | valor da parcela por contrato, contratos ausentes/extras |
| Coparticipação | rubrica 227 e 203 por funcionário |

## Saída

Relatório `Validacao_Folha_MM-AAAA.xlsx` com abas: Resumo, Apontamentos, Empréstimos, Coparticipação.
