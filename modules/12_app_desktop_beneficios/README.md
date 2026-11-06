# 🖥️ Módulo 12: App Desktop de Gestão e Auditoria de Benefícios (CustomTkinter)

![Status](https://img.shields.io/badge/Status-Concluído-brightgreen)
![UI](https://img.shields.io/badge/GUI-CustomTkinter-blue)

## 📌 Visão Geral & Pipeline Modular
Pipeline em Python e interface CustomTkinter para auditar faturamento de operadoras de saúde/odonto contra movimentações do RH.

## Estrutura

| Arquivo | Função |
|---|---|
| `config.py` | Mapeamento de colunas das 7 operadoras + Colaboradores + Demitidos. **Mudou layout? Mexa só aqui.** |
| `utils.py` | Funções utilitárias: limpar/validar/formatar CPF, normalizar nome, parse de data |
| `leitores.py` | Lê cada planilha e devolve DataFrame com colunas padronizadas em snake_case |
| `motor_auditoria.py` | Merge em cascata: CPF → matrícula → nome. Gera o consolidado |
| `exportador.py` | Gera XLSX final formatado (CPF como texto, datas DD/MM/AAAA, aba Resumo) |
| `main.py` | Ponto de entrada. Descobre arquivos na pasta, orquestra tudo |

## Dependências

```bash
pip install pandas openpyxl
```

## Uso

1. Coloque todas as planilhas numa pasta (qualquer estrutura), garantindo que o nome contém a operadora:
   - `Listagem de Ativos - Bradesco Odonto - Maio_2026.xlsx`
   - `Movimentações Fopag - HAPVIDA_Maio_2026.xlsx`
   - `COLA_Colaboradores_CSV - 18052026.xlsx`
   - `Demitidos_01.05.2026_ate_18.05.2026.CSV`

2. Rode:
   ```bash
   python main.py "C:\auditoria\entrada"
   ```
   Ou especificando saída:
   ```bash
   python main.py "C:\auditoria\entrada" "C:\auditoria\saida.xlsx"
   ```

3. O log fica em `<pasta_entrada>/logs/auditoria_YYYYMMDD_HHMMSS.log`.

## Saída

XLSX com 2 abas:
- **Consolidado**: 38 colunas com o cruzamento completo + flags de auditoria
- **Resumo**: contagem por operadora, por status, e por flag crítica

### Flags de auditoria

| Flag | Significado |
|---|---|
| `FLAG_COBRANCA_INDEVIDA` | Cobrado ativo sem movimentação no Fopag |
| `FLAG_DEMITIDO_AINDA_COBRADO` | Foi demitido mas continua na operadora |
| `FLAG_SEM_CADASTRO_RH` | Cobrado mas nem está no cadastro do RH |
| `FLAG_DEPENDENTE_SEM_TITULAR` | Dependente cobrado sem titular identificável |

## Estratégia de merge (cascata)

1. **CPF** (chave primária — mais confiável)
2. O que sobrar → tenta por **matrícula**
3. O que ainda sobrar → tenta por **nome normalizado** (sem acento, maiúsculo)
4. O resto entra no relatório com `ORIGEM_MATCH = SEM_MATCH` — **nada é descartado em silêncio**

Cada etapa loga quantos casaram e quantos sobraram. Se houver vazamento (entrada ≠ saída), aparece `ERROR: VAZAMENTO DE REGISTROS!` no log.
