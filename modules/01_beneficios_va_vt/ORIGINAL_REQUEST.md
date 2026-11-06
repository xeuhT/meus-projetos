# Original User Request

## 2026-10-06T14:15:22Z

# Teamwork Project Prompt

Desenvolver um script Python modular e em memória (Pandas + OpenPyXL) para automatizar a auditoria "04 - GOOD CARD" da Cencosud, gerando um relatório Excel estático e formatado com base em regras de negócio e cruzamento de dados de múltiplos CSVs de RH.

Working directory: `C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\VA e VT\04 - GOOD CARD`
Integrity mode: development
Requested team: Full team

## Requirements

### R1. Engenharia de Dados e Regras de Negócio (`data_loader.py` e `business_rules.py`)
- Ingestão dos relatórios originais em CSV (`COLA_Colaboradores.CSV`, `COLA_Colaboradores_Afastados.CSV`, `COLA_Demitidos.CSV`, `COLA_Ferias_Colaboradores.CSV`, glob em `FOPA_Eventos_por_*.csv`), tratando delimitadores, encoding (`latin1` ou apropriado) e metadados no topo dos arquivos (ex: pular 1 linha no arquivo de Férias e 3 linhas no FOPA). O FOPA tem separador `;`.
- Chave de Join universal estritamente por `Matrícula` (convertida e tratada como string/texto para evitar perda de zeros).
- Implementação em Pandas das 4 regras do Good Card:
  1. Cartões de Desligados: cancelamento para demitidos.
  2. Afastados > 15 dias: solicitação de bloqueio.
  3. Novos admitidos > 90 dias: elegibilidade (excluindo diretores e aprendizes).
  4. Desconto em Folha: Margem de 20% sobre o salário, apuração do evento 29621 (Saldo Devedor) e extração dos eventos 11010, 6280, 29613 e 29624.

### R2. Camada de Exportação Excel e Layout (`excel_exporter.py` e `main.py`)
- O script `main.py` deve orquestrar a execução chamando a ingestão, regras e exportação.
- O `excel_exporter.py` deve gerar um arquivo Excel final (`.xlsx`) sem fórmulas, fiel ao modelo "Conferência Good Card_Base.xlsx", contendo as abas:
  - "Orientações" (preservar layout original da aba base).
  - "Resumo da Auditoria" (dashboard com totalizadores/KPIs).
  - "Conferência" (38 colunas exatas do modelo, preenchimento dinâmico dos dados baseados na regra de negócio).
  - "Valores Importados" (11 colunas exatas conciliando evento 29613).
- Usabilidade do Excel final: Freeze panes restrito à Linha 1 ('A2'). Centralizar códigos, matrículas, datas, eventos e valores numéricos/monetários. Nomes e tratativas alinhados à esquerda. Destacar com fonte vermelha ou fundo vermelho (via OpenPyXL styles) as linhas ou células com divergência (por ex., na coluna Tratativa).

## Acceptance Criteria

### Execução e Qualidade do Código
- [ ] O script principal (`main.py`) executa do início ao fim sem lançar exceções.
- [ ] A lógica não salva arquivos CSV intermediários; processamento todo em memória.
- [ ] Tipagem correta: as chaves (`Matrícula`) devem dar merge perfeitamente.

### Qualidade do Output (Arquivo Excel gerado)
- [ ] O arquivo final é salvo em disco e pode ser lido sem corrompimentos.
- [ ] A aba "Conferência" e "Valores Importados" possuem as exatas colunas da base.
- [ ] Freeze Panes está em 'A2' e cores vermelhas são aplicadas nos casos divergentes.
