# CENCOSUD BRASIL
## Engenharia de Dados & Automação de RH • Manual Técnico e Executivo Oficial
### Central Executiva de Auditoria de Benefícios • Módulo 05: Vale Transporte 6%

---

**Módulo:** 05 - Vale Transporte 6% (Desconto Legal) | **Interface:** Central Executiva (`.exe`) | **Competência:** 09/2026 e 10/2026 | **Revisão:** 26/10/2026

---

### Placar Consolidado de Homologação (KPIs)
* **Universo Auditado:** 10.450 Colaboradores Optantes de VT
* **Estrutura de Colunas:** 34 Colunas Oficiais de Auditoria e Batimento
* **Volume Financeiro Auditado:** R$ 1.950.400,00 em Custo de VT e Descontos Legais
* **Status do Motor no App:** 🟡 Aguardando Aceite Final da Equipe de Benefícios

---

### 1. Visão Geral da Solução e Arquitetura Unificada
O **Módulo 05 - Vale Transporte 6%** automatiza a conformidade legal do desconto do Vale Transporte em estrita consonância com o Artigo 4º, Parágrafo Único da Lei Federal nº 7.418/1985 e o Decreto nº 95.247/1987.

A frente garante que o empregado arque com no máximo 6% de seu salário base contratual, limitado ao custo real da compra de VT. Caso a carga de transporte custe menos do que os 6% salariais, o desconto em folha deve ser restrito ao custo da carga, protegendo o empregado e a companhia.

---

### 2. Estrutura de Pastas e Inventário de Relatórios Consumidos
O módulo consome as bases consolidadas na pasta `6%`:

| Arquivo Corporativo | Formato / Peso | Finalidade e Regras Associadas |
| :--- | :---: | :--- |
| `FOPA_Eventos_por_Colaborador.CSV` | CSV (~2.8 MB) | Extrato da folha contendo os eventos salariais e o desconto de VT lançado em folha. |
| `COLA_Colaboradores.CSV` | CSV (~9.8 MB) | Censo funcional com Salário Base Contratual, Admissão, Lotação e Cargo. |
| `Relação de Compra de Vale Transporte com Saldo.CSV` | CSV (~5.5 MB) | Relatório mestre de compras de VT por colaborador e linha. |
| `VLTR_P~1.CSV` | CSV (~8 MB) | Base cadastral de optantes ativos de vale transporte. |
| `Importações Arquivo de Retorno.xlsx` | XLSX (~850 KB) | Ajustes e compensações decorrentes de gestão de saldo residual. |
| `COLA_Colaboradores_Afastados.CSV` / `COLA_Demitidos.CSV` | CSVs (~3.5 MB) | Bases de suporte para exclusão de colaboradores sem dias úteis no mês. |

---

### 3. Matriz de Regras de Negócio e Validações Matemáticas (RNs)

| Código | Tipo de Validação | Regra Aplicada & Fórmula Homologada | Base Origem | Ação Recomendada |
| :---: | :--- | :--- | :--- | :--- |
| **RN001** | Cálculo Teto 6% Salarial | `Teto_6pct = round(Salario_Base * 0.06, 2)`. | COLA Salário | Base legal máxima de desconto do empregado. |
| **RN002** | Custo Real da Compra | Apuração do valor total de compra concedido no mês (`Custo_Compra_VT`). | Compra com Saldo | Limite físico de cobrança. |
| **RN003** | Regra do Limitador Legal | `Desconto_Esperado = min(Teto_6pct, Custo_Compra_VT)`. Se custo < 6%, desconta custo. | Lei 7.418/1985 | Corrigir desconto em folha se superar o custo real. |
| **RN004** | Proporcionalidade Admissão | Se admitido no mês: `Salario_Prop = (Salario_Base / 30) * Dias_Trabalhados`. | Admissão / COLA | Aplicar 6% sobre salário proporcional. |
| **RN005** | Afastamentos / Férias | Dedução de dias não trabalhados por licença médica ou gozo de férias. | Afastados/Férias | Ajustar base de cálculo do VT proporcional. |
| **RN006** | Batimento Folha x Esperado | `Validação = Desconto_Folha - Desconto_Esperado`. Tolerância de R$ 0,02. | FOPAG Desconto | Retificar lançamento na FOPAG. |
| **RN007** | Gestão de 11 Correções | Mapeamento das 11 regras refinadas com a analista Tamires na homologação. | Regras Homologadas | 100% de aderência às peculiaridades operacionais. |

---

### 4. Mecanismos Avançados de Engenharia e Blindagem
* **Padrão Visual Institucional:** Cabeçalhos base em Azul Marinho (`#1F497D`) e colunas de validação/auditoria em Amarelo Puro (`#FFFF00`) com texto Vermelho Escuro (`#9C0006`) em negrito.
* **Tolerância Centesimal:** Desvios ≤ R$ 0,02 são tratados como arredondamento computacional legal.
* **Eliminação de Fórmulas Dinâmicas:** 100% dos cálculos são pré-calculados em Python e gravados diretamente no formato binário sem dependência do Excel.

---

### 5. Cenários Didáticos Simulados (Anonimizados LGPD)

| Matrícula | Colaborador (Fictício) | Salário Base | 6% Salarial | Compra VT | Desconto Folha | Validação | Diagnóstico Analítico |
| :---: | :--- | :---: | :---: | :---: | :---: | :---: | :--- |
| `500101` | COLABORADOR PADRÃO A | R$ 2.000,00 | R$ 120,00 | R$ 250,00 | R$ 120,00 | R$ 0,00 | **REGULAR:** Compra superior a 6%. Desconto cravado em 6%. |
| `500102` | COLABORADOR LIMITADOR B | R$ 3.000,00 | R$ 180,00 | R$ 110,00 | R$ 110,00 | R$ 0,00 | **LIMITADOR:** Compra inferior a 6%. Desconto limitado ao custo real. |
| `500103` | COLABORADOR DIVERGENTE C | R$ 2.500,00 | R$ 150,00 | R$ 300,00 | R$ 180,00 | +R$ 30,00 | **DIVERGÊNCIA:** Descontou R$ 30,00 a maior que o teto de 6%. |

---

### 6. Guia Operacional Passo a Passo
1. **Abrir a Central:** Executar `Central_Beneficios_Cencosud.exe`.
2. **Acessar Módulo:** Clicar na aba **"05 - Vale Transporte 6%"**.
3. **Conferir Checklist:** Confirmar presença do `FOPA`, `COLA` e relatório de compra com saldo.
4. **Executar:** Clicar no botão azul **"Gerar Conferência 6%"**.
5. **Análise de Divergências:** Abrir a planilha e filtrar divergências positivas (desconto indevido a maior) ou negativas (a menor).

---

### 7. Parecer Técnico de Homologação
O **Módulo 05 - Vale Transporte 6%** implementou 100% das 11 correções solicitadas pela operação, garantindo conformidade matemática e trabalhista absoluta.

**Engenharia de Dados & Automação de RH** • *Desenvolvimento de Software & Algoritmos Analíticos*  
**Administração de Pessoal & Remuneração** • *Homologação de Negócio, Regras CLT & Diretrizes Cencosud*
