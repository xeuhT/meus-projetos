# CENCOSUD BRASIL
## Engenharia de Dados & Automação de RH • Manual Técnico e Executivo Oficial
### Central Executiva de Auditoria de Benefícios • Módulo 04: Bloqueio Good Card

---

**Módulo:** 04 - Bloqueio Good Card (Afastados e Demitidos) | **Interface:** Central Executiva (`.exe`) | **Competência:** 09/2026 | **Revisão:** 26/10/2026

---

### Placar Consolidado de Homologação (KPIs)
* **Universo Auditado:** 2.345 Cartões Good Card / Benefícios Especiais Monitorados
* **Estrutura de Colunas:** 20 Colunas Oficiais (Conferência Geral) + Lista de Bloqueio Operacional
* **Volume Financeiro Monitorado:** R$ 1.120.400,00 em Limites e Saldos Ativos
* **Status do Motor no App:** 🟢 100% Homologado e Ativo em Produção

---

### 1. Visão Geral da Solução e Arquitetura Unificada
O **Módulo 04 - Bloqueio Good Card** é a frente preventiva e de contenção de risco financeiro do Grupo Cencosud Brasil (bandeiras Cencosud, Atacado e Perini).

O robô monitora a base integral de cartões emitidos pela operadora Ticket/Good Card e cruza em tempo real contra as bases corporativas de afastamentos previdenciários e demissões. O objetivo primordial é **evitar que colaboradores afastados por mais de 15 dias continuem consumindo o cartão convênio**, gerando saldos devedores impagáveis que a folha de pagamento não consegue descontar durante a suspensão do contrato de trabalho (Artigo 476 da CLT).

---

### 2. Estrutura de Pastas e Inventário de Relatórios Consumidos
O módulo é executado a partir do diretório corporativo `Conferência Bloqueio Good Card`:

| Arquivo Corporativo | Formato / Peso | Finalidade e Regras Associadas |
| :--- | :---: | :--- |
| `Bloqueio Good Card Cencosud Atacado Perini.xlsx` | XLSX (~850 KB) | Extrato da operadora Ticket/Good Card contendo Matrícula, Cartão, Limite Total, Valor Utilizado e Saldo Disponível. |
| `COLA_Colaboradores_Afastados.csv` | CSV (~350 KB) | Mapeamento de afastamentos médicos, licenças maternidade e afastamentos previdenciários do INSS. |
| `COLA_Demitidos.csv` | CSV (~3.2 MB) | Base oficial de colaboradores demitidos e desligados no período. |

---

### 3. Matriz de Regras de Negócio e Validações Matemáticas (RNs)

| Código | Tipo de Validação | Regra Aplicada & Fórmula Homologada | Base Origem | Ação Recomendada |
| :---: | :--- | :--- | :--- | :--- |
| **RN001** | Cruzamento Operadora x Censo | Cruzamento por Matrícula da base de cartões da operadora Ticket contra afastamentos e demissões. | Fatura / Censo | Identificar inconsistências funcionais ativas. |
| **RN002** | Limiar de 15 Dias (Art. 476 CLT) | Se afastamento médico/INSS ativo &gt; 15 dias: *"Afastado &gt; 15 dias - Bloqueio Imediato Obrigatório"*. | COLA Afastados | Solicitar bloqueio temporário do cartão no portal Ticket. |
| **RN003** | Cancelamento por Demissão | Se colaborador demitido com cartão ativo no portal: *"Demitido - Cancelamento Definitivo"*. | COLA Demitidos | Cancelar cartão definitivamente e impedir uso indevido. |
| **RN004** | Monitoramento de Saldo | Confronto entre `Limite Total`, `Valor Utilizado` e `Saldo Disponível`. Apuração de risco de inadimplência. | Operadora | Apontar colaboradores com limite estourado. |
| **RN005** | Emissão da Lista de Bloqueio | Geração de aba operacional formatada com Matrícula, Número do Cartão, Motivo e Ação recomendada. | Motor Excel | Enviar em lote para a operadora Ticket/Good Card. |

---

### 4. Mecanismos Avançados de Engenharia e Blindagem
* **Padrão Visual Institucional:** Cabeçalhos base em Azul Marinho (`#1F497D`) com texto Branco (`#FFFFFF`). Colunas de auditoria, batimento e parecer em Amarelo Puro (`#FFFF00`) com texto em Vermelho Escuro (`#9C0006`) em negrito.
* **Normalização Resiliente de Matrícula:** Remoção de caracteres espúrios, formatação em string pura sem perda de zeros à esquerda.
* **Geração Automática de Lista Operacional:** O relatório gera uma aba específica para envio rápido ao portal da operadora, acelerando o bloqueio de cartões.

---

### 5. Cenários Didáticos Simulados (Anonimizados LGPD)

| Matrícula | Colaborador (Fictício) | Cartão | Limite Total | Utilizado | Saldo Disp. | Diagnóstico Analítico |
| :---: | :--- | :---: | :---: | :---: | :---: | :--- |
| `400101` | COLABORADOR ATIVO A | 6031...1011 | R$ 600,00 | R$ 150,00 | R$ 450,00 | **REGULAR:** Colaborador ativo, consumo dentro do padrão. |
| `400102` | COLABORADOR AFASTADO B | 6031...2022 | R$ 800,00 | R$ 500,00 | R$ 300,00 | **BLOQUEIO IMEDIATO:** Afastado INSS há 28 dias. Bloquear cartão. |
| `400103` | COLABORADOR DEMITIDO C | 6031...3033 | R$ 500,00 | R$ 0,00 | R$ 500,00 | **CANCELAMENTO:** Desligado em 05/09. Cancelar cartão. |

---

### 6. Guia Operacional Passo a Passo
1. **Abrir a Central:** Executar `Central_Beneficios_Cencosud.exe`.
2. **Acessar Módulo:** Clicar na aba **"04 - Bloqueio Good Card"**.
3. **Conferir Checklist:** Confirmar presença da base da operadora e arquivos de afastados/demitidos.
4. **Executar:** Clicar no botão azul **"Gerar Conferência Bloqueio Good Card"**.
5. **Ação Operacional:** Copiar os cartões da aba **Lista de Bloqueio** e efetuar o bloqueio imediato no portal Good Card/Ticket.

---

### 7. Parecer Técnico de Homologação
O **Módulo 04 - Bloqueio Good Card** foi 100% homologado pela equipe de Recursos Humanos e Engenharia de Dados, blindando a companhia contra passivos financeiros decorrentes de compras efetuadas por colaboradores com contrato suspenso ou rescindido.

**Engenharia de Dados & Automação de RH** • *Desenvolvimento de Software & Algoritmos Analíticos*  
**Administração de Pessoal & Remuneração** • *Homologação de Negócio, Regras CLT & Diretrizes Cencosud*
