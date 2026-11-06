# CENCOSUD BRASIL
## Engenharia de Dados & Automação de RH • Manual Técnico e Executivo Oficial
### Central Executiva de Auditoria de Benefícios • Módulo 10: Intermitentes (VA e VT)

---

**Módulo:** 10 - Intermitentes (VA e VT) | **Interface:** Central Executiva (`.exe`) | **Competência:** 08/2026 | **Revisão:** 26/10/2026

---

### Placar Consolidado de Homologação (KPIs)
* **Universo Auditado:** 1.620 Colaboradores com Contrato Intermitente
* **Estrutura de Colunas:** 26 Colunas Oficiais de Validação
* **Volume Financeiro Auditado:** R$ 145.280,00 em Benefícios Intermitentes
* **Status do Motor no App:** 🟡 Pendente Envio para Validação Operacional

---

### 1. Visão Geral da Solução e Arquitetura Unificada
O **Módulo 10 - Intermitentes (VA e VT)** audita as particularidades da concessão de benefícios aos empregados com contrato de trabalho intermitente (Art. 452-A da CLT).
A solução cruza as convocações emitidas e assinadas contra os dias efetivamente trabalhados, confrontando os valores de Vale Transporte (Folha 36 e 70) e Alimentação contra os descontos em folha e a presença registrada no ponto biométrico.

---

### 2. Matriz de Regras de Negócio e Validações Matemáticas (RNs)
* **RN001 - Batimento de Convocação:** Cruzamento de dias convocados x dias com registro de ponto.
* **RN002 - Quitação de VT Intermitente:** Validação de créditos concedidos para deslocamento x desconto em folha.
* **RN003 - Vale Alimentação Intermitente:** Proporcionalidade estrita aos turnos cumpridos superiores a 6 horas.
* **RN004 - Tratamento de No-Show:** Identificação de convocações não comparecidas com estorno do benefício.
