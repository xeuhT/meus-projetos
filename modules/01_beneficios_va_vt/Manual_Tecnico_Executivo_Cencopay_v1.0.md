# CENCOSUD BRASIL
## Engenharia de Dados & Automação de RH • Manual Técnico e Executivo Oficial
### Central Executiva de Auditoria de Benefícios • Módulo 13: Cencopay

---

**Módulo:** 13 - Cencopay (Cartão Virtual Benefícios) | **Interface:** Central Executiva (`.exe`) | **Competência:** 09/2026 | **Revisão:** 26/10/2026

---

### Placar Consolidado de Homologação (KPIs)
* **Universo Auditado:** 4.924 Colaboradores Ativos com Cartão Virtual
* **Estrutura de Colunas:** 24 Colunas Oficiais de Auditoria
* **Volume Financeiro Auditado:** R$ 1.121.010,21 em Gastos e Descontos Apurados
* **Status do Motor no App:** 🟡 Pendente Envio para Validação Operacional

---

### 1. Visão Geral da Solução e Arquitetura Unificada
O **Módulo 13 - Cencopay** realiza a auditoria do Cartão Virtual Benefício Cencopay. O motor calcula o Total de Saldo Devedor (`Evento 14004 + Evento 14002`) e aplica rigorosamente o limite de **30% do salário líquido a receber** (`Evento 10990 + Evento 126`), validando o **Evento 5081 (Desconto Gasto Cartão Virtual Benefício)** e transferindo qualquer saldo excedente para o mês seguinte.
