# 💳 Módulo 03: Automação e Leitura de Extrato de Ticket Plus

![Status](https://img.shields.io/badge/Status-Concluído-brightgreen)
![Tipo](https://img.shields.io/badge/Automação-Robô%20de%20Consolidação-orange)

## 📌 Objetivo do Módulo
Automatizar a extração, leitura, higienização e consolidação dos relatórios analíticos de compras e utilizações de cartões Ticket Plus / Convênio, gerando bases unificadas para lançamento contábil e conciliação de descontos em folha de pagamento.

## ⚙️ Regras de Negócio e Funcionalidades
* **Consolidação Multi-Unidade:** Agrupamento de relatórios emitidos por filiais, centros de distribuição e escritórios corporativos.
* **Normalização Cadastral:** De-para automático de matrículas, CPFs e centros de custo entre o padrão da operadora Ticket e o sistema LG de Folha.
* **Validação de Limites de Compra:** Apuração de gastos que ultrapassem os limites contratuais ou margens consignáveis permitidas por convenção coletiva.

## 📂 Arquivos e Scripts do Módulo
* `robo_extrato_ticket.py`: Motor principal de processamento de extratos de compras.
* `consolidar_excel.py`: Script para unificação e tratamento de pastas de trabalho em lote.
* `utils.py`: Funções auxiliares de parsing de moeda, datas e normalização de strings.

## 🚀 Como Executar
```bash
python robo_extrato_ticket.py
```
