# 🍱 Módulo 01: Benefícios Corporativos e VA/VT (18 Frentes Oficiais)

![Status](https://img.shields.io/badge/Status-Concluído-brightgreen)
![Módulos](https://img.shields.io/badge/Conferências-18_Frentes-blue)

## 📌 Objetivo do Módulo
Centralizar e automatizar a conciliação mensal dos benefícios corporativos, faturamento de operadoras e cálculo de descontos em folha de pagamento (sistema LG Lugar de Gente) para milhares de colaboradores de múltiplos CNPJs.

## ⚙️ As 18 Frentes Oficiais de Conferência
1. **Vale Transporte (VT 6%):** Apuração do desconto legal dos 6% constitucionais vs. custo real de aquisição da operadora.
2. **Rateio de Centros de Custo (CVE):** Cálculo automático de rateio por unidade de lotação e custo operacional.
3. **Good Card Alimentação / Refeição:** Auditoria de créditos e desconto de evento 6280 em folha.
4. **Devoluções e Estornos:** Controle analítico de créditos cancelados ou a restituir.
5. **Cencopay:** Conciliação de recargas e subsídios em cartão próprio.
6. **Refeitório de Lojas:** Apuração de refeições servidas in loco e conferência de catracas/listas.
7. **Refeitório Escritório / Corporativo:** Descontos de refeição administrativa.
8. **Bloqueio Preventivo de Cartões:** Identificação ativa de colaboradores demitidos ou afastados para suspensão de recargas.
9. **Cesta Prêmio / Ticket Bretas:** Conferência de concessões em campanhas e premiações.
10. **Unidades de Entrega Física:** Rastreabilidade logística de distribuição de cartões e vales.
11. **Contratos Intermitentes:** Regras de proporcionalidade e convocação de benefícios.
12. **Auditores de Margem de Benefício:** Garantia de margem consignável legal para descontos.
13. **Coparticipação de Transporte:** Divergências entre faturado e descontado.
14. **Conferência de Afastados com Benefício Ativo:** Suspensão e conciliação de pagamentos indevidos.
15. **Conferência de Menores Aprendizes e Estagiários:** Carga diferenciada de benefícios conforme contrato.
16. **Auditoria de Retroativos:** Lançamentos em competências posteriores e acertos retroativos.
17. **Relatório Executivo de Glosas:** Apuração de glosas contratuais e divergências de faturamento.
18. **Fechamento Mensal Consolidado:** Emissão do pacote de conferência para liberação financeira.

## 📂 Arquivos e Scripts do Módulo
* `run_audit.py`: Orquestrador de execução em lote das conferências.
* `audit_engine.py`: Motor de regras de negócio, batimento e cálculo matemático.
* `data_loader.py`: Ingestão de bases LG, operadoras e arquivos CSV/Excel/XLSB.
* `excel_exporter.py`: Geração de relatórios com formatação executiva condicional.
* `logger_telemetry.py`: Telemetria, rastreabilidade e logs de auditoria.
* `config.py`: Parâmetros de regras, eventos da folha e caminhos operacionais.

## 🚀 Como Executar
```bash
python run_audit.py
```
