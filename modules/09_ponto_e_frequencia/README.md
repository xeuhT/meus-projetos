# ⏱️ Módulo 05: Auditoria de Ponto, Jornadas Críticas e Justificativas

![Status](https://img.shields.io/badge/Status-Concluído-brightgreen)
![Conformidade](https://img.shields.io/badge/Compliance-CLT%20%7C%20Portaria%20671-blue)

## 📌 Objetivo do Módulo
Auditoria contínua do espelho de ponto eletrônico, detecção de passivos trabalhistas decorrentes de extrapolação de jornadas legais (Art. 59 da CLT), controle de sobreaviso, conferência de escalas de folgas de operadores e gestão de justificativas e atestados médicos.

## ⚙️ Principais Relatórios e Regras Implementadas
1. **Jornadas Críticas:** Detecção de jornadas superiores a 10h diárias, interjornada inferior a 11h consecutivas e intrajornada suprimida (intervalo de refeição).
2. **Auditoria de Justificativas:** Rastreamento de atestados médicos recorrentes, afastamentos não lançados e justificativas manuais atípicas.
3. **Controle de Folgas e DSR:** Verificação do cumprimento da folga semanal remunerada a cada 6 dias de trabalho contínuo e escalas em domingos e feriados (regras específicas do varejo supermercadista).
4. **Contratos Intermitentes:** Acompanhamento de convocações, horas acordadas e aceite de jornada.

## 📂 Arquivos e Scripts do Módulo
* `gerar_relatorio_jornada_critica.py`: Motor de análise de limites de jornada e interjornada.
* `gerar_relatorio_justificativas.py`: Relatório analítico de motivos de abono e atestados.
* `gerar_relatorio_controle_folgas.py`: Auditoria de DSR e escalas 6x1 / 5x2.
* `gerar_relatorio_intermitentes.py`: Conferência de frequência e horas de intermitentes.

## 🚀 Como Executar
```bash
python gerar_relatorio_jornada_critica.py
```
