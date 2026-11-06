# 🛡️ Módulo 02: Auditoria de Seguro de Vida (Básico e Complementar 12x / 24x)

![Status](https://img.shields.io/badge/Status-Concluído-brightgreen)
![Coberturas](https://img.shields.io/badge/Planos-Básico%20%7C%2012x%20%7C%2024x-blue)

## 📌 Objetivo do Módulo
Garantir a conformidade entre os colaboradores elegíveis ao Seguro de Vida Corporativo (apólice básica) e as adesões voluntárias aos planos de capital ampliado (12x ou 24x o salário nominal), validando os descontos efetuados na folha de pagamento contra as faturas e relatórios emitidos pela seguradora.

## ⚙️ Regras de Negócio e Validações
* **Validação de Elegibilidade:** Confronto entre a base ativa de colaboradores elegíveis vs. faturas da seguradora.
* **Capital Segurado Dinâmico:** Auditoria do cálculo do prêmio com base no múltiplo salarial cadastrado (12x ou 24x o salário contratual).
* **Detecção de "Vidas Fantasmas":** Identificação de colaboradores desligados mantidos indevidamente na fatura da seguradora, gerando cobranças retroativas indevidas.
* **Geração de Arquivo de Carga:** Construção do layout de movimentação cadastral (inclusões, exclusões e alterações salariais) para envio mensal à seguradora.

## 📂 Arquivos e Scripts do Módulo
* `app_seguro_vida.py`: Interface / script principal de processamento e conciliação.
* Scripts auxiliares de geração de arquivo de carga e relatórios analíticos de divergências.

## 🚀 Como Executar
```bash
python app_seguro_vida.py
```
