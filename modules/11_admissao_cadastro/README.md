# 📋 Módulo 11: Admissão & Cadastro de Colaboradores (Censo Master)
### Primeiro Aplicativo Desktop Desenvolvido em Python — Cencosud Brasil

![Status](https://img.shields.io/badge/Status-100%25_Concluído-brightgreen)
![App Desktop](https://img.shields.io/badge/Interface-Tkinter_GUI-blue)
![Produção](https://img.shields.io/badge/Uso-Ativo_em_Produção-navy)
![ROI](https://img.shields.io/badge/ROI-17.5h%2Fmês_salvas-brightgreen)

---

## 📌 Visão Histórica & Importância Estratégica

Este foi **o primeiro aplicativo corporativo desenvolvido em Python por Matheus Cavalcanti na Cencosud Brasil**.

Antes da criação desta ferramenta, a consolidação cadastral dos novos colaboradores era feita de forma artesanal e fragmentada através de planilhas dispersas, levando dias de conferências manuais e gerando frequentes inconsistências em crachás, escalas de trabalho, cadastros do ponto e concessão indevida de Salário Família.

O aplicativo foi concebido com uma **interface gráfica própria (Tkinter)** e terminal visual de execução embutido, permitindo que os analistas de Recursos Humanos executem o processamento com zero atrito e total transparência.

---

## ⚙️ Arquitetura de Auto-Discovery & Cruzamento de 6 Relatórios LG

A ferramenta implementa um motor algorítmico inteligente de **auto-discovery por prefixo** que varre o diretório de entrada e localiza automaticamente 6 relatórios brutos emitidos pelo sistema de folha **LG Lugar de Gente**:

1. `01. Relação de Funcionários`: Base mestre de dados admissionais, cargos e salários nominais.
2. `02. Dependentes Cadastrados`: Relação de filhos e tutelados para cálculo de benefícios e deduções.
3. `03. Escalas de Trabalho`: Mapeamento de jornadas, turnos e lotações por loja/CD.
4. `04. Emissão de Crachás`: Registro de identificação física e acesso às unidades.
5. `05. Perfil de Ponto`: Parâmetros de apuração de frequência e banco de horas.
6. `06. Histórico de Folha`: Parâmetros de enquadramento tributário e previdenciário.

---

## 🎯 Entregáveis & Benefícios Operacionais

* **Relatório Geral de Admissão (Censo Master):** Consolidação em aba única com todas as informações vitais do colaborador higienizadas e validadas.
* **Saneamento de Elegibilidade ao Salário Família:** Validação cruzada estrita entre faixa de remuneração da Portaria Interministerial, certidões de nascimento válidas e laudos de vacinação/frequência escolar.
* **Economia Operacional:** **17,5 horas líquidas poupadas por mês**, em uso contínuo e sem falhas há meses pela equipe de admissão da Cencosud.

---

## 🚀 Como Executar

```bash
python app_relatorio_admissao.py
```
*(Ou através do executável standalone disponibilizado para a equipe operacional).*
