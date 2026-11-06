# 🛠️ Utilitários e Automações Complementares de RH

![Status](https://img.shields.io/badge/Status-Concluído-brightgreen)
![Segurança](https://img.shields.io/badge/LGPD-Anonimizador%20de%20Dados-red)

## 📌 Objetivo
Conjunto de ferramentas auxiliares para apoiar as rotinas operacionais, garantir conformidade com a LGPD e otimizar processos de cadastro, enquadramento e extração de dados no ecossistema de Recursos Humanos.

## ⚙️ Ferramentas Disponíveis

### 1. `anonimizar_planilhas.py` (Governança & LGPD)
* **Objetivo:** Mascaramento e anonimização irreversível de PII (CPFs, nomes de colaboradores, números de documentos, valores nominais de salários) para compartilhamento seguro de planilhas e relatórios para auditorias, homologações ou suporte externo.
* **Garantia:** Converte dados sensíveis em identificadores anônimos mantendo a integridade relacional para testes.

### 2. `automacao_salario_familia.py` (Enquadramento Legal)
* **Objetivo:** Automação do enquadramento de colaboradores com direito à cota do Salário Família conforme tabela vigente da Previdência Social, confrontando teto salarial, número de dependentes válidos e controle de comprovação de vacinação e frequência escolar.

### 3. `extrair_cargos.py` (Engenharia de Dados Cadastrais)
* **Objetivo:** Normalização, deduplicação e taxonomia da tabela de cargos e CBOs em todas as filiais e bandeiras do grupo, gerando base padronizada para enquadramentos sindicais e pesquisas salariais.

### 4. `melhora_processos_rh.py` (Otimização Contínua)
* **Objetivo:** Scripts e rotinas de automação rápida para higienização de bases e eliminação de tarefas manuais repetitivas.

## 🚀 Como Executar
```bash
python anonimizar_planilhas.py
python automacao_salario_familia.py
python extrair_cargos.py
```
