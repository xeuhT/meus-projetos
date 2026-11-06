# 🏖️ Módulo 08: Auditoria Integral de Férias e Fechamento Mensal (Projeto AD004)

![Status](https://img.shields.io/badge/Status-Concluído-brightgreen)
![Versão](https://img.shields.io/badge/Motor-v008%20Consolidado-blue)

## 📌 Documentação Oficial da Central de Auditoria de Férias e Folha

---

### 📂 Estrutura do Diretório

```text
documentacao/
│
├── Manual_Tecnico_e_Executivo_Auditoria_Ferias_Cencosud.pdf # [PDF OFICIAL] Manual Homologado (6 Páginas A4)
├── documentacao_auditoria_ferias.html                       # [FONTE HTML] Template visual corporativo estruturado
├── generate_doc_pdf.py                                      # [COMPILADOR] Script de automação e renderização PDF
├── README.md                                                # Este arquivo de especificações e guia
│
└── releases/                                                # [RELEASES & ENTREGAS]
    ├── Projeto_Auditoria_Ferias_GitLab.zip                  # Pacote de entrega para repositório corporativo
    └── Scripts_Auditoria_Ferias_AD004.zip                   # Pacote complementar de código homologado
```

---

### 📄 Conteúdo do Manual Oficial (6 Páginas)

1. **Página 1: Visão Executiva & Arquitetura Geral**
   - Contexto institucional do Grupo Cencosud Brasil.
   - Objetivos estratégicos de conciliação 360°, conformidade e segurança da informação.
   - Diagrama de blocos de alto nível dos 3 pilares de auditoria.

2. **Página 2: Motor de Auditoria da Folha Mensal (Aba 3)**
   - Conciliação entre relatório LG e espelho contábil da folha (65 colunas).
   - Mapeamento dinâmico de competência (`extract_fopa_competencia`).
   - Blindagem e validação estrutural da FOPA completa (`is_csv_fopa_folha_completa` > 30 colunas).
   - Filtro de colaboradores ativos e afastados (`is_active_afastados`).

3. **Página 3: Motor de Auditoria do Cálculo de Férias (Aba 2)**
   - Reconstituição matemática integral do cálculo de férias (86 colunas).
   - Conferência de proventos, bases de cálculo, médias, deduções (INSS, IRRF, Pensão Alimentícia) e valor líquido.

4. **Página 4: Conferência de Férias Formatada (Aba 1)**
   - Validação de prazos legais CLT (aviso prévio e pagamento com 48h de antecedência).
   - Cruzamento de feriados municipais/estaduais e impedimento de início em D-2 do descanso semanal.
   - **Regra de Negócio RN014**: Alerta de bloqueio para solicitação de 1ª parcela de 13º salário com gozo em Novembro, Dezembro ou Janeiro.

5. **Página 5: Guia Operacional da Interface Executiva Desktop**
   - Manual da interface gráfica (`app_conferencia_ferias.py` e executável `Conferencia_Ferias_Cencosud.exe`).
   - Monitoramento visual do checklist de arquivos de entrada e diagnóstico de integridade.
   - Tratamento amigável de concorrência e planilhas abertas no Excel (`FileLockModal`).

6. **Página 6: Segurança, LGPD, Troubleshooting e Roadmap**
   - Diretrizes de privacidade e conformidade com a LGPD.
   - Tabela de resolução de incidentes e mensagens de erro comuns.
   - Histórico de versões e roadmap de evolução contínua.

---

### ⚙️ Como Regerar a Documentação em PDF

A geração do PDF utiliza o Google Chrome (ou Microsoft Edge) em modo headless para garantir renderização idêntica em qualquer ambiente Windows:

```bash
# Executar a partir do diretório raiz ou da pasta documentacao/
python generate_doc_pdf.py
```

O script:
1. Gera o arquivo HTML atualizado em `documentacao/` e na raiz.
2. Aciona o Chrome Headless com os parâmetros de margem e tamanho A4 calibrados.
3. Valida com a biblioteca `PyMuPDF` que o documento possui exatamente 6 páginas.
4. Salva a versão final tanto em `documentacao/` quanto na raiz do projeto (`Manual_Tecnico_e_Executivo_Auditoria_Ferias_Cencosud.pdf`).
