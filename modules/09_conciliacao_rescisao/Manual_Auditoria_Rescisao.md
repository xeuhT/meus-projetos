# CENCOSUD BRASIL
Engenharia de Dados & Automação de RH • Manual Técnico e Executivo Completo

## Motor Analítico de Auditoria de Rescisões e FOPAG (Projeto Rescisão - Versão V001)
**Módulo:** Auditoria de Rescisão | **Status:** Produção Inicial (V1) | **Atualizado em:** 10/09/2026

---

### 1. Visão Geral da Solução e Arquitetura Modular

O **Projeto Rescisão V001** é uma solução de Engenharia de Dados desenvolvida para realizar a conciliação automatizada da folha de pagamento massiva de rescisões da Cencosud. O sistema substitui a conferência manual de PDFs individuais, validando o cálculo das rubricas legais (CLT) e descontos (INSS) em nível de centavos contra o extrato gerado pelo sistema de RH (LG).

**Cobertura Integral de Dados:** O motor consome a base de folha de pagamento massiva (formato binário `.xlsb`), agrupa os registros por matrícula do colaborador e executa o processamento vetorial de auditoria, detectando anomalias financeiras e gerando um laudo por linha processada em segundos.

**Arquitetura Técnica do Sistema:**
* **`auditor_massivo.py`**: Motor principal de ETL (Extract, Transform, Load) e auditoria. Orquestra a extração dos dados binários, corrige anomalias nativas de formatação, aplica as regras de negócio em cascata e exporta o relatório executivo.
* **Módulo de Heurísticas**: Lógica matemática embutida para cenários em que o sistema fonte não explicita os parâmetros de cálculo (como avos de férias omitidos ou dias de atestados rateados).
* **Renderização OpenPyXL**: Módulo de formatação que injeta estilos visuais e highlights condicionais diretamente na planilha de output, sem necessidade de macro.

---

### 2. Inventário de Bases de Entrada

Atualmente, o processo unificou a auditoria em uma única base massiva contendo a explosão completa da folha, dispensando PDFs.

| Categoria da Base | Padrão de Nome do Arquivo / Formato | Função no Processo de Auditoria |
| :--- | :--- | :--- |
| **FOPAG Rescisão (Massiva)** | `RFolha Rescisão Cencosud_*.xlsb` | Base mestre e extrato detalhado contendo todos os eventos de provento e desconto, dados de admissão/demissão, referência e valor do sistema. |

---

### 3. Matriz de Regras de Negócio e Inconsistências (RNs)

O motor já possui cobertura para as regras vitais e de alta complexidade matemática do escopo de rescisão. *(Nota: Matriz será expandida nas próximas versões).*

| Código | Rubrica / Evento | Regra Aplicada no Motor Python | Tratamento de Exceções |
| :--- | :--- | :--- | :--- |
| **721** | Saldo de Salário | `(Base Salarial / 30) * Dias Ref` | Delineamento algorítmico que abate faltas, Auxílio Doença (137) e Acidente (138). |
| **725** | Aviso Prévio Indenizado | `(Base / 30) * (30 dias + 3 dias/ano trabalhado)` | Consome a Rubrica 20208 para determinar os dias exatos de direito pela lei. |
| **737** | Férias Vencidas | *Matemática Reversa:* `Valor Sistema / Valor Diário` | Infere a quantidade exata de dias pagos quando a referência do sistema vem zerada. |
| **738** | Férias Proporcionais | *Matemática Reversa:* `Valor Sistema / Valor de 1 Avo` | Infere os avos fracionados considerando eventuais perdas de dias de direito (Rubrica 20305). |
| **747/748/749** | Terços de Férias | `Base do Evento Principal / 3` | Audita o terço constitucional sobre Vencidas, Proporcionais e Indenizadas. |
| **5013** | INSS Mês | **Tabela Progressiva 2026** fatiada em 4 faixas (7.5% a 14%) | Abate quantias previamente descontadas (ex: INSS Férias - Rubrica 5015) para evitar falso-positivo. |

---

### 4. Mecanismos Avançados de Blindagem de Dados

Para garantir auditoria de nível sênior, a Versão V001 conta com quatro proteções algorítmicas desenvolvidas para lidar com vícios estruturais das bases do RH LG:

**A. Correção Nativa de Datas Seriais (Excel Time Fix)**
Bases massivas `.xlsb` gravam datas como números seriais corrompidos (Ex: 46239). O script implementa uma conversão forçada `pd.to_datetime(origin='1899-12-30')`, devolvendo o formato limpo `DD/MM/AAAA` ao usuário final.

**B. Heurística Dinâmica de Saldo de Salário**
O sistema de RH frequentemente rateia dias trabalhados contra dias de atestado/DSR de formas imprevisíveis, não deduzindo o total de dias da coluna de referência principal. 
*A Solução:* O Python calcula permutações de possíveis abatimentos monetários (usando as rubricas 137, 138, 152, 20417) e escolhe matematicamente o cenário exato que o RH utilizou, zerando falsas divergências.

**C. Matemática Reversa de Férias e Tolerância a Frações**
Como a base extraída oculta o número de "Avos" na coluna referência (`0.0`), o script usa a matemática reversa para deduzir o fator aplicado. O motor é capaz de rastrear avos fracionados (ex: `0.5 avos` ou `2.65 avos`) causados por redução de dias de direito por faltas ou regimes de tempo parcial.

**D. Filtro de Tolerância a Arredondamentos**
Para isolar falhas graves, qualquer divergência de sistema `<= R$ 1,00` é automaticamente suprimida e classificada como `OK`, absorvendo perdas de arredondamentos de ponto flutuante do sistema legado LG.

---

### 5. Output e Formatação do Relatório Executivo (XLSX)

O relatório gerado pelo sistema é salvo em formato `.xlsx` nativo e possui os seguintes aprimoramentos de legibilidade operacional:

* **Colunas de Auditoria**: Foram adicionadas três colunas matrizes de controle:
  * `VALOR CALCULADO PYTHON`: O valor auditado e homologado pelo motor.
  * `DIVERGÊNCIA`: Diferença monetária absoluta detectada.
  * `STATUS AUDITORIA`: String detalhada explicando a conclusão (ex: `OK: Reversa (0.5 avos inferidos sobre 15 dias direito)`).
* **Destaque Condicional Vermelho (Automação de UI)**: Toda a extensão da linha (row) que apresentar a string `DIVERGENTE` é pintada automaticamente com o preenchimento `#FFCCCC` (vermelho-claro). O analista bate o olho exclusivamente nas anomalias.
* **Alertas Preditivos de INSS**: Casos de anomalias por recibos ausentes (ex: Férias do Mês) sinalizam: `Aviso: INSS possivelmente já descontado em recibo de Férias/Adiantamento`.
