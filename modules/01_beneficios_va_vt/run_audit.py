"""
Script Principal de Orquestração da Auditoria de 6% (Vale Transporte).
Uso:
  python run_audit.py [--pasta "..."] [--competencia "MM/AAAA"] [--saida "..."]
"""

import os
import sys
import argparse
from datetime import datetime

# Adicionar pasta local ao path
sys.path.append(os.path.dirname(__file__))

from logger_telemetry import AuditLogger
from data_loader import DataLoader
from audit_engine import AuditEngine
from excel_exporter import ExcelExporter

def main():
    parser = argparse.ArgumentParser(description="Automação da Auditoria de Vale Transporte (6%)")
    parser.add_argument(
        "--pasta",
        type=str,
        default=r"C:\Users\matcaval\OneDrive - Cencosud\Documentos\SISTEMAS RH\VA e VT\6%\Relatórios Originais",
        help="Caminho da pasta contendo os relatórios originais em CSV"
    )
    parser.add_argument(
        "--competencia",
        type=str,
        default=None,
        help="Competência de apuração no formato MM/AAAA (se omitido, infere automaticamente)"
    )
    parser.add_argument(
        "--saida",
        type=str,
        default=None,
        help="Caminho completo do arquivo de saída .xlsx gerado"
    )

    args = parser.parse_args()

    # Inicializar Logger
    log_dir = os.path.join(os.path.dirname(__file__), "logs")
    logger = AuditLogger(log_dir=log_dir)

    try:
        logger.log(f"Pasta de Origem: {args.pasta}")
        if not os.path.exists(args.pasta):
            raise FileNotFoundError(f"A pasta de origem não foi encontrada: {args.pasta}")

        # 1. Carga de Dados
        loader = DataLoader(args.pasta, logger=logger)
        
        # Detectar Competência
        if args.competencia:
            loader.competencia = args.competencia
            loader.mes_competencia = int(args.competencia.split("/")[0])
            loader.ano_competencia = int(args.competencia.split("/")[1])
            logger.log(f"Competência informada via parâmetro: {loader.competencia}")
        else:
            loader.infer_competencia()
            
        logger.record_step("Detecção de Parâmetros e Competência")

        # Carregar bases
        logger.log("Carregando bases de dados de origem...")
        df_folha = loader.load_folha()
        df_dem = loader.load_demitidos()
        df_afast = loader.load_afastados()
        df_ferias = loader.load_ferias()
        df_nom = loader.load_nomina()
        df_compra = loader.load_compra()
        df_retorno = loader.load_retorno()
        df_orientacoes = loader.load_orientacoes()
        
        logger.record_step("Carga e Limpeza Prévia das Bases")

        # 2. Motor de Auditoria e Cálculo em Pandas
        engine = AuditEngine(competencia=loader.competencia, logger=logger)
        df_auditado, df_compra_enriquecida, df_retorno_enriquecida = engine.process(
            df_folha=df_folha,
            df_dem=df_dem,
            df_afast=df_afast,
            df_ferias=df_ferias,
            df_nom=df_nom,
            df_compra=df_compra,
            df_retorno=df_retorno
        )
        logger.record_step("Processamento do Motor de Auditoria (Pandas)")

        # 3. Definição do Arquivo de Saída
        if not args.saida:
            comp_clean = loader.competencia.replace("/", "_")
            pasta_destino = os.path.dirname(args.pasta)  # Pasta 6%
            args.saida = os.path.join(pasta_destino, f"Conferência 6% _Auditada_{comp_clean}.xlsx")

        # 4. Exportação Excel com Destaques Visuais
        exporter = ExcelExporter(args.saida, logger=logger)
        divergencias_total = exporter.export(
            df_conferencia=df_auditado,
            df_compra=df_compra_enriquecida,
            df_retorno=df_retorno_enriquecida,
            df_orientacoes=df_orientacoes
        )
        logger.record_step("Exportação da Planilha Estática Formatada")

        # 5. Finalizar Telemetria
        summary_stats = {
            "Total de Colaboradores Auditados": len(df_auditado),
            "Colaboradores com Compra de VT": int((df_auditado["Custo Total Vale Transporte"] > 0).sum()),
            "Colaboradores com Desconto em Folha": int((df_auditado["6290 - Vale Transporte "] > 0).sum()),
            "Inconsistências / Divergências Apontadas": divergencias_total,
            "Registros Conformes (OK)": len(df_auditado) - divergencias_total
        }
        logger.finalize(summary_stats)
        
        print("\n" + "="*70)
        print("AUDITORIA CONCLUÍDA COM SUCESSO!")
        print(f"Arquivo gerado: {args.saida}")
        print(f"Total de registros: {len(df_auditado):,}")
        print(f"Divergências destacadas em vermelho: {divergencias_total:,}")
        print("="*70)

    except Exception as e:
        logger.log(f"ERRO CRÍTICO NO PROCESSAMENTO: {e}")
        import traceback
        logger.log(traceback.format_exc())
        raise e

if __name__ == "__main__":
    main()
