"""Teste isolado usando somente BASICO Noturno 2026.xlsx.
Altere CAMINHO_ARQUIVO se desejar testar fora da estrutura definitiva.
"""
from pathlib import Path
import sqlite3
from openpyxl import load_workbook

from importar_horarios import localizar_aba, carregar_dicionario_professores, extrair_registros_turma

CAMINHO_ARQUIVO = Path(r"C:\Users\391665\OneDrive\EMCA\Horários\2026\04 - BASICO\BASICO Noturno 2026.xlsx")

wb = load_workbook(CAMINHO_ARQUIVO, data_only=True)
base = localizar_aba(wb, "Base")
dic = carregar_dicionario_professores(base)

for aba in ("R3", "S3", "T3"):
    registros = extrair_registros_turma(localizar_aba(wb, aba), aba, "NOTURNO", dic)
    setembro = [r for r in registros if r["data"].startswith("2026-09-")]
    print(f"{aba}: {len(registros)} registros no ano / {len(setembro)} em setembro")
    for r in setembro[:5]:
        print("  ", r)
