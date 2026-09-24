from pathlib import Path

ANO_PADRAO = 2026

PASTA_HORARIOS = Path(r"C:\Users\assum\OneDrive\EMCA\Horários\2026")
PASTA_PROJETO = Path(r"C:\Projetos\LivroPontoEMCA")
CAMINHO_BANCO = PASTA_PROJETO / "dados" / "livro_ponto_emca.db"
PASTA_EXPORTACOES = PASTA_PROJETO / "exportacoes"

# Cada item representa uma turma que deve ser importada.
# "arquivo" é relativo à pasta C:\Users\391665\OneDrive\EMCA\Horários\2026
TURMAS = [
    {
        "turma": "AVI 26",
        "turno": "NOTURNO",
        "arquivo": Path("01 - AVI") / "AVI NOTURNO 2026.xlsx",
        "aba_horario": "Rel FATEPI",
    },
    {
        "turma": "CEL 26",
        "turno": "NOTURNO",
        "arquivo": Path("02 - CEL") / "CEL NOTURNO 2026.xlsx",
        "aba_horario": "Rel FATEPI",
    },
    {
        "turma": "GMP 01",
        "turno": "NOTURNO",
        "arquivo": Path("03 - GMP") / "GMP NOTURNO 2026.xlsx",
        "aba_horario": "GMP01",
    },
    {
        "turma": "GMP 02",
        "turno": "NOTURNO",
        "arquivo": Path("03 - GMP") / "GMP NOTURNO 2026.xlsx",
        "aba_horario": "GMP02",
    },
    {
        "turma": "R3",
        "turno": "NOTURNO",
        "arquivo": Path("04 - BASICO") / "BASICO Noturno 2026.xlsx",
        "aba_horario": "R3",
    },
    {
        "turma": "S3",
        "turno": "NOTURNO",
        "arquivo": Path("04 - BASICO") / "BASICO Noturno 2026.xlsx",
        "aba_horario": "S3",
    },
    {
        "turma": "T3",
        "turno": "NOTURNO",
        "arquivo": Path("04 - BASICO") / "BASICO Noturno 2026.xlsx",
        "aba_horario": "T3",
    },
    {
        "turma": "U3",
        "turno": "VESPERTINO",
        "arquivo": Path("04 - BASICO") / "BASICO Vespertino 2026.xlsx",
        "aba_horario": "U3",
    },
]

MESES_PT = {
    1: "JANEIRO",
    2: "FEVEREIRO",
    3: "MARÇO",
    4: "ABRIL",
    5: "MAIO",
    6: "JUNHO",
    7: "JULHO",
    8: "AGOSTO",
    9: "SETEMBRO",
    10: "OUTUBRO",
    11: "NOVEMBRO",
    12: "DEZEMBRO",
}

# ============================================================
# EXCEÇÕES DE ATRIBUIÇÃO DE PROFESSOR
# ============================================================

EXCECOES_PROFESSOR = {
    ("S3", "PSO"): [
        "CLÁUDIA CRISTINA AROUCA NORBERTO",
    ],
}
