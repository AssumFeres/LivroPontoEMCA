from __future__ import annotations

import calendar
import re
import shutil
import sqlite3
import unicodedata
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as ExcelImage

from config import CAMINHO_BANCO, MESES_PT, PASTA_EXPORTACOES, PASTA_PROJETO


# ============================================================
# CONFIGURAÇÃO DO MODELO
# ============================================================

PASTA_MODELOS = PASTA_PROJETO / "modelos"
CAMINHO_MODELO_PONTO = (
    PASTA_MODELOS / "Modelo_de_Ponto_Mensal_Instrutor.xlsx"
)

CAMINHO_LOGO_FAPETI = PASTA_MODELOS / "logo_Fatepi.jpg"

ABA_MODELO = "Prof. (2)"

# No modelo novo:
# A = Data
# B = Dia da semana
# C:H = Vespertino
# I:N = Noturno
# O = Extracurricular
# P = Assinatura
LINHA_INICIAL_DIAS = 9
LINHA_FINAL_DIAS = 31

COLUNAS_VESPERTINO = {
    "aula_01": ("C", "D", "E"),
    "aula_02": ("F", "G", "H"),
}

COLUNAS_NOTURNO = {
    "aula_01": ("I", "J", "K"),
    "aula_02": ("L", "M", "N"),
}

DIAS_SEMANA_PT = {
    0: "Segunda-feira",
    1: "Terça-feira",
    2: "Quarta-feira",
    3: "Quinta-feira",
    4: "Sexta-feira",
    5: "Sábado",
    6: "Domingo",
}


# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================

def limpar_nome_arquivo(texto: str) -> str:
    """
    Remove caracteres inválidos para nomes de arquivo no Windows.
    """
    texto = texto.strip()
    texto = re.sub(r'[<>:"/\\|?*]', "", texto)
    texto = re.sub(r"\s+", " ", texto)
    return texto


def remover_acentos(texto: str) -> str:
    """
    Usado somente para gerar nomes de arquivos mais simples.
    """
    normalizado = unicodedata.normalize("NFKD", texto)
    return "".join(
        caractere
        for caractere in normalizado
        if not unicodedata.combining(caractere)
    )


def nome_arquivo_professor(professor: str, ano: int, mes: int) -> str:
    nome = remover_acentos(professor.upper())
    nome = limpar_nome_arquivo(nome)
    return f"LIVRO_PONTO_{ano}_{mes:02d}_{nome}.xlsx"


def dias_uteis_mes(ano: int, mes: int) -> list[date]:
    """
    Retorna somente segunda a sexta-feira.
    Feriados NÃO são retirados, conforme regra definida para o livro ponto.
    """
    ultimo_dia = calendar.monthrange(ano, mes)[1]

    resultado: list[date] = []

    for dia in range(1, ultimo_dia + 1):
        data_atual = date(ano, mes, dia)

        if data_atual.weekday() <= 4:
            resultado.append(data_atual)

    return resultado


def normalizar_turno(turno: str | None, turma: str) -> str:
    """
    Usa o turno gravado no banco.
    Como segurança, U3 é considerada VESPERTINO.
    As demais turmas são NOTURNO.
    """
    if turno:
        turno_normalizado = turno.strip().upper()

        if turno_normalizado in {"VESPERTINO", "NOTURNO"}:
            return turno_normalizado

    if turma.strip().upper() == "U3":
        return "VESPERTINO"

    return "NOTURNO"


def sigla_valida(valor: str | None) -> str | None:
    if valor is None:
        return None

    texto = str(valor).strip()

    if not texto:
        return None

    return texto


def juntar_siglas(siglas: list[str]) -> str | None:
    """
    Caso haja mais de uma disciplina no mesmo professor/data/turno/bloco,
    mantém todas sem sobrescrever informação.
    """
    unicas: list[str] = []

    for sigla in siglas:
        sigla = sigla.strip()

        if sigla and sigla not in unicas:
            unicas.append(sigla)

    if not unicas:
        return None

    return " / ".join(unicas)


def preencher_tres_colunas(
    ws,
    linha: int,
    colunas: tuple[str, str, str],
    valor: str | None,
) -> None:
    for coluna in colunas:
        ws[f"{coluna}{linha}"] = valor


def limpar_area_dias(ws) -> None:
    """
    Limpa somente os valores das linhas de calendário.
    A formatação do modelo é preservada.
    """
    for linha in range(LINHA_INICIAL_DIAS, LINHA_FINAL_DIAS + 1):
        for coluna in range(1, 17):  # A:P
            ws.cell(row=linha, column=coluna).value = None


def obter_colunas_tabela(
    conexao: sqlite3.Connection,
    tabela: str,
) -> set[str]:
    linhas = conexao.execute(
        f"PRAGMA table_info({tabela})"
    ).fetchall()

    return {str(linha[1]) for linha in linhas}


def obter_coluna_professor(
    conexao: sqlite3.Connection,
) -> str:
    """
    Compatibilidade com diferentes versões do banco.

    Aceita, em ordem:
    - professor
    - nome_completo
    - nome_professor
    - instrutor
    """
    colunas = obter_colunas_tabela(conexao, "aulas")

    candidatos = (
        "professor",
        "nome_completo",
        "nome_professor",
        "instrutor",
    )

    for candidato in candidatos:
        if candidato in colunas:
            return candidato

    raise RuntimeError(
        "Não foi encontrada na tabela 'aulas' uma coluna de professor. "
        f"Colunas existentes: {', '.join(sorted(colunas))}"
    )


def obter_coluna_turno(
    conexao: sqlite3.Connection,
) -> str | None:
    colunas = obter_colunas_tabela(conexao, "aulas")

    if "turno" in colunas:
        return "turno"

    return None


def obter_professores_mes(
    conexao: sqlite3.Connection,
    ano: int,
    mes: int,
) -> list[str]:
    inicio = f"{ano:04d}-{mes:02d}-01"

    if mes == 12:
        proximo = f"{ano + 1:04d}-01-01"
    else:
        proximo = f"{ano:04d}-{mes + 1:02d}-01"

    coluna_professor = obter_coluna_professor(conexao)

    sql = f"""
        SELECT DISTINCT "{coluna_professor}"
        FROM aulas
        WHERE data >= ?
          AND data < ?
          AND "{coluna_professor}" IS NOT NULL
          AND TRIM("{coluna_professor}") <> ''
        ORDER BY "{coluna_professor}"
    """

    linhas = conexao.execute(
        sql,
        (inicio, proximo),
    ).fetchall()

    return [str(linha[0]).strip() for linha in linhas]


def obter_aulas_professor(
    conexao: sqlite3.Connection,
    professor: str,
    ano: int,
    mes: int,
) -> list[sqlite3.Row]:
    inicio = f"{ano:04d}-{mes:02d}-01"

    if mes == 12:
        proximo = f"{ano + 1:04d}-01-01"
    else:
        proximo = f"{ano:04d}-{mes + 1:02d}-01"

    coluna_professor = obter_coluna_professor(conexao)
    coluna_turno = obter_coluna_turno(conexao)

    if coluna_turno:
        expressao_turno = '"turno"'
    else:
        # Para bancos antigos: o turno será deduzido pela turma.
        expressao_turno = "NULL"

    sql = f"""
        SELECT
            turma,
            {expressao_turno} AS turno,
            data,
            "{coluna_professor}" AS professor,
            aula_01,
            aula_02
        FROM aulas
        WHERE "{coluna_professor}" = ?
          AND data >= ?
          AND data < ?
        ORDER BY data, turma
    """

    return conexao.execute(
        sql,
        (professor, inicio, proximo),
    ).fetchall()


def obter_resumo_professor(
    aulas: list[sqlite3.Row],
) -> tuple[str, str]:
    disciplinas: set[str] = set()
    turmas: set[str] = set()

    for aula in aulas:
        turmas.add(str(aula["turma"]).strip())

        aula_01 = sigla_valida(aula["aula_01"])
        aula_02 = sigla_valida(aula["aula_02"])

        if aula_01:
            disciplinas.add(aula_01)

        if aula_02:
            disciplinas.add(aula_02)

    disciplinas_texto = ", ".join(sorted(disciplinas))
    turmas_texto = ", ".join(sorted(turmas))

    return disciplinas_texto, turmas_texto


def agrupar_aulas_por_data(
    aulas: list[sqlite3.Row],
) -> dict[str, dict[str, dict[str, list[str]]]]:
    """
    Estrutura gerada:

    {
        "2026-09-15": {
            "VESPERTINO": {
                "aula_01": ["OTB"],
                "aula_02": ["OTB"],
            },
            "NOTURNO": {
                "aula_01": ["SEL"],
                "aula_02": ["INS"],
            },
        }
    }
    """
    dados = defaultdict(
        lambda: {
            "VESPERTINO": {
                "aula_01": [],
                "aula_02": [],
            },
            "NOTURNO": {
                "aula_01": [],
                "aula_02": [],
            },
        }
    )

    for aula in aulas:
        data_aula = str(aula["data"])
        turma = str(aula["turma"])
        turno = normalizar_turno(aula["turno"], turma)

        aula_01 = sigla_valida(aula["aula_01"])
        aula_02 = sigla_valida(aula["aula_02"])

        if aula_01:
            dados[data_aula][turno]["aula_01"].append(aula_01)

        if aula_02:
            dados[data_aula][turno]["aula_02"].append(aula_02)

    return dict(dados)



def inserir_logo_fapeti(ws) -> None:
    """
    Insere o logo da FAPETI somente se a planilha ainda não tiver imagem.
    """

    # Se o modelo já preservou alguma imagem, não insere outra
    if ws._images:
        return

    if not CAMINHO_LOGO_FAPETI.exists():
        print(
            "[AVISO] Logo da FAPETI não encontrado em: "
            f"{CAMINHO_LOGO_FAPETI}"
        )
        return

    logo = ExcelImage(str(CAMINHO_LOGO_FAPETI))

    logo.width = 217
    logo.height = 85

    ws.add_image(logo, "I1")


# ============================================================
# GERAÇÃO DO ARQUIVO DE UM PROFESSOR
# ============================================================

def gerar_livro_professor(
    professor: str,
    ano: int,
    mes: int,
    aulas: list[sqlite3.Row],
    pasta_saida: Path,
    caminho_modelo: Path = CAMINHO_MODELO_PONTO,
) -> Path:
    if not caminho_modelo.exists():
        raise FileNotFoundError(
            "Modelo do livro ponto não encontrado em:\n"
            f"{caminho_modelo}\n\n"
            "Crie a pasta 'modelos' na raiz do projeto e coloque nela "
            "o arquivo com o nome:\n"
            "Modelo_de_Ponto_Mensal_Instrutor.xlsx"
        )

    pasta_saida.mkdir(parents=True, exist_ok=True)

    caminho_saida = pasta_saida / nome_arquivo_professor(
        professor,
        ano,
        mes,
    )

    shutil.copy2(caminho_modelo, caminho_saida)

    wb = load_workbook(caminho_saida)

    if ABA_MODELO not in wb.sheetnames:
        raise ValueError(
            f"A aba '{ABA_MODELO}' não foi encontrada no modelo."
        )

    ws = wb[ABA_MODELO]

    inserir_logo_fapeti(ws)

    # --------------------------------------------------------
    # CABEÇALHO
    # --------------------------------------------------------

    disciplinas, turmas = obter_resumo_professor(aulas)

    ws["B3"] = professor

    # Mantém valor de data real para o Excel.
    ws["B4"] = datetime(ano, mes, 1)
    ws["B4"].number_format = "mmmm/yyyy"

    ws["B5"] = disciplinas
    ws["O5"] = turmas

    # --------------------------------------------------------
    # LIMPEZA DA ÁREA DO MÊS
    # --------------------------------------------------------

    limpar_area_dias(ws)

    # --------------------------------------------------------
    # PREENCHIMENTO DOS DIAS ÚTEIS
    # --------------------------------------------------------

    dias = dias_uteis_mes(ano, mes)

    capacidade = LINHA_FINAL_DIAS - LINHA_INICIAL_DIAS + 1

    if len(dias) > capacidade:
        raise ValueError(
            f"O modelo possui espaço para {capacidade} dias úteis, "
            f"mas {MESES_PT[mes]}/{ano} possui {len(dias)}."
        )

    aulas_por_data = agrupar_aulas_por_data(aulas)

    for indice, data_atual in enumerate(dias):
        linha = LINHA_INICIAL_DIAS + indice

        ws[f"A{linha}"] = datetime(
            data_atual.year,
            data_atual.month,
            data_atual.day,
        )
        ws[f"A{linha}"].number_format = "dd/mm/yyyy"

        ws[f"B{linha}"] = DIAS_SEMANA_PT[data_atual.weekday()]

        chave_data = data_atual.isoformat()
        dados_dia = aulas_por_data.get(chave_data)

        if not dados_dia:
            continue

        # ---------------- VESPERTINO ----------------

        vespertino_01 = juntar_siglas(
            dados_dia["VESPERTINO"]["aula_01"]
        )
        vespertino_02 = juntar_siglas(
            dados_dia["VESPERTINO"]["aula_02"]
        )

        preencher_tres_colunas(
            ws,
            linha,
            COLUNAS_VESPERTINO["aula_01"],
            vespertino_01,
        )
        preencher_tres_colunas(
            ws,
            linha,
            COLUNAS_VESPERTINO["aula_02"],
            vespertino_02,
        )

        # ---------------- NOTURNO ----------------

        noturno_01 = juntar_siglas(
            dados_dia["NOTURNO"]["aula_01"]
        )
        noturno_02 = juntar_siglas(
            dados_dia["NOTURNO"]["aula_02"]
        )

        preencher_tres_colunas(
            ws,
            linha,
            COLUNAS_NOTURNO["aula_01"],
            noturno_01,
        )
        preencher_tres_colunas(
            ws,
            linha,
            COLUNAS_NOTURNO["aula_02"],
            noturno_02,
        )

        # ----------------------------------------------------
        # EXTRACURRICULAR
        # Preenche X sempre que houver qualquer aula no dia.
        # ----------------------------------------------------

        possui_aula = any(
            (
                vespertino_01,
                vespertino_02,
                noturno_01,
                noturno_02,
            )
        )

        if possui_aula:
            ws[f"O{linha}"] = "X"

    # --------------------------------------------------------
    # LIMPA A COLUNA DE ASSINATURA
    # --------------------------------------------------------

    for linha in range(LINHA_INICIAL_DIAS, LINHA_FINAL_DIAS + 1):
        ws[f"P{linha}"] = None

    # --------------------------------------------------------
    # ORGANIZA AS ABAS DO ARQUIVO FINAL
    # --------------------------------------------------------

    # Exclui a folha antiga do modelo, se existir.
    if "Prof." in wb.sheetnames and "Prof." != ABA_MODELO:
        del wb["Prof."]

    # Renomeia a folha nova para um nome mais simples.
    ws.title = "Ponto"

    # Deixa Ponto como aba ativa.
    wb.active = wb.sheetnames.index("Ponto")

    wb.save(caminho_saida)

    return caminho_saida


# ============================================================
# GERAÇÃO DE TODOS OS PROFESSORES DO MÊS
# ============================================================

def gerar_livros_ponto(
    ano: int,
    mes: int,
    caminho_banco: str | Path = CAMINHO_BANCO,
    caminho_modelo: str | Path = CAMINHO_MODELO_PONTO,
) -> list[Path]:
    if not 1 <= mes <= 12:
        raise ValueError("Mês deve estar entre 1 e 12.")

    caminho_banco = Path(caminho_banco)
    caminho_modelo = Path(caminho_modelo)

    if not caminho_banco.exists():
        raise FileNotFoundError(
            f"Banco de dados não encontrado: {caminho_banco}"
        )

    if not caminho_modelo.exists():
        raise FileNotFoundError(
            f"Modelo não encontrado: {caminho_modelo}"
        )

    pasta_saida = (
        PASTA_EXPORTACOES
        / "livros_ponto"
        / f"{ano}_{mes:02d}_{MESES_PT[mes]}"
    )

    pasta_saida.mkdir(parents=True, exist_ok=True)

    conexao = sqlite3.connect(caminho_banco)
    conexao.row_factory = sqlite3.Row

    try:
        professores = obter_professores_mes(
            conexao,
            ano,
            mes,
        )

        if not professores:
            raise ValueError(
                f"Nenhum professor com aula encontrado em "
                f"{MESES_PT[mes]}/{ano}."
            )

        arquivos_gerados: list[Path] = []

        print(
            f"\nGerando livros ponto de "
            f"{MESES_PT[mes]}/{ano}..."
        )

        print(f"Professores encontrados: {len(professores)}\n")

        for numero, professor in enumerate(
            professores,
            start=1,
        ):
            aulas = obter_aulas_professor(
                conexao,
                professor,
                ano,
                mes,
            )

            caminho = gerar_livro_professor(
                professor=professor,
                ano=ano,
                mes=mes,
                aulas=aulas,
                pasta_saida=pasta_saida,
                caminho_modelo=caminho_modelo,
            )

            arquivos_gerados.append(caminho)

            print(
                f"[{numero:02d}/{len(professores):02d}] "
                f"{professor}"
            )

        print(
            "\nLivros ponto gerados com sucesso em:\n"
            f"{pasta_saida}"
        )

        return arquivos_gerados

    finally:
        conexao.close()


# ============================================================
# EXECUÇÃO DIRETA
# ============================================================

if __name__ == "__main__":
    ano = int(input("Ano: ").strip())
    mes = int(input("Mês (1-12): ").strip())

    arquivos = gerar_livros_ponto(
        ano=ano,
        mes=mes,
    )

    print(f"\nTotal de arquivos gerados: {len(arquivos)}")
