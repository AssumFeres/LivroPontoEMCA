from __future__ import annotations

import calendar
import re
import shutil
import sqlite3
from collections import defaultdict
from copy import deepcopy
from datetime import date, datetime
from io import BytesIO
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.drawing.image import (
    Image as ExcelImage,
)

from config import (
    CAMINHO_BANCO,
    MESES_PT,
    PASTA_EXPORTACOES,
    PASTA_PROJETO,
)


# ============================================================
# CONFIGURAÇÃO DO MODELO
# ============================================================

PASTA_MODELOS = (
    PASTA_PROJETO
    / "modelos"
)

CAMINHO_MODELO_PONTO = (
    PASTA_MODELOS
    / "Modelo_de_Ponto_Mensal_Instrutor.xlsx"
)

ABA_MODELO = "Prof. (2)"

LINHA_INICIAL_DIAS = 9
LINHA_FINAL_DIAS = 31

COLUNAS_VESPERTINO = {
    "aula_01": (
        "C",
        "D",
        "E",
    ),

    "aula_02": (
        "F",
        "G",
        "H",
    ),
}

COLUNAS_NOTURNO = {
    "aula_01": (
        "I",
        "J",
        "K",
    ),

    "aula_02": (
        "L",
        "M",
        "N",
    ),
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

def dias_uteis_mes(
    ano: int,
    mes: int,
) -> list[date]:
    ultimo_dia = calendar.monthrange(
        ano,
        mes,
    )[1]

    return [
        date(
            ano,
            mes,
            dia,
        )
        for dia in range(
            1,
            ultimo_dia + 1,
        )
        if date(
            ano,
            mes,
            dia,
        ).weekday() <= 4
    ]


def normalizar_turno(
    turno: str | None,
    turma: str,
) -> str:
    if turno:
        turno_normalizado = (
            turno
            .strip()
            .upper()
        )

        if turno_normalizado in {
            "VESPERTINO",
            "NOTURNO",
        }:
            return turno_normalizado

    if turma.strip().upper() == "U3":
        return "VESPERTINO"

    return "NOTURNO"


def sigla_valida(
    valor: str | None,
) -> str | None:
    if valor is None:
        return None

    texto = str(
        valor
    ).strip()

    return texto or None


def formatar_sigla_livro(
    sigla: str | None,
    segundo_instrutor: bool,
) -> str | None:
    """
    No banco a disciplina permanece sem asterisco.

    O asterisco é incluído somente na apresentação
    do Livro Ponto quando o professor é o segundo instrutor.

    Exemplo:
        SEL   -> primeiro instrutor
        SEL * -> segundo instrutor
    """

    sigla = sigla_valida(
        sigla
    )

    if not sigla:
        return None

    if segundo_instrutor:
        return f"{sigla} *"

    return sigla


def juntar_siglas(
    siglas: list[str],
) -> str | None:
    unicas: list[str] = []

    for sigla in siglas:
        sigla = sigla.strip()

        if (
            sigla
            and sigla not in unicas
        ):
            unicas.append(
                sigla
            )

    if not unicas:
        return None

    return " / ".join(
        unicas
    )


def preencher_tres_colunas(
    ws,
    linha: int,
    colunas: tuple[
        str,
        str,
        str,
    ],
    valor: str | None,
) -> None:
    for coluna in colunas:
        ws[
            f"{coluna}{linha}"
        ] = valor


def limpar_area_dias(
    ws,
) -> None:
    for linha in range(
        LINHA_INICIAL_DIAS,
        LINHA_FINAL_DIAS + 1,
    ):
        for coluna in range(
            1,
            17,
        ):
            ws.cell(
                row=linha,
                column=coluna,
            ).value = None


def obter_colunas_tabela(
    conexao: sqlite3.Connection,
    tabela: str,
) -> set[str]:
    linhas = conexao.execute(
        f"PRAGMA table_info({tabela})"
    ).fetchall()

    return {
        str(linha[1])
        for linha in linhas
    }


def obter_coluna_professor(
    conexao: sqlite3.Connection,
) -> str:
    colunas = obter_colunas_tabela(
        conexao,
        "aulas",
    )

    for candidato in (
        "professor",
        "nome_completo",
        "nome_professor",
        "instrutor",
    ):
        if candidato in colunas:
            return candidato

    raise RuntimeError(
        "Não foi encontrada na tabela 'aulas' "
        "uma coluna de professor. "
        "Colunas existentes: "
        f"{', '.join(sorted(colunas))}"
    )


def obter_coluna_turno(
    conexao: sqlite3.Connection,
) -> str | None:
    colunas = obter_colunas_tabela(
        conexao,
        "aulas",
    )

    if "turno" in colunas:
        return "turno"

    return None


def obter_professores_mes(
    conexao: sqlite3.Connection,
    ano: int,
    mes: int,
) -> list[str]:
    inicio = (
        f"{ano:04d}-"
        f"{mes:02d}-01"
    )

    if mes == 12:
        proximo = (
            f"{ano + 1:04d}-"
            "01-01"
        )
    else:
        proximo = (
            f"{ano:04d}-"
            f"{mes + 1:02d}-01"
        )

    coluna_professor = (
        obter_coluna_professor(
            conexao
        )
    )

    sql = f"""
        SELECT DISTINCT
            "{coluna_professor}"

        FROM aulas

        WHERE data >= ?
          AND data < ?

          AND "{coluna_professor}"
              IS NOT NULL

          AND TRIM(
              "{coluna_professor}"
          ) <> ''

        ORDER BY
            "{coluna_professor}"
    """

    linhas = conexao.execute(
        sql,
        (
            inicio,
            proximo,
        ),
    ).fetchall()

    return [
        str(
            linha[0]
        ).strip()
        for linha in linhas
    ]


def obter_aulas_professor(
    conexao: sqlite3.Connection,
    professor: str,
    ano: int,
    mes: int,
) -> list[sqlite3.Row]:
    inicio = (
        f"{ano:04d}-"
        f"{mes:02d}-01"
    )

    if mes == 12:
        proximo = (
            f"{ano + 1:04d}-"
            "01-01"
        )
    else:
        proximo = (
            f"{ano:04d}-"
            f"{mes + 1:02d}-01"
        )

    coluna_professor = (
        obter_coluna_professor(
            conexao
        )
    )

    colunas = obter_colunas_tabela(
        conexao,
        "aulas",
    )

    if "turno" in colunas:
        expressao_turno = '"turno"'
    else:
        expressao_turno = "NULL"

    # Compatibilidade com banco ainda não migrado.
    if (
        "segundo_instrutor_aula_01"
        in colunas
    ):
        expressao_segundo_01 = (
            '"segundo_instrutor_aula_01"'
        )
    else:
        expressao_segundo_01 = "0"

    if (
        "segundo_instrutor_aula_02"
        in colunas
    ):
        expressao_segundo_02 = (
            '"segundo_instrutor_aula_02"'
        )
    else:
        expressao_segundo_02 = "0"

    sql = f"""
        SELECT
            turma,

            {expressao_turno}
                AS turno,

            data,

            "{coluna_professor}"
                AS professor,

            aula_01,
            aula_02,

            {expressao_segundo_01}
                AS segundo_instrutor_aula_01,

            {expressao_segundo_02}
                AS segundo_instrutor_aula_02

        FROM aulas

        WHERE "{coluna_professor}" = ?
          AND data >= ?
          AND data < ?

        ORDER BY
            data,
            turma
    """

    return conexao.execute(
        sql,
        (
            professor,
            inicio,
            proximo,
        ),
    ).fetchall()


def obter_resumo_professor(
    aulas: list[sqlite3.Row],
) -> tuple[str, str]:
    disciplinas: set[str] = set()
    turmas: set[str] = set()

    for aula in aulas:
        turmas.add(
            str(
                aula["turma"]
            ).strip()
        )

        aula_01 = sigla_valida(
            aula["aula_01"]
        )

        aula_02 = sigla_valida(
            aula["aula_02"]
        )

        if aula_01:
            disciplinas.add(
                aula_01
            )

        if aula_02:
            disciplinas.add(
                aula_02
            )

    disciplinas_texto = ", ".join(
        sorted(
            disciplinas
        )
    )

    turmas_texto = ", ".join(
        sorted(
            turmas
        )
    )

    return (
        disciplinas_texto,
        turmas_texto,
    )


def agrupar_aulas_por_data(
    aulas: list[sqlite3.Row],
) -> dict:
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
        data_aula = str(
            aula["data"]
        )

        turma = str(
            aula["turma"]
        )

        turno = normalizar_turno(
            aula["turno"],
            turma,
        )

        aula_01 = formatar_sigla_livro(
            aula["aula_01"],
            bool(
                aula[
                    "segundo_instrutor_aula_01"
                ]
            ),
        )

        aula_02 = formatar_sigla_livro(
            aula["aula_02"],
            bool(
                aula[
                    "segundo_instrutor_aula_02"
                ]
            ),
        )

        if aula_01:
            dados[
                data_aula
            ][
                turno
            ][
                "aula_01"
            ].append(
                aula_01
            )

        if aula_02:
            dados[
                data_aula
            ][
                turno
            ][
                "aula_02"
            ].append(
                aula_02
            )

    return dict(
        dados
    )


def nome_aba_professor(
    nome: str,
    existentes: set[str],
) -> str:
    nome = re.sub(
        r'[\\/*?:\[\]]',
        "",
        nome.strip(),
    )

    nome = re.sub(
        r"\s+",
        " ",
        nome,
    )

    if len(nome) > 31:
        nome = nome[:31]

    base = (
        nome
        or "Professor"
    )

    candidato = base
    contador = 2

    while candidato in existentes:
        sufixo = (
            f"_{contador}"
        )

        candidato = (
            f"{base[:31-len(sufixo)]}"
            f"{sufixo}"
        )

        contador += 1

    return candidato


# ============================================================
# IMAGENS DO MODELO
# ============================================================

def capturar_imagens_modelo(
    ws_modelo,
) -> list[dict]:
    """
    Lê as imagens do modelo UMA ÚNICA VEZ.

    Isso evita o erro:
        I/O operation on closed file
    """

    imagens_cache: list[dict] = []

    for imagem in getattr(
        ws_modelo,
        "_images",
        [],
    ):
        dados = imagem._data()

        imagens_cache.append(
            {
                "dados": dados,

                "width":
                    imagem.width,

                "height":
                    imagem.height,

                "anchor":
                    deepcopy(
                        imagem.anchor
                    ),
            }
        )

    return imagens_cache


def aplicar_imagens_na_planilha(
    ws_destino,
    imagens_cache: list[dict],
) -> None:
    """
    Cria uma cópia independente das imagens
    em cada aba de professor.
    """

    for item in imagens_cache:
        fluxo = BytesIO(
            item["dados"]
        )

        nova_imagem = ExcelImage(
            fluxo
        )

        nova_imagem.width = (
            item["width"]
        )

        nova_imagem.height = (
            item["height"]
        )

        nova_imagem.anchor = deepcopy(
            item["anchor"]
        )

        ws_destino.add_image(
            nova_imagem
        )


# ============================================================
# PREENCHIMENTO DO LIVRO PONTO
# ============================================================

def preencher_planilha_professor(
    ws,
    professor: str,
    ano: int,
    mes: int,
    aulas: list[sqlite3.Row],
) -> None:
    (
        disciplinas,
        turmas,
    ) = obter_resumo_professor(
        aulas
    )

    ws["B3"] = professor

    ws["B4"] = datetime(
        ano,
        mes,
        1,
    )

    ws["B4"].number_format = (
        "mmmm/yyyy"
    )

    ws["B5"] = disciplinas
    ws["O5"] = turmas

    limpar_area_dias(
        ws
    )

    dias = dias_uteis_mes(
        ano,
        mes,
    )

    capacidade = (
        LINHA_FINAL_DIAS
        - LINHA_INICIAL_DIAS
        + 1
    )

    if len(dias) > capacidade:
        raise ValueError(
            "O modelo possui espaço para "
            f"{capacidade} dias úteis, "
            f"mas {MESES_PT[mes]}/{ano} "
            f"possui {len(dias)}."
        )

    aulas_por_data = (
        agrupar_aulas_por_data(
            aulas
        )
    )

    for (
        indice,
        data_atual,
    ) in enumerate(
        dias
    ):
        linha = (
            LINHA_INICIAL_DIAS
            + indice
        )

        ws[
            f"A{linha}"
        ] = datetime(
            data_atual.year,
            data_atual.month,
            data_atual.day,
        )

        ws[
            f"A{linha}"
        ].number_format = (
            "dd/mm/yyyy"
        )

        ws[
            f"B{linha}"
        ] = DIAS_SEMANA_PT[
            data_atual.weekday()
        ]

        dados_dia = (
            aulas_por_data.get(
                data_atual.isoformat()
            )
        )

        if not dados_dia:
            continue

        vespertino_01 = juntar_siglas(
            dados_dia[
                "VESPERTINO"
            ][
                "aula_01"
            ]
        )

        vespertino_02 = juntar_siglas(
            dados_dia[
                "VESPERTINO"
            ][
                "aula_02"
            ]
        )

        noturno_01 = juntar_siglas(
            dados_dia[
                "NOTURNO"
            ][
                "aula_01"
            ]
        )

        noturno_02 = juntar_siglas(
            dados_dia[
                "NOTURNO"
            ][
                "aula_02"
            ]
        )

        preencher_tres_colunas(
            ws,
            linha,
            COLUNAS_VESPERTINO[
                "aula_01"
            ],
            vespertino_01,
        )

        preencher_tres_colunas(
            ws,
            linha,
            COLUNAS_VESPERTINO[
                "aula_02"
            ],
            vespertino_02,
        )

        preencher_tres_colunas(
            ws,
            linha,
            COLUNAS_NOTURNO[
                "aula_01"
            ],
            noturno_01,
        )

        preencher_tres_colunas(
            ws,
            linha,
            COLUNAS_NOTURNO[
                "aula_02"
            ],
            noturno_02,
        )

        # Extracurricular:
        # marca X em todo dia em que houver aula.
        if any(
            (
                vespertino_01,
                vespertino_02,
                noturno_01,
                noturno_02,
            )
        ):
            ws[
                f"O{linha}"
            ] = "X"

    # Assinatura sempre fica em branco.
    for linha in range(
        LINHA_INICIAL_DIAS,
        LINHA_FINAL_DIAS + 1,
    ):
        ws[
            f"P{linha}"
        ] = None

    ws.sheet_view.showGridLines = False

    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1

    ws.sheet_properties.pageSetUpPr.fitToPage = (
        True
    )


# ============================================================
# GERAÇÃO DO ARQUIVO CONSOLIDADO
# ============================================================

def gerar_livros_ponto(
    ano: int,
    mes: int,
    caminho_banco: str | Path = CAMINHO_BANCO,
    caminho_modelo: str | Path = CAMINHO_MODELO_PONTO,
) -> list[Path]:
    """
    Gera UM arquivo mensal com uma aba
    para cada professor.
    """

    if not 1 <= mes <= 12:
        raise ValueError(
            "Mês deve estar entre 1 e 12."
        )

    caminho_banco = Path(
        caminho_banco
    )

    caminho_modelo = Path(
        caminho_modelo
    )

    if not caminho_banco.exists():
        raise FileNotFoundError(
            "Banco de dados não encontrado: "
            f"{caminho_banco}"
        )

    if not caminho_modelo.exists():
        raise FileNotFoundError(
            "Modelo não encontrado: "
            f"{caminho_modelo}"
        )

    pasta_saida = (
        PASTA_EXPORTACOES
        / "livros_ponto"
    )

    pasta_saida.mkdir(
        parents=True,
        exist_ok=True,
    )

    caminho_saida = (
        pasta_saida
        / (
            f"LIVRO_PONTO_"
            f"{ano}_"
            f"{mes:02d}_"
            f"{MESES_PT[mes]}.xlsx"
        )
    )

    shutil.copy2(
        caminho_modelo,
        caminho_saida,
    )

    conexao = sqlite3.connect(
        caminho_banco
    )

    conexao.row_factory = (
        sqlite3.Row
    )

    try:
        professores = obter_professores_mes(
            conexao,
            ano,
            mes,
        )

        if not professores:
            raise ValueError(
                "Nenhum professor com aula "
                "encontrado em "
                f"{MESES_PT[mes]}/{ano}."
            )

        wb = load_workbook(
            caminho_saida
        )

        if (
            ABA_MODELO
            not in wb.sheetnames
        ):
            raise ValueError(
                f"A aba '{ABA_MODELO}' "
                "não foi encontrada no modelo."
            )

        ws_modelo = wb[
            ABA_MODELO
        ]

        # Lê as imagens uma única vez.
        imagens_cache = (
            capturar_imagens_modelo(
                ws_modelo
            )
        )

        # Mantém somente a folha modelo
        # antes de gerar as abas.
        for nome in list(
            wb.sheetnames
        ):
            if nome != ABA_MODELO:
                del wb[nome]

        existentes = set(
            wb.sheetnames
        )

        print(
            "\nGerando livro ponto "
            "consolidado de "
            f"{MESES_PT[mes]}/{ano}..."
        )

        print(
            "Professores encontrados: "
            f"{len(professores)}\n"
        )

        for (
            numero,
            professor,
        ) in enumerate(
            professores,
            start=1,
        ):
            aulas = obter_aulas_professor(
                conexao,
                professor,
                ano,
                mes,
            )

            ws = wb.copy_worksheet(
                ws_modelo
            )

            aplicar_imagens_na_planilha(
                ws,
                imagens_cache,
            )

            nome_aba = nome_aba_professor(
                professor,
                existentes,
            )

            ws.title = nome_aba

            existentes.add(
                nome_aba
            )

            preencher_planilha_professor(
                ws=ws,
                professor=professor,
                ano=ano,
                mes=mes,
                aulas=aulas,
            )

            print(
                f"[{numero:02d}/"
                f"{len(professores):02d}] "
                f"{professor}"
            )

        # Remove a folha utilizada como molde.
        del wb[
            ABA_MODELO
        ]

        wb.active = 0

        wb.save(
            caminho_saida
        )

        print(
            "\nArquivo gerado "
            "com sucesso em:"
        )

        print(
            caminho_saida
        )

        return [
            caminho_saida
        ]

    finally:
        conexao.close()


if __name__ == "__main__":
    ano = int(
        input(
            "Ano: "
        ).strip()
    )

    mes = int(
        input(
            "Mês (1-12): "
        ).strip()
    )

    arquivos = gerar_livros_ponto(
        ano=ano,
        mes=mes,
    )

    print(
        "\nTotal de arquivos gerados: "
        f"{len(arquivos)}"
    )
