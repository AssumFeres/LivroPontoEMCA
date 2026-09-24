from __future__ import annotations

import re
from collections import defaultdict
from datetime import date, datetime

from openpyxl import load_workbook

from banco import conectar, criar_banco
from config import (
    ANO_PADRAO,
    CAMINHO_BANCO,
    EXCECOES_PROFESSOR,
    PASTA_HORARIOS,
    TURMAS,
)


SEPARADOR_PROFESSORES = re.compile(r"\s*/\s*")


def normalizar_texto(valor) -> str:
    if valor is None:
        return ""

    return str(valor).strip()


def normalizar_sigla(valor) -> str:
    texto = normalizar_texto(valor).upper()

    if texto in {
        "",
        "0",
        "0.0",
        "NONE",
        "NAN",
    }:
        return ""

    return texto


def dividir_professores(
    nome_completo: str,
) -> list[tuple[str, bool]]:
    """
    A coluna E da aba Base pode trazer professores separados por "/".

    Exemplo:

        RODNEY FAGUNDES DOS SANTOS /
        MANUEL DE LEMOS GASPAR

    Regra:
        - nome antes da "/" = primeiro instrutor;
        - nome depois da "/" = segundo instrutor.

    Retorno:

        [
            ("RODNEY FAGUNDES DOS SANTOS", False),
            ("MANUEL DE LEMOS GASPAR", True),
        ]

    O booleano True indica SEGUNDO INSTRUTOR.
    """

    partes = [
        parte.strip()
        for parte in SEPARADOR_PROFESSORES.split(
            nome_completo or ""
        )
    ]

    resultado: list[tuple[str, bool]] = []
    vistos: set[str] = set()

    for indice, parte in enumerate(partes):
        if not parte:
            continue

        chave = parte.casefold()

        if chave in vistos:
            continue

        vistos.add(chave)

        segundo_instrutor = indice > 0

        resultado.append(
            (
                parte,
                segundo_instrutor,
            )
        )

    return resultado


def localizar_aba(
    workbook,
    nome_desejado: str,
):
    """
    Localiza a aba ignorando espaços extras
    e diferença entre maiúsculas/minúsculas.
    """

    procurado = nome_desejado.strip().casefold()

    for nome_real in workbook.sheetnames:
        if (
            nome_real.strip().casefold()
            == procurado
        ):
            return workbook[nome_real]

    raise KeyError(
        f"Aba '{nome_desejado}' não encontrada. "
        f"Abas disponíveis: {workbook.sheetnames}"
    )


def construir_dicionario_base(
    workbook,
) -> dict[str, list[tuple[str, bool]]]:
    """
    Exemplo de retorno:

        {
            "MAT": [
                ("MARCOS CESAR FARIA", False),
            ],

            "OTB": [
                ("RODNEY FAGUNDES DOS SANTOS", False),
                ("MANUEL DE LEMOS GASPAR", True),
            ],
        }

    A aba Base utiliza:

        Coluna A = Código da Disciplina
        Coluna E = Nome Completo
    """

    ws = localizar_aba(
        workbook,
        "Base",
    )

    dicionario: dict[
        str,
        list[tuple[str, bool]],
    ] = {}

    for linha in ws.iter_rows(
        min_row=3,
        values_only=True,
    ):
        codigo = normalizar_sigla(
            linha[0]
            if len(linha) > 0
            else None
        )

        nome = normalizar_texto(
            linha[4]
            if len(linha) > 4
            else None
        )

        if not codigo or not nome:
            continue

        professores = dividir_professores(
            nome
        )

        if professores:
            dicionario[codigo] = professores

    return dicionario


def obter_professores_disciplina(
    turma: str,
    sigla: str,
    dicionario_professores: dict[
        str,
        list[tuple[str, bool]],
    ],
) -> list[tuple[str, bool]]:
    """
    Primeiro verifica exceções manuais.
    Depois utiliza o dicionário da aba Base.

    Professores definidos nas exceções são considerados,
    por padrão, PRIMEIROS INSTRUTORES.
    """

    turma = turma.strip().upper()
    sigla = sigla.strip().upper()

    chave_excecao = (
        turma,
        sigla,
    )

    if chave_excecao in EXCECOES_PROFESSOR:
        resultado: list[
            tuple[str, bool]
        ] = []

        for item in EXCECOES_PROFESSOR[
            chave_excecao
        ]:
            # Permite futuramente usar:
            # ("NOME", True)
            if (
                isinstance(item, tuple)
                and len(item) == 2
            ):
                nome = str(item[0]).strip()
                segundo = bool(item[1])
            else:
                nome = str(item).strip()
                segundo = False

            if nome:
                resultado.append(
                    (
                        nome,
                        segundo,
                    )
                )

        return resultado

    return dicionario_professores.get(
        sigla,
        [],
    )


def converter_data(
    valor,
) -> date | None:
    if valor is None:
        return None

    if isinstance(
        valor,
        datetime,
    ):
        return valor.date()

    if isinstance(
        valor,
        date,
    ):
        return valor

    return None


def registrar_log(
    conn,
    cfg: dict,
    tipo: str,
    mensagem: str,
) -> None:
    conn.execute(
        """
        INSERT INTO log_importacao (
            turma,
            arquivo,
            aba,
            tipo,
            mensagem
        )
        VALUES (?, ?, ?, ?, ?)
        """,
        (
            cfg["turma"],
            str(cfg["arquivo"]),
            cfg["aba_horario"],
            tipo,
            mensagem,
        ),
    )


def importar_turma(
    conn,
    cfg: dict,
    ano: int = ANO_PADRAO,
) -> int:
    caminho_arquivo = (
        PASTA_HORARIOS
        / cfg["arquivo"]
    )

    if not caminho_arquivo.exists():
        registrar_log(
            conn,
            cfg,
            "ERRO",
            (
                "Arquivo não encontrado: "
                f"{caminho_arquivo}"
            ),
        )

        print(
            "[ERRO] Arquivo não encontrado: "
            f"{caminho_arquivo}"
        )

        return 0

    # data_only=True é indispensável.
    # AULA 01 e AULA 02 são fórmulas nas planilhas.
    wb = load_workbook(
        caminho_arquivo,
        data_only=True,
        read_only=True,
    )

    try:
        base = construir_dicionario_base(
            wb
        )

        ws = localizar_aba(
            wb,
            cfg["aba_horario"],
        )

        # Remove somente os registros da turma/ano
        # que será reimportada.
        conn.execute(
            """
            DELETE FROM aulas
             WHERE turma = ?
               AND substr(data, 1, 4) = ?
            """,
            (
                cfg["turma"],
                str(ano),
            ),
        )

        quantidade = 0

        for linha in ws.iter_rows(
            min_row=3,
            min_col=1,
            max_col=3,
            values_only=True,
        ):
            data_aula = converter_data(
                linha[0]
            )

            if (
                data_aula is None
                or data_aula.year != ano
            ):
                continue

            sigla_01 = normalizar_sigla(
                linha[1]
            )

            sigla_02 = normalizar_sigla(
                linha[2]
            )

            # Um mesmo professor pode participar
            # de Aula 01 e Aula 02 na mesma data/turma.
            registros = defaultdict(
                lambda: {
                    "aula_01": "",
                    "aula_02": "",

                    "segundo_instrutor_aula_01": 0,
                    "segundo_instrutor_aula_02": 0,
                }
            )

            # --------------------------------------------
            # AULA 01
            # --------------------------------------------

            if sigla_01:
                professores = (
                    obter_professores_disciplina(
                        cfg["turma"],
                        sigla_01,
                        base,
                    )
                )

                if professores:
                    for (
                        professor,
                        segundo_instrutor,
                    ) in professores:
                        registros[
                            professor
                        ]["aula_01"] = sigla_01

                        registros[
                            professor
                        ][
                            "segundo_instrutor_aula_01"
                        ] = (
                            1
                            if segundo_instrutor
                            else 0
                        )

                else:
                    registrar_log(
                        conn,
                        cfg,
                        "SIGLA_NAO_ENCONTRADA",
                        (
                            f"{data_aula:%d/%m/%Y} "
                            f"- Aula 01: "
                            f"'{sigla_01}' "
                            "não existe na aba Base."
                        ),
                    )

            # --------------------------------------------
            # AULA 02
            # --------------------------------------------

            if sigla_02:
                professores = (
                    obter_professores_disciplina(
                        cfg["turma"],
                        sigla_02,
                        base,
                    )
                )

                if professores:
                    for (
                        professor,
                        segundo_instrutor,
                    ) in professores:
                        registros[
                            professor
                        ]["aula_02"] = sigla_02

                        registros[
                            professor
                        ][
                            "segundo_instrutor_aula_02"
                        ] = (
                            1
                            if segundo_instrutor
                            else 0
                        )

                else:
                    registrar_log(
                        conn,
                        cfg,
                        "SIGLA_NAO_ENCONTRADA",
                        (
                            f"{data_aula:%d/%m/%Y} "
                            f"- Aula 02: "
                            f"'{sigla_02}' "
                            "não existe na aba Base."
                        ),
                    )

            # --------------------------------------------
            # GRAVAÇÃO NO BANCO
            # --------------------------------------------

            for (
                professor,
                aulas,
            ) in registros.items():
                conn.execute(
                    """
                    INSERT INTO aulas (
                        turma,
                        turno,
                        data,
                        nome_completo,
                        aula_01,
                        aula_02,
                        segundo_instrutor_aula_01,
                        segundo_instrutor_aula_02,
                        arquivo_origem,
                        aba_origem
                    )
                    VALUES (
                        ?, ?, ?, ?, ?, ?,
                        ?, ?, ?, ?
                    )

                    ON CONFLICT (
                        turma,
                        data,
                        nome_completo
                    )

                    DO UPDATE SET
                        turno =
                            excluded.turno,

                        aula_01 =
                            excluded.aula_01,

                        aula_02 =
                            excluded.aula_02,

                        segundo_instrutor_aula_01 =
                            excluded.segundo_instrutor_aula_01,

                        segundo_instrutor_aula_02 =
                            excluded.segundo_instrutor_aula_02,

                        arquivo_origem =
                            excluded.arquivo_origem,

                        aba_origem =
                            excluded.aba_origem
                    """,
                    (
                        cfg["turma"],
                        cfg["turno"],
                        data_aula.isoformat(),
                        professor,
                        aulas["aula_01"],
                        aulas["aula_02"],
                        aulas[
                            "segundo_instrutor_aula_01"
                        ],
                        aulas[
                            "segundo_instrutor_aula_02"
                        ],
                        str(caminho_arquivo),
                        ws.title,
                    ),
                )

                quantidade += 1

        registrar_log(
            conn,
            cfg,
            "OK",
            (
                f"{quantidade} "
                "registros importados."
            ),
        )

        return quantidade

    finally:
        wb.close()


def importar_todos(
    ano: int = ANO_PADRAO,
) -> int:
    # criar_banco também executa a migração
    # necessária para bancos antigos.
    criar_banco(
        CAMINHO_BANCO,
        ano,
    )

    total = 0

    with conectar(
        CAMINHO_BANCO
    ) as conn:
        for cfg in TURMAS:
            quantidade = importar_turma(
                conn,
                cfg,
                ano,
            )

            total += quantidade

            print(
                f"{cfg['turma']}: "
                f"{quantidade} registros"
            )

    print(
        f"\nTotal importado: "
        f"{total} registros"
    )

    return total


if __name__ == "__main__":
    importar_todos()
