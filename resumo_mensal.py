from __future__ import annotations

from collections import defaultdict

from banco import conectar
from config import CAMINHO_BANCO


def atualizar_resumo_mensal(ano: int, mes: int) -> int:
    inicio = f"{ano:04d}-{mes:02d}-01"
    if mes == 12:
        fim = f"{ano + 1:04d}-01-01"
    else:
        fim = f"{ano:04d}-{mes + 1:02d}-01"

    with conectar(CAMINHO_BANCO) as conn:
        linhas = conn.execute(
            """
            SELECT nome_completo, turma, aula_01, aula_02
            FROM aulas
            WHERE data >= ? AND data < ?
            ORDER BY nome_completo, data, turma
            """,
            (inicio, fim),
        ).fetchall()

        resumo = defaultdict(lambda: {"disciplinas": set(), "turmas": set()})

        for professor, turma, aula_01, aula_02 in linhas:
            if aula_01:
                resumo[professor]["disciplinas"].add(aula_01)
            if aula_02:
                resumo[professor]["disciplinas"].add(aula_02)
            resumo[professor]["turmas"].add(turma)

        conn.execute(
            "DELETE FROM resumo_professor_mensal WHERE ano = ? AND mes = ?",
            (ano, mes),
        )

        for professor in sorted(resumo):
            disciplinas = ", ".join(sorted(resumo[professor]["disciplinas"]))
            turmas = ", ".join(sorted(resumo[professor]["turmas"]))

            conn.execute(
                """
                INSERT INTO resumo_professor_mensal (
                    ano, mes, nome_completo, disciplinas, turmas
                )
                VALUES (?, ?, ?, ?, ?)
                """,
                (ano, mes, professor, disciplinas, turmas),
            )

    return len(resumo)


if __name__ == "__main__":
    ano = int(input("Ano: "))
    mes = int(input("Mês (1-12): "))
    qtd = atualizar_resumo_mensal(ano, mes)
    print(f"Resumo atualizado: {qtd} professores.")
