from __future__ import annotations

import sqlite3
from datetime import date, timedelta
from pathlib import Path

from config import ANO_PADRAO, CAMINHO_BANCO, TURMAS


def conectar(caminho_banco: str | Path = CAMINHO_BANCO) -> sqlite3.Connection:
    caminho = Path(caminho_banco)
    caminho.parent.mkdir(parents=True, exist_ok=True)
    conexao = sqlite3.connect(caminho)
    conexao.execute("PRAGMA foreign_keys = ON")
    return conexao


def coluna_existe(conn: sqlite3.Connection, tabela: str, coluna: str) -> bool:
    colunas = conn.execute(f"PRAGMA table_info({tabela})").fetchall()
    return any(item[1] == coluna for item in colunas)


def migrar_banco_existente(conn: sqlite3.Connection) -> None:
    """Atualiza bancos criados por versões anteriores do projeto."""

    tabelas = {
        row[0]
        for row in conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table'"
        ).fetchall()
    }

    # Versões anteriores não possuíam a coluna turno.
    if "calendario_turma" in tabelas and not coluna_existe(
        conn, "calendario_turma", "turno"
    ):
        conn.execute("ALTER TABLE calendario_turma ADD COLUMN turno TEXT")

    if "aulas" in tabelas and not coluna_existe(conn, "aulas", "turno"):
        conn.execute("ALTER TABLE aulas ADD COLUMN turno TEXT")

    # Preenche o turno usando a configuração atual das turmas.
    for cfg in TURMAS:
        if "calendario_turma" in tabelas:
            conn.execute(
                """
                UPDATE calendario_turma
                   SET turno = ?
                 WHERE turma = ?
                   AND (turno IS NULL OR TRIM(turno) = '')
                """,
                (cfg["turno"], cfg["turma"]),
            )

        if "aulas" in tabelas:
            conn.execute(
                """
                UPDATE aulas
                   SET turno = ?
                 WHERE turma = ?
                   AND (turno IS NULL OR TRIM(turno) = '')
                """,
                (cfg["turno"], cfg["turma"]),
            )


def criar_banco(
    caminho_banco: str | Path = CAMINHO_BANCO,
    ano: int = ANO_PADRAO,
) -> Path:
    caminho = Path(caminho_banco)
    caminho.parent.mkdir(parents=True, exist_ok=True)

    with conectar(caminho) as conn:
        # Primeiro cria as tabelas que ainda não existem.
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS calendario_turma (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                turma TEXT NOT NULL,
                turno TEXT NOT NULL,
                data TEXT NOT NULL,
                UNIQUE (turma, data)
            );

            CREATE TABLE IF NOT EXISTS aulas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                turma TEXT NOT NULL,
                turno TEXT NOT NULL,
                data TEXT NOT NULL,
                nome_completo TEXT NOT NULL,
                aula_01 TEXT,
                aula_02 TEXT,
                arquivo_origem TEXT NOT NULL,
                aba_origem TEXT NOT NULL,
                UNIQUE (turma, data, nome_completo)
            );

            CREATE INDEX IF NOT EXISTS idx_aulas_data
                ON aulas(data);

            CREATE INDEX IF NOT EXISTS idx_aulas_professor
                ON aulas(nome_completo);

            CREATE TABLE IF NOT EXISTS resumo_professor_mensal (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ano INTEGER NOT NULL,
                mes INTEGER NOT NULL,
                nome_completo TEXT NOT NULL,
                disciplinas TEXT NOT NULL,
                turmas TEXT NOT NULL,
                UNIQUE (ano, mes, nome_completo)
            );

            CREATE TABLE IF NOT EXISTS log_importacao (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                data_hora TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                turma TEXT,
                arquivo TEXT,
                aba TEXT,
                tipo TEXT NOT NULL,
                mensagem TEXT NOT NULL
            );
            """
        )

        # Depois migra tabelas antigas que já existiam com outra estrutura.
        migrar_banco_existente(conn)

        inicio = date(ano, 1, 1)
        fim = date(ano, 12, 31)

        for cfg in TURMAS:
            atual = inicio
            while atual <= fim:
                conn.execute(
                    """
                    INSERT OR IGNORE INTO calendario_turma
                        (turma, turno, data)
                    VALUES (?, ?, ?)
                    """,
                    (cfg["turma"], cfg["turno"], atual.isoformat()),
                )
                atual += timedelta(days=1)

        conn.commit()

    return caminho


if __name__ == "__main__":
    banco = criar_banco()
    print(f"Banco criado/atualizado em: {banco}")
