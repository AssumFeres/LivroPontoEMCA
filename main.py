from banco import criar_banco
from exportar_excel import exportar_mes
from importar_horarios import importar_todos
from resumo_mensal import atualizar_resumo_mensal


def ler_ano_mes() -> tuple[int, int]:
    ano = int(input("Ano: ").strip())
    mes = int(input("Mês (1-12): ").strip())
    if not 1 <= mes <= 12:
        raise ValueError("Mês inválido.")
    return ano, mes


def menu() -> None:
    while True:
        print("\n" + "=" * 65)
        print("LIVRO PONTO - PROFESSORES EMCA")
        print("=" * 65)
        print("1 - Criar/atualizar banco")
        print("2 - Importar todos os horários")
        print("3 - Atualizar resumo mensal")
        print("4 - Exportar mês para Excel")
        print("5 - Importar tudo e exportar um mês")
        print("0 - Sair")

        opcao = input("\nOpção: ").strip()

        try:
            if opcao == "1":
                caminho = criar_banco()
                print(f"Banco criado/atualizado: {caminho}")

            elif opcao == "2":
                importar_todos()

            elif opcao == "3":
                ano, mes = ler_ano_mes()
                qtd = atualizar_resumo_mensal(ano, mes)
                print(f"Resumo atualizado: {qtd} professores.")

            elif opcao == "4":
                ano, mes = ler_ano_mes()
                caminho = exportar_mes(ano, mes)
                print(f"Excel criado: {caminho}")

            elif opcao == "5":
                ano, mes = ler_ano_mes()
                importar_todos(ano)
                caminho = exportar_mes(ano, mes)
                print(f"Excel criado: {caminho}")

            elif opcao == "0":
                print("Encerrado.")
                break

            else:
                print("Opção inválida.")

        except Exception as exc:
            print(f"\n[ERRO] {exc}")


if __name__ == "__main__":
    menu()
