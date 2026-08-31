from relation_reader import read_black_relations
from services.player_extractor import extract_player_data
from services.normalizer import normalize_player_data
from services.year_builder import build_year_rows
from services.sheets_writer import write_year_data

def print_sample(rows, limit=15):

    print()
    print("=" * 100)
    print("AMOSTRA DO DATASET 2026")
    print("=" * 100)

    header = [
        "Nick",
        "Info",
        "Jan",
        "Fev",
        "Mar",
        "Abr",
        "Mai",
        "Jun",
        "Jul",
        "Ago",
        "Set",
        "Out",
        "Nov",
        "Dez",
    ]

    print(" | ".join(header))
    print("-" * 100)

    for row in rows[:limit]:

        printable = [
            "" if value is None else str(value)
            for value in row
        ]

        print(" | ".join(printable))


def main():

    print()
    print("=" * 70)
    print("AGENTE DE AUDITORIA - V1.1")
    print("BLACK 2026")
    print("=" * 70)

    print()
    print("Lendo relação Black...")

    records = read_black_relations()

    total = len(records)

    print(f"Planilhas encontradas: {total}")
    print()

    successes = []
    errors = []

    for index, record in enumerate(records, start=1):

        nick = record["nick"]
        url = record["link"]

        print(
            f"[{index}/{total}] "
            f"Processando: {nick}"
        )

        try:

            raw_data = extract_player_data(
                nick=nick,
                url=url,
            )

            normalized_data = normalize_player_data(
                raw_data
            )

            successes.append(normalized_data)

            print("    OK")

        except Exception as error:

            errors.append(
                {
                    "nick": nick,
                    "url": url,
                    "error": str(error),
                }
            )

            print(f"    ERRO: {error}")

    # -----------------------------------------------------
    # CONSTRUÇÃO DO DATASET ANUAL
    # -----------------------------------------------------

    year_rows = build_year_rows(successes)

    print_sample(
        year_rows,
        limit=15,
    )

    print()
    print("=" * 70)
    print("RESUMO DA EXECUÇÃO")
    print("=" * 70)

    print(f"Planilhas encontradas: {total}")
    print(f"Planilhas processadas: {len(successes)}")
    print(f"Planilhas com erro: {len(errors)}")

    print()
    print(f"Linhas geradas para 2026: {len(year_rows)}")

    expected_rows = len(successes) * 5

    print(f"Linhas esperadas: {expected_rows}")

        # -----------------------------------------------------
    # ESCRITA
    # -----------------------------------------------------

    if errors:
        print()
        print("ESCRITA CANCELADA.")
        print(
            "Existem planilhas com erro. "
            "Nenhuma alteração foi feita no destino."
        )
        return

    if len(year_rows) != expected_rows:
        print()
        print("ESCRITA CANCELADA.")
        print(
            "A quantidade de linhas geradas "
            "não corresponde à quantidade esperada."
        )
        return

    print()
    print("=" * 70)
    print("ESCREVENDO ABA 2026")
    print("=" * 70)

    result = write_year_data(year_rows)

    print()
    print(f"Aba: {result['sheet']}")
    print(f"Linhas de dados: {result['data_rows']}")
    print(f"Total com cabeçalho: {result['total_rows']}")

    print()
    print("ESCRITA CONCLUÍDA.")

    if len(year_rows) == expected_rows:
        print("Estrutura anual: OK")
    else:
        print("ATENÇÃO: quantidade inesperada de linhas.")

    if errors:

        print()
        print("-" * 70)
        print("PLANILHAS COM ERRO")
        print("-" * 70)

        for index, item in enumerate(errors, start=1):

            print()
            print(f"{index}. {item['nick']}")
            print(f"   Link: {item['url']}")
            print(f"   Erro: {item['error']}")

    print()
    print("=" * 70)
    print("V1.1 CONCLUÍDA")
    print("=" * 70)


if __name__ == "__main__":
    main()