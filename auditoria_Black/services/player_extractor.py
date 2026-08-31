import time
import re

from google_auth import get_sheets_service


MONTHS = [
    "Jan", "Fev", "Mar", "Abr", "Mai", "Jun",
    "Jul", "Ago", "Set", "Out", "Nov", "Dez",
]

METRICS = [
    "Profit",
    "Mtt",
    "Buy in",
    "ROI",
    "ABI",
]


def extract_spreadsheet_id(url):
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", url)

    if not match:
        raise ValueError(
            f"Não foi possível identificar o ID da planilha: {url}"
        )

    return match.group(1)


def extract_player_data(nick, url, max_attempts=3):
    """
    Lê Sumario!E6:P10.

    Faz novas tentativas em caso de falha transitória.
    Continua sendo somente leitura.
    """

    spreadsheet_id = extract_spreadsheet_id(url)

    last_error = None

    for attempt in range(1, max_attempts + 1):

        try:
            service = get_sheets_service()

            result = (
                service.spreadsheets()
                .values()
                .get(
                    spreadsheetId=spreadsheet_id,
                    range="Sumario!E6:P10",
                    valueRenderOption="UNFORMATTED_VALUE",
                )
                .execute()
            )

            rows = result.get("values", [])

            if not rows:
                raise ValueError(
                    f"Nenhum dado encontrado em Sumario!E6:P10 para {nick}"
                )

            metrics = {}

            for index, metric in enumerate(METRICS):

                row = rows[index] if index < len(rows) else []

                monthly_values = {}

                for month_index, month in enumerate(MONTHS):

                    value = (
                        row[month_index]
                        if month_index < len(row)
                        else None
                    )

                    monthly_values[month] = value

                metrics[metric] = monthly_values

            return {
                "nick": nick,
                "spreadsheet_id": spreadsheet_id,
                "source_url": url,
                "metrics": metrics,
            }

        except Exception as error:

            last_error = error

            if attempt < max_attempts:
                print(
                    f"    Tentativa {attempt} falhou. "
                    f"Tentando novamente..."
                )

                time.sleep(2)

    raise last_error