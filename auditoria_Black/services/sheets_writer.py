from google_auth import get_writer_service


ALLOWED_SPREADSHEET_ID = (
    "1YIpNeJ-aje_K3gUn5Gyh2dd6b75UPXZ7ORxzd1MELbI"
)

YEAR_SHEET = "2026"

HEADER = [
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


def validate_write_target(spreadsheet_id):
    """
    Barreira adicional de segurança.

    Nenhum ID diferente do explicitamente autorizado
    pode ser utilizado pelas funções de escrita.
    """

    if spreadsheet_id != ALLOWED_SPREADSHEET_ID:
        raise PermissionError(
            "ESCRITA BLOQUEADA: "
            "a planilha informada não está autorizada."
        )


def ensure_year_sheet(service, spreadsheet_id):
    """
    Cria a aba 2026 somente se ela ainda não existir.
    """

    validate_write_target(spreadsheet_id)

    metadata = (
        service.spreadsheets()
        .get(
            spreadsheetId=spreadsheet_id,
            fields="sheets.properties",
        )
        .execute()
    )

    existing_sheets = [
        sheet["properties"]["title"]
        for sheet in metadata.get("sheets", [])
    ]

    if YEAR_SHEET in existing_sheets:
        return False

    body = {
        "requests": [
            {
                "addSheet": {
                    "properties": {
                        "title": YEAR_SHEET
                    }
                }
            }
        ]
    }

    (
        service.spreadsheets()
        .batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=body,
        )
        .execute()
    )

    return True


def write_year_data(rows):
    """
    Escreve o dataset anual somente na planilha autorizada.
    """

    spreadsheet_id = ALLOWED_SPREADSHEET_ID

    validate_write_target(spreadsheet_id)

    service = get_writer_service()

    created = ensure_year_sheet(
        service,
        spreadsheet_id,
    )

    if created:
        print("Aba 2026 criada.")
    else:
        print("Aba 2026 já existe.")

    # Limpa somente A:N da aba 2026.
    (
        service.spreadsheets()
        .values()
        .clear(
            spreadsheetId=spreadsheet_id,
            range=f"{YEAR_SHEET}!A:N",
            body={},
        )
        .execute()
    )

    values = [
        HEADER,
        *rows,
    ]

    (
        service.spreadsheets()
        .values()
        .update(
            spreadsheetId=spreadsheet_id,
            range=f"{YEAR_SHEET}!A1",
            valueInputOption="RAW",
            body={
                "values": values
            },
        )
        .execute()
    )

    return {
        "sheet": YEAR_SHEET,
        "data_rows": len(rows),
        "total_rows": len(values),
    }