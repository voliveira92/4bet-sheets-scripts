from google_auth import get_sheets_service


RELATION_SPREADSHEET_ID = (
    "1KDzRCwo5H06DAwpTDM5IH_8utkiGd-4fZnKEmJf-S9A"
)

RELATION_SHEET = "2026"


def read_black_relations():
    service = get_sheets_service()

    result = (
        service.spreadsheets()
        .values()
        .get(
            spreadsheetId=RELATION_SPREADSHEET_ID,
            range=f"{RELATION_SHEET}!C2:Q",
        )
        .execute()
    )

    rows = result.get("values", [])

    records = []

    for row_number, row in enumerate(rows, start=2):

        # Como estamos lendo C:Q:
        #
        # índice 0  = C
        # índice 11 = N
        # índice 12 = O
        # índice 13 = P
        # índice 14 = Q

        nick = get_value(row, 0)

        if not nick:
            continue

        links = [
            get_value(row, 11),
            get_value(row, 12),
            get_value(row, 13),
            get_value(row, 14),
        ]

        links = [
            link
            for link in links
            if is_google_sheet_url(link)
        ]

        for link in links:
            records.append(
                {
                    "nick": nick,
                    "link": link,
                    "source_row": row_number,
                }
            )

    return records


def get_value(row, index):
    if index >= len(row):
        return ""

    value = row[index]

    if value is None:
        return ""

    return str(value).strip()


def is_google_sheet_url(value):
    if not value:
        return False

    return "docs.google.com/spreadsheets/" in value