def normalize_value(value, decimals=2):
    """
    Normaliza valores vindos do Google Sheets.

    Regras:
    - "" vira None
    - None permanece None
    - int permanece int
    - float é arredondado
    - texto permanece texto
    """

    if value is None:
        return None

    if value == "":
        return None

    if isinstance(value, int):
        return value

    if isinstance(value, float):
        return round(value, decimals)

    return value


def normalize_player_data(data):
    """
    Normaliza toda a estrutura retornada pelo player_extractor.
    """

    normalized_metrics = {}

    for metric, months in data["metrics"].items():

        normalized_months = {}

        for month, value in months.items():
            normalized_months[month] = normalize_value(value)

        normalized_metrics[metric] = normalized_months

    return {
        "nick": data["nick"],
        "spreadsheet_id": data["spreadsheet_id"],
        "source_url": data["source_url"],
        "metrics": normalized_metrics,
    }