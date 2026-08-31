METRIC_ORDER = [
    "Profit",
    "Mtt",
    "Buy in",
    "ROI",
    "ABI",
]

MONTH_ORDER = [
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


def build_year_rows(successes):
    """
    Converte as planilhas processadas em linhas no formato:

    Nick | Info | Jan | Fev | ... | Dez

    Cada planilha gera exatamente 5 linhas.
    """

    rows = []

    for player_data in successes:

        nick = player_data["nick"]
        metrics = player_data["metrics"]

        for metric in METRIC_ORDER:

            months = metrics.get(metric, {})

            row = [
                nick,
                metric,
            ]

            for month in MONTH_ORDER:
                row.append(months.get(month))

            rows.append(row)

    return rows