from __future__ import annotations

import pandas as pd

REQUIRED_COLUMNS = [
    "NOME",
    "DATA DE NASCIMENTO",
    "RESPONSÁVEL",
    "ESPECIALIDADE",
    "MÊS DE REFERÊNCIA",
]


def load_excel(path: str) -> pd.DataFrame:
    """Load an Excel file and validate required columns."""
    data = pd.read_excel(path)
    missing = [c for c in REQUIRED_COLUMNS if c not in data.columns]
    if missing:
        raise ValueError(
            f"Colunas obrigatórias não encontradas: {', '.join(missing)}"
        )
    return data
