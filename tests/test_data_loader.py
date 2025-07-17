import sys, os; sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import pandas as pd
import tempfile
import os
from gerador_relatorios import data_loader


def test_load_excel_success(tmp_path):
    df = pd.DataFrame({
        'NOME': ['A'],
        'DATA DE NASCIMENTO': ['01/01/2000'],
        'RESPONSÁVEL': ['B'],
        'ESPECIALIDADE': ['C'],
        'MÊS DE REFERÊNCIA': ['Jan/2025'],
    })
    file_path = tmp_path / 'test.xlsx'
    df.to_excel(file_path, index=False)

    loaded = data_loader.load_excel(str(file_path))
    assert loaded.equals(df)


def test_load_excel_missing_column(tmp_path):
    df = pd.DataFrame({'NOME': ['A']})
    file_path = tmp_path / 'bad.xlsx'
    df.to_excel(file_path, index=False)
    try:
        data_loader.load_excel(str(file_path))
        assert False, 'Expected ValueError'
    except ValueError:
        pass
