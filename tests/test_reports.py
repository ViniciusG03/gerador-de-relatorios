import sys, os; sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from gerador_relatorios import reports


def minimal_patient():
    return {
        'info': {
            'nome': 'Teste',
            'data_nascimento': '01/01/2000',
            'responsavel': 'Resp',
            'mes_referencia': 'Jan/2025',
        },
        'especialidades': ['PSICOTERAPIA']
    }


def test_generate_pne_report(tmp_path):
    data = minimal_patient()
    reports.generate_pne_report(data, str(tmp_path))
    expected = tmp_path / 'Relatório_PNE_Teste.docx'
    assert expected.exists()


def test_generate_tipico_report(tmp_path):
    data = minimal_patient()
    reports.generate_tipico_report(data, str(tmp_path))
    expected = tmp_path / 'Relatório_Típico_Teste.docx'
    assert expected.exists()
