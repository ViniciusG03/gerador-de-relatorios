import os
import sys


def resource_path(relative_path: str) -> str:
    """Pega o caminho absoluto para o recurso, funciona para desenvolvimento e para PyInstaller"""
    try:
        base_path = sys._MEIPASS  # type: ignore
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)
