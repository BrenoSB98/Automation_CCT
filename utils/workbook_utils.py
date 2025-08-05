import subprocess
import sys

def install_openpyxl():
    """
    Verifica se a biblioteca 'openpyxl' está instalada e a instala silenciosamente, se necessário.
    """
    try:
        import openpyxl
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl", "--quiet"])

install_openpyxl()
from openpyxl import load_workbook

def load_workbooks(source_path, destination_path):
    """
        Carrega as planilhas de origem e destino.

        Args:
            source_path (str): Caminho para o arquivo de origem.
            destination_path (str): Caminho para o arquivo de destino.

        Returns:
            tuple: Planilhas de origem e destino carregadas.
        """
    try:
        source_workbook = load_workbook(source_path, data_only=True)
        destination_workbook = load_workbook(destination_path)
        return source_workbook, destination_workbook
    except FileNotFoundError as e:
        raise FileNotFoundError(f"Erro ao abrir os arquivos: {e}")
    except PermissionError as e:
        raise PermissionError(f"Permissão negada ao acessar os arquivos: {e}")
    except Exception as e:
        raise ValueError(f"Erro ao carregar as planilhas: {e}")