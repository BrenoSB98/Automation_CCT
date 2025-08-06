import os
import shutil
import tkinter as tk
from datetime import datetime
from tkinter import filedialog
from logger import get_logger
from dotenv import load_dotenv

load_dotenv()
logger = get_logger()

backup_dir = os.getenv('BACKUP_DIR')

def selecionar_planilhas():
    """
        Abre uma interface gráfica para o usuário selecionar a planilha de destino e uma ou mais planilhas de origem.

        Returns:
            tuple:
                - str or None: Caminho para a planilha de destino selecionada, ou None se não selecionado.
                - list of str: Lista com os caminhos das planilhas de origem selecionadas.
    """

    root = tk.Tk()
    root.withdraw()

    # Seleção da planilha de destino (planilhão)
    destino = filedialog.askopenfilename(
        title="Selecione a planilha de DESTINO (planilhão)",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if not destino:
        root.destroy()
        return None, []

    # Seleção da planilha de destino (planilhas das transportadoras)
    origem = filedialog.askopenfilenames(
        title="Selecione uma ou mais planilhas de origem",
        filetypes=[("Excel files", "*.xlsx")]
    )
    root.destroy()
    return destino, list(origem)

def criar_backup(caminho_arquivo):
    """
        Cria uma cópia de backup do arquivo Excel informado no diretório especificado pela variável de ambiente BACKUP_DIR.
        O nome do backup conterá um timestamp para identificação única.

        Args:
            caminho_arquivo (str): Caminho absoluto do arquivo de destino que será copiado.

        Returns:
            str: Caminho completo do arquivo de backup criado.
    """

    if not os.path.exists(backup_dir):
        os.makedirs(backup_dir)

    _, nome = os.path.split(caminho_arquivo)
    base, ext = os.path.splitext(nome)
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    nome_backup = f"{timestamp}_{base}.backup{ext}"
    caminho_backup = os.path.join(backup_dir, nome_backup)

    # Copia o arquivo original para o novo caminho de backup
    shutil.copy(caminho_arquivo, caminho_backup)
    logger.info(f"Backup criado: {caminho_backup}")
    return caminho_backup
