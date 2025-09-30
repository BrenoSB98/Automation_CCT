import os
import sys
import tkinter as tk
from tkinter import messagebox
from interface.handlers import selecionar_planilhas, criar_backup
from processing import process_contracts, process_qdc
from utils import load_workbooks
from exception import error
from logger import get_logger

logger = get_logger()

def main():
    """
        Função principal responsável por executar o fluxo completo de importação de dados de contratos
        de transporte de gás natural a partir de múltiplas planilhas para uma planilha destino.

        Etapas:
        1. Seleciona arquivos de origem e destino via interface gráfica (Tkinter).
        2. Cria um backup automático da planilha destino.
        3. Carrega e processa cada linha das planilhas de origem, populando a planilha destino.
        4. Salva os dados e exibe mensagem de sucesso ou erro.

        Saída: encerra o processo com código de status (0 para sucesso, 1 para erro).
    """

    destino, origem_lista = selecionar_planilhas()

    if not destino:
        print("Nenhuma planilha de destino foi selecionada.")
        sys.exit(1)

    if not origem_lista:
        print("Nenhuma planilha de origem foi selecionada.")
        sys.exit(1)

    criar_backup(destino)

    try:
        for caminho_origem in origem_lista:
            nome_arquivo = os.path.basename(caminho_origem)
            try:
                logger.info(f"Carregando planilha: {nome_arquivo}")
                source_wb, dest_wb = load_workbooks(caminho_origem, destino)

                source_sheet = source_wb["Registro de Contratos"]
                register_sheet = dest_wb["REGISTRO CSTs"]
                qdc_sheet = dest_wb["QDC"]

                for row in range(3, source_sheet.max_row + 1):
                    if source_sheet.cell(row=row, column=1).value is None:
                        logger.info(f"Linha vazia encontrada na linha {row} da planilha {nome_arquivo}. Pulando para a próxima planilha.")
                        break
                    try:
                        sequential_annual = process_contracts(source_sheet, register_sheet, row)
                        process_qdc(source_sheet, qdc_sheet, row, sequential_annual)
                        logger.info(f"Linha {row} da planilha {nome_arquivo} processada com sucesso.")
                    except Exception as e:
                        error(f"Erro na linha {row} da planilha {nome_arquivo}: {e}")

                dest_wb.save(destino)
                logger.info(f"Planilha {nome_arquivo} salva no destino: {destino}")

            except Exception as e:
                error(f"Erro ao processar a planilha {nome_arquivo}: {e}")

        # Inicializa a raiz da interface Tkinter sem exibir a janela principal
        root = tk.Tk()
        root.withdraw()

        # Agenda ações para garantir que o messagebox apareça em primeiro plano
        root.after_idle(root.deiconify)
        root.after_idle(root.lift)
        root.after_idle(root.focus_force)

        # Exibe mensagem de sucesso após a carga
        messagebox.showinfo("Finalizado", "Processamento concluído com sucesso!", parent=root)
        root.destroy()

        logger.info("Processamento de todas as planilhas concluído.")
        sys.exit(0)

    except Exception as e:
        error(f"Erro fatal na execução: {e}")
        root = tk.Tk()
        root.withdraw()
        root.after_idle(root.deiconify)
        root.after_idle(root.lift)
        root.after_idle(root.focus_force)
        messagebox.showerror("Erro", f"Erro fatal: {e}", parent=root)
        root.destroy()
        sys.exit(1)

if __name__ == "__main__":
    main()
