#from config import DESTINATION_LOCAL_PATH
from processing import process_contracts, process_qdc
from utils import load_workbooks
from exception import error

def main():
    try:
        destination_path = input(f"Digite o caminho local completo da planilha de destino (planilhão): ").strip('"')
        destination_path = destination_path.replace("\\", "/")

        num_workbooks = int(input("Digite o número de planilhas que serão cadastradas: "))

        if num_workbooks <= 0:
            raise ValueError("O número de planilhas deve ser maior que zero.")

        for workbook in range(1, num_workbooks + 1):
            try:
                source_path = input(
                    f"Digite o caminho local completo da planilha {workbook} que será carregada: ").strip('"')
                source_path = source_path.replace("\\", "/")

                source_workbook, destination_workbook = load_workbooks(source_path, destination_path)

                try:
                    source_sheet = source_workbook["Registro de Contratos"]
                    register_sheet = destination_workbook["REGISTRO CSTs"]
                    qdc_sheet = destination_workbook["QDC"]

                    last_row_source = source_sheet.max_row

                    for row in range(2, last_row_source + 1):
                        try:
                            sequential_annual = process_contracts(source_sheet, register_sheet, row)
                            process_qdc(source_sheet, qdc_sheet, row, sequential_annual)
                            print(f'Linha {row} carregada com sucesso.')
                        except Exception as e:
                            error(f"Erro ao processar a linha {row} na planilha {workbook}: {e}")

                    destination_workbook.save(destination_path)
                    print(f"Carga de dados realizada com sucesso no arquivo local '{destination_path}'.")
                except KeyError as e:
                    error(f"Erro ao acessar uma das planilhas no arquivo: {e}")

            except FileNotFoundError:
                error(f"O arquivo especificado para a planilha {workbook} não foi encontrado: {source_path}")

            except Exception as e:
                error(f"Erro ao carregar a planilha {workbook}: {e}")

    except ValueError as e:
        error(f"Número de planilhas inválido: {e}")

    except Exception as e:
        error(f"Erro no programa: {e}")

if __name__ == '__main__':
    main()
