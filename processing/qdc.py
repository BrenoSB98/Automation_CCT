from mappings import CARRIERS, MODALITIES, REGIMES, PERIODICITIES
from utils import copy_format, generate_sequential_year, generate_identifier


def process_qdc(source_sheet, destination_sheet, source_row, annual_sequential):
    """
    Processa e transfere os dados da aba de origem para a aba Volume do destino,
    incluindo manipulação específica de colunas e geração de sequenciais.

    Args:
        source_sheet (Worksheet): Aba de origem do arquivo.
        destination_sheet (Worksheet): Aba Volume do arquivo de destino.
        source_row (int): Número da linha da aba de origem a ser processada.
        annual_sequential (int): Sequencial de cadastro no ano.
    """
    try:
        destination_row = destination_sheet.max_row + 1

        try:
            registration_year, sequential = generate_sequential_year(destination_sheet)
        except Exception as e:
            raise ValueError(f"Erro ao gerar sequencial no destino: {e}")

        try:
            carrier = CARRIERS.get(source_sheet.cell(row=source_row, column=1).value.upper(), "00")
            modality = MODALITIES.get(source_sheet.cell(row=source_row, column=4).value.upper(), "99")
            regime = REGIMES.get(source_sheet.cell(row=source_row, column=5).value.upper(), "99")
            periodicity = PERIODICITIES.get(source_sheet.cell(row=source_row, column=6).value.upper(), "99")
        except AttributeError as e:
            raise ValueError(f"Erro ao acessar mapeamentos na linha {source_row}: {e}")

        try:
            identifier = generate_identifier(registration_year, annual_sequential, carrier, modality, regime, periodicity)
        except Exception as e:
            raise ValueError(f"Erro ao gerar identificador na linha {source_row}: {e}")

        try:
            if (source_sheet.cell(row=source_row, column=4).value != 'Acordo-quadro ("Master")' and
                source_sheet.cell(row=source_row, column=5).value != 'Acordo-Quadro ("Master")'):

                # Sequencial Geral (coluna 1 no destino)
                last_general_sequential = destination_sheet.cell(row=destination_sheet.max_row, column=1).value or 0
                destination_sheet.cell(row=destination_row, column=1).value = last_general_sequential + 1

                # Instrumento Contratual (coluna 2 no destino)
                destination_sheet.cell(row=destination_row, column=2).value = identifier

                # Zona/Ponto (coluna 18 na origem, coluna 3 no destino)
                destination_sheet.cell(row=destination_row, column=3).value = source_sheet.cell(row=source_row, column=18).value
                copy_format(source_sheet.cell(row=source_row, column=18), destination_sheet.cell(row=destination_row, column=3))

                # Entrada/Saída (coluna 5 na origem, coluna 4 no destino)
                destination_sheet.cell(row=destination_row, column=4).value = source_sheet.cell(row=source_row, column=5).value
                copy_format(source_sheet.cell(row=source_row, column=5), destination_sheet.cell(row=destination_row, column=4))

                # Valor da QDC em mil m³/dia (coluna 17 na origem, coluna 5 no destino)
                destination_sheet.cell(row=destination_row, column=5).value = source_sheet.cell(row=source_row, column=17).value
                copy_format(source_sheet.cell(row=source_row, column=17), destination_sheet.cell(row=destination_row, column=5))

                # Inicio (coluna 7 na origem, coluna 6 no destino)
                destination_sheet.cell(row=destination_row, column=6).value = source_sheet.cell(row=source_row, column=7).value
                copy_format(source_sheet.cell(row=source_row, column=7), destination_sheet.cell(row=destination_row, column=6))

                # Término (coluna 8 na origem, coluna 7 no destino)
                destination_sheet.cell(row=destination_row, column=7).value = source_sheet.cell(row=source_row, column=8).value
                copy_format(source_sheet.cell(row=source_row, column=8), destination_sheet.cell(row=destination_row, column=7))

                # QDC Máx_ARF (coluna 9 no destino)
                value_col2 = source_sheet.cell(row=source_row, column=2).value
                destination_sheet.cell(row=destination_row, column=9).value = (
                    "sim" if value_col2 and str(value_col2).startswith("ARF") else "não"
                )
        except Exception as e:
            raise ValueError(f"Erro ao preencher destino na linha {source_row}: {e}")
    except Exception as e:
        raise ValueError(f"Erro ao processar linha {source_row}: {e}")