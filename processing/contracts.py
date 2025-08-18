from mappings import CARRIERS, MODALITIES, REGIMES, PERIODICITIES
from utils import copy_format, generate_identifier, generate_sequential_year, get_last_general_sequential
from datetime import date

def process_contracts(source_sheet, destination_sheet, source_row):
    """
    Processa e transfere os dados da aba de origem para a aba Registro do destino,
    incluindo manipulação específica de colunas e geração de sequenciais.

    Args:
        source_sheet (Worksheet): Aba de origem do arquivo.
        destination_sheet (Worksheet): Aba Registro do arquivo de destino.
        source_row (int): Número da linha da aba de origem a ser processada.

    Returns:
        annual_sequential (int): Seguencial Anual
    """
    try:
        destination_row = destination_sheet.max_row + 1

        registration_year, annual_sequential = generate_sequential_year(destination_sheet)

        try:
            carrier = CARRIERS.get(source_sheet.cell(row=source_row, column=1).value.upper(), "00")
            modality = MODALITIES.get(source_sheet.cell(row=source_row, column=4).value.upper(), "99")
            regime = REGIMES.get(source_sheet.cell(row=source_row, column=5).value.upper(), "99")
            periodicity = PERIODICITIES.get(source_sheet.cell(row=source_row, column=6).value.upper(), "99")
        except AttributeError as e:
            raise ValueError(f"Erro ao processar mapeamentos na linha {source_row}: {e}")

        identifier = generate_identifier(registration_year, annual_sequential, carrier, modality, regime, periodicity)

        try:
            # Identificador (coluna 1 no destino)
            destination_sheet.cell(row=destination_row, column=1).value = identifier

            # Sequencial Anual (coluna 2 no destino)
            destination_sheet.cell(row=destination_row, column=2).value = annual_sequential

            # Transportador (coluna 1 na origem, coluna 3 no destino)
            destination_sheet.cell(row=destination_row, column=3).value = source_sheet.cell(row=source_row, column=1).value
            copy_format(source_sheet.cell(row=source_row, column=1), destination_sheet.cell(row=destination_row, column=3))

            # Carregador (coluna 2 na origem, coluna 5 no destino - enviando apenas o valor da fórmula)
            destination_sheet.cell(row=destination_row, column=5).value = source_sheet.cell(row=source_row, column=2).value
            copy_format(source_sheet.cell(row=source_row, column=2), destination_sheet.cell(row=destination_row, column=5))

            # CNPJ Carregador (coluna 3 na origem, coluna 6 no destino)
            destination_sheet.cell(row=destination_row, column=6).value = source_sheet.cell(row=source_row, column=3).value
            copy_format(source_sheet.cell(row=source_row, column=3), destination_sheet.cell(row=destination_row, column=6))

            # Modalidade (coluna 4 na origem, coluna 7 no destino)
            destination_sheet.cell(row=destination_row, column=7).value = source_sheet.cell(row=source_row, column=4).value
            copy_format(source_sheet.cell(row=source_row, column=4), destination_sheet.cell(row=destination_row, column=7))

            # Regime de Contratação (coluna 5 na origem, coluna 8 no destino)
            destination_sheet.cell(row=destination_row, column=8).value = source_sheet.cell(row=source_row, column=5).value
            copy_format(source_sheet.cell(row=source_row, column=5), destination_sheet.cell(row=destination_row, column=8))

            # Periodicidade (coluna 6 na origem, coluna 9 no destino)
            destination_sheet.cell(row=destination_row, column=9).value = source_sheet.cell(row=source_row, column=6).value
            copy_format(source_sheet.cell(row=source_row, column=6), destination_sheet.cell(row=destination_row, column=9))

            # Início (coluna 7 na origem, coluna 10 no destino)
            destination_sheet.cell(row=destination_row, column=10).value = source_sheet.cell(row=source_row, column=7).value
            copy_format(source_sheet.cell(row=source_row, column=7), destination_sheet.cell(row=destination_row, column=10))

            # Término (coluna 8 na origem, coluna 11 no destino)
            destination_sheet.cell(row=destination_row, column=11).value = source_sheet.cell(row=source_row, column=8).value
            copy_format(source_sheet.cell(row=source_row, column=8), destination_sheet.cell(row=destination_row, column=11))

            # Local Assinatura (coluna 9 na origem, coluna 12 no destino)
            destination_sheet.cell(row=destination_row, column=12).value = source_sheet.cell(row=source_row, column=9).value
            copy_format(source_sheet.cell(row=source_row, column=9), destination_sheet.cell(row=destination_row, column=12))

            # Data Assinatura (coluna 10 na origem, coluna 13 no destino)
            destination_sheet.cell(row=destination_row, column=13).value = source_sheet.cell(row=source_row, column=10).value
            copy_format(source_sheet.cell(row=source_row, column=10), destination_sheet.cell(row=destination_row, column=13))

            # Doc SEI inteiro Teor (coluna 11 na origem, coluna 14 no destino)
            destination_sheet.cell(row=destination_row, column=14).value = source_sheet.cell(row=source_row, column=11).value
            copy_format(source_sheet.cell(row=source_row, column=11), destination_sheet.cell(row=destination_row, column=14))

            # Processo Admin (coluna 12 na origem, coluna 15 no destino)
            destination_sheet.cell(row=destination_row, column=15).value = source_sheet.cell(row=source_row, column=12).value
            copy_format(source_sheet.cell(row=source_row, column=12), destination_sheet.cell(row=destination_row, column=15))

            # Data Protocolo (coluna 13 na origem, coluna 15 no destino)
            destination_sheet.cell(row=destination_row, column=16).value = source_sheet.cell(row=source_row, column=13).value
            copy_format(source_sheet.cell(row=source_row, column=13), destination_sheet.cell(row=destination_row, column=16))

            # Ano Registro (coluna 17 no destino)
            destination_sheet.cell(row=destination_row, column=17).value = registration_year

            # Sequencial Geral (coluna 18 no destino)
            last_general_sequential = get_last_general_sequential(destination_sheet, column=18)
            destination_sheet.cell(row=destination_row, column=18).value = last_general_sequential + 1

            # Doc SEI Minuta (coluna 14 na origem, coluna 19 no destino)
            destination_sheet.cell(row=destination_row, column=19).value = source_sheet.cell(row=source_row, column=14).value
            copy_format(source_sheet.cell(row=source_row, column=14), destination_sheet.cell(row=destination_row, column=19))

            # Doc SEI versão Tarjada (coluna 15 na origem, coluna 20 no destino)
            destination_sheet.cell(row=destination_row, column=20).value = source_sheet.cell(row=source_row, column=15).value
            copy_format(source_sheet.cell(row=source_row, column=15), destination_sheet.cell(row=destination_row, column=20))

            # COD Transportadora (coluna 16 na origem, coluna 21 no destino)
            destination_sheet.cell(row=destination_row, column=21).value = source_sheet.cell(row=source_row, column=16).value
            copy_format(source_sheet.cell(row=source_row, column=16), destination_sheet.cell(row=destination_row, column=21))

            # Data de Atualização (coluna 22 no destino)
            data_atualizacao = date.today()
            destination_sheet.cell(row=destination_row, column=22).value = data_atualizacao.strftime('%d/%m/%Y')

        except Exception as e:
            raise ValueError(f"Erro ao preencher células no destino na linha {source_row}: {e}")

        return annual_sequential

    except Exception as e:
        raise ValueError(f"Erro na linha {source_row}: {e}")
