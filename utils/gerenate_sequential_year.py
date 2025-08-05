from datetime import datetime


def generate_sequential_year(destination_sheet):
    """
    Gera o sequencial anual com base no ano da última entrada na planilha.

    Args:
        destination_sheet (Worksheet): Aba de destino.

    Returns:
        tuple: Ano de registro e sequencial gerado.
    """

    try:
        current_year = datetime.now().year
        last_row = destination_sheet.max_row

        if last_row == 1:
            return current_year, 1

        try:
            last_year = destination_sheet.cell(row=last_row, column=17).value
        except AttributeError as e:
            raise ValueError(f"Erro ao acessar o ano na linha {last_row}, coluna 17: {e}")

        if last_year != current_year:
            return current_year, 1
        else:
            try:
                last_sequential = destination_sheet.cell(row=last_row, column=2).value
                return current_year, last_sequential + 1 if last_sequential else 1
            except AttributeError as e:
                raise ValueError(f"Erro ao acessar o sequencial na linha {last_row}, coluna 2: {e}")
    except Exception as e:
        raise ValueError(f"Erro ao gerar o sequencial anual: {e}")