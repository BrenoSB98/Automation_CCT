def get_last_general_sequential(dest_sheet, column=18):
    """
    Obtém o último valor de sequencial geral da coluna especificada.

    Args:
        dest_sheet (Worksheet): Planilha de destino.
        column (int): Número da coluna onde o sequencial geral está armazenado.

    Returns:
        int: Último valor do sequencial geral ou 0 se não houver dados.
    """
    try:
        for row in range(dest_sheet.max_row, 0, -1):
            try:
                value = dest_sheet.cell(row=row, column=column).value
                if value is not None:
                    return value
            except AttributeError as e:
                raise ValueError(f"Erro ao acessar a célula na linha {row}, coluna {column}: {e}")
        return 0
    except Exception as e:
        raise ValueError(f"Erro ao obter o último sequencial geral: {e}")