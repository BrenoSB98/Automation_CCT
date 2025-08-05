
def generate_identifier(year, sequential, carrier, modality, contracting_regime, periodicity):
    """
    Gera o identificador único baseado nos parâmetros fornecidos.

    Args:
        year (int): Ano de registro.
        sequential (int): Sequencial anual.
        carrier (str): Transportadora.
        modality (str): Modalidade.
        contracting_regime (str): Regime de contratação.
        periodicity (str): Periodicidade.

    Returns:
        str: Identificador único formatado.
    """
    try:
        return f"{year}.{sequential:05}.{carrier}.{modality}.{contracting_regime}.{periodicity}"
    except TypeError as e:
        raise ValueError(f"Erro ao gerar identificador: verifique os tipos dos parâmetros fornecidos. {e}")
    except Exception as e:
        raise ValueError(f"Erro ao gerar identificador: {e}")