from __future__ import annotations

from datetime import date

from .datetime_utils import parse_date


def formatar_moeda_br(valor: float) -> str:
    inteiro, decimal = f"{valor:.2f}".split(".")
    inteiro_formatado = ""
    for indice, caractere in enumerate(reversed(inteiro)):
        if indice and indice % 3 == 0:
            inteiro_formatado = "." + inteiro_formatado
        inteiro_formatado = caractere + inteiro_formatado
    return f"R$ {inteiro_formatado},{decimal}"


def formatar_horas(valor: float) -> str:
    return f"{valor:.2f} h"


def formatar_data_br(valor: str | date) -> str:
    data_referencia = parse_date(valor)
    return data_referencia.strftime("%d/%m/%Y")
