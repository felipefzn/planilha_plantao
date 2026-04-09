from __future__ import annotations


class ValidacaoErro(ValueError):
    pass


def validar_nome_responsavel(nome: str) -> str:
    nome_limpo = (nome or "").strip()
    if not nome_limpo:
        raise ValidacaoErro("Informe o nome do responsável.")
    if len(nome_limpo) < 3:
        raise ValidacaoErro("O nome do responsável deve ter pelo menos 3 caracteres.")
    return nome_limpo


def validar_horas(horas: float) -> float:
    if horas <= 0:
        raise ValidacaoErro("As horas precisam ser maiores que zero.")
    return round(horas, 2)


def validar_tipo(tipo: str) -> str:
    tipo_limpo = (tipo or "").strip()
    if tipo_limpo not in {"Sobre aviso", "Suporte"}:
        raise ValidacaoErro("Selecione um tipo válido para o lançamento.")
    return tipo_limpo
