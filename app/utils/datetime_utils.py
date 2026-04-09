from __future__ import annotations

import calendar
from datetime import date, datetime, time, timedelta

MESES_PT = {
    1: "Janeiro",
    2: "Fevereiro",
    3: "Março",
    4: "Abril",
    5: "Maio",
    6: "Junho",
    7: "Julho",
    8: "Agosto",
    9: "Setembro",
    10: "Outubro",
    11: "Novembro",
    12: "Dezembro",
}

DIAS_SEMANA_PT = {
    0: "Segunda-feira",
    1: "Terça-feira",
    2: "Quarta-feira",
    3: "Quinta-feira",
    4: "Sexta-feira",
    5: "Sábado",
    6: "Domingo",
}


def nome_mes(numero: int) -> str:
    return MESES_PT.get(numero, "")


def nome_dia_semana(data_referencia: date) -> str:
    return DIAS_SEMANA_PT.get(data_referencia.weekday(), "")


def parse_date(value: str | date) -> date:
    if isinstance(value, date):
        return value
    return datetime.strptime(value, "%Y-%m-%d").date()


def format_date_iso(value: date) -> str:
    return value.isoformat()


def parse_time(value: str | time) -> time:
    if isinstance(value, time):
        return value
    return datetime.strptime(value.strip(), "%H:%M").time()


def format_time_hhmm(value: datetime | time) -> str:
    if isinstance(value, datetime):
        return value.strftime("%H:%M")
    return value.strftime("%H:%M")


def combinar_data_hora(data_referencia: date, horario: str | time) -> datetime:
    return datetime.combine(data_referencia, parse_time(horario))


def calcular_fim_por_inicio_e_fim(
    data_referencia: date, inicio: str, fim: str, vira_dia_seguinte: bool | None = None
) -> tuple[datetime, datetime]:
    inicio_dt = combinar_data_hora(data_referencia, inicio)
    fim_dt = combinar_data_hora(data_referencia, fim)
    if vira_dia_seguinte is True or fim_dt <= inicio_dt:
        fim_dt += timedelta(days=1)
    return inicio_dt, fim_dt


def calcular_fim_por_inicio_e_horas(
    data_referencia: date, inicio: str, horas: float
) -> tuple[datetime, datetime]:
    inicio_dt = combinar_data_hora(data_referencia, inicio)
    fim_dt = inicio_dt + timedelta(hours=horas)
    return inicio_dt, fim_dt


def calcular_horas(inicio_dt: datetime, fim_dt: datetime) -> float:
    horas = (fim_dt - inicio_dt).total_seconds() / 3600
    return round(horas, 2)


def obter_semanas_do_mes(ano: int, mes: int) -> list[dict]:
    calendario = calendar.Calendar(firstweekday=0)
    semanas = []
    for indice, semana in enumerate(calendario.monthdatescalendar(ano, mes), start=1):
        inicio = semana[0]
        fim = semana[-1]
        semanas.append(
            {
                "indice": indice,
                "id": f"{inicio.isoformat()}__{fim.isoformat()}",
                "inicio": inicio,
                "fim": fim,
                "dias": semana,
                "label": f"Semana {indice} ({inicio.strftime('%d/%m')} a {fim.strftime('%d/%m')})",
            }
        )
    return semanas


def encontrar_semana_por_data(ano: int, mes: int, alvo: date) -> dict | None:
    for semana in obter_semanas_do_mes(ano, mes):
        if semana["inicio"] <= alvo <= semana["fim"]:
            return semana
    return None
