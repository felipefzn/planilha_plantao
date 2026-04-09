from __future__ import annotations

from copy import deepcopy
from uuid import uuid4

from app.models import ConfiguracaoSistema, LancamentoPlantao, ProjetoPlantao
from app.utils.datetime_utils import (
    calcular_fim_por_inicio_e_fim,
    calcular_fim_por_inicio_e_horas,
    calcular_horas,
    format_time_hhmm,
    nome_dia_semana,
    nome_mes,
    parse_date,
)
from app.utils.validators import ValidacaoErro, validar_horas, validar_tipo


class ScheduleService:
    def generate_weekly_schedule(
        self,
        projeto: ProjetoPlantao,
        config: ConfiguracaoSistema,
        week_info: dict,
        responsavel_id: str,
    ) -> int:
        responsavel = next(
            (item for item in projeto.responsaveis if item.id == responsavel_id), None
        )
        if not responsavel:
            raise ValidacaoErro("Selecione um responsável válido para a semana.")

        week_id = week_info["id"]
        projeto.atribuicoes_semanais[week_id] = {
            "responsavel_id": responsavel.id,
            "responsavel": responsavel.nome,
            "inicio": week_info["inicio"].isoformat(),
            "fim": week_info["fim"].isoformat(),
            "label": week_info["label"],
        }

        projeto.lancamentos = [
            item
            for item in projeto.lancamentos
            if not (item.origem == "automatico" and item.semana_referencia == week_id)
        ]

        criados = 0
        for dia in week_info["dias"]:
            if dia.month != projeto.mes or dia.year != projeto.ano:
                continue

            regra = self._resolve_day_rule(config, dia.weekday())
            if not regra:
                continue

            inicio_dt, fim_dt = calcular_fim_por_inicio_e_fim(
                dia,
                regra["inicio"],
                regra["fim"],
                regra["vira_dia_seguinte"],
            )
            horas = validar_horas(calcular_horas(inicio_dt, fim_dt))
            lancamento = LancamentoPlantao(
                id=str(uuid4()),
                responsavel_id=responsavel.id,
                responsavel=responsavel.nome,
                data=dia.isoformat(),
                inicio=format_time_hhmm(inicio_dt),
                fim=format_time_hhmm(fim_dt),
                tipo=regra["tipo"],
                total_horas=horas,
                valor=round(horas * config.valor_por_tipo(regra["tipo"]), 2),
                mes=nome_mes(dia.month),
                dia_semana=nome_dia_semana(dia),
                origem="automatico",
                semana_referencia=week_id,
            )
            projeto.lancamentos.append(lancamento)
            criados += 1

        self.sort_entries(projeto)
        return criados

    def generate_month_schedule(
        self,
        projeto: ProjetoPlantao,
        config: ConfiguracaoSistema,
        weeks: list[dict],
        fallback_responsavel_id: str = "",
    ) -> tuple[int, int]:
        criados_total = 0
        semanas_sem_responsavel = 0

        for week_info in weeks:
            atribuicao = projeto.atribuicoes_semanais.get(week_info["id"], {})
            responsavel_id = atribuicao.get("responsavel_id") or fallback_responsavel_id
            if not responsavel_id:
                semanas_sem_responsavel += 1
                continue

            criados_total += self.generate_weekly_schedule(
                projeto,
                config,
                week_info,
                responsavel_id,
            )

        return criados_total, semanas_sem_responsavel

    def build_entry(
        self,
        config: ConfiguracaoSistema,
        responsavel_id: str,
        responsavel_nome: str,
        data_referencia,
        inicio: str,
        tipo: str,
        fim: str = "",
        horas: float | None = None,
        numero_chamado: str = "",
        solicitante: str = "",
        cliente: str = "",
        nivel: str = "",
        observacao: str = "",
        origem: str = "manual",
        semana_referencia: str = "",
        lancamento_id: str = "",
    ) -> LancamentoPlantao:
        data_obj = parse_date(data_referencia)
        tipo_validado = validar_tipo(tipo)
        if not responsavel_id or not responsavel_nome.strip():
            raise ValidacaoErro("Responsável é obrigatório.")
        if not inicio.strip():
            raise ValidacaoErro("Hora de início é obrigatória.")
        if not fim.strip() and horas is None:
            raise ValidacaoErro("Informe hora fim ou quantidade de horas.")

        if fim.strip():
            inicio_dt, fim_dt = calcular_fim_por_inicio_e_fim(data_obj, inicio, fim)
            horas_calculadas = calcular_horas(inicio_dt, fim_dt)
        else:
            horas_calculadas = validar_horas(float(horas or 0))
            inicio_dt, fim_dt = calcular_fim_por_inicio_e_horas(
                data_obj, inicio, horas_calculadas
            )

        horas_validas = validar_horas(horas_calculadas)
        valor = round(horas_validas * config.valor_por_tipo(tipo_validado), 2)

        return LancamentoPlantao(
            id=lancamento_id or str(uuid4()),
            responsavel_id=responsavel_id,
            responsavel=responsavel_nome.strip(),
            data=data_obj.isoformat(),
            inicio=format_time_hhmm(inicio_dt),
            fim=format_time_hhmm(fim_dt),
            tipo=tipo_validado,
            total_horas=horas_validas,
            valor=valor,
            numero_chamado=numero_chamado.strip(),
            solicitante=solicitante.strip(),
            cliente=cliente.strip(),
            nivel=nivel.strip(),
            observacao=observacao.strip(),
            mes=nome_mes(data_obj.month),
            dia_semana=nome_dia_semana(data_obj),
            origem=origem,
            semana_referencia=semana_referencia,
        )

    def replace_entry(self, projeto: ProjetoPlantao, lancamento: LancamentoPlantao) -> None:
        for indice, item in enumerate(projeto.lancamentos):
            if item.id == lancamento.id:
                projeto.lancamentos[indice] = lancamento
                self.sort_entries(projeto)
                return
        projeto.lancamentos.append(lancamento)
        self.sort_entries(projeto)

    def remove_entry(self, projeto: ProjetoPlantao, lancamento_id: str) -> None:
        projeto.lancamentos = [
            item for item in projeto.lancamentos if item.id != lancamento_id
        ]

    def duplicate_entry(self, projeto: ProjetoPlantao, lancamento_id: str) -> LancamentoPlantao:
        original = next(
            (item for item in projeto.lancamentos if item.id == lancamento_id), None
        )
        if not original:
            raise ValidacaoErro("Lançamento não encontrado.")
        copia = deepcopy(original)
        copia.id = str(uuid4())
        copia.origem = "manual"
        projeto.lancamentos.append(copia)
        self.sort_entries(projeto)
        return copia

    def sort_entries(self, projeto: ProjetoPlantao) -> None:
        projeto.lancamentos.sort(
            key=lambda item: (item.data, item.inicio, item.responsavel.lower(), item.tipo)
        )

    def _resolve_day_rule(self, config: ConfiguracaoSistema, weekday: int) -> dict | None:
        if weekday <= 4:
            return {
                "tipo": "Sobre aviso",
                "inicio": config.horario_padrao_inicio_plantao,
                "fim": config.horario_padrao_fim_plantao,
                "vira_dia_seguinte": True,
            }
        if weekday == 5 and config.regras_padrao_sabado.habilitado:
            regra = config.regras_padrao_sabado
            return {
                "tipo": regra.tipo,
                "inicio": regra.inicio,
                "fim": regra.fim,
                "vira_dia_seguinte": regra.vira_dia_seguinte,
            }
        if weekday == 6 and config.regras_padrao_domingo.habilitado:
            regra = config.regras_padrao_domingo
            return {
                "tipo": regra.tipo,
                "inicio": regra.inicio,
                "fim": regra.fim,
                "vira_dia_seguinte": regra.vira_dia_seguinte,
            }
        return None
