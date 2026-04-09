from __future__ import annotations

from dataclasses import asdict, dataclass, field


@dataclass(slots=True)
class RegraDiaEspecial:
    habilitado: bool
    tipo: str
    inicio: str
    fim: str
    vira_dia_seguinte: bool

    @classmethod
    def from_dict(cls, data: dict) -> "RegraDiaEspecial":
        return cls(
            habilitado=bool(data.get("habilitado", False)),
            tipo=data.get("tipo", "Sobre aviso"),
            inicio=data.get("inicio", "00:00"),
            fim=data.get("fim", "00:00"),
            vira_dia_seguinte=bool(data.get("vira_dia_seguinte", True)),
        )

    def to_dict(self) -> dict:
        return asdict(self)


@dataclass(slots=True)
class ConfiguracaoSistema:
    valor_hora_sobre_aviso: float
    valor_hora_suporte: float
    horario_padrao_inicio_plantao: str
    horario_padrao_fim_plantao: str
    regras_padrao_sabado: RegraDiaEspecial
    regras_padrao_domingo: RegraDiaEspecial
    nome_padrao_arquivo_exportado: str
    template_excel: str = ""
    idioma_calendario: str = "pt_BR"
    tema: dict = field(default_factory=dict)

    @classmethod
    def from_dict(cls, data: dict) -> "ConfiguracaoSistema":
        return cls(
            valor_hora_sobre_aviso=float(data.get("valor_hora_sobre_aviso", 3.51)),
            valor_hora_suporte=float(data.get("valor_hora_suporte", 10.5)),
            horario_padrao_inicio_plantao=data.get(
                "horario_padrao_inicio_plantao", "20:00"
            ),
            horario_padrao_fim_plantao=data.get(
                "horario_padrao_fim_plantao", "06:00"
            ),
            regras_padrao_sabado=RegraDiaEspecial.from_dict(
                data.get("regras_padrao_sabado", {})
            ),
            regras_padrao_domingo=RegraDiaEspecial.from_dict(
                data.get("regras_padrao_domingo", {})
            ),
            nome_padrao_arquivo_exportado=data.get(
                "nome_padrao_arquivo_exportado", "plantao_{ano}_{mes}.xlsx"
            ),
            template_excel=data.get("template_excel", ""),
            idioma_calendario=data.get("idioma_calendario", "pt_BR"),
            tema=data.get("tema", {}),
        )

    def to_dict(self) -> dict:
        data = asdict(self)
        data["regras_padrao_sabado"] = self.regras_padrao_sabado.to_dict()
        data["regras_padrao_domingo"] = self.regras_padrao_domingo.to_dict()
        return data

    def valor_por_tipo(self, tipo: str) -> float:
        return (
            self.valor_hora_sobre_aviso
            if tipo.strip().lower() == "sobre aviso"
            else self.valor_hora_suporte
        )
