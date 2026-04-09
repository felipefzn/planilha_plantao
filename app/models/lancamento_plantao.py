from __future__ import annotations

from dataclasses import asdict, dataclass, field
from uuid import uuid4


@dataclass(slots=True)
class LancamentoPlantao:
    responsavel_id: str
    responsavel: str
    data: str
    inicio: str
    fim: str
    tipo: str
    total_horas: float
    valor: float
    numero_chamado: str = ""
    solicitante: str = ""
    cliente: str = ""
    nivel: str = ""
    observacao: str = ""
    mes: str = ""
    dia_semana: str = ""
    origem: str = "manual"
    semana_referencia: str = ""
    id: str = field(default_factory=lambda: str(uuid4()))

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict) -> "LancamentoPlantao":
        return cls(
            id=data.get("id") or str(uuid4()),
            responsavel_id=data.get("responsavel_id", ""),
            responsavel=data.get("responsavel", ""),
            data=data.get("data", ""),
            inicio=data.get("inicio", ""),
            fim=data.get("fim", ""),
            tipo=data.get("tipo", ""),
            total_horas=float(data.get("total_horas", 0) or 0),
            valor=float(data.get("valor", 0) or 0),
            numero_chamado=data.get("numero_chamado", ""),
            solicitante=data.get("solicitante", ""),
            cliente=data.get("cliente", ""),
            nivel=data.get("nivel", ""),
            observacao=data.get("observacao", ""),
            mes=data.get("mes", ""),
            dia_semana=data.get("dia_semana", ""),
            origem=data.get("origem", "manual"),
            semana_referencia=data.get("semana_referencia", ""),
        )
