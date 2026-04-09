from __future__ import annotations

from dataclasses import dataclass, field

from .lancamento_plantao import LancamentoPlantao
from .responsavel import Responsavel


@dataclass(slots=True)
class ProjetoPlantao:
    nome: str
    ano: int
    mes: int
    responsaveis: list[Responsavel] = field(default_factory=list)
    lancamentos: list[LancamentoPlantao] = field(default_factory=list)
    atribuicoes_semanais: dict[str, dict] = field(default_factory=dict)
    caminho_arquivo: str = ""

    def to_dict(self) -> dict:
        return {
            "nome": self.nome,
            "ano": self.ano,
            "mes": self.mes,
            "responsaveis": [item.to_dict() for item in self.responsaveis],
            "lancamentos": [item.to_dict() for item in self.lancamentos],
            "atribuicoes_semanais": self.atribuicoes_semanais,
        }

    @classmethod
    def from_dict(cls, data: dict, caminho_arquivo: str = "") -> "ProjetoPlantao":
        return cls(
            nome=data.get("nome", "Projeto de Plantão"),
            ano=int(data.get("ano") or 0),
            mes=int(data.get("mes") or 0),
            responsaveis=[
                Responsavel.from_dict(item) for item in data.get("responsaveis", [])
            ],
            lancamentos=[
                LancamentoPlantao.from_dict(item) for item in data.get("lancamentos", [])
            ],
            atribuicoes_semanais=data.get("atribuicoes_semanais", {}),
            caminho_arquivo=caminho_arquivo,
        )
