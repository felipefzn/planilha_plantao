from __future__ import annotations

from dataclasses import asdict, dataclass, field
from uuid import uuid4


@dataclass(slots=True)
class Responsavel:
    nome: str
    id: str = field(default_factory=lambda: str(uuid4()))

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict) -> "Responsavel":
        return cls(
            id=data.get("id") or str(uuid4()),
            nome=(data.get("nome") or "").strip(),
        )
