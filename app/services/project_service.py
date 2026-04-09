from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

from app.models import ProjetoPlantao, Responsavel
from app.utils.validators import ValidacaoErro, validar_nome_responsavel


class ProjectService:
    def create_empty_project(self, mes: int, ano: int) -> ProjetoPlantao:
        return ProjetoPlantao(
            nome=f"Plantão {mes:02d}/{ano}",
            ano=ano,
            mes=mes,
            responsaveis=[],
            lancamentos=[],
            atribuicoes_semanais={},
        )

    def load_project(self, path: str | Path) -> ProjetoPlantao:
        caminho = Path(path)
        with caminho.open("r", encoding="utf-8") as arquivo:
            data = json.load(arquivo)
        projeto = ProjetoPlantao.from_dict(data, caminho_arquivo=str(caminho))
        if not projeto.ano or not projeto.mes:
            hoje = datetime.now()
            projeto.ano = hoje.year
            projeto.mes = hoje.month
        return projeto

    def save_project(self, projeto: ProjetoPlantao, path: str | Path) -> None:
        caminho = Path(path)
        caminho.parent.mkdir(parents=True, exist_ok=True)
        with caminho.open("w", encoding="utf-8") as arquivo:
            json.dump(projeto.to_dict(), arquivo, ensure_ascii=False, indent=2)
        projeto.caminho_arquivo = str(caminho)

    def add_responsavel(self, projeto: ProjetoPlantao, nome: str) -> Responsavel:
        nome_limpo = validar_nome_responsavel(nome)
        if any(item.nome.lower() == nome_limpo.lower() for item in projeto.responsaveis):
            raise ValidacaoErro("Já existe um responsável com esse nome.")
        responsavel = Responsavel(nome=nome_limpo)
        projeto.responsaveis.append(responsavel)
        projeto.responsaveis.sort(key=lambda item: item.nome.lower())
        return responsavel

    def update_responsavel(
        self, projeto: ProjetoPlantao, responsavel_id: str, novo_nome: str
    ) -> None:
        nome_limpo = validar_nome_responsavel(novo_nome)
        if any(
            item.nome.lower() == nome_limpo.lower() and item.id != responsavel_id
            for item in projeto.responsaveis
        ):
            raise ValidacaoErro("Já existe outro responsável com esse nome.")

        responsavel = self.get_responsavel(projeto, responsavel_id)
        responsavel.nome = nome_limpo

        for lancamento in projeto.lancamentos:
            if lancamento.responsavel_id == responsavel_id:
                lancamento.responsavel = nome_limpo

        for atribuicao in projeto.atribuicoes_semanais.values():
            if atribuicao.get("responsavel_id") == responsavel_id:
                atribuicao["responsavel"] = nome_limpo

        projeto.responsaveis.sort(key=lambda item: item.nome.lower())

    def remove_responsavel(self, projeto: ProjetoPlantao, responsavel_id: str) -> None:
        if any(item.responsavel_id == responsavel_id for item in projeto.lancamentos):
            raise ValidacaoErro(
                "Este responsável possui lançamentos e não pode ser removido."
            )
        if any(
            atribuicao.get("responsavel_id") == responsavel_id
            for atribuicao in projeto.atribuicoes_semanais.values()
        ):
            raise ValidacaoErro(
                "Este responsável está atribuído a uma semana e não pode ser removido."
            )
        projeto.responsaveis = [
            item for item in projeto.responsaveis if item.id != responsavel_id
        ]

    def get_responsavel(self, projeto: ProjetoPlantao, responsavel_id: str) -> Responsavel:
        for responsavel in projeto.responsaveis:
            if responsavel.id == responsavel_id:
                return responsavel
        raise ValidacaoErro("Responsável não encontrado.")
