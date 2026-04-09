from __future__ import annotations

import json

from app.config.paths import config_path, ensure_runtime_structure
from app.models.configuracao_sistema import ConfiguracaoSistema


DEFAULT_CONFIG = {
    "valor_hora_sobre_aviso": 3.51,
    "valor_hora_suporte": 10.5,
    "horario_padrao_inicio_plantao": "20:00",
    "horario_padrao_fim_plantao": "06:00",
    "regras_padrao_sabado": {
        "habilitado": True,
        "tipo": "Sobre aviso",
        "inicio": "00:00",
        "fim": "00:00",
        "vira_dia_seguinte": True,
    },
    "regras_padrao_domingo": {
        "habilitado": True,
        "tipo": "Sobre aviso",
        "inicio": "00:00",
        "fim": "00:00",
        "vira_dia_seguinte": True,
    },
    "nome_padrao_arquivo_exportado": "plantao_{ano}_{mes}.xlsx",
    "template_excel": "",
    "idioma_calendario": "pt_BR",
    "tema": {
        "cor_primaria": "#1F5AA6",
        "cor_secundaria": "#2F80ED",
        "cor_fundo": "#F3F7FC",
        "cor_card": "#FFFFFF",
        "cor_texto": "#183153",
        "cor_texto_secundario": "#5B6B82",
    },
}


class ConfigService:
    def load(self) -> ConfiguracaoSistema:
        ensure_runtime_structure()
        caminho = config_path()
        try:
            with caminho.open("r", encoding="utf-8") as arquivo:
                data = json.load(arquivo)
        except (FileNotFoundError, json.JSONDecodeError):
            data = DEFAULT_CONFIG
            self.save(ConfiguracaoSistema.from_dict(data))
        return ConfiguracaoSistema.from_dict(data)

    def save(self, config: ConfiguracaoSistema) -> None:
        ensure_runtime_structure()
        with config_path().open("w", encoding="utf-8") as arquivo:
            json.dump(config.to_dict(), arquivo, ensure_ascii=False, indent=2)
