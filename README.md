# Sistema de Plantão Corporativo

Aplicação desktop Windows em Python para gerar, editar, salvar e exportar planilhas de plantão com interface gráfica em `CustomTkinter`.

## Estrutura do projeto

```text
planilha_plantao/
├── app/
│   ├── assets/
│   │   └── README.md
│   ├── config/
│   │   ├── __init__.py
│   │   └── paths.py
│   ├── models/
│   │   ├── __init__.py
│   │   ├── configuracao_sistema.py
│   │   ├── lancamento_plantao.py
│   │   ├── projeto_plantao.py
│   │   └── responsavel.py
│   ├── services/
│   │   ├── __init__.py
│   │   ├── config_service.py
│   │   ├── excel_export_service.py
│   │   ├── project_service.py
│   │   └── schedule_service.py
│   ├── ui/
│   │   ├── __init__.py
│   │   ├── dialogs.py
│   │   ├── main_window.py
│   │   └── styles.py
│   ├── utils/
│   │   ├── __init__.py
│   │   ├── datetime_utils.py
│   │   ├── formatters.py
│   │   └── validators.py
│   └── __init__.py
├── data/
│   ├── README.md
│   └── sample_project.json
├── exports/
│   └── README.md
├── config.json
├── main.py
├── PlanilhaPlantao.spec
├── README.md
└── requirements.txt
```

## Requisitos

1. Windows 10 ou superior.
2. Python 3.11 ou 3.12 instalado a partir de [python.org](https://www.python.org/downloads/windows/).
3. Durante a instalação do Python, marque a opção `Add Python to PATH`.

## Instalação das dependências

No PowerShell, dentro da pasta do projeto:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
```

Se a sua máquina usar o launcher do Python:

```powershell
py -3 -m venv .venv
.\.venv\Scripts\Activate.ps1
py -3 -m pip install --upgrade pip
py -3 -m pip install -r requirements.txt
```

## Como rodar localmente

```powershell
.\.venv\Scripts\Activate.ps1
python main.py
```

Ou:

```powershell
.\.venv\Scripts\Activate.ps1
py -3 main.py
```

## Como gerar o executável `.exe`

Com o ambiente virtual ativado:

```powershell
pyinstaller --noconfirm --clean PlanilhaPlantao.spec
```

O executável será criado em:

```text
dist\PlanilhaPlantao.exe
```

## Observações de empacotamento no Windows

1. O `PlanilhaPlantao.spec` já embute `config.json` e `data/sample_project.json`.
2. Na primeira execução do `.exe`, o sistema cria automaticamente as pastas `data` e `exports` ao lado do executável.
3. Se `config.json` não existir ao lado do `.exe`, ele será copiado automaticamente.
4. Se o Excel estiver com o arquivo aberto no momento da exportação, o sistema mostra uma mensagem amigável e a exportação é interrompida com segurança.
5. Para distribuição interna, envie a pasta `dist` gerada pelo PyInstaller.

## Como usar no dia a dia

1. Abra o sistema.
2. Escolha o mês e o ano.
3. Cadastre ou revise os responsáveis.
4. Selecione uma semana no calendário ou na grade de semanas.
5. Escolha o responsável da semana e clique em `Atribuir responsável`.
6. Clique em `Gerar plantão automático` para criar os lançamentos padrão de sobreaviso.
7. Use `Adicionar suporte` para registrar ocorrências urgentes reais.
8. Revise a grade, edite, duplique ou exclua linhas quando necessário.
9. Clique em `Salvar projeto` para persistir o trabalho em JSON.
10. Clique em `Exportar Excel` para gerar a planilha final em `.xlsx`.

## Regras padrão do sistema

- Segunda a sexta: `20:00` até `06:00` do dia seguinte, tipo `Sobre aviso`.
- Sábado: regra configurável em `config.json`. O padrão entregue é cobertura de `24h`.
- Domingo: regra configurável em `config.json`. O padrão entregue é cobertura de `24h`.
- `Sobre aviso`: `horas * 3.51`
- `Suporte`: `horas * 10.5`

## Arquivos importantes para manutenção

- `config.json`: regras do sistema e aparência.
- `data/sample_project.json`: projeto de exemplo para testes.
- `app/services/schedule_service.py`: cálculos de horas, valor e geração automática.
- `app/services/excel_export_service.py`: exportação da planilha Excel corporativa.
- `app/ui/main_window.py`: janela principal.
