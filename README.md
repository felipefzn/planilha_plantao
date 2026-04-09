# Sistema de Plantao Corporativo

Aplicacao desktop Windows para gerar, editar, salvar e exportar planilhas de plantao com interface grafica em Python.

O sistema foi pensado para equipes que precisam controlar sobreaviso e suporte fora do horario comercial sem depender de preenchimento manual em Excel.

## Visao geral

Com a aplicacao, o fluxo fica assim:

1. selecionar mes e ano
2. cadastrar os responsaveis
3. atribuir responsavel por semana
4. gerar os lancamentos automaticos de sobreaviso
5. registrar suportes urgentes manualmente
6. revisar a grade
7. salvar o projeto em JSON
8. exportar o Excel final

## Funcionalidades

- interface desktop com `CustomTkinter`
- calendario mensal
- atribuicao de responsavel por semana
- geracao automatica de plantao semanal
- geracao do mes inteiro com um clique
- cadastro, edicao e remocao de responsaveis
- cadastro manual de ocorrencias de suporte
- calculo automatico de horas e valor
- tratamento de virada de dia
- filtros e ordenacao na grade de lancamentos
- persistencia em JSON
- exportacao para Excel `.xlsx`
- abas extras de resumo no Excel
- empacotamento para `.exe` com `PyInstaller`

## Regras de negocio

Horario base considerado:

- expediente normal: `08:00` as `18:00`
- descanso: `18:00` as `20:00`
- plantao fora do horario comercial: `20:00` as `06:00` do dia seguinte
- descanso: `06:00` as `08:00`

Tipos de lancamento:

- `Sobre aviso`: prontidao sem atendimento efetivo
- `Suporte`: atendimento urgente realizado

Calculo de valor:

- `Sobre aviso`: `horas * 3.51`
- `Suporte`: `horas * 10.5`

Configuracoes padrao podem ser alteradas em [`config.json`](./config.json).

## Stack

- Python 3
- CustomTkinter
- tkcalendar
- openpyxl
- JSON
- PyInstaller

## Estrutura do projeto

```text
planilha_plantao/
|-- app/
|   |-- assets/
|   |-- config/
|   |-- models/
|   |-- services/
|   |-- ui/
|   |-- utils/
|   `-- __init__.py
|-- data/
|-- exports/
|-- config.json
|-- main.py
|-- PlanilhaPlantao.spec
|-- README.md
`-- requirements.txt
```

## Requisitos

- Windows 10 ou superior
- Python 3.11 ou 3.12 instalado
- opcao `Add Python to PATH` habilitada na instalacao

## Instalacao

No PowerShell, dentro da pasta do projeto:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
```

Se sua maquina usar `py` em vez de `python`:

```powershell
py -3 -m venv .venv
.\.venv\Scripts\Activate.ps1
py -3 -m pip install --upgrade pip
py -3 -m pip install -r requirements.txt
```

## Executando localmente

```powershell
.\.venv\Scripts\Activate.ps1
python main.py
```

Ou:

```powershell
.\.venv\Scripts\Activate.ps1
py -3 main.py
```

## Executando no VSCode

- selecione o interpretador Python do projeto
- ative o ambiente virtual
- rode `main.py` pelo terminal ou com `F5`

O repositorio ja inclui uma configuracao basica em [`.vscode/launch.json`](./.vscode/launch.json).

## Como usar

### 1. Configurar o periodo

- escolha o mes
- escolha o ano
- use o calendario para localizar a semana

### 2. Definir responsaveis

- adicione os nomes na area de cadastro
- selecione a semana
- atribua o responsavel da semana

### 3. Gerar o plantao

- `Gerar semana`: cria os lancamentos automaticos da semana selecionada
- `Gerar mes`: gera todas as semanas do mes com base nas atribuicoes

### 4. Registrar suporte manual

No modal `Adicionar Suporte`, informe:

- responsavel
- data
- hora inicio
- hora fim ou quantidade de horas
- numero do chamado
- solicitante
- cliente
- nivel
- observacao

### 5. Revisar e exportar

- filtre a grade por responsavel, mes, tipo e cliente
- edite, duplique ou exclua linhas
- salve o projeto em JSON
- exporte para Excel

## Exportacao Excel

O arquivo exportado contem:

- aba `Lancamentos`
- aba `Resumo Responsavel`
- aba `Resumo Mes Tipo`

Recursos aplicados no `.xlsx`:

- cabecalho estilizado
- filtros na linha de cabecalho
- congelamento da primeira linha
- formatacao de moeda
- ajuste de largura de colunas

## Gerando o executavel `.exe`

Com o ambiente virtual ativado:

```powershell
pyinstaller --noconfirm --clean PlanilhaPlantao.spec
```

Saida esperada:

```text
dist\PlanilhaPlantao.exe
```

## Observacoes de empacotamento

- o arquivo [`PlanilhaPlantao.spec`](./PlanilhaPlantao.spec) ja inclui `config.json` e `data/sample_project.json`
- na primeira execucao do `.exe`, as pastas `data` e `exports` sao criadas automaticamente
- se o Excel estiver aberto com o arquivo de destino, a exportacao falha com mensagem amigavel

## Arquivos importantes

- [`main.py`](./main.py): ponto de entrada
- [`app/ui/main_window.py`](./app/ui/main_window.py): janela principal
- [`app/ui/dialogs.py`](./app/ui/dialogs.py): modais de cadastro e edicao
- [`app/services/schedule_service.py`](./app/services/schedule_service.py): regras de geracao e calculo
- [`app/services/excel_export_service.py`](./app/services/excel_export_service.py): exportacao para Excel
- [`config.json`](./config.json): parametros do sistema
- [`data/sample_project.json`](./data/sample_project.json): projeto de exemplo

## Estado atual

O projeto entrega:

- cadastro de responsaveis
- geracao semanal e mensal
- lancamentos manuais de suporte
- persistencia em JSON
- exportacao consolidada em Excel

## Licenca

Uso interno. Ajuste conforme a politica do repositorio.
