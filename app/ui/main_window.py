from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk
from tkcalendar import Calendar

from app.config.paths import data_dir, exports_dir, sample_project_path
from app.models import LancamentoPlantao
from app.services.config_service import ConfigService
from app.services.excel_export_service import ExcelExportService, ExportacaoErro
from app.services.project_service import ProjectService
from app.services.schedule_service import ScheduleService
from app.ui.dialogs import LancamentoDialog, ResponsavelDialog
from app.ui.styles import card_kwargs, setup_theme
from app.utils.datetime_utils import (
    MESES_PT,
    encontrar_semana_por_data,
    nome_mes,
    obter_semanas_do_mes,
)
from app.utils.formatters import formatar_data_br, formatar_horas, formatar_moeda_br
from app.utils.validators import ValidacaoErro


class MainWindow(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.config_service = ConfigService()
        self.project_service = ProjectService()
        self.schedule_service = ScheduleService()
        self.excel_service = ExcelExportService()
        self.configuracao = self.config_service.load()
        self.tema = self.configuracao.tema

        setup_theme(self, self.tema)

        self.title("Sistema de Plantão Corporativo")
        self.geometry("1540x940")
        self.minsize(1380, 860)
        self.configure(fg_color=self.tema.get("cor_fundo", "#F3F7FC"))

        hoje = datetime.now()
        self.projeto = self.project_service.create_empty_project(hoje.month, hoje.year)
        self._seed_default_responsaveis()

        self.meses_por_nome = {nome: indice for indice, nome in MESES_PT.items()}
        self.ordem_meses = {nome: indice for indice, nome in MESES_PT.items()}
        self.colunas_tabela = {
            "responsavel": "Responsável",
            "mes": "Mês",
            "dia_semana": "Dia Semana",
            "inicio": "Início",
            "fim": "Fim",
            "data": "Data",
            "tipo": "Tipo",
            "total_horas": "Total Horas",
            "valor": "Valor",
            "numero_chamado": "Nº Chamado",
            "solicitante": "Solicitante",
            "cliente": "Cliente",
            "nivel": "Nível",
            "observacao": "Observação",
        }
        self.sort_column = "data"
        self.sort_desc = False
        self.dirty = False
        self.weeks: list[dict] = []

        self.titulo_var = ctk.StringVar()
        self.subtitulo_var = ctk.StringVar()
        self.status_var = ctk.StringVar(value="Sistema pronto para uso.")
        self.file_var = ctk.StringVar(value="Projeto atual: novo projeto não salvo")
        self.total_lancamentos_var = ctk.StringVar(value="0")
        self.total_horas_var = ctk.StringVar(value="0,00 h")
        self.total_valor_var = ctk.StringVar(value="R$ 0,00")
        self.total_suportes_var = ctk.StringVar(value="0")
        self.mes_var = ctk.StringVar(value=nome_mes(self.projeto.mes))
        self.ano_var = ctk.StringVar(value=str(self.projeto.ano))
        self.week_responsavel_var = ctk.StringVar(value="")
        self.filter_responsavel_var = ctk.StringVar(value="Todos")
        self.filter_mes_var = ctk.StringVar(value=nome_mes(self.projeto.mes))
        self.filter_tipo_var = ctk.StringVar(value="Todos")
        self.filter_cliente_var = ctk.StringVar(value="")

        self._build_layout()
        self._bind_events()
        self._refresh_all(reset_filters=True)
        self._update_title()

    def _build_layout(self) -> None:
        self.grid_rowconfigure(3, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self._build_header()
        self._build_top_cards()
        self._build_action_bar()
        self._build_table_area()
        self._build_status_bar()

    def _build_header(self) -> None:
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=20, pady=(18, 10))
        header.grid_columnconfigure(0, weight=1)

        left = ctk.CTkFrame(header, fg_color="transparent")
        left.grid(row=0, column=0, sticky="ew")

        ctk.CTkLabel(
            left,
            textvariable=self.titulo_var,
            font=("Segoe UI Semibold", 28),
            text_color=self.tema.get("cor_texto", "#183153"),
            anchor="w",
        ).grid(row=0, column=0, sticky="ew")
        ctk.CTkLabel(
            left,
            textvariable=self.subtitulo_var,
            font=("Segoe UI", 13),
            text_color=self.tema.get("cor_texto_secundario", "#5B6B82"),
            anchor="w",
        ).grid(row=1, column=0, sticky="ew", pady=(6, 0))
        ctk.CTkLabel(
            left,
            textvariable=self.file_var,
            font=("Segoe UI", 12),
            text_color="#6B7A92",
            anchor="w",
        ).grid(row=2, column=0, sticky="ew", pady=(6, 0))

        right = ctk.CTkFrame(header, fg_color="transparent")
        right.grid(row=0, column=1, sticky="e")
        buttons = [
            ("Abrir exemplo", self._open_sample_project, "#6C8FBE"),
            ("Abrir projeto", self._open_project, "#6C8FBE"),
            ("Salvar projeto", self._save_project, "#2F80ED"),
            ("Exportar Excel", self._export_excel, self.tema.get("cor_primaria", "#1F5AA6")),
        ]
        for indice, (texto, comando, cor) in enumerate(buttons):
            ctk.CTkButton(
                right,
                text=texto,
                command=comando,
                fg_color=cor,
                height=38,
                width=132,
            ).grid(row=0, column=indice, padx=(10 if indice else 0, 0))

    def _build_top_cards(self) -> None:
        area = ctk.CTkFrame(self, fg_color="transparent")
        area.grid(row=1, column=0, sticky="nsew", padx=20, pady=(0, 10))
        area.grid_columnconfigure(0, weight=2)
        area.grid_columnconfigure(1, weight=2)
        area.grid_columnconfigure(2, weight=1)

        self._build_period_card(area)
        self._build_weeks_card(area)
        self._build_responsaveis_card(area)

    def _build_period_card(self, parent) -> None:
        card = ctk.CTkFrame(parent, **card_kwargs(self.tema))
        card.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        for indice in range(2):
            card.grid_columnconfigure(indice, weight=1)

        ctk.CTkLabel(
            card,
            text="Período e Calendário",
            font=("Segoe UI Semibold", 18),
            anchor="w",
        ).grid(row=0, column=0, columnspan=2, sticky="ew", padx=18, pady=(16, 4))
        ctk.CTkLabel(
            card,
            text="Escolha o mês, o ano e navegue pelo calendário para localizar a semana do plantão.",
            anchor="w",
            text_color="#5B6B82",
        ).grid(row=1, column=0, columnspan=2, sticky="ew", padx=18, pady=(0, 14))

        ctk.CTkLabel(card, text="Mês", anchor="w").grid(
            row=2, column=0, sticky="ew", padx=(18, 9)
        )
        meses = list(MESES_PT.values())
        self.combo_mes = ctk.CTkComboBox(
            card,
            values=meses,
            variable=self.mes_var,
            state="readonly",
            command=lambda _: self._change_project_period(),
            height=38,
        )
        self.combo_mes.grid(row=3, column=0, sticky="ew", padx=(18, 9), pady=(4, 12))

        ctk.CTkLabel(card, text="Ano", anchor="w").grid(
            row=2, column=1, sticky="ew", padx=(9, 18)
        )
        anos = [str(ano) for ano in range(datetime.now().year - 3, datetime.now().year + 6)]
        self.combo_ano = ctk.CTkComboBox(
            card,
            values=anos,
            variable=self.ano_var,
            state="readonly",
            command=lambda _: self._change_project_period(),
            height=38,
        )
        self.combo_ano.grid(row=3, column=1, sticky="ew", padx=(9, 18), pady=(4, 12))

        self.calendar_frame = ctk.CTkFrame(card, fg_color="transparent")
        self.calendar_frame.grid(row=4, column=0, columnspan=2, padx=18, pady=(0, 14), sticky="nsew")
        self.calendar = Calendar(
            self.calendar_frame,
            selectmode="day",
            year=self.projeto.ano,
            month=self.projeto.mes,
            locale=self.configuracao.idioma_calendario,
            date_pattern="dd/MM/yyyy",
            background=self.tema.get("cor_primaria", "#1F5AA6"),
            foreground="white",
            headersbackground=self.tema.get("cor_primaria", "#1F5AA6"),
            headersforeground="white",
            selectbackground=self.tema.get("cor_secundaria", "#2F80ED"),
            weekendbackground="#F7FAFE",
            normalbackground="#FFFFFF",
            othermonthbackground="#F2F6FB",
            othermonthforeground="#A3B1C5",
            normalforeground=self.tema.get("cor_texto", "#183153"),
            weekendforeground=self.tema.get("cor_texto", "#183153"),
            bordercolor="#D8E4F4",
        )
        self.calendar.pack(fill="both", expand=True)

        ctk.CTkButton(
            card,
            text="Novo projeto para este período",
            command=self._new_project_for_period,
            fg_color="#EAF2FD",
            text_color=self.tema.get("cor_primaria", "#1F5AA6"),
            hover_color="#DCEAFB",
        ).grid(row=5, column=0, columnspan=2, sticky="ew", padx=18, pady=(0, 18))

    def _build_weeks_card(self, parent) -> None:
        card = ctk.CTkFrame(parent, **card_kwargs(self.tema))
        card.grid(row=0, column=1, sticky="nsew", padx=10)
        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(2, weight=1)

        ctk.CTkLabel(
            card,
            text="Semanas do mês",
            font=("Segoe UI Semibold", 18),
            anchor="w",
        ).grid(row=0, column=0, sticky="ew", padx=18, pady=(16, 4))
        ctk.CTkLabel(
            card,
            text="Atribua o responsável da semana e gere automaticamente os lançamentos padrão.",
            text_color="#5B6B82",
            anchor="w",
        ).grid(row=1, column=0, sticky="ew", padx=18, pady=(0, 14))

        tree_frame = ctk.CTkFrame(card, fg_color="transparent")
        tree_frame.grid(row=2, column=0, sticky="nsew", padx=18)
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        self.weeks_tree = ttk.Treeview(
            tree_frame,
            columns=("semana", "periodo", "responsavel"),
            show="headings",
            style="Corporate.Treeview",
            height=7,
        )
        for coluna, titulo, largura in (
            ("semana", "Semana", 100),
            ("periodo", "Período", 170),
            ("responsavel", "Responsável", 170),
        ):
            self.weeks_tree.heading(coluna, text=titulo)
            self.weeks_tree.column(coluna, width=largura, stretch=True, anchor="w")
        self.weeks_tree.grid(row=0, column=0, sticky="nsew")

        scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.weeks_tree.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.weeks_tree.configure(yscrollcommand=scroll.set)

        ctk.CTkLabel(card, text="Responsável da semana", anchor="w").grid(
            row=3, column=0, sticky="ew", padx=18, pady=(14, 0)
        )
        self.week_responsavel_combo = ctk.CTkComboBox(
            card,
            values=[""],
            variable=self.week_responsavel_var,
            state="readonly",
            height=38,
        )
        self.week_responsavel_combo.grid(row=4, column=0, sticky="ew", padx=18, pady=(6, 14))

        actions = ctk.CTkFrame(card, fg_color="transparent")
        actions.grid(row=5, column=0, sticky="ew", padx=18, pady=(0, 18))
        for indice in range(3):
            actions.grid_columnconfigure(indice, weight=1)
        ctk.CTkButton(
            actions,
            text="Atribuir responsável",
            command=self._assign_responsavel_semana,
            fg_color="#2F80ED",
        ).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ctk.CTkButton(
            actions,
            text="Gerar semana",
            command=self._generate_plantao,
            fg_color=self.tema.get("cor_primaria", "#1F5AA6"),
        ).grid(row=0, column=1, sticky="ew", padx=6)
        ctk.CTkButton(
            actions,
            text="Gerar mês",
            command=self._generate_mes,
            fg_color="#1C6BB4",
        ).grid(row=0, column=2, sticky="ew", padx=(6, 0))

    def _build_responsaveis_card(self, parent) -> None:
        card = ctk.CTkFrame(parent, **card_kwargs(self.tema))
        card.grid(row=0, column=2, sticky="nsew", padx=(10, 0))
        card.grid_rowconfigure(2, weight=1)
        card.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            card,
            text="Cadastro de responsáveis",
            font=("Segoe UI Semibold", 18),
            anchor="w",
        ).grid(row=0, column=0, sticky="ew", padx=18, pady=(16, 4))
        ctk.CTkLabel(
            card,
            text="Gerencie rapidamente quem pode assumir o plantão.",
            text_color="#5B6B82",
            anchor="w",
        ).grid(row=1, column=0, sticky="ew", padx=18, pady=(0, 14))

        frame_lista = ctk.CTkFrame(card, fg_color="transparent")
        frame_lista.grid(row=2, column=0, sticky="nsew", padx=18)
        frame_lista.grid_rowconfigure(0, weight=1)
        frame_lista.grid_columnconfigure(0, weight=1)

        self.responsaveis_tree = ttk.Treeview(
            frame_lista,
            columns=("nome",),
            show="headings",
            style="Corporate.Treeview",
            height=7,
        )
        self.responsaveis_tree.heading("nome", text="Responsável")
        self.responsaveis_tree.column("nome", width=220, stretch=True, anchor="w")
        self.responsaveis_tree.grid(row=0, column=0, sticky="nsew")

        scroll = ttk.Scrollbar(
            frame_lista, orient="vertical", command=self.responsaveis_tree.yview
        )
        scroll.grid(row=0, column=1, sticky="ns")
        self.responsaveis_tree.configure(yscrollcommand=scroll.set)

        actions = ctk.CTkFrame(card, fg_color="transparent")
        actions.grid(row=3, column=0, sticky="ew", padx=18, pady=(14, 18))
        for indice in range(3):
            actions.grid_columnconfigure(indice, weight=1)
        ctk.CTkButton(actions, text="Adicionar", command=self._add_responsavel).grid(
            row=0, column=0, sticky="ew", padx=(0, 6)
        )
        ctk.CTkButton(actions, text="Editar", command=self._edit_responsavel).grid(
            row=0, column=1, sticky="ew", padx=6
        )
        ctk.CTkButton(
            actions,
            text="Remover",
            command=self._remove_responsavel,
            fg_color="#C75D5D",
            hover_color="#B44C4C",
        ).grid(row=0, column=2, sticky="ew", padx=(6, 0))

    def _build_action_bar(self) -> None:
        card = ctk.CTkFrame(self, **card_kwargs(self.tema))
        card.grid(row=2, column=0, sticky="ew", padx=20, pady=(0, 10))
        card.grid_columnconfigure(0, weight=1)

        top = ctk.CTkFrame(card, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", padx=18, pady=(14, 10))
        top.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            top,
            text="Lançamentos e filtros",
            font=("Segoe UI Semibold", 18),
            anchor="w",
        ).grid(row=0, column=0, sticky="ew")

        action_buttons = ctk.CTkFrame(top, fg_color="transparent")
        action_buttons.grid(row=0, column=1, sticky="e")
        botoes = [
            ("Adicionar suporte", self._add_support, self.tema.get("cor_primaria", "#1F5AA6")),
            ("Editar", self._edit_selected_entry, "#2F80ED"),
            ("Duplicar", self._duplicate_selected_entry, "#6C8FBE"),
            ("Excluir", self._delete_selected_entry, "#C75D5D"),
        ]
        for indice, (texto, comando, cor) in enumerate(botoes):
            ctk.CTkButton(
                action_buttons,
                text=texto,
                command=comando,
                fg_color=cor,
                height=36,
                width=128,
            ).grid(row=0, column=indice, padx=(8 if indice else 0, 0))

        filters = ctk.CTkFrame(card, fg_color="#F7FAFE", corner_radius=14)
        filters.grid(row=1, column=0, sticky="ew", padx=18, pady=(0, 16))
        for indice in range(10):
            filters.grid_columnconfigure(indice, weight=1)

        ctk.CTkLabel(filters, text="Responsável", anchor="w").grid(
            row=0, column=0, sticky="ew", padx=(14, 8), pady=(10, 4)
        )
        self.filter_responsavel_combo = ctk.CTkComboBox(
            filters,
            values=["Todos"],
            variable=self.filter_responsavel_var,
            state="readonly",
            command=lambda _: self._apply_filters(),
            height=34,
        )
        self.filter_responsavel_combo.grid(row=1, column=0, sticky="ew", padx=(14, 8), pady=(0, 10))

        ctk.CTkLabel(filters, text="Mês", anchor="w").grid(
            row=0, column=1, sticky="ew", padx=8, pady=(10, 4)
        )
        self.filter_mes_combo = ctk.CTkComboBox(
            filters,
            values=["Todos"],
            variable=self.filter_mes_var,
            state="readonly",
            command=lambda _: self._apply_filters(),
            height=34,
        )
        self.filter_mes_combo.grid(row=1, column=1, sticky="ew", padx=8, pady=(0, 10))

        ctk.CTkLabel(filters, text="Tipo", anchor="w").grid(
            row=0, column=2, sticky="ew", padx=8, pady=(10, 4)
        )
        self.filter_tipo_combo = ctk.CTkComboBox(
            filters,
            values=["Todos", "Sobre aviso", "Suporte"],
            variable=self.filter_tipo_var,
            state="readonly",
            command=lambda _: self._apply_filters(),
            height=34,
        )
        self.filter_tipo_combo.grid(row=1, column=2, sticky="ew", padx=8, pady=(0, 10))

        ctk.CTkLabel(filters, text="Cliente", anchor="w").grid(
            row=0, column=3, columnspan=2, sticky="ew", padx=8, pady=(10, 4)
        )
        self.filter_cliente_entry = ctk.CTkEntry(
            filters, textvariable=self.filter_cliente_var, height=34
        )
        self.filter_cliente_entry.grid(
            row=1, column=3, columnspan=2, sticky="ew", padx=8, pady=(0, 10)
        )

        ctk.CTkButton(
            filters,
            text="Limpar filtros",
            command=self._clear_filters,
            fg_color="#E3ECFA",
            text_color=self.tema.get("cor_primaria", "#1F5AA6"),
            hover_color="#D7E6FB",
            height=34,
        ).grid(row=1, column=5, sticky="ew", padx=(8, 14), pady=(0, 10))

    def _build_table_area(self) -> None:
        card = ctk.CTkFrame(self, **card_kwargs(self.tema))
        card.grid(row=3, column=0, sticky="nsew", padx=20, pady=(0, 10))
        card.grid_rowconfigure(1, weight=1)
        card.grid_columnconfigure(0, weight=1)

        summary = ctk.CTkFrame(card, fg_color="transparent")
        summary.grid(row=0, column=0, sticky="ew", padx=18, pady=(16, 12))
        for indice in range(4):
            summary.grid_columnconfigure(indice, weight=1)

        self._create_summary_box(summary, 0, "Total de lançamentos", self.total_lancamentos_var)
        self._create_summary_box(summary, 1, "Horas visíveis", self.total_horas_var)
        self._create_summary_box(summary, 2, "Valor visível", self.total_valor_var)
        self._create_summary_box(summary, 3, "Ocorrências de suporte", self.total_suportes_var)

        table_wrap = ctk.CTkFrame(card, fg_color="transparent")
        table_wrap.grid(row=1, column=0, sticky="nsew", padx=18, pady=(0, 18))
        table_wrap.grid_rowconfigure(0, weight=1)
        table_wrap.grid_columnconfigure(0, weight=1)

        self.entries_tree = ttk.Treeview(
            table_wrap,
            columns=list(self.colunas_tabela),
            show="headings",
            style="Corporate.Treeview",
        )
        larguras = {
            "responsavel": 150,
            "mes": 90,
            "dia_semana": 120,
            "inicio": 75,
            "fim": 75,
            "data": 100,
            "tipo": 110,
            "total_horas": 95,
            "valor": 95,
            "numero_chamado": 120,
            "solicitante": 130,
            "cliente": 130,
            "nivel": 70,
            "observacao": 220,
        }
        for coluna, titulo in self.colunas_tabela.items():
            self.entries_tree.heading(
                coluna, text=titulo, command=lambda c=coluna: self._toggle_sort(c)
            )
            self.entries_tree.column(
                coluna,
                width=larguras[coluna],
                stretch=True,
                anchor="w",
            )
        self.entries_tree.tag_configure("automatico", background="#F7FBFF")
        self.entries_tree.tag_configure("manual", background="#FFFFFF")
        self.entries_tree.grid(row=0, column=0, sticky="nsew")

        scroll_y = ttk.Scrollbar(table_wrap, orient="vertical", command=self.entries_tree.yview)
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x = ttk.Scrollbar(table_wrap, orient="horizontal", command=self.entries_tree.xview)
        scroll_x.grid(row=1, column=0, sticky="ew")
        self.entries_tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

    def _create_summary_box(self, parent, column: int, titulo: str, variable) -> None:
        box = ctk.CTkFrame(parent, fg_color="#F7FAFE", corner_radius=14)
        box.grid(row=0, column=column, sticky="ew", padx=(0 if column == 0 else 6, 6 if column < 3 else 0))
        ctk.CTkLabel(box, text=titulo, anchor="w", text_color="#5B6B82").pack(
            anchor="w", padx=14, pady=(10, 4)
        )
        ctk.CTkLabel(
            box,
            textvariable=variable,
            font=("Segoe UI Semibold", 22),
            anchor="w",
            text_color=self.tema.get("cor_texto", "#183153"),
        ).pack(anchor="w", padx=14, pady=(0, 12))

    def _build_status_bar(self) -> None:
        ctk.CTkLabel(
            self,
            textvariable=self.status_var,
            anchor="w",
            font=("Segoe UI", 12),
            text_color="#5B6B82",
        ).grid(row=4, column=0, sticky="ew", padx=24, pady=(0, 16))

    def _bind_events(self) -> None:
        self.calendar.bind("<<CalendarSelected>>", self._on_calendar_selected)
        self.weeks_tree.bind("<<TreeviewSelect>>", lambda *_: self._sync_week_assignment_combo())
        self.entries_tree.bind("<Double-1>", lambda *_: self._edit_selected_entry())
        self.filter_cliente_entry.bind("<KeyRelease>", lambda *_: self._apply_filters())
        self.bind("<Delete>", lambda *_: self._delete_selected_entry())
        self.bind("<Control-s>", lambda *_: self._save_project())
        self.bind("<Control-o>", lambda *_: self._open_project())
        self.bind("<Control-e>", lambda *_: self._export_excel())
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _seed_default_responsaveis(self) -> None:
        if self.projeto.responsaveis:
            return
        for nome in ("Felipe Brito", "Isaak Ferreira"):
            try:
                self.project_service.add_responsavel(self.projeto, nome)
            except ValidacaoErro:
                continue

    def _refresh_all(self, reset_filters: bool = False) -> None:
        self._update_header_texts()
        self._update_file_label()
        self._refresh_period_controls()
        self._refresh_responsaveis()
        self._refresh_weeks()
        self._refresh_filter_options(reset_filters=reset_filters)
        self._apply_filters()
        self._update_title()

    def _update_header_texts(self) -> None:
        self.titulo_var.set("Planilha de Plantão Corporativa")
        self.subtitulo_var.set(
            f"Projeto: {self.projeto.nome} | Período de referência: {nome_mes(self.projeto.mes)} de {self.projeto.ano}"
        )

    def _update_file_label(self) -> None:
        if self.projeto.caminho_arquivo:
            self.file_var.set(f"Projeto atual: {self.projeto.caminho_arquivo}")
        else:
            self.file_var.set("Projeto atual: novo projeto não salvo")

    def _refresh_period_controls(self) -> None:
        self.mes_var.set(nome_mes(self.projeto.mes))
        self.ano_var.set(str(self.projeto.ano))
        data_atual = self._calendar_selected_date_or_default()
        if data_atual.month != self.projeto.mes or data_atual.year != self.projeto.ano:
            data_atual = date(self.projeto.ano, self.projeto.mes, 1)
        self.calendar.selection_set(data_atual)

    def _refresh_responsaveis(self) -> None:
        self.responsaveis_tree.delete(*self.responsaveis_tree.get_children())
        for responsavel in self.projeto.responsaveis:
            self.responsaveis_tree.insert("", "end", iid=responsavel.id, values=(responsavel.nome,))

        nomes = [item.nome for item in self.projeto.responsaveis] or [""]
        self.week_responsavel_combo.configure(values=nomes)
        if self.projeto.responsaveis and self.week_responsavel_var.get() not in nomes:
            self.week_responsavel_var.set(nomes[0])
        elif not self.projeto.responsaveis:
            self.week_responsavel_var.set("")

    def _refresh_weeks(self) -> None:
        self.weeks = obter_semanas_do_mes(self.projeto.ano, self.projeto.mes)
        self.weeks_tree.delete(*self.weeks_tree.get_children())
        for semana in self.weeks:
            atribuicao = self.projeto.atribuicoes_semanais.get(semana["id"], {})
            responsavel = atribuicao.get("responsavel", "Não atribuído")
            self.weeks_tree.insert(
                "",
                "end",
                iid=semana["id"],
                values=(
                    f"Semana {semana['indice']}",
                    f"{semana['inicio'].strftime('%d/%m')} a {semana['fim'].strftime('%d/%m')}",
                    responsavel,
                ),
            )
        self._select_week_from_calendar()

    def _refresh_filter_options(self, reset_filters: bool = False) -> None:
        responsaveis = ["Todos"] + sorted({item.responsavel for item in self.projeto.lancamentos if item.responsavel})
        if len(responsaveis) == 1:
            responsaveis.extend([item.nome for item in self.projeto.responsaveis if item.nome])
        meses = ["Todos"] + sorted(
            {item.mes for item in self.projeto.lancamentos if item.mes},
            key=lambda nome: self.ordem_meses.get(nome, 99),
        )
        if len(meses) == 1:
            meses.append(nome_mes(self.projeto.mes))

        self.filter_responsavel_combo.configure(values=responsaveis)
        self.filter_mes_combo.configure(values=meses)
        self.filter_tipo_combo.configure(values=["Todos", "Sobre aviso", "Suporte"])

        if reset_filters or self.filter_responsavel_var.get() not in responsaveis:
            self.filter_responsavel_var.set("Todos")
        if reset_filters or self.filter_mes_var.get() not in meses:
            self.filter_mes_var.set(nome_mes(self.projeto.mes))
        if reset_filters:
            self.filter_tipo_var.set("Todos")
            self.filter_cliente_var.set("")

    def _apply_filters(self) -> None:
        entries = list(self.projeto.lancamentos)

        filtro_resp = self.filter_responsavel_var.get().strip()
        filtro_mes = self.filter_mes_var.get().strip()
        filtro_tipo = self.filter_tipo_var.get().strip()
        filtro_cliente = self.filter_cliente_var.get().strip().lower()

        if filtro_resp and filtro_resp != "Todos":
            entries = [item for item in entries if item.responsavel == filtro_resp]
        if filtro_mes and filtro_mes != "Todos":
            entries = [item for item in entries if item.mes == filtro_mes]
        if filtro_tipo and filtro_tipo != "Todos":
            entries = [item for item in entries if item.tipo == filtro_tipo]
        if filtro_cliente:
            entries = [item for item in entries if filtro_cliente in item.cliente.lower()]

        entries = self._sort_entries(entries)
        self._populate_entries(entries)
        self._update_summary(entries)
        self._set_status(f"{len(entries)} lançamento(s) exibido(s) na grade.")

    def _sort_entries(self, entries: list[LancamentoPlantao]) -> list[LancamentoPlantao]:
        def sort_key(item: LancamentoPlantao):
            mapping = {
                "responsavel": item.responsavel.lower(),
                "mes": self.ordem_meses.get(item.mes, 99),
                "dia_semana": item.dia_semana,
                "inicio": item.inicio,
                "fim": item.fim,
                "data": item.data,
                "tipo": item.tipo,
                "total_horas": item.total_horas,
                "valor": item.valor,
                "numero_chamado": item.numero_chamado.lower(),
                "solicitante": item.solicitante.lower(),
                "cliente": item.cliente.lower(),
                "nivel": item.nivel.lower(),
                "observacao": item.observacao.lower(),
            }
            return mapping.get(self.sort_column, item.data)

        return sorted(entries, key=sort_key, reverse=self.sort_desc)

    def _populate_entries(self, entries: list[LancamentoPlantao]) -> None:
        self.entries_tree.delete(*self.entries_tree.get_children())
        for item in entries:
            self.entries_tree.insert(
                "",
                "end",
                iid=item.id,
                values=(
                    item.responsavel,
                    item.mes,
                    item.dia_semana,
                    item.inicio,
                    item.fim,
                    formatar_data_br(item.data),
                    item.tipo,
                    f"{item.total_horas:.2f}",
                    formatar_moeda_br(item.valor),
                    item.numero_chamado,
                    item.solicitante,
                    item.cliente,
                    item.nivel,
                    item.observacao,
                ),
                tags=(item.origem,),
            )

    def _update_summary(self, entries: list[LancamentoPlantao]) -> None:
        total_horas = sum(item.total_horas for item in entries)
        total_valor = sum(item.valor for item in entries)
        total_suportes = sum(1 for item in entries if item.tipo == "Suporte")

        self.total_lancamentos_var.set(str(len(entries)))
        self.total_horas_var.set(formatar_horas(total_horas))
        self.total_valor_var.set(formatar_moeda_br(total_valor))
        self.total_suportes_var.set(str(total_suportes))

    def _set_status(self, mensagem: str) -> None:
        self.status_var.set(mensagem)

    def _update_title(self) -> None:
        sufixo = " *" if self.dirty else ""
        self.title(f"Sistema de Plantão Corporativo{sufixo}")

    def _mark_dirty(self, mensagem: str | None = None) -> None:
        self.dirty = True
        self._update_title()
        if mensagem:
            self._set_status(mensagem)

    def _clear_dirty(self, mensagem: str | None = None) -> None:
        self.dirty = False
        self._update_title()
        if mensagem:
            self._set_status(mensagem)

    def _change_project_period(self) -> None:
        novo_mes = self.meses_por_nome.get(self.mes_var.get(), self.projeto.mes)
        novo_ano = int(self.ano_var.get())
        if novo_mes == self.projeto.mes and novo_ano == self.projeto.ano:
            return
        self.projeto.mes = novo_mes
        self.projeto.ano = novo_ano
        self.projeto.nome = f"Plantão {novo_mes:02d}/{novo_ano}"
        self._refresh_all(reset_filters=True)
        self._mark_dirty("Período do projeto atualizado.")

    def _new_project_for_period(self) -> None:
        if not self._confirm_discard_if_needed():
            return
        mes = self.meses_por_nome.get(self.mes_var.get(), datetime.now().month)
        ano = int(self.ano_var.get())
        self.projeto = self.project_service.create_empty_project(mes, ano)
        self._seed_default_responsaveis()
        self.sort_column = "data"
        self.sort_desc = False
        self.dirty = False
        self._refresh_all(reset_filters=True)
        self._set_status("Novo projeto criado para o período selecionado.")

    def _selected_week(self) -> dict | None:
        selected = self.weeks_tree.selection()
        if selected:
            week_id = selected[0]
            return next((item for item in self.weeks if item["id"] == week_id), None)
        data_calendario = self._calendar_selected_date_or_default()
        return encontrar_semana_por_data(self.projeto.ano, self.projeto.mes, data_calendario)

    def _selected_entry(self) -> LancamentoPlantao | None:
        selected = self.entries_tree.selection()
        if not selected:
            return None
        lancamento_id = selected[0]
        return next((item for item in self.projeto.lancamentos if item.id == lancamento_id), None)

    def _selected_responsavel_id(self) -> str | None:
        nome = self.week_responsavel_var.get().strip()
        for responsavel in self.projeto.responsaveis:
            if responsavel.nome == nome:
                return responsavel.id
        return None

    def _sync_week_assignment_combo(self) -> None:
        semana = self._selected_week()
        if not semana:
            return
        atribuicao = self.projeto.atribuicoes_semanais.get(semana["id"], {})
        if atribuicao.get("responsavel"):
            self.week_responsavel_var.set(atribuicao["responsavel"])

    def _assign_responsavel_semana(self) -> None:
        semana = self._selected_week()
        responsavel_id = self._selected_responsavel_id()
        if not semana:
            messagebox.showwarning("Semana obrigatória", "Selecione uma semana primeiro.", parent=self)
            return
        if not responsavel_id:
            messagebox.showwarning(
                "Responsável obrigatório",
                "Escolha um responsável para atribuir à semana.",
                parent=self,
            )
            return

        responsavel = next(item for item in self.projeto.responsaveis if item.id == responsavel_id)
        self.projeto.atribuicoes_semanais[semana["id"]] = {
            "responsavel_id": responsavel.id,
            "responsavel": responsavel.nome,
            "inicio": semana["inicio"].isoformat(),
            "fim": semana["fim"].isoformat(),
            "label": semana["label"],
        }
        self._refresh_weeks()
        self._select_week_by_id(semana["id"])
        self._mark_dirty(f"Responsável {responsavel.nome} atribuído à {semana['label']}.")

    def _generate_plantao(self) -> None:
        semana = self._selected_week()
        if not semana:
            messagebox.showwarning("Semana obrigatória", "Selecione uma semana para gerar o plantão.", parent=self)
            return

        responsavel_id = self._selected_responsavel_id()
        if not responsavel_id:
            atribuicao = self.projeto.atribuicoes_semanais.get(semana["id"], {})
            responsavel_id = atribuicao.get("responsavel_id")
            if atribuicao.get("responsavel"):
                self.week_responsavel_var.set(atribuicao["responsavel"])
        if not responsavel_id:
            messagebox.showwarning(
                "Responsável obrigatório",
                "Selecione ou atribua um responsável antes de gerar o plantão.",
                parent=self,
            )
            return

        existentes = [
            item
            for item in self.projeto.lancamentos
            if item.origem == "automatico" and item.semana_referencia == semana["id"]
        ]
        if existentes and not messagebox.askyesno(
            "Substituir lançamentos automáticos",
            "Já existem lançamentos automáticos para esta semana. Deseja substituí-los?",
            parent=self,
        ):
            return

        try:
            quantidade = self.schedule_service.generate_weekly_schedule(
                self.projeto, self.configuracao, semana, responsavel_id
            )
        except ValidacaoErro as exc:
            messagebox.showerror("Não foi possível gerar", str(exc), parent=self)
            return

        self._refresh_all(reset_filters=False)
        self._select_week_by_id(semana["id"])
        self._mark_dirty(f"{quantidade} lançamento(s) automáticos gerados para a semana.")

    def _generate_mes(self) -> None:
        if not self.weeks:
            messagebox.showwarning(
                "Semanas indisponíveis",
                "Não foi possível localizar as semanas do mês selecionado.",
                parent=self,
            )
            return

        fallback_responsavel_id = self._selected_responsavel_id() or ""
        semanas_atribuidas = sum(
            1
            for week in self.weeks
            if self.projeto.atribuicoes_semanais.get(week["id"], {}).get("responsavel_id")
        )

        if not semanas_atribuidas and not fallback_responsavel_id:
            messagebox.showwarning(
                "Responsável obrigatório",
                "Atribua responsáveis às semanas ou selecione um responsável para aplicar no mês.",
                parent=self,
            )
            return

        if fallback_responsavel_id:
            nome_responsavel = next(
                (
                    item.nome
                    for item in self.projeto.responsaveis
                    if item.id == fallback_responsavel_id
                ),
                "",
            )
            semanas_sem_responsavel = [
                week
                for week in self.weeks
                if not self.projeto.atribuicoes_semanais.get(week["id"], {}).get("responsavel_id")
            ]
            if semanas_sem_responsavel:
                aplicar = messagebox.askyesno(
                    "Aplicar responsável no mês",
                    f"As semanas sem responsável serão geradas com {nome_responsavel}. Deseja continuar?",
                    parent=self,
                )
                if not aplicar:
                    fallback_responsavel_id = ""

        try:
            quantidade, ignoradas = self.schedule_service.generate_month_schedule(
                self.projeto,
                self.configuracao,
                self.weeks,
                fallback_responsavel_id=fallback_responsavel_id,
            )
        except ValidacaoErro as exc:
            messagebox.showerror("Não foi possível gerar", str(exc), parent=self)
            return

        self._refresh_all(reset_filters=False)
        self._mark_dirty(
            f"{quantidade} lançamento(s) automáticos gerados no mês. "
            f"{ignoradas} semana(s) sem responsável foram ignoradas."
        )

    def _add_support(self) -> None:
        if not self.projeto.responsaveis:
            messagebox.showwarning(
                "Sem responsáveis",
                "Cadastre ao menos um responsável antes de adicionar suporte.",
                parent=self,
            )
            return
        dialog = LancamentoDialog(
            self,
            "Adicionar Suporte",
            self.projeto.responsaveis,
            self.configuracao,
            fixed_type="Suporte",
        )
        self.wait_window(dialog)
        if dialog.result:
            self._save_entry_from_dialog(dialog.result, None)

    def _edit_selected_entry(self) -> None:
        lancamento = self._selected_entry()
        if not lancamento:
            messagebox.showwarning(
                "Seleção obrigatória",
                "Selecione um lançamento na grade para editar.",
                parent=self,
            )
            return
        dialog = LancamentoDialog(
            self,
            "Editar Lançamento",
            self.projeto.responsaveis,
            self.configuracao,
            lancamento=lancamento,
        )
        self.wait_window(dialog)
        if dialog.result:
            self._save_entry_from_dialog(dialog.result, lancamento)

    def _save_entry_from_dialog(
        self, payload: dict, original: LancamentoPlantao | None
    ) -> None:
        try:
            lancamento = self.schedule_service.build_entry(
                self.configuracao,
                **payload,
                origem=original.origem if original else "manual",
                semana_referencia=original.semana_referencia if original else "",
                lancamento_id=original.id if original else "",
            )
            self.schedule_service.replace_entry(self.projeto, lancamento)
        except (ValidacaoErro, ValueError) as exc:
            messagebox.showerror("Dados inválidos", str(exc), parent=self)
            return

        self._refresh_all(reset_filters=False)
        self.entries_tree.selection_set(lancamento.id)
        self.entries_tree.focus(lancamento.id)
        self._mark_dirty("Lançamento salvo com sucesso.")

    def _delete_selected_entry(self) -> None:
        lancamento = self._selected_entry()
        if not lancamento:
            messagebox.showwarning(
                "Seleção obrigatória",
                "Selecione um lançamento para excluir.",
                parent=self,
            )
            return
        if not messagebox.askyesno(
            "Excluir lançamento",
            "Deseja realmente excluir o lançamento selecionado?",
            parent=self,
        ):
            return
        self.schedule_service.remove_entry(self.projeto, lancamento.id)
        self._refresh_all(reset_filters=False)
        self._mark_dirty("Lançamento excluído.")

    def _duplicate_selected_entry(self) -> None:
        lancamento = self._selected_entry()
        if not lancamento:
            messagebox.showwarning(
                "Seleção obrigatória",
                "Selecione um lançamento para duplicar.",
                parent=self,
            )
            return
        try:
            novo = self.schedule_service.duplicate_entry(self.projeto, lancamento.id)
        except ValidacaoErro as exc:
            messagebox.showerror("Não foi possível duplicar", str(exc), parent=self)
            return

        self._refresh_all(reset_filters=False)
        self.entries_tree.selection_set(novo.id)
        self.entries_tree.focus(novo.id)
        self._mark_dirty("Lançamento duplicado.")

    def _add_responsavel(self) -> None:
        dialog = ResponsavelDialog(self, "Adicionar responsável")
        self.wait_window(dialog)
        if not dialog.result:
            return
        try:
            responsavel = self.project_service.add_responsavel(self.projeto, dialog.result)
        except ValidacaoErro as exc:
            messagebox.showerror("Cadastro inválido", str(exc), parent=self)
            return

        self._refresh_all(reset_filters=False)
        self.responsaveis_tree.selection_set(responsavel.id)
        self.responsaveis_tree.focus(responsavel.id)
        self.week_responsavel_var.set(responsavel.nome)
        self._mark_dirty("Responsável adicionado com sucesso.")

    def _edit_responsavel(self) -> None:
        selected = self.responsaveis_tree.selection()
        if not selected:
            messagebox.showwarning(
                "Seleção obrigatória",
                "Selecione um responsável para editar.",
                parent=self,
            )
            return
        responsavel_id = selected[0]
        responsavel = next(
            (item for item in self.projeto.responsaveis if item.id == responsavel_id), None
        )
        if not responsavel:
            return
        dialog = ResponsavelDialog(self, "Editar responsável", valor_inicial=responsavel.nome)
        self.wait_window(dialog)
        if not dialog.result:
            return
        try:
            self.project_service.update_responsavel(self.projeto, responsavel_id, dialog.result)
        except ValidacaoErro as exc:
            messagebox.showerror("Não foi possível editar", str(exc), parent=self)
            return
        self._refresh_all(reset_filters=False)
        self.responsaveis_tree.selection_set(responsavel_id)
        self._mark_dirty("Responsável atualizado.")

    def _remove_responsavel(self) -> None:
        selected = self.responsaveis_tree.selection()
        if not selected:
            messagebox.showwarning(
                "Seleção obrigatória",
                "Selecione um responsável para remover.",
                parent=self,
            )
            return
        responsavel_id = selected[0]
        if not messagebox.askyesno(
            "Remover responsável",
            "Deseja realmente remover o responsável selecionado?",
            parent=self,
        ):
            return
        try:
            self.project_service.remove_responsavel(self.projeto, responsavel_id)
        except ValidacaoErro as exc:
            messagebox.showerror("Não foi possível remover", str(exc), parent=self)
            return
        self._refresh_all(reset_filters=False)
        self._mark_dirty("Responsável removido.")

    def _save_project(self) -> None:
        caminho = self.projeto.caminho_arquivo
        if not caminho:
            nome_padrao = f"plantao_{self.projeto.ano}_{self.projeto.mes:02d}.json"
            caminho = filedialog.asksaveasfilename(
                title="Salvar projeto",
                defaultextension=".json",
                initialdir=str(data_dir()),
                initialfile=nome_padrao,
                filetypes=[("Projeto JSON", "*.json")],
                parent=self,
            )
            if not caminho:
                return
        try:
            self.project_service.save_project(self.projeto, caminho)
        except OSError as exc:
            messagebox.showerror(
                "Falha ao salvar",
                f"Não foi possível salvar o projeto.\n\nDetalhes: {exc}",
                parent=self,
            )
            return
        self._refresh_all(reset_filters=False)
        self._clear_dirty("Projeto salvo com sucesso.")

    def _open_project(self) -> None:
        if not self._confirm_discard_if_needed():
            return
        caminho = filedialog.askopenfilename(
            title="Abrir projeto",
            initialdir=str(data_dir()),
            filetypes=[("Projeto JSON", "*.json")],
            parent=self,
        )
        if not caminho:
            return
        self._load_project_from_path(Path(caminho))

    def _open_sample_project(self) -> None:
        if not self._confirm_discard_if_needed():
            return
        caminho = sample_project_path()
        if not caminho.exists():
            messagebox.showerror(
                "Exemplo indisponível",
                "O arquivo de exemplo não foi encontrado.",
                parent=self,
            )
            return
        self._load_project_from_path(caminho)

    def _load_project_from_path(self, caminho: Path) -> None:
        try:
            self.projeto = self.project_service.load_project(caminho)
        except (OSError, ValueError) as exc:
            messagebox.showerror(
                "Falha ao abrir",
                f"Não foi possível abrir o projeto.\n\nDetalhes: {exc}",
                parent=self,
            )
            return
        self.dirty = False
        self.sort_column = "data"
        self.sort_desc = False
        self._refresh_all(reset_filters=True)
        self._clear_dirty("Projeto aberto com sucesso.")

    def _export_excel(self) -> None:
        if not self.projeto.lancamentos:
            messagebox.showwarning(
                "Sem dados para exportação",
                "Cadastre ou gere lançamentos antes de exportar para Excel.",
                parent=self,
            )
            return

        nome_padrao = self.configuracao.nome_padrao_arquivo_exportado.format(
            ano=self.projeto.ano,
            mes=f"{self.projeto.mes:02d}",
        )
        caminho = filedialog.asksaveasfilename(
            title="Exportar Excel",
            defaultextension=".xlsx",
            initialdir=str(exports_dir()),
            initialfile=nome_padrao,
            filetypes=[("Excel", "*.xlsx")],
            parent=self,
        )
        if not caminho:
            return

        try:
            arquivo = self.excel_service.export(self.projeto, self.configuracao, caminho)
        except ExportacaoErro as exc:
            messagebox.showerror("Exportação interrompida", str(exc), parent=self)
            return
        messagebox.showinfo(
            "Exportação concluída",
            f"Arquivo Excel gerado com sucesso.\n\n{arquivo}",
            parent=self,
        )
        self._set_status(f"Excel exportado para {arquivo}.")

    def _clear_filters(self) -> None:
        self.filter_responsavel_var.set("Todos")
        self.filter_mes_var.set(nome_mes(self.projeto.mes))
        self.filter_tipo_var.set("Todos")
        self.filter_cliente_var.set("")
        self._apply_filters()

    def _toggle_sort(self, column: str) -> None:
        if self.sort_column == column:
            self.sort_desc = not self.sort_desc
        else:
            self.sort_column = column
            self.sort_desc = False
        self._apply_filters()

    def _on_calendar_selected(self, _event=None) -> None:
        self._select_week_from_calendar()

    def _select_week_from_calendar(self) -> None:
        data_calendario = self._calendar_selected_date_or_default()
        semana = encontrar_semana_por_data(self.projeto.ano, self.projeto.mes, data_calendario)
        if semana:
            self._select_week_by_id(semana["id"])

    def _calendar_selected_date_or_default(self) -> date:
        fallback = date(self.projeto.ano, self.projeto.mes, 1)
        try:
            data_calendario = self.calendar.selection_get()
        except Exception:
            return fallback
        return data_calendario if isinstance(data_calendario, date) else fallback

    def _select_week_by_id(self, week_id: str) -> None:
        if week_id not in self.weeks_tree.get_children():
            return
        self.weeks_tree.selection_set(week_id)
        self.weeks_tree.focus(week_id)
        self.weeks_tree.see(week_id)
        self._sync_week_assignment_combo()

    def _confirm_discard_if_needed(self) -> bool:
        if not self.dirty:
            return True
        return messagebox.askyesno(
            "Existem alterações não salvas",
            "Há alterações não salvas no projeto atual. Deseja continuar mesmo assim?",
            parent=self,
        )

    def _on_close(self) -> None:
        if self._confirm_discard_if_needed():
            self.destroy()
