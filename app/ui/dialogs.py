from __future__ import annotations

from datetime import datetime
from tkinter import messagebox

import customtkinter as ctk
from tkcalendar import DateEntry

from app.models import ConfiguracaoSistema, LancamentoPlantao, Responsavel
from app.services.schedule_service import ScheduleService
from app.utils.formatters import formatar_horas, formatar_moeda_br
from app.utils.validators import ValidacaoErro


class BaseDialog(ctk.CTkToplevel):
    def __init__(self, parent, titulo: str, largura: int, altura: int) -> None:
        super().__init__(parent)
        self.result = None
        self.title(titulo)
        self.geometry(f"{largura}x{altura}")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.after(50, self._center)

    def _center(self) -> None:
        self.update_idletasks()
        parent = self.master
        x = parent.winfo_x() + (parent.winfo_width() // 2) - (self.winfo_width() // 2)
        y = parent.winfo_y() + (parent.winfo_height() // 2) - (self.winfo_height() // 2)
        self.geometry(f"+{max(x, 80)}+{max(y, 60)}")

    def _cancel(self) -> None:
        self.result = None
        self.destroy()


class ResponsavelDialog(BaseDialog):
    def __init__(self, parent, titulo: str, valor_inicial: str = "") -> None:
        super().__init__(parent, titulo, 420, 180)

        self.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            self,
            text="Nome do responsável",
            font=("Segoe UI Semibold", 15),
            anchor="w",
        ).grid(row=0, column=0, padx=20, pady=(20, 8), sticky="ew")

        self.nome_var = ctk.StringVar(value=valor_inicial)
        self.entry_nome = ctk.CTkEntry(self, textvariable=self.nome_var, height=40)
        self.entry_nome.grid(row=1, column=0, padx=20, sticky="ew")
        self.entry_nome.focus()

        footer = ctk.CTkFrame(self, fg_color="transparent")
        footer.grid(row=2, column=0, padx=20, pady=20, sticky="ew")
        for indice in range(2):
            footer.grid_columnconfigure(indice, weight=1)

        ctk.CTkButton(footer, text="Cancelar", fg_color="#94A3B8", command=self._cancel).grid(
            row=0, column=0, padx=(0, 6), sticky="ew"
        )
        ctk.CTkButton(footer, text="Salvar", command=self._submit).grid(
            row=0, column=1, padx=(6, 0), sticky="ew"
        )

    def _submit(self) -> None:
        self.result = self.nome_var.get().strip()
        self.destroy()


class LancamentoDialog(BaseDialog):
    def __init__(
        self,
        parent,
        titulo: str,
        responsaveis: list[Responsavel],
        config: ConfiguracaoSistema,
        lancamento: LancamentoPlantao | None = None,
        fixed_type: str | None = None,
    ) -> None:
        super().__init__(parent, titulo, 860, 640)
        self.schedule_service = ScheduleService()
        self.configuracao = config
        self.responsaveis = responsaveis
        self.lancamento = lancamento
        self.fixed_type = fixed_type
        self.responsavel_por_nome = {item.nome: item.id for item in responsaveis}

        self.grid_columnconfigure(0, weight=1)

        container = ctk.CTkScrollableFrame(self, fg_color="transparent")
        container.grid(row=0, column=0, sticky="nsew", padx=18, pady=18)
        for indice in range(2):
            container.grid_columnconfigure(indice, weight=1)

        titulo_label = ctk.CTkLabel(
            container,
            text=titulo,
            font=("Segoe UI Semibold", 20),
            anchor="w",
        )
        titulo_label.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 6))

        subtitulo = (
            "Preencha o atendimento urgente ou ajuste manualmente qualquer lançamento."
            if fixed_type == "Suporte"
            else "Revise os dados do lançamento e confirme para salvar."
        )
        ctk.CTkLabel(
            container,
            text=subtitulo,
            text_color="#5B6B82",
            anchor="w",
        ).grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 16))

        self.responsavel_var = ctk.StringVar()
        self.tipo_var = ctk.StringVar(value=fixed_type or "Sobre aviso")
        self.inicio_var = ctk.StringVar()
        self.fim_var = ctk.StringVar()
        self.horas_var = ctk.StringVar()
        self.numero_var = ctk.StringVar()
        self.solicitante_var = ctk.StringVar()
        self.cliente_var = ctk.StringVar()
        self.nivel_var = ctk.StringVar()

        self._create_labeled_combo(
            container, 2, 0, "Responsável", self.responsavel_var, list(self.responsavel_por_nome)
        )
        self._create_labeled_combo(
            container,
            2,
            1,
            "Tipo",
            self.tipo_var,
            ["Sobre aviso", "Suporte"],
            readonly=fixed_type is not None,
        )

        ctk.CTkLabel(container, text="Data", anchor="w").grid(
            row=4, column=0, sticky="ew", padx=(0, 10)
        )
        self.data_entry = DateEntry(
            container,
            date_pattern="dd/MM/yyyy",
            locale=config.idioma_calendario,
            background="#1F5AA6",
            foreground="white",
            borderwidth=1,
            font=("Segoe UI", 10),
        )
        self.data_entry.grid(row=5, column=0, sticky="ew", padx=(0, 10), pady=(4, 14), ipady=5)

        self._create_labeled_entry(container, 4, 1, "Hora início", self.inicio_var)
        self._create_labeled_entry(container, 6, 0, "Hora fim", self.fim_var)
        self._create_labeled_entry(container, 6, 1, "Quantidade de horas", self.horas_var)
        self._create_labeled_entry(container, 8, 0, "Nº Chamado", self.numero_var)
        self._create_labeled_entry(container, 8, 1, "Solicitante", self.solicitante_var)
        self._create_labeled_entry(container, 10, 0, "Cliente", self.cliente_var)
        self._create_labeled_entry(container, 10, 1, "Nível", self.nivel_var)

        ctk.CTkLabel(container, text="Observação", anchor="w").grid(
            row=12, column=0, columnspan=2, sticky="ew"
        )
        self.observacao_box = ctk.CTkTextbox(container, height=110)
        self.observacao_box.grid(
            row=13, column=0, columnspan=2, sticky="ew", pady=(4, 16)
        )

        preview = ctk.CTkFrame(container, fg_color="#EEF4FC", corner_radius=14)
        preview.grid(row=14, column=0, columnspan=2, sticky="ew", pady=(0, 14))
        for indice in range(3):
            preview.grid_columnconfigure(indice, weight=1)
        self.preview_horas = ctk.CTkLabel(preview, text="Horas calculadas: --", anchor="w")
        self.preview_horas.grid(row=0, column=0, padx=14, pady=12, sticky="ew")
        self.preview_fim = ctk.CTkLabel(preview, text="Fim previsto: --", anchor="w")
        self.preview_fim.grid(row=0, column=1, padx=14, pady=12, sticky="ew")
        self.preview_valor = ctk.CTkLabel(preview, text="Valor previsto: --", anchor="w")
        self.preview_valor.grid(row=0, column=2, padx=14, pady=12, sticky="ew")

        actions = ctk.CTkFrame(container, fg_color="transparent")
        actions.grid(row=15, column=0, columnspan=2, sticky="ew", pady=(4, 8))
        for indice in range(3):
            actions.grid_columnconfigure(indice, weight=1)

        ctk.CTkButton(
            actions,
            text="Atualizar cálculos",
            fg_color="#2F80ED",
            command=self._update_preview,
        ).grid(row=0, column=0, padx=(0, 6), sticky="ew")
        ctk.CTkButton(
            actions,
            text="Cancelar",
            fg_color="#94A3B8",
            command=self._cancel,
        ).grid(row=0, column=1, padx=6, sticky="ew")
        ctk.CTkButton(actions, text="Salvar lançamento", command=self._submit).grid(
            row=0, column=2, padx=(6, 0), sticky="ew"
        )

        self._prefill()
        for variable in (
            self.responsavel_var,
            self.tipo_var,
            self.inicio_var,
            self.fim_var,
            self.horas_var,
        ):
            variable.trace_add("write", lambda *_: self._update_preview())

    def _create_labeled_entry(self, parent, row: int, column: int, label: str, variable) -> None:
        padx = (0, 10) if column == 0 else (10, 0)
        ctk.CTkLabel(parent, text=label, anchor="w").grid(
            row=row, column=column, sticky="ew", padx=padx
        )
        ctk.CTkEntry(parent, textvariable=variable, height=38).grid(
            row=row + 1, column=column, sticky="ew", padx=padx, pady=(4, 14)
        )

    def _create_labeled_combo(
        self,
        parent,
        row: int,
        column: int,
        label: str,
        variable,
        values: list[str],
        readonly: bool = False,
    ) -> None:
        padx = (0, 10) if column == 0 else (10, 0)
        ctk.CTkLabel(parent, text=label, anchor="w").grid(
            row=row, column=column, sticky="ew", padx=padx
        )
        combo = ctk.CTkComboBox(
            parent,
            variable=variable,
            values=values or [""],
            state="readonly",
            height=38,
        )
        combo.grid(row=row + 1, column=column, sticky="ew", padx=padx, pady=(4, 14))
        if values:
            combo.set(values[0] if not variable.get() else variable.get())

    def _prefill(self) -> None:
        hoje = datetime.now().date()
        self.data_entry.set_date(hoje)
        if self.responsaveis:
            self.responsavel_var.set(self.responsaveis[0].nome)

        if self.lancamento:
            self.responsavel_var.set(self.lancamento.responsavel)
            self.tipo_var.set(self.fixed_type or self.lancamento.tipo)
            self.data_entry.set_date(datetime.strptime(self.lancamento.data, "%Y-%m-%d").date())
            self.inicio_var.set(self.lancamento.inicio)
            self.fim_var.set(self.lancamento.fim)
            self.horas_var.set(str(self.lancamento.total_horas).replace(".", ","))
            self.numero_var.set(self.lancamento.numero_chamado)
            self.solicitante_var.set(self.lancamento.solicitante)
            self.cliente_var.set(self.lancamento.cliente)
            self.nivel_var.set(self.lancamento.nivel)
            self.observacao_box.insert("1.0", self.lancamento.observacao)
        elif self.fixed_type:
            self.tipo_var.set(self.fixed_type)
            self.inicio_var.set("20:00")
        else:
            self.inicio_var.set("20:00")
            self.fim_var.set("06:00")

        self._update_preview()

    def _build_payload(self) -> dict:
        responsavel_nome = self.responsavel_var.get().strip()
        responsavel_id = self.responsavel_por_nome.get(responsavel_nome, "")
        horas_texto = self.horas_var.get().strip().replace(",", ".")
        horas = float(horas_texto) if horas_texto else None

        return {
            "responsavel_id": responsavel_id,
            "responsavel_nome": responsavel_nome,
            "data_referencia": self.data_entry.get_date().isoformat(),
            "inicio": self.inicio_var.get().strip(),
            "fim": self.fim_var.get().strip(),
            "horas": horas,
            "tipo": self.fixed_type or self.tipo_var.get().strip(),
            "numero_chamado": self.numero_var.get().strip(),
            "solicitante": self.solicitante_var.get().strip(),
            "cliente": self.cliente_var.get().strip(),
            "nivel": self.nivel_var.get().strip(),
            "observacao": self.observacao_box.get("1.0", "end").strip(),
        }

    def _update_preview(self) -> None:
        try:
            payload = self._build_payload()
            if not payload["responsavel_id"] or not payload["inicio"]:
                raise ValidacaoErro("")
            preview = self.schedule_service.build_entry(self.configuracao, **payload)
            self.preview_horas.configure(
                text=f"Horas calculadas: {formatar_horas(preview.total_horas)}"
            )
            self.preview_fim.configure(text=f"Fim previsto: {preview.fim}")
            self.preview_valor.configure(
                text=f"Valor previsto: {formatar_moeda_br(preview.valor)}"
            )
        except Exception:
            self.preview_horas.configure(text="Horas calculadas: --")
            self.preview_fim.configure(text="Fim previsto: --")
            self.preview_valor.configure(text="Valor previsto: --")

    def _submit(self) -> None:
        try:
            self.result = self._build_payload()
        except ValueError:
            messagebox.showerror(
                "Horas inválidas",
                "A quantidade de horas deve ser um número válido.",
                parent=self,
            )
            return

        if not self.result.get("responsavel_id"):
            messagebox.showerror(
                "Responsável obrigatório",
                "Selecione um responsável para o lançamento.",
                parent=self,
            )
            return

        self.destroy()
