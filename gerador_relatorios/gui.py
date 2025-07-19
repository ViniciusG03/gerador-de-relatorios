from __future__ import annotations

import os
import tkinter as tk
from typing import Any
from tkinter import ttk, filedialog, messagebox
from collections import defaultdict
import pandas as pd

from . import data_loader
from .reports import generate_pne_report, generate_tipico_report
from .utils import resource_path


class MedicalReportGenerator:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Gerador de Relatórios FUSEX")
        self.root.geometry("900x700")

        try:
            self.root.iconbitmap(resource_path("icone.ico"))
        except Exception:
            pass

        self.excel_file: str | None = None
        self.report_type = tk.StringVar(value="PNE")
        self.data: pd.DataFrame | None = None

        self.setup_ui()

    def setup_ui(self) -> None:
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        title_label = ttk.Label(
            main_frame,
            text="🏥 Gerador de Relatórios FUSEX",
            font=("Arial", 18, "bold"),
        )
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 25))

        type_frame = ttk.LabelFrame(main_frame, text="📋 Tipo de Relatório", padding="15")
        type_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        ttk.Radiobutton(type_frame, text="🧩 PNE (Portador de Necessidades Especiais)", variable=self.report_type, value="PNE").pack(anchor=tk.W, pady=5)
        ttk.Radiobutton(type_frame, text="👤 Típico", variable=self.report_type, value="TIPICO").pack(anchor=tk.W, pady=5)

        file_frame = ttk.LabelFrame(main_frame, text="📁 Arquivo Excel", padding="15")
        file_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))

        file_inner_frame = ttk.Frame(file_frame)
        file_inner_frame.pack(fill=tk.X)

        self.file_label = ttk.Label(file_inner_frame, text="Nenhum arquivo selecionado", foreground="gray")
        self.file_label.pack(side=tk.LEFT, padx=(0, 15))

        btn_frame = ttk.Frame(file_inner_frame)
        btn_frame.pack(side=tk.RIGHT)
        ttk.Button(btn_frame, text="📁 Selecionar Arquivo", command=self.select_file).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(btn_frame, text="📄 Usar Exemplo", command=self.load_example).pack(side=tk.LEFT)

        preview_frame = ttk.LabelFrame(main_frame, text="👁️ Preview dos Dados", padding="15")
        preview_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))

        columns = ("Nome", "Data Nascimento", "Responsável", "Especialidade", "Mês Referência")
        self.tree = ttk.Treeview(preview_frame, columns=columns, show="headings", height=12)
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150)

        scrollbar_v = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar_h = ttk.Scrollbar(preview_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar_v.set, xscrollcommand=scrollbar_h.set)

        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar_v.grid(row=0, column=1, sticky=(tk.N, tk.S))
        scrollbar_h.grid(row=1, column=0, sticky=(tk.W, tk.E))

        preview_frame.grid_rowconfigure(0, weight=1)
        preview_frame.grid_columnconfigure(0, weight=1)

        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=3, pady=(15, 0))
        ttk.Button(button_frame, text="🚀 Gerar Relatórios", command=self.generate_reports).pack(side=tk.LEFT, padx=(0, 15))
        ttk.Button(button_frame, text="🗑️ Limpar", command=self.clear_data).pack(side=tk.LEFT, padx=(0, 15))
        ttk.Button(button_frame, text="❓ Ajuda", command=self.show_help).pack(side=tk.LEFT)

        self.status_var = tk.StringVar(value="✅ Pronto para uso")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W, padding="5")
        status_bar.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(15, 0))

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(2, weight=1)
        main_frame.rowconfigure(3, weight=1)

    def load_example(self) -> None:
        try:
            example_file = resource_path("fusex_tipico.xlsx")
            if os.path.exists(example_file):
                self.excel_file = example_file
                self.data = data_loader.load_excel(example_file)
                self.file_label.config(text="📄 fusex_tipico.xlsx (exemplo)", foreground="blue")
                self.load_preview()
                self.status_var.set(f"📄 Exemplo carregado: {len(self.data)} registros")
            else:
                messagebox.showerror("Erro", "Arquivo de exemplo não encontrado!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar exemplo: {str(e)}")

    def show_help(self) -> None:
        help_window = tk.Toplevel(self.root)
        help_window.title("Ajuda - Gerador de Relatórios FUSEX")
        help_window.geometry("600x400")
        help_window.transient(self.root)
        help_window.grab_set()

        help_text = """
        📋 COMO USAR O GERADOR DE RELATÓRIOS FUSEX

        1️⃣ PREPARAR PLANILHA EXCEL:
           • Colunas obrigatórias:
             - NOME
             - DATA DE NASCIMENTO
             - RESPONSÁVEL
             - ESPECIALIDADE
             - MÊS DE REFERÊNCIA

        2️⃣ SELECIONAR TIPO DE RELATÓRIO:
           • PNE: Para portadores de necessidades especiais
           • Típico: Para desenvolvimento típico

        3️⃣ CARREGAR DADOS:
           • Use "Selecionar Arquivo" para sua planilha
           • Ou "Usar Exemplo" para testar

        4️⃣ GERAR RELATÓRIOS:
           • Clique em "Gerar Relatórios"
           • Escolha pasta de destino
           • Aguarde o processamento

        ⚙️ ESPECIALIDADES SUPORTADAS:
           • Terapia ABA
           • Psicoterapia
           • Terapia Ocupacional
           • Fonoaudiologia
           • Psicomotricidade
           • Psicopedagogia
           • Nutrição (apenas PNE)
           • Fisioterapia (apenas PNE)

        📧 SUPORTE: Qualquer dúvida, entre em contato!
        """

        text_widget = tk.Text(help_window, wrap=tk.WORD, padx=20, pady=20)
        text_widget.insert(tk.END, help_text)
        text_widget.config(state=tk.DISABLED)
        text_widget.pack(fill=tk.BOTH, expand=True)
        ttk.Button(help_window, text="✅ Entendi", command=help_window.destroy).pack(pady=10)

    def select_file(self) -> None:
        file_path = filedialog.askopenfilename(title="Selecionar arquivo Excel", filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            try:
                self.excel_file = file_path
                self.data = data_loader.load_excel(file_path)
                self.file_label.config(text=f"📁 {os.path.basename(file_path)}", foreground="green")
                self.load_preview()
                self.status_var.set(f"📊 Arquivo carregado: {len(self.data)} registros")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao carregar arquivo: {str(e)}")

    def load_preview(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)

        for _, row in self.data.iterrows():
            data_nasc = row["DATA DE NASCIMENTO"]
            try:
                if pd.notna(data_nasc):
                    if hasattr(data_nasc, "strftime"):
                        data_nasc = data_nasc.strftime("%d/%m/%Y")
                    else:
                        data_nasc = str(data_nasc)
                        if "/" in data_nasc:
                            data_nasc = data_nasc.split(" ")[0]
                else:
                    data_nasc = "Data não informada"
            except Exception:
                data_nasc = str(data_nasc) if pd.notna(data_nasc) else "Data não informada"

            self.tree.insert(
                "",
                tk.END,
                values=(
                    str(row["NOME"]) if pd.notna(row["NOME"]) else "Nome não informado",
                    data_nasc,
                    str(row["RESPONSÁVEL"]) if pd.notna(row["RESPONSÁVEL"]) else "Não informado",
                    str(row["ESPECIALIDADE"]) if pd.notna(row["ESPECIALIDADE"]) else "Não informado",
                    str(row["MÊS DE REFERÊNCIA"]) if pd.notna(row["MÊS DE REFERÊNCIA"]) else "Não informado",
                ),
            )

    def clear_data(self) -> None:
        self.excel_file = None
        self.data = None
        self.file_label.config(text="Nenhum arquivo selecionado", foreground="gray")
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.status_var.set("✅ Pronto para uso")

    def generate_reports(self) -> None:
        if self.data is None:
            messagebox.showerror(
                "Erro",
                "Nenhum arquivo carregado! Use 'Selecionar Arquivo' ou 'Usar Exemplo'.",
            )
            return

        patient_data: dict[str, dict[str, Any]] = defaultdict(lambda: {"info": {}, "especialidades": []})
        for _, row in self.data.iterrows():
            nome = row["NOME"]
            if not patient_data[nome]["info"]:
                data_nasc = row["DATA DE NASCIMENTO"]
                try:
                    if pd.notna(data_nasc):
                        if hasattr(data_nasc, "strftime"):
                            data_nasc = data_nasc.strftime("%d/%m/%Y")
                        else:
                            data_nasc = str(data_nasc)
                            if "/" in data_nasc:
                                data_nasc = data_nasc.split(" ")[0]
                    else:
                        data_nasc = "Data não informada"
                except Exception:
                    data_nasc = str(data_nasc) if pd.notna(data_nasc) else "Data não informada"
                patient_data[nome]["info"] = {
                    "nome": str(row["NOME"]),
                    "data_nascimento": data_nasc,
                    "responsavel": str(row["RESPONSÁVEL"]) if pd.notna(row["RESPONSÁVEL"]) else "Não informado",
                    "mes_referencia": str(row["MÊS DE REFERÊNCIA"]) if pd.notna(row["MÊS DE REFERÊNCIA"]) else "Não informado",
                }
            if pd.notna(row["ESPECIALIDADE"]):
                patient_data[nome]["especialidades"].append(str(row["ESPECIALIDADE"]))

        output_dir = filedialog.askdirectory(title="Selecionar pasta para salvar relatórios")
        if not output_dir:
            return

        try:
            self.status_var.set("🔄 Gerando relatórios...")
            self.root.update()
            report_count = 0
            for nome, pdata in patient_data.items():
                if self.report_type.get() == "PNE":
                    generate_pne_report(pdata, output_dir)
                else:
                    generate_tipico_report(pdata, output_dir)
                report_count += 1
                self.status_var.set(f"🔄 Gerando... {report_count}/{len(patient_data)}")
                self.root.update()
            self.status_var.set(f"✅ {report_count} relatórios gerados com sucesso!")
            messagebox.showinfo(
                "Sucesso! 🎉",
                f"✅ {report_count} relatórios gerados com sucesso!\n\n📁 Pasta: {output_dir}",
            )
        except Exception as e:
            self.status_var.set("❌ Erro ao gerar relatórios")
            messagebox.showerror("Erro", f"Erro ao gerar relatórios: {str(e)}")
