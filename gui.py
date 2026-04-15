from __future__ import annotations

from pathlib import Path
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd

if __package__ in (None, ""):
    from analysis import AnalysisError, AnalysisResult, analyze_files, export_analysis
else:
    from .analysis import AnalysisError, AnalysisResult, analyze_files, export_analysis


def format_decimal(value: object) -> str:
    if isinstance(value, (int, float)):
        return f"{value:,.2f}".replace(",", " ").replace(".", ",")
    return str(value)


def format_integer(value: object) -> str:
    if isinstance(value, (int, float)):
        return f"{round(value):,}".replace(",", " ")
    return str(value)


class AnalysisApp:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("Butikkpotensial")
        self.root.geometry("1380x880")

        self.brutto_path_var = tk.StringVar()
        self.minimum_gross_var = tk.StringVar(value="18")
        self.top_n_var = tk.StringVar(value="10")

        self.compare_paths: list[str] = []
        self.analysis_result: AnalysisResult | None = None

        self.build_layout()
        self.root.bind("<Escape>", self.reset)

    def build_layout(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)

        controls = ttk.Frame(self.root, padding=16)
        controls.grid(row=0, column=0, sticky="ew")
        controls.columnconfigure(1, weight=1)

        ttk.Label(controls, text="Bruttofil").grid(row=0, column=0, sticky="w")
        ttk.Entry(controls, textvariable=self.brutto_path_var).grid(row=0, column=1, sticky="ew", padx=(8, 8))
        ttk.Button(controls, text="Velg fil", command=self.select_brutto_file).grid(row=0, column=2, sticky="ew")

        ttk.Label(controls, text="Sammenligningsfiler").grid(row=1, column=0, sticky="nw", pady=(8, 0))
        compare_frame = ttk.Frame(controls)
        compare_frame.grid(row=2, column=1, columnspan=2, sticky="ew", pady=(8, 0))
        compare_frame.columnconfigure(0, weight=1)

        self.compare_listbox = tk.Listbox(compare_frame, height=6)
        self.compare_listbox.grid(row=0, column=0, sticky="ew")

        compare_buttons = ttk.Frame(compare_frame)
        compare_buttons.grid(row=0, column=1, sticky="ns", padx=(8, 0))
        ttk.Button(compare_buttons, text="Legg til", command=self.select_compare_files).grid(row=0, column=0, sticky="ew")
        ttk.Button(compare_buttons, text="Fjern valgt", command=self.remove_selected_compare_file).grid(
            row=1, column=0, sticky="ew", pady=(8, 0)
        )
        ttk.Button(compare_buttons, text="Tøm", command=self.clear_compare_files).grid(row=2, column=0, sticky="ew", pady=(8, 0))

        options = ttk.Frame(controls)
        options.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(16, 0))
        options.columnconfigure(3, weight=1)

        ttk.Label(options, text="Minste brutto %").grid(row=0, column=0, sticky="w")
        ttk.Entry(options, width=8, textvariable=self.minimum_gross_var).grid(row=0, column=1, sticky="w", padx=(8, 16))

        ttk.Label(options, text="Toppliste per gruppe").grid(row=0, column=2, sticky="w")
        ttk.Entry(options, width=8, textvariable=self.top_n_var).grid(row=0, column=3, sticky="w", padx=(8, 16))

        actions = ttk.Frame(controls)
        actions.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(16, 0))
        ttk.Button(actions, text="Kjør analyse", command=self.run_analysis).grid(row=0, column=0, sticky="w")
        self.export_button = ttk.Button(actions, text="Eksporter til Excel", command=self.export_to_excel, state="disabled")
        self.export_button.grid(row=0, column=1, sticky="w", padx=(8, 0))

        self.notebook = ttk.Notebook(self.root)
        self.notebook.grid(row=1, column=0, sticky="nsew", padx=16, pady=(0, 16))

    def select_brutto_file(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Excel-filer", "*.xlsx *.xls")])
        if path:
            self.brutto_path_var.set(path)

    def select_compare_files(self) -> None:
        paths = filedialog.askopenfilenames(filetypes=[("Excel-filer", "*.xlsx *.xls")])
        for path in paths:
            if path not in self.compare_paths:
                self.compare_paths.append(path)
        self.refresh_compare_listbox()

    def refresh_compare_listbox(self) -> None:
        self.compare_listbox.delete(0, tk.END)
        for path in self.compare_paths:
            self.compare_listbox.insert(tk.END, path)

    def remove_selected_compare_file(self) -> None:
        selection = list(self.compare_listbox.curselection())
        for index in reversed(selection):
            self.compare_paths.pop(index)
        self.refresh_compare_listbox()

    def clear_compare_files(self) -> None:
        self.compare_paths.clear()
        self.refresh_compare_listbox()

    def run_analysis(self) -> None:
        if not self.brutto_path_var.get().strip():
            messagebox.showerror("Mangler bruttofil", "Velg bruttofila før analysen kjøres.")
            return

        if not self.compare_paths:
            messagebox.showerror("Mangler sammenligningsfiler", "Velg minst én sammenligningsfil.")
            return

        try:
            minimum_gross = float(self.minimum_gross_var.get().replace(",", "."))
            top_n = int(float(self.top_n_var.get().replace(",", ".")))
        except ValueError:
            messagebox.showerror("Ugyldige tall", "Bruttofilter og toppliste må være gyldige tall.")
            return

        if top_n <= 0:
            messagebox.showerror("Ugyldig toppliste", "Topplisten må være minst 1.")
            return

        try:
            self.analysis_result = analyze_files(
                self.compare_paths,
                self.brutto_path_var.get(),
                minimum_gross_percent=minimum_gross,
                top_n=top_n,
                normalization_mode="category",
            )
        except AnalysisError as exc:
            messagebox.showerror("Kunne ikke kjøre analyse", str(exc))
            return
        except Exception as exc:
            messagebox.showerror("Uventet feil", str(exc))
            return

        self.render_results()
        self.export_button.configure(state="normal")

    def render_results(self) -> None:
        for tab_id in self.notebook.tabs():
            self.notebook.forget(tab_id)

        if not self.analysis_result:
            return

        for category_result in self.analysis_result.category_results:
            frame = ttk.Frame(self.notebook, padding=16)
            self.notebook.add(frame, text=category_result.category_name)
            self.render_category_tab(frame, category_result)

        if self.notebook.tabs():
            self.notebook.select(0)

    def render_summary_tab(self) -> None:
        if not self.analysis_result:
            return

        summary_lines = [
            f"Butikk A: {self.analysis_result.store_a_name}",
            f"Butikk B: {self.analysis_result.store_b_name}",
            f"Bruttofilter: {self.analysis_result.minimum_gross_percent:.2f} %",
            f"Toppliste per varegruppe: {self.analysis_result.top_n}",
            f"Antall rader i bruttofila: {self.analysis_result.brutto_base_rows}",
            (
                "Normalisering: "
                + ("Bruker totalfil" if self.analysis_result.normalization_mode == "total-file" else "Per varegruppe")
            ),
        ]

        ttk.Label(
            self.summary_frame,
            text="\n".join(summary_lines),
            justify="left",
        ).grid(row=0, column=0, sticky="w")

        table_frame = ttk.Frame(self.summary_frame)
        table_frame.grid(row=1, column=0, sticky="nsew", pady=(16, 0))
        self.summary_frame.columnconfigure(0, weight=1)
        self.summary_frame.rowconfigure(1, weight=1)

        display_df = self.analysis_result.summary.copy()
        for column in display_df.columns:
            if pd.api.types.is_numeric_dtype(display_df[column]):
                display_df[column] = display_df[column].map(format_decimal)

        self.create_treeview(table_frame, display_df)

    def render_category_tab(self, frame: ttk.Frame, category_result) -> None:
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)

        display_store_a_name = "Butikk A"

        table_frame = ttk.Frame(frame)
        table_frame.grid(row=0, column=0, sticky="nsew", pady=(16, 0))

        display_df = category_result.table.copy()
        display_df = display_df.rename(columns={
            "Tallkode": "Rema ID",
            category_result.store_a_name: "Butikk A",
            "Brutto %": "Brutto %",
            "Potensiell brutto kr": "Potensiell brutto kr",
        })

        # Rename any adjusted B sales column to a stable label.
        for col in list(display_df.columns):
            if col != "Butikk B Justert" and "justert" in col.lower():
                display_df = display_df.rename(columns={col: "Butikk B Justert"})
                break

        # Only keep the columns that make the result easy to read.
        wanted_columns = [
            "Vare",
            "Rema ID",
            "Butikk A",
            "Butikk B Justert",
            "Potensial i Butikk A",
            "Brutto %",
            "Potensiell brutto kr",
        ]

        # If the exact potential column name does not exist, fall back to any column starting with Potensial i.
        if wanted_columns[4] not in display_df.columns:
            potential_columns = [c for c in display_df.columns if c.startswith("Potensial i ")] or []
            if potential_columns:
                display_df = display_df.rename(columns={potential_columns[0]: wanted_columns[4]})

        display_df = display_df[[col for col in wanted_columns if col in display_df.columns]]
        if "Vare" in display_df.columns:
            display_df["Vare"] = display_df["Vare"].map(self.strip_trailing_rema_id)

        round_columns = {
            "Butikk A",
            "Butikk B Justert",
            "Potensial i Butikk A",
        }

        for column in display_df.columns:
            if pd.api.types.is_numeric_dtype(display_df[column]):
                if column in round_columns:
                    display_df[column] = display_df[column].map(format_integer)
                else:
                    display_df[column] = display_df[column].map(format_decimal)

        tree = self.create_treeview(table_frame, display_df)

        ttk.Label(
            frame,
            text="Dobbeltklikk en rad for å kopiere Rema ID til utklippstavlen.",
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(12, 0))

    def strip_trailing_rema_id(self, value: object) -> str:
        if pd.isna(value):
            return ""
        text = str(value).strip()
        # Keep only the product name before the last dash.
        return re.sub(r"\s*-\s*[^-]*$", "", text).strip()

    def reset(self, event: tk.Event | None = None) -> None:
        self.brutto_path_var.set("")
        self.compare_paths.clear()
        self.refresh_compare_listbox()
        self.minimum_gross_var.set("18")
        self.top_n_var.set("10")
        self.analysis_result = None
        for tab_id in self.notebook.tabs():
            self.notebook.forget(tab_id)
        self.export_button.configure(state="disabled")

    def create_treeview(self, parent: ttk.Frame, dataframe: pd.DataFrame) -> ttk.Treeview:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)

        columns = list(dataframe.columns)
        tree = ttk.Treeview(parent, columns=columns, show="headings")
        tree.grid(row=0, column=0, sticky="nsew")

        vertical_scroll = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        vertical_scroll.grid(row=0, column=1, sticky="ns")
        tree.configure(yscrollcommand=vertical_scroll.set)

        for column in columns:
            if column == "Vare":
                width = 260
            elif column == "Rema ID":
                width = 120
            elif column == "Butikk B Justert":
                width = 120
            elif column == "Solsiden justert":
                width = 120
            elif column == "Brutto %":
                width = 100
            elif column == "Potensiell brutto kr":
                width = 140
            elif column.startswith("Potensial i "):
                width = 140
            else:
                width = 100

            tree.heading(column, text=column, command=lambda c=column: self.sort_treeview(tree, c))
            tree.column(column, width=width, anchor="w", stretch=True)

        for row in dataframe.itertuples(index=False):
            tree.insert("", tk.END, values=list(row))

        tree.bind("<Double-1>", lambda event, widget=tree: self.copy_tallkode_on_double_click(event, widget))
        return tree

    def sort_treeview(self, tree: ttk.Treeview, column: str) -> None:
        values = [(tree.set(item, column), item) for item in tree.get_children("")]

        def convert(value: str):
            if value is None:
                return ""
            if isinstance(value, str):
                text = value.strip()
                if not text:
                    return ""
                normalized = text.replace(" ", "").replace(",", ".")
                try:
                    return float(normalized)
                except ValueError:
                    return text.lower()
            return value

        current_desc = getattr(tree, "_sort_desc", False)
        values.sort(key=lambda pair: convert(pair[0]), reverse=current_desc)

        # Toggle sort order for next click.
        tree._sort_desc = not current_desc

        for index, (_, item) in enumerate(values):
            tree.move(item, "", index)

        # Update heading text indicator.
        for col in tree["columns"]:
            heading_text = col
            if col == column:
                heading_text += " ↓" if current_desc else " ↑"
            tree.heading(col, text=heading_text, command=lambda c=col: self.sort_treeview(tree, c))

    def copy_tallkode_on_double_click(self, event: tk.Event, tree: ttk.Treeview) -> None:
        item_id = tree.identify_row(event.y)
        if not item_id:
            return

        columns = list(tree["columns"])
        if "Rema ID" in columns:
            key = "Rema ID"
        elif "Tallkode" in columns:
            key = "Tallkode"
        else:
            return

        values = tree.item(item_id, "values")
        try:
            tallkode = values[columns.index(key)]
        except (ValueError, IndexError):
            return

        if not tallkode:
            return

        self.root.clipboard_clear()
        self.root.clipboard_append(str(tallkode))
        messagebox.showinfo("Rema ID kopiert", f"Rema ID {tallkode} er kopiert til utklippstavlen.")

    def export_to_excel(self) -> None:
        if not self.analysis_result:
            messagebox.showerror("Ingen analyse", "Kjør analysen før du eksporterer.")
            return

        suggested_name = "butikkpotensial_resultat.xlsx"
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=suggested_name,
            filetypes=[("Excel-filer", "*.xlsx")],
        )
        if not path:
            return

        try:
            export_analysis(self.analysis_result, path)
        except Exception as exc:
            messagebox.showerror("Kunne ikke eksportere", str(exc))
            return

        messagebox.showinfo("Eksportert", f"Resultatet ble lagret i:\n{Path(path)}")

    def run(self) -> None:
        self.root.mainloop()


def launch_gui() -> None:
    app = AnalysisApp()
    app.run()


if __name__ == "__main__":
    launch_gui()
