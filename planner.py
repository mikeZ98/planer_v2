import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from datetime import datetime
from typing import Optional  # ← ważne

APP_TITLE = "Konwerter zadań: Excel → Excel (rok/kwartał/zasobnik)"
APP_GEOMETRY = "820x560"

# ←← USTAW ŚCIEŻKĘ DO DOMYŚLNEGO PLIKU (Excel #1)
DEFAULT_MASTER_PATH = r"C:\Sciezka\do\Excel1_master.xlsx"

# ==== KONFIG – data „fałszywa” z nowego systemu (Panasonic) ==================
SENTINEL_CREATED = datetime(2025, 4, 10).date()   # 10.04.2025

# ====== Pomocnicze ============================================================

def best_guess_column(columns, candidates):
    import unicodedata
    def norm(s):
        s = str(s)
        s = ''.join(c for c in unicodedata.normalize('NFKD', s) if not unicodedata.combining(c))
        s = s.lower()
        for ch in ['-', '_', ' ', '/', '\\', '.', '(', ')', '[', ']', ':', ';', '|']:
            s = s.replace(ch, '')
        return s
    cols_norm = {col: norm(col) for col in columns}
    cand_norm = [norm(c) for c in candidates]
    for col, ncol in cols_norm.items():
        if ncol in cand_norm:
            return col
    best, best_len = None, 0
    for col, ncol in cols_norm.items():
        for c in cand_norm:
            if c in ncol and len(c) > best_len:
                best, best_len = col, len(c)
    return best or (columns[0] if columns else None)

def detect_sheet(path):
    try:
        xls = pd.ExcelFile(path)
        preferred = [s for s in xls.sheet_names
                     if str(s).strip().lower() in {'data','dane','tasks','zadania','sheet1','arkusz1'}]
        return preferred[0] if preferred else xls.sheet_names[0]
    except Exception:
        return None

def norm_text(s: pd.Series) -> pd.Series:
    return (s.astype(str)
              .str.strip()
              .str.replace(r'\s+', ' ', regex=True)
              .str.replace('\u200b', '', regex=False)  # zero-width space
              .replace({'nan':'', 'None':''}))

def smart_datetime(s: pd.Series) -> pd.Series:
    """Solidne parsowanie dat (PL dd.mm.rrrr i ISO). Zwraca datetime64[ns]."""
    if not isinstance(s, pd.Series):
        s = pd.Series(s)
    dt = pd.to_datetime(s, errors='coerce', dayfirst=True, infer_datetime_format=True)
    # dopalacz dla  dd.mm.yyyy HH:MM  /  dd.mm.yyyy
    mask = dt.isna() & s.astype(str).str.contains(r"\d{1,2}\.\d{1,2}\.\d{2,4}")
    if mask.any():
        vals = s[mask].astype(str).str.strip()
        dt2 = pd.to_datetime(vals, format='%d.%m.%Y', errors='coerce')
        still = dt2.isna()
        if still.any():
            dt2.loc[still] = pd.to_datetime(vals[still], format='%d.%m.%Y %H:%M', errors='coerce')
        dt.loc[mask] = dt2
    return dt

def norm_date_string(s: pd.Series) -> pd.Series:
    dt = smart_datetime(s)
    return dt.dt.strftime('%Y-%m-%d').fillna('')

def filter_bins(person: pd.Series) -> pd.Series:
    up = person.str.upper()
    return ~(up.isin(["TO DO", "INSTRUKCJA", "HOLD"]))

def make_chart_df(df, col_person, col_created, col_completed):
    # Osoba/rok/kwartał + 2 serie: utworzone (wg Data utworzenia), zakończone (wg Data ukończenia)
    person = norm_text(df[col_person]).replace('', '—')
    keep = filter_bins(person)
    person = person[keep]
    created = smart_datetime(df[col_created])[keep]
    completed = smart_datetime(df[col_completed])[keep]

    created_df = (
        pd.DataFrame({'osoba': person, 'rok': created.dt.year, 'kwartał': created.dt.quarter})
        .dropna(subset=['osoba','rok','kwartał'])
        .groupby(['osoba','rok','kwartał']).size()
        .rename('Liczba UTWORZONYCH').reset_index()
    )
    done_df = (
        pd.DataFrame({'osoba': person, 'rok': completed.dt.year, 'kwartał': completed.dt.quarter})
        .dropna(subset=['rok','kwartał'])
        .groupby(['osoba','rok','kwartał']).size()
        .rename('Liczba ZAKOŃCZONYCH').reset_index()
    )
    pivot = pd.merge(created_df, done_df, on=['osoba','rok','kwartał'], how='outer')
    pivot['Liczba UTWORZONYCH']  = pivot['Liczba UTWORZONYCH'].fillna(0).astype(int)
    pivot['Liczba ZAKOŃCZONYCH'] = pivot['Liczba ZAKOŃCZONYCH'].fillna(0).astype(int)
    pivot = pivot.sort_values(by=['osoba','rok','kwartał'],
                              key=lambda s: s.str.casefold() if s.name=='osoba' else s).reset_index(drop=True)
    return pivot

def write_excel_with_chart(path, pivot_df, aggr_df=None, details_df=None, title="ZADANIE UTWORZONE DO ZAKOŃCZONE"):
    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        book = writer.book
        if aggr_df is not None:
            aggr_df.to_excel(writer, sheet_name="Agregat", index=False)
        if details_df is not None:
            details_df.to_excel(writer, sheet_name="Szczegóły", index=False)

        # === PIVOT (oś 3-poziomowa: Osoba / Rok / Kwartał) ===
        ws_pivot = book.add_worksheet('Pivot'); writer.sheets['Pivot'] = ws_pivot
        cat_person = pivot_df['osoba'].tolist()
        cat_year   = pivot_df['rok'].astype(int).tolist()
        cat_quart  = pivot_df['kwartał'].astype(int).map(lambda q: f"KWARTAŁ{q}").tolist()

        ws_pivot.write_row(0, 0, ["Osoba (oś)", *cat_person])
        ws_pivot.write_row(1, 0, ["Rok (oś)",   *cat_year])
        ws_pivot.write_row(2, 0, ["Kwartał (oś)", *cat_quart])

        ws_pivot.write(4, 0, "Liczba UTWORZONYCH")
        ws_pivot.write_row(4, 1, pivot_df["Liczba UTWORZONYCH"].tolist())
        ws_pivot.write(5, 0, "Liczba ZAKOŃCZONYCH")
        ws_pivot.write_row(5, 1, pivot_df["Liczba ZAKOŃCZONYCH"].tolist())

        # === DODATKOWA SEKCJA: SUMA PER KWARTAŁ (bez podziału na osoby) ===
        if not pivot_df.empty:
            per_q = (pivot_df
                     .groupby(['rok', 'kwartał'], as_index=False)[['Liczba UTWORZONYCH','Liczba ZAKOŃCZONYCH']]
                     .sum()
                     .sort_values(['rok','kwartał']))
            q_labels = [f"{int(r)} | KWARTAŁ{int(k)}" for r, k in zip(per_q['rok'], per_q['kwartał'])]
            ws_pivot.write_row(8, 0, ["Kwartał (oś)", *q_labels])
            ws_pivot.write(10, 0, "UTWORZONE (suma)")
            ws_pivot.write_row(10, 1, per_q['Liczba UTWORZONYCH'].tolist())
            ws_pivot.write(11, 0, "ZAKOŃCZONE (suma)")
            ws_pivot.write_row(11, 1, per_q['Liczba ZAKOŃCZONYCH'].tolist())
        else:
            q_labels = []

        # === Arkusz z wykresami ===
        ws_chart = book.add_worksheet('Wykres'); writer.sheets['Wykres'] = ws_chart

        # Wykres 1: OSOBY × KWARTAŁY (stacked)
        chart1 = book.add_chart({'type': 'column','subtype': 'stacked'})
        last_col = len(pivot_df)
        chart1.add_series({
            'name':       "='Pivot'!$A$5",   # UTWORZONE
            'categories': ['Pivot', 0, 1, 2, last_col],
            'values':     ['Pivot', 4, 1, 4, last_col],
            'data_labels': {'value': True},
            'fill': {'color': '#43A047'},  # zielony
            'border': {'color': '#43A047'},
        })
        chart1.add_series({
            'name':       "='Pivot'!$A$6",   # ZAKOŃCZONE
            'categories': ['Pivot', 0, 1, 2, last_col],
            'values':     ['Pivot', 5, 1, 5, last_col],
            'data_labels': {'value': True},
            'fill': {'color': '#1E88E5'},  # niebieski
            'border': {'color': '#1E88E5'},
        })
        chart1.set_title({'name': title})
        chart1.set_legend({'position': 'top'})
        chart1.set_y_axis({'major_gridlines': {'visible': True}})
        ws_chart.insert_chart('B2', chart1, {'x_scale': 2.0, 'y_scale': 1.6})

        # Wykres 2: SUMA PER KWARTAŁ (CLUSTERED, bez podziału na osoby)
        if q_labels:
            chart2 = book.add_chart({'type': 'column'})  # clustered
            last_col_q = len(q_labels)
            chart2.add_series({
                'name':       "='Pivot'!$A$11",  # UTWORZONE (suma)
                'categories': ['Pivot', 8, 1, 8, last_col_q],
                'values':     ['Pivot', 10, 1, 10, last_col_q],
                'data_labels': {'value': True},
                'fill': {'color': '#43A047'},
                'border': {'color': '#43A047'},
            })
            chart2.add_series({
                'name':       "='Pivot'!$A$12",  # ZAKOŃCZONE (suma)
                'categories': ['Pivot', 8, 1, 8, last_col_q],
                'values':     ['Pivot', 11, 1, 11, last_col_q],
                'data_labels': {'value': True},
                'fill': {'color': '#1E88E5'},
                'border': {'color': '#1E88E5'},
            })
            chart2.set_title({'name': 'Suma kwartalna — Utworzone vs Zakończone'})
            chart2.set_legend({'position': 'top'})
            chart2.set_y_axis({'major_gridlines': {'visible': True}})
            # "obok" pierwszego wykresu:
            ws_chart.insert_chart('N2', chart2, {'x_scale': 1.6, 'y_scale': 1.6})


# ====== Budowa klucza unikalności (anty-duplikacja) ===========================

def build_unique_key(df: pd.DataFrame,
                     col_person: str,
                     col_task: Optional[str],
                     col_created: str,
                     col_completed: str,
                     id_col: str = "Identyfikator zadania") -> pd.Series:
    """
    Klucz unikalności:
      1) Jeśli istnieje kolumna z ID – użyj jej.
      2) W przeciwnym razie:
         - bazowo: (osoba, nazwa_zadania, data_utworzenia)
         - jeśli data_utworzenia == 12.04.2025 (SENTINEL) → POMIŃ ją w kluczu
           i dołóż data_ukończenia (jeśli jest), aby rozróżnić różne zadania.
         - jeśli brak kolumny nazwy, użyj (osoba, [data_utworzenia?], [data_ukończenia?]) z powyższą regułą.
    """
    cols = list(df.columns)
    if id_col in cols:
        return norm_text(df[id_col])

    person = norm_text(df.get(col_person, pd.Series(['']*len(df)))).str.lower()
    task   = norm_text(df.get(col_task,   pd.Series(['']*len(df)))).str.lower() if col_task else pd.Series(['']*len(df))
    created_dt   = smart_datetime(df.get(col_created,   pd.Series(['']*len(df))))
    completed_dt = smart_datetime(df.get(col_completed, pd.Series(['']*len(df))))

    created_date   = created_dt.dt.date
    completed_date = completed_dt.dt.date

    created_str   = created_date.astype(str).where(pd.notna(created_date), "")
    completed_str = completed_date.astype(str).where(pd.notna(completed_date), "")

    use_created = ~created_date.eq(SENTINEL_CREATED)

    key = (
        person + "||" +
        task.fillna("") + "||" +
        created_str.where(use_created, "") + "||" +
        completed_str.where(use_created, completed_str)
    )
    return key

# ====== GUI ===================================================================

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry(APP_GEOMETRY)

        self.path_master = tk.StringVar(value=DEFAULT_MASTER_PATH)
        self.path_user   = tk.StringVar()

        self.sheet_master = tk.StringVar()
        self.sheet_user   = tk.StringVar()

        self.person_var = tk.StringVar()
        self.created_var = tk.StringVar()
        self.completed_var = tk.StringVar()
        self.task_var = tk.StringVar()  # (opcjonalnie) nazwa zadania do klucza

        self.df_master = None   # Excel #1
        self.df = None          # scalony master + user
        self.columns = []

        self.create_widgets()
        self.load_master_on_start()

    # --- UI ---
    def create_widgets(self):
        pad = {'padx': 8, 'pady': 6}
        frm = ttk.Frame(self); frm.pack(fill='both', expand=True)

        # MASTER (auto)
        row0 = ttk.LabelFrame(frm, text="Excel #1 (domyślny – ładowany automatycznie)")
        row0.pack(fill='x', **pad)
        r0 = ttk.Frame(row0); r0.pack(fill='x', **pad)
        ttk.Entry(r0, textvariable=self.path_master).pack(side='left', fill='x', expand=True, padx=6)
        ttk.Button(r0, text="Zmień…", command=self.pick_master).pack(side='left')
        r0b = ttk.Frame(row0); r0b.pack(fill='x', **pad)
        ttk.Label(r0b, text="Arkusz:").pack(side='left')
        self.sheet_combo_master = ttk.Combobox(r0b, textvariable=self.sheet_master, state='readonly', width=30)
        self.sheet_combo_master.pack(side='left', padx=6)
        ttk.Button(r0b, text="Przeładuj Excel #1", command=self.load_master_on_start).pack(side='left')

        # USER (scalanie)
        row1 = ttk.LabelFrame(frm, text="Excel #2 (od użytkownika – nowe wiersze zostaną dopisane)")
        row1.pack(fill='x', **pad)
        r1 = ttk.Frame(row1); r1.pack(fill='x', **pad)
        ttk.Entry(r1, textvariable=self.path_user).pack(side='left', fill='x', expand=True, padx=6)
        ttk.Button(r1, text="➕ Wczytaj plik 2 i scal", command=self.load_and_merge_user).pack(side='left')

        # Mapowanie
        box = ttk.LabelFrame(frm, text="Mapowanie kolumn")
        box.pack(fill='x', **pad)

        r_task = ttk.Frame(box); r_task.pack(fill='x', **pad)
        ttk.Label(r_task, text="Nazwa zadania (opcjonalnie, do klucza):").pack(side='left')
        self.task_combo = ttk.Combobox(r_task, textvariable=self.task_var, state='readonly', width=40)
        self.task_combo.pack(side='left', fill='x', expand=True, padx=6)

        r2 = ttk.Frame(box); r2.pack(fill='x', **pad)
        ttk.Label(r2, text="Zasobnik (osoba):").pack(side='left')
        self.person_combo = ttk.Combobox(r2, textvariable=self.person_var, state='readonly', width=40)
        self.person_combo.pack(side='left', fill='x', expand=True, padx=6)

        r3 = ttk.Frame(box); r3.pack(fill='x', **pad)
        ttk.Label(r3, text="Data utworzenia:").pack(side='left')
        self.created_combo = ttk.Combobox(r3, textvariable=self.created_var, state='readonly', width=40)
        self.created_combo.pack(side='left', fill='x', expand=True, padx=6)

        r4 = ttk.Frame(box); r4.pack(fill='x', **pad)
        ttk.Label(r4, text="Data ukończenia:").pack(side='left')
        self.completed_combo = ttk.Combobox(r4, textvariable=self.completed_var, state='readonly', width=40)
        self.completed_combo.pack(side='left', fill='x', expand=True, padx=6)

        # Akcje
        actions = ttk.Frame(frm); actions.pack(fill='x', **pad)
        ttk.Button(actions, text="Podgląd wyniku", command=self.preview).pack(side='right', padx=6)
        ttk.Button(actions, text="Konwertuj i zapisz…", command=self.convert_and_save).pack(side='right')

        # Log
        self.status = tk.Text(frm, height=10); self.status.pack(fill='both', expand=True, **pad)
        self.status.configure(state='disabled')

        self.log("Start: wczytam Excel #1 (master). Potem doładuj Excel #2 – dopiszę tylko nowe wiersze.")

    def log(self, msg):
        self.status.configure(state='normal')
        self.status.insert('end', f"{msg}\n")
        self.status.see('end')
        self.status.configure(state='disabled')

    # --- Master load ---
    def pick_master(self):
        path = filedialog.askopenfilename(title="Wybierz plik Excel #1 (master)",
                                          filetypes=[("Excel files","*.xlsx *.xls *.xlsm")])
        if path:
            self.path_master.set(path)
            self.load_master_on_start()

    def load_master_on_start(self):
        path = self.path_master.get().strip()
        if not path or not os.path.isfile(path):
            self.log("Brak/nieprawidłowa ścieżka Excel #1 – pomiń wczytanie.")
            return
        try:
            xls = pd.ExcelFile(path)
            self.sheet_combo_master['values'] = xls.sheet_names
            preferred = detect_sheet(path)
            self.sheet_master.set(preferred or (xls.sheet_names[0] if xls.sheet_names else ""))
            self.df_master = pd.read_excel(path, sheet_name=self.sheet_master.get() or 0)
            self.df = self.df_master.copy()  # na starcie zestaw = master
            self.columns = list(self.df_master.columns)

            # mapowanie
            self.person_combo['values'] = self.columns
            self.created_combo['values'] = self.columns
            self.completed_combo['values'] = self.columns
            self.task_combo['values'] = self.columns

            person_guess = best_guess_column(self.columns,
                ['zasobnik','osoba','assignee','owner','wykonawca','przypisane do','assigned to','user'])
            created_guess = best_guess_column(self.columns,
                ['data utworzenia','utworzenia','created','creation date','created at','start date'])
            completed_guess = best_guess_column(self.columns,
                ['data ukończenia','ukonczenia','completed','done','closed','end date','resolution date'])
            task_guess = best_guess_column(self.columns,
                ['nazwa zadania','tytuł','title','task','nazwa'])
            if person_guess: self.person_var.set(person_guess)
            if created_guess: self.created_var.set(created_guess)
            if completed_guess: self.completed_var.set(completed_guess)
            if task_guess: self.task_var.set(task_guess)

            self.log(f"Załadowano Excel #1: {os.path.basename(path)} | Arkusz: {self.sheet_master.get()} | Wierszy: {len(self.df_master)}")
        except Exception as e:
            messagebox.showerror("Błąd", f"Nie udało się wczytać Excel #1:\n{e}")
            self.log(f"Błąd wczytywania Excel #1: {e}")

    # --- Merge user file ---
    def load_and_merge_user(self):
        if self.df_master is None:
            messagebox.showwarning("Uwaga", "Najpierw wczytaj lub ustaw Excel #1 (master).")
            return
        path = filedialog.askopenfilename(title="Wybierz plik Excel #2",
                                        filetypes=[("Excel files","*.xlsx *.xls *.xlsm")])
        if not path:
            return
        try:
            df2 = pd.read_excel(path, sheet_name=detect_sheet(path) or 0)

            # filtr dat (tylko po dacie utworzenia)
            created_col = self.created_var.get()
            created_dt = smart_datetime(df2[created_col]).dt.date
            mask_keep = (created_dt > datetime(2025, 4, 12).date())  # tylko > 12.04.2025
            # UWAGA: wszystkie <= 09.04.2025 i == 12.04.2025 odrzucamy
            df2 = df2.loc[mask_keep].copy()

            # zbuduj klucze
            key_master = build_unique_key(
                self.df_master,
                self.person_var.get(),
                self.task_var.get() if self.task_var.get() else None,
                self.created_var.get(),
                self.completed_var.get()
            )
            key_new = build_unique_key(
                df2,
                self.person_var.get(),
                self.task_var.get() if self.task_var.get() else None,
                self.created_var.get(),
                self.completed_var.get()
            )

            df2['__key__'] = key_new
            master_keys = set(key_master.dropna().tolist())
            before = len(df2)
            df2_new = df2[~df2['__key__'].isin(master_keys)].drop(columns='__key__', errors='ignore')
            added = len(df2_new)

            self.df = pd.concat([self.df_master, df2_new], ignore_index=True)
            self.log(f"Excel #2: {os.path.basename(path)} | wierszy po filtrze: {before} | nowych dopisano: {added} | razem: {len(self.df)}")
        except Exception as e:
            messagebox.showerror("Błąd", f"Scalanie nie powiodło się:\n{e}")
            self.log(f"Błąd scalania: {e}")

    # --- Podgląd / zapis ---
    def preview(self):
        if self.df is None:
            messagebox.showwarning("Uwaga", "Brak danych. Wczytaj Excel #1 lub scal z #2.")
            return
        try:
            out, details = self._aggregate_and_details(self.df,
                                                       self.person_var.get(),
                                                       self.created_var.get(),
                                                       self.completed_var.get())
            top = out.head(50)
            self.log(f"AGREGAT (pierwsze {len(top)}):\n{top.to_string(index=False)}")
            self.log("Szczegóły (podgląd 10):")
            self.log(details.head(10).to_string(index=False))
        except Exception as e:
            messagebox.showerror("Błąd", f"Podgląd nie powiódł się:\n{e}")
            self.log(f"Błąd podglądu: {e}")

    def convert_and_save(self):
        if self.df is None:
            messagebox.showwarning("Uwaga", "Brak danych. Wczytaj Excel #1 lub scal z #2.")
            return
        try:
            out, details = self._aggregate_and_details(self.df,
                                                       self.person_var.get(),
                                                       self.created_var.get(),
                                                       self.completed_var.get())
            pivot = make_chart_df(self.df,
                                  self.person_var.get(),
                                  self.created_var.get(),
                                  self.completed_var.get())

            base = os.path.splitext(self.path_master.get().strip() or "wynik")[0]
            default_out = base + "_WYNIK_z_wykresem.xlsx"
            path = filedialog.asksaveasfilename(
                title="Zapisz wynik jako",
                defaultextension=".xlsx",
                initialfile=os.path.basename(default_out),
                filetypes=[("Excel files","*.xlsx")]
            )
            if not path:
                self.log("Anulowano zapis.")
                return

            write_excel_with_chart(path, pivot_df=pivot, aggr_df=out, details_df=details,
                                   title="ZADANIE UTWORZONE DO ZAKOŃCZONE")
            self.log(f"Zapisano wynik do: {path}")
            messagebox.showinfo("Sukces", f"Zapisano wynik do:\n{path}")
        except Exception as e:
            messagebox.showerror("Błąd", f"Zapis nie powiódł się:\n{e}")
            self.log(f"Błąd zapisu: {e}")

    # --- Agregaty / szczegóły ---
    def _aggregate_and_details(self, df, col_person, col_created, col_completed):
        person_raw = norm_text(df[col_person])
        person = person_raw.replace('', '—')
        keep = filter_bins(person)
        person = person[keep]
        created = smart_datetime(df[col_created])[keep]
        completed = smart_datetime(df[col_completed])[keep]

        # Ukończone wg daty ukończenia
        done = (pd.DataFrame({'rok': completed.dt.year, 'kwartał': completed.dt.quarter, 'zasobnik': person})
                .dropna(subset=['rok','kwartał'])
                .groupby(['rok','kwartał','zasobnik'], dropna=False).size()
                .rename('zadania_ukończone').reset_index())

        # Nieukończone wg braku ukończenia – liczone po dacie utworzenia
        not_done_mask = completed.isna()
        not_done_created = created.where(not_done_mask)
        not_done = (pd.DataFrame({'rok': not_done_created.dt.year,
                                  'kwartał': not_done_created.dt.quarter,
                                  'zasobnik': person.where(not_done_mask)})
                    .dropna(subset=['rok','kwartał','zasobnik'])
                    .groupby(['rok','kwartał','zasobnik'], dropna=False).size()
                    .rename('zadania_nieukończone').reset_index())

        out = pd.merge(done, not_done, on=['rok','kwartał','zasobnik'], how='outer')
        out['zadania_ukończone'] = out['zadania_ukończone'].fillna(0).astype(int)
        out['zadania_nieukończone'] = out['zadania_nieukończone'].fillna(0).astype(int)
        out.sort_values(by=['zasobnik','rok','kwartał'],
                        key=lambda col: col.str.casefold() if col.name=='zasobnik' else col,
                        inplace=True)
        out.reset_index(drop=True, inplace=True)

        # Szczegóły – oba zbiory (ukończone i nieukończone)
        det_done = pd.DataFrame({
            'osoba': person,
            'data_utworzenia': created,
            'data_ukonczenia': completed,
            'status': pd.Series(['ukończone']*len(person)),
            'rok': completed.dt.year,
            'kwartał': completed.dt.quarter,
        }).dropna(subset=['rok','kwartał'])

        det_not_done = pd.DataFrame({
            'osoba': person.where(not_done_mask),
            'data_utworzenia': created.where(not_done_mask),
            'data_ukonczenia': completed.where(not_done_mask),
            'status': pd.Series(['nieukończone']*len(person)).where(not_done_mask),
            'rok': not_done_created.dt.year,
            'kwartał': not_done_created.dt.quarter,
        }).dropna(subset=['osoba','rok','kwartał'])

        details = pd.concat([det_done, det_not_done], ignore_index=True).sort_values(
            by=['osoba','rok','kwartał','data_utworzenia','data_ukonczenia'],
            key=lambda col: col.str.casefold() if col.name=='osoba' else col
        ).reset_index(drop=True)

        return out, details

# ---- run ---------------------------------------------------------------------

def main():
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    app = App()
    app.mainloop()

if __name__ == "__main__":
    main()
