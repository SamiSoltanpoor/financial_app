import sqlite3
import os
import sys
from collections import defaultdict

import tkinter as tk
from tkinter import ttk, messagebox

try:
    import jdatetime
    HAS_JDATETIME = True
except ImportError:
    HAS_JDATETIME = False

try:
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    import matplotlib
    matplotlib.rcParams['axes.unicode_minus'] = False
    HAS_MPL = True
except ImportError:
    HAS_MPL = False

BG_MAIN = "#0f1220"
BG_CARD = "#1a1e33"
BG_CARD_ALT = "#20264a"
FG_TEXT = "#e8e9f3"
FG_MUTED = "#9096b5"
ACCENT = "#7c5cff"
ACCENT2 = "#4fd1c5"
GREEN = "#3ddc97"
RED = "#ff6b6b"
BLUE = "#4d96ff"
YELLOW = "#ffd166"
FONT_FAMILY = "Segoe UI" if sys.platform.startswith("win") else "Arial"

JALALI_MONTH_NAMES = [
    "فروردین", "اردیبهشت", "خرداد", "تیر", "مرداد", "شهریور",
    "مهر", "آبان", "آذر", "دی", "بهمن", "اسفند",
]
JALALI_MONTH_SHORT = [
    "فرو", "ارد", "خرد", "تیر", "مرد", "شهر",
    "مهر", "آبا", "آذر", "دی", "بهم", "اسف",
]

EXPENSE_CATEGORIES = [
    "خوراک و خواربار", "حمل‌ونقل", "مسکن و اجاره", "قبوض",
    "پوشاک", "سلامت و درمان", "سرگرمی", "آموزش", "متفرقه",
]
INCOME_CATEGORIES = [
    "حقوق", "فریلنسری", "سرمایه‌گذاری", "هدیه", "پاداش", "متفرقه",
]

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "expenses.db")

def jtoday():
    """تاریخ شمسی امروز را برمی‌گرداند (jdatetime.date)."""
    if HAS_JDATETIME:
        return jdatetime.date.today()
    # اگر jdatetime نصب نبود، یک تاریخ ثابت برای جلوگیری از کرش برمی‌گردانیم
    class _Fallback:
        year, month, day = 1403, 1, 1
    return _Fallback()


def parse_jalali(text):
    """رشته‌ی 'YYYY-MM-DD' شمسی را اعتبارسنجی و پارس می‌کند."""
    if not HAS_JDATETIME:
        raise RuntimeError("کتابخانه jdatetime نصب نیست")
    parts = text.strip().split("-")
    if len(parts) != 3:
        raise ValueError("فرمت باید YYYY-MM-DD باشد")
    y, m, d = (int(p) for p in parts)
    return jdatetime.date(y, m, d)  # خودش اعتبارسنجی می‌کند


def jalali_str(jdate):
    return f"{jdate.year:04d}-{jdate.month:02d}-{jdate.day:02d}"


def month_label(year, month):
    return f"{JALALI_MONTH_NAMES[month-1]} {year}"


def fmt_money(v):
    try:
        return f"{v:,.0f}"
    except Exception:
        return str(v)

class Database:
    def __init__(self, path=DB_PATH):
        self.conn = sqlite3.connect(path)
        self.conn.execute("""
            CREATE TABLE IF NOT EXISTS transactions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                kind TEXT NOT NULL CHECK(kind IN ('income','expense')),
                amount REAL NOT NULL,
                category TEXT NOT NULL,
                note TEXT,
                tdate TEXT NOT NULL
            )
        """)
        self.conn.commit()

    def add(self, kind, amount, category, note, tdate):
        self.conn.execute(
            "INSERT INTO transactions (kind, amount, category, note, tdate) VALUES (?,?,?,?,?)",
            (kind, amount, category, note, tdate),
        )
        self.conn.commit()

    def delete(self, tid):
        self.conn.execute("DELETE FROM transactions WHERE id=?", (tid,))
        self.conn.commit()

    def all(self):
        cur = self.conn.execute(
            "SELECT id, kind, amount, category, note, tdate FROM transactions ORDER BY tdate DESC, id DESC"
        )
        return cur.fetchall()

    def for_month(self, year, month):
        prefix = f"{year:04d}-{month:02d}"
        cur = self.conn.execute(
            "SELECT id, kind, amount, category, note, tdate FROM transactions "
            "WHERE tdate LIKE ? ORDER BY tdate DESC, id DESC",
            (prefix + "%",),
        )
        return cur.fetchall()

    def years_present(self):
        cur = self.conn.execute("SELECT DISTINCT substr(tdate,1,4) FROM transactions ORDER BY 1")
        years = [int(r[0]) for r in cur.fetchall()]
        ty = jtoday().year
        if ty not in years:
            years.append(ty)
        return sorted(set(years))

class StatCard(tk.Frame):
    """کارت آماری بالای داشبورد (درآمد / هزینه / پس‌انداز)."""

    def __init__(self, master, title, color, **kw):
        super().__init__(master, bg=BG_CARD, highlightbackground=color,
                          highlightthickness=1, bd=0, **kw)
        self.title = title
        self.color = color
        tk.Label(self, text=title, bg=BG_CARD, fg=FG_MUTED,
                  font=(FONT_FAMILY, 11)).pack(anchor="w", padx=16, pady=(14, 0))
        self.value_lbl = tk.Label(self, text="0", bg=BG_CARD, fg=color,
                                   font=(FONT_FAMILY, 22, "bold"))
        self.value_lbl.pack(anchor="w", padx=16, pady=(2, 14))

    def set_value(self, text):
        self.value_lbl.config(text=text)

class ExpenseApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.db = Database()
        self.title("مدیریت خرج و مخارج (تاریخ شمسی)")
        self.geometry("1080x680")
        self.minsize(880, 600)
        self.configure(bg=BG_MAIN)

        if not HAS_JDATETIME:
            messagebox.showwarning(
                "کتابخانه‌ی jdatetime یافت نشد",
                "برای کارکرد کامل تاریخ شمسی، کتابخانه‌ی jdatetime را نصب کنید:\n\n"
                "pip install jdatetime\n\n"
                "تا زمان نصب، تاریخ‌ها به‌درستی محاسبه نمی‌شوند.",
            )

        today = jtoday()
        self.sel_year = tk.IntVar(value=today.year)
        self.sel_month = tk.IntVar(value=today.month)

        self._setup_style()
        self._build_layout()
        self.refresh_all()

    def _setup_style(self):
        style = ttk.Style(self)
        style.theme_use("clam")

        style.configure("TNotebook", background=BG_MAIN, borderwidth=0)
        style.configure("TNotebook.Tab", background=BG_CARD, foreground=FG_MUTED,
                         padding=(18, 10), font=(FONT_FAMILY, 10, "bold"))
        style.map("TNotebook.Tab",
                  background=[("selected", ACCENT)],
                  foreground=[("selected", "#ffffff")])

        style.configure("TFrame", background=BG_MAIN)
        style.configure("Card.TFrame", background=BG_CARD)

        style.configure("TLabel", background=BG_MAIN, foreground=FG_TEXT, font=(FONT_FAMILY, 10))
        style.configure("Muted.TLabel", background=BG_MAIN, foreground=FG_MUTED)
        style.configure("Card.TLabel", background=BG_CARD, foreground=FG_TEXT)

        style.configure("TButton", background=ACCENT, foreground="#ffffff",
                         font=(FONT_FAMILY, 10, "bold"), padding=(14, 8), borderwidth=0)
        style.map("TButton", background=[("active", "#6a4bee")])

        style.configure("Danger.TButton", background=RED, foreground="#ffffff")
        style.map("Danger.TButton", background=[("active", "#e05555")])

        style.configure("TCombobox", fieldbackground=BG_CARD_ALT, background=BG_CARD_ALT,
                         foreground=FG_TEXT, arrowcolor=FG_TEXT)
        style.configure("TEntry", fieldbackground=BG_CARD_ALT, foreground=FG_TEXT,
                         insertcolor=FG_TEXT)
        style.configure("TRadiobutton", background=BG_MAIN, foreground=FG_TEXT,
                         font=(FONT_FAMILY, 10))
        style.map("TRadiobutton", background=[("active", BG_MAIN)])

        style.configure("Treeview", background=BG_CARD, fieldbackground=BG_CARD,
                         foreground=FG_TEXT, rowheight=28, borderwidth=0,
                         font=(FONT_FAMILY, 10))
        style.configure("Treeview.Heading", background=BG_CARD_ALT, foreground=FG_TEXT,
                         font=(FONT_FAMILY, 10, "bold"), relief="flat")
        style.map("Treeview", background=[("selected", ACCENT)])

    def _build_layout(self):
        header = tk.Frame(self, bg=BG_MAIN)
        header.pack(fill="x", padx=24, pady=(20, 10))
        tk.Label(header, text="💰 مدیریت خرج و مخارج", bg=BG_MAIN, fg=FG_TEXT,
                  font=(FONT_FAMILY, 20, "bold")).pack(side="right")
        tk.Label(header, text="کنترل کامل درآمد، هزینه و پس‌انداز شما — تقویم شمسی",
                  bg=BG_MAIN, fg=FG_MUTED, font=(FONT_FAMILY, 11)).pack(side="right", padx=12)

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=20, pady=10)
        self.nb = nb

        self.tab_dashboard = ttk.Frame(nb)
        self.tab_add = ttk.Frame(nb)
        self.tab_history = ttk.Frame(nb)
        self.tab_report = ttk.Frame(nb)

        nb.add(self.tab_dashboard, text="📊 داشبورد")
        nb.add(self.tab_add, text="➕ ثبت تراکنش")
        nb.add(self.tab_history, text="📜 تاریخچه")
        nb.add(self.tab_report, text="📈 گزارش ماهانه / سالانه")

        nb.bind("<<NotebookTabChanged>>", lambda e: self.refresh_all())

        self._build_dashboard(self.tab_dashboard)
        self._build_add_form(self.tab_add)
        self._build_history(self.tab_history)
        self._build_report(self.tab_report)

    def _build_dashboard(self, parent):
        top = tk.Frame(parent, bg=BG_MAIN)
        top.pack(fill="x", pady=(10, 20))

        tk.Label(top, text="ماه جاری:", bg=BG_MAIN, fg=FG_MUTED,
                  font=(FONT_FAMILY, 11)).pack(side="right", padx=(0, 8))
        self.dash_month_lbl = tk.Label(top, text="", bg=BG_MAIN, fg=ACCENT2,
                                        font=(FONT_FAMILY, 13, "bold"))
        self.dash_month_lbl.pack(side="right")

        cards = tk.Frame(parent, bg=BG_MAIN)
        cards.pack(fill="x")
        cards.grid_columnconfigure((0, 1, 2), weight=1, uniform="c")

        self.card_income = StatCard(cards, "درآمد این ماه", GREEN)
        self.card_expense = StatCard(cards, "هزینه این ماه", RED)
        self.card_saving = StatCard(cards, "پس‌انداز این ماه", BLUE)
        self.card_income.grid(row=0, column=2, sticky="nsew", padx=6)
        self.card_expense.grid(row=0, column=1, sticky="nsew", padx=6)
        self.card_saving.grid(row=0, column=0, sticky="nsew", padx=6)

        chart_wrap = tk.Frame(parent, bg=BG_CARD)
        chart_wrap.pack(fill="both", expand=True, pady=(20, 0))
        tk.Label(chart_wrap, text="تفکیک هزینه‌ها بر اساس دسته (ماه جاری)",
                  bg=BG_CARD, fg=FG_TEXT, font=(FONT_FAMILY, 12, "bold")).pack(
                      anchor="e", padx=16, pady=(12, 0))
        self.dash_chart_holder = tk.Frame(chart_wrap, bg=BG_CARD)
        self.dash_chart_holder.pack(fill="both", expand=True, padx=10, pady=10)

    def refresh_dashboard(self):
        today = jtoday()
        rows = self.db.for_month(today.year, today.month)
        income = sum(r[2] for r in rows if r[1] == "income")
        expense = sum(r[2] for r in rows if r[1] == "expense")
        saving = income - expense

        self.dash_month_lbl.config(text=month_label(today.year, today.month))
        self.card_income.set_value(fmt_money(income))
        self.card_expense.set_value(fmt_money(expense))
        self.card_saving.set_value(fmt_money(saving))
        self.card_saving.value_lbl.config(fg=GREEN if saving >= 0 else RED)

        for w in self.dash_chart_holder.winfo_children():
            w.destroy()

        cat_totals = defaultdict(float)
        for r in rows:
            if r[1] == "expense":
                cat_totals[r[3]] += r[2]

        if not HAS_MPL:
            tk.Label(self.dash_chart_holder, text="برای نمایش نمودار، matplotlib را نصب کنید:\npip install matplotlib",
                      bg=BG_CARD, fg=FG_MUTED, font=(FONT_FAMILY, 10)).pack(expand=True)
            return

        if not cat_totals:
            tk.Label(self.dash_chart_holder, text="در این ماه هنوز هزینه‌ای ثبت نشده است.",
                      bg=BG_CARD, fg=FG_MUTED, font=(FONT_FAMILY, 11)).pack(expand=True)
            return

        fig = Figure(figsize=(5.5, 3.6), dpi=100)
        fig.patch.set_facecolor(BG_CARD)
        ax = fig.add_subplot(111)
        ax.set_facecolor(BG_CARD)
        labels = list(cat_totals.keys())
        values = list(cat_totals.values())
        colors = [ACCENT, ACCENT2, GREEN, RED, BLUE, YELLOW, "#ff9f6b", "#c792ea", "#82aaff"]
        wedges, _texts, autotexts = ax.pie(
            values, autopct="%1.0f%%", colors=colors, textprops={"color": "#ffffff", "fontsize": 9},
            startangle=90, wedgeprops={"linewidth": 1, "edgecolor": BG_CARD},
        )
        ax.legend(wedges, labels, loc="center left", bbox_to_anchor=(1.0, 0.5),
                   frameon=False, labelcolor=FG_TEXT, fontsize=8)
        fig.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=self.dash_chart_holder)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

    def _build_add_form(self, parent):
        wrap = tk.Frame(parent, bg=BG_CARD)
        wrap.pack(fill="x", padx=40, pady=40)

        self.kind_var = tk.StringVar(value="expense")
        kind_frame = tk.Frame(wrap, bg=BG_CARD)
        kind_frame.pack(fill="x", padx=20, pady=(20, 10))
        ttk.Radiobutton(kind_frame, text="هزینه", variable=self.kind_var,
                        value="expense", command=self._refresh_category_list).pack(side="right", padx=10)
        ttk.Radiobutton(kind_frame, text="درآمد", variable=self.kind_var,
                        value="income", command=self._refresh_category_list).pack(side="right", padx=10)

        grid = tk.Frame(wrap, bg=BG_CARD)
        grid.pack(fill="x", padx=20, pady=10)

        def row(label, widget_factory, r):
            tk.Label(grid, text=label, bg=BG_CARD, fg=FG_TEXT,
                      font=(FONT_FAMILY, 11)).grid(row=r, column=1, sticky="e", padx=10, pady=10)
            w = widget_factory()
            w.grid(row=r, column=0, sticky="ew", padx=10, pady=10)
            return w

        grid.grid_columnconfigure(0, weight=1)

        self.amount_var = tk.StringVar()
        self.amount_entry = row("مبلغ (تومان)", lambda: ttk.Entry(
            grid, textvariable=self.amount_var, font=(FONT_FAMILY, 11)), 0)

        self.category_var = tk.StringVar()
        self.category_combo = row("دسته‌بندی", lambda: ttk.Combobox(
            grid, textvariable=self.category_var, state="readonly",
            values=EXPENSE_CATEGORIES, font=(FONT_FAMILY, 11)), 1)

        self.note_var = tk.StringVar()
        row("توضیحات (اختیاری)", lambda: ttk.Entry(
            grid, textvariable=self.note_var, font=(FONT_FAMILY, 11)), 2)

        self.date_var = tk.StringVar(value=jalali_str(jtoday()))
        row("تاریخ شمسی (YYYY-MM-DD)", lambda: ttk.Entry(
            grid, textvariable=self.date_var, font=(FONT_FAMILY, 11)), 3)
        tk.Label(grid, text="مثال: 1405-03-11", bg=BG_CARD, fg=FG_MUTED,
                  font=(FONT_FAMILY, 9)).grid(row=4, column=0, sticky="w", padx=10)

        self._refresh_category_list()

        btn_frame = tk.Frame(wrap, bg=BG_CARD)
        btn_frame.pack(fill="x", padx=20, pady=(10, 24))
        ttk.Button(btn_frame, text="ثبت تراکنش", command=self._submit_transaction).pack(
            side="right")
        ttk.Button(btn_frame, text="امروز", command=self._set_today).pack(side="right", padx=8)
        self.add_status_lbl = tk.Label(btn_frame, text="", bg=BG_CARD, fg=GREEN,
                                        font=(FONT_FAMILY, 10))
        self.add_status_lbl.pack(side="right", padx=14)

    def _set_today(self):
        self.date_var.set(jalali_str(jtoday()))

    def _refresh_category_list(self):
        cats = EXPENSE_CATEGORIES if self.kind_var.get() == "expense" else INCOME_CATEGORIES
        self.category_combo["values"] = cats
        if self.category_var.get() not in cats:
            self.category_var.set(cats[0])

    def _submit_transaction(self):
        amount_raw = self.amount_var.get().strip().replace(",", "")
        try:
            amount = float(amount_raw)
            if amount <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("خطا", "مبلغ را به‌صورت یک عدد مثبت وارد کنید.")
            return

        if not HAS_JDATETIME:
            messagebox.showerror("خطا", "کتابخانه jdatetime نصب نیست:\npip install jdatetime")
            return
        try:
            jd = parse_jalali(self.date_var.get())
        except ValueError:
            messagebox.showerror("خطا", "تاریخ شمسی نامعتبر است. فرمت درست: 1405-03-11")
            return

        if not self.category_var.get():
            messagebox.showerror("خطا", "یک دسته‌بندی انتخاب کنید.")
            return

        self.db.add(
            self.kind_var.get(), amount, self.category_var.get(),
            self.note_var.get().strip(), jalali_str(jd),
        )
        self.amount_var.set("")
        self.note_var.set("")
        self.add_status_lbl.config(text="✓ با موفقیت ثبت شد")
        self.after(2000, lambda: self.add_status_lbl.config(text=""))
        self.refresh_all()

    def _build_history(self, parent):
        filt = tk.Frame(parent, bg=BG_MAIN)
        filt.pack(fill="x", pady=(10, 10))

        tk.Label(filt, text="فیلتر بر اساس ماه:", bg=BG_MAIN, fg=FG_MUTED,
                  font=(FONT_FAMILY, 10)).pack(side="right", padx=6)

        self.hist_year_cb = ttk.Combobox(filt, state="readonly", width=8, font=(FONT_FAMILY, 10))
        self.hist_year_cb.pack(side="right", padx=4)
        self.hist_month_cb = ttk.Combobox(filt, state="readonly", width=10, font=(FONT_FAMILY, 10),
                                           values=JALALI_MONTH_NAMES)
        self.hist_month_cb.pack(side="right", padx=4)
        ttk.Button(filt, text="اعمال فیلتر", command=self.refresh_history).pack(side="right", padx=8)
        ttk.Button(filt, text="نمایش همه", command=lambda: self.refresh_history(show_all=True)).pack(
            side="right", padx=4)

        cols = ("date", "kind", "category", "amount", "note", "id")
        self.tree = ttk.Treeview(parent, columns=cols, show="headings", selectmode="browse")
        headings = {
            "date": "تاریخ شمسی", "kind": "نوع", "category": "دسته‌بندی",
            "amount": "مبلغ", "note": "توضیحات", "id": "شناسه",
        }
        widths = {"date": 110, "kind": 70, "category": 140, "amount": 110, "note": 250, "id": 0}
        for c in cols:
            self.tree.heading(c, text=headings[c])
            self.tree.column(c, width=widths[c], anchor="center", stretch=(c != "id"))
        self.tree.column("id", width=0, stretch=False)
        self.tree["displaycolumns"] = ("date", "kind", "category", "amount", "note")
        self.tree.pack(fill="both", expand=True, pady=6)

        self.tree.tag_configure("income", foreground=GREEN)
        self.tree.tag_configure("expense", foreground=RED)

        bottom = tk.Frame(parent, bg=BG_MAIN)
        bottom.pack(fill="x", pady=6)
        ttk.Button(bottom, text="حذف تراکنش انتخاب‌شده", style="Danger.TButton",
                   command=self._delete_selected).pack(side="right")
        self.hist_summary_lbl = tk.Label(bottom, text="", bg=BG_MAIN, fg=FG_MUTED,
                                          font=(FONT_FAMILY, 10))
        self.hist_summary_lbl.pack(side="left")

    def refresh_history(self, show_all=False):
        years = self.db.years_present()
        self.hist_year_cb["values"] = years
        today = jtoday()
        if not self.hist_year_cb.get():
            self.hist_year_cb.set(today.year)
        if not self.hist_month_cb.get():
            self.hist_month_cb.set(JALALI_MONTH_NAMES[today.month - 1])

        for i in self.tree.get_children():
            self.tree.delete(i)

        if show_all:
            rows = self.db.all()
        else:
            try:
                yr = int(self.hist_year_cb.get())
                mo = JALALI_MONTH_NAMES.index(self.hist_month_cb.get()) + 1
            except (ValueError, IndexError):
                yr, mo = today.year, today.month
            rows = self.db.for_month(yr, mo)

        income = expense = 0
        for tid, kind, amount, category, note, tdate in rows:
            kind_fa = "درآمد" if kind == "income" else "هزینه"
            self.tree.insert("", "end", values=(tdate, kind_fa, category,
                                                  fmt_money(amount), note or "-", tid),
                              tags=(kind,))
            if kind == "income":
                income += amount
            else:
                expense += amount

        self.hist_summary_lbl.config(
            text=f"جمع درآمد: {fmt_money(income)}   |   جمع هزینه: {fmt_money(expense)}   |   "
                 f"پس‌انداز: {fmt_money(income - expense)}   |   تعداد تراکنش: {len(rows)}"
        )

    def _delete_selected(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("توجه", "ابتدا یک تراکنش را از لیست انتخاب کنید.")
            return
        item = self.tree.item(sel[0])
        tid = item["values"][-1]
        if messagebox.askyesno("تأیید حذف", "از حذف این تراکنش مطمئن هستید؟"):
            self.db.delete(tid)
            self.refresh_all()

    def _build_report(self, parent):
        top = tk.Frame(parent, bg=BG_MAIN)
        top.pack(fill="x", pady=(10, 10))
        tk.Label(top, text="سال شمسی:", bg=BG_MAIN, fg=FG_MUTED,
                  font=(FONT_FAMILY, 10)).pack(side="right", padx=6)
        self.report_year_cb = ttk.Combobox(top, state="readonly", width=8, font=(FONT_FAMILY, 10))
        self.report_year_cb.pack(side="right")
        ttk.Button(top, text="بروزرسانی نمودار", command=self.refresh_report).pack(side="right", padx=10)

        self.report_chart_holder = tk.Frame(parent, bg=BG_CARD)
        self.report_chart_holder.pack(fill="both", expand=True, pady=10)

    def refresh_report(self):
        years = self.db.years_present()
        self.report_year_cb["values"] = years
        if not self.report_year_cb.get():
            self.report_year_cb.set(jtoday().year)
        try:
            year = int(self.report_year_cb.get())
        except ValueError:
            year = jtoday().year

        for w in self.report_chart_holder.winfo_children():
            w.destroy()

        if not HAS_MPL:
            tk.Label(self.report_chart_holder,
                      text="برای نمایش نمودار، matplotlib را نصب کنید:\npip install matplotlib",
                      bg=BG_CARD, fg=FG_MUTED, font=(FONT_FAMILY, 10)).pack(expand=True)
            return

        incomes, expenses, savings = [], [], []
        for m in range(1, 13):
            rows = self.db.for_month(year, m)
            inc = sum(r[2] for r in rows if r[1] == "income")
            exp = sum(r[2] for r in rows if r[1] == "expense")
            incomes.append(inc)
            expenses.append(exp)
            savings.append(inc - exp)

        fig = Figure(figsize=(9, 4.2), dpi=100)
        fig.patch.set_facecolor(BG_CARD)
        ax = fig.add_subplot(111)
        ax.set_facecolor(BG_CARD)

        x = range(12)
        width = 0.27
        ax.bar([i - width for i in x], incomes, width=width, label="درآمد", color=GREEN)
        ax.bar(list(x), expenses, width=width, label="هزینه", color=RED)
        ax.bar([i + width for i in x], savings, width=width, label="پس‌انداز", color=BLUE)

        ax.set_xticks(list(x))
        ax.set_xticklabels(JALALI_MONTH_SHORT, color=FG_TEXT, fontsize=9)
        ax.tick_params(colors=FG_TEXT)
        for spine in ax.spines.values():
            spine.set_color(BG_CARD_ALT)
        ax.axhline(0, color=BG_CARD_ALT, linewidth=1)
        ax.legend(facecolor=BG_CARD, edgecolor=BG_CARD, labelcolor=FG_TEXT)
        ax.set_title(f"گزارش سالانه {year} (تقویم شمسی)", color=FG_TEXT, fontsize=12)
        fig.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=self.report_chart_holder)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

        total_income = sum(incomes)
        total_expense = sum(expenses)
        summary = tk.Label(
            self.report_chart_holder,
            text=(f"جمع کل سال {year}   —   درآمد: {fmt_money(total_income)}   |   "
                  f"هزینه: {fmt_money(total_expense)}   |   "
                  f"پس‌انداز خالص: {fmt_money(total_income - total_expense)}"),
            bg=BG_CARD, fg=FG_TEXT, font=(FONT_FAMILY, 11, "bold"),
        )
        summary.pack(pady=(0, 12))

    def refresh_all(self):
        self.refresh_dashboard()
        self.refresh_history()
        self.refresh_report()


if __name__ == "__main__":
    app = ExpenseApp()
    app.mainloop()