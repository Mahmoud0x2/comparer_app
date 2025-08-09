import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
from thefuzz import fuzz
import numpy as np
import re
import threading
import queue
from datetime import datetime
import sv_ttk  # مكتبة النمط العصري

# -----------------------------------------------------------------------------
# الفئة الرئيسية لتطبيق الواجهة الرسومية المطور والاحترافي
# -----------------------------------------------------------------------------
class SkuMultiToolApp:
    def __init__(self, root):
        """
        إعداد النافذة الرئيسية، النمط الحديث، وجميع عناصر الواجهة.
        يتطلب تثبيت: pip install sv-ttk
        """
        self.root = root
        self.root.title("مجموعة أدوات SKU الاحترافية V5.0")
        self.root.geometry("1100x800")
        self.root.minsize(950, 700)

        # إعداد النمط الحديث
        sv_ttk.set_theme("dark")

        self.log_queue = queue.Queue()
        
        # إنشاء واجهة التبويبات
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(expand=True, fill="both", padx=10, pady=10)

        # إنشاء التبويبات
        self.tab1 = ttk.Frame(self.notebook, padding=10)
        self.tab2 = ttk.Frame(self.notebook, padding=10)

        self.notebook.add(self.tab1, text='مقارنة SKUs')
        self.notebook.add(self.tab2, text='سحب بيانات المنتجات')

        self.create_comparison_tab()
        self.create_fetcher_tab()
        
        self.process_queue()

    # --------------------------------------------------------------------
    # القسم الخاص بإنشاء تبويب "مقارنة SKUs"
    # --------------------------------------------------------------------
    def create_comparison_tab(self):
        self.c_file1_path = tk.StringVar()
        self.c_file1_sheet = tk.StringVar()
        self.c_file1_col = tk.StringVar()
        self.c_file2_path = tk.StringVar()
        self.c_file2_sheet = tk.StringVar()
        self.c_file2_col = tk.StringVar()
        self.c_similarity_threshold = tk.IntVar(value=85)
        self.c_save_option = tk.StringVar(value="new_file")

        top_frame = ttk.Frame(self.tab1)
        top_frame.pack(fill=tk.X)
        bottom_frame = ttk.Frame(self.tab1)
        bottom_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        files_frame = ttk.Frame(top_frame)
        files_frame.pack(fill=tk.X, expand=True)
        self.c_file1_widgets = self.create_file_input_frame(files_frame, "الملف الأول", self.c_file1_path, self.c_file1_sheet, self.c_file1_col, 'c_file1')
        self.c_file2_widgets = self.create_file_input_frame(files_frame, "الملف الثاني", self.c_file2_path, self.c_file2_sheet, self.c_file2_col, 'c_file2')
        
        control_frame = ttk.LabelFrame(top_frame, text="الإعدادات والتحكم", padding="10")
        control_frame.pack(fill=tk.X, expand=True, pady=10)

        settings_pane = ttk.Frame(control_frame)
        settings_pane.pack(side=tk.LEFT, padx=10, fill=tk.Y)
        ttk.Label(settings_pane, text="عتبة التشابه (%):").pack(anchor='w')
        ttk.Spinbox(settings_pane, from_=0, to=100, textvariable=self.c_similarity_threshold, width=10).pack(anchor='w', fill='x', pady=(0,10))
        ttk.Label(settings_pane, text="خيارات الحفظ:").pack(anchor='w', pady=(10,0))
        ttk.Radiobutton(settings_pane, text="ملف جديد", variable=self.c_save_option, value="new_file").pack(anchor='w')
        self.c_save_file1_rb = ttk.Radiobutton(settings_pane, text="صفحة جديدة في الملف الأول", variable=self.c_save_option, value="file1", state='disabled')
        self.c_save_file1_rb.pack(anchor='w')
        self.c_save_file2_rb = ttk.Radiobutton(settings_pane, text="صفحة جديدة في الملف الثاني", variable=self.c_save_option, value="file2", state='disabled')
        self.c_save_file2_rb.pack(anchor='w')

        run_pane = ttk.Frame(control_frame)
        run_pane.pack(side=tk.LEFT, padx=20, fill=tk.BOTH, expand=True)
        self.c_run_button = ttk.Button(run_pane, text="بدء المقارنة", command=self.start_comparison_thread, style="Accent.TButton")
        self.c_run_button.pack(pady=10, ipady=5, fill='x')
        self.c_progress_bar = ttk.Progressbar(run_pane, mode='indeterminate')
        self.c_progress_bar.pack(fill='x', pady=5)
        
        stats_container = ttk.LabelFrame(bottom_frame, text="ملخص النتائج", padding="10")
        stats_container.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        cards_holder_frame = ttk.Frame(stats_container)
        cards_holder_frame.pack()
        self.c_confirmed_var = self.create_stat_card(cards_holder_frame, "✓ تطابق مؤكد")
        self.c_fuzzy_var = self.create_stat_card(cards_holder_frame, "~ تطابق تقريبي")
        self.c_unmatched1_var = self.create_stat_card(cards_holder_frame, "✗ غير متطابق (ملف 1)")
        self.c_unmatched2_var = self.create_stat_card(cards_holder_frame, "✗ غير متطابق (ملف 2)")

        self.c_log_container = ttk.LabelFrame(bottom_frame, text="سجل العمليات", padding="10")
        self.c_log_container.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.c_log_text = scrolledtext.ScrolledText(self.c_log_container, wrap=tk.WORD, state="disabled", font=('Consolas', 9))
        self.c_log_text.pack(expand=True, fill="both")

    # --------------------------------------------------------------------
    # القسم الخاص بإنشاء تبويب "سحب بيانات المنتجات"
    # --------------------------------------------------------------------
    def create_fetcher_tab(self):
        self.f_file_path = tk.StringVar()
        self.f_source_sheet = tk.StringVar()
        self.f_source_col = tk.StringVar()
        self.f_search_sheet = tk.StringVar()
        self.f_search_col = tk.StringVar()

        top_frame = ttk.Frame(self.tab2)
        top_frame.pack(fill=tk.X)
        bottom_frame = ttk.Frame(self.tab2)
        bottom_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        # إطار اختيار الملف
        file_frame = ttk.LabelFrame(top_frame, text="ملف العمل", padding="10")
        file_frame.pack(fill=tk.X, expand=True, pady=5)
        ttk.Label(file_frame, text="مسار الملف:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(file_frame, textvariable=self.f_file_path, state="readonly").grid(row=1, column=0, sticky="ew", padx=5, pady=2)
        ttk.Button(file_frame, text="اختر...", command=self.f_select_file, style="Accent.TButton").grid(row=1, column=1, sticky="e", padx=5, pady=2)
        file_frame.columnconfigure(0, weight=1)

        # إطار الإعدادات
        settings_frame = ttk.Frame(top_frame)
        settings_frame.pack(fill=tk.X, expand=True, pady=5)
        
        source_frame = ttk.LabelFrame(settings_frame, text="1. الصفحة المصدر (البيانات الكاملة)", padding="10")
        source_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Label(source_frame, text="اختر الصفحة:").pack(anchor='w')
        self.f_source_sheet_combo = ttk.Combobox(source_frame, textvariable=self.f_source_sheet, state="disabled")
        self.f_source_sheet_combo.pack(fill='x', pady=5)
        ttk.Label(source_frame, text="اختر عمود المطابقة:").pack(anchor='w')
        self.f_source_col_combo = ttk.Combobox(source_frame, textvariable=self.f_source_col, state="disabled")
        self.f_source_col_combo.pack(fill='x', pady=5)

        search_frame = ttk.LabelFrame(settings_frame, text="2. صفحة البحث (قائمة المنتجات)", padding="10")
        search_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Label(search_frame, text="اختر الصفحة:").pack(anchor='w')
        self.f_search_sheet_combo = ttk.Combobox(search_frame, textvariable=self.f_search_sheet, state="disabled")
        self.f_search_sheet_combo.pack(fill='x', pady=5)
        ttk.Label(search_frame, text="اختر عمود المنتجات:").pack(anchor='w')
        self.f_search_col_combo = ttk.Combobox(search_frame, textvariable=self.f_search_col, state="disabled")
        self.f_search_col_combo.pack(fill='x', pady=5)
        
        # ربط الأحداث
        self.f_source_sheet_combo.bind("<<ComboboxSelected>>", lambda e: self.f_on_sheet_select('source'))
        self.f_search_sheet_combo.bind("<<ComboboxSelected>>", lambda e: self.f_on_sheet_select('search'))

        # قسم التشغيل والنتائج
        run_frame = ttk.LabelFrame(bottom_frame, text="التشغيل والنتائج", padding="10")
        run_frame.pack(fill=tk.BOTH, expand=True)
        self.f_run_button = ttk.Button(run_frame, text="بدء سحب البيانات", command=self.start_fetcher_thread, style="Accent.TButton")
        self.f_run_button.pack(pady=10, ipady=5, fill='x')
        self.f_progress_bar = ttk.Progressbar(run_frame, mode='indeterminate')
        self.f_progress_bar.pack(fill='x', pady=5)
        self.f_summary_label = ttk.Label(run_frame, text="في انتظار البدء...", justify=tk.LEFT, font=('Segoe UI', 12))
        self.f_summary_label.pack(fill='x', pady=10)

    # --------------------------------------------------------------------
    # طرق مساعدة مشتركة وعامة
    # --------------------------------------------------------------------
    def create_file_input_frame(self, parent, title, file_var, sheet_var, col_var, frame_id):
        frame = ttk.LabelFrame(parent, text=title, padding="10")
        frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, pady=5)
        ttk.Label(frame, text="مسار الملف:").grid(row=0, column=0, columnspan=3, sticky="w", padx=5, pady=2)
        ttk.Entry(frame, textvariable=file_var, state="readonly").grid(row=1, column=0, sticky="ew", padx=5, pady=2)
        ttk.Button(frame, text="اختر...", command=lambda: self.c_select_file(file_var, sheet_var, col_var), style="Accent.TButton").grid(row=1, column=1, sticky="e", padx=5, pady=2)
        if frame_id == 'c_file2':
            ttk.Button(frame, text="نفس الملف الأول", command=self.c_use_same_file).grid(row=1, column=2, sticky="e", padx=5, pady=2)
        frame.columnconfigure(0, weight=1)
        ttk.Label(frame, text="الصفحة (Sheet):").grid(row=2, column=0, columnspan=3, sticky="w", padx=5, pady=2)
        sheet_combo = ttk.Combobox(frame, textvariable=sheet_var, state="disabled")
        sheet_combo.grid(row=3, column=0, columnspan=3, sticky="ew", padx=5, pady=2)
        ttk.Label(frame, text="عمود SKU:").grid(row=4, column=0, columnspan=3, sticky="w", padx=5, pady=2)
        col_combo = ttk.Combobox(frame, textvariable=col_var, state="disabled")
        col_combo.grid(row=5, column=0, columnspan=3, sticky="ew", padx=5, pady=2)
        sheet_combo.bind("<<ComboboxSelected>>", lambda e, s=sheet_var, c=col_var, f=file_var: self.c_on_sheet_select(f, s, c))
        return {'sheet': sheet_combo, 'col': col_combo}

    def create_stat_card(self, parent, title):
        card_frame = ttk.Frame(parent, padding=10, style='Card.TFrame')
        card_frame.pack(side=tk.LEFT, fill='x', expand=True, pady=5, padx=5)
        ttk.Label(card_frame, text=title, font=('Segoe UI', 11, 'bold')).pack()
        value_var = tk.StringVar(value="0")
        ttk.Label(card_frame, textvariable=value_var, font=('Segoe UI', 22, 'bold'), foreground='#007ACC').pack()
        return value_var

    def log(self, message, tool='comparison'):
        log_text = self.c_log_text if tool == 'comparison' else self.f_summary_label
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        if tool == 'comparison':
            log_text.config(state="normal")
            log_text.insert(tk.END, f"[{timestamp}] {message}\n")
            log_text.config(state="disabled")
            log_text.see(tk.END)
        else: # For fetcher tool, update the summary label
             log_text.config(text=f"[{timestamp}] {message}")

    def process_queue(self):
        try:
            msg_type, data = self.log_queue.get_nowait()
            tool = data.get('tool', 'comparison')

            if tool == 'comparison':
                if msg_type == "LOG": self.log(data['payload'], tool)
                elif msg_type == "UPDATE_SUMMARY": 
                    self.c_confirmed_var.set(data['payload']['confirmed'])
                    self.c_fuzzy_var.set(data['payload']['fuzzy'])
                    self.c_unmatched1_var.set(data['payload']['unmatched1'])
                    self.c_unmatched2_var.set(data['payload']['unmatched2'])
                elif msg_type == "DONE":
                    self.c_progress_bar.stop()
                    self.c_run_button.config(state="normal")
                    self.log(data['payload']['log'], tool)
                    self.handle_save(data['payload']['df'], tool)
                elif msg_type == "ERROR":
                    self.c_progress_bar.stop()
                    self.c_run_button.config(state="normal")
                    self.log(f"خطأ فادح: {data['payload']}", tool)
                    messagebox.showerror("حدث خطأ", f"فشلت العملية.\n{data['payload']}")
            
            elif tool == 'fetcher':
                if msg_type == "LOG": self.log(data['payload'], tool)
                elif msg_type == "DONE":
                    self.f_progress_bar.stop()
                    self.f_run_button.config(state="normal")
                    self.log(data['payload']['log'], tool)
                    self.handle_save(data['payload']['df'], tool)
                elif msg_type == "ERROR":
                    self.f_progress_bar.stop()
                    self.f_run_button.config(state="normal")
                    self.log(f"خطأ فادح: {data['payload']}", tool)
                    messagebox.showerror("حدث خطأ", f"فشلت العملية.\n{data['payload']}")

        except queue.Empty:
            pass
        finally:
            self.root.after(100, self.process_queue)
    
    def handle_save(self, results_df, tool):
        if results_df.empty:
            messagebox.showinfo("اكتملت العملية", "لا توجد نتائج للحفظ.")
            return

        if not messagebox.askyesno("حفظ النتائج", "اكتمل التحليل. هل ترغب في حفظ ملف النتائج الآن؟"):
            self.log("تم اختيار عدم حفظ النتائج.", tool)
            return
        
        save_option = self.c_save_option.get() if tool == 'comparison' else 'new_file' # Fetcher always saves to new file or same file
        
        try:
            if save_option == "new_file":
                save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], title="حفظ ملف النتائج")
                if not save_path: self.log("تم إلغاء عملية الحفظ.", tool); return
                results_df.to_excel(save_path, index=False)
                self.log(f"تم حفظ النتائج في ملف جديد: {save_path}", tool)
            else: # Save to existing file
                file_path = self.c_file1_path.get() if save_option == "file1" else self.c_file2_path.get()
                output_sheet = f"Comparison_{datetime.now().strftime('%y%m%d_%H%M')}"
                with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    results_df.to_excel(writer, sheet_name=output_sheet, index=False)
                self.log(f"تم حفظ النتائج في صفحة '{output_sheet}'.", tool)
            messagebox.showinfo("اكتمل الحفظ", "تم حفظ النتائج بنجاح.")
        except Exception as e:
            self.log(f"فشل حفظ الملف: {e}", tool)
            messagebox.showerror("خطأ في الحفظ", f"لم يتمكن من حفظ الملف.\n{e}")

    # --------------------------------------------------------------------
    # طرق خاصة بتبويب "مقارنة SKUs"
    # --------------------------------------------------------------------
    def c_use_same_file(self):
        first_file = self.c_file1_path.get()
        if not first_file: messagebox.showwarning("لا يوجد ملف", "الرجاء اختيار الملف الأول أولاً."); return
        self.c_file2_path.set(first_file)
        self.log("تم استخدام نفس الملف الأول للملف الثاني.", 'comparison')
        self.c_update_sheets(self.c_file2_path, self.c_file2_sheet, self.c_file2_col)
        self.c_save_file2_rb.config(state='normal')

    def c_select_file(self, file_var, sheet_var, col_var):
        path = filedialog.askopenfilename(title="اختر ملف إكسل", filetypes=(("Excel Files", "*.xlsx *.xls"),))
        if not path: return
        file_var.set(path)
        self.log(f"تم اختيار الملف: {path.split('/')[-1]}", 'comparison')
        self.c_update_sheets(file_var, sheet_var, col_var)
        if file_var == self.c_file1_path: self.c_save_file1_rb.config(state='normal')
        if file_var == self.c_file2_path: self.c_save_file2_rb.config(state='normal')

    def c_update_sheets(self, file_var, sheet_var, col_var):
        try:
            xls = pd.ExcelFile(file_var.get())
            widgets = self.c_file1_widgets if file_var == self.c_file1_path else self.c_file2_widgets
            widgets['sheet']['values'] = xls.sheet_names
            widgets['sheet'].config(state="readonly")
            sheet_var.set(''); col_var.set(''); widgets['col'].config(state="disabled")
        except Exception as e: messagebox.showerror("خطأ في الملف", f"لا يمكن قراءة الملف.\n{e}")

    def c_on_sheet_select(self, file_var, sheet_var, col_var):
        if not sheet_var.get(): return
        try:
            df = pd.read_excel(file_var.get(), sheet_name=sheet_var.get(), nrows=1)
            widgets = self.c_file1_widgets if file_var == self.c_file1_path else self.c_file2_widgets
            widgets['col']['values'] = list(df.columns)
            widgets['col'].config(state="readonly")
        except Exception as e: messagebox.showerror("خطأ في الصفحة", f"لا يمكن قراءة الأعمدة.\n{e}")

    def start_comparison_thread(self):
        params = {"file1": self.c_file1_path.get(), "sheet1": self.c_file1_sheet.get(), "col1": self.c_file1_col.get(),
                  "file2": self.c_file2_path.get(), "sheet2": self.c_file2_sheet.get(), "col2": self.c_file2_col.get(),
                  "threshold": self.c_similarity_threshold.get()}
        if not all(p for p in params.values() if isinstance(p, str)):
            messagebox.showwarning("مدخلات ناقصة", "الرجاء ملء جميع الحقول لكلا الملفين."); return
        self.c_confirmed_var.set("0"); self.c_fuzzy_var.set("0"); self.c_unmatched1_var.set("0"); self.c_unmatched2_var.set("0")
        self.c_run_button.config(state="disabled")
        self.log("بدء عملية المقارنة...", 'comparison')
        self.c_progress_bar.start()
        threading.Thread(target=run_comparison_logic, args=(params, self.log_queue), daemon=True).start()

    # --------------------------------------------------------------------
    # طرق خاصة بتبويب "سحب البيانات"
    # --------------------------------------------------------------------
    def f_select_file(self):
        path = filedialog.askopenfilename(title="اختر ملف إكسل", filetypes=(("Excel Files", "*.xlsx *.xls"),))
        if not path: return
        self.f_file_path.set(path)
        self.log(f"تم اختيار الملف: {path.split('/')[-1]}", 'fetcher')
        try:
            xls = pd.ExcelFile(path)
            sheet_names = xls.sheet_names
            self.f_source_sheet_combo['values'] = sheet_names
            self.f_search_sheet_combo['values'] = sheet_names
            self.f_source_sheet_combo.config(state='readonly')
            self.f_search_sheet_combo.config(state='readonly')
        except Exception as e: messagebox.showerror("خطأ في الملف", f"لا يمكن قراءة الملف.\n{e}")

    def f_on_sheet_select(self, type):
        sheet_var = self.f_source_sheet.get() if type == 'source' else self.f_search_sheet.get()
        col_combo = self.f_source_col_combo if type == 'source' else self.f_search_col_combo
        if not sheet_var: return
        try:
            df = pd.read_excel(self.f_file_path.get(), sheet_name=sheet_var, nrows=1)
            col_combo['values'] = list(df.columns)
            col_combo.config(state='readonly')
        except Exception as e: messagebox.showerror("خطأ في الصفحة", f"لا يمكن قراءة الأعمدة.\n{e}")

    def start_fetcher_thread(self):
        params = {"file_path": self.f_file_path.get(), "source_sheet": self.f_source_sheet.get(), "source_col": self.f_source_col.get(),
                  "search_sheet": self.f_search_sheet.get(), "search_col": self.f_search_col.get()}
        if not all(params.values()):
            messagebox.showwarning("مدخلات ناقصة", "الرجاء ملء جميع الحقول."); return
        self.f_run_button.config(state="disabled")
        self.log("بدء عملية سحب البيانات...", 'fetcher')
        self.f_progress_bar.start()
        threading.Thread(target=run_fetching_logic, args=(params, self.log_queue), daemon=True).start()


# -----------------------------------------------------------------------------
# منطق المعالجة في الخلفية
# -----------------------------------------------------------------------------
def run_comparison_logic(params, log_queue):
    def clean_sku(sku): return re.sub(r'[^a-zA-Z0-9]', '', str(sku)).lower() if sku else ""
    try:
        df1 = pd.read_excel(params['file1'], sheet_name=params['sheet1'])
        df2 = pd.read_excel(params['file2'], sheet_name=params['sheet2'])
        original_set1 = set(df1[params['col1']].dropna().astype(str))
        original_set2 = set(df2[params['col2']].dropna().astype(str))
        log_queue.put(("LOG", {'tool': 'comparison', 'payload': f"ملف 1: {len(original_set1)} SKU | ملف 2: {len(original_set2)} SKU"}))
        
        final_results_list, summary = [], {'confirmed': 0, 'fuzzy': 0}
        log_queue.put(("LOG", {'tool': 'comparison', 'payload': "المرحلة 1: البحث عن التطابقات المؤكدة..."}))
        cleaned_map2 = {clean_sku(s): s for s in original_set2}
        set1 = set(original_set1)
        items_to_remove1, items_to_remove2 = set(), set()
        for item1 in set1:
            cleaned_item1 = clean_sku(item1)
            if cleaned_item1 and cleaned_item1 in cleaned_map2:
                item2 = cleaned_map2[cleaned_item1]
                if item2 not in items_to_remove2:
                    final_results_list.append((item1, item2, 'تطابق مؤكد'))
                    items_to_remove1.add(item1); items_to_remove2.add(item2)
                    summary['confirmed'] += 1; del cleaned_map2[cleaned_item1]
        set1 -= items_to_remove1; set2 = set(original_set2) - items_to_remove2

        if set1 and set2:
            log_queue.put(("LOG", {'tool': 'comparison', 'payload': f"المرحلة 2: البحث التقريبي الشامل (> {params['threshold']}%) ..."}))
            fuzzy_pairs = []
            for item1 in set1:
                for item2 in set2:
                    score = fuzz.ratio(str(item1), str(item2))
                    if score >= params['threshold']: fuzzy_pairs.append((score, item1, item2))
            fuzzy_pairs.sort(key=lambda x: x[0], reverse=True)
            items_to_remove1_fuzzy, items_to_remove2_fuzzy = set(), set()
            for score, item1, item2 in fuzzy_pairs:
                if item1 not in items_to_remove1_fuzzy and item2 not in items_to_remove2_fuzzy:
                    final_results_list.append((item1, item2, f'تقريبي ({score}%)'))
                    items_to_remove1_fuzzy.add(item1); items_to_remove2_fuzzy.add(item2)
                    summary['fuzzy'] += 1
            set1 -= items_to_remove1_fuzzy; set2 -= items_to_remove2_fuzzy

        log_queue.put(("LOG", {'tool': 'comparison', 'payload': "المرحلة 3: تحديد العناصر غير المتطابقة..."}))
        for item1 in set1: final_results_list.append((item1, '', "موجود فقط في ملف 1"))
        for item2 in set2: final_results_list.append(('', item2, "موجود فقط في ملف 2"))
        
        summary.update({'unmatched1': len(set1), 'unmatched2': len(set2)})
        log_queue.put(("UPDATE_SUMMARY", {'tool': 'comparison', 'payload': summary}))
        
        results_df = pd.DataFrame(final_results_list, columns=[params['sheet1'], params['sheet2'], 'الحالة'])
        log_queue.put(("DONE", {'tool': 'comparison', 'payload': {'df': results_df, 'log': "اكتمل التحليل بنجاح."}}))
    except Exception as e:
        log_queue.put(("ERROR", {'tool': 'comparison', 'payload': f"{type(e).__name__}: {e}"}))

def run_fetching_logic(params, log_queue):
    try:
        log_queue.put(("LOG", {'tool': 'fetcher', 'payload': f"قراءة البيانات من الملف..."}))
        df_source = pd.read_excel(params['file_path'], sheet_name=params['source_sheet'])
        df_search = pd.read_excel(params['file_path'], sheet_name=params['search_sheet'])
        
        search_list = df_search[params['search_col']].dropna().unique().tolist()
        log_queue.put(("LOG", {'tool': 'fetcher', 'payload': f"تم العثور على {len(search_list)} عنصر للبحث عنه."}))
        
        filtered_data = df_source[df_source[params['source_col']].isin(search_list)]
        log_queue.put(("LOG", {'tool': 'fetcher', 'payload': f"تم العثور على {len(filtered_data)} صف مطابق."}))
        
        if filtered_data.empty:
            log_queue.put(("DONE", {'tool': 'fetcher', 'payload': {'df': pd.DataFrame(), 'log': "لم يتم العثور على أي منتجات مطابقة."}}))
            return

        output_sheet_name = f"Matched_{params['search_sheet']}"
        log_queue.put(("DONE", {'tool': 'fetcher', 'payload': {'df': filtered_data, 'log': f"اكتمل السحب! النتائج جاهزة للحفظ."}}))

    except Exception as e:
        log_queue.put(("ERROR", {'tool': 'fetcher', 'payload': f"{type(e).__name__}: {e}"}))

# -----------------------------------------------------------------------------
# نقطة انطلاق البرنامج
# -----------------------------------------------------------------------------
if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = SkuMultiToolApp(root)
        root.mainloop()
    except ImportError:
        messagebox.showerror("مكتبة ناقصة", "الرجاء تثبيت النمط الحديث أولاً:\npip install sv-ttk")
