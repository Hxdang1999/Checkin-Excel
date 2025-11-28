import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import traceback

class TimeAttendanceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Tool Lọc Công - v13.0 (Tự động xóa dòng thiếu Tên/ID)")
        self.root.geometry("1300x768")
        
        # Biến dữ liệu
        self.raw_df = None
        self.filtered_df = None
        self.cols_to_drop = []
        self.col_time_original = None
        self.col_last_name = None
        self.col_first_name = None

        # ================= HEADER =================
        top_frame = tk.Frame(root, bg="#eceff1", pady=15)
        top_frame.pack(side=tk.TOP, fill=tk.X)
        btn_style = {"font": ("Segoe UI", 10, "bold"), "width": 18, "height": 2, "relief": tk.RAISED, "bd": 2}

        # 1. IMPORT
        self.btn_import = tk.Button(top_frame, text="📂 1. Chọn File", command=self.import_file, bg="#03A9F4", fg="white", **btn_style)
        self.btn_import.pack(side=tk.LEFT, padx=15)

        # NÚT CHỌN CỘT
        self.btn_select_cols = tk.Button(top_frame, text="🗑 Tùy chọn cột...", command=self.open_column_selector, state=tk.DISABLED, bg="#CFD8DC", fg="#546E7A", font=("Segoe UI", 9))
        self.btn_select_cols.pack(side=tk.LEFT, padx=5)

        # 2. XỬ LÝ
        self.btn_filter = tk.Button(top_frame, text="⚡ 2. Lọc & Xuất", command=self.filter_data, bg="#FF9800", fg="white", **btn_style)
        self.btn_filter.pack(side=tk.LEFT, padx=15)

        # 3. XUẤT
        self.btn_export = tk.Button(top_frame, text="💾 3. Lưu Excel", command=self.export_file, bg="#4CAF50", fg="white", **btn_style)
        self.btn_export.pack(side=tk.LEFT, padx=15)

        # ================= MAIN VIEW =================
        self.paned_window = tk.PanedWindow(root, orient=tk.HORIZONTAL, sashwidth=8, bg="#90a4ae")
        self.paned_window.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Khung Trái
        frame_left = tk.LabelFrame(self.paned_window, text=" Dữ liệu gốc ", font=("Segoe UI", 10, "bold"))
        self.paned_window.add(frame_left, width=600, stretch="always")
        self.tree_left = self.create_treeview(frame_left)

        # Khung Phải
        frame_right = tk.LabelFrame(self.paned_window, text=" Kết quả ", font=("Segoe UI", 10, "bold"), fg="#D84315")
        self.paned_window.add(frame_right, width=600, stretch="always")
        self.tree_right = self.create_treeview(frame_right)

        self.status_lbl = tk.Label(root, text="Sẵn sàng", bd=1, relief=tk.SUNKEN, anchor=tk.W, bg="#cfd8dc")
        self.status_lbl.pack(side=tk.BOTTOM, fill=tk.X)

    def create_treeview(self, parent):
        sc_y = tk.Scrollbar(parent); sc_y.pack(side=tk.RIGHT, fill=tk.Y)
        sc_x = tk.Scrollbar(parent, orient='horizontal'); sc_x.pack(side=tk.BOTTOM, fill=tk.X)
        tree = ttk.Treeview(parent, yscrollcommand=sc_y.set, xscrollcommand=sc_x.set)
        tree.pack(fill=tk.BOTH, expand=True)
        sc_y.config(command=tree.yview); sc_x.config(command=tree.xview)
        return tree

    # ================= 1. IMPORT =================
    def import_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Data Files", "*.csv *.xlsx *.xls")])
        if not file_path: return

        try:
            self.status_lbl.config(text="Đang đọc file...")
            self.root.update()

            df = None
            if file_path.endswith('.csv'):
                try: df = pd.read_csv(file_path, encoding='utf-8', on_bad_lines='skip')
                except: df = pd.read_csv(file_path, encoding='utf-16', on_bad_lines='skip', sep='\t')
            else:
                df = pd.read_excel(file_path)

            header_idx = -1
            if df is not None:
                for i in range(min(10, len(df))):
                    row_str = str(df.iloc[i].values).lower()
                    if any(x in row_str for x in ['time', 'personnel id', 'mã nhân viên', 'giờ']):
                        header_idx = i + 1; break
            
            if header_idx > 0:
                if file_path.endswith('.csv'):
                    try: df = pd.read_csv(file_path, skiprows=header_idx, encoding='utf-8', on_bad_lines='skip')
                    except: df = pd.read_csv(file_path, skiprows=header_idx, encoding='utf-16', on_bad_lines='skip', sep='\t')
                else:
                    df = pd.read_excel(file_path, skiprows=header_idx)

            if df is None or df.empty: raise ValueError("File rỗng")

            df.columns = df.columns.str.strip()
            
            # --- TÌM CỘT THÔNG MINH ---
            self.col_time_original = next((c for c in df.columns if any(x in str(c).lower() for x in ['time', 'giờ', 'thời gian'])), None)
            self.col_last_name = next((c for c in df.columns if any(x in str(c).lower() for x in ['last name', 'họ', 'họ đệm'])), None)
            
            possible_names = [c for c in df.columns if any(x in str(c).lower() for x in ['first name', 'tên', 'name'])]
            self.col_first_name = None
            for c in possible_names:
                c_low = str(c).lower()
                if "device" in c_low or "area" in c_low or "department" in c_low: continue
                self.col_first_name = c; break

            self.raw_df = df
            
            # Mặc định xóa
            self.cols_to_drop = ["Ngay", "Gio", "Name"]
            
            self.btn_select_cols.config(state=tk.NORMAL, bg="#FFCCBC", fg="red") 
            self.display_data(self.tree_left, self.raw_df)
            self.clear_tree(self.tree_right)
            self.status_lbl.config(text=f"Đã tải {len(self.raw_df)} dòng.")
            
            messagebox.showinfo("Thông báo", "Tải file lên thành công")
            
        except Exception as e:
            messagebox.showerror("Lỗi Import", f"Chi tiết lỗi:\n{e}")

    # ================= 2. CHỌN CỘT =================
    def open_column_selector(self):
        if self.raw_df is None: return

        win = tk.Toplevel(self.root); win.title("Chọn cột cần XÓA")
        win.geometry("500x600")
        
        lbl = tk.Label(win, text="Tích vào cột bạn muốn XÓA (Ẩn đi).\n(Mặc định Ngay, Gio, Name đã được tích)", 
                       fg="red", font=("Arial", 10, "bold"), pady=10)
        lbl.pack()

        frame_chk = tk.Frame(win); frame_chk.pack(fill=tk.BOTH, expand=True, padx=10)
        canvas = tk.Canvas(frame_chk); scrollbar = tk.Scrollbar(frame_chk, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)
        
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True); scrollbar.pack(side="right", fill="y")

        display_list = []
        original_cols = list(self.raw_df.columns)
        
        for col in original_cols:
            display_list.append({"name": col, "is_new": False})
            if self.col_time_original and col == self.col_time_original:
                display_list.append({"name": "Ngay", "is_new": True})
                display_list.append({"name": "Gio", "is_new": True})
            if self.col_last_name and col == self.col_last_name:
                 display_list.append({"name": "Name", "is_new": True})

        self.chk_vars = {}
        for item in display_list:
            col_name = item["name"]
            is_new = item["is_new"]
            
            default_val = False 
            if col_name in self.cols_to_drop: default_val = True

            var = tk.BooleanVar(value=default_val)
            
            display_text = col_name
            if col_name == self.col_time_original: display_text += " (Gốc)"
            
            fg_color = "black"
            font_style = ("Arial", 9)
            if is_new:
                fg_color = "#009688"
                font_style = ("Arial", 9, "bold")
                display_text += " (Mới)"

            chk = tk.Checkbutton(scrollable_frame, text=display_text, variable=var, anchor='w', fg=fg_color, font=font_style)
            chk.pack(fill='x', padx=5, pady=2)
            self.chk_vars[col_name] = var

        btn_confirm = tk.Button(win, text="Xác Nhận", command=lambda: self.confirm_cols_drop(win), bg="#4CAF50", fg="white", height=2)
        btn_confirm.pack(pady=10, fill=tk.X, padx=20)

    def confirm_cols_drop(self, window):
        self.cols_to_drop = [col for col, var in self.chk_vars.items() if var.get()]
        window.destroy()
        self.btn_select_cols.config(text=f"🗑 Sẽ xóa {len(self.cols_to_drop)} cột", bg="#FFAB91")

    # ================= 3. LỌC & TẠO CỘT =================
    def filter_data(self):
        if self.raw_df is None: return
        try:
            df = self.raw_df.copy()
            col_id = next((c for c in df.columns if any(x in str(c).lower() for x in ['id', 'mã', 'personnel'])), None)
            
            if not col_id or not self.col_time_original:
                messagebox.showerror("Lỗi", "Không xác định được ID hoặc Cột thời gian gốc."); return

            # --- TÍNH NĂNG MỚI: XÓA DÒNG THIẾU THÔNG TIN ---
            # Xác định các cột bắt buộc phải có dữ liệu
            required_cols = [col_id]
            if self.col_last_name: required_cols.append(self.col_last_name)
            if self.col_first_name: required_cols.append(self.col_first_name)
            
            # Đếm trước khi xóa
            before_count = len(df)
            
            # Thực hiện xóa các dòng mà ID, Họ hoặc Tên bị NaN (trống)
            df.dropna(subset=required_cols, inplace=True)
            
            # Đếm sau khi xóa
            dropped_count = before_count - len(df)
            # -----------------------------------------------

            # 1. Xử lý Thời gian
            df[self.col_time_original] = pd.to_datetime(df[self.col_time_original], dayfirst=True, errors='coerce')
            df.dropna(subset=[self.col_time_original], inplace=True)
            
            idx_time = df.columns.get_loc(self.col_time_original)
            df.insert(idx_time + 1, 'Ngay', df[self.col_time_original].dt.strftime('%d/%m/%Y'))
            df.insert(idx_time + 2, 'Gio', df[self.col_time_original].dt.strftime('%H:%M:%S'))

            # 2. Xử lý Tên (Tạo cột Name)
            if self.col_last_name and self.col_first_name:
                if self.col_last_name in df.columns and self.col_first_name in df.columns:
                    full_name = df[self.col_last_name].fillna('').astype(str) + " " + df[self.col_first_name].fillna('').astype(str)
                    full_name = full_name.str.strip().str.replace(r'\s+', ' ', regex=True)
                    
                    idx_name = df.columns.get_loc(self.col_last_name)
                    df.insert(idx_name + 1, 'Name', full_name)

            # 3. Lọc trùng
            df.sort_values(by=[col_id, self.col_time_original], inplace=True)
            df = df.drop_duplicates(subset=[col_id, 'Ngay'], keep='first')

            # 4. Xóa cột
            real_cols_to_drop = [c for c in self.cols_to_drop if c != col_id]
            if real_cols_to_drop:
                existing = [c for c in real_cols_to_drop if c in df.columns]
                df.drop(columns=existing, inplace=True, errors='ignore')

            self.filtered_df = df
            self.display_data(self.tree_right, self.filtered_df)
            
            self.status_lbl.config(text=f"Hoàn tất! Kết quả: {len(df)} dòng.")
            
            # Thông báo kết quả + số dòng bị xóa (nếu có)
            msg = "Lọc thành công"
            if dropped_count > 0:
                msg += f"\n\n(Lưu ý: Đã tự động loại bỏ {dropped_count} dòng do thiếu ID hoặc Tên)"
                
            messagebox.showinfo("Thông báo", msg)
            
        except Exception as e:
            messagebox.showerror("Lỗi Lọc", f"{str(e)}\n{traceback.format_exc()}")

    # ================= 4. XUẤT =================
    def export_file(self):
        if self.filtered_df is None: return
        f = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if not f: return
        try:
            with pd.ExcelWriter(f, engine='xlsxwriter') as writer:
                self.filtered_df.to_excel(writer, index=False, sheet_name='CheckinList')
                ws = writer.sheets['CheckinList']
                for i, col in enumerate(self.filtered_df.columns): ws.set_column(i, i, 15)
            os.startfile(f)
        except Exception as e: messagebox.showerror("Lỗi lưu file", str(e))

    def display_data(self, tree, df):
        self.clear_tree(tree)
        if df is None: return
        tree["columns"] = list(df.columns); tree.column("#0", width=0, stretch=tk.NO)
        for c in df.columns: tree.heading(c, text=c); tree.column(c, width=100)
        for _, row in df.head(1000).astype(str).iterrows(): tree.insert("", tk.END, values=list(row))

    def clear_tree(self, tree): tree.delete(*tree.get_children())

if __name__ == "__main__":
    root = tk.Tk()
    app = TimeAttendanceApp(root)
    root.mainloop()