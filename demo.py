import sys
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from fpdf import FPDF
import os
import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import numpy as np
import traceback

class CustomerAppAnalyzer:
    def __init__(self, root):
        self.root = root
        self.root.title("Phân tích khách hàng cài app theo khu vực")
        self.root.geometry("900x600")

        # Biến lưu trữ dữ liệu
        self.customer_data = None
        self.app_data = None
        self.result_data = None
        self.detailed_data = None
        self.not_installed_data = None  # Biến lưu danh sách khách hàng chưa cài app

        # Tạo giao diện
        self.create_widgets()

    def create_widgets(self):
        # Frame chọn file
        file_frame = ttk.LabelFrame(self.root, text="Chọn File Dữ Liệu")
        file_frame.pack(fill="x", padx=10, pady=10)

        # Chọn file khách hàng
        ttk.Label(file_frame, text="File danh sách khách hàng (Excel):").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.customer_file_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.customer_file_var, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="Chọn File...", command=self.select_customer_file).grid(row=0, column=2, padx=5, pady=5)

        # Chọn file cài app
        ttk.Label(file_frame, text="File danh sách khách hàng đã cài app (CSV):").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.app_file_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.app_file_var, width=50).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="Chọn File...", command=self.select_app_file).grid(row=1, column=2, padx=5, pady=5)

        # Hiển thị thông tin
        info_frame = ttk.Frame(self.root)
        info_frame.pack(fill="x", padx=10)

        self.customer_count_var = tk.StringVar(value="Tổng số khách hàng: 0")
        self.app_count_var = tk.StringVar(value="Số khách hàng đã cài app: 0")

        ttk.Label(info_frame, textvariable=self.customer_count_var).pack(side="left", padx=10)
        ttk.Label(info_frame, textvariable=self.app_count_var).pack(side="left", padx=10)

        # Nút phân tích và danh sách
        button_frame = ttk.Frame(self.root)
        button_frame.pack(fill="x", padx=10, pady=5)

        ttk.Button(button_frame, text="Phân tích dữ liệu", command=self.analyze_data).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Danh sách khách hàng theo khu vực", command=self.show_detailed_list).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Danh sách khách hàng chưa cài app", command=self.show_not_installed_list).pack(side="left", padx=5)

        # Nút xuất báo cáo (tổng thể)
        self.export_excel_button = ttk.Button(button_frame, text="Xuất Excel (Tổng thể)", command=self.export_excel, state="disabled")
        self.export_excel_button.pack(side="right", padx=5)
        self.export_pdf_button = ttk.Button(button_frame, text="Xuất PDF (Tổng thể)", command=self.export_pdf, state="disabled")
        self.export_pdf_button.pack(side="right", padx=5)

        # Nút xuất báo cáo (chi tiết)
        self.export_detailed_excel_button = ttk.Button(button_frame, text="Xuất Excel (Chi tiết)", command=self.export_detailed_excel, state="disabled")
        self.export_detailed_excel_button.pack(side="right", padx=5)
        self.export_detailed_pdf_button = ttk.Button(button_frame, text="Xuất PDF (Chi tiết)", command=self.export_detailed_pdf, state="disabled")
        self.export_detailed_pdf_button.pack(side="right", padx=5)

        # Nút xuất báo cáo (chưa cài app)
        self.export_not_installed_excel_button = ttk.Button(button_frame, text="Xuất Excel (Chưa cài app)", command=self.export_not_installed_excel, state="disabled")
        self.export_not_installed_excel_button.pack(side="right", padx=5)
        self.export_not_installed_pdf_button = ttk.Button(button_frame, text="Xuất PDF (Chưa cài app)", command=self.export_not_installed_pdf, state="disabled")
        self.export_not_installed_pdf_button.pack(side="right", padx=5)

        # Notebook cho kết quả
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=5)

        # Tab bảng dữ liệu (summary)
        self.table_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.table_frame, text="Bảng dữ liệu")

        columns = ("region", "total", "app_installed", "percentage")
        self.tree = ttk.Treeview(self.table_frame, columns=columns, show="headings")
        self.tree.heading("region", text="Mã khu vực")
        self.tree.heading("total", text="Số khách hàng")
        self.tree.heading("app_installed", text="Số khách hàng đã cài app")
        self.tree.heading("percentage", text="Tỷ lệ (%)")
        self.tree.column("region", width=100, anchor="center")
        self.tree.column("total", width=150, anchor="center")
        self.tree.column("app_installed", width=150, anchor="center")
        self.tree.column("percentage", width=100, anchor="center")
        scrollbar = ttk.Scrollbar(self.table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Tab danh sách chi tiết
        self.detailed_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.detailed_frame, text="Danh sách chi tiết")

        detailed_columns = ("id", "name", "region", "street")
        self.detailed_tree = ttk.Treeview(self.detailed_frame, columns=detailed_columns, show="headings")
        self.detailed_tree.heading("id", text="Mã danh bộ")
        self.detailed_tree.heading("name", text="Tên khách hàng")
        self.detailed_tree.heading("region", text="Mã khu vực")
        self.detailed_tree.heading("street", text="Tên đường")
        self.detailed_tree.column("id", width=150, anchor="center")
        self.detailed_tree.column("name", width=200, anchor="w")
        self.detailed_tree.column("region", width=100, anchor="center")
        self.detailed_tree.column("street", width=300, anchor="w")
        detailed_scrollbar = ttk.Scrollbar(self.detailed_frame, orient="vertical", command=self.detailed_tree.yview)
        self.detailed_tree.configure(yscrollcommand=detailed_scrollbar.set)
        self.detailed_tree.pack(side="left", fill="both", expand=True)
        detailed_scrollbar.pack(side="right", fill="y")

        # Tab danh sách chưa cài app
        self.not_installed_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.not_installed_frame, text="Danh sách chưa cài app")

        not_installed_columns = ("id", "name", "region", "street")
        self.not_installed_tree = ttk.Treeview(self.not_installed_frame, columns=not_installed_columns, show="headings")
        self.not_installed_tree.heading("id", text="Mã danh bộ")
        self.not_installed_tree.heading("name", text="Tên khách hàng")
        self.not_installed_tree.heading("region", text="Mã khu vực")
        self.not_installed_tree.heading("street", text="Tên đường")
        self.not_installed_tree.column("id", width=150, anchor="center")
        self.not_installed_tree.column("name", width=200, anchor="w")
        self.not_installed_tree.column("region", width=100, anchor="center")
        self.not_installed_tree.column("street", width=300, anchor="w")
        not_installed_scrollbar = ttk.Scrollbar(self.not_installed_frame, orient="vertical", command=self.not_installed_tree.yview)
        self.not_installed_tree.configure(yscrollcommand=not_installed_scrollbar.set)
        self.not_installed_tree.pack(side="left", fill="both", expand=True)
        not_installed_scrollbar.pack(side="right", fill="y")

        # Tab biểu đồ
        self.chart_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.chart_frame, text="Biểu đồ")

        # Thông tin tổng hợp
        self.summary_var = tk.StringVar()
        self.summary_label = ttk.Label(self.root, textvariable=self.summary_var, font=("Arial", 10, "bold"))
        self.summary_label.pack(fill="x", padx=10, pady=5)

    def select_customer_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.customer_file_var.set(file_path)
            self.load_customer_data(file_path)

    def select_app_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if file_path:
            self.app_file_var.set(file_path)
            self.load_app_data(file_path)

    def load_customer_data(self, file_path):
        try:
            xls = pd.ExcelFile(file_path)
            sheet_name = xls.sheet_names[0]
            df_check = pd.read_excel(xls, sheet_name=sheet_name, nrows=0)
            if 'RegionCode' not in df_check.columns:
                messagebox.showerror("Lỗi Cột", f"File khách hàng '{os.path.basename(file_path)}' phải chứa cột 'RegionCode'.")
                self.customer_file_var.set("")
                return
            if 'CustomerRegistationNo' not in df_check.columns:
                messagebox.showerror("Lỗi Cột", f"File khách hàng '{os.path.basename(file_path)}' phải chứa cột 'CustomerRegistationNo'.")
                self.customer_file_var.set("")
                return
            self.customer_data = pd.read_excel(file_path)
            self.customer_data = self.customer_data.drop_duplicates(subset=['CustomerRegistationNo'], keep='first')
            print("Customer data columns:", self.customer_data.columns.tolist())
            print("Customer data shape after removing duplicates:", self.customer_data.shape)
            self.customer_count_var.set(f"Tổng số khách hàng: {len(self.customer_data)}")
            print("Đã đọc dữ liệu khách hàng:")
            print(self.customer_data.head())
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể đọc file khách hàng: {str(e)}")
            traceback.print_exc()

    def load_app_data(self, file_path):
        try:
            df_check = pd.read_csv(file_path, nrows=0)
            if 'CustomerRegistrationNo' not in df_check.columns:
                messagebox.showerror("Lỗi Cột", f"File cài app '{os.path.basename(file_path)}' phải chứa cột 'CustomerRegistrationNo'.")
                self.app_file_var.set("")
                return
            self.app_data = pd.read_csv(file_path)
            self.app_data = self.app_data.drop_duplicates(subset=['CustomerRegistrationNo'], keep='first')
            print("App data columns:", self.app_data.columns.tolist())
            print("App data shape after removing duplicates:", self.app_data.shape)
            self.app_count_var.set(f"Số khách hàng đã cài app: {len(self.app_data)}")
            print("Đã đọc dữ liệu cài app:")
            print(self.app_data.head())
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể đọc file cài app: {str(e)}")
            traceback.print_exc()

    def analyze_data(self):
        if self.customer_data is None or self.app_data is None:
            messagebox.showwarning("Cảnh báo", "Vui lòng tải cả hai file dữ liệu trước khi phân tích.")
            return

        required_customer_cols = ['CustomerRegistationNo', 'RegionCode', 'CustomerName', 'StreetCode', 'Name']
        required_app_cols = ['CustomerRegistrationNo']

        if not all(col in self.customer_data.columns for col in required_customer_cols):
            missing_cols = [col for col in required_customer_cols if col not in self.customer_data.columns]
            messagebox.showerror("Lỗi Cột", f"File khách hàng thiếu các cột: {', '.join(missing_cols)}.")
            return
        if not all(col in self.app_data.columns for col in required_app_cols):
            missing_cols = [col for col in required_app_cols if col not in self.app_data.columns]
            messagebox.showerror("Lỗi Cột", f"File cài app thiếu các cột: {', '.join(missing_cols)}.")
            return

        try:
            customer_df = self.customer_data.copy()
            app_df = self.app_data.copy()
            customer_df['CustomerRegistationNo'] = customer_df['CustomerRegistationNo'].astype(str).str.replace('\.0$', '', regex=True).str.strip()
            app_df['CustomerRegistrationNo'] = app_df['CustomerRegistrationNo'].astype(str).str.strip()
            app_customer_ids = set(app_df['CustomerRegistrationNo'])
            customer_df['has_installed_app'] = customer_df['CustomerRegistationNo'].isin(app_customer_ids)

            region_analysis = customer_df.groupby('RegionCode').agg(
                total_customers=('CustomerRegistationNo', 'nunique'),
                app_installed=('has_installed_app', 'sum')
            ).reset_index()
            region_analysis['percentage'] = (region_analysis['app_installed'] / region_analysis['total_customers'] * 100).fillna(0).round(2)
            region_analysis['percentage'] = region_analysis['percentage'].clip(upper=100)
            self.result_data = region_analysis

            self.detailed_data = customer_df[customer_df['has_installed_app']][
                ['CustomerRegistationNo', 'CustomerName', 'RegionCode', 'StreetCode', 'Name']
            ].copy()
            print("Detailed data (customers who installed app):")
            print(self.detailed_data.head())
            print("Number of detailed records:", len(self.detailed_data))

            self.not_installed_data = customer_df[~customer_df['has_installed_app']][
                ['CustomerRegistationNo', 'CustomerName', 'RegionCode', 'StreetCode', 'Name']
            ].copy()
            print("Not installed data (customers who have not installed app):")
            print(self.not_installed_data.head())
            print("Number of not installed records:", len(self.not_installed_data))

            self.tree.delete(*self.tree.get_children())
            for i, row in self.result_data.iterrows():
                self.tree.insert("", "end", values=(
                    row['RegionCode'],
                    f"{row['total_customers']:,}",
                    f"{row['app_installed']:,}",
                    f"{row['percentage']:.2f}%"
                ))

            total_customers_overall = customer_df['CustomerRegistationNo'].nunique()
            total_app_installed_overall = self.detailed_data['CustomerRegistationNo'].nunique()
            total_percentage_overall = (total_app_installed_overall / total_customers_overall * 100) if total_customers_overall > 0 else 0

            self.summary_var.set(f"Tổng hợp: {total_customers_overall:,} khách hàng | {total_app_installed_overall:,} đã cài app | Tỷ lệ: {total_percentage_overall:.2f}%")
            self.customer_count_var.set(f"Tổng số khách hàng: {total_customers_overall:,}")
            self.app_count_var.set(f"Số khách hàng đã cài app: {total_app_installed_overall:,}")

            self.create_chart()

            self.export_excel_button['state'] = 'normal'
            self.export_pdf_button['state'] = 'normal'
            self.export_detailed_excel_button['state'] = 'disabled'
            self.export_detailed_pdf_button['state'] = 'disabled'
            self.export_not_installed_excel_button['state'] = 'disabled'
            self.export_not_installed_pdf_button['state'] = 'disabled'

            self.notebook.select(self.table_frame)

        except KeyError as e:
            messagebox.showerror("Lỗi Cột", f"Không tìm thấy cột cần thiết: {str(e)}.")
            traceback.print_exc()
        except Exception as e:
            messagebox.showerror("Lỗi", f"Lỗi khi phân tích dữ liệu: {str(e)}")
            traceback.print_exc()

    def show_detailed_list(self):
        if self.detailed_data is None:
            messagebox.showwarning("Cảnh báo", "Vui lòng phân tích dữ liệu trước khi xem danh sách chi tiết.")
            return
        try:
            self.detailed_tree.delete(*self.detailed_tree.get_children())
            sorted_detailed_data = self.detailed_data.sort_values(['RegionCode', 'StreetCode']).copy()
            for col in ['CustomerRegistationNo', 'CustomerName', 'RegionCode', 'Name']:
                sorted_detailed_data[col] = sorted_detailed_data[col].fillna('').astype(str).str.replace('\n|\r', '', regex=True)
            for i, row in sorted_detailed_data.iterrows():
                self.detailed_tree.insert("", "end", values=(
                    row['CustomerRegistationNo'],
                    row['CustomerName'],
                    row['RegionCode'],
                    row['Name']
                ))
            self.export_detailed_excel_button['state'] = 'normal'
            self.export_detailed_pdf_button['state'] = 'normal'
            self.notebook.select(self.detailed_frame)
            messagebox.showinfo("Thông báo", f"Đã hiển thị {len(sorted_detailed_data)} khách hàng đã cài app.")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Lỗi khi hiển thị danh sách chi tiết: {str(e)}")
            traceback.print_exc()

    def show_not_installed_list(self):
        if self.not_installed_data is None:
            messagebox.showwarning("Cảnh báo", "Vui lòng phân tích dữ liệu trước khi xem danh sách chưa cài app.")
            return
        try:
            self.not_installed_tree.delete(*self.not_installed_tree.get_children())
            sorted_not_installed_data = self.not_installed_data.sort_values(['RegionCode', 'StreetCode']).copy()
            for col in ['CustomerRegistationNo', 'CustomerName', 'RegionCode', 'Name']:
                sorted_not_installed_data[col] = sorted_not_installed_data[col].fillna('').astype(str).str.replace('\n|\r', '', regex=True)
            for i, row in sorted_not_installed_data.iterrows():
                self.not_installed_tree.insert("", "end", values=(
                    row['CustomerRegistationNo'],
                    row['CustomerName'],
                    row['RegionCode'],
                    row['Name']
                ))
            self.export_not_installed_excel_button['state'] = 'normal'
            self.export_not_installed_pdf_button['state'] = 'normal'
            self.notebook.select(self.not_installed_frame)
            messagebox.showinfo("Thông báo", f"Đã hiển thị {len(sorted_not_installed_data)} khách hàng chưa cài app.")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Lỗi khi hiển thị danh sách chưa cài app: {str(e)}")
            traceback.print_exc()

    def create_chart(self):
        if self.result_data is None or self.result_data.empty:
            return
        for widget in self.chart_frame.winfo_children():
            widget.destroy()
        try:
            chart_data = self.result_data.sort_values('RegionCode').copy()
            regions = chart_data['RegionCode'].astype(str)
            total = chart_data['total_customers']
            installed = chart_data['app_installed']
            fig, ax = plt.subplots(figsize=(10, 5))
            x = np.arange(len(regions))
            width = 0.35
            rects1 = ax.bar(x - width/2, total, width, label='Tổng số khách hàng')
            rects2 = ax.bar(x + width/2, installed, width, label='Đã cài app')
            ax.set_title('Thống kê khách hàng đã cài app theo khu vực', fontsize=14)
            ax.set_xlabel('Mã khu vực', fontsize=10)
            ax.set_ylabel('Số lượng', fontsize=10)
            ax.set_xticks(x)
            ax.set_xticklabels(regions, rotation=45, ha="right")
            ax.legend()
            ax.tick_params(axis='x', labelsize=8)
            ax.tick_params(axis='y', labelsize=8)
            def autolabel(rects):
                for rect in rects:
                    height = rect.get_height()
                    if height > 0:
                        ax.annotate(f'{int(height):,}',
                                    xy=(rect.get_x() + rect.get_width() / 2, height),
                                    xytext=(0, 3),
                                    textcoords="offset points",
                                    ha='center', va='bottom', fontsize=7)
            autolabel(rects1)
            autolabel(rects2)
            fig.tight_layout()
            canvas = FigureCanvasTkAgg(fig, master=self.chart_frame)
            canvas.draw()
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể tạo biểu đồ: {str(e)}")
            traceback.print_exc()

    def export_excel(self):
        if self.result_data is None:
            messagebox.showwarning("Cảnh báo", "Không có dữ liệu tổng thể để xuất.")
            return
        current_date = datetime.datetime.now().strftime("%Y%m%d_%H%M")
        default_filename = f"BaoCao_TongThe_{current_date}.xlsx"
        current_dir = os.getcwd()
        default_dir = os.path.join(current_dir, "result")
        try:
            if not os.path.exists(default_dir):
                os.makedirs(default_dir)
        except OSError as e:
            messagebox.showerror("Lỗi Tạo Thư Mục", f"Không thể tạo thư mục '{default_dir}': {str(e)}")
            return
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Lưu báo cáo Excel (Tổng thể)",
            initialdir=default_dir,
            initialfile=default_filename
        )
        if not file_path:
            return
        try:
            clean_data = self.result_data.copy()
            clean_data.fillna(0, inplace=True)
            clean_data['percentage_decimal'] = clean_data['percentage'] / 100
            clean_data['percentage_decimal'] = clean_data['percentage_decimal'].clip(upper=1)
            writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
            clean_data.to_excel(writer, sheet_name='BaoCao', index=False,
                              columns=['RegionCode', 'total_customers', 'app_installed', 'percentage_decimal'],
                              startrow=2)
            workbook = writer.book
            worksheet = writer.sheets['BaoCao']
            header_format = workbook.add_format({
                'bold': True, 'text_wrap': False, 'valign': 'vcenter', 'align': 'center',
                'fg_color': '#D7E4BC', 'border': 1, 'font_name': 'Arial', 'font_size': 10
            })
            number_format = workbook.add_format({
                'num_format': '#,##0', 'text_wrap': False, 'border': 1,
                'font_name': 'Arial', 'font_size': 10, 'valign': 'vcenter'
            })
            percent_format = workbook.add_format({
                'num_format': '0.00%', 'text_wrap': False, 'border': 1, 'font_name': 'Arial',
                'font_size': 10, 'valign': 'vcenter'
            })
            total_format = workbook.add_format({
                'bold': True, 'num_format': '#,##0', 'text_wrap': False, 'border': 1,
                'font_name': 'Arial', 'font_size': 10, 'valign': 'vcenter'
            })
            percent_total_format = workbook.add_format({
                'bold': True, 'num_format': '0.00%', 'text_wrap': False, 'border': 1,
                'font_name': 'Arial', 'font_size': 10, 'valign': 'vcenter'
            })
            title_format = workbook.add_format({
                'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter',
                'font_name': 'Arial'
            })
            date_format = workbook.add_format({
                'align': 'left', 'font_name': 'Arial', 'font_size': 9
            })
            region_format = workbook.add_format({
                'text_wrap': False, 'border': 1, 'align': 'center',
                'font_name': 'Arial', 'font_size': 10, 'valign': 'vcenter'
            })
            worksheet.merge_range('A1:D1', 'BÁO CÁO KHÁCH HÀNG ĐÃ CÀI APP THEO KHU VỰC', title_format)
            worksheet.set_row(0, 20)
            worksheet.write(1, 0, f'Ngày xuất báo cáo: {datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")}', date_format)
            headers = ['Mã khu vực', 'Số khách hàng', 'Số khách hàng đã cài app', 'Tỷ lệ (%)']
            for col_num, value in enumerate(headers):
                worksheet.write(2, col_num, value, header_format)
            worksheet.set_row(2, 18)
            worksheet.set_column('A:A', 15, region_format)
            worksheet.set_column('B:B', 20, number_format)
            worksheet.set_column('C:C', 25, number_format)
            worksheet.set_column('D:D', 15, percent_format)
            total_row_index = len(self.result_data) + 3
            total_customers = self.result_data['total_customers'].sum()
            total_app_installed = self.result_data['app_installed'].sum()
            total_percentage_decimal = (total_app_installed / total_customers) if total_customers > 0 else 0
            worksheet.write(total_row_index, 0, 'TỔNG CỘNG', total_format)
            worksheet.write(total_row_index, 1, total_customers, total_format)
            worksheet.write(total_row_index, 2, total_app_installed, total_format)
            worksheet.write(total_row_index, 3, total_percentage_decimal, percent_total_format)
            worksheet.freeze_panes(3, 0)
            writer.close()
            messagebox.showinfo("Thông báo", f"Xuất báo cáo Excel (Tổng thể) thành công!\nĐã lưu tại: {file_path}")
        except PermissionError:
            messagebox.showerror("Lỗi Quyền", f"Không có quyền ghi file vào đường dẫn '{os.path.dirname(file_path)}'.")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể xuất file Excel: {str(e)}")
            traceback.print_exc()

    def export_pdf(self):
        if self.result_data is None:
            messagebox.showwarning("Cảnh báo", "Không có dữ liệu tổng thể để xuất.")
            return
        current_date = datetime.datetime.now().strftime("%Y%m%d_%H%M")
        default_filename = f"BaoCao_TongThe_{current_date}.pdf"
        current_dir = os.getcwd()
        default_dir = os.path.join(current_dir, "result")
        try:
            if not os.path.exists(default_dir):
                os.makedirs(default_dir)
        except OSError as e:
            messagebox.showerror("Lỗi Tạo Thư Mục", f"Không thể tạo thư mục '{default_dir}': {str(e)}")
            return
        file_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            title="Lưu báo cáo PDF (Tổng thể)",
            initialdir=default_dir,
            initialfile=default_filename
        )
        if not file_path:
            return
        pdf = FPDF()
        pdf.add_page()
        try:
            try:
                font_path = os.path.join(os.path.dirname(__file__), 'DejaVuSansCondensed.ttf')
                if not os.path.exists(font_path):
                    font_path_alt = os.path.join(os.getcwd(), 'DejaVuSansCondensed.ttf')
                    if os.path.exists(font_path_alt):
                        font_path = font_path_alt
                    else:
                        messagebox.showwarning("Thiếu Font", "Không tìm thấy font 'DejaVuSansCondensed.ttf'. Sử dụng font mặc định.")
                        pdf.set_font('Arial', '', 10)
                        font_name = 'Arial'
                else:
                    pdf.add_font('DejaVu', '', font_path, uni=True)
                    pdf.set_font('DejaVu', '', 10)
                    font_name = 'DejaVu'
            except RuntimeError as e:
                messagebox.showwarning("Lỗi Font", f"Lỗi khi tải font DejaVu: {e}. Sử dụng font mặc định.")
                pdf.set_font('Arial', '', 10)
                font_name = 'Arial'
            pdf.set_font(font_name, '', 16)
            pdf.cell(0, 10, 'BÁO CÁO KHÁCH HÀNG ĐÃ CÀI APP THEO KHU VỰC', 0, 1, 'C')
            pdf.ln(5)
            pdf.set_font(font_name, '', 10)
            pdf.cell(0, 10, f'Ngày xuất báo cáo: {datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")}', 0, 1)
            pdf.ln(5)
            col_width = [30, 50, 55, 30]
            row_height = 8
            header_font_size = 10
            data_font_size = 9
            pdf.set_fill_color(211, 211, 211)
            pdf.set_font(font_name, '', header_font_size)
            pdf.cell(col_width[0], row_height, 'Mã khu vực', 1, 0, 'C', 1)
            pdf.cell(col_width[1], row_height, 'Số khách hàng', 1, 0, 'C', 1)
            pdf.cell(col_width[2], row_height, 'Số khách hàng đã cài app', 1, 0, 'C', 1)
            pdf.cell(col_width[3], row_height, 'Tỷ lệ (%)', 1, 1, 'C', 1)
            pdf.set_font(font_name, '', data_font_size)
            for i, row in self.result_data.iterrows():
                pdf.cell(col_width[0], row_height, str(row['RegionCode']), 1, 0, 'C')
                pdf.cell(col_width[1], row_height, f"{int(row['total_customers']):,}", 1, 0, 'R')
                pdf.cell(col_width[2], row_height, f"{int(row['app_installed']):,}", 1, 0, 'R')
                pdf.cell(col_width[3], row_height, f"{row['percentage']:.2f}%", 1, 1, 'R')
            total_customers = self.result_data['total_customers'].sum()
            total_app_installed = self.result_data['app_installed'].sum()
            total_percentage = (total_app_installed / total_customers * 100) if total_customers > 0 else 0
            pdf.set_font(font_name, '', header_font_size)
            pdf.cell(col_width[0], row_height, 'TỔNG CỘNG', 1, 0, 'C', 1)
            pdf.cell(col_width[1], row_height, f"{int(total_customers):,}", 1, 0, 'R', 1)
            pdf.cell(col_width[2], row_height, f"{int(total_app_installed):,}", 1, 0, 'R', 1)
            pdf.cell(col_width[3], row_height, f"{total_percentage:.2f}%", 1, 1, 'R', 1)
            pdf.ln(10)
            chart_path = 'temp_chart_for_pdf.png'
            try:
                fig_pdf, ax_pdf = plt.subplots(figsize=(8, 5))
                chart_data_pdf = self.result_data.sort_values('RegionCode').copy()
                regions_pdf = chart_data_pdf['RegionCode'].astype(str)
                total_pdf = chart_data_pdf['total_customers']
                installed_pdf = chart_data_pdf['app_installed']
                x_pdf = np.arange(len(regions_pdf))
                width_pdf = 0.35
                rects1_pdf = ax_pdf.bar(x_pdf - width_pdf/2, total_pdf, width_pdf, label='Tổng số khách hàng')
                rects2_pdf = ax_pdf.bar(x_pdf + width_pdf/2, installed_pdf, width_pdf, label='Đã cài app')
                ax_pdf.set_title('Thống kê khách hàng đã cài app theo khu vực', fontsize=12)
                ax_pdf.set_xlabel('Mã khu vực', fontsize=9)
                ax_pdf.set_ylabel('Số lượng', fontsize=9)
                ax_pdf.set_xticks(x_pdf)
                ax_pdf.set_xticklabels(regions_pdf, rotation=45, ha='right', fontsize=7)  # Fixed line
                ax_pdf.tick_params(axis='y', labelsize=8)
                ax_pdf.legend(fontsize=8)
                def autolabel_pdf(rects, axis):
                    for rect in rects:
                        height = rect.get_height()
                        if height > 0:
                            axis.annotate(f'{int(height):,}',
                                        xy=(rect.get_x() + rect.get_width() / 2, height),
                                        xytext=(0, 2), textcoords="offset points",
                                        ha='center', va='bottom', fontsize=6)
                autolabel_pdf(rects1_pdf, ax_pdf)
                autolabel_pdf(rects2_pdf, ax_pdf)
                fig_pdf.tight_layout()
                fig_pdf.savefig(chart_path, bbox_inches='tight', dpi=150)
                plt.close(fig_pdf)
                pdf.set_font(font_name, '', 11)
                pdf.cell(0, 10, 'Biểu đồ thống kê theo khu vực:', 0, 1, 'L')
                pdf.ln(2)
                img_w = 160
                page_w = pdf.w - 2 * pdf.l_margin
                img_x = (page_w - img_w) / 2 + pdf.l_margin
                pdf.image(chart_path, x=img_x, w=img_w)
            except Exception as e:
                messagebox.showerror("Lỗi Biểu Đồ PDF", f"Không thể tạo biểu đồ: {str(e)}")
                traceback.print_exc()
            finally:
                if os.path.exists(chart_path):
                    try:
                        os.remove(chart_path)
                    except Exception as e_remove:
                        print(f"Warning: Could not remove temp chart file '{chart_path}': {e_remove}")
            pdf.output(file_path)
            messagebox.showinfo("Thông báo", f"Xuất báo cáo PDF (Tổng thể) thành công!\nĐã lưu tại: {file_path}")
        except PermissionError:
            messagebox.showerror("Lỗi Quyền", f"Không có quyền ghi file vào đường dẫn '{os.path.dirname(file_path)}'.")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể xuất file PDF: {str(e)}")
            traceback.print_exc()

    def export_detailed_excel(self):
        if self.detailed_data is None or self.detailed_data.empty:
            messagebox.showwarning("Cảnh báo", "Không có dữ liệu chi tiết để xuất.")
            return

        # Tạo tên file mặc định
        current_date = datetime.datetime.now().strftime("%Y%m%d_%H%M")
        default_filename = f"BaoCao_ChiTiet_{current_date}.xlsx"

        # Tạo thư mục /result nếu chưa tồn tại
        current_dir = os.getcwd()
        default_dir = os.path.join(current_dir, "result")
        try:
            if not os.path.exists(default_dir):
                os.makedirs(default_dir)
        except OSError as e:
            messagebox.showerror("Lỗi Tạo Thư Mục", f"Không thể tạo thư mục '{default_dir}': {str(e)}")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Lưu báo cáo Excel (Chi tiết)",
            initialdir=default_dir,
            initialfile=default_filename
        )

        if not file_path:
            return

        try:
            # Chuẩn bị dữ liệu chi tiết
            clean_data = self.detailed_data.sort_values(['RegionCode', 'StreetCode']).copy()
            clean_data = clean_data[['CustomerRegistationNo', 'CustomerName', 'RegionCode', 'Name']]
            clean_data.rename(columns={
                'CustomerRegistationNo': 'Mã danh bộ',
                'CustomerName': 'Tên khách hàng',
                'RegionCode': 'Mã khu vực',
                'Name': 'Tên đường'
            }, inplace=True)

            # Thay thế NaN bằng chuỗi rỗng và loại bỏ ký tự xuống dòng
            for col in clean_data.columns:
                clean_data[col] = clean_data[col].fillna('').astype(str).str.replace('\n|\r', '', regex=True).str.strip()

            # Tạo Excel Writer
            writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
            workbook = writer.book

            # Định dạng
            header_format = workbook.add_format({
                'bold': True, 'text_wrap': False, 'valign': 'vcenter', 'align': 'center',
                'fg_color': '#D7E4BC', 'border': 1, 'font_name': 'Arial', 'font_size': 10
            })
            text_format = workbook.add_format({
                'text_wrap': False, 'valign': 'vcenter', 'border': 1,
                'font_name': 'Arial', 'font_size': 10
            })
            center_format = workbook.add_format({
                'text_wrap': False, 'valign': 'vcenter', 'border': 1, 'align': 'center',
                'font_name': 'Arial', 'font_size': 10
            })
            title_format = workbook.add_format({
                'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter',
                'font_name': 'Arial'
            })
            date_format = workbook.add_format({
                'align': 'left', 'font_name': 'Arial', 'font_size': 9
            })

            # Lấy danh sách các khu vực duy nhất
            regions = clean_data['Mã khu vực'].unique()

            # Xuất dữ liệu cho từng khu vực vào sheet riêng
            for region in regions:
                # Lọc dữ liệu cho khu vực hiện tại
                region_data = clean_data[clean_data['Mã khu vực'] == region].copy()
                if region_data.empty:
                    continue  # Bỏ qua nếu khu vực không có dữ liệu

                # Tạo tên sheet (giới hạn 31 ký tự, thay ký tự không hợp lệ)
                sheet_name = f"Region_{region}"[:31].replace('/', '_').replace('\\', '_')

                # Ghi dữ liệu vào sheet
                region_data.to_excel(writer, sheet_name=sheet_name, index=False, startrow=2)

                # Lấy worksheet
                worksheet = writer.sheets[sheet_name]

                # Thêm tiêu đề lớn và ngày xuất
                worksheet.merge_range('A1:D1', f'DANH SÁCH KHÁCH HÀNG ĐÃ CÀI APP - KHU VỰC {region}', title_format)
                worksheet.set_row(0, 20)
                worksheet.write(1, 0, f'Ngày xuất báo cáo: {datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")}', date_format)

                # Áp dụng định dạng cho tiêu đề cột
                for col_num, value in enumerate(region_data.columns):
                    worksheet.write(2, col_num, value, header_format)
                worksheet.set_row(2, 18)

                # Thiết lập độ rộng cột
                worksheet.set_column('A:A', 20, center_format)  # Mã danh bộ
                worksheet.set_column('B:B', 35, text_format)    # Tên khách hàng
                worksheet.set_column('C:C', 15, center_format)  # Mã khu vực
                worksheet.set_column('D:D', 50, text_format)    # Tên đường

                # Freeze panes
                worksheet.freeze_panes(3, 0)

            # Lưu file
            writer.close()

            messagebox.showinfo("Thông báo", f"Xuất báo cáo Excel (Chi tiết) thành công!\nĐã lưu tại: {file_path}")

        except PermissionError:
            messagebox.showerror("Lỗi Quyền", f"Không có quyền ghi file vào đường dẫn '{os.path.dirname(file_path)}' hoặc file đang được mở.")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể xuất file Excel chi tiết: {str(e)}")
            traceback.print_exc()

    def export_detailed_pdf(self):
        if self.detailed_data is None or self.detailed_data.empty:
            messagebox.showwarning("Cảnh báo", "Không có dữ liệu chi tiết để xuất.")
            return
        current_date = datetime.datetime.now().strftime("%Y%m%d_%H%M")
        default_filename = f"BaoCao_ChiTiet_{current_date}.pdf"
        current_dir = os.getcwd()
        default_dir = os.path.join(current_dir, "result")
        try:
            if not os.path.exists(default_dir):
                os.makedirs(default_dir)
        except OSError as e:
            messagebox.showerror("Lỗi Tạo Thư Mục", f"Không thể tạo thư mục '{default_dir}': {str(e)}")
            return
        file_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            title="Lưu báo cáo PDF (Chi tiết)",
            initialdir=default_dir,
            initialfile=default_filename
        )
        if not file_path:
            return
        pdf = FPDF(orientation='L')
        pdf.add_page()
        try:
            try:
                font_path = os.path.join(os.path.dirname(__file__), 'DejaVuSansCondensed.ttf')
                if not os.path.exists(font_path):
                    font_path_alt = os.path.join(os.getcwd(), 'DejaVuSansCondensed.ttf')
                    if os.path.exists(font_path_alt):
                        font_path = font_path_alt
                    else:
                        messagebox.showwarning("Thiếu Font", "Không tìm thấy font 'DejaVuSansCondensed.ttf'. Sử dụng font mặc định.")
                        pdf.set_font('Arial', '', 10)
                        font_name = 'Arial'
                else:
                    pdf.add_font('DejaVu', '', font_path, uni=True)
                    pdf.set_font('DejaVu', '', 10)
                    font_name = 'DejaVu'
            except RuntimeError as e:
                messagebox.showwarning("Lỗi Font", f"Lỗi khi tải font DejaVu: {e}. Sử dụng font mặc định.")
                pdf.set_font('Arial', '', 10)
                font_name = 'Arial'
            pdf.set_font(font_name, '', 16)
            pdf.cell(0, 10, 'DANH SÁCH KHÁCH HÀNG ĐÃ CÀI APP', 0, 1, 'C')
            pdf.ln(5)
            pdf.set_font(font_name, '', 10)
            pdf.cell(0, 10, f'Ngày xuất báo cáo: {datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")}', 0, 1, 'L')
            pdf.ln(5)
            total_width = pdf.w - 2 * pdf.l_margin
            col_width = [total_width * 0.15, total_width * 0.25, total_width * 0.10, total_width * 0.50]
            row_height = 7
            header_font_size = 9
            data_font_size = 8
            pdf.set_fill_color(211, 211, 211)
            pdf.set_font(font_name, '', header_font_size)
            pdf.cell(col_width[0], row_height, 'Mã danh bộ', 1, 0, 'C', 1)
            pdf.cell(col_width[1], row_height, 'Tên khách hàng', 1, 0, 'C', 1)
            pdf.cell(col_width[2], row_height, 'Mã khu vực', 1, 0, 'C', 1)
            pdf.cell(col_width[3], row_height, 'Tên đường', 1, 1, 'C', 1)
            pdf.set_font(font_name, '', data_font_size)
            sorted_detailed_data_pdf = self.detailed_data.sort_values(['RegionCode', 'StreetCode']).copy()
            for col in ['CustomerRegistationNo', 'CustomerName', 'RegionCode', 'Name']:
                sorted_detailed_data_pdf[col] = sorted_detailed_data_pdf[col].fillna('').astype(str).str.replace('\n|\r', '', regex=True).str.strip()
            for i, row in sorted_detailed_data_pdf.iterrows():
                if pdf.get_y() + row_height > pdf.page_break_trigger:
                    pdf.add_page()
                    pdf.set_fill_color(211, 211, 211)
                    pdf.set_font(font_name, '', header_font_size)
                    pdf.cell(col_width[0], row_height, 'Mã danh bộ', 1, 0, 'C', 1)
                    pdf.cell(col_width[1], row_height, 'Tên khách hàng', 1, 0, 'C', 1)
                    pdf.cell(col_width[2], row_height, 'Mã khu vực', 1, 0, 'C', 1)
                    pdf.cell(col_width[3], row_height, 'Tên đường', 1, 1, 'C', 1)
                    pdf.set_font(font_name, '', data_font_size)
                pdf.cell(col_width[0], row_height, row['CustomerRegistationNo'], 1, 0, 'C')
                x_before_name = pdf.get_x()
                y_before_name = pdf.get_y()
                pdf.multi_cell(col_width[1], row_height, row['CustomerName'], 1, 'L')
                y_after_name = pdf.get_y()
                pdf.set_xy(x_before_name + col_width[1], y_before_name)
                pdf.cell(col_width[2], row_height, row['RegionCode'], 1, 0, 'C')
                x_before_street = pdf.get_x()
                y_before_street = pdf.get_y()
                pdf.multi_cell(col_width[3], row_height, row['Name'], 1, 'L')
                y_after_street = pdf.get_y()
                pdf.set_y(max(y_after_name, y_after_street))
            pdf.output(file_path)
            messagebox.showinfo("Thông báo", f"Xuất báo cáo PDF (Chi tiết) thành công!\nĐã lưu tại: {file_path}")
        except PermissionError:
            messagebox.showerror("Lỗi Quyền", f"Không có quyền ghi file vào đường dẫn '{os.path.dirname(file_path)}' hoặc file đang được mở.")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể xuất file PDF chi tiết: {str(e)}")
            traceback.print_exc()

    def export_not_installed_excel(self):
        if self.not_installed_data is None or self.not_installed_data.empty:
            messagebox.showwarning("Cảnh báo", "Không có dữ liệu khách hàng chưa cài app để xuất.")
            return

        # Tạo tên file mặc định
        current_date = datetime.datetime.now().strftime("%Y%m%d_%H%M")
        default_filename = f"BaoCao_ChuaCaiApp_{current_date}.xlsx"

        # Tạo thư mục /result nếu chưa tồn tại
        current_dir = os.getcwd()
        default_dir = os.path.join(current_dir, "result")
        try:
            if not os.path.exists(default_dir):
                os.makedirs(default_dir)
        except OSError as e:
            messagebox.showerror("Lỗi Tạo Thư Mục", f"Không thể tạo thư mục '{default_dir}': {str(e)}")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Lưu báo cáo Excel (Chưa cài app)",
            initialdir=default_dir,
            initialfile=default_filename
        )

        if not file_path:
            return

        try:
            # Chuẩn bị dữ liệu
            clean_data = self.not_installed_data.sort_values(['RegionCode', 'StreetCode']).copy()
            clean_data = clean_data[['CustomerRegistationNo', 'CustomerName', 'RegionCode', 'Name']]
            clean_data.rename(columns={
                'CustomerRegistationNo': 'Mã danh bộ',
                'CustomerName': 'Tên khách hàng',
                'RegionCode': 'Mã khu vực',
                'Name': 'Tên đường'
            }, inplace=True)

            # Thay thế NaN và loại bỏ ký tự xuống dòng
            for col in clean_data.columns:
                clean_data[col] = clean_data[col].fillna('').astype(str).str.replace('\n|\r', '', regex=True).str.strip()

            # Tạo Excel Writer
            writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
            workbook = writer.book

            # Định dạng
            header_format = workbook.add_format({
                'bold': True, 'text_wrap': False, 'valign': 'vcenter', 'align': 'center',
                'fg_color': '#D7E4BC', 'border': 1, 'font_name': 'Arial', 'font_size': 10
            })
            text_format = workbook.add_format({
                'text_wrap': False, 'valign': 'vcenter', 'border': 1,
                'font_name': 'Arial', 'font_size': 10
            })
            center_format = workbook.add_format({
                'text_wrap': False, 'valign': 'vcenter', 'border': 1, 'align': 'center',
                'font_name': 'Arial', 'font_size': 10
            })
            title_format = workbook.add_format({
                'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter',
                'font_name': 'Arial'
            })
            date_format = workbook.add_format({
                'align': 'left', 'font_name': 'Arial', 'font_size': 9
            })

            # Lấy danh sách các khu vực duy nhất
            regions = clean_data['Mã khu vực'].unique()

            # Xuất dữ liệu cho từng khu vực vào sheet riêng
            for region in regions:
                # Lọc dữ liệu cho khu vực hiện tại
                region_data = clean_data[clean_data['Mã khu vực'] == region].copy()
                if region_data.empty:
                    continue  # Bỏ qua nếu khu vực không có dữ liệu

                # Tạo tên sheet
                sheet_name = f"Region_{region}"[:31].replace('/', '_').replace('\\', '_')

                # Ghi dữ liệu vào sheet
                region_data.to_excel(writer, sheet_name=sheet_name, index=False, startrow=2)

                # Lấy worksheet
                worksheet = writer.sheets[sheet_name]

                # Thêm tiêu đề lớn và ngày xuất
                worksheet.merge_range('A1:D1', f'DANH SÁCH KHÁCH HÀNG CHƯA CÀI APP - KHU VỰC {region}', title_format)
                worksheet.set_row(0, 20)
                worksheet.write(1, 0, f'Ngày xuất báo cáo: {datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")}', date_format)

                # Áp dụng định dạng cho tiêu đề cột
                for col_num, value in enumerate(region_data.columns):
                    worksheet.write(2, col_num, value, header_format)
                worksheet.set_row(2, 18)

                # Thiết lập độ rộng cột
                worksheet.set_column('A:A', 20, center_format)  # Mã danh bộ
                worksheet.set_column('B:B', 35, text_format)    # Tên khách hàng
                worksheet.set_column('C:C', 15, center_format)  # Mã khu vực
                worksheet.set_column('D:D', 50, text_format)    # Tên đường

                # Freeze panes
                worksheet.freeze_panes(3, 0)

            # Lưu file
            writer.close()

            messagebox.showinfo("Thông báo", f"Xuất báo cáo Excel (Chưa cài app) thành công!\nĐã lưu tại: {file_path}")

        except PermissionError:
            messagebox.showerror("Lỗi Quyền", f"Không có quyền ghi file vào đường dẫn '{os.path.dirname(file_path)}' hoặc file đang được mở.")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể xuất file Excel chưa cài app: {str(e)}")
            traceback.print_exc()

    def export_not_installed_pdf(self):
        if self.not_installed_data is None or self.not_installed_data.empty:
            messagebox.showwarning("Cảnh báo", "Không có dữ liệu khách hàng chưa cài app để xuất.")
            return
        current_date = datetime.datetime.now().strftime("%Y%m%d_%H%M")
        default_filename = f"BaoCao_ChuaCaiApp_{current_date}.pdf"
        current_dir = os.getcwd()
        default_dir = os.path.join(current_dir, "result")
        try:
            if not os.path.exists(default_dir):
                os.makedirs(default_dir)
        except OSError as e:
            messagebox.showerror("Lỗi Tạo Thư Mục", f"Không thể tạo thư mục '{default_dir}': {str(e)}")
            return
        file_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            title="Lưu báo cáo PDF (Chưa cài app)",
            initialdir=default_dir,
            initialfile=default_filename
        )
        if not file_path:
            return
        pdf = FPDF(orientation='L')
        pdf.add_page()
        try:
            try:
                font_path = os.path.join(os.path.dirname(__file__), 'DejaVuSansCondensed.ttf')
                if not os.path.exists(font_path):
                    font_path_alt = os.path.join(os.getcwd(), 'DejaVuSansCondensed.ttf')
                    if os.path.exists(font_path_alt):
                        font_path = font_path_alt
                    else:
                        messagebox.showwarning("Thiếu Font", "Không tìm thấy font 'DejaVuSansCondensed.ttf'. Sử dụng font mặc định.")
                        pdf.set_font('Arial', '', 10)
                        font_name = 'Arial'
                else:
                    pdf.add_font('DejaVu', '', font_path, uni=True)
                    pdf.set_font('DejaVu', '', 10)
                    font_name = 'DejaVu'
            except RuntimeError as e:
                messagebox.showwarning("Lỗi Font", f"Lỗi khi tải font DejaVu: {e}. Sử dụng font mặc định.")
                pdf.set_font('Arial', '', 10)
                font_name = 'Arial'
            pdf.set_font(font_name, '', 16)
            pdf.cell(0, 10, 'DANH SÁCH KHÁCH HÀNG CHƯA CÀI APP', 0, 1, 'C')
            pdf.ln(5)
            pdf.set_font(font_name, '', 10)
            pdf.cell(0, 10, f'Ngày xuất báo cáo: {datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")}', 0, 1, 'L')
            pdf.ln(5)
            total_width = pdf.w - 2 * pdf.l_margin
            col_width = [total_width * 0.15, total_width * 0.25, total_width * 0.10, total_width * 0.50]
            row_height = 7
            header_font_size = 9
            data_font_size = 8
            pdf.set_fill_color(211, 211, 211)
            pdf.set_font(font_name, '', header_font_size)
            pdf.cell(col_width[0], row_height, 'Mã danh bộ', 1, 0, 'C', 1)
            pdf.cell(col_width[1], row_height, 'Tên khách hàng', 1, 0, 'C', 1)
            pdf.cell(col_width[2], row_height, 'Mã khu vực', 1, 0, 'C', 1)
            pdf.cell(col_width[3], row_height, 'Tên đường', 1, 1, 'C', 1)
            pdf.set_font(font_name, '', data_font_size)
            sorted_not_installed_data = self.not_installed_data.sort_values(['RegionCode', 'StreetCode']).copy()
            for col in ['CustomerRegistationNo', 'CustomerName', 'RegionCode', 'Name']:
                sorted_not_installed_data[col] = sorted_not_installed_data[col].fillna('').astype(str).str.replace('\n|\r', '', regex=True).str.strip()
            for i, row in sorted_not_installed_data.iterrows():
                if pdf.get_y() + row_height > pdf.page_break_trigger:
                    pdf.add_page()
                    pdf.set_fill_color(211, 211, 211)
                    pdf.set_font(font_name, '', header_font_size)
                    pdf.cell(col_width[0], row_height, 'Mã danh bộ', 1, 0, 'C', 1)
                    pdf.cell(col_width[1], row_height, 'Tên khách hàng', 1, 0, 'C', 1)
                    pdf.cell(col_width[2], row_height, 'Mã khu vực', 1, 0, 'C', 1)
                    pdf.cell(col_width[3], row_height, 'Tên đường', 1, 1, 'C', 1)
                    pdf.set_font(font_name, '', data_font_size)
                pdf.cell(col_width[0], row_height, row['CustomerRegistationNo'], 1, 0, 'C')
                x_before_name = pdf.get_x()
                y_before_name = pdf.get_y()
                pdf.multi_cell(col_width[1], row_height, row['CustomerName'], 1, 'L')
                y_after_name = pdf.get_y()
                pdf.set_xy(x_before_name + col_width[1], y_before_name)
                pdf.cell(col_width[2], row_height, row['RegionCode'], 1, 0, 'C')
                x_before_street = pdf.get_x()
                y_before_street = pdf.get_y()
                pdf.multi_cell(col_width[3], row_height, row['Name'], 1, 'L')
                y_after_street = pdf.get_y()
                pdf.set_y(max(y_after_name, y_after_street))
            pdf.output(file_path)
            messagebox.showinfo("Thông báo", f"Xuất báo cáo PDF (Chưa cài app) thành công!\nĐã lưu tại: {file_path}")
        except PermissionError:
            messagebox.showerror("Lỗi Quyền", f"Không có quyền ghi file vào đường dẫn '{os.path.dirname(file_path)}' hoặc file đang được mở.")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể xuất file PDF chưa cài app: {str(e)}")
            traceback.print_exc()

if __name__ == "__main__":
    # Thêm xử lý đường dẫn khi đóng gói bằng PyInstaller
    if getattr(sys, 'frozen', False):
        application_path = os.path.dirname(sys.executable)
    elif __file__:
        application_path = os.path.dirname(__file__)
    else:
        application_path = os.getcwd()
    os.chdir(application_path)
    print(f"Current working directory: {os.getcwd()}")

    root = tk.Tk()
    try:
        icon_path = os.path.join(application_path, 'app_icon.ico')
        if os.path.exists(icon_path):
            root.iconbitmap(icon_path)
    except Exception as e:
        print(f"Could not set application icon: {e}")

    app = CustomerAppAnalyzer(root)
    root.mainloop()