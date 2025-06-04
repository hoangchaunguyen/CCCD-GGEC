"""
EXCEL CONSOLIDATOR GUI
======================

Giao diện người dùng sử dụng Tkinter cho công cụ gom dữ liệu Excel. 
Cung cấp giao diện trực quan để chọn thư mục, file xuất, và hiển thị 
tiến trình xử lý.

Cấu trúc chính:
---------------
class ExcelConsolidatorApp:
    Quản lý toàn bộ giao diện và tương tác người dùng

Các thành phần chính:
1. __init__(root: tk.Tk)
2. create_widgets()
3. style_widgets()
4. browse_input()
5. browse_output()
6. start_consolidation_thread()
7. update_progress(value: int, text: str)
8. start_consolidation()

Chi tiết từng phương thức:
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from excel_consolidator import ExcelConsolidator
import os
import threading
import time
import logging

# Cấu hình logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("ExcelConsolidator")
logger.setLevel(logging.INFO)

class ExcelConsolidatorApp:
    def __init__(self, root: tk.Tk):
        """
        Khởi tạo ứng dụng giao diện
        
        Args:
            root: Cửa sổ chính của ứng dụng Tkinter
            
        Attributes:
            input_path (StringVar): Đường dẫn thư mục nguồn
            output_path (StringVar): Đường dẫn file xuất
            status (StringVar): Thông báo trạng thái
            progress_value (IntVar): Giá trị tiến trình (0-100)
            progress_text (StringVar): Mô tả tiến trình
        """
        self.root = root
        self.root.title("Excel Consolidator")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        # Biến giao diện
        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.status = tk.StringVar(value="Sẵn sàng")
        self.progress_value = tk.IntVar(value=0)
        self.progress_text = tk.StringVar(value="")
        
        # Thiết lập giao diện
        self.create_widgets()
        self.style_widgets()
        
        # Hướng log vào giao diện
        self.setup_log_redirect()

    def create_widgets(self):
        """
        Tạo và bố trí các thành phần giao diện:
        
        Cấu trúc chính:
        - Main Frame: Khung chứa chính
        - Input Section: Chọn thư mục nguồn
        - Output Section: Chọn file xuất
        - Progress Section: Thanh tiến trình
        - Status Section: Hiển thị trạng thái
        - Button Section: Nút thực hiện và thoát
        - Log Console: Hiển thị log hệ thống
        """
        main_frame = ttk.Frame(self.root, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # ---- Tiêu đề ứng dụng ----
        title = ttk.Label(
            main_frame, 
            text="CÔNG CỤ GOM DỮ LIỆU EXCEL",
            font=("Arial", 16, "bold"),
            foreground="#2C3E50"
        )
        title.grid(row=0, column=0, columnspan=3, pady=15)
        
        # ---- Khu vực nhập liệu ----
        input_frame = ttk.LabelFrame(main_frame, text=" Thư mục nguồn ")
        input_frame.grid(row=1, column=0, columnspan=3, sticky=tk.EW, pady=10, padx=5)
        
        ttk.Label(input_frame, text="Thư mục chứa file Excel:").pack(
            side=tk.LEFT, padx=(10, 5), pady=10)
            
        input_entry = ttk.Entry(input_frame, textvariable=self.input_path, width=60)
        input_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5), pady=5)
        
        ttk.Button(
            input_frame, 
            text="Duyệt...", 
            command=self.browse_input,
            width=10
        ).pack(side=tk.LEFT, padx=(0, 10), pady=5)
        
        # ---- Khu vực xuất dữ liệu ----
        output_frame = ttk.LabelFrame(main_frame, text=" File kết quả ")
        output_frame.grid(row=2, column=0, columnspan=3, sticky=tk.EW, pady=10, padx=5)
        
        ttk.Label(output_frame, text="Đường dẫn file xuất:").pack(
            side=tk.LEFT, padx=(10, 5), pady=10)
            
        output_entry = ttk.Entry(output_frame, textvariable=self.output_path, width=60)
        output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5), pady=5)
        
        ttk.Button(
            output_frame, 
            text="Chọn...", 
            command=self.browse_output,
            width=10
        ).pack(side=tk.LEFT, padx=(0, 10), pady=5)
        
        # ---- Thanh tiến trình ----
        progress_frame = ttk.LabelFrame(main_frame, text=" Tiến trình xử lý ")
        progress_frame.grid(row=3, column=0, columnspan=3, sticky=tk.EW, pady=15, padx=5)
        
        ttk.Label(progress_frame, text="Trạng thái:").pack(anchor=tk.W, padx=10, pady=(10, 0))
        self.status_label = ttk.Label(
            progress_frame, 
            textvariable=self.status, 
            foreground="#3498DB",
            font=("Arial", 9, "bold")
        )
        self.status_label.pack(fill=tk.X, padx=10, pady=(0, 5))
        
        ttk.Label(progress_frame, text="Tiến độ:").pack(anchor=tk.W, padx=10)
        progress_bar = ttk.Progressbar(
            progress_frame, 
            variable=self.progress_value, 
            maximum=100,
            mode='determinate'
        )
        progress_bar.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(
            progress_frame, 
            textvariable=self.progress_text,
            foreground="#27AE60"
        ).pack(anchor=tk.W, padx=10, pady=(0, 10))
        
        # ---- Nút điều khiển ----
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=4, column=0, columnspan=3, pady=20)
        
        self.consolidate_btn = ttk.Button(
            btn_frame, 
            text="BẮT ĐẦU GOM DỮ LIỆU", 
            command=self.start_consolidation_thread, 
            width=25,
            style="Accent.TButton"
        )
        self.consolidate_btn.pack(side=tk.LEFT, padx=20)
        
        ttk.Button(
            btn_frame, 
            text="THOÁT ỨNG DỤNG", 
            command=self.root.destroy, 
            width=20
        ).pack(side=tk.RIGHT, padx=20)
        
        # ---- Bảng log hệ thống ----
        log_frame = ttk.LabelFrame(main_frame, text=" Nhật ký xử lý ")
        log_frame.grid(row=5, column=0, columnspan=3, sticky=tk.NSEW, pady=10, padx=5)
        
        self.log_text = tk.Text(log_frame, height=10, wrap=tk.WORD, bg="#2C3E50", fg="#ECF0F1")
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        scrollbar = ttk.Scrollbar(
            log_frame, 
            command=self.log_text.yview
        )
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)

    def style_widgets(self):
        """
        Thiết lập phong cách cho các thành phần giao diện
        
        Sử dụng ttk.Style để tạo giao diện hiện đại, 
        phù hợp với tiêu chuẩn ứng dụng Windows
        """
        style = ttk.Style()
        
        # Cấu hình style chung
        style.configure("TFrame", background="#ECF0F1")
        style.configure("TLabel", background="#ECF0F1", font=("Arial", 9))
        style.configure("TButton", font=("Arial", 9, "bold"), padding=6)
        style.configure("TEntry", font=("Arial", 9))
        style.configure("TLabelframe", background="#ECF0F1")
        style.configure("TLabelframe.Label", background="#ECF0F1", font=("Arial", 10, "bold"))
        
        # Style đặc biệt cho nút chính
        style.map("Accent.TButton",
            background=[("active", "#27AE60"), ("pressed", "#2ECC71")],
            foreground=[("active", "white"), ("pressed", "white")]
        )
        style.configure("Accent.TButton", background="#2ECC71", foreground="white")

    def browse_input(self):
        """
        Mở hộp thoại chọn thư mục nguồn
        
        Chức năng:
            - Hiển thị dialog chọn thư mục
            - Tự động đề xuất file xuất mặc định
            - Cập nhật biến input_path
        """
        folder = filedialog.askdirectory(
            title="Chọn thư mục chứa file Excel",
            mustexist=True
        )
        if folder:
            self.input_path.set(folder)
            # Tạo tên file xuất mặc định
            default_output = os.path.join(folder, "consolidated_data.xlsx")
            self.output_path.set(default_output)

    def browse_output(self):
        """
        Mở hộp thoại chọn file xuất
        
        Chức năng:
            - Hiển thị dialog lưu file
            - Chỉ định định dạng .xlsx
            - Kiểm tra ghi đè file nếu cần
        """
        file = filedialog.asksaveasfilename(
            title="Lưu file kết quả",
            defaultextension=".xlsx",
            filetypes=[
                ("Excel files", "*.xlsx"),
                ("All files", "*.*")
            ]
        )
        if file:
            self.output_path.set(file)

    def setup_log_redirect(self):
        """
        Thiết lập chuyển hướng log hệ thống vào giao diện
        
        Tạo custom logging handler để hiển thị log
        trong text widget của giao diện
        """
        class TextHandler(logging.Handler):
            def __init__(self, text_widget):
                super().__init__()
                self.text_widget = text_widget
                self.setFormatter(logging.Formatter(
                    '%(asctime)s - %(levelname)s - %(message)s',
                    datefmt='%H:%M:%S'
                ))
            
            def emit(self, record):
                msg = self.format(record)
                self.text_widget.insert(tk.END, msg + "\n")
                self.text_widget.see(tk.END)
        
        # Đăng ký handler với hệ thống log
        text_handler = TextHandler(self.log_text)
        logger.addHandler(text_handler)

    def start_consolidation_thread(self):
        """
        Khởi chạy tiến trình gom dữ liệu trong luồng riêng
        
        Chức năng:
            - Vô hiệu hóa nút trong khi xử lý
            - Xóa log cũ
            - Khởi tạo và chạy thread
        """
        self.consolidate_btn.config(state=tk.DISABLED)
        self.log_text.delete(1.0, tk.END)
        self.progress_value.set(0)
        self.progress_text.set("")
        
        thread = threading.Thread(
            target=self.start_consolidation, 
            daemon=True
        )
        thread.start()

    def update_progress(self, value: int, text: str):
        """
        Cập nhật giao diện tiến trình
        
        Args:
            value: Giá trị phần trăm (0-100)
            text: Mô tả tiến trình
        """
        self.progress_value.set(value)
        self.progress_text.set(text)
        self.root.update_idletasks()

    def start_consolidation(self):
        """
        Quy trình chính gom dữ liệu Excel
        
        Xử lý:
            1. Kiểm tra đầu vào hợp lệ
            2. Khởi tạo công cụ ExcelConsolidator
            3. Quét và đọc file
            4. Hiển thị tiến trình chi tiết
            5. Xử lý lỗi và thông báo
        
        Quy trình tiến trình:
            - 0-10%: Khởi tạo
            - 10-80%: Đọc file (10% + 70% * (i/total_files))
            - 80-100%: Gom dữ liệu và lưu file
        """
        input_folder = self.input_path.get()
        output_file = self.output_path.get()
        
        # Validate input
        if not input_folder or not output_file:
            messagebox.showerror("Lỗi", "Vui lòng nhập đầy đủ thông tin")
            self.consolidate_btn.config(state=tk.NORMAL)
            return
        
        try:
            # --- GIAI ĐOẠN KHỞI TẠO ---
            self.update_progress(5, "Đang khởi tạo công cụ...")
            self.status.set("Kiểm tra đường dẫn dữ liệu")
            consolidator = ExcelConsolidator(input_folder)
            
            # --- GIAI ĐOẠN QUÉT FILE ---
            self.update_progress(10, "Đang quét thư mục...")
            self.status.set("Tìm kiếm file Excel")
            excel_files = consolidator._scan_files()
            
            if not excel_files:
                self.status.set("Không tìm thấy file Excel!")
                messagebox.showwarning("Cảnh báo", "Thư mục không chứa file Excel hợp lệ")
                self.consolidate_btn.config(state=tk.NORMAL)
                return
                
            self.status.set(f"Đã tìm thấy {len(excel_files)} file")
            
            # --- GIAI ĐOẠN ĐỌC DỮ LIỆU ---
            total_files = len(excel_files)
            for i, file in enumerate(excel_files):
                rel_path = consolidator._get_relative_path(file)
                
                # Cập nhật tiến trình
                progress = 10 + int(70 * i / total_files)
                self.update_progress(
                    progress,
                    f"Đang xử lý file {i+1}/{total_files}: {rel_path}"
                )
                self.status.set(f"Đọc dữ liệu: {rel_path}")
                
                # Đọc file và xử lý
                consolidator.file_data[rel_path] = consolidator._read_excel(file)
                
                # Làm mới giao diện sau mỗi file
                if i % 5 == 0:
                    self.root.update()
            
            # Kiểm tra dữ liệu đọc được
            if not consolidator.all_keys:
                self.status.set("Không có dữ liệu hợp lệ!")
                messagebox.showwarning(
                    "Cảnh báo", 
                    "Không tìm thấy dữ liệu key-value trong các file"
                )
                self.consolidate_btn.config(state=tk.NORMAL)
                return
                
            # --- GIAI ĐOẠN GOM DỮ LIỆU ---
            self.update_progress(85, "Đang tổng hợp dữ liệu...")
            self.status.set(f"Tổng hợp {len(consolidator.all_keys)} keys")
            success = consolidator.consolidate(output_file)
            
            if success:
                # --- HOÀN THÀNH ---
                self.update_progress(100, "Hoàn thành!")
                summary = consolidator.get_summary()
                msg = (
                    f"ĐÃ HOÀN THÀNH!\n\n"
                    f"• Số file: {summary['total_files']}\n"
                    f"• Số key: {summary['total_keys']}\n"
                    f"• File kết quả: {output_file}"
                )
                self.status.set(msg)
                messagebox.showinfo("Thành công", msg)
            else:
                self.status.set("Quá trình gom dữ liệu thất bại")
                messagebox.showerror("Lỗi", "Đã xảy ra lỗi khi gom dữ liệu")
        
        except Exception as e:
            logger.exception("Lỗi hệ thống")
            self.status.set(f"LỖI: {str(e)}")
            messagebox.showerror(
                "Lỗi nghiêm trọng", 
                f"Ứng dụng gặp sự cố:\n{str(e)}\n\nXem chi tiết trong nhật ký"
            )
        finally:
            self.consolidate_btn.config(state=tk.NORMAL)

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelConsolidatorApp(root)
    root.mainloop()