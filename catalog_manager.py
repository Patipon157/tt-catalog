"""
Thailand Trophy - Catalog File Manager
โปรแกรมจัดการไฟล์ Catalog สำหรับ Thailand Trophy

ความสามารถ:
1. ดูรายการไฟล์ทั้งหมดใน folder
2. จัดเรียง/แยกหมวดไฟล์ (รูปภาพ, เอกสาร, อื่นๆ)
3. ค้นหาไฟล์ตามชื่อ
4. สำรองข้อมูล (Backup)
"""

import os
import shutil
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime


# ===== ตั้งค่า folder เริ่มต้น =====
DEFAULT_FOLDER = r"N:\1.ThailandTrophy\7.Catalog"

# ===== ประเภทไฟล์ =====
IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp", ".tiff", ".tif", ".svg", ".ico"}
DOCUMENT_EXTENSIONS = {".pdf", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".txt", ".csv", ".rtf"}


def get_file_category(filename):
    """จำแนกประเภทไฟล์"""
    ext = os.path.splitext(filename)[1].lower()
    if ext in IMAGE_EXTENSIONS:
        return "รูปภาพ"
    elif ext in DOCUMENT_EXTENSIONS:
        return "เอกสาร"
    else:
        return "อื่นๆ"


def format_size(size_bytes):
    """แปลงขนาดไฟล์ให้อ่านง่าย"""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    elif size_bytes < 1024 * 1024 * 1024:
        return f"{size_bytes / (1024 * 1024):.1f} MB"
    else:
        return f"{size_bytes / (1024 * 1024 * 1024):.1f} GB"


class CatalogManager:
    def __init__(self, root):
        self.root = root
        self.root.title("Thailand Trophy - Catalog Manager")
        self.root.geometry("900x650")
        self.root.resizable(True, True)

        self.folder_path = DEFAULT_FOLDER
        self.all_files = []

        self.setup_ui()
        self.load_files()

    def setup_ui(self):
        """สร้างหน้าจอโปรแกรม"""
        # ===== ส่วนบน: เลือก Folder =====
        top_frame = ttk.LabelFrame(self.root, text="Folder", padding=10)
        top_frame.pack(fill="x", padx=10, pady=5)

        self.folder_var = tk.StringVar(value=self.folder_path)
        folder_entry = ttk.Entry(top_frame, textvariable=self.folder_var, width=60)
        folder_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))

        ttk.Button(top_frame, text="เลือก Folder...", command=self.browse_folder).pack(side="left", padx=2)
        ttk.Button(top_frame, text="โหลดใหม่", command=self.load_files).pack(side="left", padx=2)

        # ===== ส่วนปุ่มคำสั่ง =====
        action_frame = ttk.LabelFrame(self.root, text="คำสั่ง", padding=10)
        action_frame.pack(fill="x", padx=10, pady=5)

        ttk.Button(action_frame, text="จัดแยกหมวดไฟล์", command=self.organize_files).pack(side="left", padx=5)
        ttk.Button(action_frame, text="สำรองข้อมูล (Backup)", command=self.backup_files).pack(side="left", padx=5)

        # ช่องค้นหา
        ttk.Label(action_frame, text="  ค้นหา:").pack(side="left", padx=(20, 5))
        self.search_var = tk.StringVar()
        self.search_var.trace("w", lambda *args: self.filter_files())
        search_entry = ttk.Entry(action_frame, textvariable=self.search_var, width=25)
        search_entry.pack(side="left", padx=2)

        # ตัวกรองประเภท
        ttk.Label(action_frame, text="  ประเภท:").pack(side="left", padx=(10, 5))
        self.filter_var = tk.StringVar(value="ทั้งหมด")
        filter_combo = ttk.Combobox(
            action_frame,
            textvariable=self.filter_var,
            values=["ทั้งหมด", "รูปภาพ", "เอกสาร", "อื่นๆ"],
            state="readonly",
            width=10,
        )
        filter_combo.pack(side="left", padx=2)
        filter_combo.bind("<<ComboboxSelected>>", lambda e: self.filter_files())

        # ===== ส่วนสรุป =====
        self.summary_var = tk.StringVar(value="กำลังโหลด...")
        summary_label = ttk.Label(self.root, textvariable=self.summary_var, font=("", 10))
        summary_label.pack(fill="x", padx=10, pady=2)

        # ===== ส่วนตารางแสดงไฟล์ =====
        table_frame = ttk.Frame(self.root)
        table_frame.pack(fill="both", expand=True, padx=10, pady=5)

        columns = ("name", "category", "size", "modified")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=20)

        self.tree.heading("name", text="ชื่อไฟล์", command=lambda: self.sort_column("name"))
        self.tree.heading("category", text="ประเภท", command=lambda: self.sort_column("category"))
        self.tree.heading("size", text="ขนาด", command=lambda: self.sort_column("size"))
        self.tree.heading("modified", text="วันที่แก้ไข", command=lambda: self.sort_column("modified"))

        self.tree.column("name", width=350)
        self.tree.column("category", width=100)
        self.tree.column("size", width=100)
        self.tree.column("modified", width=150)

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # ===== ส่วนล่าง: สถานะ =====
        self.status_var = tk.StringVar(value="พร้อมใช้งาน")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief="sunken", padding=5)
        status_bar.pack(fill="x", padx=10, pady=5)

    def browse_folder(self):
        """เลือก folder ใหม่"""
        folder = filedialog.askdirectory(title="เลือก Folder Catalog")
        if folder:
            self.folder_path = folder
            self.folder_var.set(folder)
            self.load_files()

    def load_files(self):
        """โหลดรายการไฟล์จาก folder"""
        self.folder_path = self.folder_var.get()
        self.all_files = []

        if not os.path.exists(self.folder_path):
            self.summary_var.set(f"ไม่พบ folder: {self.folder_path}")
            self.status_var.set("ไม่พบ folder - กรุณาเลือก folder ใหม่")
            self.tree.delete(*self.tree.get_children())
            return

        try:
            for root_dir, dirs, files in os.walk(self.folder_path):
                for filename in files:
                    filepath = os.path.join(root_dir, filename)
                    try:
                        stat = os.stat(filepath)
                        self.all_files.append({
                            "name": filename,
                            "path": filepath,
                            "relative": os.path.relpath(filepath, self.folder_path),
                            "category": get_file_category(filename),
                            "size": stat.st_size,
                            "modified": datetime.fromtimestamp(stat.st_mtime),
                        })
                    except OSError:
                        continue

            # นับสรุป
            images = sum(1 for f in self.all_files if f["category"] == "รูปภาพ")
            docs = sum(1 for f in self.all_files if f["category"] == "เอกสาร")
            others = sum(1 for f in self.all_files if f["category"] == "อื่นๆ")
            total_size = sum(f["size"] for f in self.all_files)

            self.summary_var.set(
                f"ทั้งหมด {len(self.all_files)} ไฟล์  |  "
                f"รูปภาพ: {images}  |  เอกสาร: {docs}  |  อื่นๆ: {others}  |  "
                f"ขนาดรวม: {format_size(total_size)}"
            )
            self.status_var.set(f"โหลดสำเร็จ - {self.folder_path}")
            self.filter_files()

        except PermissionError:
            self.summary_var.set("ไม่มีสิทธิ์เข้าถึง folder นี้")
            self.status_var.set("Permission Denied")
        except Exception as e:
            self.summary_var.set(f"เกิดข้อผิดพลาด: {e}")
            self.status_var.set("Error")

    def filter_files(self):
        """กรองไฟล์ตามคำค้นหาและประเภท"""
        self.tree.delete(*self.tree.get_children())

        search_text = self.search_var.get().lower()
        category_filter = self.filter_var.get()

        for f in self.all_files:
            # กรองตามคำค้นหา
            if search_text and search_text not in f["name"].lower():
                continue
            # กรองตามประเภท
            if category_filter != "ทั้งหมด" and f["category"] != category_filter:
                continue

            self.tree.insert("", "end", values=(
                f["relative"],
                f["category"],
                format_size(f["size"]),
                f["modified"].strftime("%Y-%m-%d %H:%M"),
            ))

        shown = len(self.tree.get_children())
        self.status_var.set(f"แสดง {shown} จาก {len(self.all_files)} ไฟล์")

    def sort_column(self, col):
        """เรียงลำดับตามคอลัมน์"""
        items = [(self.tree.set(k, col), k) for k in self.tree.get_children()]
        items.sort()
        for index, (val, k) in enumerate(items):
            self.tree.move(k, "", index)

    def organize_files(self):
        """จัดแยกหมวดไฟล์เข้า sub-folder ตามประเภท"""
        if not self.all_files:
            messagebox.showwarning("แจ้งเตือน", "ไม่มีไฟล์ให้จัดเรียง\nกรุณาโหลดไฟล์ก่อน")
            return

        # ยืนยันก่อนทำ
        result = messagebox.askyesno(
            "ยืนยันการจัดแยกหมวด",
            f"จะจัดแยกไฟล์ {len(self.all_files)} ไฟล์ เข้า folder:\n\n"
            f"  📁 รูปภาพ/\n"
            f"  📁 เอกสาร/\n"
            f"  📁 อื่นๆ/\n\n"
            f"ใน {self.folder_path}\n\n"
            f"ต้องการดำเนินการหรือไม่?",
        )

        if not result:
            return

        moved = 0
        errors = 0

        for f in self.all_files:
            category_folder = os.path.join(self.folder_path, f["category"])
            os.makedirs(category_folder, exist_ok=True)

            src = f["path"]
            dst = os.path.join(category_folder, f["name"])

            # ถ้าไฟล์อยู่ใน sub-folder ของหมวดนั้นอยู่แล้ว ข้าม
            if os.path.dirname(f["path"]) == category_folder:
                continue

            # ถ้าชื่อซ้ำ เพิ่มเลขต่อท้าย
            if os.path.exists(dst):
                name, ext = os.path.splitext(f["name"])
                counter = 1
                while os.path.exists(dst):
                    dst = os.path.join(category_folder, f"{name}_{counter}{ext}")
                    counter += 1

            try:
                shutil.move(src, dst)
                moved += 1
            except Exception:
                errors += 1

        messagebox.showinfo(
            "เสร็จสิ้น",
            f"จัดแยกหมวดเสร็จแล้ว!\n\n"
            f"ย้ายสำเร็จ: {moved} ไฟล์\n"
            f"ผิดพลาด: {errors} ไฟล์",
        )
        self.load_files()

    def backup_files(self):
        """สำรองข้อมูลไปยัง folder ที่เลือก"""
        if not self.all_files:
            messagebox.showwarning("แจ้งเตือน", "ไม่มีไฟล์ให้สำรอง\nกรุณาโหลดไฟล์ก่อน")
            return

        backup_dir = filedialog.askdirectory(title="เลือก folder สำหรับเก็บ Backup")
        if not backup_dir:
            return

        # สร้าง folder backup พร้อมวันที่
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_folder = os.path.join(backup_dir, f"Catalog_Backup_{timestamp}")

        result = messagebox.askyesno(
            "ยืนยันการสำรองข้อมูล",
            f"จะ copy ไฟล์ {len(self.all_files)} ไฟล์ ไปที่:\n\n"
            f"{backup_folder}\n\n"
            f"ต้องการดำเนินการหรือไม่?",
        )

        if not result:
            return

        self.status_var.set("กำลังสำรองข้อมูล...")
        self.root.update()

        copied = 0
        errors = 0

        try:
            shutil.copytree(self.folder_path, backup_folder)
            copied = len(self.all_files)
        except Exception:
            # ถ้า copytree ไม่ได้ ทำทีละไฟล์
            os.makedirs(backup_folder, exist_ok=True)
            for f in self.all_files:
                try:
                    rel_path = os.path.relpath(f["path"], self.folder_path)
                    dst = os.path.join(backup_folder, rel_path)
                    os.makedirs(os.path.dirname(dst), exist_ok=True)
                    shutil.copy2(f["path"], dst)
                    copied += 1
                except Exception:
                    errors += 1

        messagebox.showinfo(
            "เสร็จสิ้น",
            f"สำรองข้อมูลเสร็จแล้ว!\n\n"
            f"Copy สำเร็จ: {copied} ไฟล์\n"
            f"ผิดพลาด: {errors} ไฟล์\n\n"
            f"เก็บไว้ที่: {backup_folder}",
        )
        self.status_var.set(f"Backup เสร็จ - {backup_folder}")


def main():
    root = tk.Tk()
    CatalogManager(root)
    root.mainloop()


if __name__ == "__main__":
    main()
