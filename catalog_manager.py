"""
Thailand Trophy - Catalog File Manager
โปรแกรมจัดการไฟล์ Catalog สำหรับ Thailand Trophy

ความสามารถ:
1. ดูรายการไฟล์ทั้งหมดใน folder
2. จัดเรียง/แยกหมวดไฟล์ (รูปภาพ, เอกสาร, อื่นๆ)
3. ค้นหาไฟล์ตามชื่อ
4. สำรองข้อมูล (Backup)
5. เปิดไฟล์ได้โดยดับเบิ้ลคลิก
6. คลิกขวา: เปิด / เปลี่ยนชื่อ / ลบ / Copy path
7. Progress bar สำหรับงานที่ใช้เวลานาน
"""

import os
import shutil
import subprocess
import platform
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
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


def open_file(filepath):
    """เปิดไฟล์ด้วยโปรแกรมที่ติดตั้งไว้บนเครื่อง"""
    try:
        if platform.system() == "Windows":
            os.startfile(filepath)
        elif platform.system() == "Darwin":
            subprocess.Popen(["open", filepath])
        else:
            subprocess.Popen(["xdg-open", filepath])
    except Exception as e:
        messagebox.showerror("เปิดไฟล์ไม่ได้", f"ไม่สามารถเปิดไฟล์:\n{filepath}\n\n{e}")


def open_folder_in_explorer(folder_path):
    """เปิด folder ใน File Explorer"""
    try:
        if platform.system() == "Windows":
            os.startfile(folder_path)
        elif platform.system() == "Darwin":
            subprocess.Popen(["open", folder_path])
        else:
            subprocess.Popen(["xdg-open", folder_path])
    except Exception as e:
        messagebox.showerror("เปิด folder ไม่ได้", f"{e}")


class CatalogManager:
    def __init__(self, root):
        self.root = root
        self.root.title("Thailand Trophy - Catalog Manager")
        self.root.geometry("950x700")
        self.root.resizable(True, True)

        self.folder_path = DEFAULT_FOLDER
        self.all_files = []
        self.sort_reverse = {}  # ติดตามทิศทางการเรียงแต่ละคอลัมน์

        self.setup_ui()
        self.setup_context_menu()
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
        ttk.Button(top_frame, text="เปิดใน Explorer", command=self.open_current_folder).pack(side="left", padx=2)
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

        # ===== Progress Bar (ซ่อนไว้ แสดงเมื่อทำงาน) =====
        self.progress_frame = ttk.Frame(self.root)
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_label = tk.StringVar(value="")
        ttk.Label(self.progress_frame, textvariable=self.progress_label).pack(side="left", padx=(0, 10))
        self.progress_bar = ttk.Progressbar(
            self.progress_frame, variable=self.progress_var, maximum=100, length=400
        )
        self.progress_bar.pack(side="left", fill="x", expand=True)

        # ===== ส่วนตารางแสดงไฟล์ =====
        table_frame = ttk.Frame(self.root)
        table_frame.pack(fill="both", expand=True, padx=10, pady=5)

        columns = ("name", "category", "size", "modified")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=20)

        self.tree.heading("name", text="ชื่อไฟล์", command=lambda: self.sort_column("name"))
        self.tree.heading("category", text="ประเภท", command=lambda: self.sort_column("category"))
        self.tree.heading("size", text="ขนาด", command=lambda: self.sort_column("size"))
        self.tree.heading("modified", text="วันที่แก้ไข", command=lambda: self.sort_column("modified"))

        self.tree.column("name", width=400)
        self.tree.column("category", width=100)
        self.tree.column("size", width=100)
        self.tree.column("modified", width=150)

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # ===== Bind events =====
        self.tree.bind("<Double-1>", self.on_double_click)

        # ===== ส่วนล่าง: สถานะ =====
        self.status_var = tk.StringVar(value="พร้อมใช้งาน  |  ดับเบิ้ลคลิกเพื่อเปิดไฟล์  |  คลิกขวาสำหรับตัวเลือกเพิ่มเติม")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief="sunken", padding=5)
        status_bar.pack(fill="x", padx=10, pady=5)

    def setup_context_menu(self):
        """สร้าง Right-click context menu"""
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="เปิดไฟล์", command=self.ctx_open_file)
        self.context_menu.add_command(label="เปิด Folder ที่อยู่", command=self.ctx_open_containing_folder)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="เปลี่ยนชื่อ...", command=self.ctx_rename_file)
        self.context_menu.add_command(label="Copy Path", command=self.ctx_copy_path)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="ลบไฟล์", command=self.ctx_delete_file)

        self.tree.bind("<Button-3>", self.on_right_click)

    # ===== Helper: หาข้อมูลไฟล์จากแถวที่เลือก =====

    def get_selected_file(self):
        """ดึงข้อมูลไฟล์จากแถวที่เลือกใน treeview"""
        selection = self.tree.selection()
        if not selection:
            return None
        item = self.tree.item(selection[0])
        relative_path = item["values"][0]
        for f in self.all_files:
            if f["relative"] == relative_path:
                return f
        return None

    # ===== Double-click & Right-click handlers =====

    def on_double_click(self, event):
        """ดับเบิ้ลคลิกเพื่อเปิดไฟล์"""
        f = self.get_selected_file()
        if f:
            open_file(f["path"])

    def on_right_click(self, event):
        """คลิกขวาเพื่อแสดง context menu"""
        row = self.tree.identify_row(event.y)
        if row:
            self.tree.selection_set(row)
            self.context_menu.post(event.x_root, event.y_root)

    def ctx_open_file(self):
        """Context menu: เปิดไฟล์"""
        f = self.get_selected_file()
        if f:
            open_file(f["path"])

    def ctx_open_containing_folder(self):
        """Context menu: เปิด folder ที่ไฟล์อยู่"""
        f = self.get_selected_file()
        if f:
            folder = os.path.dirname(f["path"])
            open_folder_in_explorer(folder)

    def ctx_rename_file(self):
        """Context menu: เปลี่ยนชื่อไฟล์"""
        f = self.get_selected_file()
        if not f:
            return

        new_name = simpledialog.askstring(
            "เปลี่ยนชื่อไฟล์",
            f"ชื่อเดิม: {f['name']}\n\nใส่ชื่อใหม่:",
            initialvalue=f["name"],
            parent=self.root,
        )

        if not new_name or new_name == f["name"]:
            return

        new_path = os.path.join(os.path.dirname(f["path"]), new_name)

        if os.path.exists(new_path):
            messagebox.showwarning("ชื่อซ้ำ", f"มีไฟล์ชื่อ '{new_name}' อยู่แล้ว")
            return

        try:
            os.rename(f["path"], new_path)
            self.status_var.set(f"เปลี่ยนชื่อสำเร็จ: {f['name']} -> {new_name}")
            self.load_files()
        except Exception as e:
            messagebox.showerror("เปลี่ยนชื่อไม่ได้", f"{e}")

    def ctx_copy_path(self):
        """Context menu: Copy path ไปยัง clipboard"""
        f = self.get_selected_file()
        if f:
            self.root.clipboard_clear()
            self.root.clipboard_append(f["path"])
            self.status_var.set(f"Copy path แล้ว: {f['path']}")

    def ctx_delete_file(self):
        """Context menu: ลบไฟล์"""
        f = self.get_selected_file()
        if not f:
            return

        result = messagebox.askyesno(
            "ยืนยันการลบ",
            f"ต้องการลบไฟล์นี้หรือไม่?\n\n{f['name']}\n\n"
            f"ขนาด: {format_size(f['size'])}\n"
            f"ที่อยู่: {f['path']}",
        )
        if not result:
            return

        try:
            os.remove(f["path"])
            self.status_var.set(f"ลบสำเร็จ: {f['name']}")
            self.load_files()
        except Exception as e:
            messagebox.showerror("ลบไม่ได้", f"{e}")

    # ===== Progress Bar =====

    def show_progress(self, label="กำลังทำงาน..."):
        """แสดง progress bar"""
        self.progress_var.set(0)
        self.progress_label.set(label)
        self.progress_frame.pack(fill="x", padx=10, pady=2, before=self.tree.master)
        self.root.update()

    def update_progress(self, current, total):
        """อัปเดต progress bar"""
        if total > 0:
            pct = (current / total) * 100
            self.progress_var.set(pct)
            self.progress_label.set(f"{current}/{total}")
            self.root.update()

    def hide_progress(self):
        """ซ่อน progress bar"""
        self.progress_frame.pack_forget()
        self.root.update()

    # ===== Folder & File operations =====

    def open_current_folder(self):
        """เปิด folder ปัจจุบันใน Explorer"""
        path = self.folder_var.get()
        if os.path.exists(path):
            open_folder_in_explorer(path)
        else:
            messagebox.showwarning("ไม่พบ Folder", f"ไม่พบ folder:\n{path}")

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
        """เรียงลำดับตามคอลัมน์ (สลับ ขึ้น/ลง)"""
        reverse = self.sort_reverse.get(col, False)
        items = [(self.tree.set(k, col), k) for k in self.tree.get_children()]
        items.sort(reverse=reverse)
        for index, (val, k) in enumerate(items):
            self.tree.move(k, "", index)
        self.sort_reverse[col] = not reverse

    def organize_files(self):
        """จัดแยกหมวดไฟล์เข้า sub-folder ตามประเภท"""
        if not self.all_files:
            messagebox.showwarning("แจ้งเตือน", "ไม่มีไฟล์ให้จัดเรียง\nกรุณาโหลดไฟล์ก่อน")
            return

        result = messagebox.askyesno(
            "ยืนยันการจัดแยกหมวด",
            f"จะจัดแยกไฟล์ {len(self.all_files)} ไฟล์ เข้า folder:\n\n"
            f"  รูปภาพ/\n"
            f"  เอกสาร/\n"
            f"  อื่นๆ/\n\n"
            f"ใน {self.folder_path}\n\n"
            f"ต้องการดำเนินการหรือไม่?",
        )
        if not result:
            return

        self.show_progress("กำลังจัดแยกหมวด...")
        moved = 0
        errors = 0
        total = len(self.all_files)

        for i, f in enumerate(self.all_files):
            category_folder = os.path.join(self.folder_path, f["category"])
            os.makedirs(category_folder, exist_ok=True)

            src = f["path"]
            dst = os.path.join(category_folder, f["name"])

            if os.path.dirname(f["path"]) == category_folder:
                self.update_progress(i + 1, total)
                continue

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

            self.update_progress(i + 1, total)

        self.hide_progress()

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

        self.show_progress("กำลังสำรองข้อมูล...")
        copied = 0
        errors = 0
        total = len(self.all_files)

        os.makedirs(backup_folder, exist_ok=True)
        for i, f in enumerate(self.all_files):
            try:
                rel_path = os.path.relpath(f["path"], self.folder_path)
                dst = os.path.join(backup_folder, rel_path)
                os.makedirs(os.path.dirname(dst), exist_ok=True)
                shutil.copy2(f["path"], dst)
                copied += 1
            except Exception:
                errors += 1
            self.update_progress(i + 1, total)

        self.hide_progress()

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
