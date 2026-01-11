import os
import shutil
import ctypes
import sys
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from threading import Thread

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

class AppDataMover:
    def __init__(self, root):
        self.root = root
        self.root.title("Windows AppData 安全轉移與掃描工具")
        self.root.geometry("700x650")
        self.setup_ui()

    def setup_ui(self):
        # --- 掃描區塊 ---
        scan_frame = ttk.LabelFrame(self.root, text="第一步：掃描 AppData 大檔案", padding="10")
        scan_frame.pack(fill="x", padx=10, pady=5)

        self.scan_btn = ttk.Button(scan_frame, text="掃描 AppData 佔用空間", command=self.start_scan_thread)
        self.scan_btn.pack(side="top", pady=5)

        self.tree = ttk.Treeview(scan_frame, columns=("path", "size"), show="headings", height=8)
        self.tree.heading("path", text="資料夾路徑")
        self.tree.heading("size", text="大小 (MB)")
        self.tree.column("path", width=450)
        self.tree.column("size", width=100, anchor="center")
        self.tree.pack(fill="x", pady=5)
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)

        # --- 轉移區塊 ---
        move_frame = ttk.LabelFrame(self.root, text="第二步：設定轉移目標", padding="10")
        move_frame.pack(fill="x", padx=10, pady=5)

        ttk.Label(move_frame, text="來源路徑 (自動帶入或手動選擇):").pack(anchor="w")
        self.entry_src = ttk.Entry(move_frame, width=70)
        self.entry_src.pack(pady=2)
        ttk.Button(move_frame, text="瀏覽來源", command=self.browse_src).pack(anchor="e")

        # 使用 r"" 修正 SyntaxWarning
        ttk.Label(move_frame, text=r"目的根目錄 (例如 D:\AppData_Backup):").pack(anchor="w", pady=(10,0))
        self.entry_dst = ttk.Entry(move_frame, width=70)
        self.entry_dst.pack(pady=2)
        ttk.Button(move_frame, text="瀏覽目的地", command=self.browse_dst).pack(anchor="e")

        # --- 執行與狀態 ---
        self.status_label = ttk.Label(self.root, text="狀態：準備就緒", font=("Arial", 10, "bold"), foreground="blue")
        self.status_label.pack(pady=10)

        self.run_btn = ttk.Button(self.root, text="開始安全轉移與建立連結", command=self.move_and_link)
        self.run_btn.pack(pady=5, ipadx=20, ipady=5)

    # --- 功能邏輯 ---

    def get_size(self, start_path='.'):
        """計算資料夾總大小"""
        total_size = 0
        try:
            for dirpath, dirnames, filenames in os.walk(start_path):
                for f in filenames:
                    fp = os.path.join(dirpath, f)
                    if not os.path.islink(fp):
                        total_size += os.path.getsize(fp)
        except Exception: 
            pass # 修正之前的語法錯誤
        return total_size

    def start_scan_thread(self):
        self.scan_btn.config(state="disabled")
        self.status_label.config(text="狀態：正在掃描 AppData... 請稍候", foreground="orange")
        Thread(target=self.scan_appdata, daemon=True).start()

    def scan_appdata(self):
        # 獲取環境變數路徑
        appdata_roaming = os.environ.get('APPDATA') 
        appdata_local = os.environ.get('LOCALAPPDATA') 
        
        targets = []
        # 掃描 Local, Roaming, 以及 LocalLow
        scan_list = [appdata_roaming, appdata_local]
        # 嘗試抓取 LocalLow (通常在 Local 同級)
        locallow = os.path.join(os.path.dirname(appdata_local), "LocalLow")
        if os.path.exists(locallow):
            scan_list.append(locallow)

        for base in scan_list:
            if base and os.path.exists(base):
                try:
                    for folder in os.listdir(base):
                        full_path = os.path.join(base, folder)
                        if os.path.isdir(full_path) and not os.path.islink(full_path):
                            size_mb = self.get_size(full_path) / (1024 * 1024)
                            if size_mb > 50: # 降低門檻到 50MB 讓結果更明顯
                                targets.append((full_path, round(size_mb, 2)))
                except Exception:
                    continue
        
        targets.sort(key=lambda x: x[1], reverse=True)
        self.root.after(0, lambda: self.update_tree(targets))

    def update_tree(self, data):
        for i in self.tree.get_children():
            self.tree.delete(i)
        for item in data:
            self.tree.insert("", "end", values=item)
        self.scan_btn.config(state="normal")
        self.status_label.config(text="狀態：掃描完成", foreground="green")

    def on_tree_select(self, event):
        selection = self.tree.selection()
        if selection:
            selected_item = selection[0]
            path = self.tree.item(selected_item)['values'][0]
            self.entry_src.delete(0, tk.END)
            self.entry_src.insert(0, path)

    def browse_src(self):
        path = filedialog.askdirectory()
        if path:
            self.entry_src.delete(0, tk.END)
            self.entry_src.insert(0, path)

    def browse_dst(self):
        path = filedialog.askdirectory()
        if path:
            self.entry_dst.delete(0, tk.END)
            self.entry_dst.insert(0, path)

    def move_and_link(self):
        src = self.entry_src.get().replace("/", "\\")
        dst_parent = self.entry_dst.get().replace("/", "\\")
        
        if not src or not dst_parent:
            messagebox.showerror("錯誤", "請確認來源與目的路徑已填寫")
            return

        folder_name = os.path.basename(src)
        dst = os.path.join(dst_parent, folder_name)
        temp_backup = src + "_bak"

        if os.path.exists(dst):
            messagebox.showerror("錯誤", "目標位置已存在同名資料夾！")
            return

        try:
            self.status_label.config(text="狀態：正在複製檔案至新位置...", foreground="orange")
            self.root.update()
            shutil.copytree(src, dst)

            self.status_label.config(text="狀態：建立系統連結中...", foreground="orange")
            os.rename(src, temp_backup)
            
            # 使用 subprocess 執行 mklink
            cmd = f'mklink /d "{src}" "{dst}"'
            process = subprocess.run(cmd, shell=True, capture_output=True, text=True)

            if process.returncode == 0:
                # 再次確認連結是否成功建立
                if os.path.islink(src):
                    shutil.rmtree(temp_backup)
                    messagebox.showinfo("成功", f"轉移成功！\n資料夾已移至：{dst}")
                    self.status_label.config(text="狀態：轉移完成", foreground="green")
                else:
                    raise Exception("符號連結驗證失敗")
            else:
                raise Exception(process.stderr)

        except Exception as e:
            self.status_label.config(text="狀態：失敗，執行復原...", foreground="red")
            # 復原邏輯
            if os.path.exists(temp_backup):
                if os.path.exists(src): 
                    if os.path.islink(src): os.remove(src)
                    else: shutil.rmtree(src)
                os.rename(temp_backup, src)
            if os.path.exists(dst): 
                shutil.rmtree(dst)
            messagebox.showerror("失敗", f"出錯了：{e}\n已嘗試復原資料。")
if __name__ == "__main__": 
    if not is_admin():
        # 修正：增加引號包裹路徑，避免 OneDrive 的空白或特殊字元導致失敗
        script = os.path.abspath(__file__)
        params = f'"{script}"'
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, params, None, 1)
        sys.exit() # 確保舊進度完全關閉
    else:
        root = tk.Tk()
        # 強制將視窗移至最上層，避免它躲在後台
        root.lift()
        root.attributes('-topmost', True)
        root.after_set_interval = root.attributes('-topmost', False) 
        app = AppDataMover(root)
        root.mainloop()