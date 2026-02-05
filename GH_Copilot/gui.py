#!/usr/bin/env python3
"""
CCW 逐字稿分析工具 - GUI 版本
使用 Tkinter 提供圖形化界面
"""

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import tkinter.ttk as ttk
from pathlib import Path
from ccw_analyzer import CCWTranscriptAnalyzer


class CCWAnalyzerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("公督盟委員會逐字稿分析工具")
        self.root.geometry("900x700")
        self.root.resizable(True, True)
        
        self.analyzer = CCWTranscriptAnalyzer()
        self.current_speakers = None
        self.theme = "light"
        
        # 設定 ttk 主題與樣式，避免 macOS 預設 Aqua 導致按鈕全白
        self._configure_styles()
        
        self.setup_ui()

    def _configure_styles(self):
        """設定 ttk 主題與按鈕樣式，確保在 macOS 上顏色可生效。"""
        style = ttk.Style()
        # 優先使用可自訂顏色的主題
        try:
            style.theme_use("clam")
        except Exception:
            # 若無 clam，退回預設主題
            pass
        self.style = style
        self._apply_theme(self.theme)

    def _apply_theme(self, theme: str):
        """根據主題(light/dark)套用配色。"""
        self.theme = theme
        light_mode = theme == "light"
        bg = "#f7f8fa" if light_mode else "#1e1f22"
        fg = "#111111" if light_mode else "#e6e6e6"
        header_bg = "#f0f2f5" if light_mode else "#2b2d31"
        text_bg = "#ffffff" if light_mode else "#2b2d31"
        text_fg = fg

        self.root.configure(bg=bg)

        s = self.style
        # 全域基底
        s.configure("TFrame", background=bg)
        s.configure("TLabel", background=bg, foreground=fg)
        s.configure("Header.TFrame", background=header_bg)
        s.configure("Header.TLabel", background=header_bg, foreground=fg, font=("Arial", 12, "bold"))

        # 按鈕樣式
        s.configure("Primary.TButton", background="#4472C4", foreground="white", padding=6)
        s.map("Primary.TButton", background=[("active", "#3a64ad"), ("pressed", "#345a9b")], foreground=[("disabled", "#ddd")])

        s.configure("Success.TButton", background="#70AD47", foreground="white", padding=6)
        s.map("Success.TButton", background=[("active", "#61963e"), ("pressed", "#568637")], foreground=[("disabled", "#ddd")])

        s.configure("Danger.TButton", background="#FF6B6B", foreground="white", padding=6)
        s.map("Danger.TButton", background=[("active", "#e55f5f"), ("pressed", "#cc5454")], foreground=[("disabled", "#eee")])

        s.configure("Secondary.TButton", background="#999999", foreground="white", padding=6)
        s.map("Secondary.TButton", background=[("active", "#888888"), ("pressed", "#777777")], foreground=[("disabled", "#eee")])

        # 已建立的文字區同步背景/前景
        if hasattr(self, "text_input"):
            try:
                self.text_input.configure(background=text_bg, foreground=text_fg, insertbackground=text_fg)
            except Exception:
                pass
        if hasattr(self, "result_text"):
            try:
                self.result_text.configure(background=text_bg, foreground=text_fg, insertbackground=text_fg)
            except Exception:
                pass
    
    def setup_ui(self):
        """建立用戶界面"""
        # 頂部框架 - 委員名單設定
        top_frame = ttk.Frame(self.root, style="Header.TFrame", padding=(8, 8, 8, 8))
        top_frame.pack(fill=tk.X, padx=10, pady=(6, 0))
        
        ttk.Label(top_frame, text="委員名單:", style="Header.TLabel").pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            top_frame,
            text="📁 匯入名單 (txt)",
            command=self.import_speaker_list,
            style="Primary.TButton"
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            top_frame,
            text="🗑️ 清除限制",
            command=self.clear_speaker_list,
            style="Danger.TButton"
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            top_frame,
            text="📝 貼上名單",
            command=self.open_paste_speaker_window,
            style="Secondary.TButton"
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            top_frame,
            text="🌗 主題",
            command=self.toggle_theme,
            style="Secondary.TButton"
        ).pack(side=tk.LEFT, padx=5)
        
        self.speaker_list_label = ttk.Label(top_frame, text="未設定", style="Header.TLabel")
        self.speaker_list_label.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)
        
        # 分隔線
        tk.Frame(self.root, height=2, bg="#ddd").pack(fill=tk.X, padx=10, pady=5)
        
        # 主體框架
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 左側 - 檔案操作
        left_frame = ttk.Frame(main_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        
        ttk.Label(left_frame, text="逐字稿檔案", font=("Arial", 12, "bold")).pack(anchor=tk.W, pady=(0, 5))
        
        button_frame = ttk.Frame(left_frame)
        button_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(
            button_frame,
            text="📄 讀取 Word (.docx)",
            command=self.load_word_file,
            style="Success.TButton"
        ).pack(fill=tk.X, pady=3)
        
        ttk.Button(
            button_frame,
            text="📋 讀取文字檔 (.txt)",
            command=self.load_text_file,
            style="Success.TButton"
        ).pack(fill=tk.X, pady=3)
        
        # 文本輸入
        ttk.Label(left_frame, text="或直接貼上逐字稿:", font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(10, 5))
        
        self.text_input = scrolledtext.ScrolledText(left_frame, height=15, width=50, wrap=tk.WORD)
        self.text_input.pack(fill=tk.BOTH, expand=True, pady=5)
        # 調整文字區背景，避免整體視覺過白
        try:
            if self.theme == "light":
                self.text_input.configure(background="#ffffff", foreground="#111", insertbackground="#111")
            else:
                self.text_input.configure(background="#2b2d31", foreground="#e6e6e6", insertbackground="#e6e6e6")
        except Exception:
            pass
        
        ttk.Button(
            left_frame,
            text="✓ 解析文本",
            command=self.parse_text_input,
            style="Primary.TButton"
        ).pack(pady=5)
        
        # 右側 - 統計結果
        right_frame = ttk.Frame(main_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(10, 0))
        
        ttk.Label(right_frame, text="統計結果", font=("Arial", 12, "bold")).pack(anchor=tk.W, pady=(0, 5))
        
        self.result_text = scrolledtext.ScrolledText(right_frame, height=20, width=45, wrap=tk.WORD)
        self.result_text.pack(fill=tk.BOTH, expand=True, pady=5)
        try:
            if self.theme == "light":
                self.result_text.configure(background="#ffffff", foreground="#111", insertbackground="#111")
            else:
                self.result_text.configure(background="#2b2d31", foreground="#e6e6e6", insertbackground="#e6e6e6")
        except Exception:
            pass
        
        button_frame2 = ttk.Frame(right_frame)
        button_frame2.pack(fill=tk.X, pady=5)
        
        ttk.Button(
            button_frame2,
            text="🔍 查詢特定委員",
            command=self.search_speaker,
            style="Primary.TButton"
        ).pack(side=tk.LEFT, padx=3)
        
        ttk.Button(
            button_frame2,
            text="📊 輸出 Excel",
            command=self.export_to_excel,
            style="Success.TButton"
        ).pack(side=tk.LEFT, padx=3)
        
        ttk.Button(
            button_frame2,
            text="📑 多檔分表匯出",
            command=self.export_multiple_to_excel,
            style="Success.TButton"
        ).pack(side=tk.LEFT, padx=3)
        
        ttk.Button(
            button_frame2,
            text="🔄 刷新",
            command=self.refresh_summary,
            style="Secondary.TButton"
        ).pack(side=tk.LEFT, padx=3)
    
    def import_speaker_list(self):
        """匯入委員名單 txt 檔"""
        file_path = filedialog.askopenfilename(
            title="選擇委員名單檔案",
            filetypes=[("文字檔", "*.txt"), ("所有檔案", "*.*")]
        )
        
        if file_path:
            try:
                speakers = self.analyzer.load_speaker_list_from_file(file_path)
                self.current_speakers = speakers
                
                display_text = ", ".join(speakers)
                if len(display_text) > 50:
                    display_text = display_text[:47] + "..."
                
                self.speaker_list_label.config(text=f"✓ {len(speakers)} 位委員 | {display_text}")
                messagebox.showinfo("成功", f"已載入 {len(speakers)} 位委員")
                
            except Exception as e:
                messagebox.showerror("錯誤", f"載入失敗: {e}")
    
    def clear_speaker_list(self):
        """清除委員限制"""
        self.analyzer.clear_speaker_list()
        self.current_speakers = None
        self.speaker_list_label.config(text="未設定")
        messagebox.showinfo("成功", "已清除委員限制，將抓取所有發言人")
    
    def load_word_file(self):
        """讀取 Word 檔案"""
        file_path = filedialog.askopenfilename(
            title="選擇逐字稿檔案",
            filetypes=[("Word 檔", "*.docx"), ("所有檔案", "*.*")]
        )
        
        if file_path:
            try:
                # 清空之前的數據，確保每個檔案獨立計算
                self.analyzer.clear_data()
                self.analyzer.load_from_word(file_path)
                self.refresh_summary()
                messagebox.showinfo("成功", f"已讀取檔案")
            except Exception as e:
                messagebox.showerror("錯誤", f"讀取失敗: {e}")
    
    def load_text_file(self):
        """讀取文字檔案"""
        file_path = filedialog.askopenfilename(
            title="選擇逐字稿檔案",
            filetypes=[("文字檔", "*.txt"), ("所有檔案", "*.*")]
        )
        
        if file_path:
            try:
                # 清空之前的數據
                self.analyzer.clear_data()
                with open(file_path, 'r', encoding='utf-8') as f:
                    text = f.read()
                self.text_input.delete(1.0, tk.END)
                self.text_input.insert(1.0, text)
                messagebox.showinfo("成功", "已讀取檔案")
            except Exception as e:
                messagebox.showerror("錯誤", f"讀取失敗: {e}")
    
    def parse_text_input(self):
        """解析輸入的文本"""
        text = self.text_input.get(1.0, tk.END)
        
        if not text.strip():
            messagebox.showwarning("警告", "請輸入逐字稿文本")
            return
        
        try:
            # 清空之前的數據
            self.analyzer.clear_data()
            self.analyzer.parse_transcript(text)
            self.refresh_summary()
            messagebox.showinfo("成功", f"已解析 {len(self.analyzer.get_all_speakers())} 位發言人")
        except Exception as e:
            messagebox.showerror("錯誤", f"解析失敗: {e}")
    
    def refresh_summary(self):
        """刷新統計結果"""
        self.result_text.delete(1.0, tk.END)
        
        speakers = self.analyzer.get_all_speakers()
        if not speakers:
            self.result_text.insert(tk.END, "尚未載入任何逐字稿")
            return
        
        # 標題
        result = "【所有委員發言統計】\n"
        result += "-" * 40 + "\n"
        result += f"{'委員名字':<15} {'發言次數':>10} {'總字數':>10}\n"
        result += "-" * 40 + "\n"
        
        # 統計數據
        total_count = 0
        total_chars = 0
        
        for speaker in sorted(speakers):
            data = self.analyzer.speaker_data[speaker]
            count = data["count"]
            chars = data["total_chars"]
            total_count += count
            total_chars += chars
            
            result += f"{speaker:<15} {count:>10} {chars:>10}\n"
        
        # 總計
        result += "-" * 40 + "\n"
        result += f"{'總計':<15} {total_count:>10} {total_chars:>10}\n"
        
        self.result_text.insert(tk.END, result)
    
    def search_speaker(self):
        """查詢特定委員"""
        speakers = self.analyzer.get_all_speakers()
        
        if not speakers:
            messagebox.showwarning("警告", "尚未載入任何逐字稿")
            return
        
        # 建立查詢窗口
        search_window = tk.Toplevel(self.root)
        search_window.title("查詢委員")
        search_window.geometry("400x300")
        try:
            search_window.configure(bg="#f7f8fa")
        except Exception:
            pass
        
        ttk.Label(search_window, text="選擇委員:", font=("Arial", 11, "bold")).pack(pady=10)
        
        # 委員列表
        listbox = tk.Listbox(search_window, font=("Arial", 10), height=10)
        listbox.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        for speaker in sorted(speakers):
            listbox.insert(tk.END, speaker)
        
        def show_detail():
            selection = listbox.curselection()
            if not selection:
                messagebox.showwarning("警告", "請選擇一位委員")
                return
            
            speaker_name = listbox.get(selection[0])
            data = self.analyzer.analyze_speaker(speaker_name)
            
            # 顯示詳細結果
            detail_window = tk.Toplevel(search_window)
            detail_window.title(f"{speaker_name} 的發言統計")
            detail_window.geometry("600x500")
            try:
                detail_window.configure(bg="#f7f8fa")
            except Exception:
                pass
            
            # 統計信息
            info_text = f"委員: {data['name']}\n"
            info_text += f"發言次數: {data['count']}\n"
            info_text += f"總字數: {data['total_chars']}\n"
            info_text += f"平均字數: {data['total_chars'] / data['count']:.1f}\n\n"
            info_text += "=" * 50 + "\n"
            info_text += "發言內容:\n"
            info_text += "=" * 50 + "\n\n"
            
            for i, speech in enumerate(data['speeches'], 1):
                info_text += f"{i}. {speech}\n\n"
            
            detail_display = scrolledtext.ScrolledText(detail_window, wrap=tk.WORD, font=("Arial", 10))
            detail_display.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            detail_display.insert(1.0, info_text)
            detail_display.config(state=tk.DISABLED)
        
        ttk.Button(
            search_window,
            text="查看詳情",
            command=show_detail,
            style="Primary.TButton"
        ).pack(pady=10)

    def open_paste_speaker_window(self):
        """開啟視窗，讓使用者直接貼上委員名單（每行一人）。"""
        win = tk.Toplevel(self.root)
        win.title("貼上委員名單")
        win.geometry("420x360")
        try:
            # 套用目前主題背景
            if self.theme == "light":
                win.configure(bg="#f7f8fa")
            else:
                win.configure(bg="#1e1f22")
        except Exception:
            pass

        ttk.Label(win, text="每行一個名字，例如：\n陳委員培瑜\n王委員義川", font=("Arial", 10)).pack(padx=10, pady=8, anchor=tk.W)
        text = scrolledtext.ScrolledText(win, height=12, width=46, wrap=tk.WORD)
        text.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        try:
            if self.theme == "light":
                text.configure(background="#ffffff", foreground="#111", insertbackground="#111")
            else:
                text.configure(background="#2b2d31", foreground="#e6e6e6", insertbackground="#e6e6e6")
        except Exception:
            pass

        btn_row = ttk.Frame(win)
        btn_row.pack(fill=tk.X, padx=10, pady=8)

        def apply():
            content = text.get(1.0, tk.END)
            speakers = [line.strip() for line in content.splitlines() if line.strip()]
            if not speakers:
                messagebox.showwarning("警告", "請輸入至少一位委員名字")
                return
            self.analyzer.set_speaker_list(speakers)
            self.current_speakers = speakers
            display_text = ", ".join(speakers)
            if len(display_text) > 50:
                display_text = display_text[:47] + "..."
            self.speaker_list_label.config(text=f"✓ {len(speakers)} 位委員 | {display_text}")
            messagebox.showinfo("成功", f"已設定 {len(speakers)} 位委員")
            win.destroy()

        ttk.Button(btn_row, text="套用", command=apply, style="Success.TButton").pack(side=tk.RIGHT, padx=5)
        ttk.Button(btn_row, text="取消", command=win.destroy, style="Secondary.TButton").pack(side=tk.RIGHT)

    def toggle_theme(self):
        """在淺色與深色主題間切換。"""
        new_theme = "dark" if self.theme == "light" else "light"
        self._apply_theme(new_theme)
    
    def export_to_excel(self):
        """輸出到 Excel"""
        if not self.analyzer.get_all_speakers():
            messagebox.showwarning("警告", "尚未載入任何逐字稿")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="另存為",
            defaultextension=".xlsx",
            filetypes=[("Excel 檔", "*.xlsx"), ("所有檔案", "*.*")]
        )
        
        if file_path:
            try:
                self.analyzer.export_to_excel(file_path)
                messagebox.showinfo("成功", f"已輸出到: {Path(file_path).name}")
            except Exception as e:
                messagebox.showerror("錯誤", f"輸出失敗: {e}")

    def export_multiple_to_excel(self):
        """選擇多個檔案，分工作表匯出到同一個 Excel。"""
        file_paths = filedialog.askopenfilenames(
            title="選擇逐字稿檔案 (可多選)",
            filetypes=[("Word/文字檔", "*.docx *.txt"), ("Word 檔", "*.docx"), ("文字檔", "*.txt"), ("所有檔案", "*.*")]
        )

        if not file_paths:
            return

        output_path = filedialog.asksaveasfilename(
            title="另存為",
            defaultextension=".xlsx",
            filetypes=[("Excel 檔", "*.xlsx"), ("所有檔案", "*.*")]
        )

        if not output_path:
            return

        try:
            self.analyzer.export_multiple_files_to_excel(file_paths, output_path)
            messagebox.showinfo("成功", f"已匯出 {len(file_paths)} 個檔案到: {Path(output_path).name}")
        except Exception as e:
            messagebox.showerror("錯誤", f"匯出失敗: {e}")


def main():
    """主程式"""
    root = tk.Tk()
    app = CCWAnalyzerGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
