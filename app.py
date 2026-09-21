"""
app.py
QuickClaim 可视化操作小窗口
"""

import os
import shutil
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
import webbrowser

from extractor import extract_ticket
from excel_builder import build_expense_excel

APP_VERSION = "v1.1.0"
AUTHOR_NAME = "kingcat"
AUTHOR_X_URL = "https://x.com/kingcat"
SPONSOR_URL = "https://kingcat.com/sponsor"


def sanitize_str(s: str) -> str:
    """清理文件名中 Windows 不允许的特殊字符"""
    for char in ['\\', '/', ':', '*', '?', '"', '<', '>', '|', ' ']:
        s = s.replace(char, '')
    return s.strip('_')


def detect_claim_persons(target_path: Path, default_user: str) -> list[str]:
    system_ignore = {"高铁", "打车", "发票", "住宿", "新建文件夹", "pdf", "差旅报销"}
    sub_dirs = [
        d for d in target_path.iterdir()
        if d.is_dir() and d.name not in system_ignore and not d.name.endswith("差旅报销")
    ]

    person_names = []
    for d in sub_dirs:
        if any(f.suffix.lower() == ".pdf" for f in d.rglob("*.pdf")):
            person_names.append(d.name)

    if person_names:
        return sorted(person_names)
    return [default_user]


def resolve_project_name(target_path: Path) -> str:
    generic_keywords = {
        "发票", "票据", "报销", "差旅", "新建文件夹", 
        "pdf", "新建", "feiyong", "fp", "files", "ticket", "invoice"
    }
    curr = target_path.resolve()
    if curr.name.endswith("差旅报销") or curr.name.endswith("报销"):
        return curr.name.replace("差旅报销", "").replace("报销", "")
    
    temp_dir = curr
    while temp_dir and temp_dir.parent != temp_dir:
        name_clean = temp_dir.name.lower().strip()
        if name_clean not in generic_keywords and len(name_clean) > 2 and not name_clean.endswith(":"):
            return temp_dir.name
        temp_dir = temp_dir.parent
    return curr.name


class QuickClaimApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"QuickClaim - 差旅报销助手 {APP_VERSION}")
        self.root.geometry("680x780")
        self.root.resizable(False, False)

        self.folder_var = tk.StringVar(value="")
        self.project_var = tk.StringVar(value="")
        self.name_var = tk.StringVar(value="kingcat")
        self.meal_var = tk.StringVar(value="0.0")
        self.prefix_var = tk.StringVar(value="报销")
        self.suffix_var = tk.StringVar(value="")
        self.preview_train_var = tk.StringVar(value="")
        self.preview_taxi_var = tk.StringVar(value="")

        self.detected_persons = [self.name_var.get()]

        self.prefix_var.trace_add("write", self._update_preview)
        self.suffix_var.trace_add("write", self._update_preview)
        self.folder_var.trace_add("write", self._update_preview)

        self._build_menu()
        self._build_ui()
        self._update_preview()

    def _build_menu(self):
        menubar = tk.Menu(self.root)
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label=f"关于助手 ({APP_VERSION})", command=self._show_about)
        help_menu.add_command(label="☕ 赞赏作者", command=self._show_sponsor)
        help_menu.add_separator()
        help_menu.add_command(label="访问作者主页", command=lambda: webbrowser.open(AUTHOR_X_URL))
        menubar.add_cascade(label="帮助", menu=help_menu)
        self.root.config(menu=menubar)

    def _show_about(self):
        info = (
            f"QuickClaim 差旅报销助手\n"
            f"版本: {APP_VERSION}\n"
            f"作者: @{AUTHOR_NAME}\n\n"
            f"主要功能:\n"
            f"• 支持项目多人员独立文件夹自动识别与分类汇总\n"
            f"• 12306 铁路数电客票与网约车电子行程单解析\n"
            f"• 个人打车行程单自动优先核算，明细自动标注(有票)\n"
            f"• 左侧大单元格合并、连续顺号、个人小计、全局大计\n"
        )
        messagebox.showinfo("关于 QuickClaim", info)

    def _show_sponsor(self):
        top = tk.Toplevel(self.root)
        top.title("☕ 赞赏与支持")
        top.geometry("380x250")
        top.resizable(False, False)
        top.transient(self.root)
        top.grab_set()

        frame = ttk.Frame(top, padding=20)
        frame.pack(fill='both', expand=True)

        ttk.Label(frame, text="感谢使用 QuickClaim 差旅报销助手！", font=("微软雅黑", 10, "bold")).pack(pady=(0, 8))
        ttk.Label(
            frame, 
            text="如果这个小工具帮您节省了整理发票和做报销单的时间，\n欢迎请作者喝杯咖啡 ☕ 持续支持后续更新维护！",
            justify="center",
            font=("微软雅黑", 9)
        ).pack(pady=(0, 16))

        btn_pay = tk.Button(
            frame,
            text="❤ 打开赞赏主页 (微信 / 支付宝)",
            bg="#E74C3C",
            fg="white",
            font=("微软雅黑", 9, "bold"),
            relief="flat",
            cursor="hand2",
            command=lambda: webbrowser.open(SPONSOR_URL)
        )
        btn_pay.pack(fill='x', ipady=6, pady=(0, 10))
        ttk.Button(frame, text="关闭", command=top.destroy).pack()

    def _build_ui(self):
        pad = {'padx': 16, 'pady': 5}

        # Header: 版本号与赞赏
        f_top_header = ttk.Frame(self.root)
        f_top_header.pack(fill='x', padx=16, pady=(10, 4))

        ttk.Label(f_top_header, text="QuickClaim 差旅报销助手", font=("微软雅黑", 11, "bold")).pack(side='left')
        lbl_v = tk.Label(f_top_header, text=f" {APP_VERSION} ", bg="#EAECEE", fg="#5D6D7E", font=("Consolas", 9, "bold"), relief="solid", bd=1)
        lbl_v.pack(side='left', padx=(8, 0))

        btn_sponsor = tk.Button(
            f_top_header, text="☕ 赞赏作者", bg="#FFF0F0", fg="#E74C3C",
            activebackground="#FADBD8", activeforeground="#C0392B",
            font=("微软雅黑", 9, "bold"), relief="solid", bd=1, cursor="hand2",
            padx=8, pady=1, command=self._show_sponsor
        )
        btn_sponsor.pack(side='right')

        # 1. 票据目录与项目名称
        frame_path = ttk.LabelFrame(self.root, text=" 票据目录与项目名称 ", padding=8)
        frame_path.pack(fill='x', **pad)

        f_dir = ttk.Frame(frame_path)
        f_dir.pack(fill='x')
        ttk.Entry(f_dir, textvariable=self.folder_var, font=("微软雅黑", 9)).pack(side='left', fill='x', expand=True, padx=(0, 8))
        ttk.Button(f_dir, text="浏览文件夹...", command=self._select_dir).pack(side='right')

        f_proj = ttk.Frame(frame_path)
        f_proj.pack(fill='x', pady=(8, 0))
        ttk.Label(f_proj, text="项目名称:").pack(side='left')
        ttk.Entry(f_proj, textvariable=self.project_var, font=("微软雅黑", 9, "bold")).pack(side='left', fill='x', expand=True, padx=(8, 0))

        # 2. 报销人员与餐补
        frame_params = ttk.LabelFrame(self.root, text=" 报销人员与餐补设置 ", padding=8)
        frame_params.pack(fill='x', **pad)

        f_row = ttk.Frame(frame_params)
        f_row.pack(fill='x', pady=(0, 4))
        ttk.Label(f_row, text="默认姓名:").pack(side='left')
        self.entry_name = ttk.Entry(f_row, textvariable=self.name_var, width=12)
        self.entry_name.pack(side='left', padx=(6, 24))

        ttk.Label(f_row, text="餐补定额:").pack(side='left')
        ttk.Entry(f_row, textvariable=self.meal_var, width=8).pack(side='left', padx=(6, 4))
        ttk.Label(f_row, text="元 / 天").pack(side='left')

        f_persons = ttk.Frame(frame_params)
        f_persons.pack(fill='x', pady=(4, 0))
        ttk.Label(f_persons, text="识别人员:", width=9).pack(side='left', anchor='n', pady=2)
        self.lst_persons = tk.Listbox(f_persons, height=3, font=("微软雅黑", 9), relief="solid", bd=1)
        self.lst_persons.pack(side='left', fill='x', expand=True)
        self.lst_persons.insert(tk.END, self.name_var.get())

        # 3. 命名自定义
        frame_naming = ttk.LabelFrame(self.root, text=" 票据文件重命名自定义 (前后缀自由选填) ", padding=8)
        frame_naming.pack(fill='x', **pad)

        f_inputs = ttk.Frame(frame_naming)
        f_inputs.pack(fill='x', pady=(0, 6))
        ttk.Label(f_inputs, text="前缀:").pack(side='left')
        ttk.Entry(f_inputs, textvariable=self.prefix_var, width=14).pack(side='left', padx=(4, 8))
        ttk.Label(f_inputs, text="＋ [时间+类别+明细+金额] ＋", foreground="#0066CC", font=("微软雅黑", 9, "bold")).pack(side='left', padx=(0, 8))
        ttk.Label(f_inputs, text="后缀:").pack(side='left')
        ttk.Entry(f_inputs, textvariable=self.suffix_var, width=14).pack(side='left', padx=(4, 0))

        f_prev = ttk.Frame(frame_naming)
        f_prev.pack(fill='x', pady=(4, 0))
        row_prev1 = ttk.Frame(f_prev)
        row_prev1.pack(fill='x', pady=2)
        ttk.Label(row_prev1, text="高铁预览:", foreground="#666666", width=9).pack(side='left')
        ttk.Label(row_prev1, textvariable=self.preview_train_var, foreground="#D9534F", font=("Consolas", 9, "bold")).pack(side='left')

        row_prev2 = ttk.Frame(f_prev)
        row_prev2.pack(fill='x', pady=2)
        ttk.Label(row_prev2, text="打车预览:", foreground="#666666", width=9).pack(side='left')
        ttk.Label(row_prev2, textvariable=self.preview_taxi_var, foreground="#008080", font=("Consolas", 9, "bold")).pack(side='left')

        # 4. 执行按钮
        self.btn_run = tk.Button(
            self.root, text="▶ 运行统计并一键归档", bg="#0066CC", fg="white",
            font=("微软雅黑", 10, "bold"), height=2, command=self._start_task
        )
        self.btn_run.pack(fill='x', padx=16, pady=8)

        # 5. 状态日志
        frame_log = ttk.LabelFrame(self.root, text=" 运行过程与状态监控 ", padding=8)
        frame_log.pack(fill='both', expand=True, **pad)
        self.txt_log = tk.Text(frame_log, bg="#1E1E1E", fg="#D4D4D4", font=("Consolas", 9), wrap='word')
        self.txt_log.pack(fill='both', expand=True)

        # 6. 底部信息
        frame_footer = ttk.Frame(self.root)
        frame_footer.pack(fill='x', padx=16, pady=(0, 6))
        ttk.Label(frame_footer, text=f"{APP_VERSION} | 差旅发票自动化整理工具", foreground="#888888").pack(side='left')
        link_author = tk.Label(frame_footer, text=f"联系作者: @{AUTHOR_NAME}", fg="#0066CC", cursor="hand2")
        link_author.pack(side='right')
        link_author.bind("<Button-1>", lambda e: webbrowser.open(AUTHOR_X_URL))

    def _assemble_name(self, prefix: str, core_segment: str, suffix: str) -> str:
        parts = []
        if prefix:
            parts.append(prefix)
        parts.append(core_segment)
        if suffix:
            parts.append(suffix)
        return "_".join(parts) + ".pdf"

    def _update_preview(self, *args):
        p = sanitize_str(self.prefix_var.get())
        s = sanitize_str(self.suffix_var.get())
        train_core = "2026-08-27_高铁_连云港-南京南_158.00元"
        self.preview_train_var.set(self._assemble_name(p, train_core, s))
        taxi_core = "2026-08-29_打车_高德打车(有票)_68.75元"
        self.preview_taxi_var.set(self._assemble_name(p, taxi_core, s))

    def _select_dir(self):
        selected = filedialog.askdirectory(title="选择包含发票的项目目录")
        if not selected:
            return

        sel_path = Path(selected)
        self.folder_var.set(str(sel_path))
        self._log(f"\n[*] 已选定目录: {sel_path}")

        proj = resolve_project_name(sel_path)
        self.project_var.set(proj)
        self._log(f"[*] 自动提取项目名称: 【{proj}】")

        curr_user = self.name_var.get().strip() or "kingcat"
        self.detected_persons = detect_claim_persons(sel_path, curr_user)

        self.lst_persons.delete(0, tk.END)
        for p in self.detected_persons:
            self.lst_persons.insert(tk.END, p)

        if len(self.detected_persons) > 1:
            self.entry_name.config(state='disabled')
            self._log(f"[*] 检测到多人子文件夹，共 {len(self.detected_persons)} 人: {', '.join(self.detected_persons)}")
        else:
            self.entry_name.config(state='normal')
            self._log(f"[*] 处于单人模式: 报销人【{self.detected_persons[0]}】")

    def _log(self, text: str):
        # 使用 after 保证多线程写 UI 绝对安全
        self.root.after(0, self._append_log, text)

    def _append_log(self, text: str):
        self.txt_log.insert(tk.END, text + "\n")
        self.txt_log.see(tk.END)

    def _start_task(self):
        path = self.folder_var.get().strip()
        if not path or not Path(path).exists():
            messagebox.showwarning("提示", "请先选择有效的票据目录！")
            return

        self.btn_run.config(state='disabled', text="正在处理中...")
        threading.Thread(target=self._run_claim, daemon=True).start()

    def _run_claim(self):
        try:
            root_dir = Path(self.folder_var.get().strip()).resolve()
            proj_name = self.project_var.get().strip()
            folder_proj_title = proj_name if proj_name else "差旅报销"

            try:
                daily_meal = float(self.meal_var.get().strip())
            except ValueError:
                daily_meal = 0.0

            # 统一在父级建立清晰的报销汇总目录
            workspace = root_dir.parent / f"{folder_proj_title}差旅报销"
            workspace.mkdir(exist_ok=True)

            self._log(f"\n================ 开始差旅核算与归档 ================")
            self._log(f"[*] 项目表头: 【{proj_name}】 | 餐补定额: ¥{daily_meal}/天")
            self._log(f"[*] 归档目标目录: {workspace}")

            user_prefix = sanitize_str(self.prefix_var.get())
            user_suffix = sanitize_str(self.suffix_var.get())

            all_final_tickets = []
            is_multi_mode = len(self.detected_persons) > 1

            for person in self.detected_persons:
                self._log(f"\n---> 正在处理人员: 【{person}】")
                person_source_dir = root_dir / person if is_multi_mode else root_dir
                
                pdf_files = [
                    f for f in person_source_dir.rglob("*.pdf")
                    if not f.name.startswith("~") and "差旅报销" not in f.as_posix()
                ]

                if not pdf_files:
                    self._log(f"[!] 未在【{person}】目录下发现 PDF 文件，跳过。")
                    continue

                parsed_results = []
                for f in pdf_files:
                    try:
                        info = extract_ticket(str(f))
                        if info.get("valid"):
                            info["person"] = person
                            info["orig_file"] = f
                            parsed_results.append(info)
                            self._log(f"  [✓ 识别有效] {info.get('type')} | {info.get('date')} | ¥{info.get('amount')} ({f.name})")
                        else:
                            self._log(f"  [✕ 非有效票据-保持原地不动] {f.name}")
                    except Exception as e:
                        self._log(f"  [!] 文件解析异常: {f.name} ({e})")

                if not parsed_results:
                    continue

                # 归档保留人员文件夹
                person_archive_dir = workspace / person if is_multi_mode else workspace
                for item in parsed_results:
                    src_file = Path(item["orig_file"])
                    c_type = item.get("type", "其它")
                    c_date = item.get("date", "未知日期")
                    c_desc = sanitize_str(item.get("desc", "费用"))
                    if c_type == "打车" and "(有票)" not in c_desc:
                        c_desc += "(有票)"
                    c_amt = f"{float(item.get('amount', 0.0)):.2f}"

                    core_segment = f"{c_date}_{c_type}_{c_desc}_{c_amt}元"
                    new_name = self._assemble_name(user_prefix, core_segment, user_suffix)

                    type_dir = person_archive_dir / c_type
                    type_dir.mkdir(parents=True, exist_ok=True)

                    dest_file = type_dir / new_name
                    idx = 1
                    base_stem = dest_file.stem
                    while dest_file.exists() and dest_file != src_file:
                        dest_file = type_dir / f"{base_stem}_{idx}.pdf"
                        idx += 1

                    if src_file.exists() and src_file != dest_file:
                        moved = False
                        for _ in range(3):
                            try:
                                shutil.copy2(str(src_file), str(dest_file))
                                moved = True
                                break
                            except Exception:
                                time.sleep(0.3)
                        if moved:
                            self._log(f"  └─ 归档复制 -> {person}/{c_type}/{dest_file.name}")

                # 个人行程单与发票核验
                itineraries = [t for t in parsed_results if t.get("is_itinerary")]
                taxi_invoices = [t for t in parsed_results if t.get("type") == "打车" and not t.get("is_itinerary")]
                other_tickets = [t for t in parsed_results if t.get("type") != "打车"]

                person_final_tickets = []

                if itineraries:
                    itin_sum = sum(float(t.get("amount", 0.0)) for t in itineraries)
                    inv_sum = sum(float(t.get("amount", 0.0)) for t in taxi_invoices)
                    self._log(f"[*] 【{person}】行程单总额: ¥{itin_sum:.2f} | 发票总额: ¥{inv_sum:.2f}")

                    for itin in itineraries:
                        if itin.get("sub_items"):
                            for sub in itin["sub_items"]:
                                sub["person"] = person
                                if "(有票)" not in sub.get("desc", ""):
                                    sub["desc"] = f"{sub.get('desc', '高德打车')}(有票)"
                                person_final_tickets.append(sub)
                        else:
                            itin["person"] = person
                            if "(有票)" not in itin.get("desc", ""):
                                itin["desc"] = f"{itin.get('desc', '高德打车')}(有票)"
                            person_final_tickets.append(itin)
                else:
                    for inv in taxi_invoices:
                        inv["person"] = person
                        if "(有票)" not in inv.get("desc", ""):
                            inv["desc"] = f"{inv.get('desc', '高德打车')}(有票)"
                        person_final_tickets.append(inv)

                for oth in other_tickets:
                    oth["person"] = person
                    person_final_tickets.append(oth)

                all_final_tickets.extend(person_final_tickets)

            if not all_final_tickets:
                self._log("[!] 没有提取到任何有效票据，停止生成报销单。")
                return

            # 输出 Excel
            self._log("\n[*] 正在排版并生成最终 Excel 报销单...")
            excel_filename = f"{folder_proj_title}_差旅报销汇总表.xlsx"
            excel_path = workspace / excel_filename

            build_expense_excel(
                tickets=all_final_tickets,
                output_path=str(excel_path),
                project_name=proj_name,
                daily_meal_allowance=daily_meal
            )

            self._log(f"[★] 全流程处理完毕！报销单已保存至:\n    {excel_path}")
            messagebox.showinfo("完成", f"差旅报销汇总表已生成！\n\n项目目录: {workspace.name}\n请在文件夹内查看归档文件及 Excel。")

        except Exception as e:
            self._log(f"[异常崩溃] {str(e)}")
            messagebox.showerror("运行异常", str(e))
        finally:
            self.btn_run.config(state='normal', text="▶ 运行统计并一键归档")


if __name__ == '__main__':
    root = tk.Tk()
    app = QuickClaimApp(root)
    root.mainloop()