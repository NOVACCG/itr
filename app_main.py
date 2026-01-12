"""Tabbed launcher for ITR tools."""

import os
import sys
import tkinter as tk
from tkinter import messagebox, ttk

from modules.itr_autofill_tab import APP_NAME, APP_VERSION, ITRAutofillTab
from modules.na_check_tab import NACheckTab

BASE_DIR = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
OUTPUT_ROOT = os.path.join(BASE_DIR, "output")
REPORT_ROOT = os.path.join(BASE_DIR, "report")


class ClosableNotebook(ttk.Notebook):
    def __init__(self, master: tk.Misc, **kwargs):
        super().__init__(master, **kwargs)
        self._img_close = tk.PhotoImage(
            "img_close",
            data=(
                "R0lGODlhDAAMAIABAAAAAP///yH5BAEKAAEALAAAAAAMAAwAAAIZjI+py+0Po5y02ouz3rz7D4bi"
                "SJbmiaaqKo1KZpV2JgA7"
            ),
        )
        self._img_close_active = tk.PhotoImage(
            "img_close_active",
            data=(
                "R0lGODlhDAAMAIABAP8AAP///yH5BAEKAAEALAAAAAAMAAwAAAIZjI+py+0Po5y02ouz3rz7D4biSJ"
                "bmiaaqKo1KZpV2JgA7"
            ),
        )
        self._img_close_pressed = tk.PhotoImage(
            "img_close_pressed",
            data=(
                "R0lGODlhDAAMAIABAOwAAP///yH5BAEKAAEALAAAAAAMAAwAAAIZjI+py+0Po5y02ouz3rz7D4biSJ"
                "bmiaaqKo1KZpV2JgA7"
            ),
        )
        self._style_name = "ClosableNotebook"
        style = ttk.Style(self)
        style.element_create(
            "close",
            "image",
            "img_close",
            ("active", "pressed", "img_close_pressed"),
            ("active", "img_close_active"),
            border=4,
            sticky="",
        )
        style.layout(
            f"{self._style_name}.Tab",
            [
                (
                    "Notebook.tab",
                    {
                        "sticky": "nswe",
                        "children": [
                            (
                                "Notebook.padding",
                                {
                                    "side": "top",
                                    "sticky": "nswe",
                                    "children": [
                                        ("Notebook.focus", {"side": "top", "sticky": "nswe", "children": [
                                            ("Notebook.label", {"side": "left", "sticky": ""}),
                                            ("close", {"side": "left", "sticky": ""}),
                                        ]}),
                                    ],
                                },
                            )
                        ],
                    },
                )
            ],
        )
        self.configure(style=self._style_name)
        self._closable_tabs: set[str] = set()
        self.bind("<ButtonPress-1>", self._on_close_press, True)
        self.bind("<ButtonRelease-1>", self._on_close_release, True)
        self._pressed_tab: str | None = None

    def add_closable(self, child: tk.Misc, **kwargs) -> None:
        super().add(child, **kwargs)
        tab_id = self.tabs()[-1]
        self._closable_tabs.add(tab_id)

    def forget(self, tab_id: str) -> None:
        self._closable_tabs.discard(tab_id)
        super().forget(tab_id)

    def _on_close_press(self, event: tk.Event) -> None:
        element = self.identify(event.x, event.y)
        if "close" not in element:
            return
        tab = self.index(f"@{event.x},{event.y}")
        if tab == "":
            return
        tab_id = self.tabs()[tab]
        if tab_id not in self._closable_tabs:
            return
        self.state(["pressed"])
        self._pressed_tab = tab_id

    def _on_close_release(self, event: tk.Event) -> None:
        if self._pressed_tab is None:
            return
        element = self.identify(event.x, event.y)
        tab = self.index(f"@{event.x},{event.y}")
        tab_id = self.tabs()[tab] if tab != "" else None
        if "close" in element and tab_id == self._pressed_tab:
            self.forget(tab_id)
            self.event_generate("<<NotebookTabClosed>>", data=tab_id)
        self.state(["!pressed"])
        self._pressed_tab = None


def open_folder(path: str) -> None:
    os.makedirs(path, exist_ok=True)
    try:
        os.startfile(path)
    except Exception as exc:
        messagebox.showerror("错误", f"无法打开文件夹: {exc}")


def open_help(parent: tk.Misc) -> None:
    win = tk.Toplevel(parent)
    win.title("使用说明")
    win.geometry("1120x760")

    container = ttk.Frame(win, padding=10)
    container.pack(fill="both", expand=True)
    container.columnconfigure(1, weight=1)
    container.rowconfigure(0, weight=1)

    sections = {
        "软件介绍": (
            "适用场景：\n"
            "- 把 ITR PDF 表格字段自动填充为 Excel 台账数据。\n"
            "- 对无法自动识别的字段，允许人工校对后再导出。\n\n"
            "核心产出：\n"
            "- output/ 下导出的 PDF（填好字段）\n"
            "- report/ 下的报告（定位问题、统计空字段）"
        ),
        "注意事项": (
            "1）Excel 必须是 .xlsx 格式；.xls 需先另存为 .xlsx。\n"
            "2）输出目录会自动生成：output/ 与 report/。\n"
            "3）预设修改后必须确认/保存，否则主功能会被锁定。\n"
            "4）如果解析失败，请先确认：\n"
            "   - PDF 是否可正常打开\n"
            "   - 匹配键正则与 Excel 表头是否配置正确\n"
            "   - 预设页数与 PDF 实际页数是否一致"
        ),
        "ITR 自动预填使用流程": (
            "准备工作：\n"
            "1）准备 ITR PDF（可多套、多文件）\n"
            "2）准备 Excel 台账（.xlsx）\n\n"
            "操作步骤：\n"
            "1）进入“预设管理”选择或新建预设\n"
            "2）设置 Excel 表头行、匹配键、字段映射\n"
            "3）先做“PDF 定位测试（画框）”确认定位\n"
            "4）回到主界面选择 Excel 和 PDF\n"
            "5）点击“解析&预填”，检查左侧 ITR 列表\n"
            "6）在右侧列表人工修改字段（如 Serial Number）\n"
            "7）点击“导出填好的PDF + report.xlsx”\n\n"
            "输出在哪里：\n"
            "- output/itr_autofill/filled/<batch>/\n"
            "- report/itr_autofill/<batch>/report.xlsx\n\n"
            "常见错误处理：\n"
            "- 解析无结果：检查匹配键正则是否能在 PDF 中抓到 Tag。\n"
            "- 字段为空：检查 excel_col_norm 是否与表头归一化一致。\n"
            "- 定位偏移：先使用“测试PDF定位(画框)”确认位置。"
        ),
        "NA 自动勾选使用流程": (
            "准备工作：\n"
            "1）准备 ITR PDF（可多选）\n\n"
            "操作步骤：\n"
            "1）点击“批量导入 PDF”\n"
            "2）点击“解析（抓锚点）”完成结构识别\n"
            "3）需要验证时点击“测试（生成框图 PDF）”\n"
            "4）确认无误后点击“打勾（NA）”\n\n"
            "输出在哪里：\n"
            "- output/na_check/test/<batch>/（测试框图）\n"
            "- output/na_check/filled/<batch>/（打勾结果）\n"
            "- report/ex_check/<batch>/（如有）\n\n"
            "常见错误处理：\n"
            "- 解析失败：检查 PDF 是否扫描件/不可选中文本。\n"
            "- 打勾位置偏移：先生成测试框图确认表格边界。\n"
            "- 无输出：确认是否先完成“解析”步骤。"
        ),
        "常见问题 FAQ": (
            "Q：report.xlsx 有什么用？\n"
            "A：记录每套 ITR 的匹配情况、空字段清单，便于补录与核对。\n\n"
            "Q：output 和 report 目录在哪里？\n"
            "A：程序根目录下的 output/ 与 report/，主页右下角按钮可直接打开。\n\n"
            "Q：如何确认匹配键正确？\n"
            "A：在预设里设置 PDF 提取正则，并通过测试 PDF 验证 Tag 是否被识别。"
        ),
    }

    nav = tk.Listbox(container, height=12)
    for title in sections:
        nav.insert(tk.END, title)
    nav.grid(row=0, column=0, sticky="ns", padx=(0, 10))

    detail_frame = ttk.Frame(container)
    detail_frame.grid(row=0, column=1, sticky="nsew")
    detail_frame.rowconfigure(0, weight=1)
    detail_frame.columnconfigure(0, weight=1)

    text = tk.Text(detail_frame, wrap="word")
    text.grid(row=0, column=0, sticky="nsew")
    scrollbar = ttk.Scrollbar(detail_frame, orient="vertical", command=text.yview)
    scrollbar.grid(row=0, column=1, sticky="ns")
    text.configure(yscrollcommand=scrollbar.set)

    def show_section(_event=None):
        sel = nav.curselection()
        if not sel:
            return
        title = nav.get(sel[0])
        text.config(state="normal")
        text.delete("1.0", tk.END)
        text.insert("1.0", sections.get(title, ""))
        text.config(state="disabled")

    nav.bind("<<ListboxSelect>>", show_section)
    nav.selection_set(0)
    show_section()


def main() -> None:
    root = tk.Tk()
    root.title(f"{APP_NAME} {APP_VERSION} - 工具合集")
    root.geometry("1500x900")

    notebook = ClosableNotebook(root)
    notebook.pack(fill=tk.BOTH, expand=True)

    tabs: dict[str, ttk.Frame] = {}
    tab_keys: dict[str, str] = {}

    def focus_tab(tab_id: str) -> None:
        frame = tabs.get(tab_id)
        if frame is not None:
            notebook.select(frame)

    def add_tool_tab(tab_id: str, title: str, content_builder) -> None:
        if tab_id in tabs:
            frame = tabs[tab_id]
            notebook.select(frame)
            return

        container = ttk.Frame(notebook)
        content = content_builder(container)
        content.pack(fill="both", expand=True)
        notebook.add_closable(container, text=title)
        new_tab_id = notebook.tabs()[-1]
        tabs[tab_id] = container
        tab_keys[new_tab_id] = tab_id
        notebook.select(container)

    home = ttk.Frame(notebook)
    notebook.add(home, text="主页")

    style = ttk.Style(root)
    style.layout("HomeTab", style.layout("TNotebook.Tab"))
    notebook.tab(home, style="HomeTab")

    home.columnconfigure(0, weight=1)
    home.rowconfigure(2, weight=1)

    header = ttk.Frame(home)
    header.grid(row=0, column=0, sticky="ew", padx=16, pady=(16, 8))
    title = ttk.Label(header, text=f"{APP_NAME} {APP_VERSION} - 工具合集", font=("TkDefaultFont", 16, "bold"))
    title.pack(anchor="w")
    subtitle = ttk.Label(header, text="作者：马瑞泽")
    subtitle.pack(anchor="w", pady=(4, 0))

    ttk.Separator(home, orient="horizontal").grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 12))

    content = ttk.Frame(home)
    content.grid(row=2, column=0, sticky="nsew", padx=16)
    ttk.Label(content, text="功能入口：").pack(anchor="w")
    btns = ttk.Frame(content)
    btns.pack(anchor="w", pady=(10, 0))
    ttk.Button(
        btns,
        text="打开 ITR 自动预填",
        command=lambda: add_tool_tab("itr_autofill", "ITR 自动预填", ITRAutofillTab),
    ).pack(side="left", padx=(0, 8))
    ttk.Button(
        btns,
        text="打开 NA 自动勾选",
        command=lambda: add_tool_tab("na_check", "NA 自动勾选", NACheckTab),
    ).pack(side="left", padx=(0, 8))

    bottom = ttk.Frame(home)
    bottom.grid(row=3, column=0, sticky="sew", padx=16, pady=16)
    bottom.columnconfigure(0, weight=1)
    actions = ttk.Frame(bottom)
    actions.grid(row=0, column=0, sticky="e")
    ttk.Button(actions, text="使用说明", command=lambda: open_help(home)).pack(side="left", padx=(0, 8))
    ttk.Button(actions, text="打开 output", command=lambda: open_folder(OUTPUT_ROOT)).pack(side="left", padx=(0, 8))
    ttk.Button(actions, text="打开 report", command=lambda: open_folder(REPORT_ROOT)).pack(side="left")

    def _on_tab_closed(event: tk.Event) -> None:
        tab_id = event.data
        key = tab_keys.pop(tab_id, None)
        if key:
            tabs.pop(key, None)

    notebook.bind("<<NotebookTabClosed>>", _on_tab_closed)

    root.mainloop()


if __name__ == "__main__":
    main()