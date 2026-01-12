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


def open_folder(path: str) -> None:
    os.makedirs(path, exist_ok=True)
    try:
        os.startfile(path)
    except Exception as exc:
        messagebox.showerror("错误", f"无法打开文件夹: {exc}")


def open_help(parent: tk.Misc) -> None:
    win = tk.Toplevel(parent)
    win.title("使用说明")
    win.geometry("1020x720")
    txt = tk.Text(win, wrap="word")
    txt.pack(fill="both", expand=True, padx=10, pady=10)

    help_text = f"""==============================
ITR PDF 自动预填工具 —— 使用说明
（当前版本：{APP_VERSION}）
==============================

【软件目的】
本工具用于将 ITR PDF 中的空白表格字段，自动或半自动填写为 Excel 台账中的设备信息，减少打印后手写的工作量，并支持人工校对与修改后统一导出。

--------------------------------
一、整体运行流程（新手必看）
--------------------------------
1）准备文件：准备好 ITR PDF（可包含多套 ITR）和对应的 Excel 台账
2）进入“预设管理”：新建或选择一个预设（Preset）
3）配置 Excel：
   - 【Excel 格式】本工具仅支持 .xlsx（新版工作簿）；如果是 .xls（老格式）请先在 Excel 里“另存为” .xlsx 后再导入
   - 【表头行 header_row】指真正的列标题所在行，按 0 开始计数，例如第三行是表头，则 header_row=2
   - 【表头归一化】为兼容不同台账列名写法，程序会将表头归一化为：转大写 + 删除非字母数字字符
     例：Tag No. -> TAGNO；Equipment Tag -> EQUIPMENTTAG；Temp Class -> TEMPCLASS
4）配置匹配键（Match Key）：定义如何从 PDF 提取 Tag/设备编号，并用它去匹配 Excel 行
   - 【匹配键归一化】匹配时会对键值做：去空格 + 转大写，因此类似 627-30-SKT-01-Ex / 627-30-SKT-01 的差异可通过候选规则匹配
5）配置 ITR 拆分规则：填写“每套 ITR 页数（itr_pages_per_set）”
6）配置字段映射：定义 PDF 中每个字段来源（Excel / 手动 / 常量 / 规则）
7）先做一次【PDF 定位测试（画框）】人工确认定位是否正确
8）回到主界面：解析 → 人工检查 → 导出填写完成的 PDF

--------------------------------
二、Excel 列名归一化（一定要看）
--------------------------------
软件内部使用“归一化表头名”来匹配 Excel 列：
- 全部转大写
- 删除空格、下划线、点号、斜杠等非字母数字字符
因此你在字段映射里填 excel_col_norm 时，应填写归一化后的列名（例如 TAGNO、MODEL、IPRATING）。

--------------------------------
三、ITR 拆分规则（非常重要）
--------------------------------
当 PDF 没有页码标记时，软件按“每套 ITR 页数（itr_pages_per_set）”固定拆分：
- 例如一套 ITR 由 4 页组成，则设置为 4；两页则设置为 2
- 只要你确认“一套 ITR 的页数”，就可以稳定工作

--------------------------------
四、字段映射（字段来源怎么理解）
--------------------------------
字段映射决定“这个字段从哪里来”：
- EXCEL：从 Excel 当前匹配行取值（由 excel_col_norm 指定列）
- MANUAL：解析后在主界面人工填写/修改（推荐用于 Serial Number 这类现场补录项）
- CONST：固定值（例如你希望某字段总是空，就把 source=CONST，const_value 留空）
- RULE：规则（少量场景使用，按软件内置规则生成）

提示：导出时会自动换行并缩放字体，尽量不超出格子。

--------------------------------
五、预设管理（新建/保存/另存为）
--------------------------------
- 新建预设：创建一个新的预设草稿
- 保存预设：保存当前预设
  * 如果你在右侧把“预设名”改了，再点“保存预设”，程序会视为【重命名】，不会留下旧的 NewPreset 文件
- 另存为：会保留原预设，并保存一份副本

--------------------------------
六、PDF 定位测试（画框）
--------------------------------
点击“PDF 定位测试（画框）”后，会在 output/itr_autofill/test 文件夹中生成带标记的测试 PDF：
- 蓝框：识别到的 Page 区域/页码区域（用于辅助检查拆分/定位）
- 红框：识别到的可填写空白区域（最终会在这里写入内容）
你也可以点击“打开测试文件夹”快速查看结果。

"""
    txt.insert("1.0", help_text)
    txt.config(state="disabled")


def main() -> None:
    root = tk.Tk()
    root.title(f"{APP_NAME} {APP_VERSION} - 工具合集")
    root.geometry("1500x900")

    notebook = ttk.Notebook(root)
    notebook.pack(fill=tk.BOTH, expand=True)

    tabs: dict[str, ttk.Frame] = {}

    def focus_tab(tab_id: str) -> None:
        frame = tabs.get(tab_id)
        if frame is not None:
            notebook.select(frame)

    def close_tab(tab_id: str) -> None:
        frame = tabs.pop(tab_id, None)
        if frame is None:
            return
        notebook.forget(frame)

    def build_tab_controls(parent: ttk.Frame, tab_id: str, title: str) -> None:
        bar = ttk.Frame(parent)
        bar.pack(fill="x", padx=10, pady=(10, 0))
        ttk.Label(bar, text=title).pack(side="left")
        ttk.Button(bar, text="使用说明", command=lambda: open_help(parent)).pack(side="right")
        ttk.Button(bar, text="关闭标签页", command=lambda: close_tab(tab_id)).pack(side="right", padx=8)
        ttk.Button(bar, text="打开 report", command=lambda: open_folder(REPORT_ROOT)).pack(side="right", padx=8)
        ttk.Button(bar, text="打开 output", command=lambda: open_folder(OUTPUT_ROOT)).pack(side="right", padx=8)

    def add_tool_tab(tab_id: str, title: str, content_builder) -> None:
        if tab_id in tabs:
            focus_tab(tab_id)
            return

        container = ttk.Frame(notebook)
        build_tab_controls(container, tab_id, title)
        content = content_builder(container)
        content.pack(fill="both", expand=True)
        notebook.add(container, text=title)
        tabs[tab_id] = container
        notebook.select(container)

    home = ttk.Frame(notebook)
    notebook.add(home, text="主页")

    title = ttk.Label(home, text=f"{APP_NAME} {APP_VERSION}", font=("TkDefaultFont", 16, "bold"))
    title.pack(anchor="w", padx=16, pady=(16, 8))

    intro = ttk.Label(home, text="请选择要打开的工具标签页：")
    intro.pack(anchor="w", padx=16)

    btns = ttk.Frame(home)
    btns.pack(anchor="w", padx=16, pady=10)
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

    tools = ttk.Frame(home)
    tools.pack(anchor="w", padx=16, pady=(10, 0))
    ttk.Button(tools, text="使用说明", command=lambda: open_help(home)).pack(side="left", padx=(0, 8))
    ttk.Button(tools, text="打开 output", command=lambda: open_folder(OUTPUT_ROOT)).pack(side="left", padx=(0, 8))
    ttk.Button(tools, text="打开 report", command=lambda: open_folder(REPORT_ROOT)).pack(side="left", padx=(0, 8))

    root.mainloop()


if __name__ == "__main__":
    main()
