from __future__ import annotations

import json
import threading
import traceback
from pathlib import Path
from tkinter import (
    BOTH,
    END,
    LEFT,
    RIGHT,
    Button,
    Entry,
    Frame,
    Label,
    Scrollbar,
    Text,
    Tk,
    Toplevel,
    filedialog,
    messagebox,
)

from bomcheck_app.binding_library import BindingLibrary, BindingProject
from bomcheck_app.config import load_config
from bomcheck_app.excel_processor import ExcelProcessor

CONFIG_PATH = Path("config.json")


class Application:
    def __init__(self, root: Tk):
        self.root = root
        self.root.title("料号检测系统")
        self.config = load_config(CONFIG_PATH)
        self.binding_library = BindingLibrary(self.config.binding_library)
        self.binding_library.load()
        self.processor = ExcelProcessor(self.config)
        self.selected_file: Path | None = None
        self._build_ui()

    def _build_ui(self) -> None:
        file_frame = Frame(self.root)
        file_frame.pack(fill=BOTH, padx=10, pady=10)

        Label(file_frame, text="选择BOM Excel文件：").pack(side=LEFT)
        self.file_entry = Entry(file_frame, width=50)
        self.file_entry.pack(side=LEFT, padx=5)
        Button(file_frame, text="浏览", command=self._choose_file).pack(side=LEFT)

        action_frame = Frame(self.root)
        action_frame.pack(fill=BOTH, padx=10, pady=5)
        Button(action_frame, text="执行", command=self._execute).pack(side=LEFT)
        Button(action_frame, text="编辑绑定料号", command=self._open_binding_editor).pack(side=LEFT, padx=5)

        result_frame = Frame(self.root)
        result_frame.pack(fill=BOTH, expand=True, padx=10, pady=10)
        Label(result_frame, text="执行结果：").pack(anchor="w")

        text_container = Frame(result_frame)
        text_container.pack(fill=BOTH, expand=True)
        scrollbar = Scrollbar(text_container)
        scrollbar.pack(side=RIGHT, fill="y")
        self.result_text = Text(text_container, height=15, bg="white", fg="black")
        self.result_text.pack(side=LEFT, fill=BOTH, expand=True)
        self.result_text.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.result_text.yview)

    def _choose_file(self) -> None:
        file_path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx"), ("Excel", "*.xlsm")])
        if file_path:
            self.selected_file = Path(file_path)
            self.file_entry.delete(0, END)
            self.file_entry.insert(0, str(file_path))

    def _execute(self) -> None:
        if not self.selected_file:
            messagebox.showerror("错误", "请先选择Excel文件")
            return
        thread = threading.Thread(target=self._run_execution, daemon=True)
        thread.start()

    def _run_execution(self) -> None:
        try:
            result = self.processor.execute(self.selected_file, self.binding_library)
        except Exception as exc:  # pragma: no cover - runtime safety
            traceback.print_exc()
            self._update_result_box(f"执行失败：{exc}\n{traceback.format_exc()}", success=False)
            return

        lines = [
            f"失效料号数量：{result.replacement_summary.total_invalid_found}",
            f"已替换数量：{result.replacement_summary.total_replaced}",
            "",
            "绑定料号统计：",
        ]
        for binding_result in result.binding_results:
            lines.append(f"- {binding_result.project_desc} ({binding_result.index_part_no})，主料数量：{binding_result.matched_quantity}")
            for group_result in binding_result.requirement_results:
                lines.append(
                    f"  · {group_result.group_name}：需求 {group_result.required_qty}，可用 {group_result.available_qty}，缺少 {group_result.missing_qty}"
                )
        if result.missing_items:
            lines.append("")
            lines.append("缺失物料：")
            for item in result.missing_items:
                lines.append(f"- {item.part_no} {item.desc} 缺少 {item.missing_qty}")
        if result.important_hits:
            lines.append("")
            lines.append("重要物料：")
            for hit in result.important_hits:
                lines.append(f"- {hit.keyword}（{hit.converted_keyword}）：{hit.total_quantity}")
        success = not result.has_missing
        self._update_result_box("\n".join(lines), success=success)

    def _update_result_box(self, message: str, success: bool) -> None:
        def update():
            self.result_text.delete(1.0, END)
            self.result_text.insert(END, message)
            self.result_text.configure(bg="#d4edda" if success else "#f8d7da")

        self.root.after(0, update)

    def _open_binding_editor(self) -> None:
        BindingEditor(self.root, self.binding_library)


class BindingEditor:
    def __init__(self, master, binding_library: BindingLibrary):
        self.binding_library = binding_library
        self.top = Toplevel(master)
        self.top.title("绑定料号编辑")
        self._build_ui()
        self._load_data()

    def _build_ui(self) -> None:
        self.text = Text(self.top, wrap="none")
        self.text.pack(fill=BOTH, expand=True, padx=10, pady=10)

        button_frame = Frame(self.top)
        button_frame.pack(fill=BOTH, padx=10, pady=5)
        Button(button_frame, text="保存", command=self._save).pack(side=LEFT)
        Button(button_frame, text="重新载入", command=self._load_data).pack(side=LEFT, padx=5)
        Button(button_frame, text="新增模板", command=self._add_template).pack(side=LEFT, padx=5)
        Button(button_frame, text="导入Excel", command=self._import_excel).pack(side=LEFT, padx=5)
        Button(button_frame, text="导出Excel", command=self._export_excel).pack(side=LEFT, padx=5)

    def _load_data(self) -> None:
        self.binding_library.load()
        content = json.dumps([project.to_dict() for project in self.binding_library.iter_projects()], ensure_ascii=False, indent=2)
        self.text.delete(1.0, END)
        self.text.insert(END, content)

    def _save(self) -> None:
        raw_text = self.text.get(1.0, END).strip() or "[]"
        try:
            data = json.loads(raw_text)
        except json.JSONDecodeError as exc:
            messagebox.showerror("错误", f"JSON格式错误：{exc}")
            return
        if not isinstance(data, list):
            messagebox.showerror("错误", "数据格式必须为数组")
            return
        try:
            projects = [BindingProject.from_dict(item) for item in data]
        except Exception as exc:  # pragma: no cover - runtime safety
            messagebox.showerror("错误", f"解析失败：{exc}")
            return
        self.binding_library.projects = projects
        try:
            self.binding_library.save()
        except Exception as exc:
            messagebox.showerror("错误", f"保存失败：{exc}")
            return
        messagebox.showinfo("完成", "保存成功")

    def _add_template(self) -> None:
        template = {
            "projectDesc": "项目描述",
            "indexPartNo": "索引料号",
            "indexPartDesc": "索引料号描述",
            "requiredGroups": [
                {
                    "groupName": "分组名称",
                    "number": 1,
                    "choices": [
                        {
                            "partNo": "料号",
                            "desc": "描述",
                            "conditionMode": "",
                            "conditionPartNos": [],
                            "number": 1,
                        }
                    ],
                }
            ],
        }
        current = self.text.get(1.0, END).strip()
        try:
            data = json.loads(current) if current else []
        except json.JSONDecodeError:
            data = []
        data.append(template)
        self.text.delete(1.0, END)
        self.text.insert(END, json.dumps(data, ensure_ascii=False, indent=2))

    def _import_excel(self) -> None:
        file_path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx"), ("Excel", "*.xlsm")])
        if not file_path:
            return
        try:
            self.binding_library.import_excel(Path(file_path))
        except Exception as exc:
            messagebox.showerror("错误", f"导入失败：{exc}")
            return
        self._load_data()
        messagebox.showinfo("完成", "导入成功")

    def _export_excel(self) -> None:
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if not file_path:
            return
        try:
            self.binding_library.export_excel(Path(file_path))
        except Exception as exc:
            messagebox.showerror("错误", f"导出失败：{exc}")
            return
        messagebox.showinfo("完成", "导出成功")


def main() -> None:
    root = Tk()
    Application(root)
    root.mainloop()


if __name__ == "__main__":
    main()
