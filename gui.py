"""Graphical interface for the BOM check system."""
from __future__ import annotations

import json
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import List

from binding_library import BindingLibrary, BindingProject
from bom_processor import BomProcessor, ProcessResult
from config_manager import AppConfig, ConfigManager

SUCCESS_COLOR = "#d4edda"
ERROR_COLOR = "#f8d7da"
NEUTRAL_COLOR = "#ffffff"


class Application(tk.Tk):
    """Main Tk application."""

    def __init__(self, config_manager: ConfigManager):
        super().__init__()
        self.title("料号检测系统")
        self.geometry("900x600")
        self.resizable(True, True)

        self.config_manager = config_manager
        self.config = self.config_manager.load()

        self._bom_path_var = tk.StringVar()
        self._invalid_path_var = tk.StringVar(value=str(self.config.invalid_database))
        self._binding_path_var = tk.StringVar(value=str(self.config.binding_database))
        self._important_path_var = tk.StringVar(value=str(self.config.important_materials))

        self._create_widgets()
        self._processor = self._build_processor()

    # ------------------------------------------------------------------

    def _create_widgets(self) -> None:
        top_frame = ttk.Frame(self)
        top_frame.pack(fill=tk.X, padx=10, pady=10)

        ttk.Label(top_frame, text="选择BOM Excel文件:").grid(row=0, column=0, sticky=tk.W, pady=2)
        bom_entry = ttk.Entry(top_frame, textvariable=self._bom_path_var, width=60)
        bom_entry.grid(row=0, column=1, sticky=tk.W, pady=2)
        ttk.Button(top_frame, text="浏览", command=self._browse_bom).grid(row=0, column=2, padx=5, pady=2)

        ttk.Separator(top_frame, orient=tk.HORIZONTAL).grid(row=1, column=0, columnspan=3, sticky="ew", pady=8)

        ttk.Label(top_frame, text="失效料号数据库:").grid(row=2, column=0, sticky=tk.W)
        invalid_entry = ttk.Entry(top_frame, textvariable=self._invalid_path_var, width=60)
        invalid_entry.grid(row=2, column=1, sticky=tk.W)
        ttk.Button(top_frame, text="浏览", command=lambda: self._browse_config_path(self._invalid_path_var)).grid(row=2, column=2, padx=5)

        ttk.Label(top_frame, text="绑定料号系统库:").grid(row=3, column=0, sticky=tk.W)
        binding_entry = ttk.Entry(top_frame, textvariable=self._binding_path_var, width=60)
        binding_entry.grid(row=3, column=1, sticky=tk.W)
        ttk.Button(top_frame, text="浏览", command=lambda: self._browse_config_path(self._binding_path_var)).grid(row=3, column=2, padx=5)

        ttk.Label(top_frame, text="重要物料描述库:").grid(row=4, column=0, sticky=tk.W)
        important_entry = ttk.Entry(top_frame, textvariable=self._important_path_var, width=60)
        important_entry.grid(row=4, column=1, sticky=tk.W)
        ttk.Button(top_frame, text="浏览", command=lambda: self._browse_config_path(self._important_path_var)).grid(row=4, column=2, padx=5)

        button_frame = ttk.Frame(top_frame)
        button_frame.grid(row=5, column=0, columnspan=3, sticky=tk.W, pady=10)
        ttk.Button(button_frame, text="保存配置", command=self._save_config).pack(side=tk.LEFT, padx=4)
        ttk.Button(button_frame, text="执行", command=self._execute).pack(side=tk.LEFT, padx=4)
        ttk.Button(button_frame, text="编辑绑定料号库", command=self._open_binding_editor).pack(side=tk.LEFT, padx=4)

        result_frame = ttk.LabelFrame(self, text="执行结果")
        result_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self._result_text = tk.Text(result_frame, wrap=tk.WORD, state=tk.DISABLED, bg=NEUTRAL_COLOR)
        self._result_text.pack(fill=tk.BOTH, expand=True)

    # ------------------------------------------------------------------
    # Event handlers

    def _browse_bom(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")])
        if path:
            self._bom_path_var.set(path)

    def _browse_config_path(self, variable: tk.StringVar) -> None:
        path = filedialog.askopenfilename()
        if path:
            variable.set(path)

    def _save_config(self) -> None:
        config = self._gather_config()
        self.config_manager.save(config)
        self.config = config
        self._processor = self._build_processor()
        messagebox.showinfo("提示", "配置已保存")

    def _execute(self) -> None:
        bom_path = Path(self._bom_path_var.get().strip())
        if not bom_path.exists():
            messagebox.showerror("错误", "请选择有效的BOM文件")
            return
        self.config = self._gather_config()
        self._processor = self._build_processor()
        try:
            result = self._processor.process(bom_path)
        except Exception as exc:  # pylint: disable=broad-except
            messagebox.showerror("执行失败", str(exc))
            self._write_result_text(f"执行失败: {exc}", ERROR_COLOR)
            return
        self._display_result(result)
        messagebox.showinfo("完成", f"处理完成，结果已写入 {bom_path}")

    def _open_binding_editor(self) -> None:
        BindingLibraryEditor(self, BindingLibrary(self._gather_config().binding_database))

    # ------------------------------------------------------------------

    def _display_result(self, result: ProcessResult) -> None:
        lines = [
            f"失效料号数量: {result.invalid_count}",
            f"已替换数量: {result.replaced_count}",
            f"数量列: 第{result.quantity_column}列",
            "",
            "未替换料号:",
        ]
        if result.unreplaced_items:
            for item in result.unreplaced_items:
                lines.append(f"  - {item.part_no} {item.desc}")
        else:
            lines.append("  无")
        lines.append("")
        lines.append("绑定料号统计:")
        if result.project_results:
            for project in result.project_results:
                lines.append(f"* {project.project_desc} ({project.index_part_no}) 数量: {project.index_quantity}")
                for group in project.groups:
                    status = "OK" if group.missing <= 0.000001 else f"缺少 {group.missing}"
                    lines.append(
                        f"    - {group.group_name}: 需求 {group.required}, 可用 {group.available}, {status}"
                    )
        else:
            lines.append("  无匹配项目")
        lines.append("")
        lines.append("缺失物料:")
        if result.missing_items:
            for item in result.missing_items:
                lines.append(f"  - {item.part_no}: {item.desc} 缺少 {item.quantity}")
        else:
            lines.append("  无")
        lines.append("")
        lines.append("重要物料:")
        if result.important_materials:
            for item in result.important_materials:
                lines.append(f"  - {item.part_no}: {item.desc} 数量 {item.quantity}")
        else:
            lines.append("  无")

        color = SUCCESS_COLOR if result.is_success else ERROR_COLOR
        self._write_result_text("\n".join(lines), color)

    def _write_result_text(self, text: str, color: str) -> None:
        self._result_text.configure(state=tk.NORMAL)
        self._result_text.delete("1.0", tk.END)
        self._result_text.insert(tk.END, text)
        self._result_text.configure(state=tk.DISABLED, bg=color)

    def _gather_config(self) -> AppConfig:
        return AppConfig.from_mapping(
            {
                "invalid_database": self._invalid_path_var.get().strip(),
                "binding_database": self._binding_path_var.get().strip(),
                "important_materials": self._important_path_var.get().strip(),
            },
            base_dir=Path.cwd(),
        )

    def _build_processor(self) -> BomProcessor:
        return BomProcessor(
            invalid_database=Path(self._invalid_path_var.get().strip()),
            binding_library_path=Path(self._binding_path_var.get().strip()),
            important_materials_path=Path(self._important_path_var.get().strip()),
        )


class BindingLibraryEditor(tk.Toplevel):
    """Window for editing binding projects."""

    def __init__(self, master: tk.Tk, library: BindingLibrary):  # type: ignore[override]
        super().__init__(master)
        self.title("绑定料号系统库编辑")
        self.geometry("800x500")
        self.resizable(True, True)
        self.library = library
        self.projects: List[BindingProject] = self.library.load()

        self._create_widgets()
        self._refresh_tree()

    def _create_widgets(self) -> None:
        toolbar = ttk.Frame(self)
        toolbar.pack(fill=tk.X, pady=5)
        ttk.Button(toolbar, text="新增", command=self._add_project).pack(side=tk.LEFT, padx=4)
        ttk.Button(toolbar, text="编辑", command=self._edit_project).pack(side=tk.LEFT, padx=4)
        ttk.Button(toolbar, text="删除", command=self._delete_project).pack(side=tk.LEFT, padx=4)
        ttk.Button(toolbar, text="导入Excel", command=self._import_excel).pack(side=tk.LEFT, padx=4)
        ttk.Button(toolbar, text="导出Excel", command=self._export_excel).pack(side=tk.LEFT, padx=4)
        ttk.Button(toolbar, text="保存到文件", command=self._save_to_file).pack(side=tk.LEFT, padx=4)
        ttk.Button(toolbar, text="刷新", command=self._reload).pack(side=tk.LEFT, padx=4)

        content = ttk.Panedwindow(self, orient=tk.HORIZONTAL)
        content.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        tree_frame = ttk.Frame(content)
        self._tree = ttk.Treeview(tree_frame, columns=("index", "groups"), show="headings", selectmode="browse")
        self._tree.heading("index", text="索引料号")
        self._tree.heading("groups", text="分组数量")
        self._tree.pack(fill=tk.BOTH, expand=True)
        self._tree.bind("<<TreeviewSelect>>", lambda _: self._display_selected())
        content.add(tree_frame, weight=1)

        text_frame = ttk.Frame(content)
        self._json_text = tk.Text(text_frame, wrap=tk.NONE)
        self._json_text.pack(fill=tk.BOTH, expand=True)
        ttk.Button(text_frame, text="复制", command=self._copy_json).pack(anchor=tk.E, pady=4)
        content.add(text_frame, weight=1)

    def _refresh_tree(self) -> None:
        for item in self._tree.get_children():
            self._tree.delete(item)
        for idx, project in enumerate(self.projects):
            self._tree.insert("", tk.END, iid=str(idx), values=(project.index_part_no, len(project.required_groups)))
        self._display_selected()

    def _display_selected(self) -> None:
        selection = self._tree.selection()
        self._json_text.delete("1.0", tk.END)
        if not selection:
            return
        project = self.projects[int(selection[0])]
        self._json_text.insert(tk.END, json.dumps(project.to_mapping(), ensure_ascii=False, indent=2))

    def _copy_json(self) -> None:
        text = self._json_text.get("1.0", tk.END).strip()
        if not text:
            return
        self.clipboard_clear()
        self.clipboard_append(text)
        messagebox.showinfo("提示", "已复制到剪贴板")

    def _add_project(self) -> None:
        template = {
            "projectDesc": "",
            "indexPartNo": "",
            "indexPartDesc": "",
            "requiredGroups": [],
        }
        editor = ProjectEditorDialog(self, template)
        self.wait_window(editor)
        if editor.result:
            self.projects.append(editor.result)
            self._refresh_tree()

    def _edit_project(self) -> None:
        selection = self._tree.selection()
        if not selection:
            messagebox.showwarning("提示", "请选择要编辑的项目")
            return
        index = int(selection[0])
        project = self.projects[index]
        editor = ProjectEditorDialog(self, project.to_mapping())
        self.wait_window(editor)
        if editor.result:
            self.projects[index] = editor.result
            self._refresh_tree()
            self._tree.selection_set(str(index))

    def _delete_project(self) -> None:
        selection = self._tree.selection()
        if not selection:
            messagebox.showwarning("提示", "请选择要删除的项目")
            return
        index = int(selection[0])
        if messagebox.askyesno("确认", "确定要删除该项目吗？"):
            del self.projects[index]
            self._refresh_tree()

    def _import_excel(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")])
        if not path:
            return
        try:
            self.projects = self.library.import_from_excel(Path(path))
            messagebox.showinfo("提示", "导入成功")
            self._refresh_tree()
        except Exception as exc:  # pylint: disable=broad-except
            messagebox.showerror("导入失败", str(exc))

    def _export_excel(self) -> None:
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel文件", "*.xlsx")])
        if not path:
            return
        try:
            self.library.export_to_excel(Path(path), self.projects)
            messagebox.showinfo("提示", "导出成功")
        except Exception as exc:  # pylint: disable=broad-except
            messagebox.showerror("导出失败", str(exc))

    def _save_to_file(self) -> None:
        try:
            self.library.save(self.projects)
            messagebox.showinfo("提示", "保存成功")
        except Exception as exc:  # pylint: disable=broad-except
            messagebox.showerror("保存失败", str(exc))

    def _reload(self) -> None:
        self.projects = self.library.load()
        self._refresh_tree()


class ProjectEditorDialog(tk.Toplevel):
    """Simple JSON editor dialog for binding projects."""

    def __init__(self, master: tk.Toplevel, data: dict):  # type: ignore[override]
        super().__init__(master)
        self.title("编辑项目")
        self.geometry("600x500")
        self.resizable(True, True)
        self.result: BindingProject | None = None

        ttk.Label(self, text="请以JSON格式编辑项目：").pack(anchor=tk.W, padx=10, pady=5)
        self._text = tk.Text(self, wrap=tk.NONE)
        self._text.pack(fill=tk.BOTH, expand=True, padx=10)
        self._text.insert(tk.END, json.dumps(data, ensure_ascii=False, indent=2))

        button_frame = ttk.Frame(self)
        button_frame.pack(fill=tk.X, pady=10)
        ttk.Button(button_frame, text="取消", command=self.destroy).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="确定", command=self._confirm).pack(side=tk.RIGHT, padx=5)

    def _confirm(self) -> None:
        raw = self._text.get("1.0", tk.END)
        try:
            data = json.loads(raw)
            project = BindingProject.from_mapping(data)
        except Exception as exc:  # pylint: disable=broad-except
            messagebox.showerror("错误", f"解析失败: {exc}")
            return
        self.result = project
        self.destroy()
