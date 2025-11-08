from __future__ import annotations

import threading
import traceback
from pathlib import Path
from tkinter import (
    BOTH,
    END,
    LEFT,
    RIGHT,
    Y,
    Button,
    Entry,
    Frame,
    Label,
    Listbox,
    Scrollbar,
    StringVar,
    Text,
    Tk,
    Toplevel,
    filedialog,
    messagebox,
)
from tkinter import ttk

from bomcheck_app.binding_library import BindingChoice, BindingGroup, BindingLibrary, BindingProject
from bomcheck_app.config import load_config
from bomcheck_app.excel_processor import ExcelProcessor, SaveWorkbookError, format_quantity_text
from bomcheck_app.models import ExecutionResult

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
        self._execution_lock = threading.Lock()
        self._execution_thread: threading.Thread | None = None
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
        self.execute_button = Button(action_frame, text="执行", command=self._execute)
        self.execute_button.pack(side=LEFT)
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
        if not self._execution_lock.acquire(blocking=False):
            messagebox.showinfo("处理中", "上一轮执行尚未完成，请稍候再试。")
            return
        self.execute_button.config(state="disabled")
        thread = threading.Thread(target=self._run_execution, daemon=True)
        self._execution_thread = thread
        thread.start()

    def _run_execution(self) -> None:
        try:
            try:
                result = self.processor.execute(self.selected_file, self.binding_library)
            except SaveWorkbookError as error:
                result = error.result
                self._handle_save_error(error)
            except Exception as exc:  # pragma: no cover - runtime safety
                traceback.print_exc()
                self._update_result_box(f"执行失败：{exc}\n{traceback.format_exc()}", success=False)
                return

            summary_lines = self._build_summary_lines(result)
            success = not result.has_missing
            self._update_result_box("\n".join(summary_lines), success=success)
        finally:
            self.root.after(0, self._on_execution_complete)

    def _on_execution_complete(self) -> None:
        if hasattr(self, "execute_button"):
            self.execute_button.config(state="normal")
        if self._execution_lock.locked():
            self._execution_lock.release()
        self._execution_thread = None

    def _update_result_box(self, message: str, success: bool) -> None:
        def update():
            self.result_text.delete(1.0, END)
            self.result_text.insert(END, message)
            self.result_text.configure(bg="#d4edda" if success else "#f8d7da")

        self.root.after(0, update)

    def _handle_save_error(self, error: SaveWorkbookError) -> None:
        decision_event = threading.Event()
        default_extension = error.path.suffix or ".xlsx"

        def prompt() -> None:
            message = (
                f"无法写入文件：{error.path}\n"
                "该文件可能已在其他程序中打开。是否将结果另存为其他文件？"
            )
            save_elsewhere = messagebox.askyesno("文件被占用", message)
            if save_elsewhere:
                new_path = filedialog.asksaveasfilename(
                    title="另存结果",
                    defaultextension=default_extension,
                    filetypes=[("Excel", "*.xlsx"), ("Excel", "*.xlsm")],
                    initialfile=error.path.name,
                )
                if new_path:
                    try:
                        error.workbook.save(new_path)
                        messagebox.showinfo("保存成功", f"结果已保存到：{new_path}")
                    except PermissionError:
                        messagebox.showerror("保存失败", "目标文件正在使用中，未能保存结果。")
                    except Exception as exc:  # pragma: no cover - user feedback
                        messagebox.showerror("保存失败", f"另存失败：{exc}")
            decision_event.set()

        self.root.after(0, prompt)
        decision_event.wait()

    def _open_binding_editor(self) -> None:
        BindingEditor(self.root, self.binding_library)

    def _build_summary_lines(self, result: ExecutionResult) -> list[str]:
        lines = [
            f"失效料号数量：{format_quantity_text(result.replacement_summary.total_invalid_found)}",
            f"已标记失效料号数量：{format_quantity_text(result.replacement_summary.total_invalid_previously_marked)}",
            f"已替换数量：{format_quantity_text(result.replacement_summary.total_replaced)}",
        ]
        lines.extend(self._summarize_binding_results(result))
        lines.extend(self._summarize_missing_items(result))
        lines.extend(self._summarize_important_hits(result))
        lines.extend(self._summarize_debug_logs(result))
        return lines

    def _summarize_binding_results(self, result: ExecutionResult) -> list[str]:
        binding_group_count = sum(
            len(res.requirement_results) for res in result.binding_results
        )
        lines = [
            "",
            (
                "绑定料号统计：找到 "
                f"{format_quantity_text(len(result.binding_results))} 组项目，"
                f"需求分组 {format_quantity_text(binding_group_count)} 组"
            ),
        ]
        if not result.binding_results:
            lines.append("（未找到匹配的绑定项目）")
            return lines

        for binding_result in result.binding_results:
            project_header = (
                f"- {binding_result.project_desc} ({binding_result.index_part_no})，"
                f"主料数量：{format_quantity_text(binding_result.matched_quantity)}"
            )
            lines.append(project_header)

            for group_result in binding_result.requirement_results:
                group_line = (
                    "  · "
                    + f"{group_result.group_name}：需求 {format_quantity_text(group_result.required_qty)}，"
                    + f"可用 {format_quantity_text(group_result.available_qty)}，"
                    + f"缺少 {format_quantity_text(group_result.missing_qty)}"
                )
                lines.append(group_line)

                if group_result.matched_details:
                    matched_pairs = [
                        f"{part}:{format_quantity_text(qty)}"
                        for part, qty in group_result.matched_details.items()
                    ]
                    lines.append("    满足料号：" + ", ".join(matched_pairs))

                if group_result.missing_choices:
                    lines.append(
                        "    缺少料号：" + ", ".join(group_result.missing_choices)
                    )

        return lines

    def _summarize_missing_items(self, result: ExecutionResult) -> list[str]:
        if not result.missing_items:
            return []

        lines = ["", "缺失物料："]
        for item in result.missing_items:
            lines.append(
                f"- {item.part_no} {item.desc} 缺少 {format_quantity_text(item.missing_qty)}"
            )
        return lines

    def _summarize_important_hits(self, result: ExecutionResult) -> list[str]:
        lines = [
            "",
            f"重要物料统计：找到 {format_quantity_text(len(result.important_hits))} 组",
        ]
        if not result.important_hits:
            lines.append("（无重要物料命中）")
            return lines

        for hit in result.important_hits:
            lines.append(
                f"- {hit.keyword}（{hit.converted_keyword}）：{format_quantity_text(hit.total_quantity)}"
            )
            if hit.matched_parts:
                matched_text = ", ".join(
                    f"{part}:{format_quantity_text(qty)}"
                    for part, qty in hit.matched_parts.items()
                )
                lines.append(f"    命中料号：{matched_text}")
        return lines

    def _summarize_debug_logs(self, result: ExecutionResult) -> list[str]:
        if not result.debug_logs:
            return []

        return ["", "调试信息：", *[f"- {log}" for log in result.debug_logs]]

class BindingEditor:
    def __init__(self, master, binding_library: BindingLibrary):
        self.binding_library = binding_library
        self.top = Toplevel(master)
        self.top.title("绑定料号编辑")
        self.projects: list[BindingProject] = []
        self.selected_project_index: int | None = None
        self.selected_group_index: int | None = None
        self.selected_choice_index: int | None = None
        self.project_clipboard: BindingProject | None = None
        self.group_clipboard: BindingGroup | None = None
        self._build_ui()
        self._load_data()

    def _build_ui(self) -> None:
        main_frame = Frame(self.top)
        main_frame.pack(fill=BOTH, expand=True)

        # Project list
        project_frame = Frame(main_frame)
        project_frame.pack(side=LEFT, fill=Y, padx=10, pady=10)
        Label(project_frame, text="项目列表").pack(anchor="w")
        self.project_list = Listbox(
            project_frame,
            exportselection=False,
            height=15,
            activestyle="none",
            selectmode="browse",
        )
        self.project_list.pack(fill=Y, expand=True)
        self.project_list.bind("<<ListboxSelect>>", lambda _event: self._on_project_select())
        Button(project_frame, text="新增项目", command=self._add_project).pack(fill=BOTH, pady=2)
        Button(project_frame, text="删除项目", command=self._remove_project).pack(fill=BOTH, pady=2)
        Button(project_frame, text="上移", command=lambda: self._move_project(-1)).pack(fill=BOTH, pady=2)
        Button(project_frame, text="下移", command=lambda: self._move_project(1)).pack(fill=BOTH, pady=2)
        Button(project_frame, text="复制项目", command=self._copy_project).pack(fill=BOTH, pady=2)
        Button(project_frame, text="粘贴项目", command=self._paste_project).pack(fill=BOTH, pady=2)

        # Detail panel
        detail_frame = Frame(main_frame)
        detail_frame.pack(side=LEFT, fill=BOTH, expand=True, padx=10, pady=10)

        basic_frame = Frame(detail_frame)
        basic_frame.pack(fill=BOTH)
        Label(basic_frame, text="项目描述：").grid(row=0, column=0, sticky="w")
        self.project_desc_var = StringVar()
        Entry(basic_frame, textvariable=self.project_desc_var).grid(row=0, column=1, sticky="ew")
        Label(basic_frame, text="索引料号：").grid(row=1, column=0, sticky="w")
        self.project_index_var = StringVar()
        Entry(basic_frame, textvariable=self.project_index_var).grid(row=1, column=1, sticky="ew")
        Label(basic_frame, text="索引描述：").grid(row=2, column=0, sticky="w")
        self.project_index_desc_var = StringVar()
        Entry(basic_frame, textvariable=self.project_index_desc_var).grid(row=2, column=1, sticky="ew")
        basic_frame.columnconfigure(1, weight=1)

        # Groups
        group_frame = Frame(detail_frame)
        group_frame.pack(fill=BOTH, expand=True, pady=(10, 0))
        group_left = Frame(group_frame)
        group_left.pack(side=LEFT, fill=Y)
        Label(group_left, text="需求分组").pack(anchor="w")
        self.group_list = Listbox(
            group_left,
            exportselection=False,
            height=10,
            activestyle="none",
            selectmode="browse",
        )
        self.group_list.pack(fill=Y, expand=True)
        self.group_list.bind("<<ListboxSelect>>", lambda _event: self._on_group_select())
        Button(group_left, text="新增分组", command=self._add_group).pack(fill=BOTH, pady=2)
        Button(group_left, text="删除分组", command=self._remove_group).pack(fill=BOTH, pady=2)
        Button(group_left, text="复制分组", command=self._copy_group).pack(fill=BOTH, pady=2)
        Button(group_left, text="粘贴分组", command=self._paste_group).pack(fill=BOTH, pady=2)

        group_detail = Frame(group_frame)
        group_detail.pack(side=LEFT, fill=BOTH, expand=True, padx=10)
        Label(group_detail, text="分组名称：").grid(row=0, column=0, sticky="w")
        self.group_name_var = StringVar()
        Entry(group_detail, textvariable=self.group_name_var).grid(row=0, column=1, sticky="ew")
        Label(group_detail, text="需求数量：").grid(row=1, column=0, sticky="w")
        self.group_number_var = StringVar()
        Entry(group_detail, textvariable=self.group_number_var).grid(row=1, column=1, sticky="ew")
        group_detail.columnconfigure(1, weight=1)

        # Choices table
        choice_frame = Frame(group_detail)
        choice_frame.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=(10, 0))
        group_detail.rowconfigure(2, weight=1)
        Label(choice_frame, text="可选料号").pack(anchor="w")
        columns = ("part_no", "desc", "condition_mode", "condition_part_nos", "number")
        self.choice_tree = ttk.Treeview(choice_frame, columns=columns, show="headings", height=6)
        headings = {
            "part_no": "料号",
            "desc": "描述",
            "condition_mode": "条件模式",
            "condition_part_nos": "条件料号",
            "number": "数量",
        }
        for key, title in headings.items():
            self.choice_tree.heading(key, text=title)
            self.choice_tree.column(key, width=100, anchor="w")
        self.choice_tree.pack(fill=BOTH, expand=True)
        self.choice_tree.bind("<<TreeviewSelect>>", lambda _event: self._on_choice_select())

        choice_edit = Frame(choice_frame)
        choice_edit.pack(fill=BOTH, pady=(5, 0))
        Label(choice_edit, text="料号：").grid(row=0, column=0, sticky="w")
        self.choice_part_var = StringVar()
        Entry(choice_edit, textvariable=self.choice_part_var).grid(row=0, column=1, sticky="ew")
        Label(choice_edit, text="描述：").grid(row=1, column=0, sticky="w")
        self.choice_desc_var = StringVar()
        Entry(choice_edit, textvariable=self.choice_desc_var).grid(row=1, column=1, sticky="ew")
        Label(choice_edit, text="条件模式：").grid(row=2, column=0, sticky="w")
        self.choice_mode_var = StringVar()
        Entry(choice_edit, textvariable=self.choice_mode_var).grid(row=2, column=1, sticky="ew")
        Label(choice_edit, text="条件料号：").grid(row=3, column=0, sticky="w")
        self.choice_condition_var = StringVar()
        Entry(choice_edit, textvariable=self.choice_condition_var).grid(row=3, column=1, sticky="ew")
        Label(choice_edit, text="数量：").grid(row=4, column=0, sticky="w")
        self.choice_number_var = StringVar()
        Entry(choice_edit, textvariable=self.choice_number_var).grid(row=4, column=1, sticky="ew")
        choice_edit.columnconfigure(1, weight=1)

        choice_btn_frame = Frame(choice_frame)
        choice_btn_frame.pack(fill=BOTH, pady=(5, 0))
        Button(choice_btn_frame, text="新增料号", command=self._add_choice).pack(side=LEFT, padx=2)
        Button(choice_btn_frame, text="更新料号", command=self._update_choice).pack(side=LEFT, padx=2)
        Button(choice_btn_frame, text="删除料号", command=self._remove_choice).pack(side=LEFT, padx=2)

        # Bottom action buttons
        button_frame = Frame(self.top)
        button_frame.pack(fill=BOTH, padx=10, pady=5)
        Button(button_frame, text="保存", command=self._save).pack(side=LEFT)
        Button(button_frame, text="重新载入", command=self._load_data).pack(side=LEFT, padx=5)
        Button(button_frame, text="导入Excel", command=self._import_excel).pack(side=LEFT, padx=5)
        Button(button_frame, text="导出Excel", command=self._export_excel).pack(side=LEFT, padx=5)

    def _load_data(self) -> None:
        self.binding_library.load()
        self.projects = [BindingProject.from_dict(project.to_dict()) for project in self.binding_library.iter_projects()]
        self.selected_project_index = None
        self.selected_group_index = None
        self.selected_choice_index = None
        self._refresh_project_list()
        self._clear_project_fields()
        if self.projects:
            self.project_list.selection_set(0)
            self._on_project_select()

    def _refresh_project_list(self) -> None:
        self.project_list.delete(0, END)
        for project in self.projects:
            display = f"{project.project_desc or '未命名'} ({project.index_part_no or '-'})"
            self.project_list.insert(END, display)

    def _clear_project_fields(self) -> None:
        self.project_desc_var.set("")
        self.project_index_var.set("")
        self.project_index_desc_var.set("")
        self.group_list.delete(0, END)
        self.group_name_var.set("")
        self.group_number_var.set("")
        for item in self.choice_tree.get_children():
            self.choice_tree.delete(item)
        self.choice_part_var.set("")
        self.choice_desc_var.set("")
        self.choice_mode_var.set("")
        self.choice_condition_var.set("")
        self.choice_number_var.set("")

    def _on_project_select(self) -> None:
        selection = self.project_list.curselection()
        if self.selected_project_index is not None:
            self._commit_choice_fields()
            self._commit_group_fields()
            self._commit_project_fields()
        if not selection:
            self.selected_project_index = None
            self._clear_project_fields()
            return
        self.selected_project_index = selection[0]
        project = self.projects[self.selected_project_index]
        self.project_desc_var.set(project.project_desc)
        self.project_index_var.set(project.index_part_no)
        self.project_index_desc_var.set(project.index_part_desc)
        self._refresh_group_list()

    def _commit_project_fields(self) -> None:
        if self.selected_project_index is None:
            return
        project = self.projects[self.selected_project_index]
        project.project_desc = self.project_desc_var.get().strip()
        project.index_part_no = self.project_index_var.get().strip()
        project.index_part_desc = self.project_index_desc_var.get().strip()
        display = f"{project.project_desc or '未命名'} ({project.index_part_no or '-'})"
        if 0 <= self.selected_project_index < self.project_list.size():
            current_selection = self.project_list.curselection()
            preserve_selection = (
                len(current_selection) == 1
                and current_selection[0] == self.selected_project_index
            )
            self.project_list.delete(self.selected_project_index)
            self.project_list.insert(self.selected_project_index, display)
            if preserve_selection:
                self.project_list.selection_clear(0, END)
                self.project_list.selection_set(self.selected_project_index)

    def _refresh_group_list(self) -> None:
        self.group_list.delete(0, END)
        self.selected_group_index = None
        self.selected_choice_index = None
        for group in self.projects[self.selected_project_index].required_groups:
            display = f"{group.group_name or '未命名'} (需求:{group.number})"
            self.group_list.insert(END, display)
        self.group_name_var.set("")
        self.group_number_var.set("")
        for item in self.choice_tree.get_children():
            self.choice_tree.delete(item)
        if self.projects[self.selected_project_index].required_groups:
            self.group_list.selection_set(0)
            self._on_group_select()

    def _on_group_select(self) -> None:
        if self.selected_group_index is not None:
            self._commit_choice_fields()
            self._commit_group_fields()
        selection = self.group_list.curselection()
        if not selection:
            self.selected_group_index = None
            self.group_name_var.set("")
            self.group_number_var.set("")
            self._clear_choice_fields()
            return
        self.selected_group_index = selection[0]
        group = self.projects[self.selected_project_index].required_groups[self.selected_group_index]
        self.group_name_var.set(group.group_name)
        self.group_number_var.set(str(group.number))
        self._refresh_choice_list(auto_select_first=True)

    def _commit_group_fields(self) -> None:
        if self.selected_project_index is None or self.selected_group_index is None:
            return
        group = self.projects[self.selected_project_index].required_groups[self.selected_group_index]
        group.group_name = self.group_name_var.get().strip()
        try:
            group.number = float(self.group_number_var.get()) if self.group_number_var.get().strip() else 1.0
        except ValueError:
            group.number = 1.0
        display = f"{group.group_name or '未命名'} (需求:{group.number})"
        current_selection = self.group_list.curselection()
        preserve_selection = self.selected_group_index in current_selection
        if 0 <= self.selected_group_index < self.group_list.size():
            self.group_list.delete(self.selected_group_index)
            self.group_list.insert(self.selected_group_index, display)
            if preserve_selection:
                self.group_list.selection_set(self.selected_group_index)

    def _refresh_choice_list(self, auto_select_first: bool = False) -> None:
        self.choice_tree.selection_remove(self.choice_tree.selection())
        for item in self.choice_tree.get_children():
            self.choice_tree.delete(item)
        group = self.projects[self.selected_project_index].required_groups[self.selected_group_index]
        for idx, choice in enumerate(group.choices):
            self.choice_tree.insert(
                "",
                "end",
                iid=str(idx),
                values=(
                    choice.part_no,
                    choice.desc,
                    choice.condition_mode or "",
                    ",".join(choice.condition_part_nos),
                    choice.number if choice.number is not None else "",
                ),
            )
        self._clear_choice_fields()
        if auto_select_first and group.choices:
            first_id = "0"
            self.choice_tree.focus(first_id)
            self.choice_tree.selection_set(first_id)
            self._on_choice_select()
        self.choice_tree.update_idletasks()

    def _clear_choice_fields(self) -> None:
        self.choice_tree.selection_remove(self.choice_tree.selection())
        self.choice_tree.focus("")
        self.choice_part_var.set("")
        self.choice_desc_var.set("")
        self.choice_mode_var.set("")
        self.choice_condition_var.set("")
        self.choice_number_var.set("")
        self.selected_choice_index = None

    def _on_choice_select(self) -> None:
        selection = self.choice_tree.selection()
        if self.selected_choice_index is not None:
            self._commit_choice_fields()
        if not selection:
            self.selected_choice_index = None
            self._clear_choice_fields()
            return
        self.selected_choice_index = int(selection[0])
        choice = self.projects[self.selected_project_index].required_groups[self.selected_group_index].choices[
            self.selected_choice_index
        ]
        self.choice_part_var.set(choice.part_no)
        self.choice_desc_var.set(choice.desc)
        self.choice_mode_var.set(choice.condition_mode or "")
        self.choice_condition_var.set(",".join(choice.condition_part_nos))
        self.choice_number_var.set("" if choice.number is None else str(choice.number))

    def _commit_choice_fields(self) -> None:
        if (
            self.selected_project_index is None
            or self.selected_group_index is None
            or self.selected_choice_index is None
        ):
            return
        group = self.projects[self.selected_project_index].required_groups[self.selected_group_index]
        if self.selected_choice_index >= len(group.choices):
            return
        choice = group.choices[self.selected_choice_index]
        choice.part_no = self.choice_part_var.get().strip()
        choice.desc = self.choice_desc_var.get().strip()
        choice.condition_mode = self.choice_mode_var.get().strip() or None
        condition_raw = self.choice_condition_var.get().strip()
        choice.condition_part_nos = [item.strip() for item in condition_raw.split(",") if item.strip()]
        try:
            choice.number = float(self.choice_number_var.get()) if self.choice_number_var.get().strip() else None
        except ValueError:
            choice.number = None
        values = (
            choice.part_no,
            choice.desc,
            choice.condition_mode or "",
            ",".join(choice.condition_part_nos),
            choice.number if choice.number is not None else "",
        )
        item_id = str(self.selected_choice_index)
        if item_id in self.choice_tree.get_children():
            self.choice_tree.item(item_id, values=values)

    def _add_project(self) -> None:
        self._commit_all()
        new_project = BindingProject(project_desc="新项目", index_part_no="", index_part_desc="", required_groups=[])
        self.projects.append(new_project)
        self._refresh_project_list()
        self.project_list.selection_clear(0, END)
        new_index = len(self.projects) - 1
        self.project_list.selection_set(new_index)
        self.project_list.event_generate("<<ListboxSelect>>")

    def _remove_project(self) -> None:
        selection = self.project_list.curselection()
        if not selection:
            return
        index = selection[0]
        del self.projects[index]
        self._refresh_project_list()
        self._clear_project_fields()
        self.selected_project_index = None
        if self.projects:
            new_index = min(index, len(self.projects) - 1)
            self.project_list.selection_set(new_index)
            self._on_project_select()

    def _move_project(self, direction: int) -> None:
        selection = self.project_list.curselection()
        if not selection:
            return
        index = selection[0]
        target_index = index + direction
        if target_index < 0 or target_index >= len(self.projects):
            return
        self._commit_all()
        self.projects[index], self.projects[target_index] = (
            self.projects[target_index],
            self.projects[index],
        )
        self._refresh_project_list()
        self.selected_project_index = None
        self.selected_group_index = None
        self.selected_choice_index = None
        self.project_list.selection_clear(0, END)
        self.project_list.selection_set(target_index)
        self._on_project_select()

    def _copy_project(self) -> None:
        if self.selected_project_index is None:
            messagebox.showwarning("提示", "请先选择项目")
            return
        self._commit_all()
        project = self.projects[self.selected_project_index]
        self.project_clipboard = BindingProject.from_dict(project.to_dict())

    def _paste_project(self) -> None:
        if self.project_clipboard is None:
            messagebox.showwarning("提示", "请先复制项目")
            return
        self._commit_all()
        new_project = BindingProject.from_dict(self.project_clipboard.to_dict())
        self.projects.append(new_project)
        self._refresh_project_list()
        new_index = len(self.projects) - 1
        self.project_list.selection_clear(0, END)
        self.project_list.selection_set(new_index)
        self.project_list.event_generate("<<ListboxSelect>>")

    def _add_group(self) -> None:
        if self.selected_project_index is None:
            messagebox.showwarning("提示", "请先选择项目")
            return
        self._commit_project_fields()
        self._commit_group_fields()
        self._commit_choice_fields()
        group = BindingGroup(group_name="新分组", number=1.0, choices=[])
        self.projects[self.selected_project_index].required_groups.append(group)
        self._refresh_group_list()
        new_index = len(self.projects[self.selected_project_index].required_groups) - 1
        self.group_list.selection_set(new_index)
        self.group_list.event_generate("<<ListboxSelect>>")

    def _remove_group(self) -> None:
        if self.selected_project_index is None:
            return
        selection = self.group_list.curselection()
        if not selection:
            return
        index = selection[0]
        del self.projects[self.selected_project_index].required_groups[index]
        self._refresh_group_list()
        groups = self.projects[self.selected_project_index].required_groups
        if groups:
            new_index = min(index, len(groups) - 1)
            self.group_list.selection_set(new_index)
            self._on_group_select()

    def _copy_group(self) -> None:
        if self.selected_project_index is None or self.selected_group_index is None:
            messagebox.showwarning("提示", "请先选择分组")
            return
        self._commit_choice_fields()
        self._commit_group_fields()
        group = self.projects[self.selected_project_index].required_groups[self.selected_group_index]
        self.group_clipboard = BindingGroup.from_dict(group.to_dict())

    def _paste_group(self) -> None:
        if self.selected_project_index is None:
            messagebox.showwarning("提示", "请先选择项目")
            return
        if self.group_clipboard is None:
            messagebox.showwarning("提示", "请先复制分组")
            return
        self._commit_group_fields()
        self._commit_choice_fields()
        target_project = self.projects[self.selected_project_index]
        new_group = BindingGroup.from_dict(self.group_clipboard.to_dict())
        target_project.required_groups.append(new_group)
        self._refresh_group_list()
        new_index = len(target_project.required_groups) - 1
        self.group_list.selection_clear(0, END)
        self.group_list.selection_set(new_index)
        self.group_list.event_generate("<<ListboxSelect>>")

    def _add_choice(self) -> None:
        if self.selected_project_index is None or self.selected_group_index is None:
            messagebox.showwarning("提示", "请先选择分组")
            return
        self._commit_group_fields()
        self._commit_choice_fields()
        group = self.projects[self.selected_project_index].required_groups[self.selected_group_index]
        group.choices.append(BindingChoice(part_no="", desc=""))
        self._refresh_choice_list()
        new_index = len(group.choices) - 1
        self.choice_tree.selection_set(str(new_index))
        self._on_choice_select()

    def _update_choice(self) -> None:
        if self.selected_choice_index is None:
            return
        current_index = self.selected_choice_index
        self._commit_choice_fields()
        self._refresh_choice_list()
        group = self.projects[self.selected_project_index].required_groups[self.selected_group_index]
        if group.choices:
            target_index = min(current_index, len(group.choices) - 1)
            self.choice_tree.selection_set(str(target_index))
            self._on_choice_select()
        else:
            self._clear_choice_fields()

    def _remove_choice(self) -> None:
        if (
            self.selected_project_index is None
            or self.selected_group_index is None
            or self.selected_choice_index is None
        ):
            return
        group = self.projects[self.selected_project_index].required_groups[self.selected_group_index]
        removed_index = self.selected_choice_index
        if self.selected_choice_index < len(group.choices):
            del group.choices[self.selected_choice_index]
        self._refresh_choice_list()
        if group.choices:
            new_index = min(removed_index, len(group.choices) - 1)
            self.choice_tree.selection_set(str(new_index))
            self._on_choice_select()

    def _commit_all(self) -> None:
        self._commit_choice_fields()
        self._commit_group_fields()
        self._commit_project_fields()

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
        self._commit_all()
        original_projects = self.binding_library.projects
        self.binding_library.projects = self.projects
        try:
            self.binding_library.export_excel(Path(file_path))
        except Exception as exc:
            messagebox.showerror("错误", f"导出失败：{exc}")
            return
        finally:
            self.binding_library.projects = original_projects
        messagebox.showinfo("完成", "导出成功")

    def _save(self) -> None:
        self._commit_all()
        self.binding_library.projects = self.projects
        try:
            self.binding_library.save()
        except Exception as exc:
            messagebox.showerror("错误", f"保存失败：{exc}")
            return
        messagebox.showinfo("完成", "保存成功")


def main() -> None:
    root = Tk()
    Application(root)
    root.mainloop()


if __name__ == "__main__":
    main()
