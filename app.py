from __future__ import annotations

import sys
from dataclasses import dataclass
from pathlib import Path

from PyQt6.QtCore import QTimer, Qt
from PyQt6.QtGui import QFont
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSplitter,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)
from qfluentwidgets import Theme, setTheme, setThemeColor

from paperwrite.docx_loader import document_to_markup
from paperwrite.formatting import FormatOptions, build_docx, build_pdf
from paperwrite.preview import PdfPreviewWidget


@dataclass(frozen=True)
class TemplatePreset:
    school_name: str
    header_text: str
    cover_title: str


TEMPLATE_PRESETS = {
    "苏州大学": TemplatePreset("苏州大学", "苏州大学本科生毕业设计（论文）", "本科毕业设计（论文）"),
    "南京大学": TemplatePreset("南京大学", "南京大学本科毕业论文", "本科毕业论文"),
    "东南大学": TemplatePreset("东南大学", "东南大学本科毕业设计（论文）", "本科毕业设计（论文）"),
    "浙江大学": TemplatePreset("浙江大学", "浙江大学本科毕业设计（论文）", "本科毕业设计（论文）"),
    "复旦大学": TemplatePreset("复旦大学", "复旦大学本科毕业论文", "本科毕业论文"),
}

FIELD_DEFS = [
    ("header_text", "页眉"),
    ("thesis_title", "题目"),
    ("college", "学院"),
    ("year_grade", "年级"),
    ("major", "专业"),
    ("class_name", "班级"),
    ("student_id", "学号"),
    ("author_name", "姓名"),
    ("supervisor", "指导教师"),
    ("supervisor_title", "职称"),
    ("date_text", "提交日期"),
]

APP_NAME = "ThesisFlow"
APP_VERSION = "v0.2"


def app_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent / "_internal"
    return Path(__file__).resolve().parent


def load_app_stylesheet() -> str:
    return (app_base_dir() / "assets" / "desktop.qss").read_text(encoding="utf-8")


class EditorDialog(QDialog):
    def __init__(self, text: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("编辑正文")
        self.setWindowState(Qt.WindowState.WindowMaximized)
        self.setModal(True)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(28, 24, 28, 24)
        layout.setSpacing(16)

        shell = QWidget(self)
        shell.setObjectName("EditorShell")
        shell_layout = QVBoxLayout(shell)
        shell_layout.setContentsMargins(26, 24, 26, 24)
        shell_layout.setSpacing(18)

        chrome = QWidget(shell)
        chrome_layout = QHBoxLayout(chrome)
        chrome_layout.setContentsMargins(0, 0, 0, 0)
        chrome_layout.setSpacing(10)
        for name in ["close", "min", "max"]:
            dot = QLabel(chrome)
            dot.setObjectName(f"EditorDot{name.capitalize()}")
            dot.setFixedSize(12, 12)
            chrome_layout.addWidget(dot)
        chrome_layout.addStretch(1)

        title = QLabel("正文编辑器", shell)
        title.setObjectName("EditorTitle")
        hint = QLabel("支持使用 #、##、### 标记标题层级。保存后会自动刷新 PDF 预览。", shell)
        hint.setWordWrap(True)
        hint.setObjectName("EditorHint")

        self.editor = QTextEdit(shell)
        self.editor.setPlainText(text)
        self.editor.setObjectName("EditorText")

        buttons = QHBoxLayout()
        buttons.addStretch(1)
        cancel_btn = QPushButton("关闭", shell)
        save_btn = QPushButton("保存并返回", shell)
        save_btn.setProperty("primary", True)
        cancel_btn.clicked.connect(self.reject)
        save_btn.clicked.connect(self.accept)
        buttons.addWidget(cancel_btn)
        buttons.addWidget(save_btn)

        shell_layout.addWidget(chrome)
        shell_layout.addWidget(title)
        shell_layout.addWidget(hint)
        shell_layout.addWidget(self.editor, 1)
        shell_layout.addLayout(buttons)
        layout.addWidget(shell, 1)

        self.setStyleSheet(
            """
            QDialog { background: qlineargradient(x1:0,y1:0,x2:1,y2:1, stop:0 #EEF3F9, stop:1 #E8EEF6); }
            QWidget#EditorShell { background: rgba(255,255,255,0.94); border: 1px solid #E3EAF3; border-radius: 26px; }
            QLabel#EditorDotClose, QLabel#EditorDotMin, QLabel#EditorDotMax { border-radius: 6px; }
            QLabel#EditorDotClose { background: #FF5F57; }
            QLabel#EditorDotMin { background: #FEBC2E; }
            QLabel#EditorDotMax { background: #28C840; }
            QLabel#EditorTitle { color: #0F172A; font-size: 28px; font-weight: 700; }
            QLabel#EditorHint { color: #64748B; font-size: 13px; }
            QTextEdit#EditorText { background: white; border: 1px solid #D8DEE8; border-radius: 18px; padding: 16px; font-size: 15px; }
            QPushButton { min-height: 42px; border-radius: 12px; border: 1px solid #D8E1EC; background: white; color: #0F172A; padding: 0 16px; font-weight: 600; }
            QPushButton[primary="true"] { background: qlineargradient(x1:0,y1:0,x2:1,y2:1, stop:0 #0F172A, stop:1 #334155); border: 1px solid #0F172A; color: white; }
            """
        )

    def text(self) -> str:
        return self.editor.toPlainText()


class AboutDialog(QDialog):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle(f"关于 {APP_NAME}")
        self.resize(520, 320)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(22, 22, 22, 22)
        layout.setSpacing(14)

        shell = QWidget(self)
        shell.setObjectName("AboutShell")
        shell_layout = QVBoxLayout(shell)
        shell_layout.setContentsMargins(20, 18, 20, 18)
        shell_layout.setSpacing(12)

        badge = QLabel("T", shell)
        badge.setObjectName("AboutBadge")
        badge.setAlignment(Qt.AlignmentFlag.AlignCenter)
        badge.setFixedSize(42, 42)

        title = QLabel(APP_NAME, shell)
        title.setObjectName("AboutTitle")
        desc = QLabel("一个稳定、简约的论文导入、封面补全、实时预览与导出工作区。", shell)
        desc.setWordWrap(True)
        desc.setObjectName("AboutDesc")

        meta = QWidget(shell)
        meta.setObjectName("AboutMeta")
        meta_layout = QFormLayout(meta)
        meta_layout.setContentsMargins(12, 12, 12, 12)
        meta_layout.setSpacing(10)
        meta_layout.addRow("版本", QLabel(APP_VERSION, meta))
        meta_layout.addRow("技术栈", QLabel("PyQt6 + PyMuPDF + python-docx/report pipeline", meta))
        meta_layout.addRow("开发者", QLabel("Quartz SUDA/UESTC SYR", meta))
        meta_layout.addRow("联系", QLabel("SYRQuartz@gmail.com", meta))

        close_row = QHBoxLayout()
        close_row.addStretch(1)
        close_btn = QPushButton("关闭", shell)
        close_btn.setProperty("primary", True)
        close_btn.clicked.connect(self.accept)
        close_row.addWidget(close_btn)

        shell_layout.addWidget(badge, 0, Qt.AlignmentFlag.AlignLeft)
        shell_layout.addWidget(title)
        shell_layout.addWidget(desc)
        shell_layout.addWidget(meta)
        shell_layout.addLayout(close_row)
        layout.addWidget(shell)

        self.setStyleSheet(
            """
            QDialog { background: qlineargradient(x1:0,y1:0,x2:1,y2:1, stop:0 #EEF3F9, stop:1 #E8EEF6); }
            QWidget#AboutShell { background: rgba(255,255,255,0.94); border: 1px solid #E4EAF3; border-radius: 22px; }
            QLabel#AboutBadge { background: qlineargradient(x1:0,y1:0,x2:1,y2:1, stop:0 #0F172A, stop:1 #334155); color: white; border-radius: 21px; font-size: 18px; font-weight: 700; }
            QLabel#AboutTitle { color: #0F172A; font-size: 24px; font-weight: 700; }
            QLabel#AboutDesc { color: #64748B; font-size: 13px; }
            QWidget#AboutMeta { background: #FBFCFE; border: 1px solid #E7ECF3; border-radius: 16px; }
            QWidget#AboutMeta QLabel { color: #0F172A; font-size: 13px; }
            QFormLayout QLabel { color: #64748B; font-weight: 600; }
            QPushButton { min-height: 38px; border-radius: 10px; border: 1px solid #D8E1EC; background: white; color: #0F172A; padding: 0 14px; font-weight: 600; }
            QPushButton[primary="true"] { background: qlineargradient(x1:0,y1:0,x2:1,y2:1, stop:0 #0F172A, stop:1 #334155); border: 1px solid #0F172A; color: white; }
            """
        )


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.resize(1500, 920)
        self.setMinimumSize(1180, 760)
        self.setFont(QFont("Microsoft YaHei UI", 10))

        self.source_path: Path | None = None
        self.output_path: Path | None = None
        self.preview_path = Path.cwd() / "_preview.pdf"
        self.document_text = "# 在这里输入或导入论文正文\n\n支持使用 #、##、### 表示标题层级。"
        self.template_name = next(iter(TEMPLATE_PRESETS))
        self.preview_timer = QTimer(self)
        self.preview_timer.setSingleShot(True)
        self.preview_timer.timeout.connect(self.refresh_preview)

        self._build_ui()
        self._apply_styles()
        self._apply_template(self.template_name)
        self._update_summary()
        self.refresh_preview(show_error=False)

    def _build_ui(self) -> None:
        root = QWidget(self)
        root.setObjectName("Root")
        self.setCentralWidget(root)

        layout = QVBoxLayout(root)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(14)

        shell = QWidget(root)
        shell.setObjectName("WindowShell")
        shell_layout = QVBoxLayout(shell)
        shell_layout.setContentsMargins(14, 14, 14, 14)
        shell_layout.setSpacing(14)
        layout.addWidget(shell, 1)

        topbar = QWidget(shell)
        topbar.setObjectName("TopBar")
        topbar_layout = QHBoxLayout(topbar)
        topbar_layout.setContentsMargins(6, 2, 6, 2)
        topbar_layout.setSpacing(12)

        brand_shell = QWidget(topbar)
        brand_shell.setObjectName("BrandShell")
        brand_layout = QHBoxLayout(brand_shell)
        brand_layout.setContentsMargins(0, 0, 0, 0)
        brand_layout.setSpacing(12)
        brand_badge = QLabel("T", brand_shell)
        brand_badge.setObjectName("BrandBadge")
        brand_badge.setAlignment(Qt.AlignmentFlag.AlignCenter)
        brand_badge.setFixedSize(42, 42)
        brand_layout.addWidget(brand_badge)
        brand_copy = QVBoxLayout()
        brand_copy.setContentsMargins(0, 0, 0, 0)
        brand_copy.setSpacing(2)
        header = QLabel(APP_NAME, brand_shell)
        header.setObjectName("BrandTitle")
        sub = QLabel("Native workspace", brand_shell)
        sub.setObjectName("BrandSubtitle")
        brand_copy.addWidget(header)
        brand_copy.addWidget(sub)
        brand_layout.addLayout(brand_copy)
        topbar_layout.addWidget(brand_shell, 0)

        topbar_layout.addStretch(1)
        self.status_chip = QLabel("原生预览已就绪", topbar)
        self.status_chip.setObjectName("StatusChip")
        topbar_layout.addWidget(self.status_chip, 0)

        self.about_btn = QPushButton("关于", topbar)
        self.about_btn.setObjectName("AboutButton")
        self.about_btn.clicked.connect(self.show_about)
        topbar_layout.addWidget(self.about_btn, 0)
        shell_layout.addWidget(topbar)

        splitter = QSplitter(Qt.Orientation.Horizontal, shell)
        splitter.setChildrenCollapsible(False)
        splitter.setHandleWidth(10)
        shell_layout.addWidget(splitter, 1)

        left_panel = QWidget(splitter)
        left_panel.setObjectName("SidebarPanel")
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(12, 12, 12, 12)
        left_layout.setSpacing(10)

        scroll = QScrollArea(left_panel)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QScrollArea.Shape.NoFrame)
        scroll.setObjectName("SidebarScroll")
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        left_layout.addWidget(scroll, 1)

        scroll_body = QWidget()
        scroll_body.setObjectName("ScrollBody")
        scroll.setWidget(scroll_body)
        body = QVBoxLayout(scroll_body)
        body.setContentsMargins(0, 4, 0, 4)
        body.setSpacing(14)

        project_card = self._make_card("项目设置", "导入、输出和编辑动作集中在一起。")
        project_form = QFormLayout(project_card.layout().itemAt(2).widget())
        project_form.setContentsMargins(0, 0, 0, 0)
        project_form.setSpacing(12)
        project_form.setLabelAlignment(Qt.AlignmentFlag.AlignLeft)
        project_form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.AllNonFixedFieldsGrow)
        project_form.setRowWrapPolicy(QFormLayout.RowWrapPolicy.WrapLongRows)

        self.template_combo = QComboBox(project_card)
        self.template_combo.addItems(list(TEMPLATE_PRESETS.keys()))
        self.template_combo.currentTextChanged.connect(self._apply_template)
        project_form.addRow("模板", self.template_combo)

        self.source_line = QLineEdit(project_card)
        self.source_line.setReadOnly(True)
        self.source_line.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        src_row = QWidget(project_card)
        src_layout = QHBoxLayout(src_row)
        src_layout.setContentsMargins(0, 0, 0, 0)
        src_layout.setSpacing(8)
        src_layout.addWidget(self.source_line, 1)
        self.choose_source_btn = QPushButton("导入", src_row)
        self.choose_source_btn.setFixedWidth(74)
        self.choose_source_btn.clicked.connect(self.choose_source_file)
        src_layout.addWidget(self.choose_source_btn)
        project_form.addRow("源文件", src_row)

        self.output_line = QLineEdit(project_card)
        self.output_line.setReadOnly(True)
        self.output_line.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        out_row = QWidget(project_card)
        out_layout = QHBoxLayout(out_row)
        out_layout.setContentsMargins(0, 0, 0, 0)
        out_layout.setSpacing(8)
        out_layout.addWidget(self.output_line, 1)
        self.choose_output_btn = QPushButton("选择", out_row)
        self.choose_output_btn.setFixedWidth(74)
        self.choose_output_btn.clicked.connect(self.choose_output_file)
        out_layout.addWidget(self.choose_output_btn)
        project_form.addRow("输出位置", out_row)

        project_actions = QWidget(project_card)
        actions_layout = QHBoxLayout(project_actions)
        actions_layout.setContentsMargins(0, 0, 0, 0)
        actions_layout.setSpacing(8)
        self.edit_btn = QPushButton("编辑正文", project_actions)
        self.edit_btn.clicked.connect(self.open_editor)
        self.export_btn = QPushButton("导出 PDF / DOCX", project_actions)
        self.export_btn.setProperty("primary", True)
        self.export_btn.clicked.connect(self.save_output)
        actions_layout.addWidget(self.edit_btn)
        actions_layout.addWidget(self.export_btn)
        project_form.addRow("操作", project_actions)
        body.addWidget(project_card)

        cover_card = self._make_card("封面字段", "字段修改后自动刷新预览。")
        cover_form = QFormLayout(cover_card.layout().itemAt(2).widget())
        cover_form.setContentsMargins(0, 0, 0, 0)
        cover_form.setSpacing(12)
        cover_form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.AllNonFixedFieldsGrow)
        cover_form.setRowWrapPolicy(QFormLayout.RowWrapPolicy.WrapLongRows)

        self.include_cover = QCheckBox("启用封面", cover_card)
        self.include_cover.setChecked(True)
        self.include_cover.stateChanged.connect(lambda _: self._update_cover_state())
        cover_form.addRow("封面", self.include_cover)

        self.field_inputs: dict[str, QLineEdit] = {}
        for key, label in FIELD_DEFS:
            line = QLineEdit(cover_card)
            line.textEdited.connect(lambda _=None: self.queue_preview())
            self.field_inputs[key] = line
            cover_form.addRow(label, line)

        body.addWidget(cover_card)
        body.addStretch(1)

        right_panel = QWidget(splitter)
        right_panel.setObjectName("PreviewPanel")
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(16, 16, 16, 16)
        right_layout.setSpacing(14)

        preview_head = QWidget(right_panel)
        preview_head.setObjectName("PreviewHeader")
        preview_head_layout = QHBoxLayout(preview_head)
        preview_head_layout.setContentsMargins(0, 0, 0, 0)
        preview_head_layout.setSpacing(12)
        title = QLabel("实时 PDF 预览", preview_head)
        title.setObjectName("PreviewTitle")
        preview_head_layout.addWidget(title, 1)
        self.refresh_btn = QPushButton("刷新预览", preview_head)
        self.refresh_btn.clicked.connect(self.refresh_preview)
        preview_head_layout.addWidget(self.refresh_btn)
        right_layout.addWidget(preview_head)

        preview_stage = QWidget(right_panel)
        preview_stage.setObjectName("PreviewStage")
        preview_stage_layout = QVBoxLayout(preview_stage)
        preview_stage_layout.setContentsMargins(16, 16, 16, 16)
        preview_stage_layout.setSpacing(0)
        self.preview = PdfPreviewWidget(preview_stage)
        self.preview.setObjectName("PdfPreview")
        preview_stage_layout.addWidget(self.preview, 1)
        right_layout.addWidget(preview_stage, 1)

        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([400, 1080])

    def _make_card(self, title: str, desc: str) -> QWidget:
        card = QWidget()
        card.setObjectName("Card")
        title_label = QLabel(title, card)
        title_label.setObjectName("CardTitle")
        desc_label = QLabel(desc, card)
        desc_label.setObjectName("CardDesc")
        desc_label.setWordWrap(True)
        inner = QWidget(card)
        inner.setObjectName("CardInner")
        outer = QVBoxLayout(card)
        outer.setContentsMargins(16, 16, 16, 16)
        outer.setSpacing(8)
        outer.addWidget(title_label)
        outer.addWidget(desc_label)
        outer.addWidget(inner)
        return card

    def _apply_styles(self) -> None:
        self.setStyleSheet(load_app_stylesheet())

    def _field_value(self, key: str) -> str:
        return self.field_inputs[key].text().strip()

    def _apply_template(self, name: str) -> None:
        self.template_name = name
        preset = TEMPLATE_PRESETS[name]
        header = self.field_inputs.get("header_text")
        if header and (not header.text().strip() or header.text().strip() in [p.header_text for p in TEMPLATE_PRESETS.values()]):
            header.setText(preset.header_text)
        self.queue_preview(250)

    def _update_cover_state(self) -> None:
        self.queue_preview(250)

    def _collect_options(self) -> FormatOptions:
        preset = TEMPLATE_PRESETS[self.template_name]
        return FormatOptions(
            school_name=preset.school_name,
            header_text=self._field_value("header_text") or preset.header_text,
            thesis_title=self._field_value("thesis_title"),
            college=self._field_value("college"),
            year_grade=self._field_value("year_grade"),
            major=self._field_value("major"),
            class_name=self._field_value("class_name"),
            student_id=self._field_value("student_id"),
            author_name=self._field_value("author_name"),
            supervisor=self._field_value("supervisor"),
            supervisor_title=self._field_value("supervisor_title"),
            date_text=self._field_value("date_text"),
            include_cover=self.include_cover.isChecked(),
            source_docx_path=str(self.source_path) if self.source_path else "",
        )

    def _update_summary(self) -> None:
        pass

    def queue_preview(self, delay_ms: int = 420) -> None:
        self.status_chip.setText("等待刷新预览")
        self.preview_timer.start(delay_ms)

    def refresh_preview(self, show_error: bool = True) -> None:
        try:
            self.preview_timer.stop()
            self.status_chip.setText("正在刷新预览")
            build_pdf(self.document_text, self._collect_options(), self.preview_path)
            self.preview.load_pdf(str(self.preview_path))
            self.preview.fit_width()
            status = f"预览已更新: {self.preview_path.name}"
            self.status_chip.setText(status)
        except Exception as exc:
            self.status_chip.setText("预览生成失败")
            if show_error:
                QMessageBox.critical(self, "操作失败", f"预览生成失败: {exc}")

    def choose_source_file(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(self, "选择论文源文件", str(Path.cwd()), "Word 文档 (*.docx)")
        if not file_path:
            return
        self.source_path = Path(file_path)
        self.source_line.setText(file_path)
        try:
            markup, suggested_title = document_to_markup(file_path)
        except Exception as exc:
            QMessageBox.critical(self, "操作失败", f"导入失败: {exc}")
            return
        self.document_text = markup
        if not self.field_inputs["thesis_title"].text().strip():
            self.field_inputs["thesis_title"].setText(suggested_title)
        if self.output_path is None:
            self.output_path = self.source_path.with_name(f"{self.source_path.stem}_formatted.pdf")
            self.output_line.setText(str(self.output_path))
        self._update_summary()
        self.refresh_preview()

    def choose_output_file(self) -> None:
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "选择导出文件名",
            str(self.output_path or Path.cwd() / "formatted.pdf"),
            "导出文件 (*.pdf *.docx)",
        )
        if not file_path:
            return
        self.output_path = Path(file_path)
        self.output_line.setText(file_path)
        self._update_summary()

    def show_about(self) -> None:
        AboutDialog(self).exec()

    def open_editor(self) -> None:
        dialog = EditorDialog(self.document_text, self)
        if dialog.exec():
            self.document_text = dialog.text()
            self._update_summary()
            self.refresh_preview()

    def _export_targets(self) -> tuple[Path, Path]:
        if self.output_path is None:
            raise ValueError("请先选择输出文件位置。")
        base = self.output_path.with_suffix("") if self.output_path.suffix.lower() in {".pdf", ".docx"} else self.output_path
        return base.with_suffix(".pdf"), base.with_suffix(".docx")

    def save_output(self) -> None:
        if self.output_path is None:
            QMessageBox.critical(self, "操作失败", "请先选择输出文件位置。")
            return
        try:
            pdf_path, docx_path = self._export_targets()
            options = self._collect_options()
            build_pdf(self.document_text, options, pdf_path)
            build_docx(self.document_text, options, docx_path)
            self.output_path = pdf_path
            self.output_line.setText(str(pdf_path))
            self._update_summary()
            QMessageBox.information(self, "导出完成", f"PDF: {pdf_path.name}\nWord: {docx_path.name}")
        except Exception as exc:
            QMessageBox.critical(self, "操作失败", f"导出失败: {exc}")


def main() -> int:
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    setTheme(Theme.LIGHT)
    setThemeColor("#2563EB")
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
