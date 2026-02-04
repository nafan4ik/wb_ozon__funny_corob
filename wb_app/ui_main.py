import os
from pathlib import Path
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QLineEdit,
    QFileDialog, QTextEdit, QSpinBox, QMessageBox
)
from PySide6.QtWidgets import QCheckBox
from PySide6.QtCore import QSettings
from PySide6.QtCore import Qt

from engine.run_build import run

def _get_downloads_dir() -> Path | None:
    userprofile = os.environ.get("USERPROFILE")
    if not userprofile:
        return None

    candidates = [
        Path(userprofile) / "Downloads",
        Path(userprofile) / "Загрузки",
    ]
    for p in candidates:
        if p.exists() and p.is_dir():
            return p
    return None


def pick_default_pdfs_from_downloads(
    default_wb: str,
    default_ozon: str
) -> tuple[str, str]:
    """
    Логика:
      - смотрим ТОЛЬКО 2 самых новых файла из Downloads
      - если имя содержит 'assembly_list' -> OZON
      - если имя содержит 4 дефиса -> WB
      - если найден только один (WB или OZON) — второй оставляем дефолт
      - если оба не совпали — оставляем дефолт обоим
    """
    downloads = _get_downloads_dir()
    if not downloads:
        return default_wb, default_ozon

    # берём только файлы (можно ограничить PDF, но ты не просил — если надо, добавим фильтр)
    files = [p for p in downloads.iterdir() if p.is_file()]
    if not files:
        return default_wb, default_ozon

    # сортируем по времени изменения (самые новые сверху)
    files.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    last_two = files[:2]

    wb_candidate = None
    ozon_candidate = None

    for p in last_two:
        name = p.name.lower()

        # OZON: assembly_list
        if "assembly_list" in name:
            ozon_candidate = str(p)
            continue

        # WB: 4 дефиса в имени (без расширения)
        stem = p.stem
        if stem.count("-") == 4:
            wb_candidate = str(p)
            continue

    wb_path = wb_candidate if wb_candidate else default_wb
    ozon_path = ozon_candidate if ozon_candidate else default_ozon
    return wb_path, ozon_path


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.settings = QSettings("korob", "wb_app")
        default_wb = r"C:\korob\wb.pdf"
        default_oz = r"C:\korob\ozon.pdf"
        wb_auto, oz_auto = pick_default_pdfs_from_downloads(default_wb, default_oz)
        self.setWindowTitle("...")
        layout = QVBoxLayout()




        self.wb_pdf = QLineEdit(wb_auto)
        self.oz_pdf = QLineEdit(oz_auto)

        self.setWindowTitle("Korob WB/OZON сборка (минималка)")

        layout = QVBoxLayout()

        # WB PDF
        btn_wb = QPushButton("Выбрать WB PDF")
        btn_wb.clicked.connect(self.pick_wb_pdf)
        layout.addLayout(self.row("WB PDF:", self.wb_pdf, btn_wb))

        # WB singles dir
        self.wb_dir = QLineEdit(r"C:\korob\wb_makets_po_odnomu")
        btn_wb_dir = QPushButton("Выбрать папку WB одиночных")
        btn_wb_dir.clicked.connect(self.pick_wb_dir)
        layout.addLayout(self.row("WB макеты (по одному):", self.wb_dir, btn_wb_dir))

        # OZON PDF
        btn_oz = QPushButton("Выбрать OZON PDF")
        btn_oz.clicked.connect(self.pick_oz_pdf)
        layout.addLayout(self.row("OZON PDF:", self.oz_pdf, btn_oz))

        # OZON makets dir
        self.oz_dir = QLineEdit(r"C:\korob\ozon_makets_po_odnomu")
        btn_oz_dir = QPushButton("Выбрать папку OZON макетов")
        btn_oz_dir.clicked.connect(self.pick_oz_dir)
        layout.addLayout(self.row("OZON макеты:", self.oz_dir, btn_oz_dir))

        # GAP
        gap_row = QHBoxLayout()
        gap_row.addWidget(QLabel("Отступ между макетами (GAP):"))
        self.gap = QSpinBox()
        self.gap.setRange(0, 30)
        self.gap.setValue(4)
        gap_row.addWidget(self.gap)
        gap_row.addStretch(1)
        layout.addLayout(gap_row)

        self.cb_goblin = QCheckBox("Показывать гифку и звук при ошибках WB")
        self.cb_goblin.setChecked(self.settings.value("goblin_enabled", True, bool))
        self.cb_goblin.stateChanged.connect(
            lambda v: self.settings.setValue("goblin_enabled", bool(v))
        )
        layout.addWidget(self.cb_goblin)

        # Run
        self.btn_run = QPushButton("СОБРАТЬ ПОСТАВКУ")
        self.btn_run.clicked.connect(self.on_run)
        layout.addWidget(self.btn_run)

        # Log
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        layout.addWidget(self.log, 1)

        self.setLayout(layout)

    def row(self, title, lineedit, button):
        r = QHBoxLayout()
        r.addWidget(QLabel(title))
        r.addWidget(lineedit, 1)
        r.addWidget(button)
        return r

    def pick_wb_pdf(self):
        p, _ = QFileDialog.getOpenFileName(self, "WB PDF", "", "PDF (*.pdf)")
        if p:
            self.wb_pdf.setText(p)

    def pick_wb_dir(self):
        p = QFileDialog.getExistingDirectory(self, "WB одиночные макеты")
        if p:
            self.wb_dir.setText(p)

    def pick_oz_pdf(self):
        p, _ = QFileDialog.getOpenFileName(self, "OZON PDF", "", "PDF (*.pdf)")
        if p:
            self.oz_pdf.setText(p)

    def pick_oz_dir(self):
        p = QFileDialog.getExistingDirectory(self, "OZON макеты")
        if p:
            self.oz_dir.setText(p)

    def on_run(self):
        self.btn_run.setEnabled(False)
        self.log.append("Запуск...\n")

        try:
            result = run(
                wb_pdf=self.wb_pdf.text().strip(),
                wb_singles_dir=self.wb_dir.text().strip(),
                ozon_pdf=self.oz_pdf.text().strip(),
                ozon_singles_dir=self.oz_dir.text().strip(),  # <-- теперь это папка OZON одиночных
                gap=int(self.gap.value()),
            )

            self.log.append("✅ ГОТОВО")
            self.log.append(f"Папка результата: {result['dest_dir']}")
            self.log.append("")

            # WB
            self.log.append("WB:")
            self.log.append(f"  строк(штук) = {result['wb_rows']}")
            self.log.append(f"  одиночных в пуле = {result['wb_pool']}")
            self.log.append(f"  листов собрано = {result['wb_sheets']}")
            self.log.append(f"  папка листов = {result['wb_sheets_dir']}")
            if result["wb_not_found"]:
                self.log.append("  ❗ не найдены макеты:")
                for a in result["wb_not_found"]:
                    self.log.append(f"    - {a}")
            self.log.append("")

            # OZON
            self.log.append("OZON:")
            self.log.append(f"  строк таблицы = {result['oz_rows']}")
            self.log.append(f"  одиночных в пуле = {result['oz_pool']}")
            self.log.append(f"  листов собрано = {result['oz_sheets']}")
            self.log.append(f"  папка листов = {result['ozon_sheets_dir']}")
            if result["oz_not_found"]:
                self.log.append("  ❗ не найдены макеты:")
                for a in result["oz_not_found"]:
                    self.log.append(f"    - {a}")
            self.log.append("")

            self.log.append(f"Очистка корня поставки: удалено файлов = {result['cleanup_removed']}")
            self.log.append("")

            QMessageBox.information(self, "Готово", f"Сборка завершена.\n{result['dest_dir']}")

            # --- Гоблин: если WB макеты не найдены ---
            if result.get("wb_not_found"):
                if self.settings.value("goblin_enabled", True, bool):
                    from goblin_dialog import GoblinDialog

                    self.goblin_dialog = GoblinDialog(self)
                    self.goblin_dialog.show()

        except Exception as e:
            self.log.append(f"❌ Ошибка: {e}\n")
            QMessageBox.critical(self, "Ошибка", str(e))
        finally:
            self.btn_run.setEnabled(True)
