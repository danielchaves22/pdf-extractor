"""Interface dedicada ao módulo de Ficha Financeira."""

from __future__ import annotations

from datetime import date
from pathlib import Path
from typing import List

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QTextCursor
from PyQt6.QtWidgets import (
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QTextEdit,
    QVBoxLayout,
    QWidget,
    QComboBox,
)

from processors import FichaFinanceiraProcessor


class FichaFinanceiraWindow(QMainWindow):
    """Janela principal para geração dos CSVs da ficha financeira."""

    MONTH_NAMES = [
        "Janeiro",
        "Fevereiro",
        "Março",
        "Abril",
        "Maio",
        "Junho",
        "Julho",
        "Agosto",
        "Setembro",
        "Outubro",
        "Novembro",
        "Dezembro",
    ]

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Ficha Financeira - Geração de CSVs")
        self.setFixedSize(750, 560)

        self._pdf_paths: List[Path] = []
        self._processor = FichaFinanceiraProcessor(log_callback=self.add_log_message)

        self._build_interface()

    # ------------------------------------------------------------------
    # Interface
    # ------------------------------------------------------------------
    def _build_interface(self) -> None:
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout(central_widget)
        layout.setSpacing(12)

        header = QLabel(
            "Selecione o intervalo desejado e os PDFs da ficha financeira para gerar os CSVs (PROVENTOS e ADIC. INSALUBRIDADE PAGO)."
        )
        header.setWordWrap(True)
        header.setAlignment(Qt.AlignmentFlag.AlignLeft)
        layout.addWidget(header)

        layout.addWidget(self._create_period_box())
        layout.addWidget(self._create_files_box())
        layout.addWidget(self._create_actions_box())
        layout.addWidget(self._create_log_box(), stretch=1)

    def _create_period_box(self) -> QGroupBox:
        box = QGroupBox("🗓️ Período para geração dos CSVs")
        box_layout = QHBoxLayout(box)

        self.start_month_combo = QComboBox()
        self.start_month_combo.addItems(self.MONTH_NAMES)
        self.start_year_spin = QSpinBox()
        self.start_year_spin.setRange(1900, 2100)

        self.end_month_combo = QComboBox()
        self.end_month_combo.addItems(self.MONTH_NAMES)
        self.end_year_spin = QSpinBox()
        self.end_year_spin.setRange(1900, 2100)

        today = date.today()
        current_year = today.year
        self.start_year_spin.setValue(current_year)
        self.end_year_spin.setValue(current_year)
        self.start_month_combo.setCurrentIndex(0)
        self.end_month_combo.setCurrentIndex(today.month - 1)

        box_layout.addWidget(QLabel("Início:"))
        box_layout.addWidget(self.start_month_combo)
        box_layout.addWidget(self.start_year_spin)
        box_layout.addSpacing(20)
        box_layout.addWidget(QLabel("Fim:"))
        box_layout.addWidget(self.end_month_combo)
        box_layout.addWidget(self.end_year_spin)
        box_layout.addStretch()

        return box

    def _create_files_box(self) -> QGroupBox:
        box = QGroupBox("📎 Arquivos PDF da ficha financeira")
        box_layout = QVBoxLayout(box)

        buttons_layout = QHBoxLayout()
        add_button = QPushButton("Adicionar PDFs")
        add_button.clicked.connect(self._on_add_files)
        remove_button = QPushButton("Remover Selecionados")
        remove_button.clicked.connect(self._on_remove_selected)
        clear_button = QPushButton("Limpar Lista")
        clear_button.clicked.connect(self._on_clear_list)

        buttons_layout.addWidget(add_button)
        buttons_layout.addWidget(remove_button)
        buttons_layout.addWidget(clear_button)
        buttons_layout.addStretch()

        self.files_list = QListWidget()
        self.files_list.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)

        box_layout.addLayout(buttons_layout)
        box_layout.addWidget(self.files_list)

        return box

    def _create_actions_box(self) -> QGroupBox:
        box = QGroupBox("🚀 Processamento")
        box_layout = QVBoxLayout(box)

        self.output_label = QLabel("Nenhum processamento realizado ainda.")
        self.output_label.setWordWrap(True)

        self.generate_button = QPushButton("Gerar CSVs da Ficha")
        self.generate_button.clicked.connect(self._on_generate)

        box_layout.addWidget(self.output_label)
        box_layout.addWidget(self.generate_button)

        return box

    def _create_log_box(self) -> QGroupBox:
        box = QGroupBox("📝 Registros da execução")
        box_layout = QVBoxLayout(box)

        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setMinimumHeight(180)

        box_layout.addWidget(self.log_output)
        return box

    # ------------------------------------------------------------------
    # Eventos
    # ------------------------------------------------------------------
    def _on_add_files(self) -> None:
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Selecione os PDFs da ficha financeira",
            "",
            "Arquivos PDF (*.pdf)",
        )
        if not files:
            return

        for file_path in files:
            path = Path(file_path)
            if path not in self._pdf_paths:
                self._pdf_paths.append(path)
                self.files_list.addItem(str(path))

        self.add_log_message(f"✅ {len(files)} arquivo(s) adicionados.")

    def _on_remove_selected(self) -> None:
        selected_items = self.files_list.selectedItems()
        if not selected_items:
            return

        for item in selected_items:
            path = Path(item.text())
            if path in self._pdf_paths:
                self._pdf_paths.remove(path)
            row = self.files_list.row(item)
            self.files_list.takeItem(row)

        self.add_log_message("🗑️ Arquivos removidos da lista.")

    def _on_clear_list(self) -> None:
        self._pdf_paths.clear()
        self.files_list.clear()
        self.add_log_message("🧹 Lista de PDFs limpa.")

    def _on_generate(self) -> None:
        if not self._pdf_paths:
            QMessageBox.warning(
                self,
                "Nenhum PDF selecionado",
                "Selecione pelo menos um PDF para gerar os arquivos.",
            )
            return

        start_month = self.start_month_combo.currentIndex() + 1
        start_year = self.start_year_spin.value()
        end_month = self.end_month_combo.currentIndex() + 1
        end_year = self.end_year_spin.value()

        start_period = date(start_year, start_month, 1)
        end_period = date(end_year, end_month, 1)

        if start_period > end_period:
            QMessageBox.warning(
                self,
                "Período inválido",
                "O mês inicial precisa ser anterior ou igual ao mês final.",
            )
            return

        output_dir = self._pdf_paths[0].parent

        self.generate_button.setEnabled(False)
        self.add_log_message("🚧 Iniciando processamento da ficha financeira...")

        try:
            result = self._processor.generate_csvs(
                self._pdf_paths,
                start_period,
                end_period,
                output_dir,
            )
            outputs = result.get("outputs", [])
            if outputs:
                lines = [f"{item['label']}: {item['path']}" for item in outputs]
                joined = "\n".join(lines)
                self.output_label.setText(f"Arquivos gerados:\n{joined}")
                for item in outputs:
                    self.add_log_message(
                        f"📁 {item['label']} salvo em {item['path']}"
                    )
            else:
                self.output_label.setText("Nenhum arquivo foi gerado.")
            self.add_log_message("✅ Processamento concluído com sucesso.")

            if outputs:
                dialog_lines = [
                    "Os arquivos foram gerados em:",
                    *[f"• {item['label']}: {item['path']}" for item in outputs],
                ]
            else:
                dialog_lines = [
                    "Nenhum arquivo foi gerado para o período informado.",
                ]
            QMessageBox.information(
                self,
                "Processamento concluído",
                "\n".join(dialog_lines),
            )
        except Exception as exc:  # noqa: BLE001
            self.add_log_message(f"❌ Erro durante o processamento: {exc}")
            QMessageBox.critical(
                self,
                "Erro ao gerar o arquivo",
                f"Não foi possível gerar o arquivo:\n{exc}",
            )
        finally:
            self.generate_button.setEnabled(True)

    # ------------------------------------------------------------------
    # Logs
    # ------------------------------------------------------------------
    def add_log_message(self, message: str) -> None:
        self.log_output.append(message)
        self.log_output.moveCursor(QTextCursor.MoveOperation.End)

