#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF para Excel Desktop App - Interface PyQt6
============================================

Interface gráfica moderna usando PyQt6 que utiliza o módulo
pdf_processor_core.py para toda a lógica de processamento.

Versão 4.0.1 - PROCESSAMENTO UNIFICADO:
- Performance 10-20x superior
- Threading nativo com signals/slots thread-safe
- Virtualização automática de listas grandes
- Drag & Drop nativo do PyQt6
- Interface responsiva e moderna
- Eliminação de polling manual
- Splash Screen profissional
- Sistema de atenção para códigos duplicados (01003601 + 01003602)
- NOVO: Processamento sempre via ThreadPoolExecutor (simplificação)

Funcionalidades v4.0.1:
- Seleção múltipla de PDFs com interface otimizada
- Processamento paralelo unificado com comunicação thread-safe
- Updates em tempo real sem latência
- Histórico virtualizado para performance máxima
- Styling moderno com QSS
- Splash screen com progresso de carregamento
- Sistema automático para duplicidades de PREMIO PROD. MENSAL
- Arquitetura simplificada (sempre threads, zero inconsistência)

Dependências:
pip install PyQt6

Como usar:
python desktop_app.py

Autor: Sistema de Extração Automatizada
Data: 2025
"""

import sys
import os
import json
import uuid
import subprocess
import threading
import time
from pathlib import Path
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import List, Dict, Optional
from dataclasses import dataclass

# PyQt6 imports
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTabWidget, QLabel, QPushButton, QLineEdit, QTextEdit, QProgressBar,
    QListWidget, QListWidgetItem, QFrame, QDialog, QScrollArea,
    QCheckBox, QSpinBox, QFileDialog, QMessageBox, QSizePolicy,
    QSplitter, QGroupBox, QFormLayout, QComboBox, QSplashScreen
)
from PyQt6.QtCore import (
    Qt, QThread, pyqtSignal, QTimer, QSize, pyqtSlot, QMimeData
)
from PyQt6.QtGui import (
    QFont, QTextCursor, QPalette, QColor, QDragEnterEvent, QDropEvent, QAction,
    QPixmap, QPainter, QLinearGradient
)

# Importa o processador core
try:
    from pdf_processor_core import PDFProcessorCore
except ImportError:
    print("ERRO: Módulo pdf_processor_core.py não encontrado!")
    print("Certifique-se de que o arquivo pdf_processor_core.py está na mesma pasta.")
    sys.exit(1)

class SplashScreen(QSplashScreen):
    """Splash screen moderna com progresso de carregamento"""
    
    def __init__(self):
        # Cria pixmap simples para splash screen
        splash_pixmap = self.create_splash_pixmap()
        super().__init__(splash_pixmap)
        
        self.setWindowFlags(Qt.WindowType.SplashScreen | Qt.WindowType.WindowStaysOnTopHint)
        
        # Define fonte maior e em negrito para as mensagens de progresso
        progress_font = QFont("Arial", 13, QFont.Weight.Bold)
        self.setFont(progress_font)
        
        # Timer para simular progresso
        self.progress_value = 0
        self.progress_timer = QTimer()
        self.progress_timer.timeout.connect(self.update_progress)
        
        # Centraliza na tela
        self.center_on_screen()
        
    def create_splash_pixmap(self):
        """Cria pixmap customizado para a splash screen"""
        width, height = 480, 220  # Altura reduzida
        pixmap = QPixmap(width, height)
        pixmap.fill(QColor("#1e1e1e"))
        
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        # Fundo com gradiente
        gradient = QLinearGradient(0, 0, 0, height)
        gradient.setColorAt(0, QColor("#1e1e1e"))
        gradient.setColorAt(1, QColor("#2b2b2b"))
        painter.fillRect(pixmap.rect(), gradient)
        
        # Borda
        painter.setPen(QColor("#1f538d"))
        painter.drawRoundedRect(5, 5, width-10, height-10, 10, 10)
        
        # Título principal
        painter.setPen(QColor("#ffffff"))
        title_font = QFont("Arial", 22, QFont.Weight.Bold)
        painter.setFont(title_font)
        painter.drawText(20, 50, width-40, 40, Qt.AlignmentFlag.AlignCenter, "📄 Processamento de Folha")
        
        # Subtítulo (com espaçamento aumentado)
        subtitle_font = QFont("Arial", 14)
        painter.setFont(subtitle_font)
        painter.setPen(QColor("#aaaaaa"))
        painter.drawText(20, 100, width-40, 25, Qt.AlignmentFlag.AlignCenter, "Sistema de Automatização v4.0.1")
        
        painter.end()
        return pixmap
    
    def center_on_screen(self):
        """Centraliza splash screen na tela"""
        try:
            screen = QApplication.primaryScreen().geometry()
            splash_size = self.size()
            x = (screen.width() - splash_size.width()) // 2
            y = (screen.height() - splash_size.height()) // 2
            self.move(x, y)
        except Exception:
            # Fallback se não conseguir centralizar
            self.move(400, 300)
    
    def start_loading(self):
        """Inicia processo de carregamento simulado"""
        self.progress_value = 0
        self.progress_timer.start(30)  # Atualiza a cada 30ms para animação mais suave
        self.show()
        self.showMessage("Inicializando aplicação...", Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignBottom, QColor("#ffffff"))
        QApplication.processEvents()  # Força renderização
    
    def update_progress(self):
        """Atualiza progresso da splash screen"""
        if self.progress_value < 100:
            # Incremento variável para parecer mais natural
            if self.progress_value < 30:
                self.progress_value += 1.5  # Início mais rápido
            elif self.progress_value < 70:
                self.progress_value += 1    # Meio mais lento
            else:
                self.progress_value += 0.8  # Final mais lento
                
            current_progress = int(self.progress_value)
            
            # Atualiza mensagem baseado no progresso
            if self.progress_value < 15:
                message = f"{current_progress}% - Inicializando PyQt6..."
            elif self.progress_value < 30:
                message = f"{current_progress}% - Carregando dependências..."
            elif self.progress_value < 50:
                message = f"{current_progress}% - Configurando interface..."
            elif self.progress_value < 70:
                message = f"{current_progress}% - Preparando sistema..."
            elif self.progress_value < 85:
                message = f"{current_progress}% - Finalizando..."
            elif self.progress_value < 95:
                message = f"{current_progress}% - Quase pronto..."
            else:
                message = f"{current_progress}% - Carregando..."
                
            self.showMessage(message, Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignBottom, QColor("#ffffff"))
        else:
            self.progress_timer.stop()
            self.showMessage("100% - Aplicação pronta!", Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignBottom, QColor("#2cc985"))
    
    def set_status(self, message):
        """Define status personalizado"""
        current_val = int(self.progress_value)
        if "✅" in message:
            display_message = f"{current_val}% - {message}"
            color = QColor("#2cc985")
        elif "❌" in message:
            display_message = f"{current_val}% - {message}"
            color = QColor("#f44336")
        else:
            display_message = f"{current_val}% - {message}"
            color = QColor("#ffffff")
            
        self.showMessage(display_message, Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignBottom, color)
        QApplication.processEvents()  # Força atualização imediata
    
    def finish_loading(self, main_window):
        """Finaliza carregamento e mostra janela principal"""
        self.showMessage("100% - Aplicação pronta! Abrindo...", Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignBottom, QColor("#2cc985"))
        QApplication.processEvents()
        
        # Pequena pausa para mostrar "100%" e mensagem final
        QTimer.singleShot(500, lambda: self.finish(main_window))

# Estilo escuro moderno
DARK_STYLE = """
QMainWindow {
    background-color: #1e1e1e;
    color: white;
}

QWidget {
    background-color: #1e1e1e;
    color: white;
}

QPushButton {
    background-color: #1f538d;
    border: none;
    border-radius: 6px;
    padding: 8px 16px;
    color: white;
    font-weight: bold;
    min-height: 20px;
}

QPushButton:hover {
    background-color: #2a5f9e;
}

QPushButton:pressed {
    background-color: #1a4a7a;
}

QPushButton:disabled {
    background-color: #666;
    color: #aaa;
}

QPushButton.success {
    background-color: #2cc985;
}

QPushButton.success:hover {
    background-color: #259b6e;
}

QPushButton.secondary {
    background-color: #14375e;
}

QPushButton.secondary:hover {
    background-color: #1a4a7a;
}

QPushButton.danger {
    background-color: #dc3545;
}

QPushButton.danger:hover {
    background-color: #c82333;
}

QPushButton.attention {
    background-color: #ff9800;
}

QPushButton.attention:hover {
    background-color: #f57c00;
}

QTabWidget::pane {
    border: 1px solid #444;
    background-color: #2b2b2b;
    border-radius: 4px;
}

QTabBar::tab {
    background-color: #3b3b3b;
    color: white;
    padding: 8px 16px;
    margin: 2px;
    border-radius: 4px;
}

QTabBar::tab:selected {
    background-color: #1f538d;
}

QTabBar::tab:hover {
    background-color: #4a4a4a;
}

QProgressBar {
    border: 1px solid #444;
    border-radius: 4px;
    text-align: center;
    background-color: #2b2b2b;
    color: white;
}

QProgressBar::chunk {
    background-color: #2cc985;
    border-radius: 3px;
}

QLineEdit {
    background-color: #2b2b2b;
    border: 1px solid #444;
    border-radius: 4px;
    padding: 8px;
    color: white;
}

QLineEdit:focus {
    border-color: #1f538d;
}

QTextEdit {
    background-color: #2b2b2b;
    border: 1px solid #444;
    border-radius: 4px;
    color: white;
    selection-background-color: #1f538d;
}

QListWidget {
    background-color: #2b2b2b;
    border: 1px solid #444;
    border-radius: 4px;
    outline: none;
}

QListWidget::item {
    padding: 0px;
    border-bottom: 1px solid #444;
    background-color: transparent;
}

QListWidget::item:selected {
    background-color: transparent;
}

QListWidget::item:hover {
    background-color: #333;
}

QFrame {
    background-color: #2b2b2b;
    border: 1px solid #444;
    border-radius: 6px;
}

QFrame:hover {
    border-color: #1f538d;
}

QFrame.attention {
    border-color: #ff9800;
}

QFrame.attention:hover {
    border-color: #f57c00;
}

QGroupBox {
    background-color: #2b2b2b;
    border: 1px solid #444;
    border-radius: 6px;
    font-weight: bold;
    padding-top: 10px;
}

QGroupBox::title {
    subcontrol-origin: margin;
    left: 10px;
    padding: 0 5px 0 5px;
}

QGroupBox.attention {
    border: 2px solid #ff9800;
    background-color: #2d2416;
}

QGroupBox.attention::title {
    color: #ff9800;
}

QLabel {
    color: white;
}

QCheckBox {
    color: white;
}

QCheckBox::indicator {
    width: 18px;
    height: 18px;
}

QCheckBox::indicator:unchecked {
    background-color: #2b2b2b;
    border: 1px solid #666;
    border-radius: 3px;
}

QCheckBox::indicator:checked {
    background-color: #1f538d;
    border: 1px solid #1f538d;
    border-radius: 3px;
}

QComboBox {
    background-color: #2b2b2b;
    border: 1px solid #444;
    border-radius: 4px;
    padding: 4px;
    color: white;
}

QComboBox:hover {
    border-color: #1f538d;
}

QComboBox::drop-down {
    border: none;
}

QComboBox::down-arrow {
    width: 12px;
    height: 12px;
}

QScrollBar:vertical {
    background-color: #2b2b2b;
    width: 12px;
    border-radius: 6px;
}

QScrollBar::handle:vertical {
    background-color: #666;
    border-radius: 6px;
    min-height: 20px;
}

QScrollBar::handle:vertical:hover {
    background-color: #777;
}
"""

@dataclass
class HistoryEntry:
    """Representa uma entrada no histórico de processamentos"""
    timestamp: datetime
    pdf_file: str
    success: bool
    result_data: dict
    logs: List[str]
    is_batch: bool = False
    batch_info: Dict = None
    has_attention: bool = False  # NOVO
    attention_details: List = None  # NOVO
    
    def __post_init__(self):
        if self.batch_info is None:
            self.batch_info = {}
        if self.attention_details is None:
            self.attention_details = []

class PersistenceManager:
    """Gerencia persistência de configurações e histórico"""
    
    def __init__(self, app_dir=None):
        if app_dir is None:
            if getattr(sys, 'frozen', False):
                self.app_dir = Path(sys.executable).parent
            else:
                self.app_dir = Path(__file__).parent
        else:
            self.app_dir = Path(app_dir)
        
        self.data_dir = self.app_dir / ".data"
        self.data_dir.mkdir(exist_ok=True)
        
        self.config_file = self.data_dir / "config.json"
        self.history_file = self.data_dir / "history.json"
        
        self.session_id = str(uuid.uuid4())
        self.session_start = datetime.now()
    
    def load_config(self):
        try:
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            return {}
        except Exception as e:
            print(f"Erro ao carregar configurações: {e}")
            return {}
    
    def save_config(self, config_data):
        try:
            config_data['last_saved'] = datetime.now().isoformat()
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"Erro ao salvar configurações: {e}")
    
    def load_all_history_entries(self):
        try:
            if not self.history_file.exists():
                return []
            
            with open(self.history_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            all_entries = []
            for session in data.get('sessions', []):
                for entry_data in session.get('entries', []):
                    entry = HistoryEntry(
                        timestamp=datetime.fromisoformat(entry_data['timestamp']),
                        pdf_file=entry_data['pdf_file'],
                        success=entry_data['success'],
                        result_data=entry_data['result_data'],
                        logs=entry_data['logs'],
                        is_batch=entry_data.get('is_batch', False),
                        batch_info=entry_data.get('batch_info', {}),
                        has_attention=entry_data.get('has_attention', False),  # NOVO
                        attention_details=entry_data.get('attention_details', [])  # NOVO
                    )
                    all_entries.append(entry)
            
            return all_entries
        except Exception as e:
            print(f"Erro ao carregar histórico: {e}")
            return []
    
    def save_history_entry(self, entry_data):
        try:
            history_data = {'sessions': []}
            if self.history_file.exists():
                with open(self.history_file, 'r', encoding='utf-8') as f:
                    history_data = json.load(f)
            
            current_session = None
            for session in history_data['sessions']:
                if session['session_id'] == self.session_id:
                    current_session = session
                    break
            
            if current_session is None:
                current_session = {
                    'session_id': self.session_id,
                    'start_time': self.session_start.isoformat(),
                    'entries': []
                }
                history_data['sessions'].append(current_session)
            
            current_session['entries'].append({
                'timestamp': entry_data.timestamp.isoformat(),
                'pdf_file': entry_data.pdf_file,
                'success': entry_data.success,
                'result_data': entry_data.result_data,
                'logs': entry_data.logs[:50],
                'is_batch': getattr(entry_data, 'is_batch', False),
                'batch_info': getattr(entry_data, 'batch_info', {}),
                'has_attention': getattr(entry_data, 'has_attention', False),  # NOVO
                'attention_details': getattr(entry_data, 'attention_details', [])  # NOVO
            })
            
            history_data['sessions'] = history_data['sessions'][-10:]
            
            with open(self.history_file, 'w', encoding='utf-8') as f:
                json.dump(history_data, f, indent=2, ensure_ascii=False)
                
        except Exception as e:
            print(f"Erro ao salvar histórico: {e}")
    
    def clear_history(self):
        try:
            if self.history_file.exists():
                self.history_file.unlink()
        except Exception as e:
            print(f"Erro ao limpar histórico: {e}")

class PDFProcessorThread(QThread):
    """Thread para processamento de PDFs com signals thread-safe"""
    
    progress_updated = pyqtSignal(str, int, str)  # filename, progress, message
    pdf_completed = pyqtSignal(str, dict)         # filename, result_data
    batch_completed = pyqtSignal()
    log_message = pyqtSignal(str)
    
    def __init__(self, pdf_files, processor_factory, trabalho_dir, max_workers=3):
        super().__init__()
        self.pdf_files = pdf_files
        self.processor_factory = processor_factory
        self.trabalho_dir = trabalho_dir
        self.max_workers = max_workers
        self.completed_count = 0
        self.total_count = len(pdf_files)
    
    def run(self):
        """Executa processamento unificado com ThreadPoolExecutor para todos os casos"""
        # Sempre usa processamento em lote (ThreadPoolExecutor)
        # mesmo para 1 arquivo - simplicidade > micro-otimização
        self._process_batch()
        self.batch_completed.emit()
    
    def _process_single_pdf(self, pdf_file):
        """
        Processa um único PDF (usado tanto para casos individuais quanto em lote)
        
        Este método é chamado sempre através do ThreadPoolExecutor, 
        garantindo comportamento consistente independente da quantidade de arquivos.
        """
        filename = Path(pdf_file).name
        
        try:
            processor = self.processor_factory()
            processor.set_trabalho_dir(self.trabalho_dir)
            
            def progress_callback(progress, message=""):
                self.progress_updated.emit(filename, progress, message)
            
            def log_callback(log_message):
                self.log_message.emit(f"[{filename}] {log_message}")
            
            processor.progress_callback = progress_callback
            processor.log_callback = log_callback
            
            self.progress_updated.emit(filename, 0, "Iniciando...")
            
            if Path(pdf_file).parent == Path(self.trabalho_dir):
                pdf_filename = filename
            else:
                pdf_filename = pdf_file
            
            results = processor.process_pdf(pdf_filename)
            
            if results['success']:
                # Determina ícone baseado em atenção
                if results.get('has_attention', False):
                    icon = "⚠️"
                    message = f"⚠️ {results['total_extracted']} períodos processados (ATENÇÃO)"
                else:
                    icon = "✅"
                    message = f"✅ {results['total_extracted']} períodos processados"
                
                self.progress_updated.emit(filename, 100, message)
            else:
                self.progress_updated.emit(filename, 0, f"❌ {results['error']}")
            
            self.pdf_completed.emit(filename, results)
            
        except Exception as e:
            error_result = {'success': False, 'error': str(e)}
            self.progress_updated.emit(filename, 0, f"❌ Erro: {str(e)}")
            self.pdf_completed.emit(filename, error_result)
    
    def _process_batch(self):
        """
        Processa arquivos usando ThreadPoolExecutor (unificado para todos os casos)
        
        Sempre usa ThreadPoolExecutor mesmo para 1 arquivo, garantindo:
        - Comportamento consistente
        - Código mais simples
        - Menos bugs de inconsistência
        - Overhead desprezível (~0.01% do tempo total)
        """
        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            future_to_pdf = {
                executor.submit(self._process_single_pdf, pdf_file): pdf_file 
                for pdf_file in self.pdf_files
            }
            
            for future in as_completed(future_to_pdf):
                pdf_file = future_to_pdf[future]
                try:
                    future.result()
                except Exception as e:
                    filename = Path(pdf_file).name
                    self.progress_updated.emit(filename, 0, f"❌ Exceção: {str(e)}")
                    self.pdf_completed.emit(filename, {'success': False, 'error': str(e)})

class DropZoneWidget(QWidget):
    """Widget para drag & drop de arquivos PDF"""
    
    files_dropped = pyqtSignal(list)
    
    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.setMinimumHeight(60)
        
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        self.label = QLabel("🎯 Arraste PDFs aqui ou clique para selecionar")
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.label.setStyleSheet("""
            QLabel {
                color: #666;
                font-size: 14px;
                padding: 20px;
                border: 2px dashed #666;
                border-radius: 8px;
                background-color: #2b2b2b;
            }
        """)
        
        layout.addWidget(self.label)
        
        # Permite clicar para selecionar
        self.label.mousePressEvent = self._on_click
    
    def _on_click(self, event):
        """Abre diálogo de seleção quando clicado"""
        if hasattr(self.parent(), 'select_pdfs'):
            self.parent().select_pdfs()
    
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            pdf_files = [url.toLocalFile() for url in urls if url.toLocalFile().lower().endswith('.pdf')]
            if pdf_files:
                event.acceptProposedAction()
                self.label.setStyleSheet(self.label.styleSheet() + "border-color: #2cc985;")
    
    def dragLeaveEvent(self, event):
        self.label.setStyleSheet(self.label.styleSheet().replace("border-color: #2cc985;", ""))
    
    def dropEvent(self, event: QDropEvent):
        files = []
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            if file_path.lower().endswith('.pdf'):
                files.append(file_path)
        
        if files:
            self.files_dropped.emit(files)
        
        event.acceptProposedAction()
        self.label.setStyleSheet(self.label.styleSheet().replace("border-color: #2cc985;", ""))

class BatchProgressDialog(QDialog):
    """Dialog para mostrar progresso de processamento em lote"""
    
    def __init__(self, pdf_files, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Processando {len(pdf_files)} PDFs...")
        self.setModal(True)
        
        # Define tamanho fixo e desabilita redimensionamento/maximizar
        self.setFixedSize(850, 580)
        self.setWindowFlags(
            Qt.WindowType.Dialog | 
            Qt.WindowType.WindowCloseButtonHint
        )
        
        self.pdf_widgets = {}
        
        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        
        # Header
        self.header = QLabel(f"🔄 Processando {len(pdf_files)} PDFs em Paralelo")
        self.header.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.header.setStyleSheet("font-size: 16px; font-weight: bold; padding: 15px;")
        layout.addWidget(self.header)
        
        # Status informativo
        status_info = QLabel("Acompanhe o progresso individual de cada arquivo abaixo:")
        status_info.setAlignment(Qt.AlignmentFlag.AlignCenter)
        status_info.setStyleSheet("font-size: 12px; color: #888; padding-bottom: 10px;")
        layout.addWidget(status_info)
        
        # Lista de PDFs (sem progresso geral)
        pdfs_group = QGroupBox("📋 Status de Processamento dos Arquivos PDF")
        pdfs_layout = QVBoxLayout(pdfs_group)
        
        # Scroll area para PDFs
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)
        scroll_layout.setSpacing(4)
        
        # Cria widgets para cada PDF
        for pdf_file in pdf_files:
            filename = Path(pdf_file).name
            pdf_frame = self._create_pdf_frame(filename)
            scroll_layout.addWidget(pdf_frame)
        
        scroll_layout.addStretch()
        scroll.setWidget(scroll_widget)
        pdfs_layout.addWidget(scroll)
        
        layout.addWidget(pdfs_group)
        
        # Botão fechar (desabilitado durante processamento)
        self.close_button = QPushButton("Fechar")
        self.close_button.setEnabled(False)
        self.close_button.clicked.connect(self.accept)
        self.close_button.setFixedHeight(35)
        layout.addWidget(self.close_button)
    
    def _create_pdf_frame(self, filename):
        """Cria frame para um PDF individual"""
        frame = QFrame()
        frame.setFrameStyle(QFrame.Shape.StyledPanel)
        frame.setFixedHeight(50)  # Altura fixa para evitar oscilações
        
        layout = QHBoxLayout(frame)
        layout.setContentsMargins(8, 4, 8, 4)
        layout.setSpacing(8)
        
        # Ícone de status
        icon = QLabel("⏳")
        icon.setStyleSheet("font-size: 16px;")
        icon.setFixedSize(24, 24)
        icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        # Nome do arquivo (truncado se muito longo)
        display_name = filename
        if len(display_name) > 45:
            display_name = display_name[:42] + "..."
        
        name_label = QLabel(f"📄 {display_name}")
        name_label.setStyleSheet("font-weight: bold; font-size: 11px;")
        name_label.setFixedWidth(350)  # Largura fixa
        name_label.setToolTip(f"📄 {filename}")  # Tooltip com nome completo
        
        # Progress bar
        progress = QProgressBar()
        progress.setFixedSize(120, 18)
        progress.setMinimum(0)
        progress.setMaximum(100)
        progress.setStyleSheet("""
            QProgressBar {
                border: 1px solid #444;
                border-radius: 3px;
                text-align: center;
                font-size: 10px;
            }
            QProgressBar::chunk {
                background-color: #2cc985;
                border-radius: 2px;
            }
        """)
        
        # Status com largura fixa para evitar oscilações
        status = QLabel("Aguardando...")
        status.setStyleSheet("color: #888; font-size: 10px;")
        status.setFixedWidth(180)  # Largura fixa para evitar redimensionamento
        status.setWordWrap(False)  # Sem quebra de linha
        status.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        
        layout.addWidget(icon)
        layout.addWidget(name_label)
        layout.addWidget(progress)
        layout.addWidget(status)
        layout.addStretch()  # Preenche espaço restante
        
        self.pdf_widgets[filename] = {
            'icon': icon,
            'progress': progress,
            'status': status
        }
        
        return frame
    
    @pyqtSlot(str, int, str)
    def update_pdf_progress(self, filename, progress, message):
        """Atualiza progresso de um PDF específico"""
        if filename in self.pdf_widgets:
            widgets = self.pdf_widgets[filename]
            
            # Atualiza progresso
            widgets['progress'].setValue(progress)
            widgets['progress'].setToolTip(f"{progress}% concluído")
            
            # Trunca mensagem se muito longa para evitar oscilações
            display_message = message
            if len(display_message) > 25:
                display_message = display_message[:22] + "..."
            
            widgets['status'].setText(display_message)
            widgets['status'].setToolTip(message)  # Tooltip com mensagem completa
            
            # Atualiza ícone baseado no status (incluindo atenção)
            if "✅" in message:
                widgets['icon'].setText("✅")
                widgets['icon'].setStyleSheet("font-size: 16px; color: #2cc985;")
            elif "⚠️" in message:
                widgets['icon'].setText("⚠️")
                widgets['icon'].setStyleSheet("font-size: 16px; color: #ff9800;")
            elif "❌" in message:
                widgets['icon'].setText("❌")
                widgets['icon'].setStyleSheet("font-size: 16px; color: #f44336;")
            elif progress > 0:
                widgets['icon'].setText("🔄")
                widgets['icon'].setStyleSheet("font-size: 16px; color: #1f538d;")
            else:
                widgets['icon'].setText("⏳")
                widgets['icon'].setStyleSheet("font-size: 16px; color: #ffa726;")
    
    @pyqtSlot()
    def handle_batch_completed(self):
        """Habilita botão fechar quando processamento termina"""
        # Atualiza header para mostrar conclusão
        total_pdfs = len(self.pdf_widgets)
        self.header.setText(f"✅ Processamento Concluído - {total_pdfs} PDFs")
        self.header.setStyleSheet("font-size: 16px; font-weight: bold; padding: 15px; color: #2cc985;")
        
        self.close_button.setEnabled(True)
        self.close_button.setText("✅ Fechar")
        self.close_button.setStyleSheet("""
            QPushButton {
                background-color: #2cc985;
                border: none;
                border-radius: 6px;
                padding: 8px 16px;
                color: white;
                font-weight: bold;
                min-height: 20px;
            }
            QPushButton:hover {
                background-color: #259b6e;
            }
        """)

class HistoryItemWidget(QWidget):
    """Widget personalizado para item do histórico"""
    
    details_requested = pyqtSignal(object)  # HistoryEntry
    file_open_requested = pyqtSignal(object)  # HistoryEntry
    
    def __init__(self, entry: HistoryEntry):
        super().__init__()
        self.entry = entry
        
        layout = QHBoxLayout(self)
        layout.setContentsMargins(12, 8, 12, 8)
        layout.setSpacing(12)
        
        # Ícone de status com atenção
        if entry.success:
            if entry.has_attention:
                status_icon = "⚠️"
                status_color = "#ff9800"
            else:
                status_icon = "✅"
                status_color = "#2cc985"
        else:
            status_icon = "❌"
            status_color = "#f44336"
        
        status_label = QLabel(status_icon)
        status_label.setStyleSheet(f"font-size: 18px; color: {status_color};")
        status_label.setFixedSize(24, 24)
        status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        # Informações principais
        info_layout = QVBoxLayout()
        info_layout.setSpacing(3)
        
        # Nome do arquivo (simples, sem indicação de lote)
        if entry.success and entry.result_data.get('arquivo_final'):
            display_name = Path(entry.result_data['arquivo_final']).stem
        else:
            display_name = Path(entry.pdf_file).stem
        
        # Indicador apenas de atenção, se necessário
        if entry.has_attention:
            display_name += " (ATENÇÃO)"
        
        # Truncar nome se muito longo
        if len(display_name) > 60:
            display_name = display_name[:57] + "..."
        
        name_label = QLabel(display_name)
        name_label.setStyleSheet("font-weight: bold; font-size: 13px; color: white;")
        name_label.setToolTip(display_name)
        
        # Resultado e timestamp
        if entry.success:
            result_text = f"✓ {entry.result_data.get('total_extracted', 0)} períodos processados"
            if entry.result_data.get('person_name'):
                person_name = entry.result_data['person_name']
                if len(person_name) > 25:
                    person_name = person_name[:22] + "..."
                result_text += f" • {person_name}"
            
            if entry.has_attention:
                result_text += " • ⚠️ COM OBSERVAÇÕES"
        else:
            error_msg = entry.result_data.get('error', 'Erro desconhecido')
            if len(error_msg) > 40:
                error_msg = error_msg[:37] + "..."
            result_text = f"✗ {error_msg}"
        
        result_text += f" • {entry.timestamp.strftime('%d/%m/%Y %H:%M')}"
        
        result_label = QLabel(result_text)
        result_label.setStyleSheet("font-size: 11px; color: #aaa;")
        result_label.setToolTip(f"Processado em: {entry.timestamp.strftime('%d/%m/%Y %H:%M:%S')}")
        
        info_layout.addWidget(name_label)
        info_layout.addWidget(result_label)
        
        # Botões de ação
        buttons_layout = QHBoxLayout()
        buttons_layout.setSpacing(6)
        
        if entry.success:
            open_btn = QPushButton("📂")
            open_btn.setToolTip("Abrir arquivo Excel")
            open_btn.setStyleSheet("""
                QPushButton {
                    background-color: transparent;
                    border: none;
                }
                QPushButton:hover {
                    background-color: #259b6e;
                }
            """)
            open_btn.clicked.connect(lambda: self.file_open_requested.emit(self.entry))
            buttons_layout.addWidget(open_btn)
        
        details_btn = QPushButton("📝")
        details_btn.setToolTip("Ver logs detalhados")
        details_btn.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                border: none;
            }
            QPushButton:hover {
                background-color: #2a5f9e;
            }
        """)
        details_btn.clicked.connect(lambda: self.details_requested.emit(self.entry))
        buttons_layout.addWidget(details_btn)
        
        layout.addWidget(status_label)
        layout.addLayout(info_layout, 1)
        layout.addLayout(buttons_layout)

class HistoryDetailsDialog(QDialog):
    """Dialog para mostrar detalhes do histórico"""
    
    def __init__(self, entry: HistoryEntry, parent=None):
        super().__init__(parent)
        self.entry = entry
        self.setWindowTitle("📝 Detalhes do Histórico")
        self.setModal(True)
        
        # Define tamanho fixo e desabilita redimensionamento/maximizar  
        self.setFixedSize(750, 650)  # Aumentado para incluir seção de atenção
        self.setWindowFlags(
            Qt.WindowType.Dialog | 
            Qt.WindowType.WindowCloseButtonHint
        )
        
        layout = QVBoxLayout(self)
        
        # Header compacto
        header_frame = QFrame()
        header_layout = QVBoxLayout(header_frame)
        
        # Título (simples, sem indicação de lote)
        title_text = f"📄 {Path(entry.pdf_file).stem}"
        if entry.has_attention:
            title_text += " ⚠️"
        
        title_label = QLabel(title_text)
        title_label.setStyleSheet("font-size: 16px; font-weight: bold;")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        # Informações resumidas
        info_parts = []
        
        if entry.success:
            if entry.has_attention:
                info_parts.append("⚠️ Sucesso com Atenção")
            else:
                info_parts.append("✅ Sucesso")
        else:
            info_parts.append("❌ Erro")
        
        info_parts.append(f"🕒 {entry.timestamp.strftime('%d/%m/%Y %H:%M:%S')}")
        
        if entry.success and entry.result_data:
            if entry.result_data.get('total_extracted') is not None:
                info_parts.append(f"📊 {entry.result_data['total_extracted']} períodos")
            
            if entry.result_data.get('person_name'):
                info_parts.append(f"👤 {entry.result_data['person_name']}")
            
            if entry.result_data.get('arquivo_final'):
                arquivo_nome = Path(entry.result_data['arquivo_final']).name
                info_parts.append(f"💾 {arquivo_nome}")
        
        info_text = " • ".join(info_parts)
        info_label = QLabel(info_text)
        info_label.setStyleSheet("font-size: 11px; color: #888;")
        info_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        info_label.setWordWrap(True)
        
        header_layout.addWidget(title_label)
        header_layout.addWidget(info_label)
        layout.addWidget(header_frame)
        
        # Seção de atenção (NOVA) - prioritária
        if entry.has_attention and entry.attention_details:
            attention_group = QGroupBox("⚠️ PONTOS DE ATENÇÃO")
            attention_group.setProperty("class", "attention")
            attention_layout = QVBoxLayout(attention_group)
            
            # Informação destacada
            attention_info = QLabel(
                "🔍 Durante o processamento foram detectadas situações que requerem atenção:"
            )
            attention_info.setStyleSheet("color: #ff9800; font-weight: bold; font-size: 12px;")
            attention_info.setWordWrap(True)
            attention_layout.addWidget(attention_info)
            
            # Lista os pontos de atenção com informações estruturadas
            for i, detail in enumerate(entry.attention_details, 1):
                attention_item = QFrame()
                attention_item.setStyleSheet("""
                    QFrame {
                        background-color: #3d2d1a;
                        border: 1px solid #ff9800;
                        border-radius: 4px;
                        padding: 8px;
                        margin: 2px;
                    }
                """)
                
                attention_item_layout = QVBoxLayout(attention_item)
                attention_item_layout.setContentsMargins(8, 4, 8, 4)
                
                # Título do ponto de atenção
                titulo = QLabel(f"📋 Ponto {i}: {detail.get('periodo', 'N/A')} ({detail.get('folha_type', 'N/A')})")
                titulo.setStyleSheet("color: #ff9800; font-weight: bold; font-size: 11px;")
                attention_item_layout.addWidget(titulo)
                
                # Processa cada detalhe de atenção
                detalhes_list = detail.get('detalhes', [])
                
                # Se não há detalhes estruturados, mostra informação para dados antigos
                if not detalhes_list:
                    info_label = QLabel("💡 Dados processados com versão anterior - detalhes não disponíveis")
                    info_label.setStyleSheet("color: #ffcc80; font-size: 10px; margin-left: 15px; font-style: italic;")
                    attention_item_layout.addWidget(info_label)
                    
                    help_label = QLabel("ℹ️ Processe novamente o PDF para ver informações detalhadas")
                    help_label.setStyleSheet("color: #888; font-size: 9px; margin-left: 15px; font-style: italic;")
                    attention_item_layout.addWidget(help_label)
                    attention_layout.addWidget(attention_item)
                    continue
                
                for detalhe_raw in detalhes_list:
                    if isinstance(detalhe_raw, dict):
                        # Informação estruturada
                        tipo = detalhe_raw.get('tipo', 'desconhecido')
                        
                        if tipo == 'soma_automatica':
                            # Soma automática (PREMIO PROD, HE 100%, etc.)
                            descricao = detalhe_raw.get('descricao', 'CÓDIGOS ESPECÍFICOS')
                            codigos = detalhe_raw.get('codigos', [])
                            valor_somado = detalhe_raw.get('valor_somado', 0)
                            valores_individuais = detalhe_raw.get('valores_individuais', {})
                            
                            detalhe_label = QLabel(f"💡 SOMA AUTOMÁTICA - {descricao}")
                            detalhe_label.setStyleSheet("color: #ffcc80; font-weight: bold; font-size: 10px; margin-left: 15px;")
                            attention_item_layout.addWidget(detalhe_label)
                            
                            # Mostra códigos e valor final
                            codigos_str = ' + '.join(codigos)
                            codigos_label = QLabel(f"📋 {codigos_str} = {valor_somado}")
                            codigos_label.setStyleSheet("color: #ffcc80; font-size: 10px; margin-left: 25px;")
                            attention_item_layout.addWidget(codigos_label)
                            
                            # Mostra valores individuais
                            for codigo, valor in valores_individuais.items():
                                valor_label = QLabel(f"   • {codigo}: {valor}")
                                valor_label.setStyleSheet("color: #ffcc80; font-size: 9px; margin-left: 35px;")
                                attention_item_layout.addWidget(valor_label)
                        
                        elif tipo == 'duplicidade_descricao':
                            # Duplicidade por descrição (descoberta automática)
                            descricao = detalhe_raw.get('descricao', 'DESCRIÇÃO DESCONHECIDA')
                            codigos = detalhe_raw.get('codigos', [])
                            valores_individuais = detalhe_raw.get('valores_individuais', {})
                            colunas = detalhe_raw.get('colunas_afetadas', [])
                            
                            detalhe_label = QLabel(f"🔍 DUPLICIDADE DETECTADA - {descricao}")
                            detalhe_label.setStyleSheet("color: #ffcc80; font-weight: bold; font-size: 10px; margin-left: 15px;")
                            attention_item_layout.addWidget(detalhe_label)
                            
                            # Mostra códigos
                            codigos_str = ' + '.join(codigos)
                            codigos_label = QLabel(f"📋 {codigos_str} (verificação manual recomendada)")
                            codigos_label.setStyleSheet("color: #ffcc80; font-size: 10px; margin-left: 25px;")
                            attention_item_layout.addWidget(codigos_label)
                            
                            # Mostra colunas afetadas
                            if colunas:
                                colunas_label = QLabel(f"📊 Colunas: {', '.join(colunas)}")
                                colunas_label.setStyleSheet("color: #ffcc80; font-size: 9px; margin-left: 25px;")
                                attention_item_layout.addWidget(colunas_label)
                            
                            # Mostra valores preservados individualmente  
                            for codigo, valor in valores_individuais.items():
                                valor_label = QLabel(f"   • {codigo}: {valor} (preservado)")
                                valor_label.setStyleSheet("color: #ffcc80; font-size: 9px; margin-left: 35px;")
                                attention_item_layout.addWidget(valor_label)
                        
                        else:
                            # Formato não reconhecido - mostra detalhes como texto
                            detalhes_text = detalhe_raw.get('detalhes', str(detalhe_raw))
                            detalhe_label = QLabel(f"💡 {detalhes_text}")
                            detalhe_label.setStyleSheet("color: #ffcc80; font-size: 10px; margin-left: 15px;")
                            detalhe_label.setWordWrap(True)
                            attention_item_layout.addWidget(detalhe_label)
                    
                    else:
                        # String simples - compatibilidade com versões anteriores
                        detalhe_label = QLabel(f"💡 {detalhe_raw}")
                        detalhe_label.setStyleSheet("color: #ffcc80; font-size: 10px; margin-left: 15px;")
                        detalhe_label.setWordWrap(True)
                        attention_item_layout.addWidget(detalhe_label)
                
                attention_layout.addWidget(attention_item)
            
            # Explicação
            explanation = QLabel(
                "ℹ️ Estes pontos de atenção não impedem o funcionamento da planilha, "
                "mas indicam situações especiais que foram tratadas automaticamente."
            )
            explanation.setStyleSheet("color: #888; font-size: 10px; font-style: italic;")
            explanation.setWordWrap(True)
            attention_layout.addWidget(explanation)
            
            layout.addWidget(attention_group)
        
        # Área de logs (ajustada)
        logs_group = QGroupBox("📄 Logs Detalhados")
        logs_layout = QVBoxLayout(logs_group)
        
        self.logs_text = QTextEdit()
        self.logs_text.setReadOnly(True)
        
        # Font monospace
        font = QFont("Consolas", 10)
        font.setStyleHint(QFont.StyleHint.Monospace)
        self.logs_text.setFont(font)
        
        self._populate_logs()
        logs_layout.addWidget(self.logs_text)
        layout.addWidget(logs_group, 1)  # Máximo stretch
        
        # Botões
        buttons_layout = QHBoxLayout()
        
        if entry.success and entry.result_data.get('excel_path'):
            open_btn = QPushButton("📂 Abrir Arquivo")
            open_btn.clicked.connect(self._open_file)
            buttons_layout.addWidget(open_btn)
        
        buttons_layout.addStretch()
        
        close_btn = QPushButton("✖️ Fechar")
        close_btn.clicked.connect(self.accept)
        buttons_layout.addWidget(close_btn)
        
        layout.addLayout(buttons_layout)
    
    def _populate_logs(self):
        """Popula área de logs com informações resumidas"""
        header_info = [
            f"📄 Arquivo: {self.entry.pdf_file}",
            f"🕒 Processado em: {self.entry.timestamp.strftime('%d/%m/%Y %H:%M:%S')}",
        ]
        
        if self.entry.success and self.entry.result_data:
            if self.entry.result_data.get('person_name'):
                header_info.append(f"👤 Pessoa detectada: {self.entry.result_data['person_name']}")
            
            if self.entry.result_data.get('total_extracted') is not None:
                total = self.entry.result_data['total_extracted']
                folha_normal = self.entry.result_data.get('folha_normal_periods', 0)
                salario_13 = self.entry.result_data.get('salario_13_periods', 0)
                header_info.append(f"📊 Resultado: {total} períodos processados (FOLHA NORMAL: {folha_normal}, 13º SALÁRIO: {salario_13})")
            
            # Indica se há atenção
            if self.entry.has_attention:
                header_info.append(f"⚠️ ATENÇÃO: Processamento com observações especiais")
            
            if self.entry.result_data.get('arquivo_final'):
                header_info.append(f"💾 Arquivo final: {self.entry.result_data['arquivo_final']}")
                
            if self.entry.result_data.get('total_pages'):
                header_info.append(f"📑 Total de páginas analisadas: {self.entry.result_data['total_pages']}")
        else:
            if self.entry.result_data and self.entry.result_data.get('error'):
                header_info.append(f"❌ Erro: {self.entry.result_data['error']}")
        
        # Informações sobre processamento (sem logs extensos)
        header_info.append("")
        header_info.append("📋 RESUMO DO PROCESSAMENTO:")
        header_info.append(f"Status: {'✅ Sucesso' if self.entry.success else '❌ Falha'}")
        
        if self.entry.has_attention:
            header_info.append("⚠️ Observações: Situações especiais detectadas e tratadas automaticamente")
        
        # Apenas informações resumidas, sem logs extensos
        content = "\n".join(header_info)
        self.logs_text.setPlainText(content)
    
    def _open_file(self):
        """Abre arquivo Excel associado"""
        try:
            file_path = self.entry.result_data.get('excel_path')
            if not file_path:
                arquivo_rel = self.entry.result_data.get('arquivo_final')
                if arquivo_rel:
                    # Tenta construir caminho relativo
                    file_path = Path.cwd() / arquivo_rel
            
            if not file_path:
                QMessageBox.warning(self, "Arquivo Não Encontrado", "Caminho do arquivo não está disponível.")
                return
            
            file_path = Path(file_path)
            if not file_path.exists():
                QMessageBox.warning(self, "Arquivo Não Encontrado", f"O arquivo não foi encontrado:\n\n{file_path}")
                return
            
            if sys.platform.startswith('win'):
                os.startfile(file_path)
            elif sys.platform == 'darwin':
                subprocess.Popen(['open', str(file_path)])
            else:
                subprocess.Popen(['xdg-open', str(file_path)])
        
        except Exception as e:
            QMessageBox.critical(self, "Erro ao Abrir Arquivo", f"Não foi possível abrir o arquivo:\n\n{e}")

class MainWindow(QMainWindow):
    """Janela principal da aplicação"""
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Processamento de Folha de Pagamento v4.0.1 - PyQt6")
        
        # Define tamanho fixo e desabilita redimensionamento/maximizar
        self.setFixedSize(950, 600)
        self.setWindowFlags(
            Qt.WindowType.Window | 
            Qt.WindowType.WindowCloseButtonHint | 
            Qt.WindowType.WindowMinimizeButtonHint
        )
        
        # Estado da aplicação - INICIALIZAR PRIMEIRO
        self.selected_files = []
        self.trabalho_dir = None
        self.processing = False
        self.current_logs = []  # Inicializar antes de qualquer callback
        self.processing_history = []
        
        # Processamento
        self.processor_thread = None
        self.progress_dialog = None
        
        # Configurações
        self.max_threads = 3
        self.verbose_mode = False
        self.preferred_sheet = ""
        
        # Gerenciador de persistência
        self.persistence = PersistenceManager()
        
        # Timer para salvar configurações - CRIAR ANTES DA INTERFACE
        self.save_timer = QTimer()
        self.save_timer.setSingleShot(True)
        self.save_timer.timeout.connect(self.save_current_config)
        
        # Cria interface (depois de inicializar todas as variáveis)
        self.create_interface()
        
        # Timer para carregar dados persistidos após interface pronta
        QTimer.singleShot(100, self.load_persisted_data)
    
    def create_interface(self):
        """Cria interface principal"""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        layout = QVBoxLayout(central_widget)
        layout.setSpacing(10)
        
        # Header
        self.create_header(layout)
        
        # Tabs
        self.tab_widget = QTabWidget()
        layout.addWidget(self.tab_widget)
        
        # Cria abas
        self.create_processing_tab()
        self.create_history_tab()
        self.create_settings_tab()
        
        # Status bar
        self.statusBar().showMessage("Sistema iniciado - v4.0.1 PyQt6")
    
    def create_header(self, layout):
        """Cria header da aplicação"""
        header_frame = QFrame()
        header_layout = QVBoxLayout(header_frame)
        
        title = QLabel("📄 Processamento de Folha de Pagamento")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("font-size: 22px; font-weight: bold; padding: 10px;")
        
        subtitle = QLabel("Automatização de folhas de pagamento PDF para Excel v4.0.1 - PyQt6 Performance")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        subtitle.setStyleSheet("font-size: 11px; color: #888; padding-bottom: 10px;")
        
        header_layout.addWidget(title)
        header_layout.addWidget(subtitle)
        
        layout.addWidget(header_frame)
    
    def create_processing_tab(self):
        """Cria aba de processamento"""
        processing_widget = QWidget()
        layout = QVBoxLayout(processing_widget)
        
        # Configuração do diretório
        config_group = QGroupBox("📁 Configuração do Diretório de Trabalho")
        config_layout = QVBoxLayout(config_group)
        
        help_label = QLabel("Selecione a pasta que contém o arquivo MODELO.xlsm:")
        help_label.setStyleSheet("color: #888;")
        config_layout.addWidget(help_label)
        
        dir_layout = QHBoxLayout()
        self.dir_entry = QLineEdit()
        self.dir_entry.setPlaceholderText("Caminho para o diretório de trabalho...")
        self.dir_entry.textChanged.connect(self._on_dir_changed)
        
        self.dir_button = QPushButton("📂 Selecionar")
        self.dir_button.clicked.connect(self.select_directory)
        
        self.config_status = QLabel("⚠️ Configuração necessária")
        self.config_status.setStyleSheet("color: #ffa726; font-weight: bold;")
        
        dir_layout.addWidget(self.dir_entry, 1)
        dir_layout.addWidget(self.dir_button)
        dir_layout.addWidget(self.config_status)
        
        config_layout.addLayout(dir_layout)
        layout.addWidget(config_group)
        
        # Seleção de arquivos
        files_group = QGroupBox("📎 Seleção de Arquivos PDF")
        files_layout = QVBoxLayout(files_group)
        
        # Header com contador
        header_layout = QHBoxLayout()
        files_title = QLabel("Arquivos selecionados:")
        self.file_counter_label = QLabel("0 arquivos")
        self.file_counter_label.setStyleSheet("color: #888;")
        
        header_layout.addWidget(files_title)
        header_layout.addStretch()
        header_layout.addWidget(self.file_counter_label)
        files_layout.addLayout(header_layout)
        
        # Botões de ação
        actions_layout = QHBoxLayout()
        
        select_btn = QPushButton("📂 Selecionar PDFs")
        select_btn.clicked.connect(self.select_pdfs)
        
        clear_btn = QPushButton("🗑️ Limpar")
        clear_btn.setProperty("class", "secondary")
        clear_btn.clicked.connect(self.clear_selection)
        
        self.process_btn = QPushButton("🚀 Processar PDFs")
        self.process_btn.setProperty("class", "success")
        self.process_btn.clicked.connect(self.process_pdfs)
        self.process_btn.setEnabled(False)
        
        actions_layout.addWidget(select_btn)
        actions_layout.addWidget(clear_btn)
        actions_layout.addStretch()
        actions_layout.addWidget(self.process_btn)
        
        files_layout.addLayout(actions_layout)
        
        # Drop zone
        self.drop_zone = DropZoneWidget()
        self.drop_zone.files_dropped.connect(self.handle_dropped_files)
        files_layout.addWidget(self.drop_zone)
        
        # Lista de arquivos
        self.files_list = QListWidget()
        self.files_list.setMaximumHeight(200)
        self.files_list.setStyleSheet("""
            QListWidget {
                background-color: #2b2b2b;
                border: 1px solid #444;
                border-radius: 4px;
            }
            QListWidget::item {
                border-bottom: 1px solid #444;
                padding: 2px;
            }
            QListWidget::item:hover {
                background-color: #333;
            }
        """)
        files_layout.addWidget(self.files_list)
        
        layout.addWidget(files_group)
        layout.addStretch()
        
        self.tab_widget.addTab(processing_widget, "📄 Processamento")
    
    def create_history_tab(self):
        """Cria aba de histórico"""
        history_widget = QWidget()
        layout = QVBoxLayout(history_widget)
        
        # Header com controles
        header_layout = QHBoxLayout()
        
        title = QLabel("📊 Histórico de Processamentos")
        title.setStyleSheet("font-size: 16px; font-weight: bold;")
        
        self.history_status_label = QLabel("Nenhum PDF no histórico")
        self.history_status_label.setStyleSheet("color: #888;")
        
        clear_history_btn = QPushButton("🗑️ Limpar Histórico")
        clear_history_btn.setProperty("class", "secondary")
        clear_history_btn.clicked.connect(self.clear_history)
        
        header_layout.addWidget(title)
        header_layout.addWidget(self.history_status_label, 1)
        header_layout.addWidget(clear_history_btn)
        
        layout.addLayout(header_layout)
        
        # Lista de histórico (virtualizada automaticamente pelo QListWidget)
        self.history_list = QListWidget()
        layout.addWidget(self.history_list)
        
        self.tab_widget.addTab(history_widget, "📊 Histórico")
    
    def create_settings_tab(self):
        """Cria aba de configurações"""
        settings_widget = QWidget()
        layout = QVBoxLayout(settings_widget)
        
        # Processamento paralelo
        parallel_group = QGroupBox("🚀 Processamento (Sempre Paralelo)")
        parallel_layout = QFormLayout(parallel_group)
        
        parallel_desc = QLabel("Configure quantos PDFs podem ser processados simultaneamente. O sistema sempre usa ThreadPoolExecutor para garantir comportamento consistente.")
        parallel_desc.setStyleSheet("color: #888;")
        parallel_desc.setWordWrap(True)
        parallel_layout.addRow(parallel_desc)
        
        self.threads_combo = QComboBox()
        self.threads_combo.addItems(["1", "2", "3", "4", "5", "6"])
        self.threads_combo.setCurrentText(str(self.max_threads))
        self.threads_combo.currentTextChanged.connect(self._on_threads_changed)
        
        parallel_layout.addRow("PDFs simultâneos:", self.threads_combo)
        layout.addWidget(parallel_group)
        
        # Planilha personalizada
        sheet_group = QGroupBox("📊 Nome da Planilha")
        sheet_layout = QFormLayout(sheet_group)
        
        sheet_desc = QLabel("Especifique o nome da planilha a ser atualizada. Deixe vazio para usar 'LEVANTAMENTO DADOS' (padrão).")
        sheet_desc.setStyleSheet("color: #888;")
        sheet_desc.setWordWrap(True)
        sheet_layout.addRow(sheet_desc)
        
        self.sheet_entry = QLineEdit()
        self.sheet_entry.setPlaceholderText("LEVANTAMENTO DADOS")
        self.sheet_entry.textChanged.connect(self._on_config_changed)
        
        sheet_layout.addRow("Nome da planilha:", self.sheet_entry)
        layout.addWidget(sheet_group)
        
        # Modo verboso
        verbose_group = QGroupBox("🔍 Logs Detalhados")
        verbose_layout = QVBoxLayout(verbose_group)
        
        self.verbose_checkbox = QCheckBox("Habilitar modo verboso (logs detalhados)")
        self.verbose_checkbox.toggled.connect(self._on_config_changed)
        
        verbose_layout.addWidget(self.verbose_checkbox)
        layout.addWidget(verbose_group)
        
        # Informações sobre funcionalidades (SEÇÃO REORGANIZADA)
        features_group = QGroupBox("ℹ️ Funcionalidades Automáticas")
        features_layout = QVBoxLayout(features_group)
        
        features_desc = QLabel(
            "O sistema detecta e trata automaticamente situações especiais durante o processamento:\n\n"
            "• 📋 DUPLICIDADE DE CÓDIGOS: Quando os códigos 01003601 e 01003602 (PREMIO PROD. MENSAL) "
            "aparecem no mesmo mês, os valores são somados automaticamente.\n\n"
            "• ⚠️ MARCAÇÃO AUTOMÁTICA: Processamentos com situações especiais são marcados para referência.\n\n"
            "• 🧵 PROCESSAMENTO UNIFICADO: Sempre usa ThreadPoolExecutor para comportamento consistente.\n\n"
            "• 📝 DETALHES COMPLETOS: Informações detalhadas sobre o processamento estão sempre disponíveis."
        )
        features_desc.setStyleSheet("color: #888; font-size: 11px;")
        features_desc.setWordWrap(True)
        features_layout.addWidget(features_desc)
        
        layout.addWidget(features_group)
        
        layout.addStretch()
        
        self.tab_widget.addTab(settings_widget, "⚙️ Configurações")
    
    def _get_processor(self):
        """Factory para criar processadores individuais"""
        def create_processor():
            processor = PDFProcessorCore()
            if self.preferred_sheet:
                processor.preferred_sheet = self.preferred_sheet
            return processor
        return create_processor
    
    def select_directory(self):
        """Seleciona diretório de trabalho"""
        directory = QFileDialog.getExistingDirectory(
            self,
            "Selecione o diretório de trabalho (que contém MODELO.xlsm)",
            self.trabalho_dir or ""
        )
        
        if directory:
            self.dir_entry.setText(directory)
    
    def _on_dir_changed(self):
        """Callback quando diretório muda"""
        QTimer.singleShot(500, self.validate_config)
    
    def validate_config(self):
        """Valida configuração do diretório"""
        directory = self.dir_entry.text().strip()
        
        if not directory:
            self.config_status.setText("⚠️ Selecione um diretório")
            self.config_status.setStyleSheet("color: #ffa726; font-weight: bold;")
            return
        
        try:
            processor_factory = self._get_processor()
            processor = processor_factory()
            processor.set_trabalho_dir(directory)
            valid, message = processor.validate_trabalho_dir()
            
            if valid:
                self.config_status.setText("✅ Configuração válida")
                self.config_status.setStyleSheet("color: #2cc985; font-weight: bold;")
                self.trabalho_dir = directory
                self._update_process_button()
                self._on_config_changed()
            else:
                self.config_status.setText(f"❌ {message}")
                self.config_status.setStyleSheet("color: #f44336; font-weight: bold;")
        except Exception as e:
            self.config_status.setText(f"❌ Erro: {str(e)}")
            self.config_status.setStyleSheet("color: #f44336; font-weight: bold;")
    
    def select_pdfs(self):
        """Seleciona arquivos PDF"""
        if not self.trabalho_dir:
            QMessageBox.warning(self, "Configuração Necessária", "Configure o diretório de trabalho primeiro.")
            return
        
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Selecione arquivos PDF (um ou múltiplos)",
            self.trabalho_dir,
            "Arquivos PDF (*.pdf)"
        )
        
        if files:
            self.selected_files = files
            self.update_selected_files_display()
    
    def clear_selection(self):
        """Limpa seleção de arquivos"""
        self.selected_files = []
        self.update_selected_files_display()
    
    def handle_dropped_files(self, files):
        """Manipula arquivos arrastados"""
        self.selected_files = files
        self.update_selected_files_display()
    
    def update_selected_files_display(self):
        """Atualiza display dos arquivos selecionados"""
        count = len(self.selected_files)
        self.file_counter_label.setText(f"{count} arquivo{'s' if count != 1 else ''}")
        
        if count > 0:
            self.file_counter_label.setStyleSheet("color: #2cc985; font-weight: bold;")
        else:
            self.file_counter_label.setStyleSheet("color: #888;")
        
        # Atualiza lista com botões de remoção
        self.files_list.clear()
        for i, file_path in enumerate(self.selected_files):
            filename = Path(file_path).name
            
            # Cria widget personalizado para cada item
            item_widget = QWidget()
            item_layout = QHBoxLayout(item_widget)
            item_layout.setContentsMargins(5, 2, 5, 2)
            
            # Label com o arquivo
            file_label = QLabel(f"{i+1}. {filename}")
            file_label.setStyleSheet("color: white; font-size: 11px;")
            file_label.setToolTip(file_path)
            
            # Botão de remoção
            remove_btn = QPushButton("🗑️")
            item_layout.setContentsMargins(0, 0, 0,0)
            remove_btn.setStyleSheet("""
                QPushButton {
                    background-color: transparent;
                    border: none;
                }
                QPushButton:hover {
                    background-color: #c82333;
                }
            """)
            remove_btn.setToolTip("Remover arquivo")
            remove_btn.clicked.connect(lambda checked, idx=i: self.remove_file_at_index(idx))
            
            item_layout.addWidget(file_label, 1)
            item_layout.addWidget(remove_btn)
            
            # Adiciona item à lista
            list_item = QListWidgetItem()
            list_item.setSizeHint(QSize(0, 25))
            self.files_list.addItem(list_item)
            self.files_list.setItemWidget(list_item, item_widget)
        
        self._update_process_button()
        
        # Log da seleção
        if self.selected_files:
            filenames = [Path(f).name for f in self.selected_files]
            self.add_log_message(f"Arquivos selecionados: {', '.join(filenames)}")
    
    def remove_file_at_index(self, index):
        """Remove arquivo no índice especificado"""
        if 0 <= index < len(self.selected_files):
            removed_file = self.selected_files.pop(index)
            filename = Path(removed_file).name
            self.add_log_message(f"Arquivo removido: {filename}")
            self.update_selected_files_display()
    
    def _update_process_button(self):
        """Atualiza estado do botão processar"""
        can_process = (self.trabalho_dir is not None and 
                      len(self.selected_files) > 0 and 
                      not self.processing)
        self.process_btn.setEnabled(can_process)
    
    def process_pdfs(self):
        """Inicia processamento dos PDFs"""
        if self.processing:
            return
        
        if not self.trabalho_dir:
            QMessageBox.critical(self, "Configuração Incompleta", "Configure o diretório de trabalho primeiro.")
            return
        
        if not self.selected_files:
            QMessageBox.critical(self, "Nenhum Arquivo Selecionado", "Selecione pelo menos um arquivo PDF para processar.")
            return
        
        self.processing = True
        self.process_btn.setText("🔄 Processando...")
        self.process_btn.setEnabled(False)
        
        # Cria thread de processamento
        self.processor_thread = PDFProcessorThread(
            self.selected_files,
            self._get_processor(),
            self.trabalho_dir,
            self.max_threads
        )
        
        # Conecta signals
        self.processor_thread.progress_updated.connect(self.handle_progress_update)
        self.processor_thread.pdf_completed.connect(self.handle_pdf_completed)
        self.processor_thread.batch_completed.connect(self.handle_batch_completed)
        self.processor_thread.log_message.connect(self.add_log_message)
        
        # Mostra dialog de progresso para todos os casos
        self.progress_dialog = BatchProgressDialog(self.selected_files, self)
        self.processor_thread.progress_updated.connect(self.progress_dialog.update_pdf_progress)
        self.processor_thread.batch_completed.connect(self.progress_dialog.handle_batch_completed)
        self.progress_dialog.show()
        
        # Inicia processamento
        self.processor_thread.start()
    
    @pyqtSlot(str, int, str)
    def handle_progress_update(self, filename, progress, message):
        """Manipula atualizações de progresso"""
        self.statusBar().showMessage(f"{filename}: {message}")
    
    @pyqtSlot(str, dict)
    def handle_pdf_completed(self, filename, result_data):
        """Manipula conclusão de PDF individual"""
        # Determina se há atenção
        has_attention = result_data.get('has_attention', False)
        attention_details = []
        
        # Converte attention_periods para o formato correto
        if result_data.get('attention_periods'):
            for attention_period in result_data['attention_periods']:
                attention_details.append({
                    'periodo': attention_period.get('periodo', 'N/A'),
                    'folha_type': attention_period.get('folha_type', 'N/A'),
                    'detalhes': attention_period.get('detalhes', [])
                })
        
        # Adiciona ao histórico
        entry = HistoryEntry(
            timestamp=datetime.now(),
            pdf_file=filename,
            success=result_data['success'],
            result_data=result_data,
            logs=self.current_logs.copy(),
            is_batch=len(self.selected_files) > 1,
            batch_info={'batch_size': len(self.selected_files), 'processed_in_batch': True} if len(self.selected_files) > 1 else {},
            has_attention=has_attention,  
            attention_details=attention_details  # Estrutura corrigida
        )
        
        self.processing_history.append(entry)
        self.persistence.save_history_entry(entry)
    
    @pyqtSlot()
    def handle_batch_completed(self):
        """Manipula conclusão do processamento"""
        self.processing = False
        self.process_btn.setText("🚀 Processar PDFs")
        self.process_btn.setEnabled(True)
        
        # Fecha dialog de progresso se existir
        if self.progress_dialog:
            # Dialog será fechado automaticamente quando usuário clicar no botão
            pass
        
        # Atualiza histórico
        self.update_history_display()
        
        # Estatísticas finais
        recent_entries = self.processing_history[-len(self.selected_files):]
        successful = sum(1 for entry in recent_entries if entry.success)
        with_attention = sum(1 for entry in recent_entries if entry.has_attention)
        total = len(self.selected_files)
        
        # Mostra resultado (unificado para qualquer quantidade)
        if successful == total:
            if with_attention > 0:
                QMessageBox.information(
                    self,
                    "✅ Processamento Concluído com Atenção",
                    f"{'Todos os' if total > 1 else 'O'} {total} PDF{'s foram' if total > 1 else ' foi'} processado{'s' if total > 1 else ''} com sucesso!\n\n"
                    f"⚠️ {with_attention} arquivo{'s' if with_attention != 1 else ''} possui{'em' if with_attention != 1 else ''} observações especiais.\n\n"
                    f"📊 Verifique o histórico para mais detalhes.\n"
                    f"📂 {'Os arquivos foram salvos' if total > 1 else 'O arquivo foi salvo'} na pasta DADOS/"
                )
            else:
                QMessageBox.information(
                    self,
                    "✅ Processamento Concluído",
                    f"{'Todos os' if total > 1 else 'O'} {total} PDF{'s foram' if total > 1 else ' foi'} processado{'s' if total > 1 else ''} com sucesso!\n\n"
                    f"📊 Verifique o histórico para mais detalhes.\n"
                    f"📂 {'Os arquivos foram salvos' if total > 1 else 'O arquivo foi salvo'} na pasta DADOS/"
                )
        elif successful > 0:
            attention_text = f"\n⚠️ {with_attention} com observações especiais." if with_attention > 0 else ""
            QMessageBox.warning(
                self,
                "⚠️ Processamento Parcial",
                f"{successful} de {total} PDFs foram processados com sucesso.{attention_text}\n\n"
                f"📊 Verifique o histórico para detalhes dos arquivos que falharam.\n"
                f"📂 Os arquivos processados foram salvos na pasta DADOS/"
            )
        else:
            QMessageBox.critical(
                self,
                "❌ Processamento Falhou",
                f"Nenhum PDF foi processado com sucesso.\n\n"
                f"📊 Verifique o histórico para detalhes dos erros.\n"
                f"🔧 Certifique-se de que os PDFs estão no formato correto."
            )
        
        # Limpa seleção e vai para histórico
        self.clear_selection()
        self.tab_widget.setCurrentIndex(1)  # Aba histórico
        
        self.statusBar().showMessage("Processamento concluído")
    
    def update_history_display(self):
        """Atualiza exibição do histórico"""
        self.history_list.clear()
        
        # Adiciona entradas (mais recentes primeiro)
        for entry in reversed(self.processing_history):
            item = QListWidgetItem()
            item_widget = HistoryItemWidget(entry)
            
            # Conecta signals
            item_widget.details_requested.connect(self.show_history_details)
            item_widget.file_open_requested.connect(self.open_data_file)
            
            item.setSizeHint(QSize(0, 55))  # Altura compacta
            self.history_list.addItem(item)
            self.history_list.setItemWidget(item, item_widget)
        
        # Atualiza status (simplificado, sem distinção de lote)
        total = len(self.processing_history)
        if total > 0:
            success_count = sum(1 for h in self.processing_history if h.success)
            attention_count = sum(1 for h in self.processing_history if h.has_attention)
            
            status_text = f"{total} PDFs no histórico ({success_count} sucessos, {total - success_count} falhas)"
            if attention_count > 0:
                status_text += f" • ⚠️ {attention_count} com atenção"
            
            self.history_status_label.setText(status_text)
            
            # Cor baseada na presença de atenções
            if attention_count > 0:
                self.history_status_label.setStyleSheet("color: #ff9800; font-weight: bold;")
            else:
                self.history_status_label.setStyleSheet("color: #2cc985; font-weight: bold;")
        else:
            self.history_status_label.setText("Nenhum PDF no histórico")
            self.history_status_label.setStyleSheet("color: #888;")
    
    def show_history_details(self, entry: HistoryEntry):
        """Mostra detalhes de entrada do histórico"""
        dialog = HistoryDetailsDialog(entry, self)
        dialog.exec()
    
    def open_data_file(self, entry: HistoryEntry):
        """Abre arquivo Excel de entrada do histórico"""
        try:
            file_path = entry.result_data.get('excel_path')
            if not file_path and self.trabalho_dir:
                arquivo_rel = entry.result_data.get('arquivo_final')
                if arquivo_rel:
                    file_path = Path(self.trabalho_dir) / arquivo_rel
            
            if not file_path:
                QMessageBox.warning(self, "Arquivo Não Encontrado", "Caminho do arquivo não está disponível.")
                return
            
            file_path = Path(file_path)
            if not file_path.exists():
                QMessageBox.warning(self, "Arquivo Não Encontrado", f"O arquivo não foi encontrado:\n\n{file_path}")
                return
            
            if sys.platform.startswith('win'):
                os.startfile(file_path)
            elif sys.platform == 'darwin':
                subprocess.Popen(['open', str(file_path)])
            else:
                subprocess.Popen(['xdg-open', str(file_path)])
        
        except Exception as e:
            QMessageBox.critical(self, "Erro ao Abrir Arquivo", f"Não foi possível abrir o arquivo:\n\n{e}")
    
    def clear_history(self):
        """Limpa histórico"""
        if not self.processing_history:
            return
        
        reply = QMessageBox.question(
            self, 
            "Confirmar Limpeza", 
            "Deseja limpar todo o histórico de processamentos?\n\nEsta ação não pode ser desfeita.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No  # Botão padrão
        )
        
        # Personaliza texto dos botões
        button_yes = reply == QMessageBox.StandardButton.Yes
        
        if button_yes:
            self.processing_history.clear()
            self.persistence.clear_history()
            self.update_history_display()
            
            # Mostra confirmação de sucesso
            QMessageBox.information(
                self,
                "Histórico Limpo",
                "✅ O histórico foi limpo com sucesso!"
            )
    
    def _on_threads_changed(self, value):
        """Callback quando número de threads muda"""
        self.max_threads = int(value)
        self._on_config_changed()
        self.add_log_message(f"Threads configuradas: {value} workers (sempre ThreadPoolExecutor)")
    
    def _on_config_changed(self):
        """Callback genérico para mudanças de configuração"""
        # Verifica se save_timer existe antes de usar
        if hasattr(self, 'save_timer'):
            self.save_timer.start(1000)  # Salva após 1 segundo de inatividade
    
    def save_current_config(self):
        """Salva configuração atual"""
        config = {
            'trabalho_dir': self.trabalho_dir,
            'max_threads': self.max_threads,
            'verbose_mode': self.verbose_checkbox.isChecked(),
            'preferred_sheet': self.sheet_entry.text().strip()
        }
        
        self.persistence.save_config(config)
    
    def load_persisted_data(self):
        """Carrega dados persistidos"""
        try:
            self.add_log_message("Carregando configurações e histórico...")
            
            config = self.persistence.load_config()
            
            if config.get('trabalho_dir'):
                self.trabalho_dir = config['trabalho_dir']
                self.dir_entry.setText(self.trabalho_dir)
                self.validate_config()
            
            if config.get('max_threads'):
                self.max_threads = config['max_threads']
                self.threads_combo.setCurrentText(str(self.max_threads))
            
            if config.get('verbose_mode'):
                self.verbose_checkbox.setChecked(config['verbose_mode'])
            
            if config.get('preferred_sheet'):
                self.preferred_sheet = config['preferred_sheet']
                self.sheet_entry.setText(self.preferred_sheet)
            
            # Carrega histórico
            self.processing_history = self.persistence.load_all_history_entries()
            self.update_history_display()
            
            attention_count = sum(1 for h in self.processing_history if h.has_attention)
            if attention_count > 0:
                self.add_log_message(f"✅ Dados carregados: {len(self.processing_history)} entradas no histórico ({attention_count} com atenção)")
            else:
                self.add_log_message(f"✅ Dados carregados: {len(self.processing_history)} entradas no histórico")
            
        except Exception as e:
            self.add_log_message(f"⚠️ Erro ao carregar dados persistidos: {e}")
    
    def add_log_message(self, message):
        """Adiciona mensagem ao log"""
        # Verifica se current_logs existe antes de usar
        if not hasattr(self, 'current_logs'):
            self.current_logs = []
            
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}"
        self.current_logs.append(log_entry)
        
        # Mantém apenas últimos 100 logs
        if len(self.current_logs) > 100:
            self.current_logs = self.current_logs[-100:]
    
    def closeEvent(self, event):
        """Manipula fechamento da aplicação"""
        if self.processing:
            reply = QMessageBox.question(
                self,
                "Processamento em Andamento",
                "Há processamentos em andamento. Deseja realmente fechar a aplicação?\n\nO processamento será interrompido.",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No  # Botão padrão
            )
            
            if reply == QMessageBox.StandardButton.No:
                event.ignore()
                return
        
        self.save_current_config()
        event.accept()

def main():
    """Função principal com splash screen"""
    try:
        app = QApplication(sys.argv)
        
        # Aplica estilo escuro
        app.setStyleSheet(DARK_STYLE)
        
        # Cria e mostra splash screen
        splash = SplashScreen()
        splash.start_loading()
        
        # Permite que splash screen seja mostrada
        app.processEvents()
        
        # Detecção do ambiente (executável vs desenvolvimento)
        is_frozen = getattr(sys, 'frozen', False)
        
        if is_frozen:
            # Em executável, carregamento pode ser mais lento
            splash.set_status("Inicializando executável...")
            time.sleep(0.4)
        else:
            # Em desenvolvimento
            splash.set_status("Modo desenvolvimento...")
            time.sleep(0.2)
        
        # Simula carregamento das dependências principais
        splash.set_status("Carregando processador...")
        app.processEvents()
        time.sleep(0.3 if is_frozen else 0.2)
        
        # Verifica importações críticas
        splash.set_status("Verificando dependências...")
        app.processEvents()
        time.sleep(0.2)
        
        try:
            # Testa se consegue importar o processador
            from pdf_processor_core import PDFProcessorCore
            splash.set_status("✅ Dependências OK")
            app.processEvents()
            time.sleep(0.2)
        except ImportError as e:
            splash.set_status("❌ Erro crítico")
            app.processEvents()
            time.sleep(1)
            splash.close()
            QMessageBox.critical(None, "Erro Crítico de Dependência", 
                               f"❌ Não foi possível carregar o módulo principal:\n\n"
                               f"📄 Arquivo: pdf_processor_core.py\n"
                               f"❓ Status: Não encontrado\n\n"
                               f"🔧 Solução:\n"
                               f"Certifique-se de que todos os arquivos estão na mesma pasta do executável.\n\n"
                               f"📋 Erro técnico: {e}")
            sys.exit(1)
        
        # Carregamento da interface principal
        splash.set_status("Criando interface...")
        app.processEvents()
        time.sleep(0.3 if is_frozen else 0.2)
        
        # Cria janela principal
        splash.set_status("Inicializando PyQt6...")
        app.processEvents()
        
        # Cria window
        window = MainWindow()
        
        splash.set_status("Configurando sistema...")
        app.processEvents()
        time.sleep(0.2 if is_frozen else 0.1)
        
        splash.set_status("Preparando interface...")
        app.processEvents()
        time.sleep(0.2)
        
        # Aguarda até splash screen chegar a pelo menos 80%
        while splash.progress_value < 80:
            app.processEvents()
            time.sleep(0.02)
        
        splash.set_status("Carregando dados...")
        app.processEvents()
        time.sleep(0.2 if is_frozen else 0.1)
        
        # Aguarda carregamento completo
        while splash.progress_value < 100:
            app.processEvents()
            time.sleep(0.02)
        
        # Finaliza splash e mostra janela principal
        splash.set_status("✅ Pronto! Abrindo...")
        app.processEvents()
        time.sleep(0.3)
        
        # Fecha splash e mostra janela principal
        splash.finish_loading(window)
        window.show()
        
        # Adiciona logs de inicialização bem-sucedida
        window.add_log_message("🚀 Aplicação PyQt6 v4.0.1 iniciada com sucesso!")
        window.add_log_message("💡 Interface moderna com performance nativa carregada")
        window.add_log_message("⚡ Funcionalidades automáticas ativas")
        window.add_log_message("🔧 Processamento unificado (sempre ThreadPoolExecutor) ativo")
        if is_frozen:
            window.add_log_message("📦 Executando em modo executável (.exe)")
        else:
            window.add_log_message("🛠️ Executando em modo desenvolvimento")
        
        sys.exit(app.exec())
        
    except Exception as e:
        # Tratamento de erro crítico
        try:
            if 'splash' in locals():
                splash.close()
            QMessageBox.critical(None, "Erro Crítico na Inicialização", 
                               f"❌ Erro inesperado ao iniciar a aplicação:\n\n"
                               f"📋 Detalhes do erro:\n{e}\n\n"
                               f"🔧 A aplicação será fechada.\n\n"
                               f"💡 Tente executar novamente ou verifique se todos os arquivos estão presentes.")
        except:
            pass  # Se nem QMessageBox funcionar, apenas termina
        
        sys.exit(1)

if __name__ == "__main__":
    main()