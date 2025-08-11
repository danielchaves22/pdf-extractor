#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF para Excel Desktop App - Interface Gráfica
==============================================

Interface gráfica moderna usando CustomTkinter que utiliza o módulo
pdf_processor_core.py para toda a lógica de processamento.

Versão 3.3.2 - OTIMIZAÇÕES DE INTERFACE:
- Painel de informações removido para mais espaço
- Lista de arquivos expandida (altura 220px)
- Interface compacta e organizada
- Contador dinâmico no header
- Melhor aproveitamento do espaço vertical

Funcionalidades v3.3:
- Seleção múltipla de PDFs com interface simplificada
- Processamento paralelo com controle de threads
- Interface atualizada com botão sempre visível
- Popup de progresso com status individual de cada PDF
- Histórico em lote

Dependências:
pip install customtkinter pillow

Como usar:
python desktop_app.py

Autor: Sistema de Extração Automatizada
Data: 2025
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import os
import sys
import json
import uuid
import subprocess
from pathlib import Path
from datetime import datetime
import webbrowser
from concurrent.futures import ThreadPoolExecutor, as_completed
import queue
from typing import List, Dict, Optional
from dataclasses import dataclass

# Importa drag and drop se disponível
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except ImportError:
    HAS_DND = False
    # Fallback para tkinter normal
    TkinterDnD = tk

# Importa o processador core
try:
    from pdf_processor_core import PDFProcessorCore
except ImportError:
    print("ERRO: Módulo pdf_processor_core.py não encontrado!")
    print("Certifique-se de que o arquivo pdf_processor_core.py está na mesma pasta.")
    sys.exit(1)

# Configuração do CustomTkinter
ctk.set_appearance_mode("dark")  # "System", "Dark", "Light"
ctk.set_default_color_theme("blue")  # "blue", "green", "dark-blue"

@dataclass
class PDFProcessingStatus:
    """Status de processamento de um PDF individual"""
    filename: str
    status: str  # 'waiting', 'processing', 'completed', 'error'
    progress: int = 0
    message: str = ""
    result_data: Optional[Dict] = None
    logs: List[str] = None
    
    def __post_init__(self):
        if self.logs is None:
            self.logs = []

class BatchProcessor:
    """Gerenciador de processamento em lote de PDFs"""
    
    def __init__(self, max_workers=3):
        self.max_workers = max_workers
        self.status_queue = queue.Queue()
        self.pdf_statuses = {}
        self.completed_count = 0
        self.total_count = 0
        self.is_running = False
        
    def add_pdfs(self, pdf_files: List[str]):
        """Adiciona PDFs para processamento"""
        self.total_count = len(pdf_files)
        self.completed_count = 0
        self.pdf_statuses = {}
        
        for pdf_file in pdf_files:
            filename = Path(pdf_file).name
            self.pdf_statuses[filename] = PDFProcessingStatus(
                filename=filename,
                status='waiting'
            )
    
    def get_status(self, filename: str) -> Optional[PDFProcessingStatus]:
        """Obtém status de um PDF específico"""
        return self.pdf_statuses.get(filename)
    
    def update_status(self, filename: str, status: str, progress: int = 0, 
                     message: str = "", result_data: Dict = None, log_message: str = None):
        """Atualiza status de um PDF"""
        if filename in self.pdf_statuses:
            pdf_status = self.pdf_statuses[filename]
            pdf_status.status = status
            pdf_status.progress = progress
            pdf_status.message = message
            
            if result_data:
                pdf_status.result_data = result_data
            
            if log_message:
                pdf_status.logs.append(log_message)
            
            # Envia atualização para interface
            self.status_queue.put(('update', filename, pdf_status))
            
            # Conta conclusões
            if status in ['completed', 'error']:
                self.completed_count += 1
                if self.completed_count >= self.total_count:
                    self.is_running = False
                    self.status_queue.put(('batch_complete', None, None))
    
    def process_batch(self, pdf_files: List[str], processor_factory, trabalho_dir: str):
        """Processa lote de PDFs em paralelo"""
        self.add_pdfs(pdf_files)
        self.is_running = True
        
        def process_single_pdf(pdf_file):
            """Processa um único PDF"""
            filename = Path(pdf_file).name
            
            try:
                # Cria processador individual para esta thread
                processor = processor_factory()
                processor.set_trabalho_dir(trabalho_dir)
                
                # Callbacks customizados para este PDF
                def progress_callback(progress, message=""):
                    self.update_status(filename, 'processing', progress, message)
                
                def log_callback(log_message):
                    self.update_status(filename, 'processing', 
                                     log_message=f"[{filename}] {log_message}")
                
                # Configura callbacks
                processor.progress_callback = progress_callback
                processor.log_callback = log_callback
                
                # Inicia processamento
                self.update_status(filename, 'processing', 0, "Iniciando...")
                
                # Processa PDF
                if Path(pdf_file).parent == Path(trabalho_dir):
                    pdf_filename = filename
                else:
                    pdf_filename = pdf_file
                
                results = processor.process_pdf(pdf_filename)
                
                # Atualiza status final
                if results['success']:
                    self.update_status(filename, 'completed', 100, 
                                     f"✅ {results['total_extracted']} períodos processados", 
                                     results)
                else:
                    self.update_status(filename, 'error', 0, 
                                     f"❌ {results['error']}", results)
                
                return results
                
            except Exception as e:
                error_result = {'success': False, 'error': str(e)}
                self.update_status(filename, 'error', 0, f"❌ Erro: {str(e)}", error_result)
                return error_result
        
        # Executa processamento paralelo
        def run_parallel():
            with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
                future_to_pdf = {
                    executor.submit(process_single_pdf, pdf_file): pdf_file 
                    for pdf_file in pdf_files
                }
                
                for future in as_completed(future_to_pdf):
                    pdf_file = future_to_pdf[future]
                    try:
                        result = future.result()
                    except Exception as e:
                        filename = Path(pdf_file).name
                        self.update_status(filename, 'error', 0, f"❌ Exceção: {str(e)}")
        
        # Inicia processamento em thread separada
        thread = threading.Thread(target=run_parallel, daemon=True)
        thread.start()

class PersistenceManager:
    """Gerencia persistência de configurações e histórico"""
    
    def __init__(self, app_dir=None):
        """Inicializa gerenciador de persistência"""
        if app_dir is None:
            # Usa diretório do executável ou script
            if getattr(sys, 'frozen', False):
                # Executável PyInstaller
                self.app_dir = Path(sys.executable).parent
            else:
                # Script Python
                self.app_dir = Path(__file__).parent
        else:
            self.app_dir = Path(app_dir)
        
        # Cria diretório .data para arquivos de persistência
        self.data_dir = self.app_dir / ".data"
        self.data_dir.mkdir(exist_ok=True)
        
        self.config_file = self.data_dir / "config.json"
        self.history_file = self.data_dir / "history.json"
        
        # ID único para esta sessão
        self.session_id = str(uuid.uuid4())
        self.session_start = datetime.now()
    
    def load_config(self):
        """Carrega configurações salvas"""
        try:
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            return {}
        except Exception as e:
            print(f"Erro ao carregar configurações: {e}")
            return {}
    
    def save_config(self, config_data):
        """Salva configurações"""
        try:
            config_data['last_saved'] = datetime.now().isoformat()
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"Erro ao salvar configurações: {e}")
    
    def load_history(self, max_sessions=10):
        """Carrega histórico das últimas sessões"""
        try:
            if self.history_file.exists():
                with open(self.history_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    # Retorna apenas as últimas N sessões
                    sessions = data.get('sessions', [])[-max_sessions:]
                    return sessions
            return []
        except Exception as e:
            print(f"Erro ao carregar histórico: {e}")
            return []
    
    def save_history_entry(self, entry_data):
        """Adiciona entrada ao histórico da sessão atual"""
        try:
            # Carrega histórico existente
            history_data = {'sessions': []}
            if self.history_file.exists():
                with open(self.history_file, 'r', encoding='utf-8') as f:
                    history_data = json.load(f)
            
            # Procura sessão atual ou cria nova
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
            
            # Adiciona entrada
            current_session['entries'].append({
                'timestamp': entry_data.timestamp.isoformat(),
                'pdf_file': entry_data.pdf_file,
                'success': entry_data.success,
                'result_data': entry_data.result_data,
                'logs': entry_data.logs[:50],  # Limita logs para economizar espaço
                'is_batch': getattr(entry_data, 'is_batch', False),
                'batch_info': getattr(entry_data, 'batch_info', {})
            })
            
            # Mantém apenas últimas 10 sessões
            history_data['sessions'] = history_data['sessions'][-10:]
            
            # Salva
            with open(self.history_file, 'w', encoding='utf-8') as f:
                json.dump(history_data, f, indent=2, ensure_ascii=False)
                
        except Exception as e:
            print(f"Erro ao salvar entrada do histórico: {e}")
    
    def clear_history(self):
        """Limpa todo o histórico"""
        try:
            if self.history_file.exists():
                self.history_file.unlink()
        except Exception as e:
            print(f"Erro ao limpar histórico: {e}")
    
    def load_all_history_entries(self):
        """Carrega todas as entradas de histórico de todas as sessões"""
        try:
            sessions = self.load_history()
            all_entries = []
            
            for session in sessions:
                for entry_data in session.get('entries', []):
                    # Reconstrói objeto HistoryEntry
                    entry = HistoryEntry(
                        timestamp=datetime.fromisoformat(entry_data['timestamp']),
                        pdf_file=entry_data['pdf_file'],
                        success=entry_data['success'],
                        result_data=entry_data['result_data'],
                        logs=entry_data['logs'],
                        is_batch=entry_data.get('is_batch', False),
                        batch_info=entry_data.get('batch_info', {})
                    )
                    all_entries.append(entry)
            
            return all_entries
        except Exception as e:
            print(f"Erro ao carregar entradas do histórico: {e}")
            return []

class HistoryEntry:
    """Representa um entrada no histórico de processamentos"""
    def __init__(self, timestamp, pdf_file, success, result_data, logs, 
                 is_batch=False, batch_info=None):
        self.timestamp = timestamp
        self.pdf_file = pdf_file
        self.success = success
        self.result_data = result_data
        self.logs = logs
        self.is_batch = is_batch
        self.batch_info = batch_info if batch_info is not None else {}

class BatchProcessingPopup:
    """Popup para mostrar progresso de processamento em lote"""
    
    def __init__(self, parent):
        self.parent = parent
        self.window = None
        self.pdf_frames = {}
        self.main_progress_bar = None
        self.main_status_label = None
        self.scroll_frame = None
        
    def show(self, pdf_files: List[str]):
        """Mostra o popup de processamento em lote"""
        try:
            self.window = ctk.CTkToplevel(self.parent.root)
            self.window.title(f"Processando {len(pdf_files)} PDFs...")
            self.window.geometry("800x600")
            
            # Configurações básicas
            try:
                self.window.attributes('-alpha', 1.0)
                self.window.transient(self.parent.root)
                self.window.grab_set()  # Torna modal
                self.window.resizable(False, False)
            except:
                pass
            
            self.window.protocol("WM_DELETE_WINDOW", self._on_close_attempt)
            
            # Cria interface
            self._create_interface(pdf_files)
            
            # Posicionamento
            try:
                parent_x = self.parent.root.winfo_x()
                parent_y = self.parent.root.winfo_y()
                x = parent_x + 25
                y = parent_y + 25
                self.window.geometry(f"800x600+{x}+{y}")
            except:
                self.window.geometry("800x600+200+150")
            
            # Foco
            try:
                self.window.focus_force()
                self.window.lift()
            except:
                pass
                
        except Exception as e:
            print(f"Erro ao criar popup de processamento em lote: {e}")
            self.window = None
    
    def _create_interface(self, pdf_files: List[str]):
        """Cria interface do popup"""
        if not self.window:
            return
        
        try:
            # Header
            header_frame = ctk.CTkFrame(self.window)
            header_frame.pack(fill="x", padx=20, pady=(20, 10))
            
            title_label = ctk.CTkLabel(
                header_frame,
                text=f"🔄 Processando {len(pdf_files)} PDFs em Paralelo",
                font=ctk.CTkFont(size=16, weight="bold")
            )
            title_label.pack(pady=15)
            
            # Progresso geral
            general_frame = ctk.CTkFrame(self.window)
            general_frame.pack(fill="x", padx=20, pady=(0, 10))
            
            ctk.CTkLabel(
                general_frame,
                text="📊 Progresso Geral",
                font=ctk.CTkFont(size=14, weight="bold")
            ).pack(pady=(15, 5))
            
            self.main_progress_bar = ctk.CTkProgressBar(general_frame, height=20)
            self.main_progress_bar.pack(fill="x", padx=20, pady=(0, 10))
            self.main_progress_bar.set(0)
            
            self.main_status_label = ctk.CTkLabel(
                general_frame,
                text="Iniciando processamento...",
                font=ctk.CTkFont(size=12)
            )
            self.main_status_label.pack(padx=20, pady=(0, 15))
            
            # Lista de PDFs
            pdfs_frame = ctk.CTkFrame(self.window)
            pdfs_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))
            
            ctk.CTkLabel(
                pdfs_frame,
                text="📝 Status Individual dos PDFs",
                font=ctk.CTkFont(size=14, weight="bold")
            ).pack(pady=(15, 10))
            
            # Frame scrollável para lista de PDFs
            self.scroll_frame = ctk.CTkScrollableFrame(pdfs_frame)
            self.scroll_frame.pack(fill="both", expand=True, padx=15, pady=(0, 15))
            
            # Cria frames para cada PDF
            for pdf_file in pdf_files:
                filename = Path(pdf_file).name
                self._create_pdf_frame(filename)
                
        except Exception as e:
            print(f"Erro ao criar interface do popup em lote: {e}")
    
    def _create_pdf_frame(self, filename: str):
        """Cria frame para um PDF individual"""
        # Frame principal do PDF
        pdf_frame = ctk.CTkFrame(self.scroll_frame)
        pdf_frame.pack(fill="x", padx=5, pady=2)
        
        # Frame interno
        inner_frame = ctk.CTkFrame(pdf_frame, fg_color="transparent")
        inner_frame.pack(fill="x", padx=10, pady=8)
        
        # Ícone de status
        status_icon = ctk.CTkLabel(
            inner_frame,
            text="⏳",
            font=ctk.CTkFont(size=16),
            width=30
        )
        status_icon.pack(side="left", padx=(0, 10))
        
        # Informações do arquivo
        info_frame = ctk.CTkFrame(inner_frame, fg_color="transparent")
        info_frame.pack(side="left", fill="x", expand=True)
        
        # Nome do arquivo
        name_label = ctk.CTkLabel(
            info_frame,
            text=f"📄 {filename}",
            font=ctk.CTkFont(size=12, weight="bold"),
            anchor="w"
        )
        name_label.pack(fill="x")
        
        # Status
        status_label = ctk.CTkLabel(
            info_frame,
            text="Aguardando...",
            font=ctk.CTkFont(size=10),
            anchor="w"
        )
        status_label.pack(fill="x", pady=(2, 0))
        
        # Barra de progresso individual
        progress_bar = ctk.CTkProgressBar(info_frame, height=8)
        progress_bar.pack(fill="x", pady=(5, 0))
        progress_bar.set(0)
        
        # Armazena referências
        self.pdf_frames[filename] = {
            'frame': pdf_frame,
            'status_icon': status_icon,
            'name_label': name_label,
            'status_label': status_label,
            'progress_bar': progress_bar
        }
    
    def update_pdf_status(self, filename: str, status: PDFProcessingStatus):
        """Atualiza status de um PDF específico"""
        if filename not in self.pdf_frames:
            return
        
        frame_refs = self.pdf_frames[filename]
        
        # Atualiza ícone
        if status.status == 'waiting':
            frame_refs['status_icon'].configure(text="⏳")
        elif status.status == 'processing':
            frame_refs['status_icon'].configure(text="🔄")
        elif status.status == 'completed':
            frame_refs['status_icon'].configure(text="✅")
        elif status.status == 'error':
            frame_refs['status_icon'].configure(text="❌")
        
        # Atualiza status
        frame_refs['status_label'].configure(text=status.message)
        
        # Atualiza progresso
        frame_refs['progress_bar'].set(status.progress / 100)
    
    def update_general_progress(self, completed: int, total: int, message: str = ""):
        """Atualiza progresso geral"""
        if self.main_progress_bar:
            progress = completed / total if total > 0 else 0
            self.main_progress_bar.set(progress)
        
        if self.main_status_label:
            if message:
                self.main_status_label.configure(text=message)
            else:
                self.main_status_label.configure(
                    text=f"Processados: {completed}/{total} PDFs"
                )
    
    def _on_close_attempt(self):
        """Impede fechamento durante processamento"""
        if self.parent.processing:
            try:
                messagebox.showwarning(
                    "Processamento em andamento",
                    "Aguarde a conclusão do processamento em lote."
                )
            except:
                pass
        else:
            self.close()
    
    def close(self):
        """Fecha o popup"""
        if self.window:
            try:
                self.window.grab_release()  # Libera modal
                self.window.destroy()
                self.window = None
                
                # Força foco na janela principal
                self.parent.root.focus_force()
                self.parent.root.lift()
            except Exception as e:
                print(f"Erro ao fechar popup em lote: {e}")
                self.window = None

class HistoryDetailsWindow:
    """Janela modal para detalhes do histórico com estilo melhorado"""
    
    _instance = None  # Controle de instância única
    
    def __init__(self, parent, entry):
        # Controle de instância única
        if HistoryDetailsWindow._instance is not None:
            try:
                HistoryDetailsWindow._instance.close()
            except:
                pass
        
        self.parent = parent
        self.entry = entry
        self.window = None
        HistoryDetailsWindow._instance = self
        
        self._create_window()
    
    def _create_window(self):
        """Cria a janela de detalhes com estilo melhorado"""
        try:
            self.window = ctk.CTkToplevel(self.parent.root)
            self.window.title("📝 Detalhes do Histórico")
            self.window.geometry("750x600")
            
            # Configurações de janela modal
            try:
                self.window.transient(self.parent.root)
                self.window.grab_set()  # Torna modal
                self.window.resizable(True, True)
                self.window.attributes('-alpha', 1.0)
            except:
                pass
            
            self.window.protocol("WM_DELETE_WINDOW", self.close)
            
            # Criar interface
            self._create_interface()
            
            # Posicionamento
            self._center_window()
            
            # Foco
            try:
                self.window.focus_force()
                self.window.lift()
            except:
                pass
                
        except Exception as e:
            print(f"Erro ao criar janela de detalhes: {e}")
            self.window = None
    
    def _create_interface(self):
        """Cria a interface da janela com estilo CustomTkinter"""
        if not self.window:
            return
        
        # Container principal
        main_frame = ctk.CTkFrame(self.window)
        main_frame.pack(fill="both", expand=True, padx=15, pady=15)
        
        # Header com informações principais
        header_frame = ctk.CTkFrame(main_frame)
        header_frame.pack(fill="x", padx=15, pady=(15, 10))
        
        # Título
        title_text = f"📄 {Path(self.entry.pdf_file).stem}"
        if self.entry.is_batch and self.entry.batch_info.get('batch_size', 0) > 1:
            title_text += f" (Lote de {self.entry.batch_info['batch_size']})"
        
        title_label = ctk.CTkLabel(
            header_frame,
            text=title_text,
            font=ctk.CTkFont(size=18, weight="bold")
        )
        title_label.pack(pady=(15, 5))
        
        # Status e timestamp
        status_color = "#2cc985" if self.entry.success else "#f44336"
        status_text = "✅ Processamento bem-sucedido" if self.entry.success else "❌ Processamento com erro"
        
        status_label = ctk.CTkLabel(
            header_frame,
            text=status_text,
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=status_color
        )
        status_label.pack(pady=(0, 5))
        
        timestamp_label = ctk.CTkLabel(
            header_frame,
            text=f"🕒 {self.entry.timestamp.strftime('%d/%m/%Y %H:%M:%S')}",
            font=ctk.CTkFont(size=12),
            text_color="gray"
        )
        timestamp_label.pack(pady=(0, 15))
        
        # Informações de resultado
        if self.entry.success and self.entry.result_data:
            info_frame = ctk.CTkFrame(main_frame)
            info_frame.pack(fill="x", padx=15, pady=(0, 10))
            
            info_title = ctk.CTkLabel(
                info_frame,
                text="📊 Informações do Processamento",
                font=ctk.CTkFont(size=14, weight="bold")
            )
            info_title.pack(pady=(15, 10))
            
            # Grade de informações
            info_grid = ctk.CTkFrame(info_frame, fg_color="transparent")
            info_grid.pack(fill="x", padx=15, pady=(0, 15))
            
            info_items = []
            
            if self.entry.result_data.get('person_name'):
                info_items.append(("👤 Pessoa", self.entry.result_data['person_name']))
            
            if self.entry.result_data.get('total_extracted') is not None:
                info_items.append(("📋 Períodos processados", str(self.entry.result_data['total_extracted'])))
            
            if self.entry.result_data.get('arquivo_final'):
                info_items.append(("💾 Arquivo criado", self.entry.result_data['arquivo_final']))
            
            if self.entry.result_data.get('folha_normal_periods'):
                info_items.append(("📄 FOLHA NORMAL", f"{self.entry.result_data['folha_normal_periods']} períodos"))
            
            if self.entry.result_data.get('salario_13_periods'):
                info_items.append(("💰 13º SALÁRIO", f"{self.entry.result_data['salario_13_periods']} períodos"))
            
            # Exibe informações em pares
            for i in range(0, len(info_items), 2):
                row_frame = ctk.CTkFrame(info_grid, fg_color="transparent")
                row_frame.pack(fill="x", pady=2)
                
                # Item da esquerda
                left_item = info_items[i]
                left_frame = ctk.CTkFrame(row_frame)
                left_frame.pack(side="left", fill="x", expand=True, padx=(0, 5))
                
                ctk.CTkLabel(
                    left_frame,
                    text=left_item[0],
                    font=ctk.CTkFont(size=11, weight="bold"),
                    anchor="w"
                ).pack(fill="x", padx=10, pady=(5, 0))
                
                ctk.CTkLabel(
                    left_frame,
                    text=left_item[1],
                    font=ctk.CTkFont(size=11),
                    anchor="w"
                ).pack(fill="x", padx=10, pady=(0, 5))
                
                # Item da direita (se existir)
                if i + 1 < len(info_items):
                    right_item = info_items[i + 1]
                    right_frame = ctk.CTkFrame(row_frame)
                    right_frame.pack(side="right", fill="x", expand=True, padx=(5, 0))
                    
                    ctk.CTkLabel(
                        right_frame,
                        text=right_item[0],
                        font=ctk.CTkFont(size=11, weight="bold"),
                        anchor="w"
                    ).pack(fill="x", padx=10, pady=(5, 0))
                    
                    ctk.CTkLabel(
                        right_frame,
                        text=right_item[1],
                        font=ctk.CTkFont(size=11),
                        anchor="w"
                    ).pack(fill="x", padx=10, pady=(0, 5))
        
        # Área de logs
        logs_frame = ctk.CTkFrame(main_frame)
        logs_frame.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        
        logs_title = ctk.CTkLabel(
            logs_frame,
            text="📄 Logs Detalhados",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        logs_title.pack(pady=(15, 10))
        
        # Textbox para logs com estilo melhorado
        self.logs_textbox = ctk.CTkTextbox(
            logs_frame,
            wrap="word",
            font=ctk.CTkFont(family="Consolas", size=11),
            corner_radius=8
        )
        self.logs_textbox.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        
        # Popula logs
        self._populate_logs()
        
        # Botões de ação
        buttons_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        buttons_frame.pack(fill="x", padx=15, pady=(0, 15))
        
        # Botão fechar
        close_button = ctk.CTkButton(
            buttons_frame,
            text="✖️ Fechar",
            command=self.close,
            width=100,
            height=35,
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color="#dc3545",
            hover_color="#c82333"
        )
        close_button.pack(side="right", padx=(10, 0))
        
        # Botão abrir arquivo (apenas para sucessos)
        if self.entry.success and self.entry.result_data.get('excel_path'):
            open_button = ctk.CTkButton(
                buttons_frame,
                text="📂 Abrir Arquivo",
                command=self._open_file,
                width=120,
                height=35,
                font=ctk.CTkFont(size=12, weight="bold"),
                fg_color="#28a745",
                hover_color="#218838"
            )
            open_button.pack(side="right")
    
    def _populate_logs(self):
        """Popula a área de logs com formatação melhorada"""
        try:
            # Cabeçalho com informações básicas
            header_info = [
                f"📄 Arquivo: {self.entry.pdf_file}",
                f"🕒 Data/Hora: {self.entry.timestamp.strftime('%d/%m/%Y %H:%M:%S')}",
                f"✅ Resultado: {'Sucesso' if self.entry.success else 'Falha'}",
                ""
            ]
            
            # Informações específicas
            if self.entry.success and self.entry.result_data:
                if self.entry.result_data.get('arquivo_final'):
                    header_info.append(f"💾 Arquivo Excel: {self.entry.result_data['arquivo_final']}")
                if self.entry.result_data.get('person_name'):
                    header_info.append(f"👤 Pessoa: {self.entry.result_data['person_name']}")
                if self.entry.result_data.get('total_extracted') is not None:
                    header_info.append(f"📊 Períodos processados: {self.entry.result_data['total_extracted']}")
                if self.entry.result_data.get('total_pages'):
                    header_info.append(f"📑 Páginas analisadas: {self.entry.result_data['total_pages']}")
            else:
                if self.entry.result_data and self.entry.result_data.get('error'):
                    header_info.append(f"❌ Erro: {self.entry.result_data['error']}")
            
            header_info.extend(["", "📄 Logs Detalhados:", "=" * 50, ""])
            
            # Combina header com logs
            if self.entry.logs:
                all_lines = header_info + self.entry.logs
            else:
                all_lines = header_info + ["Nenhum log detalhado disponível."]
            
            # Insere no textbox
            content = "\n".join(all_lines)
            self.logs_textbox.insert("1.0", content)
            self.logs_textbox.configure(state="disabled")
            
        except Exception as e:
            error_text = f"Erro ao carregar logs: {str(e)}"
            self.logs_textbox.insert("1.0", error_text)
            self.logs_textbox.configure(state="disabled")
    
    def _open_file(self):
        """Abre o arquivo Excel associado"""
        try:
            file_path = self.entry.result_data.get('excel_path')
            if not file_path and hasattr(self.parent, 'trabalho_dir') and self.parent.trabalho_dir:
                arquivo_rel = self.entry.result_data.get('arquivo_final')
                if arquivo_rel:
                    file_path = Path(self.parent.trabalho_dir) / arquivo_rel
            
            if not file_path:
                messagebox.showerror("Erro", "Caminho do arquivo não disponível.")
                return
            
            file_path = Path(file_path)
            if not file_path.exists():
                messagebox.showerror("Erro", f"Arquivo não encontrado: {file_path}")
                return
            
            if sys.platform.startswith('win'):
                os.startfile(file_path)  # type: ignore[attr-defined]
            elif sys.platform == 'darwin':
                subprocess.Popen(['open', str(file_path)])
            else:
                subprocess.Popen(['xdg-open', str(file_path)])
        
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir o arquivo: {e}")
    
    def _center_window(self):
        """Centraliza a janela em relação à janela principal"""
        try:
            # Atualiza geometria
            self.window.update_idletasks()
            
            # Obtém dimensões
            width = 750
            height = 600
            
            # Obtém posição da janela principal
            parent_x = self.parent.root.winfo_x()
            parent_y = self.parent.root.winfo_y()
            parent_width = self.parent.root.winfo_width()
            parent_height = self.parent.root.winfo_height()
            
            # Calcula posição centralizada
            x = parent_x + (parent_width - width) // 2
            y = parent_y + (parent_height - height) // 2
            
            # Garante que a janela não saia da tela
            x = max(0, x)
            y = max(0, y)
            
            self.window.geometry(f"{width}x{height}+{x}+{y}")
            
        except Exception:
            # Fallback para posição padrão
            self.window.geometry("750x600+200+100")
    
    def close(self):
        """Fecha a janela e limpa referências"""
        if self.window:
            try:
                self.window.grab_release()  # Libera modal
                self.window.destroy()
                self.window = None
                
                # Foco na janela principal
                self.parent.root.focus_force()
                self.parent.root.lift()
                
            except Exception as e:
                print(f"Erro ao fechar janela de detalhes: {e}")
                self.window = None
            finally:
                HistoryDetailsWindow._instance = None

class PDFExcelDesktopApp:
    def __init__(self):
        # Configuração da janela principal usando CustomTkinter
        if HAS_DND:
            class CTkDnD(ctk.CTk, TkinterDnD.DnDWrapper):
                """Janela principal com suporte a Drag & Drop"""
                pass
            self.root = CTkDnD()
        else:
            self.root = ctk.CTk()
        
        self.root.title("Processamento de Folha de Pagamento v3.3.2")
        self.root.geometry("950x600")
        self.root.resizable(False, False)
        
        # Remove qualquer transparência da janela principal
        self.root.attributes('-alpha', 1.0)
        
        # Gerenciador de persistência
        self.persistence = PersistenceManager()
        
        # Variáveis de estado
        self.selected_files = []  # Agora lista de arquivos
        self.trabalho_dir = None
        self.processing = False
        
        # Processamento em lote
        self.batch_processor = BatchProcessor(max_workers=3)  # 3 PDFs em paralelo
        self.batch_popup = None
        
        # Histórico de processamentos
        self.processing_history = []
        self.current_logs = []
        
        # Variáveis da interface
        self.verbose_var = ctk.BooleanVar()
        self.sheet_entry = None
        self.max_threads_var = ctk.IntVar(value=3)
        
        # Processador será inicializado sob demanda
        self.processor = None

        # Configuração de estilo
        self.setup_styles()
        
        # Cria a interface
        self.create_interface()
        
        # Configura drag and drop
        self.setup_drag_drop()
        
        # Carrega configurações iniciais
        self.load_initial_config()

    def _get_processor(self):
        """Factory para criar processadores individuais"""
        def create_processor():
            processor = PDFProcessorCore()
            if self.sheet_entry and self.sheet_entry.get().strip():
                processor.preferred_sheet = self.sheet_entry.get().strip()
            return processor
        return create_processor

    def setup_styles(self):
        """Define cores e estilos customizados"""
        self.colors = {
            'primary': "#1f538d",
            'secondary': "#14375e",
            'success': "#2cc985",
            'warning': "#ffa726",
            'error': "#f44336",
            'bg_dark': "#212121",
            'bg_light': "#2b2b2b",
            'text_primary': "#ffffff",
            'text_secondary': "#b0b0b0"
        }

    def create_interface(self):
        """Cria a interface gráfica principal com abas"""
        # Container principal
        main_frame = ctk.CTkFrame(self.root)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Header global
        self.create_global_header(main_frame)
        
        # Sistema de abas
        self.tabview = ctk.CTkTabview(main_frame)
        self.tabview.pack(fill="both", expand=True, pady=(10, 0))
        
        # Cria as abas
        self.tab_processing = self.tabview.add("📄 Processamento")
        self.tab_history = self.tabview.add("📊 Histórico")
        self.tab_settings = self.tabview.add("⚙️ Configurações")
        
        # Configura abas
        self.create_processing_tab()
        self.create_history_tab()
        self.create_settings_tab()
        
        # Define aba inicial
        self.tabview.set("📄 Processamento")
        
        # Carrega dados persistidos
        self.root.after(100, self.load_persisted_data)

    def load_persisted_data(self):
        """Carrega dados persistidos"""
        try:
            config = self.persistence.load_config()
            
            if config.get('trabalho_dir'):
                self.trabalho_dir = config['trabalho_dir']
                if hasattr(self, 'dir_entry') and self.dir_entry:
                    self.dir_entry.delete(0, 'end')
                    self.dir_entry.insert(0, self.trabalho_dir)
                    self.validate_config()
            
            if config.get('verbose_mode', False):
                self.verbose_var.set(True)
            
            if config.get('max_threads'):
                self.max_threads_var.set(config['max_threads'])
                self.batch_processor.max_workers = config['max_threads']
            
            if config.get('preferred_sheet') and hasattr(self, 'sheet_entry') and self.sheet_entry:
                self.sheet_entry.delete(0, 'end')
                self.sheet_entry.insert(0, config['preferred_sheet'])
            
            self.processing_history = self.persistence.load_all_history_entries()
            if hasattr(self, 'history_status_label'):
                self.update_history_display()
                
                if self.processing_history:
                    total = len(self.processing_history)
                    success_count = sum(1 for h in self.processing_history if h.success)
                    batch_count = sum(1 for h in self.processing_history if h.is_batch)
                    individual_count = total - batch_count
                    
                    status_text = f"{total} PDFs no histórico ({success_count} sucessos, {total - success_count} falhas)"
                    if batch_count > 0:
                        status_text += f" • {batch_count} de lotes, {individual_count} individuais"
                    
                    self.history_status_label.configure(text=status_text)
            
            self.add_log_message("Configurações e histórico carregados")
            
        except Exception as e:
            self.add_log_message(f"Erro ao carregar dados persistidos: {e}")

    def save_current_config(self):
        """Salva configuração atual"""
        try:
            config = {
                'trabalho_dir': self.trabalho_dir,
                'verbose_mode': self.verbose_var.get(),
                'max_threads': self.max_threads_var.get(),
            }
            
            if self.sheet_entry and self.sheet_entry.get().strip():
                config['preferred_sheet'] = self.sheet_entry.get().strip()
            
            self.persistence.save_config(config)
            
        except Exception as e:
            self.add_log_message(f"Erro ao salvar configuração: {e}")

    def create_global_header(self, parent):
        """Cria o cabeçalho global da aplicação"""
        header_frame = ctk.CTkFrame(parent)
        header_frame.pack(fill="x", pady=(0, 8))
        
        title_label = ctk.CTkLabel(
            header_frame,
            text="📄 Processamento de Folha de Pagamento",
            font=ctk.CTkFont(size=22, weight="bold")
        )
        title_label.pack(pady=(10, 6))
        
        subtitle_label = ctk.CTkLabel(
            header_frame,
            text="Automatização de folhas de pagamento PDF para Excel v3.3.2 - Interface Otimizada + Lista Expandida",
            font=ctk.CTkFont(size=11),
            text_color=self.colors['text_secondary']
        )
        subtitle_label.pack(pady=(0, 10))

    def create_processing_tab(self):
        """Cria o conteúdo da aba de processamento"""
        self.processing_container = ctk.CTkFrame(self.tab_processing)
        self.processing_container.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Configuração do diretório
        self.create_config_section()
        
        # Seleção de arquivos com botão processar integrado
        self.create_multiple_file_section()

    def create_multiple_file_section(self):
        """Cria seção de seleção de arquivos simplificada com mais espaço para lista"""
        file_frame = ctk.CTkFrame(self.processing_container)
        file_frame.pack(fill="x", pady=(0, 8))
        
        # Header com título e contador
        header_frame = ctk.CTkFrame(file_frame, fg_color="transparent")
        header_frame.pack(fill="x", padx=20, pady=(10, 8))
        
        file_title = ctk.CTkLabel(
            header_frame,
            text="📎 Seleção de Arquivos PDF",
            font=ctk.CTkFont(size=15, weight="bold"),
            anchor="w"
        )
        file_title.pack(side="left")
        
        # Contador de arquivos (atualizado dinamicamente)
        self.file_counter_label = ctk.CTkLabel(
            header_frame,
            text="0 arquivos",
            font=ctk.CTkFont(size=12),
            text_color=self.colors['text_secondary'],
            anchor="e"
        )
        self.file_counter_label.pack(side="right")

        # Labels principais para instruções de seleção
        self.drop_main_label = ctk.CTkLabel(
            file_frame,
            text="Nenhum arquivo selecionado",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=self.colors['text_secondary'],
        )
        self.drop_main_label.pack(fill="x", padx=20)

        self.drop_sub_label = ctk.CTkLabel(
            file_frame,
            text="🎯 Arraste PDFs aqui ou use o botão abaixo",
            font=ctk.CTkFont(size=11),
            text_color=self.colors['text_secondary'],
        )
        self.drop_sub_label.pack(fill="x", padx=20, pady=(0, 8))

        # Frame para botões de ação (compacto)
        action_frame = ctk.CTkFrame(file_frame, fg_color="transparent")
        action_frame.pack(fill="x", padx=20, pady=(0, 8))
        
        # Botão seleção 
        select_button = ctk.CTkButton(
            action_frame,
            text="📂 Selecionar PDFs",
            command=self.select_pdfs,
            width=160,
            height=32,
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color=self.colors['primary']
        )
        select_button.pack(side="left", padx=(0, 8))
        
        # Botão limpar seleção
        clear_button = ctk.CTkButton(
            action_frame,
            text="🗑️ Limpar",
            command=self.clear_selection,
            width=70,
            height=32,
            font=ctk.CTkFont(size=11),
            fg_color=self.colors['secondary']
        )
        clear_button.pack(side="left", padx=(0, 20))
        
        # Área de drop compacta (opcional para drag & drop)
        if HAS_DND:
            self.drop_frame = ctk.CTkFrame(action_frame, height=32, width=200)
            self.drop_frame.pack(side="left", padx=(0, 20))
            self.drop_frame.pack_propagate(False)
            
            drop_label = ctk.CTkLabel(
                self.drop_frame,
                text="🎯 Ou arraste PDFs aqui",
                font=ctk.CTkFont(size=11),
                text_color=self.colors['text_secondary']
            )
            drop_label.pack(expand=True)
            
            # Bind para clique na área de drop
            self.drop_frame.bind("<Button-1>", lambda e: self.select_pdfs())
            drop_label.bind("<Button-1>", lambda e: self.select_pdfs())
        
        # Botão processar (mais compacto)
        self.process_button = ctk.CTkButton(
            action_frame,
            text="🚀 Processar PDFs",
            command=self.process_pdfs,
            width=180,
            height=32,
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color=self.colors['success'],
            hover_color="#259b6e"
        )
        self.process_button.pack(side="right")
        
        # Lista de arquivos selecionados (MAIS ESPAÇO - altura aumentada)
        self.files_list_frame = ctk.CTkScrollableFrame(file_frame, height=220)
        self.files_list_frame.pack(fill="x", padx=20, pady=(0, 15))

    def select_pdfs(self):
        """Seleciona PDFs (um ou múltiplos)"""
        if not self.trabalho_dir:
            messagebox.showwarning("Aviso", "Configure o diretório de trabalho primeiro.")
            return
        
        file_paths = filedialog.askopenfilenames(
            title="Selecione arquivos PDF (um ou múltiplos)",
            initialdir=self.trabalho_dir,
            filetypes=[("Arquivos PDF", "*.pdf")]
        )
        
        if file_paths:
            self.selected_files = list(file_paths)
            self.update_selected_files_display()

    def clear_selection(self):
        """Limpa seleção de arquivos"""
        self.selected_files = []
        self.update_selected_files_display()

    def update_selected_files_display(self):
        """Atualiza display dos arquivos selecionados"""
        # Atualiza labels principais
        if not self.selected_files:
            self.drop_main_label.configure(
                text="Nenhum arquivo selecionado",
                text_color=self.colors['text_secondary']
            )
            self.drop_sub_label.configure(
                text="🎯 Arraste PDFs aqui ou use o botão abaixo",
                text_color=self.colors['text_secondary']
            )
            self.file_counter_label.configure(text="0 arquivos")
        else:
            count = len(self.selected_files)
            self.drop_main_label.configure(
                text=f"📄 {count} arquivo{'s' if count > 1 else ''} selecionado{'s' if count > 1 else ''}",
                text_color=self.colors['success']
            )
            self.file_counter_label.configure(
                text=f"{count} arquivo{'s' if count > 1 else ''}"
            )
            if count == 1:
                self.drop_sub_label.configure(
                    text="Processamento individual configurado",
                    text_color=self.colors['success']
                )
            else:
                self.drop_sub_label.configure(
                    text=f"Processamento paralelo configurado ({count} arquivos)",
                    text_color=self.colors['success']
                )
        
        # Limpa lista atual
        for widget in self.files_list_frame.winfo_children():
            widget.destroy()
        
        # Adiciona arquivos à lista
        for i, file_path in enumerate(self.selected_files):
            filename = Path(file_path).name
            
            file_item = ctk.CTkFrame(self.files_list_frame)
            file_item.pack(fill="x", padx=5, pady=2)
            
            # Número do arquivo
            num_label = ctk.CTkLabel(
                file_item,
                text=f"{i+1}.",
                font=ctk.CTkFont(size=11, weight="bold"),
                width=25
            )
            num_label.pack(side="left", padx=(10, 5), pady=5)
            
            # Nome do arquivo
            name_label = ctk.CTkLabel(
                file_item,
                text=filename,
                font=ctk.CTkFont(size=11),
                anchor="w"
            )
            name_label.pack(side="left", fill="x", expand=True, padx=(0, 10), pady=5)
            
            # Botão remover
            remove_button = ctk.CTkButton(
                file_item,
                text="❌",
                command=lambda idx=i: self.remove_file(idx),
                width=30,
                height=25,
                font=ctk.CTkFont(size=10)
            )
            remove_button.pack(side="right", padx=(5, 10), pady=5)
        
        # Log da seleção
        if self.selected_files:
            filenames = [Path(f).name for f in self.selected_files]
            self.add_log_message(f"Arquivos selecionados: {', '.join(filenames)}")

    def remove_file(self, index):
        """Remove arquivo da seleção"""
        if 0 <= index < len(self.selected_files):
            removed_file = self.selected_files.pop(index)
            self.update_selected_files_display()
            self.add_log_message(f"Arquivo removido: {Path(removed_file).name}")

    def setup_drag_drop(self):
        """Configura drag and drop de arquivos"""
        if not HAS_DND:
            return
            
        try:
            self.drop_frame.drop_target_register(DND_FILES)
            self.drop_frame.dnd_bind('<<Drop>>', self.handle_drop)
        except:
            pass

    def handle_drop(self, event):
        """Manipula arquivos arrastados"""
        files = self.root.tk.splitlist(event.data)
        pdf_files = [f for f in files if f.lower().endswith('.pdf')]
        
        if pdf_files:
            self.selected_files = pdf_files
            self.update_selected_files_display()
        else:
            messagebox.showerror("Erro", "Por favor, selecione apenas arquivos PDF.")

    def create_config_section(self):
        """Cria seção de configuração do diretório"""
        config_frame = ctk.CTkFrame(self.processing_container)
        config_frame.pack(fill="x", pady=(0, 8))
        
        config_title = ctk.CTkLabel(
            config_frame,
            text="📁 Configuração do Diretório de Trabalho",
            font=ctk.CTkFont(size=15, weight="bold"),
            anchor="w"
        )
        config_title.pack(fill="x", padx=20, pady=(10, 8))
        
        dir_frame = ctk.CTkFrame(config_frame)
        dir_frame.pack(fill="x", padx=20, pady=(0, 8))
        
        help_status_frame = ctk.CTkFrame(dir_frame, fg_color="transparent")
        help_status_frame.pack(fill="x", padx=15, pady=(10, 5))
        
        dir_help = ctk.CTkLabel(
            help_status_frame,
            text="Selecione a pasta que contém o arquivo MODELO.xlsm:",
            font=ctk.CTkFont(size=11),
            text_color=self.colors['text_secondary'],
            anchor="w"
        )
        dir_help.pack(side="left", fill="x", expand=True)
        
        self.config_status = ctk.CTkLabel(
            help_status_frame,
            text="⚠️ Configuração necessária",
            font=ctk.CTkFont(size=10),
            text_color=self.colors['warning'],
            anchor="e"
        )
        self.config_status.pack(side="right", padx=(5, 0))
        
        dir_input_frame = ctk.CTkFrame(dir_frame, fg_color="transparent")
        dir_input_frame.pack(fill="x", padx=15, pady=(0, 10))
        
        self.dir_entry = ctk.CTkEntry(
            dir_input_frame,
            placeholder_text="Caminho para o diretório de trabalho...",
            font=ctk.CTkFont(size=12),
            height=32
        )
        self.dir_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        self.dir_button = ctk.CTkButton(
            dir_input_frame,
            text="📂 Selecionar",
            command=self.select_directory,
            width=110,
            height=32,
            font=ctk.CTkFont(size=12, weight="bold")
        )
        self.dir_button.pack(side="right")

    def create_settings_tab(self):
        """Cria aba de configurações com opções de processamento paralelo"""
        settings_main = ctk.CTkFrame(self.tab_settings)
        settings_main.pack(fill="both", expand=True, padx=10, pady=10)
        
        settings_title = ctk.CTkLabel(
            settings_main,
            text="⚙️ Configurações Avançadas",
            font=ctk.CTkFont(size=18, weight="bold"),
            anchor="w"
        )
        settings_title.pack(fill="x", padx=20, pady=(15, 20))
        
        # Seção de processamento paralelo
        parallel_frame = ctk.CTkFrame(settings_main)
        parallel_frame.pack(fill="x", padx=20, pady=(0, 15))
        
        parallel_title = ctk.CTkLabel(
            parallel_frame,
            text="🚀 Processamento Paralelo",
            font=ctk.CTkFont(size=15, weight="bold"),
            anchor="w"
        )
        parallel_title.pack(fill="x", padx=20, pady=(15, 8))
        
        parallel_desc = ctk.CTkLabel(
            parallel_frame,
            text="Configure quantos PDFs podem ser processados simultaneamente (1 = individual, 2+ = paralelo). Mais threads = mais rápido, mas usa mais memória.",
            font=ctk.CTkFont(size=11),
            text_color=self.colors['text_secondary'],
            anchor="w",
            wraplength=800
        )
        parallel_desc.pack(fill="x", padx=20, pady=(0, 10))
        
        threads_frame = ctk.CTkFrame(parallel_frame, fg_color="transparent")
        threads_frame.pack(fill="x", padx=20, pady=(0, 15))
        
        ctk.CTkLabel(
            threads_frame,
            text="PDFs simultâneos (1=individual, 2+=paralelo):",
            font=ctk.CTkFont(size=12)
        ).pack(side="left")
        
        self.threads_spinbox = ctk.CTkOptionMenu(
            threads_frame,
            values=["1", "2", "3", "4", "5", "6"],
            variable=self.max_threads_var,
            command=self.on_threads_changed,
            width=80
        )
        self.threads_spinbox.pack(side="right", padx=(10, 0))
        
        # Seção de planilha personalizada
        sheet_frame = ctk.CTkFrame(settings_main)
        sheet_frame.pack(fill="x", padx=20, pady=(0, 15))
        
        sheet_title = ctk.CTkLabel(
            sheet_frame,
            text="📊 Nome da Planilha",
            font=ctk.CTkFont(size=15, weight="bold"),
            anchor="w"
        )
        sheet_title.pack(fill="x", padx=20, pady=(15, 8))
        
        sheet_desc = ctk.CTkLabel(
            sheet_frame,
            text="Especifique o nome da planilha a ser atualizada. Deixe vazio para usar 'LEVANTAMENTO DADOS' (padrão).",
            font=ctk.CTkFont(size=11),
            text_color=self.colors['text_secondary'],
            anchor="w",
            wraplength=800
        )
        sheet_desc.pack(fill="x", padx=20, pady=(0, 10))
        
        self.sheet_entry = ctk.CTkEntry(
            sheet_frame,
            placeholder_text="LEVANTAMENTO DADOS",
            font=ctk.CTkFont(size=12),
            height=35
        )
        self.sheet_entry.pack(fill="x", padx=20, pady=(0, 15))
        self.sheet_entry.bind('<KeyRelease>', lambda e: self.root.after(1000, self.save_current_config))
        
        # Seção de modo verboso
        verbose_frame = ctk.CTkFrame(settings_main)
        verbose_frame.pack(fill="x", padx=20, pady=(0, 15))
        
        verbose_title = ctk.CTkLabel(
            verbose_frame,
            text="🔍 Logs Detalhados",
            font=ctk.CTkFont(size=15, weight="bold"),
            anchor="w"
        )
        verbose_title.pack(fill="x", padx=20, pady=(15, 8))
        
        self.verbose_checkbox = ctk.CTkCheckBox(
            verbose_frame,
            text="Habilitar modo verboso (logs detalhados)",
            variable=self.verbose_var,
            command=self.save_current_config,
            font=ctk.CTkFont(size=12)
        )
        self.verbose_checkbox.pack(padx=20, pady=(0, 15), anchor="w")

    def on_threads_changed(self, value):
        """Callback quando número de threads muda"""
        self.max_threads_var.set(int(value))
        self.batch_processor.max_workers = int(value)
        self.save_current_config()
        if int(value) == 1:
            self.add_log_message("Processamento configurado para: Individual (1 PDF por vez)")
        else:
            self.add_log_message(f"Processamento configurado para: Paralelo ({value} PDFs simultâneos)")

    def create_history_tab(self):
        """Cria aba de histórico"""
        history_main = ctk.CTkFrame(self.tab_history)
        history_main.pack(fill="both", expand=True, padx=10, pady=10)
        
        history_title = ctk.CTkLabel(
            history_main,
            text="📊 Histórico de Processamentos (Individual)",
            font=ctk.CTkFont(size=18, weight="bold"),
            anchor="w"
        )
        history_title.pack(fill="x", padx=20, pady=(15, 10))
        
        controls_frame = ctk.CTkFrame(history_main)
        controls_frame.pack(fill="x", padx=20, pady=(0, 10))
        
        self.clear_history_button = ctk.CTkButton(
            controls_frame,
            text="🗑️ Limpar Histórico",
            command=self.clear_history,
            width=150,
            height=35,
            font=ctk.CTkFont(size=12),
            fg_color=self.colors['secondary']
        )
        self.clear_history_button.pack(side="right", padx=15, pady=10)
        
        self.history_status_label = ctk.CTkLabel(
            controls_frame,
            text="Nenhum PDF no histórico",
            font=ctk.CTkFont(size=12),
            text_color=self.colors['text_secondary']
        )
        self.history_status_label.pack(side="left", padx=15, pady=10)
        
        self.history_list_frame = ctk.CTkScrollableFrame(history_main, height=400)
        self.history_list_frame.pack(fill="both", expand=True, padx=20, pady=(0, 15))

    def select_directory(self):
        """Seleciona diretório de trabalho"""
        directory = filedialog.askdirectory(
            title="Selecione o diretório de trabalho (que contém MODELO.xlsm)"
        )
        
        if directory:
            self.dir_entry.delete(0, 'end')
            self.dir_entry.insert(0, directory)
            self.validate_config()

    def validate_config(self):
        """Valida configuração do diretório"""
        directory = self.dir_entry.get().strip()
        
        if not directory:
            self.config_status.configure(
                text="⚠️ Selecione um diretório",
                text_color=self.colors['warning']
            )
            return
        
        try:
            processor_factory = self._get_processor()
            processor = processor_factory()
            processor.set_trabalho_dir(directory)
            valid, message = processor.validate_trabalho_dir()
            
            if valid:
                self.config_status.configure(
                    text="✅ Configuração válida",
                    text_color=self.colors['success']
                )
                self.trabalho_dir = directory
                self.update_pdf_list()
                self.save_current_config()
            else:
                self.config_status.configure(
                    text=f"❌ {message}",
                    text_color=self.colors['error']
                )
        except Exception as e:
            self.config_status.configure(
                text=f"❌ Erro: {str(e)}",
                text_color=self.colors['error']
            )

    def update_pdf_list(self):
        """Atualiza lista de PDFs disponíveis"""
        if not self.trabalho_dir:
            return

        try:
            processor_factory = self._get_processor()
            processor = processor_factory()
            processor.set_trabalho_dir(self.trabalho_dir)
            pdf_files = processor.get_pdf_files_in_trabalho_dir()
            
            if pdf_files:
                self.add_log_message(f"PDFs encontrados no diretório: {', '.join(pdf_files)}")
            else:
                self.add_log_message("Nenhum arquivo PDF encontrado no diretório.")
        except Exception as e:
            self.add_log_message(f"Erro ao listar PDFs: {e}")

    def process_pdfs(self):
        """Processa os PDFs selecionados"""
        if self.processing:
            return
        
        # Validações
        if not self.trabalho_dir:
            messagebox.showerror("Erro", "Configure o diretório de trabalho primeiro.")
            return
        
        if not self.selected_files:
            messagebox.showerror("Erro", "Selecione pelo menos um arquivo PDF.")
            return
        
        # Configura processamento
        self.processing = True
        self.process_button.configure(text="🔄 Processando...", state="disabled")
        
        # Mostra popup apropriado
        if len(self.selected_files) == 1:
            # Processamento individual (mantém compatibilidade)
            self._process_single_pdf()
        else:
            # Processamento em lote
            self._process_batch_pdfs()

    def _process_single_pdf(self):
        """Processa um único PDF (modo compatibilidade)"""
        pdf_file = self.selected_files[0]
        
        def single_progress(progress, message=""):
            self.add_log_message(f"Progresso: {progress}% - {message}")
        
        def single_log(message):
            self.add_log_message(message)
        
        def process_thread():
            try:
                processor_factory = self._get_processor()
                processor = processor_factory()
                processor.set_trabalho_dir(self.trabalho_dir)
                processor.progress_callback = single_progress
                processor.log_callback = single_log
                
                filename = Path(pdf_file).name
                results = processor.process_pdf(filename)
                
                self.root.after(0, self._single_process_complete, results, filename)
                
            except Exception as e:
                self.root.after(0, self._process_error, str(e))
        
        threading.Thread(target=process_thread, daemon=True).start()

    def _process_batch_pdfs(self):
        """Processa múltiplos PDFs em paralelo"""
        # Cria popup de processamento em lote
        self.batch_popup = BatchProcessingPopup(self)
        self.batch_popup.show(self.selected_files)
        
        # Inicia processamento em lote
        self.batch_processor.process_batch(
            self.selected_files,
            self._get_processor(),
            self.trabalho_dir
        )
        
        # Inicia monitoramento do status
        self._monitor_batch_progress()

    def _monitor_batch_progress(self):
        """
        Monitora progresso do processamento em lote de forma mais responsiva.
        Versão 3.3.1 - Melhorada para evitar congelamento da interface.
        """
        if not self.batch_processor.is_running and self.batch_processor.status_queue.empty():
            return
        
        # Processa até 5 atualizações por ciclo para manter responsividade
        updates_processed = 0
        max_updates_per_cycle = 5
        
        try:
            while updates_processed < max_updates_per_cycle:
                try:
                    # Timeout curto para não bloquear interface
                    update_type, filename, status = self.batch_processor.status_queue.get_nowait()
                    updates_processed += 1

                    if update_type == 'update' and self.batch_popup and self.batch_popup.window:
                        # Atualiza status individual
                        self.batch_popup.update_pdf_status(filename, status)

                        # Atualiza progresso geral
                        completed = self.batch_processor.completed_count
                        total = self.batch_processor.total_count
                        self.batch_popup.update_general_progress(completed, total)

                        # Log do progresso
                        if status.status == 'completed':
                            self.add_log_message(f"✅ {filename}: Concluído")
                        elif status.status == 'error':
                            self.add_log_message(f"❌ {filename}: {status.message}")
                        elif status.status == 'processing' and status.progress > 0:
                            # Log menos verboso para progresso
                            if status.progress % 25 == 0:  # Log a cada 25%
                                self.add_log_message(f"🔄 {filename}: {status.progress}% - {status.message}")

                    elif update_type == 'batch_complete':
                        # Processamento completo
                        self._batch_process_complete()
                        return

                except queue.Empty:
                    break

            # Força atualização da interface mais frequentemente
            if self.batch_popup and self.batch_popup.window:
                try:
                    self.batch_popup.window.update_idletasks()
                except Exception:
                    pass
            
            # Atualiza interface principal também
            try:
                self.root.update_idletasks()
            except Exception:
                pass

            # Continua monitoramento se ainda processando com intervalo menor
            if self.batch_processor.is_running or not self.batch_processor.status_queue.empty():
                # Intervalo reduzido para maior responsividade
                self.root.after(50, self._monitor_batch_progress)

        except Exception as e:
            self.add_log_message(f"Erro no monitoramento: {e}")
            # Em caso de erro, tenta continuar monitoramento
            if self.batch_processor.is_running:
                self.root.after(200, self._monitor_batch_progress)

    def _batch_process_complete(self):
        """Callback quando processamento em lote completa"""
        self.processing = False
        self.process_button.configure(text="🚀 Processar PDFs", state="normal")
        
        # Coleta resultados
        successful_pdfs = []
        failed_pdfs = []
        batch_timestamp = datetime.now()
        
        # Cria entradas INDIVIDUAIS para cada PDF processado
        for filename, status in self.batch_processor.pdf_statuses.items():
            if status.status == 'completed':
                successful_pdfs.append((filename, status.result_data))
                
                # Cria entrada individual para PDF bem-sucedido
                individual_entry = HistoryEntry(
                    timestamp=batch_timestamp,
                    pdf_file=filename,
                    success=True,
                    result_data=status.result_data,
                    logs=status.logs[-20:],  # Últimos 20 logs do PDF
                    is_batch=True,  # Indica que foi processado em lote
                    batch_info={
                        'batch_size': len(self.selected_files),
                        'processed_in_batch': True
                    }
                )
                
                self.processing_history.append(individual_entry)
                self.persistence.save_history_entry(individual_entry)
                
            else:
                failed_pdfs.append((filename, status.message))
                
                # Cria entrada individual para PDF que falhou
                individual_entry = HistoryEntry(
                    timestamp=batch_timestamp,
                    pdf_file=filename,
                    success=False,
                    result_data={'error': status.message},
                    logs=status.logs[-20:],  # Últimos 20 logs do PDF
                    is_batch=True,  # Indica que foi processado em lote
                    batch_info={
                        'batch_size': len(self.selected_files),
                        'processed_in_batch': True
                    }
                )
                
                self.processing_history.append(individual_entry)
                self.persistence.save_history_entry(individual_entry)
        
        self.update_history_display()
        
        # Fecha popup
        if self.batch_popup:
            self.batch_popup.close()
            self.batch_popup = None
        
        # Mostra resultado final
        total = len(self.selected_files)
        success_count = len(successful_pdfs)
        fail_count = len(failed_pdfs)
        
        if success_count == total:
            successful_names = [item[0] for item in successful_pdfs]
            messagebox.showinfo(
                "Processamento Concluído",
                f"✅ Todos os {total} PDFs foram processados com sucesso!\n\n"
                f"Arquivos processados:\n" + "\n".join([f"• {f}" for f in successful_names[:5]])
                + (f"\n... e mais {len(successful_names)-5}" if len(successful_names) > 5 else "")
                + f"\n\n📊 Cada arquivo foi adicionado individualmente ao histórico."
            )
        elif success_count > 0:
            messagebox.showwarning(
                "Processamento Parcial",
                f"⚠️ {success_count}/{total} PDFs processados com sucesso.\n\n"
                f"Sucessos: {success_count}\nFalhas: {fail_count}\n\n"
                f"📊 Todos os arquivos foram adicionados individualmente ao histórico."
            )
        else:
            messagebox.showerror(
                "Processamento Falhou",
                f"❌ Nenhum PDF foi processado com sucesso.\n\n"
                f"Total de falhas: {fail_count}\n\n"
                f"📊 Todas as falhas foram registradas individualmente no histórico."
            )
        
        # Limpa seleção e vai para histórico
        self.clear_selection()
        self.tabview.set("📊 Histórico")

    def _single_process_complete(self, results, filename):
        """Callback para processamento individual"""
        self.processing = False
        self.process_button.configure(text="🚀 Processar PDFs", state="normal")
        
        # Adiciona ao histórico como entrada individual
        entry = HistoryEntry(
            timestamp=datetime.now(),
            pdf_file=filename,
            success=results['success'],
            result_data=results,
            logs=self.current_logs.copy(),
            is_batch=False,  # Processamento individual
            batch_info=None
        )
        
        self.processing_history.append(entry)
        self.persistence.save_history_entry(entry)
        self.update_history_display()
        
        # Mostra resultado
        if results['success']:
            messagebox.showinfo(
                "Sucesso",
                f"✅ PDF processado com sucesso!\n\n"
                f"Períodos processados: {results['total_extracted']}\n"
                f"Arquivo: {results['arquivo_final']}\n\n"
                f"📊 Adicionado individualmente ao histórico."
            )
        else:
            messagebox.showerror("Erro", f"❌ Erro no processamento:\n\n{results['error']}")
        
        # Limpa e vai para histórico
        self.clear_selection()
        self.tabview.set("📊 Histórico")

    def _process_error(self, error_message):
        """Callback para erro de processamento"""
        self.processing = False
        self.process_button.configure(text="🚀 Processar PDFs", state="normal")
        
        messagebox.showerror("Erro", f"❌ Erro no processamento:\n\n{error_message}")

    def clear_history(self):
        """Limpa histórico"""
        if not self.processing_history:
            return
        
        if messagebox.askyesno("Confirmar", "Deseja limpar todo o histórico?"):
            self.processing_history.clear()
            self.persistence.clear_history()
            self.update_history_display()
            self.history_status_label.configure(text="Nenhum PDF no histórico")

    def update_history_display(self):
        """Atualiza exibição do histórico"""
        # Limpa lista atual
        for widget in self.history_list_frame.winfo_children():
            widget.destroy()
        
        # Adiciona entradas do histórico (mais recentes primeiro)
        for i, entry in enumerate(reversed(self.processing_history)):
            self.create_history_item(entry, len(self.processing_history) - 1 - i)
        
        # Atualiza status do histórico
        if hasattr(self, 'history_status_label'):
            total = len(self.processing_history)
            success_count = sum(1 for h in self.processing_history if h.success)
            batch_count = sum(1 for h in self.processing_history if h.is_batch)
            individual_count = total - batch_count
            
            if total > 0:
                status_text = f"{total} PDFs no histórico ({success_count} sucessos, {total - success_count} falhas)"
                if batch_count > 0:
                    status_text += f" • {batch_count} de lotes, {individual_count} individuais"
                self.history_status_label.configure(text=status_text)
            else:
                self.history_status_label.configure(text="Nenhum PDF no histórico")

    def create_history_item(self, entry, index):
        """Cria item visual no histórico"""
        item_frame = ctk.CTkFrame(self.history_list_frame)
        item_frame.pack(fill="x", padx=5, pady=2)
        
        inner_frame = ctk.CTkFrame(item_frame, fg_color="transparent")
        inner_frame.pack(fill="x", padx=10, pady=8)
        
        # Ícone de status com indicador de lote
        if entry.is_batch and entry.batch_info and entry.batch_info.get('processed_in_batch'):
            # PDF processado em lote
            if entry.success:
                icon = "📦✅"  # Ícone de lote + sucesso
            else:
                icon = "📦❌"  # Ícone de lote + erro
        else:
            # PDF processado individualmente
            icon = "✅" if entry.success else "❌"
        
        status_color = self.colors['success'] if entry.success else self.colors['error']
        
        status_label = ctk.CTkLabel(
            inner_frame,
            text=icon,
            font=ctk.CTkFont(size=16),
            text_color=status_color,
            width=40
        )
        status_label.pack(side="left", padx=(0, 10))
        
        # Informações principais
        info_frame = ctk.CTkFrame(inner_frame, fg_color="transparent")
        info_frame.pack(side="left", fill="x", expand=True)
        
        # Linha 1: Nome do arquivo e timestamp
        line1_frame = ctk.CTkFrame(info_frame, fg_color="transparent")
        line1_frame.pack(fill="x")
        
        # Mostra nome do arquivo individual (sem extensão se for sucesso)
        if entry.success and entry.result_data.get('arquivo_final'):
            # Remove extensão do nome do arquivo final
            display_name = f"📄 {Path(entry.result_data['arquivo_final']).stem}"
        else:
            # Usa nome do PDF original
            display_name = f"📄 {Path(entry.pdf_file).stem}"
        
        # Adiciona indicador de lote se aplicável
        if entry.is_batch and entry.batch_info and entry.batch_info.get('processed_in_batch'):
            batch_size = entry.batch_info.get('batch_size', 0)
            if batch_size > 1:
                display_name += f" (lote {batch_size})"
        
        file_label = ctk.CTkLabel(
            line1_frame,
            text=display_name,
            font=ctk.CTkFont(size=12, weight="bold"),
            anchor="w"
        )
        file_label.pack(side="left")
        
        time_label = ctk.CTkLabel(
            line1_frame,
            text=entry.timestamp.strftime("%d/%m/%Y %H:%M:%S"),
            font=ctk.CTkFont(size=10),
            text_color=self.colors['text_secondary'],
            anchor="e"
        )
        time_label.pack(side="right")
        
        # Linha 2: Resultado individual
        if entry.success:
            result_text = f"✓ {entry.result_data.get('total_extracted', 0)} períodos processados"
            if entry.result_data.get('person_name'):
                result_text += f" • {entry.result_data['person_name']}"
        else:
            result_text = f"✗ {entry.result_data.get('error', 'Erro desconhecido')}"
        
        result_label = ctk.CTkLabel(
            info_frame,
            text=result_text,
            font=ctk.CTkFont(size=10),
            text_color=status_color,
            anchor="w"
        )
        result_label.pack(fill="x", pady=(2, 0))
        
        # Botões de ação
        actions_frame = ctk.CTkFrame(inner_frame, fg_color="transparent")
        actions_frame.pack(side="right", padx=(10, 0))
        
        # Botão "Abrir Dados" (apenas para sucessos)
        if entry.success:
            open_data_button = ctk.CTkButton(
                actions_frame,
                text="📂",
                command=lambda e=entry: self.open_data_file(e),
                width=35,
                height=25,
                font=ctk.CTkFont(size=12),
                fg_color=self.colors['primary'],
                hover_color="#174a80"
            )
            open_data_button.pack(side="right", padx=(5, 0))
        
        # Botão "Ver Detalhes" - CORRIGIDO v3.3.1
        details_button = ctk.CTkButton(
            actions_frame,
            text="📝",
            command=lambda e=entry: self.show_history_details(e),
            width=35,
            height=25,
            font=ctk.CTkFont(size=12),
            fg_color=self.colors['secondary'],
            hover_color="#0f2a3f"
        )
        details_button.pack(side="right")
        
        # Hover effect no item
        def on_enter(event):
            item_frame.configure(fg_color=self.colors['bg_light'])
        
        def on_leave(event):
            item_frame.configure(fg_color=["gray92", "gray14"])

        item_frame.bind("<Enter>", on_enter)
        item_frame.bind("<Leave>", on_leave)

    def open_data_file(self, entry):
        """Abre o arquivo Excel associado à entrada do histórico"""
        try:
            file_path = entry.result_data.get('excel_path')
            if not file_path and self.trabalho_dir:
                arquivo_rel = entry.result_data.get('arquivo_final')
                if arquivo_rel:
                    file_path = Path(self.trabalho_dir) / arquivo_rel

            if not file_path:
                messagebox.showerror("Erro", "Caminho do arquivo não disponível.")
                return

            file_path = Path(file_path)
            if not file_path.exists():
                messagebox.showerror("Erro", f"Arquivo não encontrado: {file_path}")
                return

            if sys.platform.startswith('win'):
                os.startfile(file_path)  # type: ignore[attr-defined]
            elif sys.platform == 'darwin':
                subprocess.Popen(['open', str(file_path)])
            else:
                subprocess.Popen(['xdg-open', str(file_path)])

        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir o arquivo: {e}")

    def show_history_details(self, entry):
        """
        Mostra detalhes da entrada do histórico.
        Versão 3.3.1 - CORRIGIDO: Agora usa janela modal com controle de instância única.
        """
        try:
            # Cria janela de detalhes com estilo melhorado e controle modal
            details_window = HistoryDetailsWindow(self, entry)
        except Exception as e:
            # Fallback para messagebox em caso de erro
            messagebox.showerror("Erro", f"Não foi possível abrir janela de detalhes: {e}")

    def add_log_message(self, message):
        """Adiciona mensagem ao log"""
        try:
            timestamp = datetime.now().strftime("%H:%M:%S")
            log_entry = f"[{timestamp}] {message}"
            self.current_logs.append(log_entry)
        except:
            pass

    def load_initial_config(self):
        """Carrega configuração inicial"""
        def task():
            try:
                self.root.after(0, lambda: self.add_log_message("Sistema iniciado - v3.3.2 com interface otimizada e lista expandida"))
                
                # Configura status inicial
                if not self.trabalho_dir:
                    self.root.after(0, lambda: self.config_status.configure(
                        text="⚙️ Configure o diretório de trabalho",
                        text_color=self.colors['warning']
                    ))
                
            except Exception as e:
                error_message = f"Erro na inicialização: {str(e)}"
                self.root.after(0, lambda: self.add_log_message(error_message))

        threading.Thread(target=task, daemon=True).start()

    def on_closing(self):
        """Manipula fechamento da aplicação"""
        if self.processing:
            result = messagebox.askyesno(
                "Processamento em andamento",
                "Há processamentos em andamento. Deseja realmente fechar?"
            )
            if not result:
                return
        
        # Fecha janela de detalhes se estiver aberta
        if HistoryDetailsWindow._instance:
            try:
                HistoryDetailsWindow._instance.close()
            except:
                pass
        
        # Fecha popup de processamento se estiver aberto
        if self.batch_popup:
            try:
                self.batch_popup.close()
            except:
                pass
        
        self.save_current_config()
        self.root.destroy()

    def run(self):
        """Inicia a aplicação"""
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.dir_entry.bind('<KeyRelease>', lambda e: self.root.after(500, self.validate_config))
        self.root.mainloop()


def main():
    """Função principal"""
    try:
        import customtkinter
        
        if not HAS_DND:
            print("AVISO: tkinterdnd2 não instalado. Drag & drop desabilitado.")
        
        app = PDFExcelDesktopApp()
        app.run()
        
    except ImportError as e:
        error_msg = f"""
ERRO: Dependências não instaladas!

Para instalar: pip install customtkinter pillow
Opcional: pip install tkinterdnd2

Dependências faltando: {str(e)}
        """
        print(error_msg)
        sys.exit(1)
    
    except Exception as e:
        print(f"Erro ao iniciar aplicação: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()