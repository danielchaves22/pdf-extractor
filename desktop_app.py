#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF para Excel Desktop App - Interface Gráfica
==============================================

Interface gráfica moderna usando CustomTkinter que utiliza o módulo
pdf_processor_core.py para toda a lógica de processamento.

Versão 3.2 - Com Sistema de Histórico e Abas Otimizado

Dependências:
pip install customtkinter pillow

Como usar:
python desktop_app.py

Compatibilidade:
- Windows (gera .exe com PyInstaller)
- macOS (gera .app com PyInstaller)  
- Linux (gera executável com PyInstaller)

Autor: Sistema de Extração Automatizada
Data: 2025
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import os
import sys
from pathlib import Path
from datetime import datetime
import webbrowser
from tkinterdnd2 import DND_FILES, TkinterDnD

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

class HistoryEntry:
    """Representa um entrada no histórico de processamentos"""
    def __init__(self, timestamp, pdf_file, success, result_data, logs):
        self.timestamp = timestamp
        self.pdf_file = pdf_file
        self.success = success
        self.result_data = result_data
        self.logs = logs

class ProcessingPopup:
    """Popup para mostrar progresso e logs durante processamento"""
    def __init__(self, parent):
        self.parent = parent
        self.window = None
        
    def show(self):
        """Mostra o popup de processamento"""
        self.window = ctk.CTkToplevel(self.parent.root)
        self.window.title("Processando PDF...")
        
        # Remove qualquer efeito de transparência ou grab
        self.window.attributes('-alpha', 1.0)  # Força opacidade total
        
        # Configurações básicas primeiro
        self.window.transient(self.parent.root)
        self.window.protocol("WM_DELETE_WINDOW", self._on_close_attempt)
        
        # Cria interface primeiro
        self._create_interface()
        
        # Posiciona DEPOIS de criar a interface, com delay para garantir renderização
        self.window.after(50, lambda: self._position_relative_to_parent(600, 400))
        
        # Força foco sem grab
        self.window.focus_force()
        self.window.lift()
        self.window.attributes('-topmost', True)
        self.window.after(100, lambda: self.window.attributes('-topmost', False))

    def _create_interface(self):
        """Cria a interface do popup separadamente"""
    def _create_interface(self):
        """Cria a interface do popup separadamente"""
        # Header
        header_frame = ctk.CTkFrame(self.window)
        header_frame.pack(fill="x", padx=20, pady=(20, 10))
        
        title_label = ctk.CTkLabel(
            header_frame,
            text="🔄 Processando PDF...",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        title_label.pack(pady=15)
        
        # Progress section
        progress_frame = ctk.CTkFrame(self.window)
        progress_frame.pack(fill="x", padx=20, pady=(0, 10))
        
        # Barra de progresso
        self.progress_bar = ctk.CTkProgressBar(progress_frame, height=20)
        self.progress_bar.pack(fill="x", padx=20, pady=(20, 10))
        self.progress_bar.set(0)
        
        # Label de status
        self.progress_label = ctk.CTkLabel(
            progress_frame,
            text="Iniciando processamento...",
            font=ctk.CTkFont(size=12)
        )
        self.progress_label.pack(padx=20, pady=(0, 20))
        
        # Logs section
        logs_frame = ctk.CTkFrame(self.window)
        logs_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        
        logs_title = ctk.CTkLabel(
            logs_frame,
            text="📝 Log de Processamento",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        logs_title.pack(pady=(15, 10))
        
        # Textbox para logs
        self.log_textbox = ctk.CTkTextbox(
            logs_frame,
            font=ctk.CTkFont(family="Consolas", size=10),
            wrap="word"
        )
        self.log_textbox.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        
        # Conecta callbacks do parent a este popup
        self.parent.popup_progress_bar = self.progress_bar
        self.parent.popup_progress_label = self.progress_label
        self.parent.popup_log_textbox = self.log_textbox
    
    def _position_relative_to_parent(self, width, height):
        """Posiciona o popup relativamente à janela principal de forma simples"""
        try:
            # Método simples usando apenas winfo_rootx/rooty
            parent_x = self.parent.root.winfo_rootx()
            parent_y = self.parent.root.winfo_rooty()
            parent_width = self.parent.root.winfo_width()
            parent_height = self.parent.root.winfo_height()
            
            # Calcula posição centralizada
            x = parent_x + (parent_width - width) // 2
            y = parent_y + (parent_height - height) // 2
            
            # Aplica posição diretamente
            self.window.geometry(f"{width}x{height}+{x}+{y}")
            
        except Exception:
            # Fallback simples: posição fixa
            self.window.geometry(f"{width}x{height}+300+200")
    
    def _on_close_attempt(self):
        """Impede fechamento durante processamento"""
        if self.parent.processing:
            messagebox.showwarning(
                "Processamento em andamento",
                "Aguarde a conclusão do processamento."
            )
        else:
            self.close()
    
    def close(self):
        """Fecha o popup sem efeitos de transparência"""
        if self.window:
            # Desconecta callbacks primeiro
            self.parent.popup_progress_bar = None
            self.parent.popup_progress_label = None
            self.parent.popup_log_textbox = None
            
            # Destrói janela diretamente sem grab_release
            self.window.destroy()
            self.window = None
            
            # Força foco na janela principal sem grab/release
            self.parent.root.focus_force()
            self.parent.root.lift()
            self.parent.root.attributes('-alpha', 1.0)  # Força opacidade total

class PDFExcelDesktopApp:
    def __init__(self):
        # Configuração da janela principal simples
        self.root = TkinterDnD.Tk()
        self.root.title("Processamento de Folha de Pagamento v3.2")
        self.root.geometry("950x600")
        self.root.resizable(False, False)
        
        # Remove qualquer transparência da janela principal
        self.root.attributes('-alpha', 1.0)
        
        # Configuração de ícone (opcional)
        try:
            # Se tiver um ícone .ico, descomente abaixo
            # self.root.iconbitmap("icon.ico")
            pass
        except:
            pass
        
        # Variáveis de estado
        self.selected_file = None
        self.trabalho_dir = None
        self.processing = False
        
        # Histórico de processamentos
        self.processing_history = []
        self.current_logs = []
        
        # Variáveis da interface (inicializadas antes das abas)
        self.verbose_var = ctk.BooleanVar()
        self.sheet_entry = None  # Será criado na aba de configurações
        
        # Popup de processamento
        self.processing_popup = ProcessingPopup(self)
        self.popup_progress_bar = None
        self.popup_progress_label = None
        self.popup_log_textbox = None

        # Processador será inicializado sob demanda
        self.processor = None
        
        # Debug para posicionamento de popups (pode ser ativado se necessário)
        # self._debug_positioning = True  # Descomente para ativar logs de debug
        
        # Para ativar debug de posicionamento temporariamente, descomente:
        # print("Para ativar debug de posicionamento, adicione: app._debug_positioning = True")

        # Configuração de estilo
        self.setup_styles()
        
        # Cria a interface
        self.create_interface()
        
        # Configura drag and drop
        self.setup_drag_drop()
        
        # Carrega configurações
        self.load_initial_config()

    def _get_processor(self):
        """Inicializa o processador core quando necessário"""
        if self.processor is None:
            self.processor = PDFProcessorCore(
                progress_callback=self.update_progress,
                log_callback=self.add_log_message
            )
        return self.processor

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
        
        # Cria as abas na ordem correta: Processamento, Histórico, Configurações
        self.tab_processing = self.tabview.add("📄 Processamento")
        self.tab_history = self.tabview.add("📊 Histórico")
        self.tab_settings = self.tabview.add("⚙️ Configurações Avançadas")
        
        # Configura aba de processamento
        self.create_processing_tab()
        
        # Configura aba de histórico
        self.create_history_tab()
        
        # Configura aba de configurações
        self.create_settings_tab()
        
        # Define aba inicial
        self.tabview.set("📄 Processamento")

    def create_global_header(self, parent):
        """Cria o cabeçalho global da aplicação"""
        header_frame = ctk.CTkFrame(parent)
        header_frame.pack(fill="x", pady=(0, 8))
        
        # Título principal
        title_label = ctk.CTkLabel(
            header_frame,
            text="📄 Processamento de Folha de Pagamento",
            font=ctk.CTkFont(size=22, weight="bold")
        )
        title_label.pack(pady=(10, 6))
        
        # Subtítulo
        subtitle_label = ctk.CTkLabel(
            header_frame,
            text="Automatização de folhas de pagamento PDF para Excel v3.2",
            font=ctk.CTkFont(size=11),
            text_color=self.colors['text_secondary']
        )
        subtitle_label.pack(pady=(0, 10))

    def create_processing_tab(self):
        """Cria o conteúdo da aba de processamento"""
        # Container normal (sem scroll) para manter tudo visível
        self.processing_container = ctk.CTkFrame(self.tab_processing)
        self.processing_container.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Configuração do diretório de trabalho
        self.create_config_section()
        
        # Seleção de arquivo (compacta)
        self.create_compact_file_section()
        
        # Botão de ação principal (logo após seleção de arquivo)
        self.create_main_action_button()

    def create_history_tab(self):
        """Cria o conteúdo da aba de histórico"""
        # Container principal do histórico
        history_main = ctk.CTkFrame(self.tab_history)
        history_main.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Título da seção
        history_title = ctk.CTkLabel(
            history_main,
            text="📊 Histórico de Processamentos da Sessão",
            font=ctk.CTkFont(size=18, weight="bold"),
            anchor="w"
        )
        history_title.pack(fill="x", padx=20, pady=(15, 10))
        
        # Frame para controles
        controls_frame = ctk.CTkFrame(history_main)
        controls_frame.pack(fill="x", padx=20, pady=(0, 10))
        
        # Botão limpar histórico
        self.clear_history_button = ctk.CTkButton(
            controls_frame,
            text="🗑️ Limpar Histórico",
            command=self.clear_history,
            width=150,
            height=35,
            font=ctk.CTkFont(size=12),
            fg_color=self.colors['secondary'],
            hover_color="#0f2a3f"
        )
        self.clear_history_button.pack(side="right", padx=15, pady=10)
        
        # Label de status
        self.history_status_label = ctk.CTkLabel(
            controls_frame,
            text="Nenhum processamento realizado nesta sessão",
            font=ctk.CTkFont(size=12),
            text_color=self.colors['text_secondary']
        )
        self.history_status_label.pack(side="left", padx=15, pady=10)
        
        # Container da lista de histórico
        self.history_list_frame = ctk.CTkScrollableFrame(history_main, height=400)
        self.history_list_frame.pack(fill="both", expand=True, padx=20, pady=(0, 15))

    def create_settings_tab(self):
        """Cria o conteúdo da aba de configurações"""
        # Container principal das configurações
        settings_main = ctk.CTkFrame(self.tab_settings)
        settings_main.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Título da seção
        settings_title = ctk.CTkLabel(
            settings_main,
            text="⚙️ Configurações Avançadas",
            font=ctk.CTkFont(size=18, weight="bold"),
            anchor="w"
        )
        settings_title.pack(fill="x", padx=20, pady=(15, 20))
        
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
        
        verbose_desc = ctk.CTkLabel(
            verbose_frame,
            text="Ativa logs detalhados durante o processamento para diagnóstico e resolução de problemas.",
            font=ctk.CTkFont(size=11),
            text_color=self.colors['text_secondary'],
            anchor="w",
            wraplength=800
        )
        verbose_desc.pack(fill="x", padx=20, pady=(0, 10))
        
        self.verbose_checkbox = ctk.CTkCheckBox(
            verbose_frame,
            text="Habilitar modo verboso (logs detalhados)",
            variable=self.verbose_var,
            font=ctk.CTkFont(size=12)
        )
        self.verbose_checkbox.pack(padx=20, pady=(0, 15), anchor="w")
        
        # Seção de reset
        reset_frame = ctk.CTkFrame(settings_main)
        reset_frame.pack(fill="x", padx=20, pady=(0, 15))
        
        reset_title = ctk.CTkLabel(
            reset_frame,
            text="🔄 Redefinir Configurações",
            font=ctk.CTkFont(size=15, weight="bold"),
            anchor="w"
        )
        reset_title.pack(fill="x", padx=20, pady=(15, 8))
        
        reset_desc = ctk.CTkLabel(
            reset_frame,
            text="Restaura todas as configurações para os valores padrão.",
            font=ctk.CTkFont(size=11),
            text_color=self.colors['text_secondary'],
            anchor="w"
        )
        reset_desc.pack(fill="x", padx=20, pady=(0, 10))
        
        reset_button = ctk.CTkButton(
            reset_frame,
            text="🔄 Redefinir Configurações",
            command=self.reset_settings,
            width=200,
            height=35,
            font=ctk.CTkFont(size=12),
            fg_color=self.colors['secondary'],
            hover_color="#0f2a3f"
        )
        reset_button.pack(padx=20, pady=(0, 15), anchor="w")

    def reset_settings(self):
        """Redefine todas as configurações para valores padrão"""
        # Limpa campo de planilha (se existir)
        if self.sheet_entry:
            self.sheet_entry.delete(0, 'end')
        
        # Desabilita modo verboso
        self.verbose_var.set(False)
        
        # Mensagem de confirmação
        messagebox.showinfo("Configurações", "Configurações redefinidas para valores padrão.")

    def create_config_section(self):
        """Cria seção de configuração"""
        config_frame = ctk.CTkFrame(self.processing_container)
        config_frame.pack(fill="x", pady=(0, 8))
        
        # Título da seção
        config_title = ctk.CTkLabel(
            config_frame,
            text="📁 Configuração do Diretório de Trabalho",
            font=ctk.CTkFont(size=15, weight="bold"),
            anchor="w"
        )
        config_title.pack(fill="x", padx=20, pady=(10, 8))
        
        # Frame para seleção de diretório
        dir_frame = ctk.CTkFrame(config_frame)
        dir_frame.pack(fill="x", padx=20, pady=(0, 8))
        
        # Frame para label explicativo e status (horizontal)
        help_status_frame = ctk.CTkFrame(dir_frame, fg_color="transparent")
        help_status_frame.pack(fill="x", padx=15, pady=(10, 5))
        
        # Label explicativo (lado esquerdo)
        dir_help = ctk.CTkLabel(
            help_status_frame,
            text="Selecione a pasta que contém o arquivo MODELO.xlsm:",
            font=ctk.CTkFont(size=11),
            text_color=self.colors['text_secondary'],
            anchor="w"
        )
        dir_help.pack(side="left", fill="x", expand=True)
        
        # Status da configuração (lado direito)
        self.config_status = ctk.CTkLabel(
            help_status_frame,
            text="⚠️ Configuração necessária",
            font=ctk.CTkFont(size=10),
            text_color=self.colors['warning'],
            anchor="e"
        )
        self.config_status.pack(side="right", padx=(5, 0))
        
        # Frame horizontal para entrada e botão
        dir_input_frame = ctk.CTkFrame(dir_frame, fg_color="transparent")
        dir_input_frame.pack(fill="x", padx=15, pady=(0, 10))
        
        # Entrada de texto para diretório
        self.dir_entry = ctk.CTkEntry(
            dir_input_frame,
            placeholder_text="Caminho para o diretório de trabalho...",
            font=ctk.CTkFont(size=12),
            height=32
        )
        self.dir_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        # Botão para selecionar diretório
        self.dir_button = ctk.CTkButton(
            dir_input_frame,
            text="📂 Selecionar",
            command=self.select_directory,
            width=110,
            height=32,
            font=ctk.CTkFont(size=12, weight="bold")
        )
        self.dir_button.pack(side="right")

    def create_compact_file_section(self):
        """Cria seção compacta de seleção de arquivo"""
        file_frame = ctk.CTkFrame(self.processing_container)
        file_frame.pack(fill="x", pady=(0, 8))
        
        # Título da seção
        file_title = ctk.CTkLabel(
            file_frame,
            text="📎 Seleção de Arquivo PDF",
            font=ctk.CTkFont(size=15, weight="bold"),
            anchor="w"
        )
        file_title.pack(fill="x", padx=20, pady=(10, 8))
        
        # Area única de drag & drop e status (bem compacta)
        self.drop_frame = ctk.CTkFrame(file_frame, height=65)
        self.drop_frame.pack(fill="x", padx=20, pady=(0, 8))
        self.drop_frame.pack_propagate(False)
        
        # Container interno para centralizar conteúdo
        drop_content = ctk.CTkFrame(self.drop_frame, fg_color="transparent")
        drop_content.pack(expand=True, fill="both", padx=6, pady=5)
        
        # Label principal (título ou arquivo selecionado)
        self.drop_main_label = ctk.CTkLabel(
            drop_content,
            text="Nenhum arquivo selecionado",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=self.colors['text_secondary']
        )
        self.drop_main_label.pack(pady=(2, 1))
        
        # Label secundário (instruções ou caminho)
        self.drop_sub_label = ctk.CTkLabel(
            drop_content,
            text="🎯 Arraste um arquivo PDF aqui ou clique para selecionar",
            font=ctk.CTkFont(size=11),
            text_color=self.colors['text_secondary']
        )
        self.drop_sub_label.pack(pady=(0, 2))
        
        # Bind para clique na área de drop
        self.drop_frame.bind("<Button-1>", lambda e: self.select_pdf_file())
        drop_content.bind("<Button-1>", lambda e: self.select_pdf_file())
        self.drop_main_label.bind("<Button-1>", lambda e: self.select_pdf_file())
        self.drop_sub_label.bind("<Button-1>", lambda e: self.select_pdf_file())

    def create_main_action_button(self):
        """Cria botão principal de ação"""
        button_frame = ctk.CTkFrame(self.processing_container)
        button_frame.pack(fill="x", pady=(0, 8))
        
        # Frame interno para centralizar botão
        button_inner = ctk.CTkFrame(button_frame, fg_color="transparent")
        button_inner.pack(pady=10)
        
        # Botão processar (principal)
        self.process_button = ctk.CTkButton(
            button_inner,
            text="🚀 Processar PDF",
            command=self.process_pdf,
            width=280,
            height=38,
            font=ctk.CTkFont(size=15, weight="bold"),
            fg_color=self.colors['success'],
            hover_color="#259b6e"
        )
        self.process_button.pack()

    def clear_history(self):
        """Limpa o histórico de processamentos"""
        if not self.processing_history:
            return
        
        # Confirmação
        from tkinter import messagebox
        if messagebox.askyesno("Confirmar", "Deseja limpar todo o histórico da sessão?"):
            self.processing_history.clear()
            self.update_history_display()
            self.history_status_label.configure(
                text="Nenhum processamento realizado nesta sessão"
            )

    def add_to_history(self, pdf_file, success, result_data):
        """Adiciona processamento ao histórico"""
        entry = HistoryEntry(
            timestamp=datetime.now(),
            pdf_file=pdf_file,
            success=success,
            result_data=result_data,
            logs=self.current_logs.copy()
        )
        
        self.processing_history.append(entry)
        self.update_history_display()
        
        # Atualiza status
        total = len(self.processing_history)
        success_count = sum(1 for h in self.processing_history if h.success)
        self.history_status_label.configure(
            text=f"{total} processamentos realizados ({success_count} sucessos, {total - success_count} falhas)"
        )

    def update_history_display(self):
        """Atualiza a exibição do histórico"""
        # Limpa lista atual
        for widget in self.history_list_frame.winfo_children():
            widget.destroy()
        
        # Adiciona entradas do histórico (mais recentes primeiro)
        for i, entry in enumerate(reversed(self.processing_history)):
            self.create_history_item(entry, len(self.processing_history) - 1 - i)

    def create_history_item(self, entry, index):
        """Cria um item visual no histórico"""
        # Frame do item
        item_frame = ctk.CTkFrame(self.history_list_frame)
        item_frame.pack(fill="x", padx=5, pady=2)
        
        # Frame interno para layout
        inner_frame = ctk.CTkFrame(item_frame, fg_color="transparent")
        inner_frame.pack(fill="x", padx=10, pady=8)
        
        # Ícone de status
        status_icon = "✅" if entry.success else "❌"
        status_color = self.colors['success'] if entry.success else self.colors['error']
        
        status_label = ctk.CTkLabel(
            inner_frame,
            text=status_icon,
            font=ctk.CTkFont(size=16),
            text_color=status_color,
            width=30
        )
        status_label.pack(side="left", padx=(0, 10))
        
        # Informações principais
        info_frame = ctk.CTkFrame(inner_frame, fg_color="transparent")
        info_frame.pack(side="left", fill="x", expand=True)
        
        # Linha 1: Arquivo Excel final e timestamp
        line1_frame = ctk.CTkFrame(info_frame, fg_color="transparent")
        line1_frame.pack(fill="x")
        
        # Mostra nome do arquivo Excel final ao invés do PDF
        excel_filename = "Arquivo não criado"
        if entry.success and entry.result_data.get('arquivo_final'):
            excel_filename = entry.result_data['arquivo_final']
        
        file_label = ctk.CTkLabel(
            line1_frame,
            text=f"📄 {excel_filename}",
            font=ctk.CTkFont(size=12, weight="bold"),
            anchor="w"
        )
        file_label.pack(side="left")
        
        time_label = ctk.CTkLabel(
            line1_frame,
            text=entry.timestamp.strftime("%H:%M:%S"),
            font=ctk.CTkFont(size=10),
            text_color=self.colors['text_secondary'],
            anchor="e"
        )
        time_label.pack(side="right")
        
        # Linha 2: Resultado
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
        
        # Botão "Ver Detalhes"
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
        """Abre o arquivo de dados específico do processamento"""
        if not entry.success or not entry.result_data.get('arquivo_final'):
            return
        
        # Constrói caminho completo do arquivo
        if self.trabalho_dir:
            file_path = Path(self.trabalho_dir) / entry.result_data['arquivo_final']
        else:
            return
        
        if not file_path.exists():
            messagebox.showerror("Erro", f"Arquivo não encontrado:\n{file_path}")
            return
        
        try:
            # Abre arquivo no sistema
            if sys.platform == "win32":
                os.startfile(file_path)
            elif sys.platform == "darwin":  # macOS
                os.system(f"open '{file_path}'")
            else:  # Linux
                os.system(f"xdg-open '{file_path}'")
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir o arquivo:\n{e}")

    def show_history_details(self, entry):
        """Mostra detalhes de um processamento do histórico"""
        # Cria janela de detalhes
        details_window = ctk.CTkToplevel(self.root)
        details_window.title(f"Detalhes - {entry.pdf_file}")
        
        # Remove efeitos de transparência
        details_window.attributes('-alpha', 1.0)
        
        # Configurações básicas primeiro
        details_window.transient(self.root)
        
        # Cria interface primeiro
        self._create_details_interface(details_window, entry)
        
        # Posiciona DEPOIS de criar a interface, com delay
        details_window.after(50, lambda: self._position_window_relative_to_parent(details_window, 700, 600))
        
        # Força foco sem grab
        details_window.focus_force()
        details_window.lift()
        details_window.attributes('-topmost', True)
        details_window.after(100, lambda: details_window.attributes('-topmost', False))

    def _create_details_interface(self, details_window, entry):
        """Cria a interface da janela de detalhes separadamente"""
    def _create_details_interface(self, details_window, entry):
        """Cria a interface da janela de detalhes separadamente"""
        # Header
        header_frame = ctk.CTkFrame(details_window)
        header_frame.pack(fill="x", padx=20, pady=(20, 10))
        
        status_icon = "✅" if entry.success else "❌"
        status_text = "SUCESSO" if entry.success else "FALHA"
        status_color = self.colors['success'] if entry.success else self.colors['error']
        
        title_label = ctk.CTkLabel(
            header_frame,
            text=f"{status_icon} {status_text} - {entry.pdf_file}",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=status_color
        )
        title_label.pack(pady=15)
        
        time_label = ctk.CTkLabel(
            header_frame,
            text=f"Processado em: {entry.timestamp.strftime('%d/%m/%Y às %H:%M:%S')}",
            font=ctk.CTkFont(size=12),
            text_color=self.colors['text_secondary']
        )
        time_label.pack(pady=(0, 15))
        
        # Resultados
        if entry.success:
            results_frame = ctk.CTkFrame(details_window)
            results_frame.pack(fill="x", padx=20, pady=(0, 10))
            
            results_title = ctk.CTkLabel(
                results_frame,
                text="📊 Resultados do Processamento",
                font=ctk.CTkFont(size=14, weight="bold")
            )
            results_title.pack(pady=(15, 10))
            
            # Estatísticas
            stats_text = f"""
Total de períodos: {entry.result_data.get('total_extracted', 0)}
FOLHA NORMAL: {entry.result_data.get('folha_normal_periods', 0)} períodos
13 SALÁRIO: {entry.result_data.get('salario_13_periods', 0)} períodos
Arquivo criado: {entry.result_data.get('arquivo_final', 'N/A')}
"""
            if entry.result_data.get('person_name'):
                stats_text += f"Nome detectado: {entry.result_data['person_name']}\n"
            
            stats_label = ctk.CTkLabel(
                results_frame,
                text=stats_text.strip(),
                font=ctk.CTkFont(size=11),
                anchor="w",
                justify="left"
            )
            stats_label.pack(padx=15, pady=(0, 15), anchor="w")
        
        # Logs
        logs_frame = ctk.CTkFrame(details_window)
        logs_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        
        logs_title = ctk.CTkLabel(
            logs_frame,
            text="📝 Log do Processamento",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        logs_title.pack(pady=(15, 10))
        
        # Textbox com logs
        logs_textbox = ctk.CTkTextbox(
            logs_frame,
            font=ctk.CTkFont(family="Consolas", size=10)
        )
        logs_textbox.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        
        # Adiciona logs
        for log in entry.logs:
            logs_textbox.insert("end", log + "\n")
        
        logs_textbox.configure(state="disabled")
        
        # Botão fechar
        close_button = ctk.CTkButton(
            details_window,
            text="Fechar",
            command=lambda: self._close_details_window(details_window),
            width=100
        )
        close_button.pack(pady=(0, 20))

    def _position_window_relative_to_parent(self, window, width, height):
        """Posiciona uma janela relativamente à janela principal usando coordenadas exatas"""
        try:
            # Força atualização da janela principal para obter coordenadas corretas
            self.root.update_idletasks()
            
            # Obtém coordenadas absolutas da janela principal na tela
            parent_x = self.root.winfo_rootx()
            parent_y = self.root.winfo_rooty()
            parent_width = self.root.winfo_width()
            parent_height = self.root.winfo_height()
            
            # Calcula posição centralizada relativamente à janela principal
            x = parent_x + (parent_width - width) // 2
            y = parent_y + (parent_height - height) // 2
            
            # Garante que não fique fora da tela (mas prioriza estar no mesmo monitor)
            screen_width = self.root.winfo_screenwidth()
            screen_height = self.root.winfo_screenheight()
            
            # Ajusta apenas se realmente sair da tela
            if x + width > screen_width:
                x = screen_width - width - 10
            if x < 0:
                x = 10
            if y + height > screen_height:
                y = screen_height - height - 10
            if y < 0:
                y = 10
            
            # Define geometria com coordenadas absolutas
            window.geometry(f"{width}x{height}+{x}+{y}")
            
        except Exception as e:
            # Fallback: centraliza na tela
            screen_width = self.root.winfo_screenwidth()
            screen_height = self.root.winfo_screenheight()
            x = (screen_width - width) // 2
            y = (screen_height - height) // 2
            window.geometry(f"{width}x{height}+{x}+{y}")

    def _close_details_window(self, window):
        """Fecha janela de detalhes sem efeitos de transparência"""
        window.destroy()
        
        # Força foco na janela principal sem grab/release
        self.root.focus_force()
        self.root.lift()
        self.root.attributes('-alpha', 1.0)  # Força opacidade total

    def setup_drag_drop(self):
        """Configura drag and drop de arquivos"""
        try:
            self.drop_frame.drop_target_register(DND_FILES)
            self.drop_frame.dnd_bind('<<Drop>>', self.handle_drop)
        except:
            # Se TkinterDnD não funcionar, apenas desabilita drag & drop
            pass

    def handle_drop(self, event):
        """Manipula arquivos arrastados"""
        files = self.root.tk.splitlist(event.data)
        if files:
            file_path = files[0]
            if file_path.lower().endswith('.pdf'):
                self.selected_file = file_path
                self.update_selected_file_display()
            else:
                messagebox.showerror("Erro", "Por favor, selecione apenas arquivos PDF.")

    def load_initial_config(self):
        """Carrega configuração inicial em segundo plano"""

        def task():
            # Mensagem inicial na interface
            self.root.after(
                0,
                lambda: self.config_status.configure(
                    text="🔄 Carregando configurações...",
                    text_color=self.colors['text_secondary'],
                ),
            )
            self.root.after(
                0,
                lambda: self.add_log_message("Iniciando carregamento de configuração"),
            )
            try:
                processor = self._get_processor()
                processor.load_env_config()
                if processor.trabalho_dir:
                    def apply_dir():
                        self.dir_entry.delete(0, 'end')
                        self.dir_entry.insert(0, processor.trabalho_dir)
                        self.validate_config()
                        self.add_log_message("Configuração inicial carregada")

                    self.root.after(0, apply_dir)
            except Exception as e:
                self.root.after(
                    0,
                    lambda: self.add_log_message(
                        f"Erro ao carregar configuração: {e}"
                    ),
                )
            finally:
                self.root.after(
                    0,
                    lambda: self.config_status.configure(
                        text="✅ Configuração pronta",
                        text_color=self.colors['success'],
                    ),
                )

        threading.Thread(target=task, daemon=True).start()

    def select_directory(self):
        """Abre diálogo para seleção de diretório"""
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
            processor = self._get_processor()
            processor.set_trabalho_dir(directory)
            valid, message = processor.validate_trabalho_dir()
            
            if valid:
                self.config_status.configure(
                    text="✅ Configuração válida",
                    text_color=self.colors['success']
                )
                self.trabalho_dir = directory
                self.update_pdf_list()
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
        """Atualiza lista de PDFs disponíveis no diretório"""
        if not self.trabalho_dir:
            return

        try:
            pdf_files = self._get_processor().get_pdf_files_in_trabalho_dir()
            if pdf_files:
                self.add_log_message(f"PDFs encontrados: {', '.join(pdf_files)}")
            else:
                self.add_log_message("Nenhum arquivo PDF encontrado no diretório.")
        except Exception as e:
            self.add_log_message(f"Erro ao listar PDFs: {e}")

    def select_pdf_file(self):
        """Abre diálogo para seleção de PDF"""
        if not self.trabalho_dir:
            messagebox.showwarning("Aviso", "Configure o diretório de trabalho primeiro.")
            return
        
        file_path = filedialog.askopenfilename(
            title="Selecione o arquivo PDF",
            initialdir=self.trabalho_dir,
            filetypes=[
                ("Arquivos PDF", "*.pdf"),
                ("Todos os arquivos", "*.*")
            ]
        )
        
        if file_path:
            self.selected_file = file_path
            self.update_selected_file_display()

    def update_selected_file_display(self):
        """Atualiza display do arquivo selecionado"""
        if self.selected_file:
            filename = Path(self.selected_file).name
            # Atualiza labels da área compacta
            self.drop_main_label.configure(
                text=f"📄 {filename}",
                text_color=self.colors['success']
            )
            self.drop_sub_label.configure(
                text=f"Arquivo selecionado: {filename}",
                text_color=self.colors['success']
            )
            self.add_log_message(f"Arquivo selecionado: {filename}")
        else:
            # Volta ao estado inicial
            self.drop_main_label.configure(
                text="Nenhum arquivo selecionado",
                text_color=self.colors['text_secondary']
            )
            self.drop_sub_label.configure(
                text="🎯 Arraste um arquivo PDF aqui ou clique para selecionar",
                text_color=self.colors['text_secondary']
            )

    def add_log_message(self, message):
        """Adiciona mensagem ao log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}"
        
        # Adiciona ao log atual da sessão
        self.current_logs.append(log_entry)
        
        # Adiciona ao textbox do popup se estiver aberto
        if self.popup_log_textbox:
            self.popup_log_textbox.insert("end", log_entry + "\n")
            self.popup_log_textbox.see("end")
        
        # Atualiza a interface
        self.root.update_idletasks()

    def update_progress(self, progress, message=""):
        """Atualiza barra de progresso"""
        # Normaliza progresso (0-100 para 0-1)
        normalized_progress = max(0, min(100, progress)) / 100
        
        # Atualiza popup se estiver aberto
        if self.popup_progress_bar:
            self.popup_progress_bar.set(normalized_progress)
        
        if message and self.popup_progress_label:
            self.popup_progress_label.configure(text=message)
        
        # Atualiza a interface
        self.root.update_idletasks()

    def process_pdf(self):
        """Processa o PDF selecionado"""
        if self.processing:
            return
        
        # Validações
        if not self.trabalho_dir:
            messagebox.showerror("Erro", "Configure o diretório de trabalho primeiro.")
            return
        
        if not self.selected_file:
            messagebox.showerror("Erro", "Selecione um arquivo PDF primeiro.")
            return
        
        # Configura processador
        sheet_name = self.sheet_entry.get().strip() if self.sheet_entry else ""
        processor = self._get_processor()
        if sheet_name:
            processor.preferred_sheet = sheet_name
        
        # Mostra popup de processamento
        self.processing_popup.show()
        
        # Inicia processamento em thread separada
        self.processing = True
        self.process_button.configure(text="🔄 Processando...", state="disabled")
        
        thread = threading.Thread(target=self._process_pdf_thread)
        thread.daemon = True
        thread.start()

    def _process_pdf_thread(self):
        """Thread de processamento do PDF"""
        try:
            # Obtém nome do arquivo relativo ao diretório de trabalho
            if Path(self.selected_file).parent == Path(self.trabalho_dir):
                pdf_filename = Path(self.selected_file).name
            else:
                pdf_filename = self.selected_file
            
            # Processa PDF
            results = self._get_processor().process_pdf(pdf_filename)
            
            # Atualiza interface no thread principal
            self.root.after(0, self._process_complete, results)
            
        except Exception as e:
            self.root.after(0, self._process_error, str(e))

    def _process_complete(self, results):
        """Callback quando processamento completa"""
        self.processing = False
        self.process_button.configure(text="🚀 Processar PDF", state="normal")
        
        # Fecha popup de processamento
        self.processing_popup.close()
        
        # Adiciona ao histórico
        pdf_filename = Path(self.selected_file).name if self.selected_file else "Arquivo desconhecido"
        self.add_to_history(pdf_filename, results['success'], results)
        
        if results['success']:
            # Mostra resultados
            total = results['total_extracted']
            normal = results['folha_normal_periods']
            salario_13 = results['salario_13_periods']
            
            success_msg = f"✅ Processamento concluído!\n\n"
            success_msg += f"📊 Total: {total} períodos processados\n"
            if normal > 0:
                success_msg += f"📄 FOLHA NORMAL: {normal} períodos\n"
            if salario_13 > 0:
                success_msg += f"💰 13 SALÁRIO: {salario_13} períodos\n"
            
            if results.get('person_name'):
                success_msg += f"\n👤 Nome detectado: {results['person_name']}"
            else:
                success_msg += f"\n📄 Nome não detectado (usando nome do PDF)"
            
            success_msg += f"\n\n💾 Arquivo criado: {results['arquivo_final']}"
            success_msg += f"\n\n📊 Resultado adicionado ao histórico da sessão."
            
            messagebox.showinfo("Sucesso", success_msg)
        else:
            messagebox.showerror("Erro", f"Erro no processamento:\n\n{results['error']}")
        
        # Limpa interface para novo processamento
        self.clear_processing_interface()
        
        # Vai para aba apropriada
        if results['success']:
            # Se sucesso, vai para histórico
            self.tabview.set("📊 Histórico")
        else:
            # Se erro de configuração, vai para configurações
            error_msg = results.get('error', '').lower()
            if any(term in error_msg for term in ['planilha', 'sheet', 'levantamento dados']):
                self.tabview.set("⚙️ Configurações Avançadas")
            else:
                self.tabview.set("📊 Histórico")

    def clear_processing_interface(self):
        """Limpa a interface de processamento para nova operação"""
        # Limpa arquivo selecionado
        self.selected_file = None
        self.update_selected_file_display()
        
        # Limpa campos da aba de configurações (se existirem)
        if self.sheet_entry:
            self.sheet_entry.delete(0, 'end')
        self.verbose_var.set(False)
        
        # Limpa logs da sessão atual para próximo processamento
        self.current_logs = []

    def _process_error(self, error_message):
        """Callback quando processamento falha"""
        self.processing = False
        self.process_button.configure(text="🚀 Processar PDF", state="normal")
        
        # Fecha popup de processamento
        self.processing_popup.close()
        
        # Adiciona falha ao histórico
        pdf_filename = Path(self.selected_file).name if self.selected_file else "Arquivo desconhecido"
        error_result = {'success': False, 'error': error_message}
        self.add_to_history(pdf_filename, False, error_result)
        
        messagebox.showerror("Erro", f"Erro no processamento:\n\n{error_message}")
        
        # Vai para aba apropriada baseada no tipo de erro
        error_msg = error_message.lower()
        if any(term in error_msg for term in ['planilha', 'sheet', 'levantamento dados']):
            self.tabview.set("⚙️ Configurações Avançadas")
        else:
            self.tabview.set("📊 Histórico")

    def on_closing(self):
        """Manipula fechamento da aplicação"""
        if self.processing:
            result = messagebox.askyesno(
                "Processamento em andamento",
                "Há um processamento em andamento. Deseja realmente fechar a aplicação?"
            )
            if not result:
                return
        
        self.root.destroy()

    def run(self):
        """Inicia a aplicação"""
        # Bind para fechamento
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Bind para validar config quando o campo muda
        self.dir_entry.bind('<KeyRelease>', lambda e: self.root.after(500, self.validate_config))
        
        # Inicia loop principal
        self.root.mainloop()


def main():
    """Função principal"""
    try:
        # Verifica dependências
        import customtkinter
        
        # Tenta importar TkinterDnD (opcional)
        try:
            import tkinterdnd2
        except ImportError:
            print("AVISO: tkinterdnd2 não instalado. Drag & drop desabilitado.")
            print("Para habilitar: pip install tkinterdnd2")
        
        # Cria e executa aplicação
        app = PDFExcelDesktopApp()
        app.run()
        
    except ImportError as e:
        error_msg = """
ERRO: Dependências não instaladas!

Para instalar as dependências necessárias, execute:

pip install customtkinter pillow

Opcionalmente (para drag & drop):
pip install tkinterdnd2

Dependências faltando: {}
        """.format(str(e))
        
        print(error_msg)
        
        # Tenta mostrar em messagebox se tkinter básico funcionar
        try:
            import tkinter.messagebox as mb
            mb.showerror("Dependências não instaladas", error_msg)
        except:
            pass
        
        sys.exit(1)
    
    except Exception as e:
        error_msg = f"Erro ao iniciar aplicação: {e}"
        print(error_msg)
        
        try:
            import tkinter.messagebox as mb
            mb.showerror("Erro", error_msg)
        except:
            pass
        
        sys.exit(1)


if __name__ == "__main__":
    main()