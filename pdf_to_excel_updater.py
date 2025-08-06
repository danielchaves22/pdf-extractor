#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF para Excel Updater - Interface de Linha de Comando
======================================================

Versão 3.2 - Refatorado para usar PDFProcessorCore

Interface de linha de comando que utiliza o módulo pdf_processor_core.py
para toda a lógica de processamento. Elimina duplicação de código.

Funcionalidades:
- Diretório de trabalho obrigatório configurado via MODELO_DIR no .env
- Interface gráfica para seleção de PDF (opcional)
- Detecta nome da pessoa no PDF e usa para nomear arquivo final
- Processa FOLHA NORMAL (linhas 1-65) e 13 SALARIO (linhas 67+)
- Usa PDFProcessorCore para toda a lógica

Estrutura esperada no diretório de trabalho:
├── MODELO.xlsm          # Planilha modelo (obrigatório)
├── arquivo.pdf          # PDF a processar
└── DADOS/               # Pasta criada automaticamente
    └── NOME DA PESSOA.xlsm     # Resultado com nome da pessoa

Configuração .env:
MODELO_DIR=C:\\caminho\\para\\diretorio\\de\\trabalho

Uso:
python pdf_to_excel_updater.py                    # Abre seletor de arquivo
python pdf_to_excel_updater.py arquivo.pdf        # Processa arquivo específico

Autor: Sistema de Extração Automatizada
Data: 2025
"""

import argparse
import logging
import sys
import os
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

# Configuração de encoding para Windows
if sys.platform == "win32":
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
    # Tenta configurar console para UTF-8
    try:
        os.system('chcp 65001 > nul')
    except:
        pass

# Importa o processador core
try:
    from pdf_processor_core import PDFProcessorCore
except ImportError:
    print("ERRO: Módulo pdf_processor_core.py não encontrado!")
    print("Certifique-se de que o arquivo pdf_processor_core.py está na mesma pasta.")
    sys.exit(1)

# Configuração de logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def safe_print(message, fallback_message=None):
    """Imprime mensagem com fallback para sistemas sem suporte a Unicode"""
    try:
        print(message)
    except UnicodeEncodeError:
        if fallback_message:
            print(fallback_message)
        else:
            # Remove emojis e caracteres especiais
            clean_message = message.encode('ascii', errors='ignore').decode('ascii')
            print(clean_message)

class CLILogHandler:
    """Handler para logs da CLI que imprime diretamente no console"""
    
    def __init__(self, verbose=False):
        self.verbose = verbose
    
    def log_callback(self, message):
        """Callback para receber logs do processador"""
        if self.verbose:
            safe_print(message)
        elif not message.startswith('[DEBUG]'):
            # Remove prefixo [INFO], [WARNING], etc. para output mais limpo
            clean_message = message
            if message.startswith('['):
                clean_message = message.split('] ', 1)[-1] if '] ' in message else message
            safe_print(clean_message)

class PDFToExcelUpdater:
    """Wrapper da CLI para PDFProcessorCore - para compatibilidade"""
    
    def __init__(self, verbose=False):
        """Inicializa o updater usando PDFProcessorCore"""
        
        # Cria handler de logs
        self.log_handler = CLILogHandler(verbose)
        
        # Inicializa o processador core
        self.processor = PDFProcessorCore(
            progress_callback=None,  # CLI não usa progress bar
            log_callback=self.log_handler.log_callback
        )
        
        # Para compatibilidade com interface antiga
        self.preferred_sheet = None

    def load_env_config(self):
        """Carrega configurações do arquivo .env"""
        return self.processor.load_env_config()

    def select_pdf_file(self):
        """Abre diálogo para seleção de arquivo PDF no diretório de trabalho"""
        try:
            # Carrega configuração primeiro
            self.processor.load_env_config()
            
            # Cria janela invisível
            root = tk.Tk()
            root.withdraw()  # Esconde a janela principal
            root.attributes('-topmost', True)  # Mantém diálogo na frente
            
            # Verifica se há PDFs disponíveis
            pdf_files = self.processor.get_pdf_files_in_trabalho_dir()
            
            if not pdf_files:
                messagebox.showwarning(
                    "Nenhum PDF encontrado", 
                    f"Nenhum arquivo PDF encontrado no diretório de trabalho:\n{self.processor.trabalho_dir}"
                )
                root.destroy()
                return None
            
            # Abre diálogo de seleção
            selected_file = filedialog.askopenfilename(
                title="Selecione o arquivo PDF para processar",
                initialdir=self.processor.trabalho_dir,
                filetypes=[
                    ("Arquivos PDF", "*.pdf"),
                    ("Todos os arquivos", "*.*")
                ]
            )
            
            root.destroy()
            
            if selected_file:
                # Retorna apenas o nome do arquivo (sem caminho)
                return Path(selected_file).name
            else:
                return None
                
        except Exception as e:
            if 'root' in locals():
                root.destroy()
            logger.error(f"Erro ao abrir diálogo de seleção: {e}")
            return None

    def process_pdf(self, pdf_filename: str):
        """Processa PDF usando PDFProcessorCore"""
        
        # Configura planilha preferida se especificada
        if self.preferred_sheet:
            self.processor.preferred_sheet = self.preferred_sheet
        
        # Processa usando o core
        results = self.processor.process_pdf(pdf_filename)
        
        if not results['success']:
            raise ValueError(results['error'])
        
        # Retorna no formato esperado pela CLI antiga (para compatibilidade)
        extracted_data = {
            'FOLHA NORMAL': {},
            '13 SALARIO': {}
        }
        
        # Simula estrutura antiga se necessário
        return extracted_data, results['excel_path'], results['person_name']

def print_results_summary(results):
    """Imprime resumo dos resultados de forma organizada"""
    
    if not results['success']:
        safe_print(f"❌ ERRO: {results['error']}", f"ERRO: {results['error']}")
        return
    
    # Resultados básicos
    total = results['total_extracted']
    normal = results['folha_normal_periods']
    salario_13 = results['salario_13_periods']
    
    safe_print(f"\n✅ Processamento concluído: {total} períodos processados", 
               f"\nOK: Processamento concluido: {total} periodos processados")
    
    if normal > 0:
        safe_print(f"   📄 FOLHA NORMAL: {normal} períodos", 
                   f"   FOLHA NORMAL: {normal} periodos")
    if salario_13 > 0:
        safe_print(f"   💰 13 SALÁRIO: {salario_13} períodos", 
                   f"   13 SALARIO: {salario_13} periodos")
    
    # Nome detectado
    if results.get('person_name'):
        safe_print(f"\n👤 Nome detectado: {results['person_name']}", 
                   f"\nNome detectado: {results['person_name']}")
    else:
        safe_print(f"\n📄 Nome não detectado - usando nome do PDF", 
                   f"\nNome nao detectado - usando nome do PDF")
    
    # Arquivo criado
    safe_print(f"\n💾 Arquivo criado: {results['arquivo_final']}", 
               f"\nArquivo criado: {results['arquivo_final']}")
    
    # Estatísticas detalhadas se houver falhas
    if results.get('failed_periods'):
        failed_count = len(results['failed_periods'])
        success_count = results['success_periods']
        total_periods = results['total_periods']
        
        if failed_count > 0:
            safe_print(f"\n⚠️  Atenção: {failed_count} períodos falharam:", 
                       f"\nAtencao: {failed_count} periodos falharam:")
            for failed in results['failed_periods'][:3]:
                safe_print(f"   • {failed}", f"   - {failed}")
            if failed_count > 3:
                safe_print(f"   • ... e mais {failed_count - 3} períodos", 
                           f"   - ... e mais {failed_count - 3} periodos")

def main():
    """Função principal da aplicação"""
    parser = argparse.ArgumentParser(
        description='PDF para Excel Updater v3.2 - CLI usando PDFProcessorCore',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Exemplos de uso:
  python pdf_to_excel_updater.py                    # Abre seletor de arquivo
  python pdf_to_excel_updater.py arquivo.pdf        # Processa arquivo específico
  python pdf_to_excel_updater.py arquivo.pdf -v     # Modo verboso
  python pdf_to_excel_updater.py arquivo.pdf -s "PLANILHA"  # Planilha específica

Configuração:
  Configure MODELO_DIR no arquivo .env apontando para o diretório que contém MODELO.xlsm
        """
    )
    
    parser.add_argument(
        'pdf_filename', 
        nargs='?', 
        help='Nome do arquivo PDF (opcional - abrirá diálogo se não fornecido)'
    )
    parser.add_argument(
        '-s', '--sheet', 
        help='Nome da planilha específica (padrão: "LEVANTAMENTO DADOS")'
    )
    parser.add_argument(
        '-v', '--verbose', 
        action='store_true', 
        help='Modo verboso (mostra logs detalhados)'
    )
    
    args = parser.parse_args()
    
    # Configura nível de log baseado no verbose
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)
    else:
        logging.getLogger().setLevel(logging.INFO)
    
    try:
        # Cria updater
        updater = PDFToExcelUpdater(verbose=args.verbose)
        
        # Configura planilha se especificada
        if args.sheet:
            updater.preferred_sheet = args.sheet
        
        # Determina qual PDF processar
        pdf_filename = args.pdf_filename
        
        if not pdf_filename:
            # Se não foi fornecido PDF, abre diálogo de seleção
            if not args.verbose:
                safe_print("Abrindo seletor de arquivo...")
            
            pdf_filename = updater.select_pdf_file()
            
            if not pdf_filename:
                safe_print("❌ CANCELADO: Nenhum arquivo selecionado", 
                           "CANCELADO: Nenhum arquivo selecionado")
                return 0
        
        if not args.verbose:
            safe_print(f"🔄 Processando: {pdf_filename}", 
                       f"Processando: {pdf_filename}")
        
        # Processa PDF usando o core
        results = updater.processor.process_pdf(pdf_filename)
        
        # Imprime resultados
        print_results_summary(results)
        
        return 0 if results['success'] else 1
        
    except ValueError as e:
        safe_print(f"❌ ERRO: {e}", f"ERRO: {e}")
        return 1
    except Exception as e:
        safe_print(f"❌ ERRO: Erro inesperado: {e}", f"ERRO: Erro inesperado: {e}")
        if args.verbose:
            raise
        return 1

if __name__ == "__main__":
    sys.exit(main())

class CLILogHandler:
    """Handler para logs da CLI que imprime diretamente no console"""
    
    def __init__(self, verbose=False):
        self.verbose = verbose
    
    def log_callback(self, message):
        """Callback para receber logs do processador"""
        if self.verbose:
            print(message)
        elif not message.startswith('[DEBUG]'):
            # Remove prefixo [INFO], [WARNING], etc. para output mais limpo
            clean_message = message
            if message.startswith('['):
                clean_message = message.split('] ', 1)[-1] if '] ' in message else message
            print(clean_message)

class PDFToExcelUpdater:
    """Wrapper da CLI para PDFProcessorCore - para compatibilidade"""
    
    def __init__(self, verbose=False):
        """Inicializa o updater usando PDFProcessorCore"""
        
        # Cria handler de logs
        self.log_handler = CLILogHandler(verbose)
        
        # Inicializa o processador core
        self.processor = PDFProcessorCore(
            progress_callback=None,  # CLI não usa progress bar
            log_callback=self.log_handler.log_callback
        )
        
        # Para compatibilidade com interface antiga
        self.preferred_sheet = None

    def load_env_config(self):
        """Carrega configurações do arquivo .env"""
        return self.processor.load_env_config()

    def select_pdf_file(self):
        """Abre diálogo para seleção de arquivo PDF no diretório de trabalho"""
        try:
            # Carrega configuração primeiro
            self.processor.load_env_config()
            
            # Cria janela invisível
            root = tk.Tk()
            root.withdraw()  # Esconde a janela principal
            root.attributes('-topmost', True)  # Mantém diálogo na frente
            
            # Verifica se há PDFs disponíveis
            pdf_files = self.processor.get_pdf_files_in_trabalho_dir()
            
            if not pdf_files:
                messagebox.showwarning(
                    "Nenhum PDF encontrado", 
                    f"Nenhum arquivo PDF encontrado no diretório de trabalho:\n{self.processor.trabalho_dir}"
                )
                root.destroy()
                return None
            
            # Abre diálogo de seleção
            selected_file = filedialog.askopenfilename(
                title="Selecione o arquivo PDF para processar",
                initialdir=self.processor.trabalho_dir,
                filetypes=[
                    ("Arquivos PDF", "*.pdf"),
                    ("Todos os arquivos", "*.*")
                ]
            )
            
            root.destroy()
            
            if selected_file:
                # Retorna apenas o nome do arquivo (sem caminho)
                return Path(selected_file).name
            else:
                return None
                
        except Exception as e:
            if 'root' in locals():
                root.destroy()
            logger.error(f"Erro ao abrir diálogo de seleção: {e}")
            return None

    def process_pdf(self, pdf_filename: str):
        """Processa PDF usando PDFProcessorCore"""
        
        # Configura planilha preferida se especificada
        if self.preferred_sheet:
            self.processor.preferred_sheet = self.preferred_sheet
        
        # Processa usando o core
        results = self.processor.process_pdf(pdf_filename)
        
        if not results['success']:
            raise ValueError(results['error'])
        
        # Retorna no formato esperado pela CLI antiga (para compatibilidade)
        extracted_data = {
            'FOLHA NORMAL': {},
            '13 SALARIO': {}
        }
        
        # Simula estrutura antiga se necessário
        return extracted_data, results['excel_path'], results['person_name']

def print_results_summary(results):
    """Imprime resumo dos resultados de forma organizada"""
    
    if not results['success']:
        print(f"❌ ERRO: {results['error']}")
        return
    
    # Resultados básicos
    total = results['total_extracted']
    normal = results['folha_normal_periods']
    salario_13 = results['salario_13_periods']
    
    print(f"\n✅ Processamento concluído: {total} períodos processados")
    
    if normal > 0:
        print(f"   📄 FOLHA NORMAL: {normal} períodos")
    if salario_13 > 0:
        print(f"   💰 13 SALÁRIO: {salario_13} períodos")
    
    # Nome detectado
    if results.get('person_name'):
        print(f"\n👤 Nome detectado: {results['person_name']}")
    else:
        print(f"\n📄 Nome não detectado - usando nome do PDF")
    
    # Arquivo criado
    print(f"\n💾 Arquivo criado: {results['arquivo_final']}")
    
    # Estatísticas detalhadas se houver falhas
    if results.get('failed_periods'):
        failed_count = len(results['failed_periods'])
        success_count = results['success_periods']
        total_periods = results['total_periods']
        
        if failed_count > 0:
            print(f"\n⚠️  Atenção: {failed_count} períodos falharam:")
            for failed in results['failed_periods'][:3]:
                print(f"   • {failed}")
            if failed_count > 3:
                print(f"   • ... e mais {failed_count - 3} períodos")

def main():
    """Função principal da aplicação"""
    parser = argparse.ArgumentParser(
        description='PDF para Excel Updater v3.2 - CLI usando PDFProcessorCore',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Exemplos de uso:
  python pdf_to_excel_updater.py                    # Abre seletor de arquivo
  python pdf_to_excel_updater.py arquivo.pdf        # Processa arquivo específico
  python pdf_to_excel_updater.py arquivo.pdf -v     # Modo verboso
  python pdf_to_excel_updater.py arquivo.pdf -s "PLANILHA"  # Planilha específica

Configuração:
  Configure MODELO_DIR no arquivo .env apontando para o diretório que contém MODELO.xlsm
        """
    )
    
    parser.add_argument(
        'pdf_filename', 
        nargs='?', 
        help='Nome do arquivo PDF (opcional - abrirá diálogo se não fornecido)'
    )
    parser.add_argument(
        '-s', '--sheet', 
        help='Nome da planilha específica (padrão: "LEVANTAMENTO DADOS")'
    )
    parser.add_argument(
        '-v', '--verbose', 
        action='store_true', 
        help='Modo verboso (mostra logs detalhados)'
    )
    
    args = parser.parse_args()
    
    # Configura nível de log baseado no verbose
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)
    else:
        logging.getLogger().setLevel(logging.INFO)
    
    try:
        # Cria updater
        updater = PDFToExcelUpdater(verbose=args.verbose)
        
        # Configura planilha se especificada
        if args.sheet:
            updater.preferred_sheet = args.sheet
        
        # Determina qual PDF processar
        pdf_filename = args.pdf_filename
        
        if not pdf_filename:
            # Se não foi fornecido PDF, abre diálogo de seleção
            if not args.verbose:
                print("Abrindo seletor de arquivo...")
            
            pdf_filename = updater.select_pdf_file()
            
            if not pdf_filename:
                print("❌ CANCELADO: Nenhum arquivo selecionado")
                return 0
        
        if not args.verbose:
            print(f"🔄 Processando: {pdf_filename}")
        
        # Processa PDF usando o core
        results = updater.processor.process_pdf(pdf_filename)
        
        # Imprime resultados
        print_results_summary(results)
        
        return 0 if results['success'] else 1
        
    except ValueError as e:
        print(f"❌ ERRO: {e}")
        return 1
    except Exception as e:
        print(f"❌ ERRO: Erro inesperado: {e}")
        if args.verbose:
            raise
        return 1

if __name__ == "__main__":
    sys.exit(main())