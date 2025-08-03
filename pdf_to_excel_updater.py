#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF para Excel Updater - Processamento com Diretório de Trabalho
================================================================

Versão 3.0 - Modo único com diretório de trabalho configurado via .env

Funcionalidades:
- Diretório de trabalho obrigatório configurado via MODELO_DIR no .env
- Busca PDF e MODELO.xlsm no diretório de trabalho
- Cria pasta DADOS no diretório de trabalho
- Pode ser executado de qualquer local desde que .env esteja configurado

Estrutura esperada no diretório de trabalho:
├── MODELO.xlsm          # Planilha modelo (obrigatório)
├── arquivo.pdf          # PDF a processar
└── DADOS/               # Pasta criada automaticamente
    └── arquivo.xlsm     # Resultado processado

Configuração .env:
MODELO_DIR=C:\\caminho\\para\\diretorio\\de\\trabalho

Autor: Sistema de Extração Automatizada
Data: 2025
"""

import re
import pandas as pd
import pdfplumber
from datetime import datetime
from pathlib import Path
import logging
from typing import Dict, List, Optional, Tuple
import argparse
from openpyxl import load_workbook
import os
import shutil
from dotenv import load_dotenv

# Configuração de logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class PDFToExcelUpdater:
    """Classe para processar PDF usando diretório de trabalho configurado"""
    
    def __init__(self):
        """Inicializa o updater com as regras de mapeamento"""
        
        # Regras de mapeamento para colunas específicas do Excel
        self.mapping_rules = {
            # Obter da coluna ÍNDICE (com fallback especial para PRODUÇÃO)
            '01003601': {'code': 'PREMIO PROD. MENSAL', 'excel_column': 'X', 'source': 'indice', 'fallback_to_valor': True},
            '01007301': {'code': 'HORAS EXT.100%-180', 'excel_column': 'Y', 'source': 'indice'},
            '01009001': {'code': 'ADIC.NOT.25%-180', 'excel_column': 'AC', 'source': 'indice'},
            '01003501': {'code': 'HORAS EXT.75%-180', 'excel_column': 'AA', 'source': 'indice'},
            '02007501': {'code': 'DIFER.PROV. HORAS EXTRAS 75%', 'excel_column': 'AA', 'source': 'indice'},
            
            # Obter da coluna VALOR
            '09090301': {'code': 'SALARIO CONTRIB INSS', 'excel_column': 'B', 'source': 'valor'}
        }
        
        # Planilha preferida (pode ser especificada externamente)
        self.preferred_sheet = None
        
        # Diretório de trabalho
        self.trabalho_dir = None
        
        # Meses em português para conversão
        self.meses_pt = {
            'janeiro': 1, 'fevereiro': 2, 'março': 3, 'abril': 4,
            'maio': 5, 'junho': 6, 'julho': 7, 'agosto': 8,
            'setembro': 9, 'outubro': 10, 'novembro': 11, 'dezembro': 12
        }
        
        # Abreviações de meses
        self.meses_abrev = {
            'jan': 1, 'fev': 2, 'mar': 3, 'abr': 4, 'mai': 5, 'jun': 6,
            'jul': 7, 'ago': 8, 'set': 9, 'out': 10, 'nov': 11, 'dez': 12
        }

    def load_env_config(self):
        """Carrega configurações do arquivo .env"""
        # Carrega o arquivo .env se existir
        if Path('.env').exists():
            load_dotenv()
            logger.debug("Arquivo .env carregado")
        else:
            raise ValueError("Arquivo .env não encontrado. Configure MODELO_DIR no arquivo .env")
        
        # Obtém o diretório de trabalho
        self.trabalho_dir = os.getenv('MODELO_DIR')
        if not self.trabalho_dir:
            raise ValueError("MODELO_DIR não está definido no arquivo .env")
        
        # Valida se diretório existe
        trabalho_dir_path = Path(self.trabalho_dir)
        if not trabalho_dir_path.exists():
            raise ValueError(f"Diretório de trabalho não encontrado: {self.trabalho_dir}")
        
        logger.debug(f"Diretório de trabalho configurado: {self.trabalho_dir}")

    def find_pdf_file(self, pdf_filename: str) -> str:
        """Procura arquivo PDF no diretório de trabalho"""
        trabalho_dir_path = Path(self.trabalho_dir)
        
        # Se pdf_filename já tem caminho, usa direto
        if Path(pdf_filename).is_absolute() or '/' in pdf_filename or '\\' in pdf_filename:
            if Path(pdf_filename).exists():
                return pdf_filename
            else:
                raise ValueError(f"Arquivo PDF não encontrado: {pdf_filename}")
        
        # Procura no diretório de trabalho
        pdf_path = trabalho_dir_path / pdf_filename
        if pdf_path.exists():
            logger.debug(f"PDF encontrado no diretório de trabalho: {pdf_path}")
            return str(pdf_path)
        
        # Tenta com extensão .pdf se não foi fornecida
        if not pdf_filename.lower().endswith('.pdf'):
            pdf_path_with_ext = trabalho_dir_path / f"{pdf_filename}.pdf"
            if pdf_path_with_ext.exists():
                logger.debug(f"PDF encontrado (com extensão): {pdf_path_with_ext}")
                return str(pdf_path_with_ext)
        
        raise ValueError(f"Arquivo PDF '{pdf_filename}' não encontrado no diretório de trabalho: {self.trabalho_dir}")

    def copy_modelo_to_dados(self, pdf_path: str) -> str:
        """Copia o arquivo modelo para a pasta DADOS com nome baseado no PDF"""
        
        trabalho_dir_path = Path(self.trabalho_dir)
        
        # Procura pelo arquivo MODELO.xlsm no diretório de trabalho
        modelo_file = trabalho_dir_path / "MODELO.xlsm"
        if not modelo_file.exists():
            raise ValueError(f"Arquivo MODELO.xlsm não encontrado no diretório de trabalho: {self.trabalho_dir}")
        
        logger.debug(f"Arquivo modelo encontrado: {modelo_file}")
        
        # Cria pasta DADOS no diretório de trabalho
        dados_dir = trabalho_dir_path / "DADOS"
        dados_dir.mkdir(exist_ok=True)
        logger.debug(f"Pasta DADOS: {dados_dir}")
        
        # Define nome do arquivo de destino baseado no PDF
        pdf_base_name = Path(pdf_path).stem
        destino_file = dados_dir / f"{pdf_base_name}.xlsm"
        
        # Copia o modelo para o destino
        try:
            shutil.copy2(modelo_file, destino_file)
            logger.debug(f"Modelo copiado: {modelo_file} -> {destino_file}")
            return str(destino_file)
        except Exception as e:
            raise ValueError(f"Erro ao copiar modelo: {e}")

    def extract_text_from_pdf(self, pdf_path: str) -> List[str]:
        """Extrai texto de todas as páginas do PDF"""
        pages_text = []
        
        try:
            with pdfplumber.open(pdf_path) as pdf:
                logger.debug(f"Processando PDF: {pdf_path} ({len(pdf.pages)} páginas)")
                
                for i, page in enumerate(pdf.pages):
                    text = page.extract_text()
                    if text:
                        pages_text.append(text)
                    
        except Exception as e:
            logger.error(f"Erro ao processar PDF: {e}")
            raise
            
        return pages_text

    def extract_reference_date(self, text: str) -> Optional[Tuple[int, int]]:
        """Extrai a data de referência da página (mês/ano)"""
        patterns = [
            r'Referência:\s*(\w+)/(\d{4})',
            r'Referencia:\s*(\w+)/(\d{4})',
            r'Data\s*do\s*c[aá]lculo:\s*\d{2}/(\d{2})/(\d{4})',
            r'Per[ií]odo:\s*(\w+)/(\d{4})',
            r'Compet[êe]ncia:\s*(\w+)/(\d{4})',
            r'(\w+)\s*/\s*(\d{4})',
        ]
        
        found_matches = []
        
        for i, pattern in enumerate(patterns):
            matches = re.findall(pattern, text, re.IGNORECASE)
            for match in matches:
                found_matches.append((i, match))
                try:
                    if len(match) == 2:
                        mes_str = match[0].lower()
                        ano = int(match[1])
                        
                        # Tenta converter mês por nome
                        mes = self.meses_pt.get(mes_str) or self.meses_abrev.get(mes_str)
                        
                        # Se não conseguir por nome, tenta por número
                        if not mes:
                            try:
                                mes = int(mes_str)
                                if 1 <= mes <= 12:
                                    logger.debug(f"Data extraída (padrão {i+1}): {mes_str}/{ano} -> {mes}/{ano}")
                                    return (mes, ano)
                            except ValueError:
                                continue
                        else:
                            logger.debug(f"Data extraída (padrão {i+1}): {mes_str}/{ano} -> {mes}/{ano}")
                            return (mes, ano)
                except ValueError:
                    continue
        
        # Log quando não conseguir extrair data
        if found_matches:
            logger.debug(f"Padrões encontrados mas inválidos: {found_matches}")
        else:
            logger.debug("Nenhum padrão de data encontrado na página")
            # Mostra as primeiras linhas para debug
            lines = text.split('\n')[:5]
            logger.debug(f"Primeiras linhas da página: {[line.strip() for line in lines if line.strip()]}")
        
        return None

    def extract_last_two_numbers(self, line: str):
        """Extrai os dois últimos números de uma linha, ignorando campos fantasma"""
        
        def convert_to_float_robust(value_str):
            if not value_str or not value_str.strip():
                return None
                
            cleaned = value_str.strip()
            
            # Detecta formato de horas (06:34) e converte para (06,34)
            if ':' in cleaned:
                hour_pattern = r'^\d{1,2}:\d{2}$'
                if re.match(hour_pattern, cleaned):
                    logger.debug(f"Formato de horas detectado: {cleaned}")
                    return cleaned.replace(':', ',')
            
            # Remove caracteres não numéricos exceto ponto e vírgula
            cleaned = re.sub(r'[^\d.,]', '', cleaned)
            
            if not cleaned:
                return None
            
            # Estratégias de conversão para números
            try:
                # Formato brasileiro (1.234,56)
                if ',' in cleaned and cleaned.count(',') == 1:
                    return float(cleaned.replace('.', '').replace(',', '.'))
                # Formato americano (1,234.56)
                elif '.' in cleaned and cleaned.count('.') == 1 and ',' in cleaned:
                    return float(cleaned.replace(',', ''))
                # Apenas vírgula decimal (1234,56)
                elif ',' in cleaned and '.' not in cleaned:
                    return float(cleaned.replace(',', '.'))
                # Apenas ponto decimal (1234.56)
                elif '.' in cleaned and ',' not in cleaned:
                    return float(cleaned)
                else:
                    return float(cleaned)
            except ValueError:
                return None
        
        # Encontra todos os números na linha (incluindo formato de horas)
        number_pattern = r'[\d]+(?:[.,:]\d+)*'
        matches = re.findall(number_pattern, line)
        
        if len(matches) >= 2:
            # Pega os dois últimos números
            penultimo = convert_to_float_robust(matches[-2])
            ultimo = convert_to_float_robust(matches[-1])
            return penultimo, ultimo
        elif len(matches) == 1:
            # Só um número - considera como valor
            ultimo = convert_to_float_robust(matches[-1])
            return None, ultimo
        else:
            return None, None

    def extract_data_from_page(self, text: str) -> Dict[str, any]:
        """Extrai dados específicos de uma página usando as regras de mapeamento"""
        data = {}
        codes_found = []
        
        # Processa linha por linha
        lines = text.split('\n')
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
            
            # Verifica se a linha contém algum dos códigos procurados
            found_code = None
            for code in self.mapping_rules.keys():
                if code in line:
                    found_code = code
                    break
            
            if found_code:
                codes_found.append(found_code)
                
                # Extrai os dois últimos números da linha
                indice, valor = self.extract_last_two_numbers(line)
                
                rule = self.mapping_rules[found_code]
                logger.debug(f"Processando {found_code}: índice={indice}, valor={valor}")
                
                # Lógica de mapeamento com fallback
                value_to_use = None
                source_used = None
                
                if rule['source'] == 'indice':
                    if indice is not None and indice != 0:
                        value_to_use = indice
                        source_used = 'índice'
                    elif rule.get('fallback_to_valor', False) and valor is not None:
                        # Fallback para PRODUÇÃO (01003601)
                        value_to_use = valor
                        source_used = 'valor (fallback)'
                        logger.debug(f"[FALLBACK] {found_code}: Usando fallback valor pois indice vazio")
                elif rule['source'] == 'valor' and valor is not None:
                    value_to_use = valor
                    source_used = 'valor'
                
                if value_to_use is not None:
                    data[rule['excel_column']] = value_to_use
                    logger.debug(f"[OK] {found_code}: Coluna {rule['excel_column']} = {value_to_use} ({source_used})")
                else:
                    logger.debug(f"[AVISO] {found_code}: Valor não pôde ser extraído")
        
        if codes_found:
            logger.debug(f"Códigos encontrados na página: {codes_found}")
        else:
            logger.debug("Nenhum código conhecido encontrado na página")
            # Mostra algumas linhas para debug se não encontrar nada
            sample_lines = [line for line in lines if line.strip() and not line.startswith(' ')][:3]
            if sample_lines:
                logger.debug(f"Amostra de linhas: {sample_lines}")
        
        return data

    def filter_normal_pages(self, pages_text: List[str]) -> List[str]:
        """Filtra apenas páginas de folha normal (exclui 13º salário e férias)"""
        filtered_pages = []
        
        for i, text in enumerate(pages_text):
            is_normal = True
            page_type_found = False
            
            # Primeiro, procura especificamente pela linha "Tipo da folha:"
            lines = text.split('\n')
            for line in lines:
                line_clean = line.strip()
                
                # Verifica se é uma linha de tipo de folha
                if re.search(r'Tipo\s+da\s+folha\s*:', line_clean, re.IGNORECASE):
                    page_type_found = True
                    logger.debug(f"Página {i+1}: Linha tipo encontrada: '{line_clean}'")
                    
                    # Verifica se é folha normal
                    if re.search(r'FOLHA\s+NORMAL', line_clean, re.IGNORECASE):
                        is_normal = True
                        logger.debug(f"Página {i+1}: FOLHA NORMAL identificada")
                        break
                    # Verifica se é tipo especial (13º, férias, etc.)
                    elif re.search(r'13\s*[°º]?\s*SAL[AÁ]RIO|DECIMOT?ERCEIRO|F[ÉE]RIAS|ADIANTAMENTO|RESCIS[ÃA]O', line_clean, re.IGNORECASE):
                        is_normal = False
                        logger.debug(f"Página {i+1}: Tipo especial identificado: '{line_clean}'")
                        break
            
            # Se não encontrou linha "Tipo da folha:", usa filtro de fallback mais restritivo
            if not page_type_found:
                logger.debug(f"Página {i+1}: Linha 'Tipo da folha' não encontrada, aplicando filtro de fallback")
                
                # Procura por indicadores de folhas especiais no cabeçalho/início da página
                header_text = '\n'.join(lines[:10])  # Apenas primeiras 10 linhas
                
                exclusion_patterns = [
                    r'13\s*[°º]?\s*SAL[AÁ]RIO',
                    r'DECIMOT?ERCEIRO',
                    r'FOLHA\s*DE\s*F[ÉE]RIAS',
                    r'ADIANTAMENTO\s*SALARIAL',
                    r'RESCIS[ÃA]O'
                ]
                
                for pattern in exclusion_patterns:
                    if re.search(pattern, header_text, re.IGNORECASE):
                        is_normal = False
                        logger.debug(f"Página {i+1}: Tipo especial identificado no cabeçalho (fallback)")
                        break
            
            if is_normal:
                filtered_pages.append(text)
                logger.debug(f"Página {i+1}: Aceita como folha normal")
            else:
                logger.debug(f"Página {i+1}: Rejeitada (não é folha normal)")
                
        logger.debug(f"Páginas filtradas: {len(filtered_pages)} de {len(pages_text)} (folhas normais)")
        return filtered_pages

    def find_row_for_period(self, worksheet, month: int, year: int) -> Optional[int]:
        """Encontra a linha correspondente ao período (mês/ano) na planilha"""
        meses_nomes = ['', 'jan', 'fev', 'mar', 'abr', 'mai', 'jun',
                      'jul', 'ago', 'set', 'out', 'nov', 'dez']
        periodo_procurado = f"{meses_nomes[month]}/{str(year)[2:]}"
        
        logger.debug(f"Procurando período: {periodo_procurado} (mês {month}, ano {year})")
        
        found_values = []  # Para debug
        
        # Procura na coluna A (PERÍODO)
        for row_num in range(1, worksheet.max_row + 1):
            cell_value = worksheet[f'A{row_num}'].value
            
            if cell_value is None:
                continue
            
            # Armazena valores encontrados para debug (apenas primeiros 10)
            if len(found_values) < 10:
                found_values.append(f"A{row_num}: {repr(cell_value)}")
                
            # Se for string, tenta formato texto (nov/12, dez/12, etc.)
            if isinstance(cell_value, str):
                cell_str = cell_value.strip()
                
                if cell_str == periodo_procurado:
                    logger.debug(f"[OK] Período {periodo_procurado} encontrado na linha {row_num} (formato texto)")
                    return row_num
            
            # Se for data/datetime, compara mês e ano
            elif isinstance(cell_value, datetime):
                if cell_value.month == month and cell_value.year == year:
                    logger.debug(f"[OK] Período {periodo_procurado} encontrado na linha {row_num} (formato data: {cell_value})")
                    return row_num
            
            # Se for number (serial date do Excel), tenta converter
            elif isinstance(cell_value, (int, float)):
                try:
                    from datetime import timedelta
                    
                    # Excel epoch é 1900-01-01, mas com bug: conta 1900 como bissexto
                    if cell_value > 59:  # Após 28/02/1900
                        excel_date = datetime(1899, 12, 30) + timedelta(days=cell_value)
                    else:
                        excel_date = datetime(1899, 12, 31) + timedelta(days=cell_value)
                    
                    if excel_date.month == month and excel_date.year == year:
                        logger.debug(f"[OK] Período {periodo_procurado} encontrado na linha {row_num} (serial: {cell_value} -> {excel_date})")
                        return row_num
                        
                except Exception as e:
                    logger.debug(f"Erro ao converter serial date {cell_value}: {e}")
                    continue
        
        # Log de debug quando não encontrar
        logger.warning(f"[ERRO] Período {periodo_procurado} não encontrado na planilha")
        if found_values:
            logger.debug(f"Valores encontrados na coluna A: {', '.join(found_values[:5])}")
            if len(found_values) > 5:
                logger.debug(f"... e mais {len(found_values) - 5} valores")
        else:
            logger.debug("Coluna A vazia ou sem dados válidos")
        
        return None

    def update_excel_file(self, excel_path: str, extracted_data: Dict):
        """Atualiza o arquivo Excel existente com os dados extraídos"""
        try:
            # Verifica se é arquivo com macros (.xlsm)
            is_macro_enabled = excel_path.lower().endswith('.xlsm')
            
            # Carrega a planilha existente preservando macros se necessário
            if is_macro_enabled:
                workbook = load_workbook(excel_path, keep_vba=True)
            else:
                workbook = load_workbook(excel_path)
            
            worksheet = None
            
            # Se foi especificada uma planilha, tenta usar ela primeiro
            if self.preferred_sheet:
                if self.preferred_sheet in workbook.sheetnames:
                    worksheet = workbook[self.preferred_sheet]
                    logger.debug(f"Usando planilha especificada: {worksheet.title}")
                else:
                    raise ValueError(f"Planilha especificada '{self.preferred_sheet}' não encontrada")
            else:
                # Regra: "LEVANTAMENTO DADOS" é obrigatória se não especificada
                if 'LEVANTAMENTO DADOS' in workbook.sheetnames:
                    worksheet = workbook['LEVANTAMENTO DADOS']
                    logger.debug(f"Usando planilha padrão: {worksheet.title}")
                else:
                    raise ValueError("Planilha 'LEVANTAMENTO DADOS' não encontrada. Use -s para especificar outra planilha.")
            
            updates_count = 0
            failed_periods = []
            successful_periods = []
            
            # Para cada período com dados extraídos
            for (month, year), data in extracted_data.items():
                meses = ['', 'jan', 'fev', 'mar', 'abr', 'mai', 'jun',
                        'jul', 'ago', 'set', 'out', 'nov', 'dez']
                periodo = f"{meses[month]}/{str(year)[2:]}"
                
                # Encontra a linha correspondente ao período
                row_num = self.find_row_for_period(worksheet, month, year)
                
                if row_num:
                    # Atualiza cada coluna com dados
                    period_updates = 0
                    cells_already_filled = []
                    cells_updated = []
                    
                    for column, value in data.items():
                        cell_address = f"{column}{row_num}"
                        old_value = worksheet[cell_address].value
                        
                        # Só atualiza se a célula estiver vazia ou diferente
                        if old_value is None or old_value == '' or old_value == 0:
                            worksheet[cell_address] = value
                            cells_updated.append(f"{cell_address}:{value}")
                            logger.debug(f"Atualizado {cell_address}: {old_value} -> {value}")
                            updates_count += 1
                            period_updates += 1
                        else:
                            cells_already_filled.append(f"{cell_address}:{old_value}")
                            logger.debug(f"Celula {cell_address} ja preenchida: {old_value}")
                    
                    if period_updates > 0:
                        successful_periods.append(periodo)
                        logger.debug(f"[OK] {periodo}: {period_updates} células atualizadas")
                    else:
                        # Período encontrado mas nenhuma célula foi atualizada
                        failed_periods.append(f"{periodo} (células já preenchidas)")
                        logger.debug(f"[AVISO] {periodo}: todas as células já estavam preenchidas")
                else:
                    # Período não encontrado na planilha
                    failed_periods.append(f"{periodo} (linha não encontrada)")
                    logger.debug(f"[ERRO] {periodo}: linha não encontrada na planilha")
            
            # Salva as alterações preservando formato original
            workbook.save(excel_path)
            
            # Log resumido
            total_periods = len(extracted_data)
            success_periods = len(successful_periods)
            
            if success_periods == total_periods:
                logger.info(f"[OK] Processamento concluído: {success_periods} períodos atualizados")
            else:
                logger.info(f"[AVISO] Processamento concluído: {success_periods}/{total_periods} períodos atualizados")
            
            if failed_periods:
                logger.warning(f"[ERRO] Falhas em {len(failed_periods)} períodos:")
                for failed in failed_periods[:3]:  # Mostra apenas primeiros 3
                    logger.warning(f"   {failed}")
                if len(failed_periods) > 3:
                    logger.warning(f"   ... e mais {len(failed_periods) - 3} períodos")
            
            logger.debug(f"Total de células atualizadas: {updates_count}")
            
        except Exception as e:
            logger.error(f"Erro ao atualizar Excel: {e}")
            raise

    def process_pdf(self, pdf_filename: str):
        """Processa PDF usando diretório de trabalho configurado"""
        logger.debug("Iniciando processamento PDF...")
        
        # Carrega configuração obrigatória
        self.load_env_config()
        
        # Encontra PDF no diretório de trabalho
        pdf_path = self.find_pdf_file(pdf_filename)
        logger.info(f"PDF encontrado: {Path(pdf_path).name}")
        
        # Copia modelo e cria arquivo de destino
        excel_path = self.copy_modelo_to_dados(pdf_path)
        logger.info(f"Modelo copiado para: DADOS/{Path(excel_path).name}")
        
        # Extrai dados do PDF
        pages_text = self.extract_text_from_pdf(pdf_path)
        normal_pages = self.filter_normal_pages(pages_text)
        
        logger.debug(f"PDF processado: {len(pages_text)} páginas totais, {len(normal_pages)} páginas válidas (folha normal)")
        
        # Extrai dados de cada página
        extracted_data = {}
        pages_sem_data = []
        pages_sem_periodo = []
        
        for i, page_text in enumerate(normal_pages):
            logger.debug(f"\n--- Processando página {i+1} ---")
            
            # Extrai data de referência
            date_ref = self.extract_reference_date(page_text)
            if not date_ref:
                pages_sem_periodo.append(i+1)
                logger.debug(f"[AVISO] Página {i+1}: Data de referência não encontrada")
                continue
            
            mes, ano = date_ref
            meses_nomes = ['', 'jan', 'fev', 'mar', 'abr', 'mai', 'jun',
                          'jul', 'ago', 'set', 'out', 'nov', 'dez']
            periodo_str = f"{meses_nomes[mes]}/{str(ano)[2:]}"
            logger.debug(f"Data: Página {i+1}: Período identificado como {periodo_str}")
            
            # Extrai dados da página
            page_data = self.extract_data_from_page(page_text)
            
            if page_data:
                extracted_data[date_ref] = page_data
                codes_found = [code for code in self.mapping_rules.keys() 
                              if any(self.mapping_rules[code]['excel_column'] == col for col in page_data.keys())]
                logger.debug(f"[OK] Página {i+1}: {len(page_data)} campos extraídos {list(page_data.keys())}")
                logger.debug(f"  Códigos encontrados: {codes_found}")
            else:
                pages_sem_data.append((i+1, periodo_str))
                logger.debug(f"[AVISO] Página {i+1}: Período {periodo_str} identificado, mas nenhum código mapeado encontrado")
        
        # Resumo do processamento das páginas
        if pages_sem_periodo:
            logger.info(f"[AVISO] {len(pages_sem_periodo)} páginas sem período identificado: {pages_sem_periodo}")
        
        if pages_sem_data:
            periodos_sem_data = [periodo for _, periodo in pages_sem_data]
            logger.info(f"[AVISO] {len(pages_sem_data)} páginas com período mas sem dados: {periodos_sem_data}")
        
        # Atualiza Excel
        if extracted_data:
            logger.debug(f"\n--- Atualizando Excel ---")
            logger.debug(f"Períodos a processar: {len(extracted_data)}")
            self.update_excel_file(excel_path, extracted_data)
        else:
            logger.warning("ERRO: Nenhum dado foi extraído do PDF!")
            if pages_sem_periodo and pages_sem_data:
                logger.warning("Possíveis causas: formato de data não reconhecido ou códigos não encontrados")
        
        return extracted_data, excel_path

def main():
    """Função principal da aplicação"""
    parser = argparse.ArgumentParser(description='PDF para Excel Updater v3.0 - Diretório de Trabalho')
    parser.add_argument('pdf_filename', help='Nome do arquivo PDF (será buscado no diretório de trabalho)')
    parser.add_argument('-s', '--sheet', help='Nome da planilha específica (padrão: "LEVANTAMENTO DADOS")')
    parser.add_argument('-v', '--verbose', action='store_true', help='Modo verboso')
    
    args = parser.parse_args()
    
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)
    else:
        # Modo silencioso - apenas INFO e WARNING/ERROR
        logging.getLogger().setLevel(logging.INFO)
    
    try:
        # Cria updater e processa
        updater = PDFToExcelUpdater()
        
        # Se especificou planilha, usa diretamente
        if args.sheet:
            updater.preferred_sheet = args.sheet
        
        print(f"Processando: {args.pdf_filename}")
        
        extracted_data, excel_path = updater.process_pdf(args.pdf_filename)
        
        # Resumo final
        if extracted_data:
            total_extracted = len(extracted_data)
            print(f"OK: Concluído: {total_extracted} períodos processados")
            excel_name = Path(excel_path).name
            print(f"Arquivo criado: DADOS/{excel_name}")
        else:
            print("ERRO: Nenhum dado foi extraído do PDF!")
            print("Dica: Use -v para diagnóstico detalhado")
            return 1
        
        return 0
        
    except ValueError as e:
        print(f"ERRO: {e}")
        return 1
    except Exception as e:
        print(f"ERRO: Erro inesperado: {e}")
        if args.verbose:
            raise
        return 1

if __name__ == "__main__":
    import sys
    sys.exit(main())