#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF para Excel Updater - Preenche Excel Existente (.xlsx/.xlsm)
================================================================

Esta versão detecta um arquivo Excel com mesmo nome do PDF e preenche
diretamente nas colunas corretas, preservando formatação, fórmulas e macros.

Versão 2.1 - Conforme regras de negócio

Estrutura da planilha:
- Coluna A: PERÍODO (não preenche)
- Coluna B: REMUNERAÇÃO RECEBIDA  
- Coluna X: PRODUÇÃO
- Coluna Y: INDICE HE 100%
- Coluna AA: INDICE HE 75% 
- Coluna AC: INDICE ADC. NOT.

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

# Configuração de logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class PDFToExcelUpdater:
    """Classe para atualizar Excel existente com dados do PDF"""
    
    def __init__(self):
        """Inicializa o updater com as regras de mapeamento"""
        
        # Regras de mapeamento para colunas específicas do Excel (v2.1)
        self.mapping_rules = {
            # Obter da coluna ÍNDICE (com fallback especial para PRODUÇÃO)
            '01003601': {'code': 'PREMIO PROD. MENSAL', 'excel_column': 'X', 'source': 'indice', 'fallback_to_valor': True},
            '01007301': {'code': 'HORAS EXT.100%-180', 'excel_column': 'Y', 'source': 'indice'},
            '01009001': {'code': 'ADIC.NOT.25%-180', 'excel_column': 'AC', 'source': 'indice'},
            '01003501': {'code': 'HORAS EXT.75%-180', 'excel_column': 'AA', 'source': 'indice'},
            '02007501': {'code': 'DIFER.PROV. HORAS EXTRAS 75%', 'excel_column': 'AA', 'source': 'indice'},  # NOVO
            
            # Obter da coluna VALOR
            '09090301': {'code': 'SALARIO CONTRIB INSS', 'excel_column': 'B', 'source': 'valor'}
        }
        
        # Planilha preferida (pode ser especificada externamente)
        self.preferred_sheet = None
        
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

    def find_excel_file(self, pdf_path: str) -> Optional[str]:
        """Procura arquivo Excel com mesmo nome do PDF"""
        pdf_path_obj = Path(pdf_path)
        base_name = pdf_path_obj.stem
        directory = pdf_path_obj.parent
        
        # Procura por diferentes extensões (incluindo .xlsm para macros)
        possible_extensions = ['.xlsm', '.xlsx', '.xls']
        
        for ext in possible_extensions:
            excel_path = directory / f"{base_name}{ext}"
            if excel_path.exists():
                logger.debug(f"Arquivo Excel encontrado: {excel_path}")
                return str(excel_path)
        
        return None

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
        
        for pattern in patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            for match in matches:
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
                                    return (mes, ano)
                            except ValueError:
                                continue
                        else:
                            return (mes, ano)
                except ValueError:
                    continue
        
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
        
        return data

    def filter_normal_pages(self, pages_text: List[str]) -> List[str]:
        """Filtra apenas páginas de folha normal (exclui 13º salário e férias)"""
        filtered_pages = []
        
        exclusion_patterns = [
            r'13\s*[°º]?\s*SAL[AÁ]RIO',
            r'DECIMOT?ERCEIRO',
            r'F[ÉE]RIAS',
            r'FOLHA\s*DE\s*F[ÉE]RIAS',
            r'ADIANTAMENTO',
            r'RESCIS[ÃA]O'
        ]
        
        for i, text in enumerate(pages_text):
            is_normal = True
            
            for pattern in exclusion_patterns:
                if re.search(pattern, text, re.IGNORECASE):
                    is_normal = False
                    break
            
            if is_normal:
                filtered_pages.append(text)
                
        logger.debug(f"Páginas filtradas: {len(filtered_pages)} de {len(pages_text)}")
        return filtered_pages

    def find_row_for_period(self, worksheet, month: int, year: int) -> Optional[int]:
        """Encontra a linha correspondente ao período (mês/ano) na planilha"""
        logger.debug(f"Procurando período: mês {month}, ano {year}")
        
        # Procura na coluna A (PERÍODO)
        for row_num in range(1, worksheet.max_row + 1):
            cell_value = worksheet[f'A{row_num}'].value
            
            if cell_value is None:
                continue
                
            # Se for string, tenta formato texto (nov/12, dez/12, etc.)
            if isinstance(cell_value, str):
                cell_str = cell_value.strip()
                
                # Formato esperado: nov/12, dez/12, jan/13, etc.
                meses_nomes = ['', 'jan', 'fev', 'mar', 'abr', 'mai', 'jun',
                              'jul', 'ago', 'set', 'out', 'nov', 'dez']
                periodo_procurado = f"{meses_nomes[month]}/{str(year)[2:]}"
                
                if cell_str == periodo_procurado:
                    logger.debug(f"Período {periodo_procurado} encontrado na linha {row_num} (texto)")
                    return row_num
            
            # Se for data/datetime, compara mês e ano
            elif isinstance(cell_value, datetime):
                if cell_value.month == month and cell_value.year == year:
                    logger.debug(f"Período {month}/{year} encontrado na linha {row_num} (data: {cell_value})")
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
                        logger.debug(f"Período {month}/{year} encontrado na linha {row_num} (serial: {cell_value} -> {excel_date})")
                        return row_num
                        
                except Exception as e:
                    logger.debug(f"Erro ao converter serial date {cell_value}: {e}")
                    continue
        
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
                # NOVA REGRA: "LEVANTAMENTO DADOS" é obrigatória se não especificada
                if 'LEVANTAMENTO DADOS' in workbook.sheetnames:
                    worksheet = workbook['LEVANTAMENTO DADOS']
                    logger.debug(f"Usando planilha padrão: {worksheet.title}")
                else:
                    raise ValueError("Planilha 'LEVANTAMENTO DADOS' não encontrada. Use -s para especificar outra planilha.")
            
            updates_count = 0
            failed_periods = []
            
            # Para cada período com dados extraídos
            for (month, year), data in extracted_data.items():
                # Encontra a linha correspondente ao período
                row_num = self.find_row_for_period(worksheet, month, year)
                
                if row_num:
                    # Atualiza cada coluna com dados
                    period_updates = 0
                    for column, value in data.items():
                        cell_address = f"{column}{row_num}"
                        old_value = worksheet[cell_address].value
                        
                        # Só atualiza se a célula estiver vazia ou diferente
                        if old_value is None or old_value == '' or old_value == 0:
                            worksheet[cell_address] = value
                            logger.debug(f"Atualizado {cell_address}: {old_value} -> {value}")
                            updates_count += 1
                            period_updates += 1
                        else:
                            logger.debug(f"Celula {cell_address} ja preenchida: {old_value}")
                    
                    if period_updates == 0:
                        # Período encontrado mas nenhuma célula foi atualizada
                        meses = ['', 'jan', 'fev', 'mar', 'abr', 'mai', 'jun',
                                'jul', 'ago', 'set', 'out', 'nov', 'dez']
                        periodo = f"{meses[month]}/{str(year)[2:]}"
                        failed_periods.append(f"{periodo} (células já preenchidas)")
                else:
                    # Período não encontrado na planilha
                    meses = ['', 'jan', 'fev', 'mar', 'abr', 'mai', 'jun',
                            'jul', 'ago', 'set', 'out', 'nov', 'dez']
                    periodo = f"{meses[month]}/{str(year)[2:]}"
                    failed_periods.append(f"{periodo} (linha não encontrada)")
            
            # Salva as alterações preservando formato original
            workbook.save(excel_path)
            
            # Log resumido
            total_periods = len(extracted_data)
            success_periods = total_periods - len(failed_periods)
            
            logger.info(f"[OK] Processamento concluido: {success_periods}/{total_periods} periodos atualizados")
            
            if failed_periods:
                logger.warning(f"[AVISO] Falhas em {len(failed_periods)} periodos:")
                for failed in failed_periods[:5]:  # Mostra apenas primeiros 5
                    logger.warning(f"   {failed}")
                if len(failed_periods) > 5:
                    logger.warning(f"   ... e mais {len(failed_periods) - 5} periodos")
            
        except Exception as e:
            logger.error(f"Erro ao atualizar Excel: {e}")
            raise

    def process_pdf_to_excel(self, pdf_path: str, excel_path: str = None):
        """Processa PDF e atualiza Excel existente"""
        logger.debug("Iniciando processamento PDF para Excel...")
        
        # Procura arquivo Excel se não especificado
        if not excel_path:
            excel_path = self.find_excel_file(pdf_path)
            if not excel_path:
                raise ValueError(f"Arquivo Excel não encontrado para: {Path(pdf_path).stem}")
        
        # Extrai dados do PDF
        pages_text = self.extract_text_from_pdf(pdf_path)
        normal_pages = self.filter_normal_pages(pages_text)
        
        # Extrai dados de cada página
        extracted_data = {}
        
        for i, page_text in enumerate(normal_pages):
            logger.debug(f"Processando página {i+1}...")
            
            # Extrai data de referência
            date_ref = self.extract_reference_date(page_text)
            if not date_ref:
                logger.debug("Data de referência não encontrada nesta página")
                continue
            
            logger.debug(f"Data encontrada: {date_ref[0]:02d}/{date_ref[1]}")
            
            # Extrai dados da página
            page_data = self.extract_data_from_page(page_text)
            
            if page_data:
                extracted_data[date_ref] = page_data
                logger.debug(f"[OK] Dados extraidos para {date_ref[0]:02d}/{date_ref[1]}: {len(page_data)} campos")
            else:
                logger.debug(f"[AVISO] Nenhum dado extraido para {date_ref[0]:02d}/{date_ref[1]}")
        
        # Atualiza Excel existente
        if extracted_data:
            self.update_excel_file(excel_path, extracted_data)
        else:
            logger.warning("Nenhum dado foi extraído do PDF!")
        
        return extracted_data

def main():
    """Função principal da aplicação"""
    parser = argparse.ArgumentParser(description='PDF para Excel Updater v2.1')
    parser.add_argument('pdf_path', help='Caminho para o arquivo PDF')
    parser.add_argument('-e', '--excel', help='Caminho do arquivo Excel (opcional)')
    parser.add_argument('-s', '--sheet', help='Nome da planilha específica (padrão: "LEVANTAMENTO DADOS")')
    parser.add_argument('-v', '--verbose', action='store_true', help='Modo verboso')
    
    args = parser.parse_args()
    
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)
    else:
        # Modo silencioso - apenas INFO e WARNING/ERROR
        logging.getLogger().setLevel(logging.INFO)
    
    # Valida arquivo de entrada
    if not Path(args.pdf_path).exists():
        print(f"[ERRO] Arquivo PDF nao encontrado: {args.pdf_path}")
        return 1
    
    try:
        # Cria updater e processa
        updater = PDFToExcelUpdater()
        
        # Se especificou planilha, usa diretamente
        if args.sheet:
            updater.preferred_sheet = args.sheet
        
        print(f"Processando: {Path(args.pdf_path).name}")
        
        extracted_data = updater.process_pdf_to_excel(args.pdf_path, args.excel)
        
        # Detecta tipo de arquivo Excel usado
        excel_used = args.excel if args.excel else updater.find_excel_file(args.pdf_path)
        
        # Resumo final simplificado
        if extracted_data:
            print(f"[OK] Concluido: {len(extracted_data)} periodos processados")
            if excel_used:
                excel_name = Path(excel_used).name
                if excel_used.lower().endswith('.xlsm'):
                    print(f"Arquivo: {excel_name} (macros preservados)")
                else:
                    print(f"Arquivo: {excel_name}")
        else:
            print("[ERRO] Nenhum dado foi extraido do PDF!")
            return 1
        
        return 0
        
    except ValueError as e:
        print(f"[ERRO] {e}")
        return 1
    except Exception as e:
        print(f"[ERRO] Erro inesperado: {e}")
        if args.verbose:
            raise
        return 1

if __name__ == "__main__":
    import sys
    sys.exit(main())