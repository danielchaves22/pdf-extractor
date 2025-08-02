#!/usr/bin/env python3
"""
Extrator de Dados de Folha de Pagamento - PDF
=============================================

Esta aplicação extrai dados de PDFs de folha de pagamento e gera planilhas
seguindo regras específicas de mapeamento.

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

# Configuração de logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class PDFDataExtractor:
    """Classe principal para extração de dados de PDF"""
    
    def __init__(self):
        """Inicializa o extrator com as regras de mapeamento"""
        
        # Regras de mapeamento conforme especificação
        self.mapping_rules = {
            # Obter da coluna ÍNDICE
            '01003601': {'code': 'PREMIO PROD. MENSAL', 'target': 'PRODUÇÃO', 'source': 'indice'},
            '01007301': {'code': 'HORAS EXT.100%-180', 'target': 'INDICE HE 100%', 'source': 'indice'},
            '01009001': {'code': 'ADIC.NOT.25%-180', 'target': 'INDICE ADC. NOT.', 'source': 'indice'},
            '01003501': {'code': 'HORAS EXT.75%-180', 'target': 'INDICE HE 75%', 'source': 'indice'},
            
            # Obter da coluna VALOR
            '09090301': {'code': 'SALARIO CONTRIB INSS', 'target': 'REMUNERAÇÃO RECEBIDA', 'source': 'valor'}
        }
        
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

    def extract_text_from_pdf(self, pdf_path: str) -> List[str]:
        """
        Extrai texto de todas as páginas do PDF
        
        Args:
            pdf_path: Caminho para o arquivo PDF
            
        Returns:
            Lista com o texto de cada página
        """
        pages_text = []
        
        try:
            with pdfplumber.open(pdf_path) as pdf:
                logger.info(f"Processando PDF: {pdf_path} ({len(pdf.pages)} páginas)")
                
                for i, page in enumerate(pdf.pages):
                    text = page.extract_text()
                    if text:
                        pages_text.append(text)
                        logger.debug(f"Página {i+1} processada")
                    
        except Exception as e:
            logger.error(f"Erro ao processar PDF: {e}")
            raise
            
        return pages_text

    def extract_reference_date(self, text: str) -> Optional[Tuple[int, int]]:
        """
        Extrai a data de referência da página (mês/ano)
        
        Args:
            text: Texto da página
            
        Returns:
            Tupla (mês, ano) ou None se não encontrado
        """
        # Padrões para capturar data de referência
        patterns = [
            r'Referência:\s*(\w+)/(\d{4})',
            r'Referencia:\s*(\w+)/(\d{4})',
            r'Data do cálculo:\s*\d{2}/(\d{2})/(\d{4})',
            r'(\w+)/(\d{4})',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                if pattern.count('(') == 2:  # Tem mês e ano
                    mes_str = match.group(1).lower()
                    ano = int(match.group(2))
                    
                    # Tenta converter mês
                    mes = self.meses_pt.get(mes_str) or self.meses_abrev.get(mes_str)
                    if mes:
                        return (mes, ano)
                else:  # Só tem mês numérico e ano
                    mes = int(match.group(1))
                    ano = int(match.group(2))
                    return (mes, ano)
        
        return None

    def extract_data_from_page(self, text: str) -> Dict[str, float]:
        """
        Extrai dados específicos de uma página usando as regras de mapeamento
        
        Args:
            text: Texto da página
            
        Returns:
            Dicionário com os dados extraídos
        """
        data = {}
        
        # Padrão para capturar linhas com códigos e valores
        # Formato esperado: P/D código descrição índice valor
        pattern = r'([PD])\s+(\d+)\s+([^0-9]+?)\s+([\d.,]+)?\s+([\d.,]+)'
        
        matches = re.findall(pattern, text)
        
        for match in matches:
            tipo, codigo, descricao, indice_str, valor_str = match
            
            # Verifica se o código está nas regras
            if codigo in self.mapping_rules:
                rule = self.mapping_rules[codigo]
                
                try:
                    # Converte valores para float
                    indice = self._convert_to_float(indice_str) if indice_str else None
                    valor = self._convert_to_float(valor_str) if valor_str else None
                    
                    # Determina qual valor usar baseado na regra
                    if rule['source'] == 'indice' and indice is not None:
                        data[rule['target']] = indice
                    elif rule['source'] == 'valor' and valor is not None:
                        data[rule['target']] = valor
                    
                    logger.debug(f"Extraído {codigo}: {rule['target']} = {data.get(rule['target'])}")
                    
                except ValueError as e:
                    logger.warning(f"Erro ao converter valores para {codigo}: {e}")
        
        return data

    def _convert_to_float(self, value_str: str) -> float:
        """
        Converte string com formato brasileiro para float
        
        Args:
            value_str: String com o valor
            
        Returns:
            Valor convertido para float
        """
        if not value_str:
            return 0.0
            
        # Remove espaços e converte vírgula para ponto
        cleaned = value_str.strip().replace('.', '').replace(',', '.')
        return float(cleaned)

    def filter_normal_pages(self, pages_text: List[str]) -> List[str]:
        """
        Filtra apenas páginas de folha normal (exclui 13º salário e férias)
        
        Args:
            pages_text: Lista com texto das páginas
            
        Returns:
            Lista filtrada com apenas folhas normais
        """
        filtered_pages = []
        
        exclusion_patterns = [
            r'13\s*SALARIO',
            r'FERIAS',
            r'FOLHA\s*DE\s*FERIAS',
            r'DECIMO\s*TERCEIRO'
        ]
        
        for text in pages_text:
            is_normal = True
            
            for pattern in exclusion_patterns:
                if re.search(pattern, text, re.IGNORECASE):
                    is_normal = False
                    break
            
            if is_normal:
                filtered_pages.append(text)
                
        logger.info(f"Páginas filtradas: {len(filtered_pages)} de {len(pages_text)}")
        return filtered_pages

    def generate_complete_table(self, extracted_data: Dict[Tuple[int, int], Dict[str, float]], 
                              start_date: Tuple[int, int] = (11, 2012), 
                              end_date: Tuple[int, int] = (11, 2017)) -> pd.DataFrame:
        """
        Gera tabela completa com todos os meses no período especificado
        
        Args:
            extracted_data: Dados extraídos por (mês, ano)
            start_date: Data inicial (mês, ano)
            end_date: Data final (mês, ano)
            
        Returns:
            DataFrame completo
        """
        # Gera lista completa de períodos
        periods = []
        current_month, current_year = start_date
        
        while (current_year, current_month) <= (end_date[1], end_date[0]):
            periods.append((current_month, current_year))
            
            current_month += 1
            if current_month > 12:
                current_month = 1
                current_year += 1
        
        # Cria DataFrame
        columns = [
            'PERÍODO', 'REMUNERAÇÃO RECEBIDA', 'PRODUÇÃO', 'INDICE HE 100%', 'FORMULA_1',
            'INDICE HE 75%', 'FORMULA_2', 'INDICE ADC. NOT.', 'FORMULA_3'
        ]
        
        df_data = []
        
        for i, (mes, ano) in enumerate(periods):
            # Formata período
            meses_nomes = ['', 'jan', 'fev', 'mar', 'abr', 'mai', 'jun',
                          'jul', 'ago', 'set', 'out', 'nov', 'dez']
            periodo = f"{meses_nomes[mes]}/{str(ano)[2:]}"
            
            # Busca dados extraídos para este período
            data = extracted_data.get((mes, ano), {})
            
            # Linha da tabela (baseada na linha 5 sendo a primeira)
            linha = i + 5
            
            row = [
                periodo,
                data.get('REMUNERAÇÃO RECEBIDA', ''),
                data.get('PRODUÇÃO', ''),
                data.get('INDICE HE 100%', ''),
                f'=Y{linha}/100',
                data.get('INDICE HE 75%', ''),
                f'=AC{linha}/10000',
                data.get('INDICE ADC. NOT.', ''),
                f'=AC{linha}/10000'
            ]
            
            df_data.append(row)
        
        df = pd.DataFrame(df_data, columns=columns)
        return df

    def process_pdf(self, pdf_path: str, output_path: str = None) -> pd.DataFrame:
        """
        Processa um PDF completo e gera a planilha final
        
        Args:
            pdf_path: Caminho do PDF de entrada
            output_path: Caminho do arquivo de saída (opcional)
            
        Returns:
            DataFrame com os dados processados
        """
        logger.info("Iniciando processamento do PDF...")
        
        # Extrai texto do PDF
        pages_text = self.extract_text_from_pdf(pdf_path)
        
        # Filtra páginas normais
        normal_pages = self.filter_normal_pages(pages_text)
        
        # Extrai dados de cada página
        extracted_data = {}
        
        for page_text in normal_pages:
            # Extrai data de referência
            date_ref = self.extract_reference_date(page_text)
            if not date_ref:
                logger.warning("Data de referência não encontrada em uma página")
                continue
            
            # Extrai dados da página
            page_data = self.extract_data_from_page(page_text)
            
            if page_data:
                extracted_data[date_ref] = page_data
                logger.info(f"Dados extraídos para {date_ref[0]:02d}/{date_ref[1]}")
        
        # Gera tabela completa
        df = self.generate_complete_table(extracted_data)
        
        # Salva resultado se especificado
        if output_path:
            self.save_results(df, output_path)
        
        logger.info(f"Processamento concluído! {len(extracted_data)} períodos processados.")
        return df

    def save_results(self, df: pd.DataFrame, output_path: str):
        """
        Salva os resultados em arquivo Excel ou CSV
        
        Args:
            df: DataFrame com os dados
            output_path: Caminho do arquivo de saída
        """
        output_path = Path(output_path)
        
        if output_path.suffix.lower() == '.xlsx':
            # Salva como Excel com formatação
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Levantamento_Dados', index=False)
                
                # Aplica formatação básica
                worksheet = writer.sheets['Levantamento_Dados']
                
                # Ajusta largura das colunas
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width
                    
        elif output_path.suffix.lower() == '.csv':
            df.to_csv(output_path, index=False, encoding='utf-8-sig')
        else:
            raise ValueError("Formato de saída deve ser .xlsx ou .csv")
        
        logger.info(f"Resultados salvos em: {output_path}")

def main():
    """Função principal da aplicação"""
    parser = argparse.ArgumentParser(description='Extrator de Dados de Folha de Pagamento')
    parser.add_argument('pdf_path', help='Caminho para o arquivo PDF')
    parser.add_argument('-o', '--output', help='Arquivo de saída (.xlsx ou .csv)')
    parser.add_argument('-v', '--verbose', action='store_true', help='Modo verboso')
    
    args = parser.parse_args()
    
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)
    
    # Valida arquivo de entrada
    if not Path(args.pdf_path).exists():
        logger.error(f"Arquivo não encontrado: {args.pdf_path}")
        return
    
    # Define arquivo de saída se não especificado
    if not args.output:
        input_path = Path(args.pdf_path)
        args.output = input_path.parent / f"{input_path.stem}_extracted.xlsx"
    
    try:
        # Processa o PDF
        extractor = PDFDataExtractor()
        df = extractor.process_pdf(args.pdf_path, args.output)
        
        # Mostra resumo
        print(f"\n{'='*50}")
        print("RESUMO DO PROCESSAMENTO")
        print(f"{'='*50}")
        print(f"Total de períodos: {len(df)}")
        print(f"Períodos com dados: {len(df[df['REMUNERAÇÃO RECEBIDA'] != ''])}")
        print(f"Arquivo gerado: {args.output}")
        print(f"{'='*50}")
        
    except Exception as e:
        logger.error(f"Erro durante o processamento: {e}")
        raise

if __name__ == "__main__":
    main()
