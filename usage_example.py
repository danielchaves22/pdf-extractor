#!/usr/bin/env python3
"""
Exemplo de uso do Extrator de Dados PDF
=======================================

Este script demonstra como usar a aplicação tanto programaticamente
quanto pela linha de comando.
"""

from pdf_extractor import PDFDataExtractor
import pandas as pd

def exemplo_programatico():
    """Exemplo de uso programático da aplicação"""
    
    print("=== EXEMPLO PROGRAMÁTICO ===")
    
    # Cria instância do extrator
    extractor = PDFDataExtractor()
    
    # Processa o PDF
    pdf_path = "seu_arquivo.pdf"  # Substitua pelo caminho do seu PDF
    
    try:
        # Extrai dados e gera planilha
        df = extractor.process_pdf(pdf_path, "resultado_programatico.xlsx")
        
        # Mostra primeiras linhas
        print("\nPrimeiras 10 linhas:")
        print(df.head(10))
        
        # Estatísticas básicas
        print(f"\nEstatísticas:")
        print(f"Total de períodos: {len(df)}")
        
        # Conta períodos com dados
        periodos_com_dados = len(df[df['REMUNERAÇÃO RECEBIDA'] != ''])
        print(f"Períodos com dados: {periodos_com_dados}")
        
        # Mostra faixa de valores
        valores_nao_vazios = df[df['REMUNERAÇÃO RECEBIDA'] != '']['REMUNERAÇÃO RECEBIDA']
        if len(valores_nao_vazios) > 0:
            valores_numericos = pd.to_numeric(valores_nao_vazios, errors='coerce')
            print(f"Menor remuneração: R$ {valores_numericos.min():,.2f}")
            print(f"Maior remuneração: R$ {valores_numericos.max():,.2f}")
            print(f"Média remuneração: R$ {valores_numericos.mean():,.2f}")
        
    except FileNotFoundError:
        print(f"❌ Arquivo não encontrado: {pdf_path}")
        print("💡 Dica: Substitua 'seu_arquivo.pdf' pelo caminho correto")
    except Exception as e:
        print(f"❌ Erro: {e}")

def exemplo_linha_comando():
    """Mostra exemplos de uso via linha de comando"""
    
    print("\n=== EXEMPLOS LINHA DE COMANDO ===")
    print()
    print("1. Uso básico:")
    print("   python pdf_extractor.py meu_arquivo.pdf")
    print()
    print("2. Especificando arquivo de saída:")
    print("   python pdf_extractor.py meu_arquivo.pdf -o resultado.xlsx")
    print()
    print("3. Modo verboso (mostra mais detalhes):")
    print("   python pdf_extractor.py meu_arquivo.pdf -v")
    print()
    print("4. Gerando CSV ao invés de Excel:")
    print("   python pdf_extractor.py meu_arquivo.pdf -o resultado.csv")
    print()
    print("5. Exemplo completo:")
    print("   python pdf_extractor.py folha_pagamento.pdf -o dados_extraidos.xlsx -v")

def validar_ambiente():
    """Valida se o ambiente está configurado corretamente"""
    
    print("=== VALIDAÇÃO DO AMBIENTE ===")
    
    try:
        import pandas
        print(f"✅ pandas {pandas.__version__}")
    except ImportError:
        print("❌ pandas não instalado")
    
    try:
        import pdfplumber
        print(f"✅ pdfplumber {pdfplumber.__version__}")
    except ImportError:
        print("❌ pdfplumber não instalado")
    
    try:
        import openpyxl
        print(f"✅ openpyxl {openpyxl.__version__}")
    except ImportError:
        print("❌ openpyxl não instalado")
    
    print("\n💡 Para instalar dependências: pip install -r requirements.txt")

if __name__ == "__main__":
    print("🔧 EXTRATOR DE DADOS PDF - EXEMPLO DE USO")
    print("=" * 50)
    
    # Valida ambiente
    validar_ambiente()
    
    # Mostra exemplos
    exemplo_linha_comando()
    
    # Tenta executar exemplo programático
    exemplo_programatico()
    
    print("\n" + "=" * 50)
    print("✨ Exemplo concluído!")
