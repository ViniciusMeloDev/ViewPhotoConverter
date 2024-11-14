import pytesseract
from PIL import Image
import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox

# Configuração do Tesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\msys64\mingw64\bin\tesseract.exe'

def extrair_texto(caminho_imagem):
    """Extrai o texto de uma imagem usando o Tesseract OCR e organiza em linhas."""
    try:
        imagem = Image.open(caminho_imagem)
        texto = pytesseract.image_to_string(imagem)
        return texto.strip().split("\n")  # Cada linha em uma nova linha da lista
    except FileNotFoundError:
        print(f"Erro: A imagem '{caminho_imagem}' não foi encontrada.")
        return None
    except Exception as e:
        print(f"Erro ao processar a imagem {caminho_imagem}: {e}")
        return None

def processar_imagens_diretorio(diretorio_imagens):
    """Processa todas as imagens em um diretório e organiza os textos extraídos em formato de planilha."""
    dados = []
    
    for arquivo in os.listdir(diretorio_imagens):
        caminho_arquivo = os.path.join(diretorio_imagens, arquivo)
        
        if caminho_arquivo.lower().endswith(('.png', '.jpg', '.jpeg', '.tiff', '.bmp')):
            print(f"Processando imagem: {arquivo}")
            linhas_texto = extrair_texto(caminho_arquivo)
            if linhas_texto:
                for linha in linhas_texto:
                    # Organizar os dados conforme o layout
                    colunas = linha.split()  # Dividir conforme a estrutura dos dados extraídos
                    dados.append({
                        "Horários": colunas[0] if len(colunas) > 0 else "",
                        "Dias": colunas[1] if len(colunas) > 1 else "",
                        "Atividade / Professor": colunas[2] if len(colunas) > 2 else "",
                        "Idade / Morador": colunas[3] if len(colunas) > 3 else "",
                        "Frequência - Dias do Mês": " ".join(colunas[4:-1]),  # Ajuste conforme necessário
                        "Total do Mês": colunas[-1] if len(colunas) > 4 else ""
                    })
    
    return dados

def salvar_em_excel(dados, nome_arquivo="relatorio_texto_imagens.xlsx"):
    """Salva os dados extraídos em um arquivo Excel com formatação específica, baseada no layout do PDF."""
    df = pd.DataFrame(dados)
    
    # Salvando com formatação de estilo no Excel
    with pd.ExcelWriter(nome_arquivo, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Frequências Junho 2024")
        
        # Obter o workbook e worksheet para formatação
        workbook  = writer.book
        worksheet = writer.sheets["Frequências Junho 2024"]
        
        # Ajustar largura das colunas e aplicar estilos
        worksheet.set_column("A:A", 10)  # Horários
        worksheet.set_column("B:B", 15)  # Dias
        worksheet.set_column("C:C", 30)  # Atividade / Professor
        worksheet.set_column("D:D", 20)  # Idade / Morador
        worksheet.set_column("E:E", 40)  # Frequência - Dias do Mês
        worksheet.set_column("F:F", 15)  # Total do Mês
        
        # Estilo para o cabeçalho
        header_format = workbook.add_format({
            "bold": True,
            "bg_color": "#D9EAD3",
            "border": 1
        })
        
        # Estilo para as células
        cell_format = workbook.add_format({
            "text_wrap": True,
            "border": 1
        })
        
        # Aplicar o estilo do cabeçalho e do corpo
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        worksheet.set_column("A:F", None, cell_format)

    print(f"Relatório salvo como '{nome_arquivo}'")

def selecionar_diretorio():
    """Abre uma janela para selecionar o diretório de imagens."""
    diretorio = filedialog.askdirectory()
    if diretorio:
        entrada_diretorio.delete(0, tk.END)
        entrada_diretorio.insert(0, diretorio)

def gerar_relatorio():
    """Gera o relatório de textos extraídos das imagens no diretório especificado."""
    diretorio_imagens = entrada_diretorio.get()
    nome_arquivo = entrada_nome_arquivo.get() or "relatorio_texto_imagens.xlsx"
    
    if not diretorio_imagens:
        messagebox.showerror("Erro", "Por favor, selecione o diretório de imagens.")
        return
    
    # Extrai o texto das imagens no diretório
    dados_imagens = processar_imagens_diretorio(diretorio_imagens)
    
    # Salva os dados extraídos em uma planilha Excel
    if dados_imagens:
        salvar_em_excel(dados_imagens, nome_arquivo)
        messagebox.showinfo("Sucesso", f"Relatório salvo como '{nome_arquivo}'")
    else:
        messagebox.showwarning("Atenção", "Nenhum texto extraído das imagens.")

# Configuração da interface gráfica
app = tk.Tk()
app.title("Extrator de Texto para Relatório")
app.geometry("400x200")

# Rótulo e campo para o diretório de imagens
tk.Label(app, text="Diretório de Imagens:").pack(pady=5)
entrada_diretorio = tk.Entry(app, width=50)
entrada_diretorio.pack(pady=5)
botao_diretorio = tk.Button(app, text="Selecionar Diretório", command=selecionar_diretorio)
botao_diretorio.pack(pady=5)

# Rótulo e campo para o nome do arquivo Excel
tk.Label(app, text="Nome do Arquivo Excel:").pack(pady=5)
entrada_nome_arquivo = tk.Entry(app, width=50)
entrada_nome_arquivo.insert(0, "relatorio_texto_imagens.xlsx")
entrada_nome_arquivo.pack(pady=5)

# Botão para gerar o relatório
botao_gerar = tk.Button(app, text="Gerar Relatório", command=gerar_relatorio)
botao_gerar.pack(pady=10)

# Inicia o loop da interface
app.mainloop()
