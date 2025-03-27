import tkinter as tk
from tkinter import ttk, messagebox, Scrollbar
import openpyxl
import socket
import time
import qrcode
import io
import base64

def ler_arquivo_excel(caminho_arquivo):
    wb = openpyxl.load_workbook(caminho_arquivo)
    planilha = wb.active
    dados = []
    for linha in planilha.iter_rows(min_row=2, values_only=True):
        dados.append(linha)
    return dados

def imprimir_zebra(printer_ip, printer_port, zpl_code):
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.connect((printer_ip, printer_port))
        s.sendall(zpl_code.encode())
        s.close()
        return True
    except Exception as e:
        print(f"Erro ao imprimir: {e}")
        return False

def gerar_etiqueta(codigo_inicial, codigo_final, printer_ip, printer_port, dados):
    for codigo in range(codigo_inicial, codigo_final + 1):
        codigo_formatado = f"EVC{codigo:04d}"
        
        # Encontrar dados correspondentes no Excel
        dados_etiqueta = None
        for dado in dados:
            if len(dado) > 4 and dado[4] == codigo_formatado:
                dados_etiqueta = dado
                break
        
        if dados_etiqueta:
            descricao_ingles = dados_etiqueta[1]
            descricao_portugues = dados_etiqueta[2]
            material = dados_etiqueta[0]
        else:
            descricao_ingles = "N/A"
            descricao_portugues = "N/A"
            material = "N/A"
        
        qr_data = f"Material:{material};numeroserie({codigo_formatado})"
        qr = qrcode.QRCode(version=1, box_size=10, border=4)
        qr.add_data(qr_data)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white")
        
        img_buffer = io.BytesIO()
        img.save(img_buffer, format="PNG")
        img_base64 = base64.b64encode(img_buffer.getvalue()).decode()
        
        etiqueta = f"""
        ^XA
        ^CF0,30
        ^FO50,20^FD{descricao_ingles}^FS
        ^FO50,60^FD{descricao_portugues}^FS
        ^FO50,100^FDMaterial-no.: {material}^FS
        ^FO50,140^FDSerial-no.: {codigo_formatado}^FS
        ^FO50,180^FDPage 1^FS
        ^FO50,220^IMGd:{len(img_base64)},{img_base64}^FS
        ^XZ
        """
        
        if imprimir_zebra(printer_ip, printer_port, etiqueta):
            print(f"Imprimindo etiqueta para: {codigo_formatado}")
            time.sleep(1)

def imprimir_etiquetas():
    caminho_arquivo = "bd.xlsx"
    dados = ler_arquivo_excel(caminho_arquivo)
    
    materiais_descricoes = []
    for dado in dados:
        material_descricao = f"{dado[0]} - {dado[1]}"
        if material_descricao not in materiais_descricoes:
            materiais_descricoes.append(material_descricao)
    materiais_descricoes.sort()
    
    janela = tk.Tk()
    janela.title("Impressão de Etiquetas")
    
    frame_materiais = ttk.LabelFrame(janela, text="Selecionar Materiais e Descrições")
    frame_materiais.pack(padx=10, pady=10)
    
    canvas = tk.Canvas(frame_materiais)
    scrollbar = Scrollbar(frame_materiais, orient="vertical", command=canvas.yview)
    scrollable_frame = ttk.Frame(canvas)
    
    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")
        )
    )
    
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    
    vars_materiais = [tk.IntVar() for _ in materiais_descricoes]
    checkbuttons_materiais = []
    for i, material_descricao in enumerate(materiais_descricoes):
        checkbutton = ttk.Checkbutton(scrollable_frame, text=material_descricao, variable=vars_materiais[i])
        checkbutton.pack(anchor="w")
        checkbuttons_materiais.append(checkbutton)
    
    frame_codigos = ttk.LabelFrame(janela, text="Intervalo de Códigos")
    frame_codigos.pack(padx=10, pady=10)
    
    label_inicial = ttk.Label(frame_codigos, text="Código Inicial (apenas números):")
    label_inicial.grid(row=0, column=0)
    entry_inicial = ttk.Entry(frame_codigos)
    entry_inicial.grid(row=0, column=1)
    
    label_final = ttk.Label(frame_codigos, text="Código Final (apenas números):")
    label_final.grid(row=1, column=0)
    entry_final = ttk.Entry(frame_codigos)
    entry_final.grid(row=1, column=1)
    
    def imprimir():
        materiais_descricoes_selecionados = [materiais_descricoes[i] for i, var in enumerate(vars_materiais) if var.get() == 1]
        
        try:
            codigo_inicial = int(entry_inicial.get())
            codigo_final = int(entry_final.get())
        except ValueError:
            messagebox.showerror("Erro", "Por favor, insira apenas números nos campos de código inicial e final.")
            return
        
        printer_ip = "192.168.221.99"
        printer_port = 9100
        
        gerar_etiqueta(codigo_inicial, codigo_final, printer_ip, printer_port, dados)
        messagebox.showinfo("Impressão Concluída", "Impressão de etiquetas concluída.")
    
    botao_imprimir = ttk.Button(janela, text="Imprimir", command=imprimir)
    botao_imprimir.pack(pady=10)
    
    janela.mainloop()

if __name__ == "__main__":
    imprimir_etiquetas()