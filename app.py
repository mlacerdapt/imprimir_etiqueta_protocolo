import openpyxl
import zebra
import time
import qrcode
import io
import base64

# Função para ler o arquivo Excel e extrair dados necessários
def ler_arquivo_excel(caminho_arquivo):
    wb = openpyxl.load_workbook(caminho_arquivo)
    planilha = wb.active  # A primeira planilha do arquivo

    dados = []
    for linha in planilha.iter_rows(min_row=2, values_only=True):  # Ignora cabeçalho e obtém valores
        dados.append(linha)

    return dados

# Função para gerar a etiqueta para a impressora Zebra
def gerar_etiqueta(dados, codigo_inicial, codigo_final, impressora):
    for codigo in range(codigo_inicial, codigo_final + 1):
        codigo_formatado = f"EVC{codigo:04d}"
        
        # Encontra os dados correspondentes ao código no Excel
        dados_etiqueta = None
        for dado in dados:
            if dado[4] == codigo_formatado:  # Assumindo que o código está na coluna E
                dados_etiqueta = dado
                break
        
        if dados_etiqueta:
            material = dados_etiqueta[0]  # Assumindo que o material está na coluna A
            descricao_ingles = dados_etiqueta[1]  # Assumindo que a descrição em inglês está na coluna B
            descricao_portugues = dados_etiqueta[2]  # Assumindo que a descrição em português está na coluna C
            
            # Gerar QR code
            qr_data = f"Material:{material};numeroserie({codigo_formatado})"
            qr = qrcode.QRCode(version=1, box_size=10, border=4)
            qr.add_data(qr_data)
            qr.make(fit=True)
            img = qr.make_image(fill_color="black", back_color="white")
            
            # Converter imagem QR code para base64
            img_buffer = io.BytesIO()
            img.save(img_buffer, format="PNG")
            img_base64 = base64.b64encode(img_buffer.getvalue()).decode()
            
            # Gerar comando ZPL para a etiqueta
            etiqueta = f"""
            ^XA
            ^CF0,30
            ^FO50,50^FD{material}^FS
            ^FO50,100^FD{descricao_ingles}^FS
            ^FO50,150^FD{descricao_portugues}^FS
            ^FO50,200^FD{codigo_formatado}^FS
            ^FO50,250^IMGd:{len(img_base64)},{img_base64}^FS
            ^XZ
            """
            
            # Enviar comando de impressão
            impressora.output(etiqueta)
            print(f"Imprimindo etiqueta para: {codigo_formatado}")
            time.sleep(1)  # Pausa de 1 segundo entre as impressões
        else:
            print(f"Código {codigo_formatado} não encontrado no Excel.")

# Função principal
def imprimir_etiquetas():
    caminho_arquivo = "seu_arquivo.xlsx"  # Substitua pelo caminho do seu arquivo Excel
    dados = ler_arquivo_excel(caminho_arquivo)
    
    codigo_inicial = int(input("Digite o código inicial para impressão: "))
    codigo_final = int(input("Digite o código final para impressão: "))
    
    impressora = zebra.Zebra()
    impressora.setqueue("192.168.221.99")  # Substitua pelo IP da sua impressora
    
    gerar_etiqueta(dados, codigo_inicial, codigo_final, impressora)

if __name__ == "__main__":
    imprimir_etiquetas()