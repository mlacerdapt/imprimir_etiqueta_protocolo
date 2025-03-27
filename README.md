# imprimir_etiqueta_protocolo
 Sistema de gestão de impressão de protocolos.


Bibliotecas instaladas:

pip install openpyxl zebra qrcode pillow

IP Impressoras: 

PTVDCP0048 - 192.168.221.81
PTVDCP0050 - 192.168.221.99
PTVDCP0051 - 192.168.221.69
PTVDCP0052 - impressora desligada


1. Requisitos Principais

Ler Dados do Excel: Ler dados de um arquivo Excel onde a área de impressão já está definida.
Tamanho da Etiqueta: A etiqueta tem dimensões de 104mm x 40mm.
Preencher Campos: Preencher os campos da etiqueta com os dados fornecidos no Excel.
Gerar QR Code: Criar um QR code com o formato "Material;numeroserie(EVC0000)".
Incrementar Número de Série: Imprimir uma etiqueta para cada número de série no intervalo especificado.