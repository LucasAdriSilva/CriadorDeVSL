import os, base64

nomeArquivo = 'default.pptx'

with open(f'{nomeArquivo}', 'rb') as file:
    encodedString = base64.b64encode(file.read())

with open(f'{nomeArquivo.split(".")[0]}.bin', 'wb') as outfile:
    outfile.write(encodedString)

