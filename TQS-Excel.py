import os

def ler_arquivo(nomearq):
    """Lê o arquivo e retorna suas linhas. Lança uma exceção se o arquivo não existir."""
    # Obtém o diretório do script atual
    diretorio_atual = os.path.dirname(os.path.abspath(__file__))
    # Cria o caminho completo para o arquivo
    caminho_arquivo = os.path.join(diretorio_atual, nomearq)

    try:
        with open(caminho_arquivo, "r", encoding="latin-1") as arq:
            return arq.readlines()
    except FileNotFoundError:
        print(f"Erro: O arquivo '{nomearq}' não foi encontrado.")
        return None  # Retorna None se o arquivo não for encontrado

def encontrar_linhas_quantitativos(linhas):
    """Encontra as linhas entre 'Quantitativos' e 'Legenda'."""
    linhas_encontradas = []
    quantitativos_encontrados = False

    for i, linha in enumerate(linhas):
        if "Quantitativos" in linha:
            quantitativos_encontrados = True
            # Lê as próximas linhas até encontrar "Legenda"
            for j in range(i + 1, len(linhas)):
                if "Legenda" in linhas[j]:  # Verifica se a linha contém "Legenda"
                    break  # Para se encontrar "Legenda"
                linhas_encontradas.append(linhas[j].strip())  # Adiciona a linha encontrada
            break  # Sai do loop ao encontrar "Quantitativos"

    return quantitativos_encontrados, linhas_encontradas

def salvar_em_excel(palavras):
    """Salva a lista de palavras em um arquivo Excel (XML) no diretório atual."""
    # Obtém o diretório do script atual
    diretorio_atual = os.path.dirname(os.path.abspath(__file__))
    # Define o caminho completo para o arquivo
    caminho_arquivo = os.path.join(diretorio_atual, "palavras_extraidas.xml")

    with open(caminho_arquivo, "w", encoding="utf-8") as f:
        # Escreve o cabeçalho do XML para Excel
        f.write('<?xml version="1.0"?>\n')
        f.write('<?mso-application progid="Excel.Sheet"?>\n')
        f.write('<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"\n')
        f.write(' xmlns:o="urn:schemas-microsoft-com:office:office"\n')
        f.write(' xmlns:x="urn:schemas-microsoft-com:office:excel"\n')
        f.write(' xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"\n')
        f.write(' xmlns:html="http://www.w3.org/TR/REC-html40">\n')
        f.write('<Worksheet ss:Name="Sheet1">\n')
        f.write('<Table>\n')

        # Adiciona cada linha de palavras ao XML
        for linha in palavras:
            f.write('  <Row>\n')
            for palavra in linha:
                f.write(f'    <Cell><Data ss:Type="String">{palavra}</Data></Cell>\n')
            f.write('  </Row>\n')

        # Fecha as tags do XML
        f.write('</Table>\n')
        f.write('</Worksheet>\n')
        f.write('</Workbook>\n')

    print(f"Dados exportados para '{caminho_arquivo}'")

def main():
    while True:
        nomearq = input("Digite o nome do arquivo (com extensão): ")
        linhas = ler_arquivo(nomearq)

        if linhas is not None:  # Verifica se o arquivo foi lido com sucesso
            quantitativos_encontrados, linhas_encontradas = encontrar_linhas_quantitativos(linhas)

            if quantitativos_encontrados:
                # Divide cada linha encontrada em palavras
                palavras = [linha.split() for linha in linhas_encontradas]
                salvar_em_excel(palavras)
                break  # Sai do loop após salvar os dados
            else:
                print("A palavra 'Quantitativos' não foi encontrada.")
                break  # Sai do loop se 'Quantitativos' não for encontrado

if __name__ == "__main__":
    main()

