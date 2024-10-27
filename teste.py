import pandas as pd  # Importa a biblioteca pandas

def ler_arquivo(nomearq):
    """Lê o arquivo e retorna suas linhas. Lança uma exceção se o arquivo não existir."""
    try:
        with open(nomearq, "r", encoding="latin-1") as arq:
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
    """Salva a lista de palavras em um arquivo Excel."""
    df = pd.DataFrame(palavras)
    df.to_excel("palavras_extraidas.xlsx", index=False, header=False)
    print("Dados exportados para 'palavras_extraidas.xlsx'")

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