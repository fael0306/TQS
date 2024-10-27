import pandas as pd
import glob
import os

def ler_arquivo(nomearq):
    try:
        with open(nomearq, "r", encoding="latin-1") as arq:
            return arq.readlines()
    except FileNotFoundError:
        print(f"Erro: O arquivo '{nomearq}' não foi encontrado.")
        return None

def encontrar_linhas_quantitativos(linhas):
    linhas_encontradas = []
    quantitativos_encontrados = False

    for i, linha in enumerate(linhas):
        if "Quantitativos" in linha:
            quantitativos_encontrados = True
            for j in range(i + 1, len(linhas)):
                if "Legenda" in linhas[j]:
                    break  # Para se encontrar "Legenda"
                linhas_encontradas.append(linhas[j].strip())
            break  # Sai do loop ao encontrar "Quantitativos"

    return quantitativos_encontrados, linhas_encontradas

def salvar_em_excel(palavras):
    df = pd.DataFrame(palavras)
    df.to_excel("palavras_extraidas.xlsx", index=False, header=False)
    print("Dados exportados para 'palavras_extraidas.xlsx'")

def main():
    arquivos_lst = glob.glob("*.LST")
    
    if not arquivos_lst:
        print("Nenhum arquivo com extensão .LST foi encontrado na pasta.")
        return

    nomearq = arquivos_lst[0]  # Pega o primeiro arquivo encontrado com extensão .LST
    print(f"Lendo o arquivo '{nomearq}'")

    linhas = ler_arquivo(nomearq)

    if linhas is not None:
        quantitativos_encontrados, linhas_encontradas = encontrar_linhas_quantitativos(linhas)

        if quantitativos_encontrados:
            palavras = [linha.split() for linha in linhas_encontradas]
            salvar_em_excel(palavras)
        else:
            print("A palavra 'Quantitativos' não foi encontrada.")

if __name__ == "__main__":
    main()
