import csv
import pandas as pd
import time
import sys
import os

def caminho_absoluto(relativo):
        """
        Retorna o caminho correto tanto no Python normal
        quanto no executável do PyInstaller
        """
        if hasattr(sys, '_MEIPASS'):
            return os.path.join(sys._MEIPASS, relativo)
        return os.path.join(os.path.abspath("."), relativo)  

def separar_nomes():
    postos = [
        "Sd","EV", "Cb", "3º", "Sgt", "ST", "Asp", "Al", "Al/CPOR", "NPOR",
        "Ten", "2º", "1º", "Cap", "Maj", "Ten Cel", "Cel"
    ]

    a = "pronta"

    #PEGA O ARQUIVO DE TEXTO COM TODOS OS NOMES E TRANFORMA EM UMA LISTA
    with open(file=caminho_absoluto("pessoas.txt"), mode="r", encoding="utf-8") as arquivo:
        militares = [pessoa.replace("\n", "") for pessoa in arquivo.readlines()]


    with open(f"arquivo_csv_graduacao_nomes.csv", mode="w", newline="", encoding="utf-8") as arquivo:
        writer = csv.writer(arquivo)
        writer.writerow(["graduacao", "nome_completo"])


        for linha in militares:
            partes = linha.split()

            # Separa as palavras da graduação (enquanto elas estiverem na lista de postos/complementos)
            graduacao = []
            while partes and (partes[0] in postos):
                #O Pop ele retira da lista e RETORNA o valor retirado
                graduacao.append(partes.pop(0))
            if "EV" not in graduacao:
                graduacao = [posto.replace("Sd", "Sd EP") for posto in graduacao]
            graduacao = [posto.replace("ST", "SubTen") for posto in graduacao]
            partes = [parte_nome.replace("INT", "") for parte_nome in partes]
            partes = [parte_nome.replace("INF", "") for parte_nome in partes]



            nome = " ".join(partes)
            graduacao = ' '.join(graduacao)
            if graduacao != "":
                writer.writerow([graduacao, nome])

    time.sleep(2)

    return caminho_absoluto("arquivo_csv_graduacao_nomes.csv")
