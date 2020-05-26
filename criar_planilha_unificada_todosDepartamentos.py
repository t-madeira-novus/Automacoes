import pandas as pd
from tqdm import tqdm

df = pd.read_excel("P:\documentos\OneDrive - Novus Contabilidade\Doc Compartilhado\Controles\DIVISÃO DE EMPRESAS - 05.2020.xlsx")
df_dp = pd.read_excel("df_dp.xlsm")
df_final = pd.DataFrame(columns=["Códigos", 'Contábil', 'Fiscal', 'Pessoal'])
codigos = []
j = 0

for i in tqdm(df.index):
    try:
        codigo = int(df.at[i, "Codigo"])
        codigos.append(codigo)
    except:
        codigo = -1
    responsavel_contabil = str(df.at[i, "DEPARTAMENTO CONTABIL"])
    responsavel_fiscal = str(df.at[i, "DEPARTAMENTO FISCAL"])
    df_final.at[j, "Códigos"] = codigo
    df_final.at[j, "Contábil"] = responsavel_contabil
    df_final.at[j, "Fiscal"] = responsavel_fiscal
    j += 1


for codigo in tqdm(codigos):
    for i in df_dp.index:
        try:
            aux = int (df_dp.at[i, "Codigos Empresas"])
        except:
            aux = -1
        if codigo == aux :

            j = 0
            for j in df.index:
                if codigo == df.at[j, "Codigo"]:
                    df_final.at[j, 'Pessoal'] = str(df_dp.at[i, "Responsável a partir de \n02/2017"])


df_final.to_csv("responsaveis_empresas.csv", encoding='latin-1', sep=';', index=False)