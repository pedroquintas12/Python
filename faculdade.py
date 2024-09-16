import pandas as pd

df_original = pd.read_csv("C:\\Users\\pedro\\OneDrive\\Documentos\\GitHub\\Python\\dados.csv", sep=';',  encoding='ISO-8859-1')


print(df_original.describe())