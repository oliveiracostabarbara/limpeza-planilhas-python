import pandas as pd

a = pd.read_csv("dados/vendas_loja.csv", sep=";", dtype=str, encoding="utf-8-sig")
b = pd.read_excel("dados/pedidos_site.xlsx", dtype=str)
c = pd.read_excel("dados/pedidos_whatsapp.xlsx", dtype=str)

colunas = ["cliente", "telefone", "data", "produto", "valor"]
a.columns = colunas
b.columns = colunas
c.columns = colunas

a["fonte"] = "loja"
b["fonte"] = "site"
c["fonte"] = "whatsapp"

df = pd.concat([a, b, c], ignore_index=True)

# nomes: tira espaços sobrando e padroniza "Ana Alves"
df["cliente"] = df["cliente"].str.strip().str.replace(r"\s+", " ", regex=True).str.title()

# telefones: fica só com dígitos e tira o 55 do início
df["telefone"] = df["telefone"].str.replace(r"\D", "", regex=True)
df.loc[(df["telefone"].str.len() == 13) & df["telefone"].str.startswith("55"), "telefone"] = df["telefone"].str[2:]

# valores: tira "R$", ajusta vírgula/ponto e converte em número
df["valor"] = df["valor"].str.replace("R$", "", regex=False).str.strip()
tem_virgula = df["valor"].str.contains(",", na=False)
df.loc[tem_virgula, "valor"] = df["valor"].str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
df["valor"] = pd.to_numeric(df["valor"])

# datas: tenta cada formato e junta os resultados
formatos = ["%d/%m/%Y", "%Y-%m-%d", "%d/%m/%y", "%d-%m-%Y"]
datas = pd.Series(pd.NaT, index=df.index)
for f in formatos:
    datas = datas.fillna(pd.to_datetime(df["data"], format=f, errors="coerce"))
df["data"] = datas

# Passo 6: separar linhas inválidas
def motivo(linha):
    if pd.isna(linha["valor"]):
        return "sem valor"
    if pd.isna(linha["data"]):
        return "sem data"
    if pd.isna(linha["telefone"]):
        return "sem telefone"
    if len(linha["telefone"]) != 11:
        return "telefone invalido"
    return ""

df["motivo"] = df.apply(motivo, axis=1)
erros = df[df["motivo"] != ""].copy()
validos = df[df["motivo"] == ""].copy()

print(len(erros), len(validos))
print(erros["motivo"].value_counts())

# Passo 7: remover duplicados
antes = len(validos)
validos = validos.drop_duplicates(subset=["telefone", "data", "produto", "valor"], keep="first")
print("duplicados removidos:", antes - len(validos))
print("pedidos unicos:", len(validos))
print("receita:", round(validos["valor"].sum(), 2))

# Passo 8: gerar o Excel final
colunas_saida = ["cliente", "telefone", "data", "produto", "valor", "fonte"]
pedidos = validos[colunas_saida].sort_values("data").reset_index(drop=True)
erros = erros[colunas_saida + ["motivo"]].reset_index(drop=True)

with pd.ExcelWriter("saida/pedidos_consolidados.xlsx", engine="openpyxl") as arquivo:
    pedidos.to_excel(arquivo, sheet_name="Pedidos", index=False)
    erros.to_excel(arquivo, sheet_name="Erros", index=False)
    for aba in arquivo.sheets.values():
        for coluna in aba.columns:
            aba.column_dimensions[coluna[0].column_letter].width = 18
        for celula in aba["C"][1:]:
            celula.number_format = "DD/MM/YYYY"

print("Arquivo gerado:", len(pedidos), "pedidos e", len(erros), "erros")