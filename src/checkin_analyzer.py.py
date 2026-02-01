# ===== 1) IMPORTS (sempre no topo) =====
import os
from glob import glob
import pandas as pd


# ===== 2) FUNÇÕES (definem comportamentos, não executam) =====

def carregar_planilha(caminho_arquivo):
    """
    Detecta modelo de planilha e retorna dataframe + nomes das colunas corretas
    """

    # Modelo NOVO
    try:
        df = pd.read_excel(caminho_arquivo, header=12)
        if "Visitante" in df.columns and "Pagamento" in df.columns:
            print("📄 Modelo NOVO detectado →", os.path.basename(caminho_arquivo))
            return df, "Visitante", "Pagamento"
    except:
        pass

    # Modelo ANTIGO
    try:
        df = pd.read_excel(caminho_arquivo, header=2)
        if "Unnamed: 4" in df.columns and "Unnamed: 5" in df.columns:
            print("📄 Modelo ANTIGO detectado →", os.path.basename(caminho_arquivo))
            return df, "Unnamed: 4", "Unnamed: 5"
    except:
        pass

    print("Modelo não reconhecido:", caminho_arquivo)
    return None, None, None


# ===== 3) MAIN (onde o programa EXECUTA de verdade) =====

def main():

    # CONFIGURAÇÃO
    pasta = r"C:\Users\marci\Desktop\Contagem de Check-ins\Contador-Checkins\relatorios"

    # BUSCAR ARQUIVOS
    arquivos = glob(os.path.join(pasta, "*.xls*"))
    arquivos = [a for a in arquivos if "ranking_" not in os.path.basename(a)]

    if not arquivos:
        print("Nenhum arquivo Excel encontrado na pasta")
        return  # <- melhor que exit()

    rankings_por_arquivo = {}

    # PROCESSAR CADA PLANILHA
    for arquivo in arquivos:

        df, coluna_aluno, coluna_valor = carregar_planilha(arquivo)

        if df is None:
            continue

        # Padronizar colunas
        df = df.rename(columns={
            coluna_aluno: "Aluno",
            coluna_valor: "Valor"
        })

        df = df[["Aluno", "Valor"]]
        df = df.dropna(subset=["Aluno", "Valor"])

        # Limpar valores monetários
        df["Valor"] = (
            df["Valor"]
            .astype(str)
            .str.replace("R$", "", regex=False)
            .str.replace(" ", "", regex=False)
            .str.replace(".", "", regex=False)
            .str.replace(",", ".", regex=False)
        )

        df["Valor"] = pd.to_numeric(df["Valor"], errors="coerce")
        df = df.dropna(subset=["Valor"])

        # Agrupar dados
        ranking = df.groupby("Aluno").agg(
            Numero_Acessos=("Aluno", "count"),
            Valor_Total_R=("Valor", "sum")
        ).reset_index()

        ranking = ranking.sort_values(by="Numero_Acessos", ascending=False)

        nome_aba = os.path.splitext(os.path.basename(arquivo))[0]
        rankings_por_arquivo[nome_aba] = ranking

    # SALVAR RESULTADO FINAL
    saida = os.path.join(pasta, "ranking_consolidado.xlsx")

    with pd.ExcelWriter(saida, engine="openpyxl") as writer:
        for nome_aba, ranking_df in rankings_por_arquivo.items():
            ranking_df.to_excel(writer, sheet_name=nome_aba[:31], index=False)

    print(f"\n✅ Arquivo final gerado com sucesso: {saida}")


# ===== 4) PONTO DE ENTRADA =====
# Só roda main() se o arquivo for executado diretamente

if __name__ == "__main__":
    main()
