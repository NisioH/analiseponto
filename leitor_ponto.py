import pandas as pd
from datetime import date, timedelta, datetime

# Etapa 1: Data de ontem
ontem = date.today() - timedelta(days=1)

# Etapa 2: Lê os arquivos de ponto
arquivos = ['PontoColaboradores.txt', 'Ponto_Escritorio.txt', 'Ponto_Sede.txt', 'Ponto_Secador.txt']
dfs_ponto = []

for arquivo in arquivos:
    df = pd.read_csv(arquivo, header=None, names=['registro'], encoding='latin-1')
    df['NIT'] = df['registro'].str[-11:]
    df['Data'] = pd.to_datetime(df['registro'].str[8:16], format='%d%m%Y', errors='coerce')
    df['Hora'] = df['registro'].str[16:20].str.replace(r'(\d{2})(\d{2})', r'\1:\2', regex=True)
    dfs_ponto.append(df)

df_ponto = pd.concat(dfs_ponto, ignore_index=True)

# Etapa 3: Lê os funcionários
funcionarios = pd.read_excel('Funcionarios.xlsx', usecols=['Nome', 'Secao', 'NIT'])
funcionarios['NIT'] = funcionarios['NIT'].astype(str).str.zfill(11)

# Etapa 4: Junta ponto + funcionários
df_completo = df_ponto.merge(funcionarios, on='NIT', how='left')

# Etapa 5: Filtra registros de ontem
df_ontem = df_completo[df_completo['Data'].dt.date == ontem]

# Etapa 6: Conta registros por NIT
contagem = df_ontem['NIT'].value_counts()
nits_unicos = contagem[contagem == 1].index
df_unico = df_ontem[df_ontem['NIT'].isin(nits_unicos)]

# Etapa 7: Agrupa por seção
nome_arquivo_saida = 'registros_unicos_por_secao.txt'

try:
    with open(nome_arquivo_saida, 'w', encoding='utf-8') as f:
        f.write("=" * 80 + "\n")
        f.write(f"REGISTROS ÚNICOS POR SEÇÃO - {ontem.strftime('%d/%m/%Y')}\n")
        f.write("=" * 80 + "\n\n")
        f.write(f"Total de registros únicos: {len(df_unico)}\n")
        f.write(f"Data de geração: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
        f.write("=" * 80 + "\n\n")

        if len(df_unico) > 0:
            secoes = df_unico['Secao'].dropna().unique()
            for secao in sorted(secoes):
                f.write(f"SEÇÃO: {secao}\n")
                f.write("-" * 80 + "\n")
                df_secao = df_unico[df_unico['Secao'] == secao]
                for idx, row in df_secao.iterrows():
                    f.write(f"Nome: {row['Nome']}\n")
                    f.write(f"NIT:  {row['NIT']}\n")
                    f.write(f"Hora: {row['Hora']}\n")
                    f.write(f"Total de registros: 1\n")
                    f.write("-" * 80 + "\n")
                f.write("\n")
        else:
            f.write("Nenhum registro único encontrado para ontem.\n")

    print(f"\n✓ Arquivo gerado com sucesso: {nome_arquivo_saida}")
except Exception as e:
    print(f"\n✗ Erro ao salvar arquivo: {e}")
