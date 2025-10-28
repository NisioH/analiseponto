import pandas as pd
import io
from datetime import datetime

# --- CONFIGURAÇÃO DA DATA DE HOJE ---
# Usamos a data atual para filtro.
data_de_hoje = datetime.now().date()

# --- 1. CARREGAMENTO E PRÉ-PROCESSAMENTO DOS ARQUIVOS DE PONTO ---
arquivos = ['Ponto_Algodoeira.txt', 'Ponto_Escritorio.txt', 'Ponto_Sede.txt', 'Ponto_Secador.txt']
dfs_ponto = []
for arquivo in arquivos:
    df = pd.read_csv(arquivo, header=None, names=['linha_completa'], encoding='latin-1')
    df['origem'] = arquivo
    dfs_ponto.append(df)
df_ponto_raw = pd.concat(dfs_ponto, ignore_index=True)

# --- 2. CARREGAMENTO E PADRONIZAÇÃO DOS DADOS DE FUNCIONÁRIOS ---
df_funcionarios = pd.read_excel(r'Funcionarios.xlsx')
df_funcionarios['NIT_STR'] = df_funcionarios['NIT'].astype(str)
nit_to_nome = df_funcionarios.set_index('NIT_STR')['Nome'].to_dict()
nit_to_secao = df_funcionarios.set_index('NIT_STR')['Secao'].to_dict()

# --- 3. BUSCA REVERSA E EXTRAÇÃO (VALIDAÇÃO DE NIT E DATA) ---
registros_validos = []
for index, row in df_ponto_raw.iterrows():
    linha_completa = row['linha_completa']
    nit_encontrado = None
    for nit_valido in nit_to_nome.keys():
        if nit_valido in linha_completa:
            nit_encontrado = nit_valido
            break

    if nit_encontrado:
        try:
            data_str = linha_completa[10:18]
            data_do_ponto = datetime.strptime(data_str, '%d%m%Y').date()

            if data_do_ponto == data_de_hoje:
                hora_str = linha_completa[18:22]
                data_hora = datetime.strptime(data_str + hora_str, '%d%m%Y%H%M')

                registros_validos.append({
                    'NIT': nit_encontrado,
                    'Nome': nit_to_nome.get(nit_encontrado, 'N/A'),
                    'Secao': nit_to_secao.get(nit_encontrado, 'N/A'),
                    'data_hora': data_hora,
                    'origem': row['origem']
                })
        except ValueError:
            continue

df_ponto_validado = pd.DataFrame(registros_validos)

# --- 4. GERAÇÃO DO RELATÓRIO DE PRESENÇA COM HORÁRIO ---

# Calcula o primeiro ponto de cada funcionário (a menor hora)
df_primeiro_ponto = df_ponto_validado.loc[df_ponto_validado.groupby('NIT')['data_hora'].idxmin()]

df_batidos = df_primeiro_ponto[['Nome', 'Secao', 'data_hora']].copy()
df_batidos['Horario'] = df_batidos['data_hora'].dt.strftime('%H:%M')

# --- 5. SALVAMENTO NO ARQUIVO .TXT AGRUPADO COM FORMATAÇÃO MELHORADA ---

print(f"\n--- Geração do Relatório de Presença ({data_de_hoje.strftime('%d/%m/%Y')}) ---\n")

output_buffer = io.StringIO()
df_batidos = df_batidos.sort_values(by=['Secao', 'Nome'])

# Define o tamanho máximo que a coluna Nome deve ter para alinhamento
MAX_NOME_LEN = df_batidos['Nome'].str.len().max() if not df_batidos.empty else 40
# Garante que o preenchimento tenha um tamanho mínimo de 40 para não ficar muito apertado
NOME_PAD_LEN = max(MAX_NOME_LEN, 40) + 2

# Formatação do cabeçalho
output_buffer.write(f"RELATÓRIO DE PRESENÇA - DATA: {data_de_hoje.strftime('%d/%m/%Y')}\n")
output_buffer.write("-" * 80 + "\n")

# Agrupamento e escrita no arquivo
for secao, grupo in df_batidos.groupby('Secao'):
    output_buffer.write(f"\n#####################################################\n")
    output_buffer.write(f"SEÇÃO: {secao}\n")
    output_buffer.write(f"TOTAL DE FUNCIONÁRIOS NA SEÇÃO: {len(grupo)}\n")
    output_buffer.write(f"#####################################################\n")

    # Escreve o cabeçalho da tabela interna
    output_buffer.write(f"| {'NOME DO FUNCIONÁRIO':<{NOME_PAD_LEN}} | {'HORÁRIO (ENTRADA)':<18} |\n")
    output_buffer.write(f"|{'-' * (NOME_PAD_LEN + 2)}|{'-' * 20}|\n")

    # Escreve os dados com alinhamento
    for _, item in grupo.iterrows():
        # Usamos :<{NOME_PAD_LEN} para alinhar o nome à esquerda com padding fixo
        output_buffer.write(f"| {item['Nome']:<{NOME_PAD_LEN}} | {item['Horario']:<18} |\n")

    output_buffer.write("\n")

nome_arquivo_saida = f"Relatorio_Presenca_{data_de_hoje.strftime('%Y%m%d')}.txt"
with open(nome_arquivo_saida, 'w', encoding='utf-8') as f:
    f.write(output_buffer.getvalue())

print(f"Total de funcionários presentes únicos: {len(df_batidos)}")
print(f"Relatório de Presença com Horários formatado salvo em: {nome_arquivo_saida}")