import pandas as pd
import io
from datetime import datetime
from datetime import date


data_de_hoje = datetime.now().date()
#date(2025, 11, 11)

arquivos = ['Ponto_Algodoeira.txt', 'Ponto_Escritorio.txt', 'Ponto_Sede.txt', 'Ponto_Secador.txt']
dfs_ponto = []
for arquivo in arquivos:
    df = pd.read_csv(arquivo, header=None, names=['linha_completa'], encoding='latin-1')
    df['origem'] = arquivo
    dfs_ponto.append(df)
df_ponto_raw = pd.concat(dfs_ponto, ignore_index=True)

df_funcionarios = pd.read_excel(r'C:\Users\fazin\OneDrive\Documents\Nisio\Analise_Ponto\Funcionarios.xlsx')
df_funcionarios['NIT_STR'] = df_funcionarios['NIT'].astype(str)
nit_to_nome = df_funcionarios.set_index('NIT_STR')['Nome'].to_dict()
nit_to_secao = df_funcionarios.set_index('NIT_STR')['Secao'].to_dict()

# ---VALIDAÇÃO DE NIT E DATA
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

# --- FILTRAR ÚLTIMOS 3 MESES ---
# Data limite: hoje menos 3 meses
data_limite = data_de_hoje.replace(day=1)  # início do mês atual
# Para subtrair 3 meses corretamente, usamos relativedelta
from dateutil.relativedelta import relativedelta
data_limite = data_de_hoje - relativedelta(months=3)

# Filtra registros válidos dos últimos 3 meses
df_ultimos_3_meses = df_ponto_validado[df_ponto_validado['data_hora'].dt.date >= data_limite]

# Ordena por funcionário e data/hora
df_ultimos_3_meses = df_ultimos_3_meses.sort_values(by=['Nome', 'data_hora'])

# Apenas para conferência
print(f"\n--- Registros dos últimos 3 meses ---\n")
print(df_ultimos_3_meses[['Nome', 'Secao', 'data_hora', 'origem']].head(20))

# Salvar em TXT
nome_arquivo_3m_txt = f"Registros_Ultimos3Meses_{data_de_hoje.strftime('%Y%m%d')}.txt"
df_ultimos_3_meses.to_csv(nome_arquivo_3m_txt, index=False, sep=";", encoding="utf-8")

# Salvar em Excel
nome_arquivo_3m_xlsx = f"Registros_Ultimos3Meses_{data_de_hoje.strftime('%Y%m%d')}.xlsx"
df_ultimos_3_meses.to_excel(nome_arquivo_3m_xlsx, index=False)


# Filtra apenas funcionários com um único registro de ponto
df_unico_ponto = df_ponto_validado[df_ponto_validado['NIT'].map(df_ponto_validado['NIT'].value_counts()) == 1]

df_primeiro_ponto = df_unico_ponto.loc[df_unico_ponto.groupby('NIT')['data_hora'].idxmin()]

df_batidos = df_primeiro_ponto[['Nome', 'Secao', 'data_hora']].copy()
df_batidos['Horario'] = df_batidos['data_hora'].dt.strftime('%H:%M')


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


# Salvando em .txt
for secao, grupo in df_batidos.groupby('Secao'):
    output_buffer.write(f"\n#####################################################\n")
    output_buffer.write(f"SEÇÃO: {secao}\n")
    output_buffer.write(f"TOTAL DE FUNCIONÁRIOS NA SEÇÃO: {len(grupo)}\n")
    output_buffer.write(f"#####################################################\n")

    
    output_buffer.write(f"| {'NOME DO FUNCIONÁRIO':<{NOME_PAD_LEN}} | {'HORÁRIO (ENTRADA)':<18} |\n")
    output_buffer.write(f"|{'-' * (NOME_PAD_LEN + 2)}|{'-' * 20}|\n")

    # Escreve os dados com alinhamento
    for _, item in grupo.iterrows():
        
        output_buffer.write(f"| {item['Nome']:<{NOME_PAD_LEN}} | {item['Horario']:<18} |\n")

    output_buffer.write("\n")

nome_arquivo_saida = f"Relatorio_Presenca_{data_de_hoje.strftime('%Y%m%d')}.txt"
with open(nome_arquivo_saida, 'w', encoding='utf-8') as f:
    f.write(output_buffer.getvalue())


# Salvando em Excel
dados_por_secao = {}

# Agrupa os dados por seção
for secao, grupo in df_batidos.groupby('Secao'):
    
    df_secao = grupo[['Nome', 'Horario']].copy()
    df_secao.columns = ['NOME DO FUNCIONÁRIO', 'HORÁRIO (ENTRADA)']
    
   
    dados_por_secao[secao] = df_secao


nome_arquivo_saida = f"Relatorio_Presenca_{data_de_hoje.strftime('%Y%m%d')}.xlsx"

# Grava cada seção em uma aba do Excel
with pd.ExcelWriter(nome_arquivo_saida, engine='xlsxwriter') as writer:
    for secao, df_secao in dados_por_secao.items():
        
        aba = str(secao)[:31]
        df_secao.to_excel(writer, sheet_name=aba, index=False)

print(f"Total de funcionários presentes únicos: {len(df_batidos)}")
print(f"Relatório de Presença com Horários formatado salvo em: {nome_arquivo_saida}")
