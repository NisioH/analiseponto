import pandas as pd
import io
from datetime import datetime, timedelta

# 1. Define a data alvo como ONTEM
data_alvo = datetime.now().date() - timedelta(days=1)

arquivos = ['Ponto_Algodoeira.txt', 'Ponto_Escritorio.txt', 'Ponto_Sede.txt', 'Ponto_Secador.txt']
dfs_ponto = []
for arquivo in arquivos:
    df = pd.read_csv(arquivo, header=None, names=['linha_completa'], encoding='latin-1')
    df['origem'] = arquivo
    dfs_ponto.append(df)
df_ponto_raw = pd.concat(dfs_ponto, ignore_index=True)

df_funcionarios = pd.read_excel('Funcionarios.xlsx')
df_funcionarios['NIT_STR'] = df_funcionarios['NIT'].astype(str)

# Padroniza a nomenclatura da seção
df_funcionarios['Secao'] = df_funcionarios['Secao'].replace('Colaboradores sede', 'Colaboradores Sede')

nit_to_nome = df_funcionarios.set_index('NIT_STR')['Nome'].to_dict()
nit_to_secao = df_funcionarios.set_index('NIT_STR')['Secao'].to_dict()

# --- VALIDAÇÃO DE NIT E DATA ---
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

            # Filtrar os registros de ontem
            if data_do_ponto == data_alvo:
                hora_str = linha_completa[18:22]
                data_hora = datetime.strptime(data_str + hora_str, '%d%m%Y%H%M')

                # pro nome ficar limpo ('Ponto_Sede.txt' vira 'Sede')
                local_limpo = row['origem'].replace('Ponto_', '').replace('.txt', '')

                registros_validos.append({
                    'NIT': nit_encontrado,
                    'Nome': nit_to_nome.get(nit_encontrado, 'N/A'),
                    'Secao': nit_to_secao.get(nit_encontrado, 'N/A'),
                    'data_hora': data_hora,
                    'Local': local_limpo
                })
        except ValueError:
            continue

df_ponto_validado = pd.DataFrame(registros_validos)

if not df_ponto_validado.empty:
    df_ponto_validado['Horario'] = df_ponto_validado['data_hora'].dt.strftime('%H:%M')

    df_ponto_validado = df_ponto_validado.sort_values(by=['Nome', 'data_hora'])

    df_ponto_validado['Batida_Num'] = df_ponto_validado.groupby('NIT').cumcount() + 1

    # Cria uma string combinada: "08:00 (Sede)"
    df_ponto_validado['Registro_TXT'] = df_ponto_validado['Horario'] + " (" + df_ponto_validado['Local'] + ")"

    # Gira a tabela (pivot) para colocar as batidas em colunas
    df_txt = df_ponto_validado.pivot(index=['NIT', 'Nome', 'Secao'], columns='Batida_Num',
                                     values='Registro_TXT').reset_index()

    # Gira a tabela separando os valores de Horário e Local
    df_excel = df_ponto_validado.pivot(index=['NIT', 'Nome', 'Secao'], columns='Batida_Num',
                                       values=['Horario', 'Local'])

    # Renomeia as colunas do Excel para ficarem planas (Ex: 'Horario 1', 'Local 1')
    df_excel.columns = [f"{col[0]} {col[1]}" for col in df_excel.columns]
    df_excel = df_excel.reset_index()

else:
    df_txt = pd.DataFrame(columns=['NIT', 'Nome', 'Secao'])
    df_excel = pd.DataFrame(columns=['NIT', 'Nome', 'Secao'])

print(f"\n--- Geração do Relatório de Presença ({data_alvo.strftime('%d/%m/%Y')}) ---\n")

output_buffer = io.StringIO()

# --- SALVANDO O TXT ---
if not df_txt.empty:
    df_txt = df_txt.sort_values(by=['Secao', 'Nome'])

    MAX_NOME_LEN = df_txt['Nome'].str.len().max()
    NOME_PAD_LEN = max(MAX_NOME_LEN, 40) + 2

    output_buffer.write(f"RELATÓRIO DE PRESENÇA E LOCAIS - DATA: {data_alvo.strftime('%d/%m/%Y')}\n")
    output_buffer.write("-" * 120 + "\n")

    for secao, grupo in df_txt.groupby('Secao'):
        output_buffer.write(f"\n#####################################################\n")
        output_buffer.write(f"SEÇÃO: {secao}\n")
        output_buffer.write(f"TOTAL DE FUNCIONÁRIOS NA SEÇÃO: {len(grupo)}\n")
        output_buffer.write(f"#####################################################\n\n")

        # Pega a quantidade de batidas dinamicamente
        colunas_batidas = [c for c in grupo.columns if isinstance(c, int)]

        cabecalho = f"| {'NOME DO FUNCIONÁRIO':<{NOME_PAD_LEN}} |"
        linha_sep = f"|{'-' * (NOME_PAD_LEN + 2)}|"

        for col in colunas_batidas:
            cabecalho += f" {'BATIDA ' + str(col):<20} |"
            linha_sep += f"{'-' * 22}|"

        output_buffer.write(cabecalho + "\n")
        output_buffer.write(linha_sep + "\n")

        for _, item in grupo.iterrows():
            linha_func = f"| {item['Nome']:<{NOME_PAD_LEN}} |"
            for col in colunas_batidas:
                # Se não bateu o ponto nessa sequência, coloca "---"
                registro = str(item[col]) if pd.notna(item[col]) else "---"
                linha_func += f" {registro:<20} |"
            output_buffer.write(linha_func + "\n")

        output_buffer.write("\n")
else:
    output_buffer.write("Nenhum registro encontrado para a data especificada.\n")

nome_arquivo_saida_txt = f"Relatorio_Presenca_{data_alvo.strftime('%Y%m%d')}.txt"
with open(nome_arquivo_saida_txt, 'w', encoding='utf-8') as f:
    f.write(output_buffer.getvalue())

# --- SALVANDO O EXCEL ---
nome_arquivo_saida_xlsx = f"Relatorio_Presenca_{data_alvo.strftime('%Y%m%d')}.xlsx"

with pd.ExcelWriter(nome_arquivo_saida_xlsx, engine='xlsxwriter') as writer:
    if not df_excel.empty:
        df_excel = df_excel.sort_values(by=['Secao', 'Nome'])
        for secao, grupo in df_excel.groupby('Secao'):
            aba = str(secao)[:31]
            # Removemos NIT e Secao para limpar a aba
            df_secao = grupo.drop(columns=['NIT', 'Secao'])
            df_secao.rename(columns={'Nome': 'NOME DO FUNCIONÁRIO'}, inplace=True)
            df_secao.to_excel(writer, sheet_name=aba, index=False)
    else:
        pd.DataFrame(columns=['Aviso']).to_excel(writer, sheet_name='Sem Dados', index=False)

total_funcionarios = len(df_txt) if not df_txt.empty else 0
print(f"Total de funcionários processados: {total_funcionarios}")
print(f"Relatórios gerados com sucesso:\n- {nome_arquivo_saida_txt}\n- {nome_arquivo_saida_xlsx}")