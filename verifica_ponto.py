import pandas as pd
import io
from datetime import datetime, timedelta, date

escolha_dia = input('Qual data pesquisar?\n\n Para ontem digite: 1\n\n Para uma data anterior ao dia de ontem digite: 2\n\n ')

datas_alvo = []
hoje = datetime.now().date()

if escolha_dia == '1':
    ontem = hoje - timedelta(days=1)
    
    if hoje.weekday() == 0: 
        print('\n[DETECTADO SEGUNDA-FEIRA] Buscando registros acumulados de Sábado e Domingo...')
        sabado = hoje - timedelta(days=2)
        domingo = hoje - timedelta(days=1)
        datas_alvo = [sabado, domingo]
    else:
        datas_alvo = [ontem]
    
    print('\nBuscando...\n')

else:
    data_input = input('\nDigite a data no formato DD/MM/AAAA: ')
    try:
        data_digitada = datetime.strptime(data_input, '%d/%m/%Y').date()
        datas_alvo = [data_digitada]
        print('\nBuscando...\n')
    except ValueError:
        print("Data inválida. Usando a data de ontem por padrão.")
        datas_alvo = [hoje - timedelta(days=1)]

if len(datas_alvo) > 1:
    string_data_arquivo = f"{datas_alvo[0].strftime('%Y%m%d')}_a_{datas_alvo[-1].strftime('%Y%m%d')}"
    titulo_relatorio = f"PERÍODO: {datas_alvo[0].strftime('%d/%m/%Y')} a {datas_alvo[-1].strftime('%d/%m/%Y')}"
else:
    string_data_arquivo = datas_alvo[0].strftime('%Y%m%d')
    titulo_relatorio = f"DATA: {datas_alvo[0].strftime('%d/%m/%Y')}"


arquivos = ['Ponto_Algodoeira.txt', 'Ponto_Escritorio.txt', 'Ponto_Sede.txt', 'Ponto_Secador.txt']
dfs_ponto = []
for arquivo in arquivos:
    try:
        df = pd.read_csv(arquivo, header=None, names=['linha_completa'], encoding='latin-1')
        df['origem'] = arquivo
        dfs_ponto.append(df)
    except FileNotFoundError:
        print(f"Aviso: O arquivo arquivo '{arquivo}' não foi encontrado e será ignorado.")

if not dfs_ponto:
    print("Nenhum arquivo de ponto de origem encontrado. Encerrando execução.")
    exit()

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

            # Filtrar se a data do ponto está na nossa lista de datas alvo (Sábado, Domingo, etc.)
            if data_do_ponto in datas_alvo:
                hora_str = linha_completa[18:22]
                data_hora = datetime.strptime(data_str + hora_str, '%d%m%Y%H%M')

                # Pro nome ficar limpo ('Ponto_Sede.txt' vira 'Sede')
                local_limpo = row['origem'].replace('Ponto_', '').replace('.txt', '')

                registros_validos.append({
                    'NIT': nit_encontrado,
                    'Nome': nit_to_nome.get(nit_encontrado, 'N/A'),
                    'Secao': nit_to_secao.get(nit_encontrado, 'N/A'),
                    'data_do_ponto': data_do_ponto, # Mantemos o controle do dia específico da batida
                    'data_hora': data_hora,
                    'Local': local_limpo
                })
        except ValueError:
            continue

df_ponto_validado = pd.DataFrame(registros_validos)

if not df_ponto_validado.empty:
   
    dias_semana_pt = {0: 'Seg', 1: 'Ter', 2: 'Qua', 3: 'Qui', 4: 'Sex', 5: 'Sáb', 6: 'Dom'}
    df_ponto_validado['Dia_Semana'] = df_ponto_validado['data_hora'].dt.weekday.map(dias_semana_pt)
    df_ponto_validado['Horario'] = df_ponto_validado['data_hora'].dt.strftime('%H:%M')

    df_ponto_validado = df_ponto_validado.sort_values(by=['Nome', 'data_hora'])

    # O contador de batidas agora reinicia por Funcionário E por Dia (essencial para o fim de semana não misturar tudo)
    df_ponto_validado['Batida_Num'] = df_ponto_validado.groupby(['NIT', 'data_do_ponto']).cumcount() + 1

    # Formato do texto que vai para a coluna
    df_ponto_validado['Registro_TXT'] = df_ponto_validado['Dia_Semana'] + " " + df_ponto_validado['Horario'] + " (" + df_ponto_validado['Local'] + ")"

    # Gira a tabela (pivot). Incluímos 'data_do_ponto' para separar as linhas do mesmo funcionário se ele trabalhou nos dois dias
    df_txt = df_ponto_validado.pivot(index=['NIT', 'Nome', 'Secao', 'data_do_ponto'], columns='Batida_Num',
                                     values='Registro_TXT').reset_index()
else:
    df_txt = pd.DataFrame(columns=['NIT', 'Nome', 'Secao', 'data_do_ponto'])

print(f"\n--- Geração do Relatório de Presença ({titulo_relatorio}) ---\n")

output_buffer = io.StringIO()

# --- SALVANDO O TXT ---
if not df_txt.empty:
    # Ordena por Seção, Nome e Data do Ponto
    df_txt = df_txt.sort_values(by=['Secao', 'Nome', 'data_do_ponto'])

    MAX_NOME_LEN = df_txt['Nome'].str.len().max()
    NOME_PAD_LEN = max(MAX_NOME_LEN, 40) + 2

    output_buffer.write(f"RELATÓRIO DE PRESENÇA E LOCAIS - {titulo_relatorio}\n")
    output_buffer.write("-" * 140 + "\n")

    for secao, grupo in df_txt.groupby('Secao'):
        output_buffer.write(f"\n#####################################################\n")
        output_buffer.write(f"SEÇÃO: {secao}\n")
        output_buffer.write(f"TOTAL DE REGISTROS NA SEÇÃO: {len(grupo)}\n")
        output_buffer.write(f"#####################################################\n\n")

        # Pega a quantidade máxima de batidas dinamicamente
        colunas_batidas = [c for c in grupo.columns if isinstance(c, int)]

        cabecalho = f"| {'NOME DO FUNCIONÁRIO':<{NOME_PAD_LEN}} | {'DATA':<10} |"
        linha_sep = f"|{'-' * (NOME_PAD_LEN + 2)}|{'-' * 12}|"

        for col in colunas_batidas:
            cabecalho += f" {'BATIDA ' + str(col):<22} |"
            linha_sep += f"{'-' * 24}|"

        output_buffer.write(cabecalho + "\n")
        output_buffer.write(linha_sep + "\n")

        for _, item in grupo.iterrows():
            data_formatada = item['data_do_ponto'].strftime('%d/%m/%Y')
            linha_func = f"| {item['Nome']:<{NOME_PAD_LEN}} | {data_formatada:<10} |"
            
            for col in colunas_batidas:
                registro = str(item[col]) if pd.notna(item[col]) else "---"
                linha_func += f" {registro:<22} |"
            output_buffer.write(linha_func + "\n")

        output_buffer.write("\n")
else:
    output_buffer.write("Nenhum registro encontrado para a data/período especificado.\n")

nome_arquivo_saida_txt = f"Relatorio_Presenca_{string_data_arquivo}.txt"
with open(nome_arquivo_saida_txt, 'w', encoding='utf-8') as f:
    f.write(output_buffer.getvalue())

print(f"Relatório gerado com sucesso: {nome_arquivo_saida_txt}")