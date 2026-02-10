import pandas as pd
from datetime import datetime

from datetime import datetime, timedelta

hoje = datetime.now().date()
ontem = hoje - timedelta(days=1)

# Se hoje for segunda-feira, pega sexta, sábado e domingo
if hoje.weekday() == 0:  
    data_inicio = hoje - timedelta(days=3)  # sexta-feira
    data_fim = ontem  # domingo
else:
    data_inicio = ontem
    data_fim = ontem

arquivos = ['Ponto_Algodoeira.txt', 'Ponto_Escritorio.txt', 'Ponto_Sede.txt', 'Ponto_Secador.txt']
dfs_ponto = []
for arquivo in arquivos:
    df = pd.read_csv(arquivo, header=None, names=['linha_completa'], encoding='latin-1')
    df['origem'] = arquivo
    dfs_ponto.append(df)
df_ponto_raw = pd.concat(dfs_ponto, ignore_index=True)

df_funcionarios = pd.read_excel(r'C:\Users\fazin\OneDrive\Documents\Nisio\analiseponto\Funcionarios.xlsx')
df_funcionarios['NIT_STR'] = df_funcionarios['NIT'].astype(str)
df_funcionarios.rename(columns=lambda x: x.strip(), inplace=True)

nit_to_nome = df_funcionarios.set_index('NIT_STR')['Nome'].to_dict()
nit_to_secao = df_funcionarios.set_index('NIT_STR')['Secao'].to_dict()
nit_to_cpf = df_funcionarios.set_index('NIT_STR')['CPF'].to_dict()

registros_validos = []
for _, row in df_ponto_raw.iterrows():
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
            if data_inicio <= data_do_ponto <= data_fim:
                hora_str = linha_completa[18:22]
                data_hora = datetime.strptime(data_str + hora_str, '%d%m%Y%H%M')
                registros_validos.append({
                    'NIT': nit_encontrado,
                    'Nome': nit_to_nome.get(nit_encontrado, 'N/A'),
                    'Secao': nit_to_secao.get(nit_encontrado, 'N/A'),
                    'CPF': nit_to_cpf.get(nit_encontrado, 'N/A'),
                    'data_hora': data_hora,
                    'data': data_do_ponto
                })
        except ValueError:
            continue

df_ponto_validado = pd.DataFrame(registros_validos)

if not df_ponto_validado.empty:
    relatorio = []
    for (nit, dia), grupo in df_ponto_validado.groupby(['NIT', 'data']):
        nome = nit_to_nome.get(nit, 'N/A')
        secao = nit_to_secao.get(nit, 'N/A')
        cpf = nit_to_cpf.get(nit, 'N/A')

        batidas = grupo['data_hora'].sort_values().tolist()
        batidas_str = [b.strftime('%H:%M') for b in batidas]

        if len(batidas) >= 4:
            saida_almoco = batidas[1]
            retorno_almoco = batidas[2]
            intervalo_almoco = (retorno_almoco - saida_almoco).total_seconds() / 60

            if intervalo_almoco < 60: 
                relatorio.append({
                    'Nome': nome,
                    'CPF': cpf,
                    'Data': dia.strftime('%d/%m/%Y'),
                    'Secao': secao,
                    'Batidas': " ".join(batidas_str),
                    'Almoco': f"{int(intervalo_almoco//60):02d}:{int(intervalo_almoco%60):02d}"
                })

    df_relatorio = pd.DataFrame(relatorio, columns=['Nome','CPF','Data','Secao','Batidas','Almoco'])

    if not df_relatorio.empty:
        df_relatorio.to_excel("Relatorio_Almoco.xlsx", index=False)

        with open("Relatorio_Almoco.txt", 'w', encoding='utf-8') as f:
            f.write("RELATÓRIO DE INTERVALO DE ALMOÇO < 1h\n")
            f.write("="*95 + "\n")
            f.write(f"{'NOME':<30} {'CPF':<15} {'DATA':<12} {'SEÇÃO':<15} {'BATIDAS':<25} {'ALMOÇO':<8}\n")
            f.write("-"*95 + "\n")
            for _, item in df_relatorio.iterrows():
                f.write(f"{item['Nome']:<30} {item['CPF']:<15} {item['Data']:<12} {item['Secao']:<15} {item['Batidas']:<25} {item['Almoco']:<8}\n")

        print("📂 Relatórios gerados: Relatorio_Almoco.xlsx e Relatorio_Almoco.txt")
    else:
        print("✅ Nenhum funcionário fez almoço menor que 1h.")
else:
    print("⚠️ Nenhum registro válido encontrado.")
