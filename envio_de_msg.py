'import pandas as pd
import pywhatkit as kit

planilha = r"C:\Users\Jonatha\Desktop\Jonatha\programação\Envio de Msg whatsapp automática.xlsx"
data = pd.read_excel(planilha, header=1)
cmensagem = 'Msg'
cDDD = 'Whtas'
cnumero =  'Unnamed: 3'
cnumero1 =  'Unnamed: 4'
link_dynms = 'www.e-inscricao.com/dynms/conference2024'
if cmensagem in data.columns and cDDD in data.columns:
    for i, f in data.iterrows():
        tel = f[cDDD]
        tel1 = f[cnumero]
        tel2 = f[cnumero1]
        if type(tel) == float or type(tel2) == float or type(tel1) == float:
            print(f'A coluna {cnumero} não está preenchida.')
        else:
            msg = f[cmensagem]
            nome = f['Nome']
            if type(nome) == float:
                nome = ''
            sobrenome = f['Sobrenome']
            if type(sobrenome) == float:
                sobrenome = ''
            if pd.notna(tel) and pd.notna(msg):
                try:
                    mensagem_final = 'Olá ' + nome + ' ' + sobrenome + msg + ' - ' + link_dynms
                    tel = '+' + str(tel).replace(" ", "").strip() + str(tel1).replace(" ", "").strip() + str(tel2).replace(" ", "").strip() 
                    print(mensagem_final)
                    #kit.sendwhatmsg_instantly(tel, mensagem_final) 
                except Exception as e:
                    print(f'Erro ao abrir {i}: {e}')
else:
    print(f'As colunas "{cmensagem}" e/ou "{cDDD}" não foram encontradas.')
