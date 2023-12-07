import pyautogui
import pandas as pd
import time
from time import sleep
from datetime import datetime

#Several notes about this script:

#The coordinates are set according to my notebook screen, so if you don't configure them, the script won't function as intended.
#I recommend having some knowledge about using PyAutoGUI and MouseInfo.
#Be very cautious when altering information in the spreadsheet. If any field names are changed or deleted, the script may encounter errors.
#That being said, I hope you enjoy and have fun!

#Algumas observações sobre este script
#- As coodernadas estão setadas de acordo com a tela do meu notebook, logo caso você não as configure o script não irá funcionar como deve
#- Recomendo que você tenha alguma noção sobre o uso do PyautoGUI e o mouseInfo
#- Muito cuidado ao alterar as informações na planilha, caso algum nome dos campos  presentes seja mudado ou apagado o scrip poderá apresentar erros
#- Dito isso, espero que você goste e se divirta


#Busca de arquivos na planilha ultilizando o Pandas
tabela_clientes = pd.read_excel("Formulario teste advbox.xlsx")

#Verifica os dados que ja foram lidos
dados_nao_lidos = tabela_clientes[tabela_clientes['Lido'] != 1]
#Buscando os dados dos clientes
for i in range(len(tabela_clientes["CPF/CNPJ (Apenas números)"])):

    #Verificação de cadastro através do CPF
    verificar = tabela_clientes["Lido"][i]
    if verificar == 1:
       i +=1
       continue
    else:

        #Preenche o CPF/CNJP
        cpf = tabela_clientes["CPF/CNPJ (Apenas números)"][i]
        cpf_str = str(cpf)
        pyautogui.click(384,485, duration=1)
        pyautogui.write(cpf_str)
         

        # Coleta o nome do usuário
        nome = tabela_clientes["Nome completo"][i]
        pyautogui.click(386,632, duration=1)
        pyautogui.write(nome)


        #Teste de scroll
        pyautogui.scroll(-700)

        #Origem
        origem = tabela_clientes['Origem '][i]
        pyautogui.click(392,201, duration=1)
        pyautogui.write(origem)


        # #Anotações Gerais
        ant_gerais = tabela_clientes['Anotações gerais'][i]
        pyautogui.click(389,345, duration=1)
        pyautogui.write(ant_gerais)

        # #Tipo de Pessoa (Fisica ou Juridica)
        tipo_pessoa = tabela_clientes["Pessoa do tipo"][i]
        if tipo_pessoa == "Fisica":
            pyautogui.click(388,494, duration=1)
            sleep(1)
        else:
            pyautogui.click(387,527, duration=1)
            sleep(1)

        # Preenchimento RG
        rg = tabela_clientes['RG (Apenas números)'][i]
        str_rg = str(rg)
        pyautogui.click(380,694, duration=1)
        pyautogui.write(str_rg)

        
        #Scroll
        pyautogui.scroll(-770)
        

        #Data de nascimento
        nascimento = tabela_clientes["Data de nascimento"][i]
        str_nascimento = str(nascimento)
        ano, mes, dia = str_nascimento.split("-")
        pyautogui.click(386,201, duration=1)
        pyautogui.write(dia[:2])
        pyautogui.write(mes)
        pyautogui.write(ano)

        # #Estado civil
        estd_civil = tabela_clientes['Estado civil'][i]
        pyautogui.click(382,353, duration=1)
        pyautogui.write(estd_civil)

        # #Preenchimento Profissão 
        profissao = tabela_clientes["Profissão"][i]
        pyautogui.click(382,506, duration=1)
        pyautogui.write(profissao)

        # #Teste de gênero do cliente
        sexo = tabela_clientes["Sexo"][i]
        if sexo == "Feminino":
            pyautogui.click(388,682, duration=1)
            sleep(1)
        else:
            pyautogui.click(388,646, duration=1)
            sleep(1)

        #scroll
        pyautogui.scroll(-750)
        
        #Nacionalidade
        naturalidade = tabela_clientes["Nacionalidade"][i]
        str_naturalidade = str(naturalidade)
        pyautogui.click(383,210, duration=1)
        pyautogui.write(str_naturalidade)


        # #Celular
        celular = tabela_clientes["Celular (Apenas digitos)"][i]
        str_celular = str(celular)
        pyautogui.click(385,361, duration=1)
        pyautogui.write(str_celular)

        # #Telefone
        telefone = tabela_clientes["Telefone"][i]
        str_telefone = str(telefone)
        pyautogui.click(385,512, duration=1)
        pyautogui.write(str_telefone)

        # #Email
        email = tabela_clientes["E-mail"][i]
        pyautogui.click(386,664, duration=1)
        pyautogui.write(email)

        #scroll
        pyautogui.scroll(-720)

        # CEP
        cep = tabela_clientes["CEP"][i]
        str_cep = str(cep)
        pyautogui.click(381,215, duration=1)
        pyautogui.write(str_cep)

        # #Pais
        pais = tabela_clientes["País"][i]
        pyautogui.click(382,364, duration=1)
        pyautogui.write(pais)


        # #Estado
        estado = tabela_clientes["Estado"][i]
        pyautogui.click(382,512, duration=1)
        pyautogui.write(estado)


        # #Cidade 
        cidade = tabela_clientes["Cidade"][i]
        pyautogui.click(384,660, duration=1)
        pyautogui.write(cidade)

        #scroll
        pyautogui.scroll(-710)


        #Endereço
        endereco = tabela_clientes["Endereço, número"][i]
        str_end = str(endereco)
        pyautogui.click(384,219, duration=1)
        pyautogui.write(str_end)


        #Bairro
        bairro = tabela_clientes["Bairro"][i]
        pyautogui.click(381,370, duration=1)
        pyautogui.write(bairro)

        # #PIS/PASEP
        pis = tabela_clientes["PIS/PASEP"][i]
        str_pis = str(pis)
        pyautogui.click(386,521, duration=1)
        pyautogui.write(str_pis)


        # #CTPS
        ctps = tabela_clientes["CTPS"][i]
        str_ctps = str(ctps)
        pyautogui.click(385,666, duration=1)
        pyautogui.write(str_ctps)

        #scroll
        pyautogui.scroll(-700)

        #CID
        cid = tabela_clientes["CID"][i]
        str_cid = str(cid)
        pyautogui.click(380,233, duration=1)
        pyautogui.write(str_cid)

        # #Nome da mãe
        nome_mae = tabela_clientes["Nome da mãe"][i]
        pyautogui.click(381,385, duration=1)
        pyautogui.write(nome_mae)

        # #Representante
        representante = tabela_clientes["Adicionar representante"][i]
        pyautogui.click(382,533, duration=1)
        pyautogui.write(representante)

        break

        #Enviar formulario 
        pyautogui.click(385,611, duration=1)

        # #salvando o formulário
        # tabela_clientes.loc[0:i,['Lido']] = 1
        # tabela_clientes.to_excel('Formulario teste advbox.xlsx', index=False)
        # break
    








