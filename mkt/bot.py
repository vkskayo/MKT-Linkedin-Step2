"""
WARNING:

Please make sure you install the bot with `pip install -e .` in order to get all the dependencies
on your Python environment.

Also, if you are using PyCharm or another IDE, make sure that you use the SAME Python interpreter
as your IDE.

If you get an error like:
```
ModuleNotFoundError: No module named 'botcity'
```

This means that you are likely using a different Python interpreter than the one used to install the bot.
To fix this, you can either:
- Use the same interpreter as your IDE and install your bot with `pip install -e .`
- Use the same interpreter as the one used to install the bot (`pip install -e .`)

Please refer to the documentation for more information at https://documentation.botcity.dev/
"""

from botcity.web import WebBot, Browser, By
import pandas as pd
from dotenv import load_dotenv
import os #provides ways to access the Operating System and allows us to read the environment variables
import openpyxl
from datetime import datetime
import time
import smtplib
# Import the email modules we'll need
import email.message




# Uncomment the line below for integrations with BotMaestro
# Using the Maestro SDK
# from botcity.maestro import *


class Bot(WebBot):
    def action(self, execution=None):
        # Uncomment to silence Maestro errors when disconnected
        # if self.maestro:
        #     self.maestro.RAISE_NOT_CONNECTED = False

   
   
        # Uncomment to change the default Browser to Firefox
        # self.browser = Browser.FIREFOX

        # Carregando variáveis de ambiente
        load_dotenv()
        caminho_chrome_driver = os.getenv("CAMINHO_CHROME_DRIVER")
        email = os.getenv("EMAIL")
        password = os.getenv("PASSWORD")
        caminho_arquivo_pessoas_queue = os.getenv("CAMINHO_ARQUIVO_PESSOAS_QUEUE")
        caminho_arquivo_pessoas = os.getenv("CAMINHO_ARQUIVO_PESSOAS")
        caminho_arquivo_empresas = os.getenv("CAMINHO_ARQUIVO_EMPRESAS")
        caminho_arquivo_relatorio_sucesso = os.getenv("CAMINHO_ARQUIVO_RELATORIO_SUCESSO")
        caminho_arquivo_base = os.getenv("CAMINHO_ARQUIVO_BASE")
        background = os.getenv("BACKGROUND")

        self.driver_path = caminho_chrome_driver
        # Configure whether or not to run on headless mode
        self.headless = background.upper() == 'TRUE'

        # Fetch the Activity ID from the task:
        # task = self.maestro.get_task(execution.task_id)
        # activity_id = task.activity_id
        print("Inicio processo: ")
        inicio = time.time()
        
        prepararArquivo(caminho_arquivo_pessoas,caminho_arquivo_pessoas_queue)
        login(self, email, password)
        self.wait(45000)
        extrairLinkPessoas(self, caminho_arquivo_empresas, caminho_arquivo_pessoas_queue, caminho_arquivo_relatorio_sucesso)
        extrairInfoPessoas(self, caminho_arquivo_pessoas_queue, caminho_arquivo_pessoas, caminho_arquivo_relatorio_sucesso)
        logout(self)
        integrarBase(caminho_arquivo_empresas, caminho_arquivo_pessoas, caminho_arquivo_base)
        duracao_segundos, duracao_formatada = finalizar_contagem_tempo(inicio)

        print(f"Tempo total de execução: {duracao_segundos} segundos ({duracao_formatada})")
        
        # Wait for 10 seconds before closing
        #self.wait(100000)

        # Stop the browser and clean up
        #self.stop_browser()
    
def login(self, credential_email, credential_password):
 
    # Opens the BotCity website.
    self.browse("https://www.linkedin.com")
    self.maximize_window()
    # Uncomment to mark this task as finished on BotMaestro
    # self.maestro.finish_task(
    #     task_id=execution.task_id,
    #     status=AutomationTaskFinishStatus.SUCCESS,
    #     message="Task Finished OK."
    # )

    email = self.find_element(selector="/html/body/main/section[1]/div/div/form/div[1]/div[1]/div/div/input", by = By.XPATH)
    password = self.find_element(selector="/html/body/main/section[1]/div/div/form/div[1]/div[2]/div/div/input", by = By.XPATH)
    loginButton = self.find_element(selector="/html/body/main/section[1]/div/div/form/div[2]/button", by = By.XPATH)

    if email and password and loginButton:
        email.send_keys(credential_email)
        password.send_keys(credential_password)
        loginButton.click()
    else:
        print("Não encontrado elemento de login, iniciando fluxo alternativo de login")
        botaoEntrar = self.find_element(selector="/html/body/main/section[1]/div/div/a", by = By.XPATH)
        botaoEntrar.click()
        alternativeEmail = self.find_element(selector='//*[@id="username"]', by = By.XPATH) 
        alternativePassword = self.find_element(selector='//*[@id="password"]', by = By.XPATH)
        alternativeLoginButton = self.find_element(selector='//*[@id="organic-div"]/form/div[3]/button', by = By.XPATH) 

        alternativeEmail.send_keys(credential_email)
        alternativePassword.send_keys(credential_password) 
        alternativeLoginButton.click()


def logout(self):
    botaoEu = self.find_element(selector="//span[contains(.,'Eu')]", by = By.XPATH)
    botaoEu.click()
    self.wait(2000)
    botaoSair = self.find_element(selector="//a[contains(.,'Sair')]", by = By.XPATH)
    botaoSair.click()

    botaoSairDefinitivo = self.find_element(selector="/html/body/div[3]/div/div/div[2]/section/footer/button[2]", by = By.XPATH)

    if botaoSairDefinitivo:
        botaoSairDefinitivo.click()
    

def extrairLinkPessoas(self, caminho_arquivo_empresas, caminho_arquivo_pessoas_queue, caminho_arquivo_relatorio_sucesso):

    df = pd.read_excel(caminho_arquivo_pessoas_queue) 
    df_queue = pd.read_excel(caminho_arquivo_empresas)

    for index, row in df_queue.iterrows():
        try:
            linkedinEmpresa = row['linkedinEmpresa']
            numUsuariosAssociados = row['TamanhoEmpresaUsuarios']
            extrairBool = row["Extrair"]
            numeroExtracao = row["NumeroExtracao"]

            # Extrair link apenas das empresas com o valor na coluna "Extrair"
            if extrairBool != 1:
                print("Business Rule Exception:: Pulando a extração da empresa de link: " + linkedinEmpresa + "\n")
                print("Não foi solicitada sua extração.")
                continue

            numClicksScroll = round(numUsuariosAssociados * 0.75) # Fazer mais testes e ajustar essa fórmula 

            # Garantir que o mínimo de scroll seja 100 clicks
            if numClicksScroll < 100:
                numClicksScroll = 100

            colunaLinkPessoas = []
            self.browse(linkedinEmpresa)
            self.wait(6000)
            botaoPessoas = self.find_element(selector="//a[contains(.,'Pessoas')]", by = By.XPATH)
            botaoPessoas.click()
            self.scroll_down(clicks=numClicksScroll) 
            listaLinkPessoas = self.find_elements(selector='//a[contains(@aria-label,"Ver perfil de")]', by = By.XPATH)
            
            for elem in listaLinkPessoas:
                colunaLinkPessoas.append(elem.get_attribute("href"))
            
            if df.empty:
                df = pd.DataFrame({
                'linkPessoa': colunaLinkPessoas,
                'linkedinEmpresa': linkedinEmpresa,
                'Status': "Não processado"
                })
        
                df = filtrarUsuariosProcessados(caminho_arquivo_relatorio_sucesso, df)
                
                if numeroExtracao > len(df):
                # Se numeroMaior for maior, pegar todas as linhas do DataFrame
                    df.loc[:, 'Status'] = 'A processar'
                else:
                # Caso contrário, pegar até o numeroMaior
                    df.loc[:numeroExtracao - 1, 'Status'] = 'A processar'
            else:
                df2 = pd.DataFrame({
                'linkPessoa': colunaLinkPessoas,
                'linkedinEmpresa': linkedinEmpresa,
                 'Status': "Não processado"
                })

                df2 = filtrarUsuariosProcessados(caminho_arquivo_relatorio_sucesso, df2)

                if numeroExtracao > len(df2):
                # Se numeroMaior for maior, pegar todas as linhas do DataFrame
                    df2.loc[:, 'Status'] = 'A processar'
                else:
                # Caso contrário, pegar até o numeroMaior
                    df2.loc[:numeroExtracao - 1, 'Status'] = 'A processar'
                df = pd.concat([df, df2], ignore_index=True)

        except Exception as e:
            print(e)
            print("\nNão foi possível realizar a extração dos links das pessoas da empresa de link: " + linkedinEmpresa)
    df.to_excel(caminho_arquivo_pessoas_queue, index=False) 
            
def extrairInfoPessoas(self, caminho_arquivo_pessoas_queue, caminho_arquivo_pessoas, caminho_arquivo_relatorio_sucesso):
    df_pessoas = pd.read_excel(caminho_arquivo_pessoas_queue)
    df = pd. DataFrame()

    if df_pessoas.empty:
        # Enviar email de notificação indicando que a fila de execução está vazia !!!!!!!!
        raise Exception('A fila de execução de extração de pessoas está vazia')

    for index, row in df_pessoas.iterrows():
        try:
            linkedinEmpresa = row['linkedinEmpresa']
            linkPessoa = row['linkPessoa']
            status = row['Status']
       
            if status == 'Não processado':
                continue

            # Inicialização variáveis da base de pessoas
            nome = 'Não informado'
            telefone = 'Não informado'
            cargo = 'Não informado'
            email = 'Não informado'
            cargo_experiencia = 'Não informado'
            #Inicialização do dataframe da base de pessoas

            if df.empty:
                df = pd.DataFrame(columns=['linkPessoa', 'linkedinEmpresa', 'Nome', 'Cargo', 'Telefone',  'Email'])

            #Início do processo de extração dos dados das pessoas
            self.browse(linkPessoa)
            self.wait(2000)

            isMessageSectionMinimized = False
            secaoMensagem = self.find_element(selector="msg-overlay-list-bubble--is-minimized", by = By.CLASS_NAME)

            if secaoMensagem:
                isMessageSectionMinimized = True

            if not isMessageSectionMinimized:

                botaoFecharMensagens = self.find_element(selector='//header[@class="msg-overlay-bubble-header"]', by = By.XPATH)
                botaoFecharMensagens.click()

            nomeElement = self.find_element(selector='h1', by = By.TAG_NAME)  
            cargoElement = self.find_element(selector="text-body-medium", by = By.CLASS_NAME)

            experienciaElement = self.find_element(selector="//*[@id='profile-content']/div/div[2]/div/div/main/section[5]", by = By.XPATH)
            experienciaElement2 = self.find_element(selector="//*[@id='profile-content']/div/div[2]/div/div/main/section[6]", by = By.XPATH)
            experienciaElement3 = self.find_element(selector="//*[@id='profile-content']/div/div[2]/div/div/main/section[4]", by = By.XPATH)
            experienciaElement4 = self.find_element(selector="//*[@id='profile-content']/div/div[2]/div/div/main/section[3]", by = By.XPATH)

            experienciaList = ['None']
            experienciaList2 = ['None']
            experienciaList3 = ['None']
            experienciaList4 = ['None']

            if experienciaElement:
                experienciaList = experienciaElement.text.split("\n")
            
            if experienciaElement2:
                experienciaList2 = experienciaElement2.text.split("\n")

            if experienciaElement3:
                experienciaList3 = experienciaElement3.text.split("\n")
            
            if experienciaElement4:
                experienciaList4 = experienciaElement4.text.split("\n")

            if experienciaList[0] == "Experiência":
                cargo_experiencia = experienciaList[2]
            elif experienciaList2[0] == "Experiência":
                cargo_experiencia = experienciaList2[2]
            
            elif experienciaList3[0] == "Experiência":
                cargo_experiencia = experienciaList3[2]
       
            elif experienciaList4[0] == "Experiência":
                cargo_experiencia = experienciaList4[2]
            
            else:
                print("Não foi possível extrair o cargo_experiencia")
                

            nome = nomeElement.text
            cargo = cargoElement.text

            botaoInformacoesDeContato = self.find_element(selector="//a[contains(.,'Informações de contato')]", by = By.XPATH)
            botaoInformacoesDeContato.click()
            self.wait(2000)
            # Pegar email e telefone

            secaoInformacoesDeContato = self.find_element(selector="pv-profile-section__section-info", by = By.CLASS_NAME)

            listInformacoesDeContato = secaoInformacoesDeContato.text.split("\n")

            for idx,elem in enumerate(listInformacoesDeContato):
                if elem == "E-mail":
                    email = listInformacoesDeContato[idx + 1]
                if elem == "telefone":
                    telefone = listInformacoesDeContato[idx + 1]

            nova_linha = {
            'linkPessoa': linkPessoa,
            'linkedinEmpresa': linkedinEmpresa,
            'Nome': nome,
            'Cargo': cargo,
            'Telefone' : telefone,
            'Email': email,
            'CargoExperiencia': cargo_experiencia
            }

            df = pd.concat([df, pd.DataFrame([nova_linha])], ignore_index=True)
            df_pessoas.at[index, 'Status'] = "Sucesso"
        except Exception as e:
            print(e)
            print("\nNão foi possível realizar a extração das informações da pessoa de link: " + linkPessoa)
            df_pessoas.at[index, 'Status'] = "Falha"
        
    self.tab()
    self.enter()    
    # Escrever o DataFrame no arquivo Excel
    df.to_excel(caminho_arquivo_pessoas, index=False) 
    # Atualizar o relatório de execução
    relatorio = df_pessoas[(df_pessoas['Status'] == "Sucesso") | (df_pessoas['Status'] == "Falha")]
    relatorio.to_excel(caminho_arquivo_pessoas_queue, index=False)

    relatorio_sucesso_base = pd.read_excel(caminho_arquivo_relatorio_sucesso)
    relatorio_sucesso = df_pessoas[df_pessoas['Status'] == "Sucesso"]
    relatorio_sucesso = relatorio_sucesso.rename(columns={'linkPessoa': 'linkPessoa_', 'linkedinEmpresa': 'linkedinEmpresa_', 'Status': 'Status_'})
    relatorio_sucesso_base = pd.concat([relatorio_sucesso_base, relatorio_sucesso], ignore_index=True)

    relatorio_sucesso_base.to_excel(caminho_arquivo_relatorio_sucesso, index=False)


def isElementEnabled(self, seletor):
    element = self.find_element(selector=seletor, by = By.XPATH)

    if element:
        if element.is_enabled():
            return True
        else:
            print("O elemento não está ativo")
            return False
    else:
        print("O elemento não foi encontrado")
        return False
    
def prepararArquivo(caminho_arquivo_pessoas,caminho_arquivo_pessoas_queue):
 
    listOfFiles = [caminho_arquivo_pessoas,caminho_arquivo_pessoas_queue]

    wb = openpyxl.Workbook()

    for file in listOfFiles:
        wb.save(file)

def integrarBase(caminho_arquivo_empresas, caminho_arquivo_pessoas, caminho_arquivo_base):
    try:
        df_pessoas = pd.read_excel(caminho_arquivo_pessoas)
        df_empresas = pd.read_excel(caminho_arquivo_empresas, sheet_name="Base de dados")
        df_join = pd.merge(df_pessoas, df_empresas, on='linkedinEmpresa', how='inner')
        base_path = caminho_arquivo_base + "baseLinkedin-" + datetime.now().strftime("%m-%d-%Y%H-%M-%S") + ".xlsx"
        df_join.to_excel(base_path, index=False) 
        print("Base criado com sucesso.")
    except Exception as e:
        #Enviar email
        print(e)
        print("Erro na etapa de integração das bases. Por favor verificar com o suporte")
        
def finalizar_contagem_tempo(inicio):
    fim = time.time()
    duracao_segundos = fim - inicio
    duracao_formatada = time.strftime("%H:%M:%S", time.gmtime(duracao_segundos))
    return duracao_segundos, duracao_formatada

def filtrarUsuariosProcessados(caminho_arquivo_relatorio_sucesso, df):
    df_pessoas = df
    df_relatorio = pd.read_excel(caminho_arquivo_relatorio_sucesso)
    merged = df_pessoas.merge(df_relatorio, left_on=['linkPessoa', 'linkedinEmpresa'], right_on=['linkPessoa_', 'linkedinEmpresa_'], how='outer', indicator=True)
    filtered_df = merged[merged['_merge'] == 'left_only']
    result_df = filtered_df[df_pessoas.columns]
    result_df.reset_index(drop=True, inplace=True)
    return result_df

def enviarEmail():
    corpo_email = """<p>hello</p>"""
    msg = email.message.Message()
    msg['Subject'] = "Teste"
    msg['From'] = "vkskayo@gmail.com"
    msg['To'] = "vkskayo@gmail.com"

    password = 'pxedhydfirrnszfs'
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(corpo_email)

    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()

    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
    print("Email enviado")


if __name__ == '__main__':
    Bot.main()
