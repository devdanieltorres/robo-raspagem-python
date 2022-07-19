import datetime
import time
import pandas as pd
import selenium.webdriver.support
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains

# retira a mensagem de erro ao inserir comentarios raspados no excel
pd.set_option('mode.chained_assignment', None)

# ----------SELEÇÃO DE CARDS--------------

# Carrego a planilha da exportação do planner
excel = pd.read_excel("DIPMAtividades.xlsx", index_col=0)

# ler a quantidade de linhas da tabela
QtdLinhas = len(excel)

# ----------------SC1----------------------

dataHoje = datetime.date.today()
datetime.datetime.__format__(dataHoje, "%d/%m/%Y")
excel['Data de início'] = pd.to_datetime(excel['Data de início'], format='%d/%m/%Y', utc=True).dt.date

excel.loc[(excel["Nome do Bucket"] == "Em andamento"), "SC1"] = 'NÃO'
excel.loc[(excel["Nome do Bucket"] == "Feito"), "SC1"] = 'NÃO'

excel.loc[(excel["Nome do Bucket"] == "Backlog") & (excel['Data de início'] <= dataHoje), "SC1"] = 'SIM'
excel.loc[(excel["Nome do Bucket"] == "Backlog") & (excel['Data de início'] > dataHoje), "SC1"] = 'NÃO'

excel.loc[(excel["Nome do Bucket"] == "A fazer") & (excel['Data de início'] <= dataHoje), "SC1"] = 'SIM'
excel.loc[(excel["Nome do Bucket"] == "A fazer") & (excel['Data de início'] > dataHoje), "SC1"] = 'NÃO'

excel.loc[(excel["Nome do Bucket"] == "Bloqueados") & (excel['Data de início'] <= dataHoje), "SC1"] = 'SIM'
excel.loc[(excel["Nome do Bucket"] == "Bloqueados") & (excel['Data de início'] > dataHoje), "SC1"] = 'NÃO'

# ----------------SC2----------------------

excel.loc[(excel["Nome do Bucket"] == "Em andamento"), "SC2"] = 'NÃO'
excel.loc[(excel["Nome do Bucket"] == "Feito"), "SC2"] = 'NÃO'

excel.loc[(excel["Nome do Bucket"] == "Backlog") & (excel['Progresso'] == "Em andamento"), "SC2"] = 'SIM'
excel.loc[(excel["Nome do Bucket"] == "Backlog") & (excel['Progresso'] != "Em andamento"), "SC2"] = 'NÃO'

excel.loc[(excel["Nome do Bucket"] == "A fazer") & (excel['Progresso'] == "Em andamento"), "SC2"] = 'SIM'
excel.loc[(excel["Nome do Bucket"] == "A fazer") & (excel['Progresso'] != "Em andamento"), "SC2"] = 'NÃO'

excel.loc[(excel["Nome do Bucket"] == "Bloqueados") & (excel['Progresso'] == "Em andamento"), "SC2"] = 'SIM'
excel.loc[(excel["Nome do Bucket"] == "Bloqueados") & (excel['Progresso'] != "Em andamento"), "SC2"] = 'NÃO'

# ----------------SC3------------------------

excel.loc[excel["Nome do Bucket"] == "Em andamento", "SC3"] = "SIM"
excel.loc[excel["Nome do Bucket"] != "Em andamento", "SC3"] = "NÃO"

# ----------------SC4------------------------

dataInicio = "01/01/2022"
dataInicio = pd.to_datetime(dataInicio, format='%d/%m/%Y', utc=True).date()

dataFim = "30/03/2022"
dataFim = pd.to_datetime(dataFim, format='%d/%m/%Y', utc=True).date()

excel['Data de conclusão'] = pd.to_datetime(excel['Data de conclusão'], format='%d/%m/%Y', utc=True).dt.date

excel.loc[(excel["Nome do Bucket"] == "Feito") & (excel['Data de conclusão'] >= dataInicio) & (excel['Data de conclusão'] <= dataFim), "SC4"] = "SIM"
excel.loc[(excel["Nome do Bucket"] == "Feito") & (pd.isnull(excel['Data de conclusão']) == True), "SC4"] = "SIM"
excel.loc[(excel["Nome do Bucket"] == "Feito") & (excel['Data de conclusão'] < dataInicio) | (excel['Data de conclusão'] > dataFim), "SC4"] = "NÃO"
excel.loc[(excel["Nome do Bucket"] != "Feito"), "SC4"] = "NÃO"

# # ----------------Validações de preenchimento------------------------
# ----------------VP1------------------------

excel.loc[excel["Nome do Bucket"] == "Backlog", "VP1"] = 1
excel.loc[excel["Nome do Bucket"] != "Backlog", "VP1"] = 0

# ----------------VP2------------------------

excel.loc[excel["Nome do Bucket"] == "A fazer", "VP2"] = 1
excel.loc[excel["Nome do Bucket"] != "A fazer", "VP2"] = 0

# ----------------VP4------------------------

excel.loc[(excel["Nome do Bucket"] == "Em andamento") & ((pd.isnull(excel['Data de início']) == True) | (excel['Data de início'] > dataHoje) |
          (pd.isnull(excel['Data de conclusão']) == True) | (excel['Progresso'] != "Em andamento")), "VP4"] = 1

excel.loc[(excel["Nome do Bucket"] == "Em andamento") & (pd.isnull(excel['Data de início']) != True) & (excel['Data de início'] < dataHoje) &
          (pd.isnull(excel['Data de conclusão']) != True) & (excel['Progresso'] == "Em andamento"), "VP4"] = 0

excel.loc[(excel["Nome do Bucket"] != "Em andamento"), "VP4"] = 0

# ----------------VP5------------------------

excel.loc[(excel["Nome do Bucket"] == "Feito") & ((pd.isnull(excel['Data de início']) == True) | (excel['Data de início'] > dataHoje) |
          (pd.isnull(excel['Data de conclusão']) == True) | (excel['Data de conclusão'] > dataHoje) | (excel['Progresso'] != "Concluída")), "VP5"] = 1

excel.loc[(excel["Nome do Bucket"] == "Feito") & (pd.isnull(excel['Data de início']) != True) & (excel['Data de início'] < dataHoje) &
          (pd.isnull(excel['Data de conclusão']) != True) & (excel['Data de conclusão'] < dataHoje) & (excel['Progresso'] == "Concluída"), "VP5"] = 0

excel.loc[(excel["Nome do Bucket"] != "Feito"), "VP5"] = 0

# ----------Quantidade de erros de preenchimento------------------

excel['erros'] = excel['VP1'] + excel['VP2'] + excel['VP4'] + excel['VP5']

# ----------------Filtro de cards------------------------

pd_atividades = excel.loc[(excel['SC1'] == 'SIM') | (excel['SC2'] == 'SIM') | (excel['SC3'] == 'SIM') | (excel['SC4'] == 'SIM')]
pd_atividades = pd_atividades.loc[(pd_atividades['erros'] == 0)]
Feitos = excel.loc[excel["Nome do Bucket"] == 'Feito']

pd_atividades.to_excel("PlanilhaPreenchida.xlsx")

# ------------COMEÇA A AUTENTICAÇÃO------------------

# Carrego a planilha para inserir comentarios
excelComentarios = pd.read_excel("Comentarios.xlsx", index_col=0)

# Carrega o driver para abrir o navegador
s = Service("chromedriver.exe")
driver = webdriver.Chrome(service=s)

# deixa o navegador em tela cheia
driver.maximize_window()

# define e abre a URL do Planner DIPP Atividades
driver.get("https://tasks.office.com/dataprev.gov.br/pt-BR/Home/Planner/#/plantaskboard?groupId=a760b2ac-8adc-492a-853d-d16b9f08dd74&planId=CYRfUvUwVkGwLJ0ih2B3lGUADUn5")

# aguardar
time.sleep(5)

# seleciona o elemento Login para digitar o usuário
elemLogin = driver.find_element(by=By.ID, value="i0116")

# preenche o campo login com o username da credencial recuperado anteriormente e aperta tecla ENTER
elemLogin.send_keys("AQUI VAI O USUARIO" + Keys.ENTER)

# aguardar
time.sleep(5)

# seleciona o elemento senha
elemPass = driver.find_element(by=By.ID, value="i0118")

# preenche o campo senha
elemPass.send_keys("AQUI VAI A SENHA" + Keys.ENTER)

# aguardar
time.sleep(5)

# seleciona o elemento para continuar conectado "SIM"
elemConectado = driver.find_element(by=By.ID, value="idSIButton9")

# executa o click no botão "SIM"
elemConectado.click()

# aguarda carregar o planner
driver.implicitly_wait(30)

# fechar menu lateral
elemFecharMenu = driver.find_element(by=By.CLASS_NAME, value='header')
elemFecharMenu.click()

# aguardar
WebDriverWait(driver, 5)

# mover a rolagem horizonal para exibir o campo Tarefas feitas
elemIrFeitos = driver.find_element(by=By.ID, value="6resMGX0u0aVUTWevh62-WUAOHmk")
WebDriverWait(driver, 7)
actions = ActionChains(driver)
actions.move_to_element(elemIrFeitos).perform()
actions.perform()

# ----------Faz a rolagem até o fim de BACKLOG-----------------------------------------------------------------
WebDriverWait(driver, 5)
elemIrBacklog = driver.find_element(by=By.ID, value="IpPPoxpx30mKL1yvPFMffWUAHaAO")
ActionChains(driver).move_to_element(elemIrBacklog).perform()
WebDriverWait(driver, 7)
elemFinalBackulog = elemIrBacklog.find_element(by=By.CLASS_NAME, value="bottomDropZone")
WebDriverWait(driver, 7)
ActionChains(driver).move_to_element(elemIrBacklog).perform()

# ----------Faz a rolagem até o fim de A fazer-----------------------------------------------------------------
WebDriverWait(driver, 5)
elemIrAFazer = driver.find_element(by=By.ID, value="7NgLdv6IGkCG-SdF4-R38WUAJqm-")
ActionChains(driver).move_to_element(elemIrAFazer).perform()
WebDriverWait(driver, 7)
elemFinalAFazer = elemIrAFazer.find_element(by=By.CLASS_NAME, value="bottomDropZone")
WebDriverWait(driver, 7)
ActionChains(driver).move_to_element(elemFinalAFazer).perform()

# ----------Faz a rolagem até o fim de Em andamento-----------------------------------------------------------------
contadorA = 0
while contadorA <= 5:
    WebDriverWait(driver, 5)
    elemIrEmAndamento = driver.find_element(by=By.ID, value="eKrkUsGBR02dvCi3zgXVKGUAFhhg")
    ActionChains(driver).move_to_element(elemIrEmAndamento).perform()
    WebDriverWait(driver, 7)
    elemFinalEmAnda = elemIrEmAndamento.find_element(by=By.CLASS_NAME, value="bottomDropZone")
    WebDriverWait(driver, 7)
    ActionChains(driver).move_to_element(elemFinalEmAnda).perform()
    contadorA += 1

# ----------Faz a rolagem até o fim de Bloqueado-----------------------------------------------------------------
WebDriverWait(driver, 5)
elemIrBloqueado = driver.find_element(by=By.ID, value="qeGI_YEH3EeY6ayIoMXKz2UAHQ_P")
ActionChains(driver).move_to_element(elemIrBloqueado).perform()
WebDriverWait(driver, 7)
elemFinalBloqueado = elemIrBloqueado.find_element(by=By.CLASS_NAME, value="bottomDropZone")
WebDriverWait(driver, 7)
ActionChains(driver).move_to_element(elemFinalBloqueado).perform()

# exibir o campo Tarefas feitas
elemExpandirFeitos = driver.find_element(by=By.CLASS_NAME, value="sectionToggleButton")
elemExpandirFeitos.click()
time.sleep(10)

# iterar entre os itens da coluna Feitos Planner
contador = 0
while contador <= 30:
    driver.implicitly_wait(15)
    FinalFeitos = driver.find_element(by=By.XPATH, value='//*[@id="6resMGX0u0aVUTWevh62-WUAOHmk"]/div/div[2]/div[2]/div[2]')
    ActionChains(driver).move_to_element(FinalFeitos).perform()
    contador += 1

# aguardar
time.sleep(15)

# ----------------FIM DA AUTENTICAÇÃO----------------------

# Conta a linha
contadorLinha = 0

# Conta o numero do card
contadorCard = 0

# Percorre as atividades para obter os comentarios via raspagem
for index, row in pd_atividades.iterrows():
    # aguardar
    time.sleep(5)
    print(index)
    # clica no card para ser possível capturar os comentários
    try:
        driver.implicitly_wait(180)
        element = driver.find_element(By.ID, value=index)
        driver.implicitly_wait(15)
        ActionChains(driver).move_to_element(element).perform()
        driver.implicitly_wait(15)
        element.click()

        # aguarda o card ser carregado
        time.sleep(10)

        # carrega a div que contem todos os comentários
        driver.implicitly_wait(15)
        lista_comentarios = driver.find_elements(by=By.CLASS_NAME, value="commentCard")

        # contador de comentarios por card
        contadorComCard = 0

        # acrescenta o numero do card
        contadorCard += 1

        # vai rodar todos os comentarios na lista de comentarios
        for comentario in lista_comentarios:
                # pega o usuário do comentário
                usuario = comentario.find_element(by=By.CLASS_NAME, value="commenterName").text

                # pega o data e hora do comentário
                data_hora = comentario.find_element(by=By.CLASS_NAME, value="timestamp").text

                # pega o texto do do comentário
                texto_comentario = comentario.find_element(by=By.CLASS_NAME, value="content").text

                # acrescenta a quantidade de linhas
                contadorLinha += 1
                # acrescenta a quantidade de comentarios em cada card
                contadorComCard += 1

                # tira partes desnescessarias da data
                data_hora = data_hora.replace(" de ", "-")
                data_hora = data_hora.replace(" às ", " ")

                # CONVERTER TODOS OS MESES PARA NUMEROS
                data_hora = data_hora.replace("janeiro", "1")
                data_hora = data_hora.replace("fevereiro", "2")
                data_hora = data_hora.replace("março", "3")
                data_hora = data_hora.replace("abril", "4")
                data_hora = data_hora.replace("maio", "5")
                data_hora = data_hora.replace("junho", "6")
                data_hora = data_hora.replace("julho", "7")
                data_hora = data_hora.replace("agosto", "8")
                data_hora = data_hora.replace("setembro", "9")
                data_hora = data_hora.replace("outubro", "10")
                data_hora = data_hora.replace("novembro", "11")
                data_hora = data_hora.replace("dezembro", "12")

                # formata a data para ser possivel realizar o calculo
                data_hora = pd.to_datetime(data_hora, format='%d-%m-%Y %H:%M', utc=True).date()

                # apresenta na tela (usuario, data_hora, texto_comentario, contadorLinha) apenas para acompanhamento
                print(usuario, data_hora, texto_comentario, contadorLinha, sep=" --- ")

                # insere os dados na planilha
                excelComentarios['Identificação da tarefa'][contadorLinha] = index
                excelComentarios['Usuario'][contadorLinha] = usuario
                excelComentarios['Data Hora'][contadorLinha] = data_hora
                excelComentarios['Comentario'][contadorLinha] = texto_comentario
                excelComentarios['Comentarios por card'][contadorLinha] = contadorComCard
                excelComentarios['Card'][contadorLinha] = contadorCard
                excelComentarios['Contador comentarios'][contadorLinha] = contadorLinha

        # clicar para fechar o card exibido atualmente no modal
        driver.implicitly_wait(90)
        botaoFechar = driver.find_element(by=By.CSS_SELECTOR, value="button[aria-label='Fechar caixa de diálogo']")
        webdriver.ActionChains(driver).click(botaoFechar).perform()

        # contadores para acessar a celula certa da planilha
        contador1 = 0
        contador2 = 1
        # coleta a quantidade de linhas na planilha
        QtdLinhasTeste = contadorLinha - 1
        while contador1 < QtdLinhasTeste:
            contador1 += 1
            contador2 += 1
            # coleta a data do comentario atual e a data do comentario anterior
            data1 = excelComentarios['Data Hora'][contador1]
            data2 = excelComentarios['Data Hora'][contador2]
            # calcula a distancia em dias de um comentario para o outro
            resultado = data1 - data2
            # marca como certo o ultimo comentario do ultimo card da planilha
            excelComentarios.loc[(excelComentarios['Contador comentarios'] == contadorLinha), "Teste"] = 'Certo'
            # marca como errado o card que esta fora do prazo
            excelComentarios.loc[(resultado.days > 3) & (excelComentarios['Contador comentarios'] == contador1), "Teste"] = 'Errado'
            # verificar se é o ultimo comentario do card
            card1 = excelComentarios['Comentarios por card'][contador1]
            card2 = excelComentarios['Comentarios por card'][contador2]
            excelComentarios.loc[(resultado.days > 3) & (excelComentarios['Contador comentarios'] == contador1) & (card2 <= card1), "Teste"] = 'Certo'
            # marca como certo o card que esta dentro do prazo
            excelComentarios.loc[(resultado.days <= 3) & (excelComentarios['Contador comentarios'] == contador1), "Teste"] = 'Certo'

        # gera a planilha com a validação de comentarios
        excelComentarios.to_excel(r"C:\Users\DanielTorres\PycharmProjects\pythonExercicios\crawler\ComentariosPreenchidos.xlsx")

    except selenium.common.exceptions.NoSuchElementException:
        # Abre uma nova aba e vai para o site do SO
        driver.execute_script("window.open('https://tasks.office.com/dataprev.gov.br/pt-BR/Home/Planner/#/plantaskboard?groupId=a760b2ac-8adc-492a-853d-d16b9f08dd74&planId=CYRfUvUwVkGwLJ0ih2B3lGUADUn5', '_blank')")
        driver.switch_to.window(driver.window_handles[1])
        time.sleep(15)
        # fechar menu lateral
        elemFecharMenu = driver.find_element(by=By.CLASS_NAME, value='header')
        elemFecharMenu.click()

        # ----------Faz a rolagem até o fim de Feitos-----------------------------------------------------------------
        WebDriverWait(driver, 5)
        # mover a rolagem horizonal para exibir o campo Tarefas feitas
        elemIrFeitos = driver.find_element(by=By.ID, value="6resMGX0u0aVUTWevh62-WUAOHmk")
        WebDriverWait(driver, 7)
        actions = ActionChains(driver)
        actions.move_to_element(elemIrFeitos).perform()
        actions.perform()

        # ----------Faz a rolagem até o fim de BACKLOG-----------------------------------------------------------------
        WebDriverWait(driver, 5)
        elemIrBacklog = driver.find_element(by=By.ID, value="IpPPoxpx30mKL1yvPFMffWUAHaAO")
        ActionChains(driver).move_to_element(elemIrBacklog).perform()
        WebDriverWait(driver, 7)
        elemFinalBackulog = elemIrBacklog.find_element(by=By.CLASS_NAME, value="bottomDropZone")
        WebDriverWait(driver, 7)
        ActionChains(driver).move_to_element(elemIrBacklog).perform()

        # ----------Faz a rolagem até o fim de A fazer-----------------------------------------------------------------
        WebDriverWait(driver, 5)
        elemIrAFazer = driver.find_element(by=By.ID, value="7NgLdv6IGkCG-SdF4-R38WUAJqm-")
        ActionChains(driver).move_to_element(elemIrAFazer).perform()
        WebDriverWait(driver, 7)
        elemFinalAFazer = elemIrAFazer.find_element(by=By.CLASS_NAME, value="bottomDropZone")
        WebDriverWait(driver, 7)
        ActionChains(driver).move_to_element(elemFinalAFazer).perform()

        # ----------Faz a rolagem até o fim de Em andamento-----------------------------------------------------------------
        contadorA = 0
        while contadorA <= 5:
            WebDriverWait(driver, 5)
            elemIrEmAndamento = driver.find_element(by=By.ID, value="eKrkUsGBR02dvCi3zgXVKGUAFhhg")
            ActionChains(driver).move_to_element(elemIrEmAndamento).perform()
            WebDriverWait(driver, 7)
            elemFinalEmAnda = elemIrEmAndamento.find_element(by=By.CLASS_NAME, value="bottomDropZone")
            WebDriverWait(driver, 7)
            ActionChains(driver).move_to_element(elemFinalEmAnda).perform()
            contadorA += 1

        # ----------Faz a rolagem até o fim de Bloqueado-----------------------------------------------------------------
        WebDriverWait(driver, 5)
        elemIrBloqueado = driver.find_element(by=By.ID, value="qeGI_YEH3EeY6ayIoMXKz2UAHQ_P")
        ActionChains(driver).move_to_element(elemIrBloqueado).perform()
        WebDriverWait(driver, 7)
        elemFinalBloqueado = elemIrBloqueado.find_element(by=By.CLASS_NAME, value="bottomDropZone")
        WebDriverWait(driver, 7)
        ActionChains(driver).move_to_element(elemFinalBloqueado).perform()

        # exibir o campo Tarefas feitas
        elemExpandirFeitos = driver.find_element(by=By.CLASS_NAME, value="sectionToggleButton")
        elemExpandirFeitos.click()
        time.sleep(10)

        # iterar entre os itens da coluna Feitos Planner
        contador = 0
        while contador <= 15:
            driver.implicitly_wait(16)
            FinalFeitos = driver.find_element(by=By.XPATH,
                                              value='//*[@id="6resMGX0u0aVUTWevh62-WUAOHmk"]/div/div[2]/div[2]/div[2]')
            ActionChains(driver).move_to_element(FinalFeitos).perform()
            contador += 1
        # aguardar
        time.sleep(15)
        WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.ID, index)))
        element = driver.find_element(By.ID, value=index)
        driver.implicitly_wait(15)
        ActionChains(driver).move_to_element(element).perform()
        driver.implicitly_wait(15)
        webdriver.ActionChains(driver).double_click(element).perform()

        # aguarda o card ser carregado
        time.sleep(10)

        # carrega a div que contem todos os comentários
        lista_comentarios = driver.find_elements(by=By.CLASS_NAME, value="commentCard")
        driver.implicitly_wait(7)

        # acrescenta o numero do card
        contadorCard += 1

        # vai rodar todos os comentarios na lista de comentarios
        for comentario in lista_comentarios:
            # pega o usuário do comentário
            usuario = comentario.find_element(by=By.CLASS_NAME, value="commenterName").text

            # pega o data e hora do comentário
            data_hora = comentario.find_element(by=By.CLASS_NAME, value="timestamp").text

            # pega o texto do do comentário
            texto_comentario = comentario.find_element(by=By.CLASS_NAME, value="content").text

            # acrescenta a quantidade de linhas
            contadorLinha += 1

            # acrescenta a quantidade de comentarios em cada card
            contadorComCard += 1

            # tira partes desnescessarias da data
            data_hora = data_hora.replace(" de ", "-")
            data_hora = data_hora.replace(" às ", " ")

            # CONVERTER TODOS OS MESES PARA NUMEROS
            data_hora = data_hora.replace("janeiro", "1")
            data_hora = data_hora.replace("fevereiro", "2")
            data_hora = data_hora.replace("março", "3")
            data_hora = data_hora.replace("abril", "4")
            data_hora = data_hora.replace("maio", "5")
            data_hora = data_hora.replace("junho", "6")
            data_hora = data_hora.replace("julho", "7")
            data_hora = data_hora.replace("agosto", "8")
            data_hora = data_hora.replace("setembro", "9")
            data_hora = data_hora.replace("outubro", "10")
            data_hora = data_hora.replace("novembro", "11")
            data_hora = data_hora.replace("dezembro", "12")

            # formata a data para ser possivel realizar o calculo
            data_hora = pd.to_datetime(data_hora, format='%d-%m-%Y %H:%M', utc=True).date()

            # apresenta na tela (usuario, data_hora, texto_comentario, contadorLinha) apenas para acompanhamento
            print(usuario, data_hora, texto_comentario, contadorLinha, sep=" --- ")

            # insere os dados na planilha
            excelComentarios['Identificação da tarefa'][contadorLinha] = index
            excelComentarios['Usuario'][contadorLinha] = usuario
            excelComentarios['Data Hora'][contadorLinha] = data_hora
            excelComentarios['Comentario'][contadorLinha] = texto_comentario
            excelComentarios['Comentarios por card'][contadorLinha] = contadorComCard
            excelComentarios['Card'][contadorLinha] = contadorCard
            excelComentarios['Contador comentarios'][contadorLinha] = contadorLinha

        # clicar para fechar o card exibido atualmente no modal
        driver.implicitly_wait(5)
        botaoFechar = driver.find_element(by=By.CSS_SELECTOR, value="button[aria-label='Fechar caixa de diálogo']")
        botaoFechar.click()

# coleta a quantidade de linhas na planilha
QtdLinhasTeste = contadorLinha - 1
while contador1 < QtdLinhasTeste:
    contador1 += 1
    contador2 += 1
    # coleta a data do comentario e a data do comentario anterior
    data1 = excelComentarios['Data Hora'][contador1]
    data2 = excelComentarios['Data Hora'][contador2]
    # calcula a distancia em dias de um comentario para o outro
    resultado = data1 - data2
    # marca como certo o ultimo comentario do ultimo card da planilha
    excelComentarios.loc[(excelComentarios['Contador comentarios'] == contadorLinha), "Teste"] = 'Certo'
    # marca como errado o card que esta fora do prazo
    excelComentarios.loc[(resultado.days > 3) & (excelComentarios['Contador comentarios'] == contador1), "Teste"] = 'Errado'
    # verificar se é o ultimo comentario do card
    card1 = excelComentarios['Comentarios por card'][contador1]
    card2 = excelComentarios['Comentarios por card'][contador2]
    excelComentarios.loc[(resultado.days > 3) & (excelComentarios['Contador comentarios'] == contador1) & (card2 <= card1), "Teste"] = 'Certo'
    # marca como certo o card que esta dentro do prazo
    excelComentarios.loc[(resultado.days <= 3) & (excelComentarios['Contador comentarios'] == contador1), "Teste"] = 'Certo'

# gera a planilha com a validação de comentarios
excelComentarios.to_excel("ComentariosPreenchidos.xlsx")

# ------------------- Juntar as informações das duas planilhas
excel = pd.read_excel("PlanilhaPreenchida.xlsx")
Completa = pd.read_excel("teste.xlsx")
excelComentado = pd.read_excel("ComentariosPreenchidos.xlsx")
comentariosErros = pd.read_excel("ComentariosPreenchidos.xlsx")

# encontrar quantos comentarios existe por card
excelComentado['Total Comentarios'] = excelComentado.groupby('Identificação da tarefa')['Identificação da tarefa'].transform('count')

# encontra o total de erros nos cards
comentariosErros = comentariosErros.loc[comentariosErros['Teste.1'] == 'Errado']
comentariosErros['Comentarios com erros'] = comentariosErros.groupby('Identificação da tarefa')['Identificação da tarefa'].transform('count')
comentariosErros = comentariosErros[['Identificação da tarefa', 'Comentarios com erros']]
comentariosErros = comentariosErros.drop_duplicates()
print(comentariosErros)

# seleciona a coluna Identificação da tarefa e Total Comentarios e retira duplicidades
excel1 = excelComentado[['Identificação da tarefa', 'Total Comentarios']]
excel1 = excel1.drop_duplicates()
excel1 = excel1[excel1['Identificação da tarefa'].notna()]
print(excel1)

# mescla as duas tabelas para adicionar o TOTAL DE COMENTARIOS POR CARD VALIDOS e Comentarios com ERRO
excel = pd.merge(excel, excel1, how='outer')
excel = pd.merge(excel, comentariosErros, how='outer')
print(excel)

# mescla a tabela com o total de comentarios a que contenha todos os cards selecionados

excel = pd.merge(excel, Completa, how = 'outer')

print(excel)

excel.loc[pd.isnull(excel['Comentarios com erros']) == True, "Comentarios com erros"] = 0

writer = pd.ExcelWriter('planilha_completa.xlsx')
# aplica VERMELHO na linha com erro
def aplicar_estilo(linha):

    if linha['erros'] > 0:
        return ['background-color: red'] * len(linha)
    else:
        return ['background-color: white'] * len(linha)
excel.style.apply(axis=1, func=aplicar_estilo).to_excel(writer, sheet_name='Planner', index=False)
excelComentado.to_excel(writer, sheet_name='Comentarios das Tarefas', index=False)


writer.save()
