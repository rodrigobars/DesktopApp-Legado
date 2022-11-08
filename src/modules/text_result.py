def text_result():
    from os import system as cmd

    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service as ChromeService
    from webdriver_manager.chrome import ChromeDriverManager

    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    from modules.mytools import progress_bar, applyColor, loading

    # Inserindo o número do pregão
    cmd('cls')
    Pregao = input('\n'+applyColor("Digite o número do pregão (Ex: 12020): ", text_color = 5))
    Path = input('\n'+applyColor("Insira o caminho onde deseja salvar o Resultado de Julgamento: ", text_color = 5))
    cmd('cls')
    
    # Iniciando o carregamento da página
    url = "http://comprasnet.gov.br/livre/pregao/ata0.asp"

    # Informando o Path para o arquivo "chromedriver.exe" e armazenando em "driver"
    op = webdriver.ChromeOptions()
    op.add_argument('headless')
    op.add_experimental_option("excludeSwitches", ["enable-logging"])
    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=op)

    # Abrindo o google chrome no site informado
    driver.get(url)

    # Configurando o tempo de espera em segundos até que o elemento html seja encontrado
    wdw = WebDriverWait(driver, 2000)

    CodUasg = wdw.until(
        EC.presence_of_element_located((By.XPATH, "//input[@id='co_uasg']"))
    )
    CodUasg.send_keys("150182")

    NumPreg = wdw.until(EC.presence_of_element_located(
        (By.XPATH, "//input[@id='numprp']")))
    NumPreg.send_keys(Pregao)

    driver.execute_script("ValidaForm();")

    try:
        driver.find_element(By.XPATH, "//a[contains(text(), '{}')]".format(Pregao)).click()
    except:
        #driver.find_element(By.XPATH, "//center[@class= 'mensagem'][text()[contains(.,'Nenhuma Ata Encontrada.')]]")
        print(applyColor("Pregão não disponível!\n", text_color=1))
        input(applyColor("   -Pressione 'Enter' para sair...\n", text_color = 5))
        return

    ######################################################################
    # HOMOLOGAÇÃO

    try:
        driver.find_element(By.XPATH, "//input[@name='termodehomologacao']").click()
    except:
        print(applyColor("Não houveram itens homologados até o momento!\n", text_color=1))
        input(applyColor("   -Pressione 'Enter' para sair...\n", text_color = 5))
        return
    
    itensHomolog = []

    hasNextPage = True

    print(applyColor('\nVerificando itens homologados...\n', text_color=5))

    loading = loading()

    while hasNextPage:
        try:
            homolog = driver.find_elements(By.XPATH, "//tbody[tr/td/table/tbody/tr/td[text()[contains(., 'Homologado')]]]/tr[1]/td[contains(text(), 'Item:')]")
        except:
            homolog = None

        if homolog != None and homolog.__len__() > 0:
            for item in range(homolog.__len__()):
                itensHomolog.append(int(homolog[item].text[6:]))
                print(applyColor(next(loading), text_color=5), end='\r')

        try:
            driver.find_element(By.ID, "proximas")
            driver.execute_script("javascript:PaginarItens('Proxima');")
        except:
            hasNextPage = False

    driver.execute_script("javascript:voltar();")

    Owners = wdw.until(
        EC.presence_of_element_located(
            (By.XPATH, "//input[@id='btnResultadoFornecr']")
        )
    )
    Owners.click()

    # Armazena a quantidade total de empresas vencedoras
    num_companys = driver.find_elements(By.XPATH, "//td[contains(text(), 'Item')]").__len__()

    box = {}

    indexStartEnd = [1]
    start = 0
    end = 2
    startRealIndex = 1
    print('\n  '+applyColor(f"\n Empresas: {num_companys} ", text_color = 7, background_color = 5)+'\n')
    print(applyColor('Iniciando a coleta dos itens por empresa...', text_color=5))

    for company in range(1, num_companys+1):
        driver.execute_script(
            '''window.getElementByXpath = function (path){
                return document.evaluate(path, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;}
            '''
        )

        indexStartEnd = driver.execute_script(f"""
            // Selecionando a empresa na tabela
            range = document.createRange();
            sel = window.getSelection();
            sel.removeAllRanges();
            var root_node = document.getElementsByClassName("td")[0];
            
            // ###########################################################################
            // ##    O comando abaixo apenas seleciona a tabela do "start" até o "end"  ##
            // ###########################################################################
            // ##                                                                       ##
            // ## range.setStart(root_node.getElementsByClassName('tex3')[{start}], 1); ##
            // ## range.setEnd(root_node.getElementsByClassName('tex3b')[{end}], 1);    ##
            // ## sel.addRange(range);                                                  ##
            // ##                                                                       ##
            // ###########################################################################

            // Retornando o índice do 'start'
            var childStart = root_node.getElementsByClassName('tex3')[{start}]
            var parent = childStart.parentNode
            var indexStart = Array.prototype.indexOf.call(parent.children, childStart)

            // Retornando o índice do 'end'
            var childEnd = root_node.getElementsByClassName('tex3b')[{end}]
            var parent = childEnd.parentNode
            var indexEnd = Array.prototype.indexOf.call(parent.children, childEnd)

            // Retornando o CNPJ
            var cnpj = window.getElementByXpath('/html/body/table[2]/tbody/tr[{startRealIndex}]/td/b').textContent

            // Retornando o NOME
            var name = window.getElementByXpath('/html/body/table[2]/tbody/tr[{startRealIndex}]/td/text()').textContent

            // var itens = []
            // for (var item = {startRealIndex+2}; item < {end}; item=item+2) itens.push(window.getElementByXpath('/html/body/table[2]/tbody/tr[item]/td/text()').textContent)

            //removi do array cnpj, name
            return Array(indexStart, indexEnd, cnpj, name) 
        """)

        cnpj = indexStartEnd[2].rstrip()
        name = indexStartEnd[3][3:].rstrip()
        companyItens = int(((indexStartEnd[1]-indexStartEnd[0])-3)/2)

        itens = []
        for item in range(1,companyItens+1):
            item_atual = int(wdw.until(EC.presence_of_element_located(
                    (By.XPATH, f'/html/body/table[2]/tbody/tr[{startRealIndex+item*2}]/td[1]'))).text)
            if item_atual in itensHomolog:
                itens.append(item_atual)
        
        if itens:
            empresaFormatada = f" {name}({cnpj}):{(str(itens)[1:-1]).replace(' ', '')}"

            if company!=num_companys:
                empresaFormatada+=';'
            else:
                empresaFormatada+='.'

            box.update({company:empresaFormatada})

        start += int((companyItens*2)+1)
        end += 3
        startRealIndex += int((companyItens*2)+4)

        progress_bar(company, num_companys+1)

    op = webdriver.ChromeOptions()
    # op.add_argument("--start-maximized")
    op.add_experimental_option("excludeSwitches", ["enable-logging"])
    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=op)
    driver.set_window_size(500, 500)

    # Abrindo o google chrome no site informado
    driver.get(url)

    # Configurando o tempo de espera em segundos até que o elemento html seja encontrado
    wdw = WebDriverWait(driver, 2000)

    CodUasg = wdw.until(
        EC.presence_of_element_located((By.XPATH, "//input[@id='co_uasg']"))
    )
    CodUasg.send_keys("150182")

    NumPreg = wdw.until(EC.presence_of_element_located(
        (By.XPATH, "//input[@id='numprp']")))
    NumPreg.send_keys(Pregao)

    driver.execute_script("ValidaForm();")

    PrClick = wdw.until(
        EC.presence_of_element_located(
            (By.XPATH, "//a[contains(text(), '{}')]".format(Pregao))
        )
    )
    PrClick.click()

    AtaPr = wdw.until(
        EC.presence_of_element_located(
            (By.NAME, f"150182-{Pregao}-1")
        )
    )
    AtaPr.click()

    print(applyColor('Iniciando a coleta dos itens cancelados...\n', text_color = 5))

    wdw.until(
        EC.presence_of_element_located(
            (By.XPATH, "//td[contains(text(), 'MINISTÉRIO DA EDUCAÇÃO')]")
        )
    )

    itensCancelados = driver.find_elements(By.XPATH, "//tbody[tr/td/table/tbody/tr/td[text()[contains(., 'Cancelado no julgamento')]]]/tr[1]/td[contains(text(), 'Item')]")
    if not itensCancelados:
        itensCancelados = None

    itensDesertos = driver.find_elements(By.XPATH, "//tbody[tr/td/table/tbody/tr/td[text()[contains(., 'Cancelado por inexistência de proposta')]]]/tr[1]/td[contains(text(), 'Item')]")
    if not itensDesertos:
        itensDesertos = None

    itensRecurso = driver.find_elements(By.XPATH, "//tbody[tr/td/table/tbody/tr/td[text()[contains(., 'intenção de recurso')]]]/tr[1]/td[contains(text(), 'Item')]")
    if not itensRecurso:
        itensRecurso = None

    if itensCancelados != None:
        itensCanc = []
        for item in range(itensCancelados.__len__()):
            itensCanc.append(int(itensCancelados[item].text[6:]))
        if itensCancelados.__len__() == 2:
            canceladoPlural = ' ITEM CANCELADO:'
        else:
            canceladoPlural = ' ITENS CANCELADOS:'
        itensCanc = (canceladoPlural+(str(itensCanc)[1:-1]).replace(' ', ''))+'.'
    else:
        itensCancelados = False

    if itensDesertos != None:
        itensDese = []
        for item in range(itensDesertos.__len__()):
            itensDese.append(int(itensDesertos[item].text[6:]))
        if itensDesertos.__len__() == 1:
            desertoPlural = ' ITEM DESERTO:'
        else:
            desertoPlural = ' ITENS DESERTOS:'
        itensDese = (desertoPlural+(str(itensDese)[1:-1]).replace(' ', ''))+'.'
    else:
        itensDesertos = False

    if itensRecurso != None:
        itensRecu = []
        for item in range(itensRecurso.__len__()-1):
            itensRecu.append(int(itensRecurso[item].text[6:]))
        if itensRecurso.__len__() == 1:
            desertoPlural = ' ITEM EM RECURSO:'
        else:
            desertoPlural = ' ITENS EM RECURSO:'
        itensRecu = (desertoPlural+(str(itensRecu)[1:-1]).replace(' ', ''))+'.'
    else:
        itensRecurso = False

    aux = num_companys
    if itensCancelados:
        aux+=1
        box.update({f'{aux}':f'{itensCanc}'})
    if itensDesertos:
        aux+=1
        box.update({f'{aux}':f'{itensDese}'})

    def Text(num_pregao, num_companys, box):
        if num_companys > 1:
            itemPlural, empresaPlural = 'os seguintes itens', 'às empresas'
        elif companyItens == 1:
            itemPlural, empresaPlural = 'o seguinte item', 'à empresa'
        else:
            itemPlural, empresaPlural = 'os seguintes itens', 'a empresa'

        content = f'A Pró-Reitoria de Administração da Universidade Federal Fluminense torna público o resultado do julgamento do Pregão {num_pregao[:-4]}/PROAD/{num_pregao[-4:]}, tendo sido adjudicado e homologado {itemPlural} em relação {empresaPlural}:'

        for value in box.values():
            content += value

        return content

    content = Text(Pregao, num_companys, box)

    qtdlinhas = content.__len__()/47

    auxStart = 0
    auxEnd = 48

    with open(rf"{Path}\\texto de resultado {Pregao[:-4]}.2022.txt", 'w') as texto_resultado:
        if type(qtdlinhas) == int:
            for linha in range(qtdlinhas):
                texto_resultado.write(f"{content[auxStart:auxEnd]}\n")
                auxStart = auxEnd
                auxEnd += 47
            texto_resultado.write('\nESTA PUBLICAÇÃO EQUIVALE À PUBLICAÇÃO DA ATA DE\n REGISTRO DE PREÇOS.')
        else:
            for linha in range(int(qtdlinhas)+1):
                texto_resultado.write(f"{content[auxStart:auxEnd]}\n")
                if linha==qtdlinhas:
                    texto_resultado.write(f"{content[auxStart:(auxEnd+qtdlinhas%47)]}")
                auxStart = auxEnd
                auxEnd += 47
            texto_resultado.write('\nESTA PUBLICAÇÃO EQUIVALE À PUBLICAÇÃO DA ATA DE\n REGISTRO DE PREÇOS.')

    print('-------------------------- Resumo dos itens do pregão --------------------------')
    print(applyColor(f">>>  ITENS HOMOLOGADOS:{str(itensHomolog)[1:-1].replace(' ', '')+'.'}", text_color = 2))
    if itensCancelados:
        print(applyColor(f">>> {itensCanc}", text_color=1))
    if itensDesertos:
        print(applyColor(f">>> {itensDese}", text_color=1))
    if itensRecurso:
        print(applyColor(f">>> {itensRecu}", text_color=3))
    print('--------------------------------------------------------------------------------')

    driver.quit()

    input(applyColor("\nConcluído... \n\n    -Pressione 'Enter' para sair...\n", text_color = 5))