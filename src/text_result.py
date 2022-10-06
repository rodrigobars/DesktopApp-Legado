def text_result():
    #XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    #   Checagem de dependências...   X
    #XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

    import importlib.util
    from os import system as cmd
    from time import sleep

    package_names = ['pip', 'selenium', 'webdriver_manager']

    print('\nChecando dependências...\n')
    for package in package_names.copy():
        spec = importlib.util.find_spec(package)
        if spec is not None:
            package_names.remove(package)
            print(f'{package:<20} OK')
        else:
            print(f'{package:<20} X')

    if 'pip' in package_names:
        cmd('python -m ensurepip')

    if package_names:
        print('\nInstalando dependências...\n')
        for i in range(1,4):
            print(i)
            sleep(1)
        cmd(f"pip install {' '.join(package_names)}")

    #XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service as ChromeService
    from webdriver_manager.chrome import ChromeDriverManager

    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    from mytools import progress_bar, applyColor

    # Inserindo o número do pregão
    cmd('cls')
    Pregao = input('\n'+applyColor("Digite o número do pregão (Ex: 12020): ", text_color = 5))
    Path = input('\n'+applyColor("Insira o caminho onde deseja salvar o Resultado de Julgamento: ", text_color = 5))

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

    PrClick = wdw.until(
        EC.presence_of_element_located(
            (By.XPATH, "//a[contains(text(), '{}')]".format(Pregao))
        )
    )
    PrClick.click()

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
    companyInfo = {}
    start = 0
    end = 2
    startRealIndex = 1

    print('\n'+applyColor(f" Empresas: {num_companys} ", text_color = 0, background_color = 5)+'\n')
    print(applyColor('Iniciando a coleta dos itens por empresa...', text_color=5))

    for company in range(1, num_companys+1):
        driver.execute_script(
            '''window.getElementByXpath = function (path){
                return document.evaluate(path, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;}
            '''
        )

    #

        # Preciso checar quantos itens a empresa tem...
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
            itens.append(int(wdw.until(EC.presence_of_element_located(
                    (By.XPATH, f'/html/body/table[2]/tbody/tr[{startRealIndex+item*2}]/td[1]'))).text))
        
        empresaFormatada = f"{name}({cnpj}):{(str(itens)[1:-1]).replace(' ', '')}"

        if company!=num_companys:
            empresaFormatada+=';'
        else:
            empresaFormatada+='.'

        box.update({f"{company}":f" {empresaFormatada}"})

        #print(empresaFormatada)

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

    try:
        itensCancelados = driver.find_elements(By.XPATH, "//tbody[tr/td/table/tbody/tr/td[text()[contains(., 'Cancelado no julgamento')]]]/tr[1]/td")
    except:
        itensCancelados = None

    try:
        itensDesertos = driver.find_elements(By.XPATH, "//tbody[tr/td/table/tbody/tr/td[text()[contains(., 'Cancelado por inexistência de proposta')]]]/tr[1]/td")
    except:
        itensDesertos = None

    if itensCancelados != None and itensCancelados.__len__() > 1:
        itensCanc = []
        for item in range(itensCancelados.__len__()-1):
            itensCanc.append(int(itensCancelados[item].text[6:]))
        if itensCancelados.__len__() == 2:
            canceladoPlural = ' ITEM CANCELADO:'
        else:
            canceladoPlural = ' ITENS CANCELADOS:'
        itensCanc = (canceladoPlural+(str(itensCanc)[1:-1]).replace(' ', ''))+'.'
    else:
        itensCancelados = False

    if itensDesertos != None and itensDesertos.__len__() > 0:
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

    aux = num_companys
    if itensCancelados is not False:
        aux+=1
        box.update({f'{aux}':f'{itensCanc}'})
    if itensDesertos is not False:
        aux+=1
        box.update({f'{aux}':f'{itensDese}'})

    # for key, value in box.items():
    #         print(key, value)

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

    #text_path = r"C:\Users\Rodrigo.DESKTOP-ST2DND8\Desktop\Work\Texto de Resultado\2022"

    auxStart = 0
    auxEnd = 48

    with open(rf"{Path}\\texto de resultado {Pregao[:-4]}.2022.txt", 'w') as texto_resultado:
        if type(qtdlinhas) == int:
            for linha in range(qtdlinhas):
                # print(content[auxStart:auxEnd])
                texto_resultado.write(f"{content[auxStart:auxEnd]}\n")
                auxStart = auxEnd
                auxEnd += 47
            texto_resultado.write('\nESTA PUBLICAÇÃO EQUIVALE À PUBLICAÇÃO DA ATA DE\n REGISTRO DE PREÇOS.')
        else:
            for linha in range(int(qtdlinhas)+1):
                # print(content[auxStart:auxEnd])
                texto_resultado.write(f"{content[auxStart:auxEnd]}\n")
                if linha==qtdlinhas:
                    texto_resultado.write(f"{content[auxStart:(auxEnd+qtdlinhas%47)]}")
                    # print(content[auxStart:(auxEnd+qtdlinhas%47)])
                auxStart = auxEnd
                auxEnd += 47
            texto_resultado.write('\nESTA PUBLICAÇÃO EQUIVALE À PUBLICAÇÃO DA ATA DE\n REGISTRO DE PREÇOS.')
            # print('\nESTA PUBLICAÇÃO EQUIVALE À PUBLICAÇÃO DA ATA DE\n REGISTRO DE PREÇOS.')

    input(applyColor("Concluído... \n\n    -Pressione 'Enter' para sair...\n", text_color = 5))

if __name__ == '__main__':
    text_result()