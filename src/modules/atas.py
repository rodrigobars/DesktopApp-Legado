def atas():
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service as ChromeService
    from webdriver_manager.chrome import ChromeDriverManager
    from time import sleep
    from datetime import date
    import pyautogui as gui
    import win32com.client as win32
    import os
    #from modules.mytools import progress_bar

    # Limpando o gen.py antes de ocorrer erro
    os.system("powershell Remove-Item -path $env:LOCALAPPDATA\Temp\gen_py -recurse")
    os.system("cls")

    ################################
    # Consertando o documento Word #
    ################################

    def actual_month():
        month = {
            1: "Janeiro",
            2: "Fevereiro",
            3: "Março",
            4: "Abril",
            5: "Maio",
            6: "Junho",
            7: "Julho",
            8: "Agosto",
            9: "Setembro",
            10: "Outubro",
            11: "Novembro",
            12: "Dezembro"
        }
        return month.get(date.today().month, "Invalid month")


    def startWord():
        dou = input("\nResultado do dou (Ex: 01/01/2000): ")+','
        ata_path = input("\nInsira o caminho da Ata: ")
        term_path = input("\nInsira o caminho do Termo: ")
        os.system("cls")
        # Abrindo o programa Word e setando a visibilidade como verdadeira

        global word
        word = win32.gencache.EnsureDispatch('Word.Application')
        word.Visible = False

        # Abrindo o caminho do documento Word
        wordDoc = word.Documents.Open(rf'{ata_path}')
        gui.getWindowsWithTitle('Word')[0].minimize()

        # Consertando os títulos
        wordDoc.Content.GoTo(3, 1, 1).Select()
        find = word.Selection.Find
        find.Text = "ANEXO III DO EDITAL DO "
        find.Replacement.Text = ""
        find.Execute(Replace=1, Forward=True)
        find.Text = "MINUTA "
        find.Execute(Replace=1, Forward=True)

        # Inserindo a data de publicação
        # Lembrando que é necessário remover o highlight
        wordDoc.Content.GoTo(3, 1, 1).Select()
        find = word.Selection.Find

        find.Text = "...../...../20.....,"

        find.Replacement.Text = f"{dou}"
        find.Replacement.Highlight = False
        find.Execute(Replace=1, Forward=True)

        wordDoc.Content.GoTo(3, 1, 1).Select()

        return wordDoc, term_path

    #####################################
    # Abrindo o navegador para o início #
    #####################################

    # Implementar o Type Annotation


    def runStartWork(wordDoc):
        wordDoc.Tables(1).Delete()

        content = wordDoc.Content
        content.Find.Text = "proposta(s) são as que seguem:"
        content.Find.Execute(Forward=True)

        # Colocando o cursor no final do documento
        CurPos = content.Words(1).End
        # Pegando o número de parágrafos até a ocorrência da palavra do content que é um "Find"
        NumParagraphs = wordDoc.Range(0, CurPos).Paragraphs.Count

        # Iniciando o carregamento da página
        url = "http://comprasnet.gov.br/livre/pregao/ata0.asp"

        # Informando o Path para o arquivo "chromedriver.exe" e armazenando em "driver"
        op = webdriver.ChromeOptions()
        #op.add_argument("--headless")
        op.add_experimental_option("excludeSwitches", ["enable-logging"])
        driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=op)

        # Abrindo o google chrome no site informado
        driver.get(url)

        # Navegador em tela cheia
        driver.maximize_window()

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
        
        print(f"Empresas: {num_companys}")

        indexStartEnd = [1]
        companyInfo = {}
        start = 0
        end = 2
        startRealIndex = 1
        for company in range(1, num_companys+1):
            driver.execute_script(
                '''window.getElementByXpath = function (path){
                    return document.evaluate(path, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;}
                '''
            )

            sleep(1)

            # Preciso checar quantos itens a empresa tem...
            indexStartEnd = driver.execute_script(f"""
                // Selecionando a empresa na tabela
                range = document.createRange();
                sel = window.getSelection();
                sel.removeAllRanges();
                var root_node = document.getElementsByClassName("td")[0];
                range.setStart(root_node.getElementsByClassName('tex3')[{start}], 1);
                range.setEnd(root_node.getElementsByClassName('tex3b')[{end}], 1);
                sel.addRange(range);

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

                //removi do array cnpj, name
                return Array(indexStart, indexEnd, cnpj, name) 
            """)
            cnpj = indexStartEnd[2].rstrip()
            name = indexStartEnd[3][3:].rstrip()
            companyInfo.update({cnpj: name})
            ###########################################################################################
            # Preciso usar essa informação pra formatar a primeira linha do word e elaborar os termos #
            ###########################################################################################

            gui.hotkey('ctrl', 'c')

            wordDoc.Paragraphs.Add(wordDoc.Paragraphs(NumParagraphs+1).Range)
            wordDoc.Paragraphs.Add(wordDoc.Paragraphs(NumParagraphs+1).Range)
            wordDoc.Paragraphs.Add(wordDoc.Paragraphs(NumParagraphs+1).Range)
            wordDoc.Paragraphs.Add(wordDoc.Paragraphs(NumParagraphs+1).Range)
            wordDoc.Content.Paragraphs(NumParagraphs+2).Range.Paste()
            wordDoc.Paragraphs.Add(wordDoc.Paragraphs(NumParagraphs+1).Range)
            wordDoc.Paragraphs.Add(wordDoc.Paragraphs(NumParagraphs+1).Range)
            wordDoc.Paragraphs.Add(wordDoc.Paragraphs(NumParagraphs+1).Range)
            wordDoc.Paragraphs.Add(wordDoc.Paragraphs(NumParagraphs+1).Range)

            companyItens = int(((indexStartEnd[1]-indexStartEnd[0])-3)/2)
            start += int((companyItens*2)+1)
            end += 3
            startRealIndex += int((companyItens*2)+4)

        # Fechando o navegador
        driver.quit()

        # Varredura de colunas
        for i in range(1, num_companys+1):
            actual_table = wordDoc.Tables(i)
            endOfTable = False
            line = 3

            # Removendo e Mesclando o cabeçalho
            actual_table.Cell(2, 5).Range.Delete()
            actual_table.Cell(2, 4).Merge(wordDoc.Tables(i).Cell(2, 5))

            # Varredura de linhas da tabela atual
            while endOfTable == False:
                try:
                    actual_table.Cell(line, 5).Range.Delete()
                    actual_table.Cell(line, 4).Merge(
                        wordDoc.Tables(i).Cell(line, 5))
                    actual_table.Cell(line+1, 0).Range.Select()
                    word.Selection.Paragraphs.Alignment = 0
                    #gui.hotkey('ctrl', 'l')
                    line += 2
                except:
                    endOfTable = True

            # Colocando borda na tabela atual
            for j in range(1, 5):
                border = wordDoc.Tables(i).Borders.Item(-j)
                border.LineStyle = 21

            # Formatando as informações de cada empresa
            companyInfoDisplay = f"RAZÃO SOCIAL: {list(companyInfo.values())[abs(num_companys-i)]} \nCNPJ: {list(companyInfo.keys())[abs(num_companys-i)]} \nENDEREÇO: , CEP: \nTELEFONE: \nE-MAIL: \nDADOS BANCARIOS: \nREPRESENTANTE: , CPF: "

            actual_table = wordDoc.Tables(i)
            actual_table.Cell(1, 1).Range.Delete()
            actual_table.Cell(1, 1).Range.InsertAfter(companyInfoDisplay)


            # Mudando nome da fonte, tamanho e negrito
            textSelected = actual_table.Cell(1, 1).Range.Font
            textSelected.Name = "Calibri"
            textSelected.Size = "10,5"
            textSelected.Bold = True

        # Salva as alterações pendentes no mesmo arquivo automaticamente, sem solicitar ao usuário se -1
        wordDoc.Close(-1, 0)

        return companyInfo


    def startTerms(companyInfo, term_path):
        word = win32.gencache.EnsureDispatch('Word.Application')
        word.Visible = False

        print(f"\nElaborando os Termos de Responsabilidade...\n")

        for idx, (key, value) in enumerate(companyInfo.items()):
            sleep(1)
            # Lembrando que a procura parte de onde o cursos está para frente, por isso é necessário colocá-lo no início
            # Removendo parte do título
            if key == list(companyInfo.keys())[0]:
                wordDoc = word.Documents.Open(rf'{term_path}')  
                sleep(1)
                wordDoc.Content.GoTo(3, 1, 1).Select()
                find = word.Selection.Find
                find.Text = "ANEXO IV DO EDITAL DO "
                find.Replacement.Text = " "
                find.Execute(Replace=1, Forward=True)
                sleep(1)
                # Preenchendo com o nome da empresa
                wordDoc.Content.GoTo(3, 1, 1).Select()
                find.Text = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
                find.Replacement.Text = f"{value}"
                find.Replacement.Highlight = False
                find.Execute(Replace=1, Forward=True)
                sleep(1)
                # Preenchendo com o cnpj da empresa
                wordDoc.Content.GoTo(3, 1, 1).Select()
                find.Text = "XXXXXXXXXXXXXXXXXXXXXXX"
                find.Replacement.Text = f"{key}"
                find.Replacement.Highlight = False
                find.Execute(Replace=1, Forward=True)

                # Preenchendo com dia de hoje
                wordDoc.Content.GoTo(3, 1, 1).Select()
                find.Text = "____"
                find.Replacement.Text = f"{date.today().day}"
                find.Execute(Replace=1, Forward=True)

                # Preenchendo com o mês de hoje
                wordDoc.Content.GoTo(3, 1, 1).Select()
                find.Text = "___________"
                find.Replacement.Text = f"{actual_month()}"
                find.Execute(Replace=1, Forward=True)
                sleep(1)
                wordDoc.SaveAs2(
                    rf"{os.path.dirname(term_path)}\\{' '.join(value.split()[0:2]).upper()}", 16)
            else:
                # Preenchendo com o nome da empresa
                wordDoc.Content.GoTo(3, 1, 1).Select()
                find.Text = valueAux
                find.Replacement.Text = f"{value}"
                find.Replacement.Highlight = False
                find.Execute(Replace=1, Forward=True)
                sleep(1)
                # Preenchendo com o cnpj da empresa
                wordDoc.Content.GoTo(3, 1, 1).Select()
                find.Text = keyAux
                find.Replacement.Text = f"{key}"
                find.Replacement.Highlight = False
                find.Execute(Replace=1, Forward=True)
                sleep(1)
                wordDoc.SaveAs2(
                    rf"{os.path.dirname(term_path)}\\{' '.join(value.split()[0:2]).upper()}", 16)
                sleep(1)
            keyAux = key
            valueAux = value

        with open(rf"{os.path.dirname(term_path)}\\mails.txt", 'w') as mails:
            for companyName in companyInfo.values():
                mails.write(f"{' '.join(companyName.split()[0:2]).upper()};\n")

        word.Quit()

    Pregao = str(input("Digite o número do pregão (Ex: 12000): "))
    wordDocTerm = startWord()
    companyInfo = runStartWork(
        wordDoc=wordDocTerm[0])
    startTerms(companyInfo=companyInfo, term_path=wordDocTerm[1])
    print("Finalizado \U0001F7E2")