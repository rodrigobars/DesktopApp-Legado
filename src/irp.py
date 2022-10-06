def irp():

    #XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    #   Checagem de dependências...   X
    #XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

    import importlib.util
    from os import system as cmd
    from time import sleep

    package_names = ['pip', 'pandas', 'selenium', 'webdriver_manager', 'PySimpleGUI', 'openpyxl']

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

    import pandas as pd
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.support.select import Select
    from selenium.webdriver.chrome.service import Service as ChromeService
    from webdriver_manager.chrome import ChromeDriverManager
    import PySimpleGUI as gui
    from math import ceil

    ############################################################################

    def CritérioDeValor(driver):
        wdw = WebDriverWait(driver, 20)
        CritériodeValor = wdw.until(
            EC.presence_of_element_located((By.ID, "idComboCriterioValor"))
        )
        Selected = Select(CritériodeValor)
        Selected.select_by_index(2)


    def PreçoUni(x, driver):
        wdw = WebDriverWait(driver, 20)
        ValorUnitário = wdw.until(
            EC.presence_of_element_located(
                (By.XPATH, '//*[@id="div_Item"]/table[2]/tbody/tr[2]/td[5]/input')
            )
        )
        ValorUnitário.send_keys(str(x * 1000))


    def ValorSigiloso(driver):
        wdw = WebDriverWait(driver, 20)
        ValorSigiloso = wdw.until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    "//input[@type='radio' and @name='item.valorCaraterSigiloso' and @value='2']",
                )
            )
        )
        ValorSigiloso.click()


    def Descrição(x, driver):
        # Verificando se o primeiro campo de descrição está usado...
        wdw = WebDriverWait(driver, 20)
        try:
            driver.find_element(By.XPATH, 
                "//textarea[contains(@class, 'fieldReadOnly')]")
            Segundo = driver.find_element(By.XPATH, 
                "//textarea[contains(@name, 'item.descricaoComplementar')]"
            )
            Segundo.send_keys(x)

        except:
            Primeiro = driver.find_element(By.XPATH, 
                "//textarea[contains(@name, 'item.descricaoDetalhada')]"
            )
            Primeiro.clear()
            Primeiro.send_keys(x)

    def localEquantidade(x, driver):
        driver.find_element(By.NAME, 'itemIRPMunicipioEntrega.municipio.nome').send_keys('Niterói/RJ')
        driver.execute_script('consultarMunicipio(this);')
        Janelas = driver.window_handles
        driver.switch_to.window(Janelas[1])
        driver.execute_script("window.opener.retornarConsultaMunicipio('58653;Niterói/RJ');window.close();")
        driver.switch_to.window(Janelas[0])
        qtd = driver.find_element(By.XPATH, 
            '//*[@id="div_Item"]/fieldset/table/tbody/tr[2]/td[3]/input'
        )
        qtd.send_keys(str(x))
        driver.execute_script("incluiMunicipio('divLocaisEntregaItemIRP');")

    def WindowManager(driver):
        sleep(1)
        Janelas = driver.window_handles
        try:
            driver.switch_to.window(Janelas[1])
            driver.find_element(By.XPATH, 
                "//title[contains(text(), 'Catálogo de materiais/serviços')]"
            )
            # Segunda janela é de CatMat (Correto)
        except:
            try:
                secundário = Janelas[2]
                driver.close()
                driver.switch_to.window(secundário)
                # Aqui 3 janelas foram abertas, como a segunda janela não é a de CatMat, fechamos ela e damos foco a segunda.
            except:
                if len(Janelas) == 2:
                    # Na arquitetura do código, quando chega aqui o atual handle é [0]
                    driver.switch_to.window(Janelas[1])
                    driver.close()
                    driver.switch_to.window(Janelas[0])
                    # Aqui na hora de salvar abriu um pop-up, portanto, devemos fecha-lo

    def Inclusão_Item(x, driver, início):

        for item in range(início, (len(x.iloc[:, 0]))):
            # Ele precisa ir para a página de onde pertence o item
            if (item+1)>20:
                aux = (item+1)/20
                if type(aux) == int:
                    driver.find_element(By.XPATH, f"//a[@title='Ir para página {aux}']").click()
                else:
                    aux = ceil(aux)
                    driver.find_element(By.XPATH, f"//a[@title='Ir para página {aux}']").click()
            sleep(2)

            # Auxílio no index, pois todas as páginas vão de 1 à 20, sem importar 
            aux2 = (item+1)%20 if item+1>20 else item+1
            if aux2 == 0:
                aux2 = 20
            sleep(1)
            driver.find_element(By.XPATH, f"//*[@id='itensIRP']/tbody/tr[{aux2}]/td[9]/a").click()

            # Seleciona o critério de valor
            CritérioDeValor(driver)

            # Seleciona o preço unitário
            PreçoUni(x.iloc[item, 4], driver)

            # Seleciona o valor sigiloso
            ValorSigiloso(driver)

            # Preenche a descrição
            Descrição(x.iloc[item, 0], driver)

            # Preenche a quantidade
            localEquantidade(x.iloc[item, 3], driver)

            # Por fim, é necessário salvar
            ####Elemento final da programação####
            sleep(0.5)
            driver.execute_script("alterarItemIRP();")

            sleep(1)
            WindowManager(driver)
            driver.execute_script('cancelarAlteracaoInclusaoItemIRP()')
            sleep(2)
            WindowManager(driver)
            sleep(1)

    def Programa(Login, Senha, NúmeroIRP, BaseDado, início, Path):
        # Chamada de tela
        url = "https://www.comprasnet.gov.br/seguro/loginPortal.asp"
        # Informando o Path para o arquivo "chromedriver.exe" e armazenando em "driver"
        op = webdriver.ChromeOptions()
        op.add_argument("--start-maximized")
        op.add_experimental_option("excludeSwitches", ["enable-logging"])
        driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=op)

        driver.get(url)

        wdw = WebDriverWait(driver, 20)

        # Selecionando perfil "Governo"
        driver.execute_script("mudaPerfilBotao(2);")

        # Digitando o CPF
        cpf = driver.find_element(By.ID, "txtLogin")
        cpf.send_keys(Login)
        sleep(0.1)

        # Digitando a SENHA
        senha = driver.find_element(By.ID, "txtSenha")
        senha.send_keys(Senha)
        sleep(0.1)

        # Pressionando o botão "ACESSAR"
        
        driver.execute_script("frmLoginGoverno_submit(); return false;")
        
        sleep(5)

        # Entrando na IRP
        dropdownMenuUser = wdw.until(
            EC.element_to_be_clickable((By.XPATH, "//button[@id='dropdownMenuUser']"))
        )
        dropdownMenuUser.click()


        IrpModule = wdw.until(
            EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'IRP')]"))
        )
        IrpModule.click()

        # ---  Nível de acesso = https://www.comprasnet.gov.br/seguro/indexgov.asp ---
        # Alterando entre Frames do site
        sleep(4)
        atual = driver.window_handles[1]
        driver.close()
        driver.switch_to.window(atual)

        irpmodule2 = wdw.until(EC.element_to_be_clickable(((By.ID, "mi_0_5"))))
        irpmodule2.click()

        IrpModule3 = wdw.until(EC.element_to_be_clickable((By.ID, "mi_0_7")))
        IrpModule3.click()

        NumIrp = driver.find_element(By.XPATH, 
            '//*[@id="corpo"]/form/table/tbody/tr[1]/td/table/tbody/tr[3]/td/input'
        )
        NumIrp.click()

        Parâmetro = driver.find_element(By.XPATH, 
            '//*[@id="numeroIrp"]/table/tbody/tr/td/input'
        )
        Parâmetro.send_keys(NúmeroIRP)

        driver.find_element(By.XPATH, '//*[@id="btnConsultar"]').click()

        driver.find_element(By.XPATH, 
            '//*[@id="listaIrps"]/tbody/tr/td[7]/a').click()

        driver.execute_script(
            "try{closeCalendar();}catch(e){};if(document.forms['manterIRPForm'].elements['irp.codigoIrp'].value != ''){ manterIRPForm.abaVisivelAtual.value=2;AlternarAbas('td_Item','div_Item');showHideBotaoDivulgar();showHideBarraBotaoIRP();}"
        )

        # Chamando a função de inclusão de itens:
        Inclusão_Item(BaseDado, driver, início)
        print("\n\n\nSUCESSO!\n\n\n")

    gui.theme_background_color("Beige")

    class Gui:
        def __init__(self):
            self.layout = [
                [
                    gui.Text(
                        "Perfil: Governo",
                        font="Vicasso",
                        text_color="black",
                        background_color="Beige",
                    )
                ],
                [
                    gui.Text(
                        "Login: ",
                        font="Vicasso",
                        size=(5, 1),
                        text_color="black",
                        background_color="Beige",
                    ),
                    gui.Input(size=(20, 1), key="Login"),
                ],
                [
                    gui.Text(
                        "Senha: ",
                        font="Vicasso",
                        size=(5, 1),
                        text_color="black",
                        background_color="Beige",
                    ),
                    gui.Input(
                        size=(20, 1), key="Senha", password_char="*", do_not_clear=False
                    ),
                ],
                [gui.T("", background_color="Beige")],
                [gui.Text("_" * 66, background_color="Orange")],
                [
                    gui.Text(
                        "Nº IRP:",
                        font="Vicasso",
                        size=(5, 1),
                        text_color="black",
                        background_color="Beige",
                    ),
                    gui.Input(size=(10, 1), key="IRP"),
                ],
                [gui.T("", background_color="Beige")],
                [gui.Text("_" * 66, background_color="Black")],
                [
                    gui.Text(
                        "Selecione a planilha:",
                        font="Vicasso",
                        text_color="black",
                        background_color="Beige",
                    )
                ],
                [
                    gui.Input(key="Path", size=(57, 1)),
                    gui.FileBrowse("Buscar", target=("Path")),
                ],
                [
                    gui.Text(
                        "Incluir a partir do item: ",
                        font="Vicasso",
                        text_color="black",
                        background_color="Beige",
                    ),
                    gui.Input(size=(3, 1), key="Início"),
                ],
                [gui.T("", background_color="Beige")],
                [gui.Button("Iniciar", size=(10, 1), key="_START_")],
            ]

            self.window = gui.Window("SIASGNet IRP").Layout(self.layout)

        def Iniciar(self):
            while True:
                Start = False
                while Start == False:
                    print("\nColetando...")
                    self.event, self.values = self.window.Read()
                    Login = self.values["Login"]
                    Senha = self.values["Senha"]
                    NúmeroIRP = self.values["IRP"]
                    Path = self.values["Path"]
                    BaseDado = pd.read_excel(Path)
                    BaseDado = BaseDado.set_index("ITEM")
                    BaseDado["FORNECIMENTO"] = BaseDado["FORNECIMENTO"].astype(str)
                    início = int(self.values["Início"]) - 1
                    if início < 0:
                        início = int(
                            input("\nItem inválido, digite novamente item inicial: \n")
                        )
                    elif início > len(BaseDado.iloc[:, 0]):
                        início = int(
                            input(
                                "\nItem inválido, total de itens: {}\n".format(
                                    len(BaseDado.iloc[:, 0])
                                )
                            )
                        )
                    elif Path == "":
                        print("Selecione a planilha a ser incluída!")
                        Start = False
                    else:
                        print("iniciando...")
                        Start = True

                if self.event is None:
                    print("parou")
                    break

                if self.event == "_START_" and Start == True:
                    Programa(Login, Senha, NúmeroIRP, BaseDado, início, Path)

    UserInterface = Gui()
    UserInterface.Iniciar()
    
if __name__ == '__main__':
    irp()