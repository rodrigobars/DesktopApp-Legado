from modules.text_result import text_result
from modules.scrap_to_excel import scrap_to_excel
from modules.atas import atas
from modules.irp import irp
from modules.mytools import applyColor

from os import system as cmd    

def main(module):

    #XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    #   Checagem de dependências...   X
    #XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

    import importlib.util
    from os import system as cmd
    from time import sleep

    package_names = ['pip', 'pandas', 'selenium', 'webdriver_manager', 'openpyxl', 'win32com', 'jinja2', 'pyautogui', 'PySimpleGUI']

    print('\nChecando dependências...\n')

    print("=======================")
    for package in package_names.copy():
        spec = importlib.util.find_spec(package)
        if spec is not None:
            package_names.remove(package)
            print(f'{package:<20} OK')
        else:
            print(f'{package:<20} X')
    print("=======================")

    if 'pip' in package_names:
        cmd('python -m ensurepip')

    if package_names:
        print('\nInstalando dependências...\n')
        for i in range(1,4):
            print(abs(4-i))
            sleep(1)
        cmd(f"pip install {' '.join(package_names)}")

    print('\nIniciando...\n')
    for i in range(1,4):
        print(abs(4-i))
        sleep(1)

    #XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

    match module:

        case '1':
            atas()

        case '2':
            scrap_to_excel()

        case '3':
            text_result()

        case '4':
            irp()

if __name__ == '__main__':
    cmd('cls')

    module = 0
    while True:
        module = input(
        applyColor('''
        _________________________________________
        |                                         |
        |   1 : Ata e Termos de Responsabilidade  |
        |   2 : Planilha de empresas vencedoras   |
        |   3 : Resultado de julgamento           |
        |   4 : IRP                               |
        |_________________________________________|
        
        Código do módulo: ''', text_color=2)
        )

        if module in ["1","2","3","4"]:
            break
        else:
            cmd("cls")
            print(applyColor(">>> Informe um número válido <<<", text_color=1))

    main(module)