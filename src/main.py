from modules.text_result import text_result
from modules.scrap_to_excel import scrap_to_excel
from modules.atas import atas
from modules.irp import irp
from modules.mytools import applyColor, check_packages
from time import sleep
import os

def main():
    os.system('cls')

    menu_options = {'1':atas, '2':scrap_to_excel, '3':text_result, '4':irp}
    
    def set_menu_options():
        table_menu = '''
         _________________________________________
        |                                         |
        |   1 : Ata e Termos de Responsabilidade  |
        |   2 : Planilha de empresas vencedoras   |
        |   3 : Resultado de julgamento           |
        |   4 : IRP                               |
        |_________________________________________|

        Código do módulo: '''

        module = input(applyColor(table_menu, text_color=2))
        return module

    while (module := set_menu_options()) not in menu_options.keys():
       #os.system('cls')
       print(applyColor(f"{'':<8}>>> Informe um número válido <<<", text_color=1), end='\r')

    # Verifica se as dependências estão instaladas
    check_packages()

    print('\nIniciando...\n')
    for i in range(1,4):
        print(abs(4-i))
        sleep(1)

    menu_options[module]()

if __name__ == '__main__':
    main()