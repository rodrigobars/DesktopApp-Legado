from text_result import text_result
from scrap_to_excel import scrap_to_excel
from atas import atas
from mytools import applyColor
from irp import irp

from os import system as cmd

def main(module):

    match module:

        case '1':
            atas()

        case '2':
            scrap_to_excel()

        case '3':
            text_result()

        case '4':
            irp()

        case _:
            print('Não feito ainda...')


if __name__ == '__main__':
    cmd('cls')
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
    main(module)