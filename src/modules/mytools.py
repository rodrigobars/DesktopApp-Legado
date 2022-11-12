import colorama

def progress_bar(progress, total, color=colorama.Fore.MAGENTA):
    percent = 100 * ((progress) / float(total-1))
    bar = '█' * (int(percent/2)) + ' ' * (50 - (int(percent/2)))
    if progress == 1:
        print('\n')
    elif progress+1 != total:
        print(color + f"\r >>> |{bar}| {percent:.2f}%", end='\r')
    else:
        print(colorama.Fore.LIGHTGREEN_EX + f"\r >>> |{bar}| {percent:.2f}%", end='\r')
        print(colorama.Fore.RESET, '\n\n')

def applyColor(text, text_format = 1, text_color = 0, background_color = 0):
    """
    text format code:
    - 1 - None
    - 2 - Bold
    - 3 - Italic
    - 4 - Underline
    - 7 - Negative

    color code:
    - 0    - white
    - 1    - red
    - 2    - green
    - 3    - yellow
    - 4    - blue
    - 5    - purple
    - 6    - light blue
    - 7    - gray
    """
    return f"\033[{str(text_format)};{'3'+str(text_color)};{'4'+str(background_color)}m{text}\033[m"

def loading_dots():
    dot = "-"
    idx2 = 1
    side = '>'
    while True:
        yield dot
        if side == '>':
            idx2 += 1
            if idx2==5:    
                dot += "-"
                idx2 = 1
            if dot == "----":
                dot = 'ˍ'
                side = '<'
        if side == '<':
            idx2 += 1
            if idx2==5:    
                dot += "ˍ"
                idx2 = 1
            if dot == "ˍˍˍˍ":
                dot = '-'
                side = '>'

def loading_circle():
    simbol = {1:'◜', 2:'◝', 3:'◞', 4:'◟'}
    idx = 1
    while True:
        yield simbol[idx]
        idx += 1
        if idx == 5: idx = 1

def check_packages():
    import importlib.util
    from time import sleep
    import os

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
        os.system('python -m ensurepip')

    if package_names:
        print('\nInstalando dependências...\n')
        for i in range(1,4):
            print(abs(4-i))
            sleep(1)
        os.system(f"pip install {' '.join(package_names)}")