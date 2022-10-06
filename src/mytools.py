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