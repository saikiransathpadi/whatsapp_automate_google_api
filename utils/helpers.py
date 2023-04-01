from datetime import date

def get_today_date_formatted(format):
    today = date.today()
    return today.strftime(format)

def get_name_with_prefix(name):
    name_prefix = 'Dr.'
    if name.lower()[0:2] == 'dr':
        name_prefix=''
    return name_prefix + str(name).strip()

def print_colored(message, color="reset", style="format"):
    colors = {
        'black': '\u001b[30m',
        'red': '\u001b[31m',
        'green': '\u001b[32m',
        'yellow': '\u001b[33m',
        'blue': '\u001b[34m',
        'magenta': '\u001b[35m',
        'cyan': '\u001b[36m',
        'white': '\u001b[37m',
        'reset': '\u001b[0m'
    }
    if color not in colors:
        color = 'reset'
    try:
        message = getattr(message, style)()
    except AttributeError:
        print("Invalid style", style)
        pass
    print(f'{colors[color]}{message}{colors["reset"]}')
