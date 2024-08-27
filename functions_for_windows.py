import os

def get_user_home_folder():
    if os.name == 'posix':  # Verifica se é um sistema tipo Unix (Linux, macOS, etc.)
        return os.getenv('HOME')
    elif os.name == 'nt':  # Verifica se é um sistema Windows
        return os.getenv('USERPROFILE')
    else:
        raise OSError('Sistema operacional não suportado')


def create_folder(folder_path, folder_name):
    folder_full_path = os.path.join(folder_path, folder_name)
    os.makedirs(folder_full_path, exist_ok=True)

print(get_user_home_folder())


def error():
    ...