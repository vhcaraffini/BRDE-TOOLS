import os

def get_mother_folder():
    caminho = os.path.dirname(os.path.realpath(__file__))
    return caminho


def create_folder(folder_path, folder_name):
    folder_full_path = os.path.join(folder_path, folder_name)
    os.makedirs(folder_full_path, exist_ok=True)

print(get_mother_folder())


def error():
    ...