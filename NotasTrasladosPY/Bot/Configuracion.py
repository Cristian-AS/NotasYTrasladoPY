import json
import os

# Obtén la ruta del directorio actual donde se encuentra el archivo
current_dir = os.path.dirname(os.path.abspath(__file__))

# Mueve un nivel hacia arriba
workfolder_path = os.path.dirname(current_dir)

def create_folder(path):
    """Crea una carpeta si no existe."""
    if not os.path.exists(path):
        os.mkdir(path)
        print(f"Creada la carpeta en: {path}")
        return True
    else:
        print(f"La carpeta ya existe en: {path}")
        return False

def read_file(file_path):
    """Lee un archivo y retorna su contenido."""
    with open(file_path, encoding='utf-8') as file:
        content = file.read()
    return content

def main():
    config_file = os.path.join(workfolder_path, "config.json")

    # Leer el archivo de configuración JSON
    with open(config_file, mode='r', encoding='utf-8') as fp:
        data = json.load(fp)

    # Crear las carpetas necesarias
    folders_to_create = ["reports", "templates", "database", "email"]
    all_folders_created = all(create_folder(os.path.join(workfolder_path, folder_name)) for folder_name in folders_to_create)

    # Leer el mensaje de correo electrónico
    email_message_path = os.path.join(workfolder_path, "email", data['email']['email_message'])
    email_message = read_file(email_message_path)

    if all_folders_created:
        print("Se han creado todas las carpetas necesarias.")
    else:
        print("No se han creado todas las carpetas necesarias.")

    if email_message:
        print("Se ha leído correctamente el mensaje de correo electrónico.")
    else:
        print("No se ha podido leer el mensaje de correo electrónico.")

if __name__ == "__main__":
    main()