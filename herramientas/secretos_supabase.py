"""Guarda las credenciales de Supabase en el Administrador de credenciales de Windows.

Quedan cifradas para el usuario de Windows de esta PC, fuera de git, OneDrive y
Dropbox. Es una herramienta de administración: la app no la importa y no va en el .exe.

    python herramientas/secretos_supabase.py lista
    python herramientas/secretos_supabase.py guardar db_password
    python herramientas/secretos_supabase.py copiar db_password
    python herramientas/secretos_supabase.py borrar db_password
"""
import getpass
import subprocess
import sys

import keyring

SERVICIO = 'VyC-Supabase'
SECRETOS = {
    'cuenta_email': 'Correo de la cuenta de Supabase',
    'cuenta_password': 'Contraseña de la cuenta de Supabase',
    'db_password': 'Contraseña de la base de datos',
    'secret_key': 'Clave secreta del proyecto (sb_secret_...)',
}


def obtener(nombre):
    """Devuelve un secreto para scripts de administración (migraciones, respaldos)."""
    valor = keyring.get_password(SERVICIO, nombre)
    if valor is None:
        raise KeyError(f"Falta '{nombre}'. Guárdalo con: python herramientas/secretos_supabase.py guardar {nombre}")
    return valor


def _lista():
    for nombre, descripcion in SECRETOS.items():
        estado = 'guardado' if keyring.get_password(SERVICIO, nombre) else 'falta'
        print(f"{nombre:16} {estado:9} {descripcion}")


def _guardar(nombre):
    valor = getpass.getpass(f"{SECRETOS[nombre]}: ")
    if not valor:
        sys.exit("No se guardó: el valor está vacío.")
    keyring.set_password(SERVICIO, nombre, valor)
    print(f"{nombre} guardado.")


def _copiar(nombre):
    """Pasa el valor al portapapeles para no mostrarlo en pantalla."""
    subprocess.run('clip', input=obtener(nombre).encode('utf-16-le'), check=True)
    print(f"{nombre} copiado al portapapeles.")


def _borrar(nombre):
    keyring.delete_password(SERVICIO, nombre)
    print(f"{nombre} borrado.")


def main():
    acciones = {'guardar': _guardar, 'copiar': _copiar, 'borrar': _borrar}
    argumentos = sys.argv[1:]
    if argumentos == ['lista']:
        _lista()
    elif len(argumentos) == 2 and argumentos[0] in acciones and argumentos[1] in SECRETOS:
        acciones[argumentos[0]](argumentos[1])
    else:
        sys.exit(__doc__)


if __name__ == '__main__':
    main()
