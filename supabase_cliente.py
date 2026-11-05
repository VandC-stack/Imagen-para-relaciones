"""Conexión de la aplicación con Supabase.

Solo usa la URL del proyecto y la clave publicable (`supabase_config.json`). Esa
clave viaja dentro del .exe y es pública por diseño: lo que protege los datos son
las reglas RLS de la base. Las claves secretas se guardan aparte, en el
Administrador de credenciales de Windows (ver herramientas/secretos_supabase.py),
y nunca deben llegar aquí.
"""
import base64
import json
import os
import sys
from functools import lru_cache

import httpx
from supabase import Client, create_client

ARCHIVO_CONFIG = 'supabase_config.json'
TIEMPO_ESPERA = 10


class ConfiguracionSupabaseError(Exception):
    """La configuración falta o trae una clave que no debe usarse en la app."""


def _ruta_config():
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, ARCHIVO_CONFIG)


def _es_clave_secreta(clave):
    """Detecta claves con acceso total, tanto del formato nuevo como del anterior."""
    if clave.startswith('sb_secret_'):
        return True
    partes = clave.split('.')
    if len(partes) != 3:
        return False
    try:
        relleno = '=' * (-len(partes[1]) % 4)
        datos = json.loads(base64.urlsafe_b64decode(partes[1] + relleno))
    except ValueError:
        return False
    return datos.get('role') == 'service_role'


def cargar_configuracion():
    """Devuelve (url, clave_publicable) desde el entorno o `supabase_config.json`."""
    url = os.environ.get('SUPABASE_URL', '')
    clave = os.environ.get('SUPABASE_PUBLISHABLE_KEY', '')
    if not (url and clave):
        try:
            with open(_ruta_config(), encoding='utf-8') as f:
                config = json.load(f)
        except (OSError, ValueError) as e:
            raise ConfiguracionSupabaseError(f"No se pudo leer {ARCHIVO_CONFIG}: {e}") from e
        url = config.get('url', '')
        clave = config.get('publishable_key', '')

    url, clave = url.strip().rstrip('/'), clave.strip()
    if not url.startswith('https://') or not clave:
        raise ConfiguracionSupabaseError(f"{ARCHIVO_CONFIG} debe traer 'url' (https) y 'publishable_key'.")
    if _es_clave_secreta(clave):
        raise ConfiguracionSupabaseError("Se configuró una clave secreta. La aplicación solo admite la clave publicable.")
    return url, clave


@lru_cache(maxsize=1)
def obtener_cliente() -> Client:
    """Devuelve el cliente compartido; se crea la primera vez que se necesita."""
    return create_client(*cargar_configuracion())


def probar_conexion():
    """Comprueba que el proyecto responde y acepta la clave. Devuelve (ok, mensaje)."""
    try:
        url, clave = cargar_configuracion()
        respuesta = httpx.get(f"{url}/auth/v1/health", headers={'apikey': clave}, timeout=TIEMPO_ESPERA)
    except ConfiguracionSupabaseError as e:
        return False, str(e)
    except httpx.HTTPError as e:
        return False, f"Sin conexión con Supabase: {e}"
    if respuesta.status_code == 200:
        return True, "Conexión correcta"
    if respuesta.status_code in (401, 403):
        return False, "Supabase rechazó la clave publicable."
    return False, f"Supabase respondió con el código {respuesta.status_code}."
