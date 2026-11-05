"""Carga normas, clientes e inspectores desde los JSON de una carpeta data/.

Usa la clave secreta guardada en el Administrador de credenciales de Windows. Sin
--aplicar solo arma los datos e informa. Se niega a cargar sobre tablas con datos,
para no duplicar registros; si algo falla a medias, vacía lo que alcanzó a cargar.

    python base_de_datos/cargar_catalogos.py data
    python base_de_datos/cargar_catalogos.py data --aplicar
"""
import json
import os
import sys
from datetime import datetime

import keyring
from supabase import create_client

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from supabase_cliente import cargar_configuracion

SERVICIO_SECRETOS = 'VyC-Supabase'
# En orden de carga; se vacían en orden inverso.
TABLAS = ('normas', 'clientes', 'cliente_direcciones', 'inspectores', 'inspector_normas')


def _leer(carpeta, archivo):
    with open(os.path.join(carpeta, archivo), encoding='utf-8') as f:
        return json.load(f)


def _texto(valor):
    """Quita espacios y convierte lo vacío en None."""
    return str(valor).strip() or None if valor is not None else None


def _fecha(valor, formato, avisos, origen):
    valor = _texto(valor)
    if not valor:
        return None
    try:
        return datetime.strptime(valor, formato).date().isoformat()
    except ValueError:
        avisos.append(f"Fecha inválida en {origen}: {valor}")
        return None


def filas_normas(normas, inspectores):
    """Devuelve las normas y cuáles se agregaron por no estar en Normas.json."""
    filas = [
        {'nom': _texto(n['NOM']), 'nombre': _texto(n.get('NOMBRE')), 'capitulo': _texto(n.get('CAPITULO'))}
        for n in normas
    ]
    conocidas = {f['nom'] for f in filas}
    faltantes = sorted({_texto(n) for i in inspectores for n in i.get('Normas acreditadas') or []} - conocidas)
    return filas + [{'nom': nom, 'nombre': None, 'capitulo': None} for nom in faltantes], faltantes


def fila_cliente(cliente, avisos):
    nombre = _texto(cliente['CLIENTE'])
    return {
        'rfc': _texto(cliente['RFC']),
        'nombre': nombre,
        'numero_contrato': _texto(cliente.get('NÚMERO_DE_CONTRATO')),
        'fecha_contrato': _fecha(cliente.get('FECHA_DE_CONTRATO'), '%d/%m/%Y', avisos, nombre),
        'actividad': _texto(cliente.get('ACTIVIDAD')),
        'curp': _texto(cliente.get('CURP')),
    }


def filas_direcciones(cliente):
    return [
        {
            'calle_numero': _texto(d.get('CALLE Y NO')),
            'colonia': _texto(d.get('COLONIA O POBLACION') or d.get('COLONIA')),
            'municipio': _texto(d.get('MUNICIPIO O ALCADIA')),
            'estado': _texto(d.get('CIUDAD O ESTADO')),
            'cp': _texto(d.get('CP')),
            'servicio': _texto(d.get('SERVICIO')),
        }
        for d in cliente.get('DIRECCIONES') or []
    ]


def fila_inspector(inspector, avisos):
    nombre = _texto(inspector['NOMBRE DE INSPECTOR'])
    # Algunas fechas traen hora; solo importa el día.
    fecha = _texto(inspector.get('Fecha de acreditación'))
    return {
        'nombre': nombre,
        'correo': _texto(inspector.get('CORREO')),
        'firma': _texto(inspector['FIRMA']),
        'puesto': _texto(inspector.get('Puesto')),
        'vigencia': _texto(inspector.get('VIGENCIA')),
        'fecha_acreditacion': _fecha(fecha[:10] if fecha else None, '%Y-%m-%d', avisos, nombre),
        'referencia': _texto(inspector.get('Referencia')),
        'imagen': _texto(inspector.get('IMAGEN')),
    }


def _insertar(cliente, tabla, filas):
    return cliente.table(tabla).insert(filas).execute().data


def _cargar(cliente, normas, clientes, inspectores, avisos):
    ids_normas = {f['nom']: f['id'] for f in _insertar(cliente, 'normas', normas)}
    # De uno en uno para conocer el id que recibe cada cliente e inspector.
    for origen in clientes:
        cliente_id = _insertar(cliente, 'clientes', fila_cliente(origen, avisos))[0]['id']
        direcciones = [{**d, 'cliente_id': cliente_id} for d in filas_direcciones(origen)]
        if direcciones:
            _insertar(cliente, 'cliente_direcciones', direcciones)
    for origen in inspectores:
        inspector_id = _insertar(cliente, 'inspectores', fila_inspector(origen, avisos))[0]['id']
        acreditadas = {_texto(n) for n in origen.get('Normas acreditadas') or []}
        if acreditadas:
            _insertar(cliente, 'inspector_normas', [
                {'inspector_id': inspector_id, 'norma_id': ids_normas[nom]} for nom in acreditadas
            ])


def _vaciar(cliente):
    for tabla in reversed(TABLAS):
        filtro = 'inspector_id' if tabla == 'inspector_normas' else 'id'
        cliente.table(tabla).delete().gte(filtro, 0).execute()


def main():
    argumentos = [a for a in sys.argv[1:] if a != '--aplicar']
    if len(argumentos) != 1:
        sys.exit(__doc__)
    carpeta, aplicar = argumentos[0], '--aplicar' in sys.argv

    normas_json = _leer(carpeta, 'Normas.json')
    clientes = _leer(carpeta, 'Clientes.json')
    inspectores = _leer(carpeta, 'Firmas.json')

    avisos = []
    normas, faltantes = filas_normas(normas_json, inspectores)
    print(f"{'normas':22}{len(normas):6}")
    print(f"{'clientes':22}{len(clientes):6}")
    print(f"{'cliente_direcciones':22}{sum(len(filas_direcciones(c)) for c in clientes):6}")
    print(f"{'inspectores':22}{len(inspectores):6}")
    print(f"{'inspector_normas':22}{sum(len({_texto(n) for n in i.get('Normas acreditadas') or []}) for i in inspectores):6}")
    print('Normas agregadas sin descripción:', ', '.join(faltantes) or 'ninguna')
    for origen in clientes:
        fila_cliente(origen, avisos)
    for origen in inspectores:
        fila_inspector(origen, avisos)
    print('Avisos:', len(avisos))
    for aviso in avisos:
        print('  ', aviso)
    if not aplicar:
        print('SIMULACIÓN: no se guardó nada')
        return

    secreto = keyring.get_password(SERVICIO_SECRETOS, 'secret_key')
    if not secreto:
        sys.exit("Falta secret_key en el Administrador de credenciales de Windows.")
    cliente = create_client(cargar_configuracion()[0], secreto)

    for tabla in TABLAS:
        total = cliente.table(tabla).select('*', count='exact').limit(1).execute().count
        if total:
            sys.exit(f"public.{tabla} ya tiene {total} filas. No se cargó nada.")
    try:
        _cargar(cliente, normas, clientes, inspectores, [])
    except Exception:
        _vaciar(cliente)
        raise
    print('APLICADO')


if __name__ == '__main__':
    main()
