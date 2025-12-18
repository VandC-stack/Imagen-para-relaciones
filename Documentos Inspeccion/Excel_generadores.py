# -- Este archivo genera el control de folios Anual y Reporte EMA -- #

import os
import json
import pandas as pd
from datetime import datetime

# REPORTE ANUAL #
def generar_control_folios_anual(historial_path, tabla_backups_dir, ruta_salida, year=None, start_date=None, end_date=None, export_cache=None):
    """Genera un Excel con el control de folios anual o por rango.
    - Puede leer del `export_cache` (archivo JSON generado por la app) si se provee.
    - start_date y end_date aceptan cadenas en formato dd/mm/YYYY o ISO (YYYY-MM-DD).
    - Los encabezados y columnas siguen la especificación del usuario.
    """
    # Preferir cache si existe
    rows = []
    try:
        if export_cache and os.path.exists(export_cache):
            with open(export_cache, 'r', encoding='utf-8') as f:
                ec = json.load(f)
            tabla = ec.get('ema', []) if isinstance(ec, dict) else []
            # tabla contiene filas ya enriquecidas (ver app._generar_datos_exportable)
            for r in tabla:
                rows.append(r)
        else:
            # Fallback: intentar leer tabla_de_relacion en backups_dir (último backup)
            tabla_file = os.path.join(tabla_backups_dir, 'tabla_de_relacion.json')
            if not os.path.exists(tabla_file):
                # buscar cualquier backup en la carpeta
                if os.path.exists(tabla_backups_dir):
                    files = sorted([os.path.join(tabla_backups_dir, f) for f in os.listdir(tabla_backups_dir) if f.endswith('.json')])
                    if files:
                        tabla_file = files[-1]
            tabla = []
            if os.path.exists(tabla_file):
                with open(tabla_file, 'r', encoding='utf-8') as f:
                    try:
                        tabla = json.load(f)
                    except Exception:
                        tabla = []
            # Construir filas desde tabla
            for r in tabla:
                try:
                    solicitud_full = r.get('ENCABEZADO', '') or r.get('SOLICITUD_ENCABEZADO', '') or r.get('SOLICITUD','')
                    sol_parts = str(solicitud_full).split()[-1] if solicitud_full else ''
                    cliente = r.get('EMPRESA','') or r.get('EMPRESA_VISITADA', r.get('CLIENTE',''))
                    numero_contrato = ''
                    rfc = ''
                    # intentar obtener datos de Clientes.json si está junto al historial
                    clientes_path = os.path.join(os.path.dirname(historial_path), 'Clientes.json')
                    if os.path.exists(clientes_path):
                        try:
                            with open(clientes_path, 'r', encoding='utf-8') as cf:
                                cl = json.load(cf)
                                if isinstance(cl, list):
                                    for c in cl:
                                        if (c.get('CLIENTE','').upper() == (cliente or '').upper()):
                                            numero_contrato = c.get('NÚMERO_DE_CONTRATO','')
                                            rfc = c.get('RFC','')
                                            break
                        except Exception:
                            pass

                    rows.append({
                        'NUMERO_SOLICITUD': sol_parts,
                        'CLIENTE': cliente,
                        'NUMERO_CONTRATO': numero_contrato,
                        'RFC': rfc,
                        'CURP': '',
                        'PRODUCTO_VERIFICADO': r.get('DESCRIPCION',''),
                        'MARCAS': r.get('MARCA',''),
                        'NOM': r.get('CLASIF UVA') or r.get('CLASIF_UVA') or r.get('NOM',''),
                        'TIPO_DOCUMENTO': r.get('TIPO DE DOCUMENTO') or r.get('TIPO_DE_DOCUMENTO',''),
                        'DOCUMENTO_EMITIDO': solicitud_full,
                        'FECHA_DOCUMENTO_EMITIDO': r.get('FECHA DE VERIFICACION') or r.get('FECHA_DE_VERIFICACION') or '',
                        'VERIFICADOR': r.get('VERIFICADOR') or r.get('INSPECTOR',''),
                        'PEDIMENTO_IMPORTACION': r.get('PEDIMENTO',''),
                        'FECHA_DESADUANAMIENTO': r.get('FECHA DE ENTRADA') or r.get('FECHA_ENTRADA',''),
                        'FECHA_VISITA': r.get('FECHA DE VERIFICACION') or r.get('FECHA_DE_VERIFICACION') or '',
                        'MODELOS': r.get('CODIGO',''),
                        'SOL_EMA': sol_parts,
                        'FOLIO_EMA': str(r.get('FOLIO','')).zfill(6) if str(r.get('FOLIO','')).strip() else '',
                        'INSP_EMA': r.get('VERIFICADOR') or r.get('INSPECTOR','')
                    })
                except Exception:
                    continue
    except Exception:
        rows = []

    # Filtrar por año / rango de fechas si se solicita
    def _parse_date(s):
        if not s:
            return None
        s = str(s).strip()
        # aceptar dd/mm/YYYY o ISO
        try:
            if '/' in s:
                dparts = s.split('/')
                return datetime(int(dparts[2]), int(dparts[1]), int(dparts[0]))
            elif '-' in s and len(s.split('-')[0]) == 4:
                return datetime.fromisoformat(s.split('T')[0])
        except Exception:
            try:
                return datetime.fromisoformat(s)
            except Exception:
                return None
        return None

    if start_date:
        sd = _parse_date(start_date)
    else:
        sd = None
    if end_date:
        ed = _parse_date(end_date)
    else:
        ed = None
    if year and not sd and not ed:
        try:
            sd = datetime(int(year), 1, 1)
            ed = datetime(int(year), 12, 31)
        except Exception:
            sd = ed = None

    if sd or ed:
        filtered = []
        for r in rows:
            d = r.get('FECHA_DOCUMENTO_EMITIDO') or r.get('FECHA_VISITA')
            pd = _parse_date(d)
            if pd:
                if sd and pd < sd:
                    continue
                if ed and pd > ed:
                    continue
            filtered.append(r)
        rows = filtered

    # Reordenar columnas según especificación
    columns = ['NUMERO_SOLICITUD','CLIENTE','NUMERO_CONTRATO','RFC','CURP','PRODUCTO_VERIFICADO','MARCAS','NOM','TIPO_DOCUMENTO','DOCUMENTO_EMITIDO','FECHA_DOCUMENTO_EMITIDO','VERIFICADOR','PEDIMENTO_IMPORTACION','FECHA_DESADUANAMIENTO','FECHA_VISITA','MODELOS','SOL_EMA','FOLIO_EMA','INSP_EMA']

    df = pd.DataFrame(rows)
    # Reindex to have columns in correct order (add missing)
    for c in columns:
        if c not in df.columns:
            df[c] = ''
    df = df[columns]

    os.makedirs(os.path.dirname(ruta_salida), exist_ok=True)
    df.to_excel(ruta_salida, index=False)

    try:
        with open(ruta_salida + '.json', 'w', encoding='utf-8') as f:
            json.dump(rows, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

    return ruta_salida





# REPORTE EMA #
def generar_reporte_ema(tabla_de_relacion_path, historial_path, ruta_salida, filtros=None, export_cache=None):
    """Genera el reporte EMA Excel con los encabezados solicitados.
    Si `export_cache` existe, lo usa como fuente preferente.
    """
    filtros = filtros or {}
    rows = []
    try:
        if export_cache and os.path.exists(export_cache):
            with open(export_cache, 'r', encoding='utf-8') as f:
                ec = json.load(f)
            rows = ec.get('ema', []) if isinstance(ec, dict) else []
        else:
            # Cargar tabla de relación directamente
            tabla = []
            if os.path.exists(tabla_de_relacion_path):
                with open(tabla_de_relacion_path, 'r', encoding='utf-8') as f:
                    try:
                        tabla = json.load(f)
                    except Exception:
                        tabla = []

            # Cargar clientes
            clientes = {}
            clientes_path = os.path.join(os.path.dirname(historial_path), 'Clientes.json')
            if os.path.exists(clientes_path):
                try:
                    with open(clientes_path, 'r', encoding='utf-8') as f:
                        cl = json.load(f)
                        if isinstance(cl, list):
                            for c in cl:
                                clientes[c.get('CLIENTE','').upper()] = c
                except Exception:
                    pass

            for r in tabla:
                try:
                    solicitud_full = r.get('ENCABEZADO', '') or r.get('SOLICITUD_ENCABEZADO', '') or r.get('SOLICITUD','')
                    sol_parts = str(solicitud_full).split()[-1] if solicitud_full else ''
                    cliente = r.get('EMPRESA','') or r.get('EMPRESA_VISITADA', r.get('CLIENTE',''))
                    cliente_key = (cliente or '').upper()
                    cliente_info = clientes.get(cliente_key, {})

                    rows.append({
                        'NUMERO_SOLICITUD': sol_parts,
                        'FECHA_INSPECCION': r.get('FECHA DE VERIFICACION') or r.get('FECHA_DE_VERIFICACION') or '',
                        'NUMERO_DICTAMEN': str(r.get('FOLIO','')).zfill(6) if str(r.get('FOLIO','')).strip() else '',
                        'NUMERO_CONTRATO': cliente_info.get('NÚMERO_DE_CONTRATO',''),
                        'TIPO_DOCUMENTO': r.get('TIPO DE DOCUMENTO') or r.get('TIPO_DE_DOCUMENTO',''),
                        'FECHA_DOCUMENTO_EMITIDO': r.get('FECHA DE VERIFICACION') or r.get('FECHA_DE_VERIFICACION') or '',
                        'PRODUCTO_VERIFICADO': r.get('DESCRIPCION',''),
                        'FECHA_DESADUANAMIENTO': r.get('FECHA DE ENTRADA') or r.get('FECHA_ENTRADA',''),
                        'FECHA_VISITA': r.get('FECHA DE VERIFICACION') or r.get('FECHA_DE_VERIFICACION') or '',
                        'OBSERVACIONES': 'N/A',
                        'INSPECTOR': r.get('VERIFICADOR') or r.get('INSPECTOR',''),
                        'PERSONAS_APOYO': 'N/A',
                        'NOM': r.get('CLASIF UVA') or r.get('CLASIF_UVA') or r.get('NOM','')
                    })
                except Exception:
                    continue
    except Exception:
        rows = []

    # Orden y columnas solicitadas
    columns = ['NUMERO_SOLICITUD','FECHA_INSPECCION','NUMERO_DICTAMEN','NUMERO_CONTRATO','TIPO_DOCUMENTO','FECHA_DOCUMENTO_EMITIDO','PRODUCTO_VERIFICADO','FECHA_DESADUANAMIENTO','FECHA_VISITA','OBSERVACIONES','INSPECTOR','PERSONAS_APOYO','NOM']

    df = pd.DataFrame(rows)
    for c in columns:
        if c not in df.columns:
            df[c] = ''
    df = df[columns]

    os.makedirs(os.path.dirname(ruta_salida), exist_ok=True)
    df.to_excel(ruta_salida, index=False)
    try:
        with open(ruta_salida + '.json', 'w', encoding='utf-8') as f:
            json.dump(rows, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

    return ruta_salida

