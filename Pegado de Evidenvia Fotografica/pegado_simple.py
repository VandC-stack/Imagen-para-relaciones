import os
from docx import Document
from main import (
    obtener_rutas,
    indexar_imagenes,
    buscar_imagen_index,
    insertar_imagen_con_transparencia,
    extraer_codigos,
    extraer_codigos_pdf,
    normalizar_cadena_alnum_mayus,
    norm_path_key,
    insertar_imagenes_en_pdf_placeholder,
)
from registro_fallos import registrar_fallo, limpiar_registro, mostrar_registro, LOG_FILE


def procesar_simple():
    limpiar_registro()

    ruta_docs, ruta_imgs = obtener_rutas()
    if not ruta_docs or not ruta_imgs:
        return

    # Índice normal de imágenes para el modo simple
    index = indexar_imagenes(ruta_imgs)

    # Ahora modo simple procesa Word y PDF
    archivos = [
        f for f in os.listdir(ruta_docs)
        if (f.lower().endswith(".docx") or f.lower().endswith(".pdf")) and not f.startswith("~$")
    ]

    for archivo in archivos:
        ruta_doc = os.path.join(ruta_docs, archivo)
        ext = os.path.splitext(archivo)[1].lower()

        # ========================================
        # WORD
        # ========================================
        if ext == ".docx":
            print(f"Procesando documento (modo simple DOCX): {ruta_doc}")

            doc = Document(ruta_doc)
            codigos = extraer_codigos(doc)

            if not codigos:
                registrar_fallo(archivo)
                continue

            imagen_insertada = False
            usadas_paths = set()
            usadas_bases = set()

            for p in doc.paragraphs:
                 if "${imagen}" in (p.text or ""):
                    p.clear()
                    run = p.add_run()

                    for codigo in codigos:
                        img_path = buscar_imagen_index(index, codigo, usadas_paths, usadas_bases)
                        if img_path:
                            usadas_paths.add(norm_path_key(img_path))
                            usar_base = normalizar_cadena_alnum_mayus(os.path.splitext(os.path.basename(img_path))[0])
                            usadas_bases.add(usar_base)
                            insertar_imagen_con_transparencia(run, img_path)
                            imagen_insertada = True

                    break

            if not imagen_insertada:
                registrar_fallo(archivo)

            doc.save(ruta_doc)
            print(f"Documento DOCX actualizado: {ruta_doc}")
            continue

        # ========================================
        # PDF
        # ========================================
        if ext == ".pdf":
            print(f"Procesando documento (modo simple PDF): {ruta_doc}")

            codigos = extraer_codigos_pdf(ruta_doc)

            if not codigos:
                registrar_fallo(archivo)
                continue

            rutas_imagenes = []
            usadas_paths = set()
            usadas_bases = set()

            for codigo in codigos:
                img_path = buscar_imagen_index(index, codigo, usadas_paths, usadas_bases)
                if img_path:
                    rutas_imagenes.append(img_path)
                    usadas_paths.add(norm_path_key(img_path))
                    usar_base = normalizar_cadena_alnum_mayus(os.path.splitext(os.path.basename(img_path))[0])
                    usadas_bases.add(usar_base)

            if not rutas_imagenes:
                registrar_fallo(archivo)
                continue

            exito = insertar_imagenes_en_pdf_placeholder(ruta_doc, rutas_imagenes)
            if not exito:
                registrar_fallo(archivo)

    mostrar_registro()
    if os.path.exists(LOG_FILE):
        os.startfile(LOG_FILE)  