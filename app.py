#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Servidor web para generación de Comentarios Técnicos – Pacific Control SAC.
Ejecutar: python app.py  →  abrir http://localhost:5002

Flujo en dos pasos:
  1) POST /analizar  → extrae los PDFs, evalúa contra los LMP y devuelve la tabla
                       (JSON) para que el usuario la revise y corrija en pantalla.
  2) POST /generar   → recibe la tabla ya corregida por el usuario (JSON) y genera
                       el Word con exactamente esos valores.
"""

import os, io, uuid, base64, shutil, importlib, sys
from flask import Flask, request, render_template, jsonify

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50 MB máx

OUTPUT_DIR = '/tmp/output'
os.makedirs(OUTPUT_DIR, exist_ok=True)


def _cargar_generar():
    """Recarga generar.py en cada request para asegurar código actualizado."""
    if 'generar' in sys.modules:
        del sys.modules['generar']
    import generar as gen
    return gen


@app.route('/')
def index():
    return render_template('index.html')


# ═══════════════════════════════════════════════════════════════════════════════
# PASO 1 — Analizar: extraer + evaluar, devolver tabla para revisión (sin Word)
# ═══════════════════════════════════════════════════════════════════════════════
@app.route('/analizar', methods=['POST'])
def analizar():
    try:
        # ── Validar campos ───────────────────────────────────────────────────
        numero = request.form.get('numero', '').strip()
        if not numero:
            return jsonify({'error': 'Ingresa el N° del Comentario Técnico.'}), 400

        pdfs_files = request.files.getlist('pdfs')
        lmp_file   = request.files.get('lmp')

        if not pdfs_files or all(f.filename == '' for f in pdfs_files):
            return jsonify({'error': 'Adjunta al menos un PDF de Informe de Ensayo.'}), 400
        if not lmp_file or lmp_file.filename == '':
            return jsonify({'error': 'Adjunta el archivo de LMP (Excel o PDF).'}), 400

        # ── Directorio temporal por request ──────────────────────────────────
        tmp = os.path.join(OUTPUT_DIR, uuid.uuid4().hex)
        os.makedirs(tmp)

        try:
            gen = _cargar_generar()

            # Guardar PDFs
            pdf_paths = []
            for f in pdfs_files:
                if f.filename.lower().endswith('.pdf'):
                    dest = os.path.join(tmp, f.filename)
                    f.save(dest)
                    pdf_paths.append(dest)

            if not pdf_paths:
                return jsonify({'error': 'Los archivos adjuntados no son PDFs válidos.'}), 400

            # Guardar LMP
            lmp_ext  = os.path.splitext(lmp_file.filename)[1].lower()
            lmp_path = os.path.join(tmp, f'lmp{lmp_ext}')
            lmp_file.save(lmp_path)

            # ── Extraer informes ──────────────────────────────────────────────
            informes = []
            log = []
            for p in sorted(pdf_paths):
                inf = gen.extraer_pdf(p)
                informes.append(inf)
                micro = ' [microbiológicos]' if inf['tiene_micro'] else ''
                log.append(f"N° {inf['numero']} — {len(inf['resultados'])} parámetros{micro}")

            # ── Cargar LMP y construir tabla evaluada (sin generar el Word) ────
            lmp   = gen.cargar_lmp(lmp_path)
            datos = gen.construir_datos(numero, informes, lmp)

            # ── Respuesta: la tabla para que el usuario la revise ─────────────
            return jsonify({
                'ok':            True,
                'datos':         datos,          # numero, encabezado, numeros, tiene_micro, filas, sin_lmp
                'log':           log,
                'lmp_count':     len(lmp),
                'total_params':  len(datos['filas']),
                'sin_lmp_total': len(datos['sin_lmp']),
            })

        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    except Exception as e:
        return jsonify({'error': f'Error inesperado: {str(e)}'}), 500


# ═══════════════════════════════════════════════════════════════════════════════
# PASO 2 — Generar: recibir la tabla ya corregida y producir el Word
# ═══════════════════════════════════════════════════════════════════════════════
@app.route('/generar', methods=['POST'])
def generar():
    try:
        payload = request.get_json(silent=True) or {}
        datos   = payload.get('datos')

        if not datos or not str(datos.get('numero', '')).strip():
            return jsonify({'error': 'Faltan datos del Comentario Técnico.'}), 400
        if not datos.get('filas'):
            return jsonify({'error': 'No hay parámetros en la tabla para generar el documento.'}), 400

        gen = _cargar_generar()

        numero     = str(datos['numero']).strip()
        nombre_doc = f'Comentario Técnico N° {numero}.docx'

        # Generar el Word en memoria y devolverlo en la MISMA respuesta (base64).
        # Así la descarga no depende de un archivo guardado en /tmp entre dos
        # peticiones distintas, algo poco fiable en serverless (Vercel) porque
        # cada petición puede atenderla una instancia diferente.
        buf = io.BytesIO()
        no_conformes = gen.generar_word_desde_datos(datos, buf)

        return jsonify({
            'ok':           True,
            'archivo':      nombre_doc,
            'docx_b64':     base64.b64encode(buf.getvalue()).decode('ascii'),
            'no_conformes': no_conformes,
            'total_params': len([f for f in datos['filas']
                                 if (f.get('analisis') or f.get('resultado'))]),
            'tiene_micro':  bool(datos.get('tiene_micro')),
        })

    except Exception as e:
        return jsonify({'error': f'Error inesperado: {str(e)}'}), 500


if __name__ == '__main__':
    print('\n  Comentarios Técnicos — Pacific Control SAC')
    print('  Abre tu navegador en:  http://localhost:5002\n')
    app.run(debug=False, port=5002)
