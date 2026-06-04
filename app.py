#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Servidor web para generación de Comentarios Técnicos – Pacific Control SAC.
Ejecutar: python app.py  →  abrir http://localhost:5002

Flujo en dos pasos:
  1) POST /analizar  → extrae los PDFs, evalúa contra DECRETOS.xlsx (fijo en el
                       servidor) y devuelve la tabla (JSON) para que el usuario
                       la revise y corrija en pantalla, incluyendo la conclusión.
  2) POST /generar   → recibe la tabla y la conclusión ya corregidas por el
                       usuario (JSON) y genera el Word con exactamente esos valores.
"""

import os, io, uuid, base64, shutil, sys
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
        if not pdfs_files or all(f.filename == '' for f in pdfs_files):
            return jsonify({'error': 'Adjunta al menos un PDF de Informe de Ensayo.'}), 400

        # ── Directorio temporal por request ──────────────────────────────────
        tmp = os.path.join(OUTPUT_DIR, uuid.uuid4().hex)
        os.makedirs(tmp)

        try:
            gen = _cargar_generar()

            # Verificar que el archivo de LMP existe en el servidor
            if not os.path.isfile(gen.LMP_PATH):
                return jsonify({'error': 'No se encontró DECRETOS.xlsx en el servidor.'}), 500

            # Guardar PDFs
            pdf_paths = []
            for f in pdfs_files:
                if f.filename.lower().endswith('.pdf'):
                    dest = os.path.join(tmp, f.filename)
                    f.save(dest)
                    pdf_paths.append(dest)

            if not pdf_paths:
                return jsonify({'error': 'Los archivos adjuntados no son PDFs válidos.'}), 400

            # ── Extraer informes ──────────────────────────────────────────────
            informes = []
            log = []
            for p in sorted(pdf_paths):
                inf = gen.extraer_pdf(p)
                informes.append(inf)
                micro = ' [microbiológicos]' if inf['tiene_micro'] else ''
                log.append(f"N° {inf['numero']} — {len(inf['resultados'])} parámetros{micro}")

            # ── Cargar LMP desde DECRETOS.xlsx (fijo en el servidor) ──────────
            lmp   = gen.cargar_lmp(gen.LMP_PATH)
            datos = gen.construir_datos(numero, informes, lmp)

            # ── Respuesta: tabla + conclusión para que el usuario las revise ──
            return jsonify({
                'ok':            True,
                'datos':         datos,
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
# PASO 2 — Generar: recibir tabla + conclusión corregidas y producir el Word
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

        # ── Generar ambos PDFs en memoria ─────────────────────────────────────
        no_conformes = [f.get('analisis', '').strip()
                        for f in datos.get('filas', [])
                        if (f.get('evaluacion') or '').strip() == 'NO CONFORME']

        buf_pdf_cf = io.BytesIO()
        gen.generar_pdf(datos, buf_pdf_cf, con_fondo=True)

        buf_pdf_sf = io.BytesIO()
        gen.generar_pdf(datos, buf_pdf_sf, con_fondo=False)

        base = f'Comentario Técnico N° {numero}'
        return jsonify({
            'ok':            True,
            'no_conformes':  no_conformes,
            'total_params':  len([f for f in datos['filas']
                                  if (f.get('analisis') or f.get('resultado'))]),
            'tiene_micro':   bool(datos.get('tiene_micro')),
            'nombre_pdf_cf': f'{base} (Con Fondo).pdf',
            'nombre_pdf_sf': f'{base} (Sin Fondo).pdf',
            'pdf_cf_b64':    base64.b64encode(buf_pdf_cf.getvalue()).decode('ascii'),
            'pdf_sf_b64':    base64.b64encode(buf_pdf_sf.getvalue()).decode('ascii'),
        })

    except Exception as e:
        return jsonify({'error': f'Error inesperado: {str(e)}'}), 500


if __name__ == '__main__':
    print('\n  Comentarios Técnicos — Pacific Control SAC')
    print('  Abre tu navegador en:  http://localhost:5002\n')
    app.run(debug=False, port=5002)
