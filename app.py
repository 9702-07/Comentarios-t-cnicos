#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Servidor web para generación de Comentarios Técnicos – Pacific Control SAC.
Ejecutar: python app.py  →  abrir http://localhost:5000
"""

import os, uuid, shutil, importlib, sys
from flask import Flask, request, render_template, send_file, jsonify

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50 MB máx

OUTPUT_DIR = '/tmp/output'
os.makedirs(OUTPUT_DIR, exist_ok=True)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/generar', methods=['POST'])
def generar():
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
            # Recargar generar.py en cada request para asegurar código actualizado
            if 'generar' in sys.modules:
                del sys.modules['generar']
            import generar as gen

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

            # ── Cargar LMP ────────────────────────────────────────────────────
            lmp = gen.cargar_lmp(lmp_path)

            # ── Generar Word ──────────────────────────────────────────────────
            nombre_doc = f'Comentario Técnico N° {numero}.docx'
            output     = os.path.join(OUTPUT_DIR, nombre_doc)
            no_conf, sin_lmp = gen.generar_word(numero, informes, lmp, output)

            # ── Respuesta ─────────────────────────────────────────────────────
            return jsonify({
                'ok':           True,
                'archivo':      nombre_doc,
                'log':          log,
                'no_conformes': no_conf,
                'sin_lmp':      sin_lmp[:10],
                'sin_lmp_total': len(sin_lmp),
                'lmp_count':    len(lmp),
                'total_params': sum(len(i['resultados']) for i in informes),
                'tiene_micro':  any(i['tiene_micro'] for i in informes),
            })

        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    except Exception as e:
        return jsonify({'error': f'Error inesperado: {str(e)}'}), 500


@app.route('/descargar/<path:nombre>')
def descargar(nombre):
    path = os.path.join(OUTPUT_DIR, nombre)
    if not os.path.isfile(path):
        return 'Archivo no encontrado', 404
    return send_file(path, as_attachment=True, download_name=nombre)


if __name__ == '__main__':
    print('\n  Comentarios Técnicos — Pacific Control SAC')
    print('  Abre tu navegador en:  http://localhost:5000\n')
    app.run(debug=False, port=5002)
