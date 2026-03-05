"""
Servidor Flask — Corrector de Formato Hoja de Ruta
====================================================
Sirve la interfaz HTML en / y procesa archivos en POST /corregir-hr

Uso local:
  pip install -r requirements.txt
  python servidor.py
"""

import io
import re
import os
from collections import Counter

import pandas as pd
from flask import Flask, request, send_file, jsonify, Response
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

SKIP_KEYWORDS = ['Parentesco', 'Firma', 'Motivo', 'Nombre y Apellido',
                 'Destinatario', 'Dirección', 'Nombre']


def detectar_numero_hr(data):
    if len(data) > 1:
        for val in data[1]:
            if val is None:
                continue
            match = re.search(r'(\d{4})\s*-\s*(\d+)', str(val))
            if match:
                return int(match.group(2))
    return 0


def detectar_distrito(data):
    contador = Counter()
    for row in data:
        if not row or len(row) <= 10:
            continue
        val = str(row[10] or '').strip()
        if not val or any(kw in val for kw in SKIP_KEYWORDS):
            continue
        partes = val.split()
        if partes:
            ultima = partes[-1].upper()
            if len(ultima) > 3:
                contador[ultima] += 1
    return contador.most_common(1)[0][0] if contador else 'AREQUIPA'


def extraer_entradas(data, hr_numero, distrito):
    entries = []
    for i, row in enumerate(data):
        if not row or len(row) <= 3:
            continue
        cell_d = str(row[3] or '').strip()
        if re.match(r'^LS\d+CW$', cell_d, re.IGNORECASE):
            barcode = cell_d
            name = str(row[10] or '').strip() if len(row) > 10 else ''
            address = ''
            for offset in range(1, 13):
                next_row = data[i + offset] if (i + offset) < len(data) else None
                if not next_row:
                    break
                k_val = next_row[10] if len(next_row) > 10 else None
                if k_val:
                    k_str = str(k_val).strip()
                    if k_str and not any(kw in k_str for kw in SKIP_KEYWORDS):
                        address = k_str
                        break
            cod_hr = ''
            for offset in range(1, 13):
                next_row = data[i + offset] if (i + offset) < len(data) else None
                if not next_row:
                    break
                d_val = next_row[3] if len(next_row) > 3 else None
                if d_val is not None:
                    d_str = str(int(round(d_val))) if isinstance(d_val, float) else str(d_val).strip()
                    if re.match(r'^\d{7,12}$', d_str):
                        cod_hr = d_str
                        break
            entries.append([hr_numero, distrito, barcode, name, address, cod_hr])
    return entries


def corregir_excel(file_bytes, filename):
    try:
        df = pd.read_excel(io.BytesIO(file_bytes), header=None, engine='xlrd')
    except Exception:
        df = pd.read_excel(io.BytesIO(file_bytes), header=None, engine='openpyxl')

    data = df.values.tolist()
    hr_numero = detectar_numero_hr(data)
    distrito = detectar_distrito(data)
    entries = extraer_entradas(data, hr_numero, distrito)

    if not entries:
        raise ValueError('No se encontraron entradas con código de barras (LS...CW).')

    headers = ['NUMERO DE HOJA DE RUTA ', 'DISTRITO', 'CODIGO DE BARRAS ',
               'NOMBRE ', 'DIRECCION', 'CODIGO DE HOJA DE RUTA ']
    df_out = pd.DataFrame(entries, columns=headers)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_out.to_excel(writer, index=False, sheet_name='Sheet1')
        ws = writer.sheets['Sheet1']
        for col_idx, width in enumerate([24, 15, 18, 40, 60, 24], start=1):
            ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = width

    output.seek(0)
    out_name = re.sub(r'\.xlsx?$', '_CORREGIDO.xlsx', filename, flags=re.IGNORECASE)
    return output.read(), out_name, len(entries), hr_numero, distrito


# ─── Rutas ────────────────────────────────────────────────────────────────────

@app.route('/', methods=['GET'])
def index():
    """Sirve la interfaz HTML directamente."""
    html_path = os.path.join(os.path.dirname(__file__), 'index.html')
    with open(html_path, 'r', encoding='utf-8') as f:
        return Response(f.read(), mimetype='text/html')


@app.route('/corregir-hr', methods=['POST'])
def corregir_hr():
    if 'data' not in request.files:
        return jsonify({'error': True, 'mensaje': 'Envía el archivo en el campo "data".'}), 400

    archivo = request.files['data']
    ext = archivo.filename.rsplit('.', 1)[-1].lower()
    if ext not in ('xls', 'xlsx'):
        return jsonify({'error': True, 'mensaje': 'Solo se aceptan .xls o .xlsx'}), 400

    try:
        out_bytes, out_name, total, hr_num, distrito = corregir_excel(archivo.read(), archivo.filename)
    except ValueError as e:
        return jsonify({'error': True, 'mensaje': str(e)}), 400
    except Exception as e:
        return jsonify({'error': True, 'mensaje': f'Error interno: {e}'}), 500

    response = send_file(io.BytesIO(out_bytes),
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                         as_attachment=True,
                         download_name=out_name)
    response.headers['X-Total-Entries'] = str(total)
    response.headers['X-Hoja-De-Ruta'] = str(hr_num)
    response.headers['X-Distrito'] = distrito
    return response


if __name__ == '__main__':
    print("=" * 50)
    print("  Corrector HR → http://localhost:5000")
    print("=" * 50)
    app.run(host='0.0.0.0', port=5000, debug=False)
