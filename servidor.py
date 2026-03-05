"""
Servidor Flask — Corrector de Formato Hoja de Ruta
====================================================
Recibe un archivo .xls/.xlsx con formato físico multi-fila,
lo corrige y devuelve el archivo corregido listo para descargar.

Uso:
  pip install -r requirements.txt
  python servidor.py

Endpoint:
  POST /corregir-hr
  Body: multipart/form-data  →  campo "data" con el archivo Excel
"""

import io
import re
from collections import Counter

import pandas as pd
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS

app = Flask(__name__)
CORS(app)  # Permite llamadas desde cualquier origen (HTML estático, otro dominio, etc.)


# ---------------------------------------------------------------------------
# Lógica de corrección (portada del código JS que usaba n8n)
# ---------------------------------------------------------------------------

SKIP_KEYWORDS = ['Parentesco', 'Firma', 'Motivo', 'Nombre y Apellido',
                 'Destinatario', 'Dirección', 'Nombre']


def detectar_numero_hr(data: list) -> int:
    """Auto-detecta el número de Hoja de Ruta buscando patrón YYYY - NNNNNN."""
    if len(data) > 1:
        row2 = data[1]
        for val in row2:
            if val is None:
                continue
            match = re.search(r'(\d{4})\s*-\s*(\d+)', str(val))
            if match:
                return int(match.group(2))
    return 0


def detectar_distrito(data: list) -> str:
    """Auto-detecta el distrito por la última palabra más frecuente en las direcciones (col K = índice 10)."""
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
    if contador:
        return contador.most_common(1)[0][0]
    return 'AREQUIPA'


def extraer_entradas(data: list, hr_numero: int, distrito: str) -> list:
    """Extrae todas las filas con código de barras LS...CW y sus datos asociados."""
    entries = []
    for i, row in enumerate(data):
        if not row or len(row) <= 3:
            continue
        cell_d = str(row[3] or '').strip()

        # Detectar código de barras: empieza con LS y termina con CW
        if re.match(r'^LS\d+CW$', cell_d, re.IGNORECASE):
            barcode = cell_d
            name = str(row[10] or '').strip() if len(row) > 10 else ''

            # Buscar dirección en columna K (índice 10) en filas siguientes
            address = ''
            for offset in range(1, 13):
                next_row = data[i + offset] if (i + offset) < len(data) else None
                if next_row is None:
                    break
                k_val = next_row[10] if len(next_row) > 10 else None
                if k_val:
                    k_str = str(k_val).strip()
                    if k_str and not any(kw in k_str for kw in SKIP_KEYWORDS):
                        address = k_str
                        break

            # Buscar código de HR en columna D (índice 3): número de 7-12 dígitos
            cod_hr = ''
            for offset in range(1, 13):
                next_row = data[i + offset] if (i + offset) < len(data) else None
                if next_row is None:
                    break
                d_val = next_row[3] if len(next_row) > 3 else None
                if d_val is not None:
                    if isinstance(d_val, (int, float)):
                        d_str = str(int(round(d_val)))
                    else:
                        d_str = str(d_val).strip()
                    if re.match(r'^\d{7,12}$', d_str):
                        cod_hr = d_str
                        break

            entries.append([hr_numero, distrito, barcode, name, address, cod_hr])

    return entries


def corregir_excel(file_bytes: bytes, filename: str) -> tuple:
    """
    Aplica la corrección completa al archivo Excel.
    Devuelve (bytes_corregido, nombre_archivo_salida).
    """
    # Leer el archivo (soporta .xls y .xlsx)
    try:
        df = pd.read_excel(io.BytesIO(file_bytes), header=None, engine='xlrd')
    except Exception:
        df = pd.read_excel(io.BytesIO(file_bytes), header=None, engine='openpyxl')

    data = df.values.tolist()

    hr_numero = detectar_numero_hr(data)
    distrito = detectar_distrito(data)
    entries = extraer_entradas(data, hr_numero, distrito)

    if not entries:
        raise ValueError(
            'No se encontraron entradas con código de barras (LS...CW) en el archivo. '
            'Verifica que el archivo tenga el formato esperado.'
        )

    # Construir DataFrame de salida con 6 columnas
    headers = [
        'NUMERO DE HOJA DE RUTA ',
        'DISTRITO',
        'CODIGO DE BARRAS ',
        'NOMBRE ',
        'DIRECCION',
        'CODIGO DE HOJA DE RUTA '
    ]
    df_out = pd.DataFrame(entries, columns=headers)

    # Serializar a bytes .xlsx
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_out.to_excel(writer, index=False, sheet_name='Sheet1')

        # Ajustar ancho de columnas
        ws = writer.sheets['Sheet1']
        col_widths = [24, 15, 18, 40, 60, 24]
        for col_idx, width in enumerate(col_widths, start=1):
            col_letter = ws.cell(row=1, column=col_idx).column_letter
            ws.column_dimensions[col_letter].width = width

    output.seek(0)

    # Nombre del archivo de salida
    out_name = re.sub(r'\.xlsx?$', '_CORREGIDO.xlsx', filename, flags=re.IGNORECASE)
    out_name = out_name.replace('formato_incorrecto', 'formato_correcto').replace('_incorrecto', '_correcto')

    return output.read(), out_name, len(entries), hr_numero, distrito


# ---------------------------------------------------------------------------
# Endpoints Flask
# ---------------------------------------------------------------------------

@app.route('/', methods=['GET'])
def index():
    return jsonify({
        'servicio': 'Corrector de Formato Hoja de Ruta',
        'version': '1.0',
        'endpoint': 'POST /corregir-hr  →  multipart/form-data, campo "data"'
    })


@app.route('/corregir-hr', methods=['POST'])
def corregir_hr():
    # Validar que se envió un archivo
    if 'data' not in request.files:
        return jsonify({'error': True, 'mensaje': 'No se recibió ningún archivo. Envía el Excel en el campo "data".'}), 400

    archivo = request.files['data']
    if archivo.filename == '':
        return jsonify({'error': True, 'mensaje': 'El archivo no tiene nombre.'}), 400

    ext = archivo.filename.rsplit('.', 1)[-1].lower()
    if ext not in ('xls', 'xlsx'):
        return jsonify({'error': True, 'mensaje': 'Solo se aceptan archivos .xls o .xlsx'}), 400

    try:
        file_bytes = archivo.read()
        out_bytes, out_name, total, hr_num, distrito = corregir_excel(file_bytes, archivo.filename)
    except ValueError as e:
        return jsonify({'error': True, 'mensaje': str(e)}), 400
    except Exception as e:
        return jsonify({'error': True, 'mensaje': f'Error interno al procesar el archivo: {e}'}), 500

    response = send_file(
        io.BytesIO(out_bytes),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=out_name
    )
    response.headers['X-Total-Entries'] = str(total)
    response.headers['X-Hoja-De-Ruta'] = str(hr_num)
    response.headers['X-Distrito'] = distrito
    return response


# ---------------------------------------------------------------------------

if __name__ == '__main__':
    print("=" * 55)
    print("  Corrector de Formato HR — Servidor Flask")
    print("  http://localhost:5000")
    print("  POST /corregir-hr  →  campo 'data' con el Excel")
    print("=" * 55)
    # host='0.0.0.0' para que sea accesible desde la red / servidor
    app.run(host='0.0.0.0', port=5000, debug=False)
