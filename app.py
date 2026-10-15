"""Web UI for Travel Document Extractor.
Run:  python app.py   -> open http://localhost:5000
Upload PDFs / images / a ZIP, see the JSON output on the same page.
"""
import io, json, tempfile, traceback
from pathlib import Path
from flask import Flask, request, jsonify, render_template, send_file
import travel_extractor as te

base_dir = Path(__file__).resolve().parent
te.load_dotenv_simple(base_dir / '.env')
te.TESSERACT_PATH = te.configure_tesseract()

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024  # 200 MB


def extract_rows(files, mode):
    """files: list of Path. Returns (rows, log)."""
    import os
    use_ai = mode == 'ai' or (mode == 'auto' and bool(os.getenv('AI_API_KEY')))
    rows, log = [], []
    for p in files:
        try:
            text, method = te.extract_local_text(p)
            rec = te.ai_parse(text, p.name, p) if use_ai else te.local_parse(text, p.name)
            if (not use_ai) and rec.get('Type of Document') == 'Hotel Booking' and method == 'image-ocr':
                rec = te.hotel_consensus(rec, getattr(te.ocr_image, 'last_variants', []), p.name)
            rec['Extraction Method'] = ('AI vision + ' + method) if use_ai else method
            rec['Source File'] = rec.get('Source File') or p.name
            ess = ['Vendor', 'Amount Charged', 'Passenger Name', 'Type of Document']
            rec['Review Needed'] = 'Yes' if any(rec.get(k) in ('', None) for k in ess) else 'No'
            rows.append(rec)
            log.append({'file': p.name, 'ok': True, 'msg': f"{rec.get('Type of Document')} ({method})"})
        except Exception as e:
            log.append({'file': p.name, 'ok': False, 'msg': f'{type(e).__name__}: {e}'})
            rec = {k: '' for k in te.COLUMNS}
            rec.update({'Source File': p.name, 'Extraction Method': 'error', 'Review Needed': 'Yes'})
            rows.append(rec)
    return rows, log


@app.route('/')
def index():
    import os
    return render_template('index.html', ai_available=bool(os.getenv('AI_API_KEY')))


@app.route('/api/extract', methods=['POST'])
def api_extract():
    uploads = request.files.getlist('files')
    mode = request.form.get('mode', 'auto')
    if not uploads:
        return jsonify(error='No files uploaded'), 400
    try:
        with tempfile.TemporaryDirectory() as td:
            td = Path(td)
            inbox = td / 'in'; inbox.mkdir()
            for i, f in enumerate(uploads):
                name = Path(f.filename).name or f'file{i}'
                f.save(inbox / name)
            files = te.collect_input(inbox, td / 'work')
            if not files:
                return jsonify(error='No supported files. Supported: ' + ', '.join(sorted(te.SUPPORTED)) + ', .zip'), 400
            rows, log = extract_rows(files, mode)
            out = td / 'out.json'
            te.write_json(rows, out)
            data = json.loads(out.read_text(encoding='utf-8'))
            xl = td / 'out.xlsx'
            te.write_excel(rows, xl)
            app.config['LAST_XLSX'] = xl.read_bytes()
        data['log'] = log
        return jsonify(data)
    except Exception as e:
        return jsonify(error=f'{type(e).__name__}: {e}', trace=traceback.format_exc()), 500


@app.route('/api/download.xlsx')
def download_xlsx():
    b = app.config.get('LAST_XLSX')
    if not b:
        return 'Nothing extracted yet', 404
    return send_file(io.BytesIO(b), as_attachment=True, download_name='travel_documents.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


if __name__ == '__main__':
    print('Open http://localhost:5000')
    app.run(host='127.0.0.1', port=5000, debug=False)
