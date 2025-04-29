from flask import Flask, render_template, request, redirect, url_for, send_file, flash
from werkzeug.utils import secure_filename
import os, io, pandas as pd, datetime
from docx import Document
from docxcompose.composer import Composer

import MAIN
# Prevent GUI pop-ups in Main
class _DummySplash:
    def destroy(self): pass
MAIN.show_splash2       = lambda *a,**k: _DummySplash()
MAIN.show_splash        = lambda *a,**k: _DummySplash()
MAIN.open_file          = lambda *a,**k: None
MAIN.messagebox.showinfo  = lambda *a,**k: None
MAIN.messagebox.showerror = lambda *a,**k: None
MAIN.root       = None
# Locate templates in our project folder
_APP_ROOT = os.path.dirname(os.path.abspath(__file__))
MAIN.resource_path = lambda rel: os.path.join(_APP_ROOT, 'templates', os.path.basename(rel))
# Avoid double-processing
MAIN.preprocess_excel = lambda path, filters=None: path

# Bring in Main’s tag-building primitives
from MAIN import (
    expand_template_to_3x3_fixed,
    expand_template_to_4x5_fixed_scaled,
    process_chunk,
    chunk_records,
    reapply_table_cell_spacing_only,
    FONT_SCHEME_HORIZONTAL,
    FONT_SCHEME_VERTICAL,
    FONT_SCHEME_MINI,
    SCALE_FACTOR,
    run_full_process_inventory_slips,
    _fetch_and_match as fetch_and_match_json
)

# Flask config
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXT   = {'xlsx','csv'}
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.secret_key = 'replace_with_a_secure_random_key'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def allowed_file(fn):
    return '.' in fn and fn.rsplit('.',1)[1].lower() in ALLOWED_EXT

# === INDEX ===
@app.route('/')
def index():
    df = None
    cache_p = os.path.join(app.config['UPLOAD_FOLDER'],'current_df.pkl')
    if os.path.exists(cache_p):
        try: df = pd.read_pickle(cache_p)
        except: df = None

    filter_cols = {
        'vendor':       'Vendor',
        'brand':        'Product Brand',
        'product_type': 'Product Type*',
        'lineage':      'Lineage',
        'weight':       'CombinedWeight',
        'strain':       'Product Strain'
    }

    options = {}
    for k,col in filter_cols.items():
        options[k] = sorted(df[col].dropna().unique()) if df is not None and col in df.columns else []

    sel = {k: request.args.get(k,'') for k in filter_cols}

    available = []
    if df is not None and 'Product Name*' in df.columns:
        tmp = df.copy()
        for k,col in filter_cols.items():
            v = sel[k]
            if v:
                tmp = tmp[tmp[col].astype(str)==v]
        available = sorted(tmp['Product Name*'].dropna().tolist())

    return render_template('index.html',
        options=options,
        selected_filters=sel,
        available=available,
        selected=[],
        json_url_value=request.args.get('json_url','')
    )

# === UPLOAD ===
@app.route('/upload', methods=['POST'])
def upload():
    if 'data_file' not in request.files:
        flash('No file provided','error')
        return redirect(url_for('index'))
    f = request.files['data_file']
    if f.filename=='' or not allowed_file(f.filename):
        flash('Invalid file type','error')
        return redirect(url_for('index'))

    fn = secure_filename(f.filename)
    save_path = os.path.join(app.config['UPLOAD_FOLDER'], fn)
    f.save(save_path)

    cleaned = MAIN.preprocess_excel(save_path)
    df = pd.read_excel(cleaned, engine='openpyxl')
    df.to_pickle(os.path.join(UPLOAD_FOLDER,'current_df.pkl'))

    flash('File uploaded and preprocessed','success')
    return redirect(url_for('index'))

# === IN-MEMORY TAG BUILDING ===
def build_tags_docx(df, template_type):
    # choose orientation
    if template_type in ('horizontal','vertical'):
        orientation = template_type
        tpl_file = f"{orientation}.docx"
        scheme   = FONT_SCHEME_HORIZONTAL if orientation=='horizontal' else FONT_SCHEME_VERTICAL
        fixed_buf = expand_template_to_3x3_fixed(os.path.join('templates',tpl_file))
        chunk_size = 9
    else:
        orientation = 'mini'
        scheme      = FONT_SCHEME_MINI
        fixed_buf   = expand_template_to_4x5_fixed_scaled(os.path.join('templates','mini.docx'),
                                                          scale_factor=SCALE_FACTOR)
        chunk_size  = 30

    records = df.to_dict('records')
    pages = []
    for chunk in chunk_records(records, chunk_size):
        pages.append(process_chunk((chunk, fixed_buf, scheme, orientation, SCALE_FACTOR)))

    if not pages:
        raise RuntimeError("No tags to generate")

    master = Document(io.BytesIO(pages[0]))
    comp   = Composer(master)
    for b in pages[1:]:
        comp.append(Document(io.BytesIO(b)))

    if orientation!='mini':
        reapply_table_cell_spacing_only(master, spacing_inches=0.03)

    buf = io.BytesIO()
    master.save(buf)
    buf.seek(0)
    return buf.getvalue()

# === GENERATE ===
@app.route('/generate/<template>', methods=['POST'])
def generate(template):
    cache_p = os.path.join(UPLOAD_FOLDER,'current_df.pkl')
    if not os.path.exists(cache_p):
        flash('Please upload first','error')
        return redirect(url_for('index'))
    df = pd.read_pickle(cache_p)

    tags = request.form.getlist('selected_tags')
    if not tags:
        flash('Please move at least one tag into Selected','error')
        return redirect(url_for('index'))

    df = df[df['Product Name*'].isin(tags)]

    try:
        if template in ('horizontal','vertical','mini'):
            data = build_tags_docx(df, template)
        elif template=='inventory':
            url = request.form.get('json_url','')
            if url:
                matched = fetch_and_match_json(url, df)
                data = run_full_process_inventory_slips(matched)
            else:
                data = run_full_process_inventory_slips(df)
        else:
            flash('Unknown template','error')
            return redirect(url_for('index'))
    except Exception as e:
        flash(f"Error generating document: {e}",'error')
        return redirect(url_for('index'))

    return send_file(
        io.BytesIO(data),
        as_attachment=True,
        download_name=f'{template}_tags.docx',
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )

if __name__=='__main__':
    app.run(debug=True, host='0.0.0.0', port=8000)
