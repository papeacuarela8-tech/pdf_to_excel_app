import os
import shutil
import tempfile
from flask import Flask, request, render_template, send_file, redirect, url_for, flash
from werkzeug.utils import secure_filename
from parser import parse_pdf_invoice

app = Flask(__name__)
app.secret_key = 'papeleria_acuarela_secret_key'
app.config['MAX_CONTENT_LENGTH'] = 25 * 1024 * 1024

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        flash('No se seleccionó ningún archivo', 'error')
        return redirect(url_for('index'))
    
    file = request.files['file']
    if file.filename == '':
        flash('No se seleccionó ningún archivo', 'error')
        return redirect(url_for('index'))
    
    if file and file.filename.lower().endswith('.pdf'):
        try:
            markup_percent = float(request.form.get('markup_percent', 65))
            if not 0 <= markup_percent <= 1000:
                raise ValueError
        except (TypeError, ValueError):
            flash('El margen de venta debe ser un porcentaje entre 0 y 1000.', 'error')
            return redirect(url_for('index'))

        # Create a temporary directory for processing
        temp_dir = tempfile.mkdtemp()
        safe_filename = secure_filename(file.filename) or 'factura.pdf'
        pdf_path = os.path.join(temp_dir, safe_filename)
        file.save(pdf_path)
        
        base_name = os.path.splitext(safe_filename)[0]
        excel_name = f"Precios_Venta_{base_name}.xlsx"
        excel_path = os.path.join(temp_dir, excel_name)
        
        try:
            # Parse the invoice PDF and generate the styled Excel
            parse_pdf_invoice(pdf_path, excel_path, markup_percent=markup_percent)
            
            # Send file for download
            response = send_file(
                excel_path,
                as_attachment=True,
                download_name=excel_name,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
            # The response must finish streaming before its temporary files are removed.
            response.call_on_close(lambda: shutil.rmtree(temp_dir, ignore_errors=True))
            return response
        except Exception as e:
            # Clean up temp_dir on failure
            try:
                shutil.rmtree(temp_dir, ignore_errors=True)
            except Exception:
                pass
            flash(f'Error al procesar el archivo: {str(e)}', 'error')
            return redirect(url_for('index'))
        
    flash('El archivo debe ser en formato PDF', 'error')
    return redirect(url_for('index'))

@app.errorhandler(413)
def file_too_large(_error):
    flash('El PDF supera el tamaño máximo permitido de 25 MB.', 'error')
    return redirect(url_for('index'))

if __name__ == '__main__':
    print("Iniciando servidor local de Papelería Acuarela en http://localhost:5000")
    app.run(host='127.0.0.1', port=5000, debug=True)
