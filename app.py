import os
import shutil
import threading
import queue
from flask import Flask, request, jsonify, send_from_directory, Response
from werkzeug.utils import secure_filename
from flask_cors import CORS # Para permitir a comunicação

# Importa a função principal que adaptamos do seu script
from seu_script_de_extracao import run_extraction_process_web_mode

# --- Configuração ---
UPLOAD_FOLDER = 'uploads'
GENERATED_REPORTS_FOLDER = 'generated_reports'
ALLOWED_EXTENSIONS = {'pdf', 'xls', 'xlsx'}

app = Flask(__name__, static_folder='static')
CORS(app) # Habilita CORS para a aplicação

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['GENERATED_REPORTS_FOLDER'] = GENERATED_REPORTS_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # Limite de 100MB para uploads

# Cria as pastas se não existirem
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(GENERATED_REPORTS_FOLDER, exist_ok=True)

# Fila para comunicação de progresso
progress_queue = queue.Queue()

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# --- Rota Principal (Interface) ---
@app.route('/')
def index():
    return send_from_directory(app.static_folder, 'index.html')

# --- Rota para Stream de Progresso ---
@app.route('/progress')
def progress():
    def generate():
        while True:
            try:
                message = progress_queue.get(timeout=30)
                if message == 'DONE':
                    yield f"data: {message}\n\n"
                    break
                yield f"data: {message}\n\n"
            except queue.Empty:
                yield f"data: ping\n\n"
    return Response(generate(), mimetype='text/event-stream')

# --- Rota para a Extração ---
@app.route('/upload_and_extract', methods=['POST'])
def upload_and_extract():
    if 'pdf_files' not in request.files or 'excel_file' not in request.files:
        return jsonify({"status": "error", "message": "Nenhum arquivo PDF ou Excel enviado."}), 400

    pdf_files = request.files.getlist('pdf_files')
    excel_file = request.files['excel_file']

    if not pdf_files or not excel_file:
        return jsonify({"status": "error", "message": "Nenhum arquivo PDF ou Excel selecionado."}), 400

    # --- Limpa pastas temporárias para uma nova execução ---
    for folder in [UPLOAD_FOLDER, GENERATED_REPORTS_FOLDER]:
        if os.path.exists(folder):
            shutil.rmtree(folder)
        os.makedirs(folder, exist_ok=True)

    # --- Salva os arquivos PDF ---
    for pdf in pdf_files:
        if pdf and allowed_file(pdf.filename):
            filename = secure_filename(pdf.filename)
            pdf.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
        else:
            return jsonify({"status": "error", "message": f"Tipo de arquivo PDF não permitido: {pdf.filename}"}), 400

    # --- Salva o arquivo Excel ---
    if excel_file and allowed_file(excel_file.filename):
        excel_filename = secure_filename(excel_file.filename)
        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_filename)
        excel_file.save(excel_path)
    else:
        return jsonify({"status": "error", "message": "Tipo de arquivo Excel não permitido ou nome inválido."}), 400

    try:
        # Limpa a fila de progresso
        while not progress_queue.empty():
            progress_queue.get()
        
        # Conta o total de PDFs
        total_pdfs = len([f for f in os.listdir(app.config['UPLOAD_FOLDER']) if f.lower().endswith('.pdf')])
        progress_queue.put(f"0/{total_pdfs}")
        
        # --- A MÁGICA ACONTECE AQUI ---
        # Chama a função adaptada do seu script, passando as pastas e a fila de progresso
        output_files = run_extraction_process_web_mode(
            pdf_upload_folder=app.config['UPLOAD_FOLDER'],
            excel_percentual_path=excel_path,
            output_report_folder=app.config['GENERATED_REPORTS_FOLDER'],
            progress_callback=lambda current, total: progress_queue.put(f"{current}/{total}")
        )
        # --------------------------------
        
        # Sinaliza conclusão
        progress_queue.put('DONE')

        # Prepara a resposta JSON com links para download
        response_data = {
            "status": "success",
            "message": "Extração e geração de relatórios concluídas com sucesso!",
            "download_links": {}
        }
        for key, filename in output_files.items():
            response_data["download_links"][key] = f"/download/{filename}"
            
        return jsonify(response_data), 200

    except Exception as e:
        print(f"Erro durante a extração: {e}")
        progress_queue.put('DONE')
        return jsonify({"status": "error", "message": f"Erro interno durante a extração: {str(e)}"}), 500

# --- Rota para Download dos Relatórios ---
@app.route('/download/<filename>')
def download_file(filename):
    return send_from_directory(app.config['GENERATED_REPORTS_FOLDER'], filename, as_attachment=True)

# --- Inicia o Servidor ---
if __name__ == '__main__':
    print("\nServidor rodando! Acesse http://127.0.0.1:5000 no seu navegador.\n")
    app.run(debug=True, port=5000)