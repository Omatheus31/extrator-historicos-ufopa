import os
import shutil
import queue
import argparse
import threading
from flask import Flask, request, jsonify, send_from_directory, Response
from werkzeug.utils import secure_filename
from flask_cors import CORS

from seu_script_de_extracao import run_extraction_process_web_mode


ALLOWED_EXTENSIONS = {'pdf', 'xls', 'xlsx'}

# Delay (em segundos) antes de apagar uploads/relatórios gerados após conclusão
# Pode ser sobrescrito pela variável de ambiente `EXTRACTION_CLEANUP_SECONDS`
CLEANUP_DELAY_SECONDS = int(os.getenv('EXTRACTION_CLEANUP_SECONDS', '120'))


def make_app(base_dir: str):
    upload_folder = os.path.join(base_dir, 'uploads')
    generated_folder = os.path.join(base_dir, 'generated_reports')

    app = Flask(__name__, static_folder='static')
    CORS(app)

    app.config['UPLOAD_FOLDER'] = upload_folder
    app.config['GENERATED_REPORTS_FOLDER'] = generated_folder
    app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB

    os.makedirs(upload_folder, exist_ok=True)
    os.makedirs(generated_folder, exist_ok=True)

    return app


def get_base_dir_from_args_or_env():
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument('--base-dir', dest='base_dir', help='Pasta base para uploads e reports')
    args, _ = parser.parse_known_args()
    base_dir = args.base_dir or os.getenv('EXTRACTION_BASE_DIR') or os.getcwd()
    return os.path.abspath(base_dir)


BASE_DIR = get_base_dir_from_args_or_env()
app = make_app(BASE_DIR)

# Fila para comunicação de progresso
progress_queue = queue.Queue()


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def index():
    return send_from_directory(app.static_folder, 'index.html')


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


@app.route('/upload_and_extract', methods=['POST'])
def upload_and_extract():
    skip_percentuais = request.form.get('skip_percentuals') in ('1', 'true', 'on', 'yes')

    if 'pdf_files' not in request.files:
        return jsonify({"status": "error", "message": "Nenhum arquivo PDF enviado."}), 400

    pdf_files = request.files.getlist('pdf_files')
    excel_file = request.files.get('excel_file') if 'excel_file' in request.files else None

    if not pdf_files:
        return jsonify({"status": "error", "message": "Nenhum arquivo PDF selecionado."}), 400

    if not skip_percentuais and not excel_file:
        return jsonify({"status": "error", "message": "Nenhum arquivo Excel de percentuais selecionado."}), 400

    # limpa pastas temporárias
    for folder in [app.config['UPLOAD_FOLDER'], app.config['GENERATED_REPORTS_FOLDER']]:
        if os.path.exists(folder):
            shutil.rmtree(folder)
        os.makedirs(folder, exist_ok=True)

    # salva PDFs
    for pdf in pdf_files:
        if pdf and allowed_file(pdf.filename):
            filename = secure_filename(pdf.filename)
            pdf.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
        else:
            return jsonify({"status": "error", "message": f"Tipo de arquivo PDF não permitido: {pdf.filename}"}), 400

    excel_path = None
    if excel_file:
        if allowed_file(excel_file.filename):
            excel_filename = secure_filename(excel_file.filename)
            excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_filename)
            excel_file.save(excel_path)
        else:
            return jsonify({"status": "error", "message": "Tipo de arquivo Excel não permitido ou nome inválido."}), 400

    try:
        # limpa fila de progresso
        while not progress_queue.empty():
            progress_queue.get()

        total_pdfs = len([f for f in os.listdir(app.config['UPLOAD_FOLDER']) if f.lower().endswith('.pdf')])
        progress_queue.put(f"0/{total_pdfs}")

        output_files = run_extraction_process_web_mode(
            pdf_upload_folder=app.config['UPLOAD_FOLDER'],
            excel_percentual_path=excel_path,
            output_report_folder=app.config['GENERATED_REPORTS_FOLDER'],
            progress_callback=lambda current, total: progress_queue.put(f"{current}/{total}")
        )

        progress_queue.put('DONE')

        response_data = {
            "status": "success",
            "message": "Extração e geração de relatórios concluídas com sucesso!",
            "download_links": {}
        }
        for key, filename in output_files.items():
            response_data["download_links"][key] = f"/download/{filename}"

        # Agende limpeza dos uploads e dos arquivos gerados após um delay
        try:
            upload_folder = app.config['UPLOAD_FOLDER']
            generated_folder = app.config['GENERATED_REPORTS_FOLDER']
            def cleanup_dirs(u_folder, g_folder):
                try:
                    for folder in (u_folder, g_folder):
                        if os.path.exists(folder):
                            shutil.rmtree(folder)
                        os.makedirs(folder, exist_ok=True)
                    print(f"Limpeza concluída: {u_folder}, {g_folder}")
                except Exception as e:
                    print(f"Erro durante limpeza programada: {e}")

            def schedule_cleanup(delay, u_folder, g_folder):
                t = threading.Timer(delay, cleanup_dirs, args=(u_folder, g_folder))
                t.daemon = True
                t.start()
                print(f"Agendada limpeza em {delay}s para: {u_folder}, {g_folder}")

            schedule_cleanup(CLEANUP_DELAY_SECONDS, upload_folder, generated_folder)
        except Exception as e:
            print(f"Não foi possível agendar limpeza: {e}")

        return jsonify(response_data), 200

    except Exception as e:
        print(f"Erro durante a extração: {e}")
        progress_queue.put('DONE')
        return jsonify({"status": "error", "message": f"Erro interno durante a extração: {str(e)}"}), 500


@app.route('/download/<filename>')
def download_file(filename):
    return send_from_directory(app.config['GENERATED_REPORTS_FOLDER'], filename, as_attachment=True)


if __name__ == '__main__':
    print(f"\nServidor rodando! Base dir: {BASE_DIR}")
    print("Acesse http://127.0.0.1:5000 no seu navegador.\n")
    app.run(debug=False, port=5000, use_reloader=False)
import os
import shutil
import queue
import argparse
from flask import Flask, request, jsonify, send_from_directory, Response
from werkzeug.utils import secure_filename
from flask_cors import CORS

from seu_script_de_extracao import run_extraction_process_web_mode


ALLOWED_EXTENSIONS = {'pdf', 'xls', 'xlsx'}


def make_app(base_dir: str):
    upload_folder = os.path.join(base_dir, 'uploads')
    generated_folder = os.path.join(base_dir, 'generated_reports')

    app = Flask(__name__, static_folder='static')
    CORS(app)

    app.config['UPLOAD_FOLDER'] = upload_folder
    app.config['GENERATED_REPORTS_FOLDER'] = generated_folder
    app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB

    os.makedirs(upload_folder, exist_ok=True)
    os.makedirs(generated_folder, exist_ok=True)

    return app


def get_base_dir_from_args_or_env():
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument('--base-dir', dest='base_dir', help='Pasta base para uploads e reports')
    args, _ = parser.parse_known_args()
    base_dir = args.base_dir or os.getenv('EXTRACTION_BASE_DIR') or os.getcwd()
    return os.path.abspath(base_dir)


BASE_DIR = get_base_dir_from_args_or_env()
app = make_app(BASE_DIR)

# Fila para comunicação de progresso
progress_queue = queue.Queue()


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def index():
    return send_from_directory(app.static_folder, 'index.html')


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


@app.route('/upload_and_extract', methods=['POST'])
def upload_and_extract():
    skip_percentuais = request.form.get('skip_percentuals') in ('1', 'true', 'on', 'yes')

    if 'pdf_files' not in request.files:
        return jsonify({"status": "error", "message": "Nenhum arquivo PDF enviado."}), 400

    pdf_files = request.files.getlist('pdf_files')
    excel_file = request.files.get('excel_file') if 'excel_file' in request.files else None

    if not pdf_files:
        return jsonify({"status": "error", "message": "Nenhum arquivo PDF selecionado."}), 400

    if not skip_percentuais and not excel_file:
        return jsonify({"status": "error", "message": "Nenhum arquivo Excel de percentuais selecionado."}), 400

    # limpa pastas temporárias
    for folder in [app.config['UPLOAD_FOLDER'], app.config['GENERATED_REPORTS_FOLDER']]:
        if os.path.exists(folder):
            shutil.rmtree(folder)
        os.makedirs(folder, exist_ok=True)

    # salva PDFs
    for pdf in pdf_files:
        if pdf and allowed_file(pdf.filename):
            filename = secure_filename(pdf.filename)
            pdf.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
        else:
            return jsonify({"status": "error", "message": f"Tipo de arquivo PDF não permitido: {pdf.filename}"}), 400

    excel_path = None
    if excel_file:
        if allowed_file(excel_file.filename):
            excel_filename = secure_filename(excel_file.filename)
            excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_filename)
            excel_file.save(excel_path)
        else:
            return jsonify({"status": "error", "message": "Tipo de arquivo Excel não permitido ou nome inválido."}), 400

    try:
        # limpa fila de progresso
        while not progress_queue.empty():
            progress_queue.get()

        total_pdfs = len([f for f in os.listdir(app.config['UPLOAD_FOLDER']) if f.lower().endswith('.pdf')])
        progress_queue.put(f"0/{total_pdfs}")

        output_files = run_extraction_process_web_mode(
            pdf_upload_folder=app.config['UPLOAD_FOLDER'],
            excel_percentual_path=excel_path,
            output_report_folder=app.config['GENERATED_REPORTS_FOLDER'],
            progress_callback=lambda current, total: progress_queue.put(f"{current}/{total}")
        )

        progress_queue.put('DONE')

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


@app.route('/download/<filename>')
def download_file(filename):
    return send_from_directory(app.config['GENERATED_REPORTS_FOLDER'], filename, as_attachment=True)


if __name__ == '__main__':
    print(f"\nServidor rodando! Base dir: {BASE_DIR}")
    print("Acesse http://127.0.0.1:5000 no seu navegador.\n")
    app.run(debug=True, port=5000, use_reloader=False)
import os
import shutil
import queue
import argparse
from flask import Flask, request, jsonify, send_from_directory, Response
from werkzeug.utils import secure_filename
from flask_cors import CORS

from seu_script_de_extracao import run_extraction_process_web_mode


ALLOWED_EXTENSIONS = {'pdf', 'xls', 'xlsx'}


def make_app(base_dir: str):
    upload_folder = os.path.join(base_dir, 'uploads')
    generated_folder = os.path.join(base_dir, 'generated_reports')

    app = Flask(__name__, static_folder='static')
    CORS(app)
import os
import shutil
import queue
import argparse
from flask import Flask, request, jsonify, send_from_directory, Response
from werkzeug.utils import secure_filename
from flask_cors import CORS

from seu_script_de_extracao import run_extraction_process_web_mode


ALLOWED_EXTENSIONS = {'pdf', 'xls', 'xlsx'}


def make_app(base_dir: str):
    upload_folder = os.path.join(base_dir, 'uploads')
    generated_folder = os.path.join(base_dir, 'generated_reports')

    app = Flask(__name__, static_folder='static')
    CORS(app)

    app.config['UPLOAD_FOLDER'] = upload_folder
    app.config['GENERATED_REPORTS_FOLDER'] = generated_folder
    app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB

    os.makedirs(upload_folder, exist_ok=True)
    os.makedirs(generated_folder, exist_ok=True)

    return app


def get_base_dir_from_args_or_env():
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument('--base-dir', dest='base_dir', help='Pasta base para uploads e reports')
    args, _ = parser.parse_known_args()
    base_dir = args.base_dir or os.getenv('EXTRACTION_BASE_DIR') or os.getcwd()
    return os.path.abspath(base_dir)


BASE_DIR = get_base_dir_from_args_or_env()
app = make_app(BASE_DIR)

# Fila para comunicação de progresso
progress_queue = queue.Queue()


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def index():
    return send_from_directory(app.static_folder, 'index.html')


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


@app.route('/upload_and_extract', methods=['POST'])
def upload_and_extract():
    skip_percentuais = request.form.get('skip_percentuals') in ('1', 'true', 'on', 'yes')

    if 'pdf_files' not in request.files:
        return jsonify({"status": "error", "message": "Nenhum arquivo PDF enviado."}), 400

    pdf_files = request.files.getlist('pdf_files')
    excel_file = request.files.get('excel_file') if 'excel_file' in request.files else None

    if not pdf_files:
        return jsonify({"status": "error", "message": "Nenhum arquivo PDF selecionado."}), 400

    if not skip_percentuais and not excel_file:
        return jsonify({"status": "error", "message": "Nenhum arquivo Excel de percentuais selecionado."}), 400

    # limpa pastas temporárias
    for folder in [app.config['UPLOAD_FOLDER'], app.config['GENERATED_REPORTS_FOLDER']]:
        if os.path.exists(folder):
            shutil.rmtree(folder)
        os.makedirs(folder, exist_ok=True)

    # salva PDFs
    for pdf in pdf_files:
        if pdf and allowed_file(pdf.filename):
            filename = secure_filename(pdf.filename)
            pdf.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
        else:
            return jsonify({"status": "error", "message": f"Tipo de arquivo PDF não permitido: {pdf.filename}"}), 400

    excel_path = None
    if excel_file:
        if allowed_file(excel_file.filename):
            excel_filename = secure_filename(excel_file.filename)
            excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_filename)
            excel_file.save(excel_path)
        else:
            return jsonify({"status": "error", "message": "Tipo de arquivo Excel não permitido ou nome inválido."}), 400

    try:
        # limpa fila de progresso
        while not progress_queue.empty():
            progress_queue.get()

        total_pdfs = len([f for f in os.listdir(app.config['UPLOAD_FOLDER']) if f.lower().endswith('.pdf')])
        progress_queue.put(f"0/{total_pdfs}")

        output_files = run_extraction_process_web_mode(
            pdf_upload_folder=app.config['UPLOAD_FOLDER'],
            excel_percentual_path=excel_path,
            output_report_folder=app.config['GENERATED_REPORTS_FOLDER'],
            progress_callback=lambda current, total: progress_queue.put(f"{current}/{total}")
        )

        progress_queue.put('DONE')

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


@app.route('/download/<filename>')
def download_file(filename):
    return send_from_directory(app.config['GENERATED_REPORTS_FOLDER'], filename, as_attachment=True)


if __name__ == '__main__':
    print(f"\nServidor rodando! Base dir: {BASE_DIR}")
    print("Acesse http://127.0.0.1:5000 no seu navegador.\n")
    app.run(debug=True, port=5000, use_reloader=False)
import argparse
from flask import Flask, request, jsonify, send_from_directory, Response
from werkzeug.utils import secure_filename
from flask_cors import CORS # Para permitir a comunicação

# Importa a função principal que adaptamos do seu script
from seu_script_de_extracao import run_extraction_process_web_mode


# --- Configuração ---
# Permite sobrescrever a pasta base onde os dados (uploads/generated_reports)
# serão criados. Prioriza argumento de linha de comando, depois variável de
# ambiente `EXTRACTION_BASE_DIR`, e por fim o diretório atual de trabalho.
ALLOWED_EXTENSIONS = {'pdf', 'xls', 'xlsx'}


def make_app(base_dir: str):
    """Cria a instância Flask usando `base_dir` como diretório base para
    `uploads/` e `generated_reports/`. Isso permite que quem clona o repo
    escolha onde armazenar os dados."""

    upload_folder = os.path.join(base_dir, 'uploads')
    generated_folder = os.path.join(base_dir, 'generated_reports')

    app = Flask(__name__, static_folder='static')
    CORS(app)

    app.config['UPLOAD_FOLDER'] = upload_folder
    app.config['GENERATED_REPORTS_FOLDER'] = generated_folder
    app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB

    # Cria as pastas se não existirem
    os.makedirs(upload_folder, exist_ok=True)
    os.makedirs(generated_folder, exist_ok=True)

    return app


def get_base_dir_from_args_or_env():
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument('--base-dir', dest='base_dir', help='Pasta base para uploads e reports')
    args, _ = parser.parse_known_args()
    base_dir = args.base_dir or os.getenv('EXTRACTION_BASE_DIR') or os.getcwd()
    return os.path.abspath(base_dir)


BASE_DIR = get_base_dir_from_args_or_env()
app = make_app(BASE_DIR)


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
    for folder in [app.config['UPLOAD_FOLDER'], app.config['GENERATED_REPORTS_FOLDER']]:
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
    # Se o usuário passar --base-dir ao executar `python app.py --base-dir <pasta>`
    # ou configurar a variável de ambiente `EXTRACTION_BASE_DIR`, a app
    # já terá sido criada apontando para essa pasta.
    print(f"\nServidor rodando! Base dir: {BASE_DIR}")
    print("Acesse http://127.0.0.1:5000 no seu navegador.\n")
    app.run(debug=True, port=5000)
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
    app.run(debug=False, port=5000, use_reloader=False)