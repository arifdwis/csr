import os
import sys
import json
import http.server
import socketserver
import webbrowser
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CODE_DIR = os.path.join(BASE_DIR, "code")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
sys.path.insert(0, CODE_DIR)

from data_loader import load_and_clean_data

PORT = 8501
DIRECTORY = os.path.dirname(os.path.abspath(__file__))

# Pre-load cleaned dataset
print("[1/3] Memproses dan memvalidasi dataset Excel CSR Samarinda...")
clean_json_path = os.path.join(OUTPUT_DIR, "data_cleaned.json")
if os.path.exists(clean_json_path):
    with open(clean_json_path, 'r', encoding='utf-8') as f:
        CLEANED_DATA = json.load(f)
else:
    df = load_and_clean_data()
    CLEANED_DATA = df.to_dict(orient='records')
print(f"[2/3] Berhasil memproses {len(CLEANED_DATA)} baris data terverifikasi.")

class CSRDashboardHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=DIRECTORY, **kwargs)

    def end_headers(self):
        self.send_header('Cache-Control', 'no-store, no-cache, must-revalidate, max-age=0')
        self.send_header('Pragma', 'no-cache')
        self.send_header('Expires', '0')
        super().end_headers()

    def do_GET(self):
        if self.path == '/api/data':
            self.send_response(200)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            clean_json_path = os.path.join(OUTPUT_DIR, "data_cleaned.json")
            if os.path.exists(clean_json_path):
                with open(clean_json_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
            else:
                data = CLEANED_DATA
            self.wfile.write(json.dumps(data, ensure_ascii=False).encode('utf-8'))
        elif self.path.startswith('/') and (self.path == '/' or self.path.startswith('/index.html') or self.path.startswith('/?')):
            self.path = '/index.html'
            return super().do_GET()
        else:
            return super().do_GET()

    def log_message(self, format, *args):
        # Clean logging
        return

def run_server(port=PORT):
    # Try preferred port, or fall back if busy
    for p in [port, 8080, 8000, 8888]:
        try:
            handler = CSRDashboardHandler
            socketserver.TCPServer.allow_reuse_address = True
            with socketserver.TCPServer(("", p), handler) as httpd:
                print("==================================================================")
                print(f"  🏛️  DASHBOARD EKSEKUTIF ANALISIS CSR PEMKOT SAMARINDA AKTIF")
                print(f"  🌐  Buka di Browser: http://localhost:{p}")
                print(f"  📊  Data Engine: 806 Entri Terstandarisasi (2022-2026)")
                print("==================================================================")
                
                # Auto-open browser
                try:
                    webbrowser.open(f"http://localhost:{p}")
                except Exception:
                    pass
                    
                httpd.serve_forever()
                break
        except OSError:
            print(f"Port {p} sedang digunakan, mencoba port berikutnya...")

if __name__ == '__main__':
    run_server()
