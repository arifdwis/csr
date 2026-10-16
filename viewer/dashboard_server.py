"""Server lokal untuk prototipe dashboard CSR/TSP Samarinda.

- /api/data  : baris hasil normalisasi (output/data_cleaned.json)
- /api/rekap : rekonsiliasi per lembar terhadap subtotal dan sheet Total Excel (output/rekap_sumber.json)
Data diolah ulang otomatis bila file Excel sumber lebih baru daripada hasil olahan.
Server hanya mendengarkan di 127.0.0.1 karena data memuat nama pemberi perorangan.
"""
import http.server
import json
import os
import socketserver
import sys
import webbrowser

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CODE_DIR = os.path.join(BASE_DIR, "code")
sys.path.insert(0, CODE_DIR)

from data_loader import EXCEL_FILE, OUTPUT_JSON, OUTPUT_REKAP_JSON, load_and_clean_data  # noqa: E402

HOST = os.environ.get("CSR_HOST", "127.0.0.1")
PORT = int(os.environ.get("CSR_PORT", "8501"))
DIRECTORY = os.path.dirname(os.path.abspath(__file__))


def ensure_fresh_output():
    """Olah ulang Excel bila hasil olahan belum ada atau lebih lama dari sumbernya."""
    outputs = [OUTPUT_JSON, OUTPUT_REKAP_JSON]
    if not all(os.path.exists(p) for p in outputs) or \
            any(os.path.getmtime(p) < os.path.getmtime(EXCEL_FILE) for p in outputs):
        print("Memproses ulang Excel sumber...")
        load_and_clean_data()


def read_json(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


class CSRDashboardHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=DIRECTORY, **kwargs)

    def handle(self):
        try:
            super().handle()
        except (BrokenPipeError, ConnectionResetError):
            pass

    def end_headers(self):
        self.send_header("Cache-Control", "no-store, max-age=0")
        super().end_headers()

    def send_json(self, payload):
        try:
            body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        except (BrokenPipeError, ConnectionResetError):
            pass

    def do_GET(self):
        path = self.path.split("?", 1)[0]
        if path in ("/api/data", "/api/rekap"):
            ensure_fresh_output()
            return self.send_json(read_json(OUTPUT_JSON if path == "/api/data" else OUTPUT_REKAP_JSON))
        if path in ("/", "/index.html"):
            self.path = "/index.html"
            try:
                return super().do_GET()
            except (BrokenPipeError, ConnectionResetError):
                pass
            return
        # Hanya index.html yang disajikan dari folder viewer.
        self.send_error(404)

    def log_message(self, format, *args):
        return


def run_server():
    ensure_fresh_output()
    rows = read_json(OUTPUT_JSON)
    rekap = read_json(OUTPUT_REKAP_JSON)
    socketserver.TCPServer.allow_reuse_address = True
    for p in [PORT, 8080, 8000, 8888]:
        try:
            with socketserver.TCPServer((HOST, p), CSRDashboardHandler) as httpd:
                print("=" * 66)
                print("  Prototipe Dashboard CSR/TSP Pemkot Samarinda")
                print(f"  Buka di browser : http://localhost:{p}")
                print(f"  Data            : {len(rows)} baris olahan, {rekap['jumlah_baris_dalam_rekap']} masuk rekap")
                print(f"  Total rekap     : Rp {rekap['total_dalam_rekap']:,.0f}".replace(",", "."))
                print("=" * 66)
                try:
                    webbrowser.open(f"http://localhost:{p}")
                except Exception:
                    pass
                httpd.serve_forever()
                return
        except OSError:
            print(f"Port {p} sedang digunakan, mencoba port berikutnya...")


if __name__ == "__main__":
    run_server()
