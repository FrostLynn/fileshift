import io
import os
import json
import urllib.parse
from http import HTTPStatus
from http.server import HTTPServer, BaseHTTPRequestHandler
from socketserver import ThreadingMixIn
from converter.engine import convert_file, SUPPORTED_CONVERSIONS
from converter.inspect import inspect_file

STATIC_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "public")

class ThreadedHTTPServer(ThreadingMixIn, HTTPServer):
    daemon_threads = True

class ConverterHandler(BaseHTTPRequestHandler):
    server_version = "FileShift/1.0"

    def send_json_response(self, data: dict, status: int = 200):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self):
        self.send_response(HTTPStatus.NO_CONTENT)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, X-Filename, X-From-Format, X-To-Format, X-Delimiter, X-Sheet-Index")
        self.end_headers()

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        if path == "/api/health":
            supported = [{"from": s, "to": d} for s, d in SUPPORTED_CONVERSIONS.keys()]
            self.send_json_response({
                "status": "online",
                "app": "FileShift File Converter",
                "version": "1.0.0",
                "supported_conversions": supported
            })
            return

        # Serve static files
        if path == "/" or path == "":
            file_path = os.path.join(STATIC_DIR, "index.html")
        else:
            rel_path = path.lstrip("/")
            file_path = os.path.join(STATIC_DIR, rel_path)

        # Basic path traversal guard
        real_file_path = os.path.realpath(file_path)
        real_static_dir = os.path.realpath(STATIC_DIR)
        if not real_file_path.startswith(real_static_dir) or not os.path.isfile(real_file_path):
            self.send_response(404)
            self.send_header("Content-Type", "text/plain; charset=utf-8")
            self.end_headers()
            self.wfile.write(b"404 Not Found")
            return

        mime_types = {
            ".html": "text/html; charset=utf-8",
            ".css": "text/css; charset=utf-8",
            ".js": "application/javascript; charset=utf-8",
            ".json": "application/json; charset=utf-8",
            ".svg": "image/svg+xml",
            ".png": "image/png",
            ".ico": "image/x-icon"
        }
        ext = os.path.splitext(real_file_path)[1].lower()
        content_type = mime_types.get(ext, "application/octet-stream")

        with open(real_file_path, "rb") as f:
            content = f.read()

        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(content)))
        self.send_header("Cache-Control", "no-cache")
        self.end_headers()
        self.wfile.write(content)

    def parse_request_data(self) -> dict:
        content_length = int(self.headers.get("Content-Length", 0))
        if content_length <= 0:
            raise ValueError("No data provided in request body.")

        raw_body = self.rfile.read(content_length)
        content_type_header = self.headers.get("Content-Type", "")

        # 1. Direct raw binary payload with headers
        if "multipart/form-data" not in content_type_header:
            filename = urllib.parse.unquote(self.headers.get("X-Filename", "document"))
            from_fmt = self.headers.get("X-From-Format", "")
            to_fmt = self.headers.get("X-To-Format", "")
            delimiter = self.headers.get("X-Delimiter", "")
            sheet_index = self.headers.get("X-Sheet-Index", "0")
            return {
                "file_bytes": raw_body,
                "filename": filename,
                "from_fmt": from_fmt,
                "to_fmt": to_fmt,
                "delimiter": delimiter or None,
                "sheet_index": int(sheet_index) if sheet_index.isdigit() else 0
            }

        # 2. Parse multipart/form-data
        boundary_marker = None
        for part in content_type_header.split(";"):
            part = part.strip()
            if part.startswith("boundary="):
                boundary_marker = part.split("=", 1)[1].strip('"').encode("utf-8")
                break

        if not boundary_marker:
            raise ValueError("Invalid multipart/form-data boundary.")

        boundary = b"--" + boundary_marker
        parts = raw_body.split(boundary)

        file_bytes = None
        filename = "document"
        from_fmt = ""
        to_fmt = ""
        delimiter = None
        sheet_index = 0

        for part in parts:
            if not part or part == b"--\r\n" or part == b"--":
                continue
            if b"\r\n\r\n" in part:
                headers_data, body_data = part.split(b"\r\n\r\n", 1)
                # Strip trailing \r\n
                if body_data.endswith(b"\r\n"):
                    body_data = body_data[:-2]

                headers_text = headers_data.decode("latin-1", errors="replace")
                
                # Check field name
                name_match = None
                for line in headers_text.splitlines():
                    if "Content-Disposition:" in line:
                        for item in line.split(";"):
                            item = item.strip()
                            if item.startswith('name="'):
                                name_match = item[6:-1]
                            elif item.startswith('filename="'):
                                filename = item[10:-1]

                if name_match == "file":
                    file_bytes = body_data
                elif name_match == "from":
                    from_fmt = body_data.decode("utf-8", errors="replace").strip()
                elif name_match == "to":
                    to_fmt = body_data.decode("utf-8", errors="replace").strip()
                elif name_match == "delimiter":
                    delimiter = body_data.decode("utf-8", errors="replace").strip() or None
                elif name_match == "sheet_index":
                    s_str = body_data.decode("utf-8", errors="replace").strip()
                    sheet_index = int(s_str) if s_str.isdigit() else 0

        if file_bytes is None:
            raise ValueError("No file attached in multipart form.")

        return {
            "file_bytes": file_bytes,
            "filename": filename,
            "from_fmt": from_fmt,
            "to_fmt": to_fmt,
            "delimiter": delimiter,
            "sheet_index": sheet_index
        }

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        try:
            req_data = self.parse_request_data()
            file_bytes = req_data["file_bytes"]
            filename = req_data["filename"]
            from_fmt = req_data["from_fmt"]
            to_fmt = req_data["to_fmt"]

            # Infer from_fmt from filename extension if omitted
            if not from_fmt and "." in filename:
                from_fmt = filename.rsplit(".", 1)[-1].lower()

            if path == "/api/inspect":
                insp_result = inspect_file(file_bytes, filename, format_hint=from_fmt)
                self.send_json_response(insp_result, status=200)
                return

            elif path == "/api/convert":
                if not from_fmt or not to_fmt:
                    self.send_json_response({
                        "error": "Parameter format asal ('from') dan tujuan ('to') wajib diisi."
                    }, status=400)
                    return

                options = {
                    "delimiter": req_data.get("delimiter"),
                    "sheet_index": req_data.get("sheet_index", 0)
                }

                out_bytes, out_name, mime_type = convert_file(
                    file_bytes=file_bytes,
                    from_fmt=from_fmt,
                    to_fmt=to_fmt,
                    original_filename=filename,
                    options=options
                )

                self.send_response(200)
                self.send_header("Content-Type", mime_type)
                self.send_header("Content-Disposition", f'attachment; filename="{out_name}"')
                self.send_header("Content-Length", str(len(out_bytes)))
                self.send_header("X-Converted-Filename", out_name)
                self.send_header("X-Original-Size", str(len(file_bytes)))
                self.send_header("X-Converted-Size", str(len(out_bytes)))
                self.send_header("Access-Control-Allow-Origin", "*")
                self.send_header("Access-Control-Expose-Headers", "Content-Disposition, X-Converted-Filename, X-Original-Size, X-Converted-Size")
                self.end_headers()
                self.wfile.write(out_bytes)
                return

            else:
                self.send_json_response({"error": "Endpoint tidak ditemukan."}, status=404)

        except Exception as e:
            self.send_json_response({"error": str(e)}, status=400)

def run_server(port: int = 8080, host: str = "0.0.0.0"):
    server_address = (host, port)
    httpd = ThreadedHTTPServer(server_address, ConverterHandler)
    print(f"FileShift Converter Server running at http://{host}:{port}/")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nStopping server...")
        httpd.server_close()

if __name__ == "__main__":
    import sys
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 8080
    run_server(port=port)
