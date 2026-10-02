import unittest
import io
import json
from server import ConverterHandler

class MockSocket:
    def __init__(self, request_bytes):
        self._rfile = io.BytesIO(request_bytes)
        self._wfile = io.BytesIO()

    def makefile(self, mode, *args, **kwargs):
        if "r" in mode:
            return self._rfile
        return self._wfile

    def sendall(self, data):
        self._wfile.write(data)

class TestServerHandler(unittest.TestCase):
    def test_health_route(self):
        raw_req = b"GET /api/health HTTP/1.1\r\nHost: localhost\r\n\r\n"
        mock_sock = MockSocket(raw_req)
        handler = ConverterHandler(mock_sock, ("127.0.0.1", 12345), None)

        output = mock_sock._wfile.getvalue()
        self.assertIn(b"200 OK", output)
        self.assertIn(b"FileShift File Converter", output)

    def test_convert_route(self):
        body = b"Item,Qty\nNotebook,12\nPen,50\n"
        headers = (
            "POST /api/convert HTTP/1.1\r\n"
            "Host: localhost\r\n"
            "Content-Type: application/octet-stream\r\n"
            "X-Filename: stock.csv\r\n"
            "X-From-Format: csv\r\n"
            "X-To-Format: xlsx\r\n"
            f"Content-Length: {len(body)}\r\n\r\n"
        ).encode("utf-8")
        raw_req = headers + body
        mock_sock = MockSocket(raw_req)
        handler = ConverterHandler(mock_sock, ("127.0.0.1", 12345), None)

        output = mock_sock._wfile.getvalue()
        self.assertIn(b"200 OK", output)
        self.assertIn(b"stock.xlsx", output)

    def test_inspect_route(self):
        body = b"ColA,ColB\n1,2\n3,4\n"
        headers = (
            "POST /api/inspect HTTP/1.1\r\n"
            "Host: localhost\r\n"
            "Content-Type: application/octet-stream\r\n"
            "X-Filename: sample.csv\r\n"
            "X-From-Format: csv\r\n"
            f"Content-Length: {len(body)}\r\n\r\n"
        ).encode("utf-8")
        raw_req = headers + body
        mock_sock = MockSocket(raw_req)
        handler = ConverterHandler(mock_sock, ("127.0.0.1", 12345), None)

        output = mock_sock._wfile.getvalue()
        self.assertIn(b"200 OK", output)
        self.assertIn(b"table", output)

if __name__ == "__main__":
    unittest.main()
