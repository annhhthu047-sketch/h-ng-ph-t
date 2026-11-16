import os
import json
import cgi
import http.server
import socketserver
import tempfile
import urllib.parse
from agent_engine import AgentDataEngine

PORT = 8501
STATIC_DIR = os.path.join(os.path.dirname(__file__), "static")
engine = AgentDataEngine()

# Nạp sẵn file mẫu mặc định khi khởi động
sample_path = os.path.join(os.path.dirname(__file__), "hung_phat_sample_data.xlsx")
if os.path.exists(sample_path):
    engine.load_from_excel_or_csv(sample_path)

class AgentRequestHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=STATIC_DIR, **kwargs)

    def do_POST(self):
        url_path = urllib.parse.urlparse(self.path).path
        
        if url_path == "/api/use-sample-data":
            global engine
            engine = AgentDataEngine()
            if os.path.exists(sample_path):
                engine.load_from_excel_or_csv(sample_path)
                self.send_json_response({"success": True, "message": "Đã nạp file mẫu thành công."})
            else:
                self.send_json_response({"success": False, "error": "Không tìm thấy file mẫu."}, status=404)

        elif url_path == "/api/run-pipeline":
            try:
                # Xử lý multipart form data nếu có upload file
                content_type = self.headers.get("Content-Type", "")
                
                if "multipart/form-data" in content_type:
                    form = cgi.FieldStorage(
                        fp=self.rfile,
                        headers=self.headers,
                        environ={'REQUEST_METHOD': 'POST',
                                 'CONTENT_TYPE': self.headers['Content-Type']}
                    )

                    file_item = form['file'] if 'file' in form else None
                    sheet_url_item = form.getvalue('sheet_url', None)

                    if file_item is not None and file_item.filename:
                        # Lưu tạm file để engine đọc
                        suffix = ".xlsx" if file_item.filename.endswith(".xlsx") else ".csv"
                        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as temp:
                            temp.write(file_item.file.read())
                            temp_path = temp.name
                        engine.load_from_excel_or_csv(temp_path)
                    elif sheet_url_item:
                        engine.load_from_google_sheet_url(sheet_url_item)
                
                # Thực hiện tuần tự 3 bước của AI Agent theo sơ đồ TO-BE
                # Bước 1: Tiền xử lý & Chuẩn hóa
                engine.run_step_1_normalization()

                # Bước 2: Đối chiếu & Phát hiện mâu thuẫn
                conflicts_res = engine.run_step_2_conflict_and_gap_analysis()

                # Bước 3: Đánh giá & Xếp hạng địa bàn
                eval_res = engine.run_step_3_evaluation_and_scoring()

                response_data = {
                    "success": True,
                    "audit_log": engine.audit_log,
                    "conflicts": conflicts_res,
                    "ranking": eval_res["ranking"],
                    "recommendations": eval_res["recommendations"]
                }
                self.send_json_response(response_data)

            except Exception as e:
                self.send_json_response({"success": False, "error": str(e)}, status=500)

        else:
            self.send_error(404, "Endpoint not found")

    def send_json_response(self, data, status=200):
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(json.dumps(data, ensure_ascii=False).encode("utf-8"))

def run_server():
    socketserver.TCPServer.allow_reuse_address = True
    with socketserver.TCPServer(("", PORT), AgentRequestHandler) as httpd:
        print(f"===============================================================")
        print(f"🤖 HƯNG PHÁT AI AGENT SERVER ĐANG CHẠY TẠI: http://localhost:{PORT}")
        print(f"===============================================================")
        httpd.serve_forever()

if __name__ == "__main__":
    run_server()

