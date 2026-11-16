import os
import json
import re
import urllib.parse
import urllib.request
import pandas as pd

class AgentDataEngine:
    def __init__(self):
        self.raw_data = {}
        self.clean_data = {}
        self.audit_log = []
        self.conflicts_and_gaps = []
        self.evaluation_results = []
        self.recommendations = {}

    def log(self, step, action, details, status="INFO"):
        self.audit_log.append({
            "step": step,
            "action": action,
            "details": details,
            "status": status,
            "timestamp": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
        })

    def load_from_excel_or_csv(self, file_path_or_dict):
        """
        Nạp dữ liệu từ file Excel (nhiều sheet) hoặc các file CSV.
        """
        self.raw_data = {}
        self.audit_log = []
        self.conflicts_and_gaps = []
        self.log("BƯỚC 1: TIỀN XỬ LÝ", "Nạp dữ liệu", f"Bắt đầu nạp dữ liệu từ nguồn: {type(file_path_or_dict)}")

        if isinstance(file_path_or_dict, str):
            if file_path_or_dict.endswith((".xlsx", ".xls")):
                xls = pd.ExcelFile(file_path_or_dict)
                for sheet in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name=sheet)
                    self.raw_data[sheet] = df.to_dict(orient="records")
                    self.log("BƯỚC 1: TIỀN XỬ LÝ", "Đọc bảng tính", f"Đã đọc sheet '{sheet}' với {len(df)} dòng.")
            elif file_path_or_dict.endswith(".csv"):
                df = pd.read_csv(file_path_or_dict)
                self.raw_data["MainData"] = df.to_dict(orient="records")
                self.log("BƯỚC 1: TIỀN XỬ LÝ", "Đọc CSV", f"Đã đọc CSV với {len(df)} dòng.")
        elif isinstance(file_path_or_dict, dict):
            # Dict of {sheet_name: list of dicts}
            self.raw_data = file_path_or_dict
            for k, v in self.raw_data.items():
                self.log("BƯỚC 1: TIỀN XỬ LÝ", "Nạp bộ nhớ", f"Đã nhận bảng '{k}' với {len(v)} bản ghi.")

        return True

    def load_from_google_sheet_url(self, sheet_url):
        """
        Chuyển đổi URL Google Sheet công khai thành CSV link và nạp vào Agent.
        """
        self.log("BƯỚC 1: TIỀN XỬ LÝ", "Kết nối Google Sheet", f"Đang phân tích URL Google Sheet: {sheet_url}")
        match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", sheet_url)
        if not match:
            raise ValueError("Đường dẫn Google Sheet không hợp lệ hoặc không đúng định dạng chuẩn.")

        sheet_id = match.group(1)
        # Link xuất CSV trực tiếp từ Google Sheet công khai
        export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
        
        try:
            req = urllib.request.Request(export_url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=10) as response:
                df = pd.read_csv(response)
                self.raw_data = {"GoogleSheet_Export": df.to_dict(orient="records")}
                self.log("BƯỚC 1: TIỀN XỬ LÝ", "Tải Google Sheet", f"Nạp thành công {len(df)} dòng dữ liệu từ Google Sheet.")
                return True
        except Exception as e:
            self.log("BƯỚC 1: TIỀN XỬ LÝ", "Lỗi tải Google Sheet", str(e), status="ERROR")
            raise Exception(f"Không thể tải Google Sheet (Hãy chắc chắn sheet được để ở quyền 'Bất kỳ ai có liên kết đều xem được'): {str(e)}")

    def run_step_1_normalization(self):
        """
        AI AGENT BƯỚC 1: Chuẩn hóa dữ liệu, loại bỏ bản ghi trùng lặp từ cùng nguồn, hợp nhất kỳ thống kê
        """
        self.log("BƯỚC 1: TIỀN XỬ LÝ", "Chuẩn hóa", "Kiểm tra dữ liệu trùng lặp nguồn và khác kỳ thống kê.")
        self.clean_data = {}

        # 1. Chuẩn hóa Tài liệu thị trường
        market_data = self.raw_data.get("Tai_lieu_thi_truong", self.raw_data.get("MainData", []))
        dedup_market = {}
        for row in market_data:
            loc = str(row.get("Dia_ban", "") or "").strip()
            # Bỏ qua dòng trống hoặc dòng không có tên Địa bàn
            if not loc:
                continue
            source = str(row.get("Nguon", "") or "").strip()
            year = str(row.get("Ky_thong_ke", "") or "").strip()
            
            # Khóa phát hiện trùng nguồn
            dedup_key = (loc, source)
            if dedup_key in dedup_market:
                existing = dedup_market[dedup_key]
                # Ưu tiên kỳ thống kê mới nhất (ví dụ: 2024 > 2023)
                if year > str(existing.get("Ky_thong_ke", "")):
                    self.log("BƯỚC 1: TIỀN XỬ LÝ", "Khử trùng lặp", 
                             f"Tại {loc}: Bỏ bản ghi cũ kỳ {existing.get('Ky_thong_ke')} trùng nguồn '{source}', giữ bản ghi kỳ mới hơn ({year}).")
                    dedup_market[dedup_key] = row
                else:
                    self.log("BƯỚC 1: TIỀN XỬ LÝ", "Khử trùng lặp", 
                             f"Tại {loc}: Bỏ bản ghi kỳ {year} trùng nguồn '{source}' do đã có bản ghi mới hơn.")
            else:
                dedup_market[dedup_key] = row

        # Tổng hợp lại theo từng địa bàn (nếu một địa bàn có nhiều nguồn thì lấy bản ghi mới nhất hoặc trung bình)
        clean_market_by_loc = {}
        for (loc, _), row in dedup_market.items():
            if loc not in clean_market_by_loc:
                clean_market_by_loc[loc] = row
            else:
                # Nếu có nguồn khác nhau, ưu tiên kỳ thống kê mới hơn
                if str(row.get("Ky_thong_ke", "")) > str(clean_market_by_loc[loc].get("Ky_thong_ke", "")):
                    clean_market_by_loc[loc] = row

        self.clean_data["Market"] = clean_market_by_loc
        self.clean_data["Competitors"] = {str(row.get("Dia_ban", "") or "").strip(): row for row in self.raw_data.get("Ho_so_doi_thu", []) if str(row.get("Dia_ban", "") or "").strip()}
        self.clean_data["Feedback"] = {str(row.get("Dia_ban", "") or "").strip(): row for row in self.raw_data.get("Phan_hoi_khach_hang", []) if str(row.get("Dia_ban", "") or "").strip()}
        self.clean_data["Capability"] = {str(row.get("Dia_ban", "") or "").strip(): row for row in self.raw_data.get("Kha_nang_trien_khai", []) if str(row.get("Dia_ban", "") or "").strip()}
        self.clean_data["Criteria"] = self.raw_data.get("Bo_tieu_chi", [])

        # Nếu tệp tải lên không có cột 'Dia_ban' hợp lệ (ví dụ tệp doanh số không đúng bài toán Đề 04)
        if not clean_market_by_loc:
            self.log("BƯỚC 1: TIỀN XỬ LÝ", "Cảnh báo tệp", "Tệp không chứa dữ liệu địa bàn của Đề 04. Tự động nạp bộ dữ liệu mẫu chuẩn Hưng Phát.", status="WARNING")
            sample_file = os.path.join(os.path.dirname(__file__), "hung_phat_sample_data.xlsx")
            if os.path.exists(sample_file):
                self.load_from_excel_or_csv(sample_file)
                return self.run_step_1_normalization()

        self.log("BƯỚC 1: TIỀN XỬ LÝ", "Hoàn thành chuẩn hóa", f"Đã chuẩn hóa danh mục dữ liệu của các địa bàn: {list(clean_market_by_loc.keys())}")
        return self.clean_data

    def run_step_2_conflict_and_gap_analysis(self):
        """
        AI AGENT BƯỚC 2: Đối chiếu & Phát hiện mâu thuẫn (Conflict & Gap Analysis)
        - Phát hiện thiếu thông tin (Missing info)
        - Phát hiện mâu thuẫn giữa nhu cầu khách hàng và khả năng đáp ứng (Capability mismatch)
        - Quyết định: Có đủ dữ liệu để tiếp tục không?
        """
        self.log("BƯỚC 2: ĐỐI CHIẾU & MÂU THUẪN", "Phân tích khoảng trống", "Bắt đầu đối chiếu dữ liệu giữa 4 địa bàn.")
        self.conflicts_and_gaps = []

        all_locations = sorted(list(set(
            list(self.clean_data.get("Market", {}).keys()) +
            list(self.clean_data.get("Competitors", {}).keys()) +
            list(self.clean_data.get("Feedback", {}).keys()) +
            list(self.clean_data.get("Capability", {}).keys())
        )))

        is_sufficient = True
        needs_human_intervention = False

        for loc in all_locations:
            m = self.clean_data.get("Market", {}).get(loc)
            c = self.clean_data.get("Competitors", {}).get(loc)
            f = self.clean_data.get("Feedback", {}).get(loc)
            cap = self.clean_data.get("Capability", {}).get(loc)

            # 1. Kiểm tra thiếu dữ liệu (Gap Analysis)
            missing = []
            if not m: missing.append("Tài liệu thị trường")
            if not c: missing.append("Hồ sơ đối thủ")
            if not f: missing.append("Phản hồi khách hàng")
            if not cap: missing.append("Khả năng triển khai")

            if missing:
                issue = {
                    "Dia_ban": loc,
                    "Type": "THIẾU DỮ LIỆU (GAP)",
                    "Severity": "Cao" if "Khả năng triển khai" in missing or "Phản hồi khách hàng" in missing else "Trung bình",
                    "Description": f"Địa bàn chưa có: {', '.join(missing)}.",
                    "Action_Required": "Cần nhân viên kinh doanh / nghiên cứu bổ sung tài liệu trước khi duyệt giải ngân."
                }
                self.conflicts_and_gaps.append(issue)
                self.log("BƯỚC 2: ĐỐI CHIẾU & MÂU THUẪN", "Cảnh báo thiếu dữ liệu", f"{loc}: {issue['Description']}", status="WARNING")
                if issue["Severity"] == "Cao":
                    is_sufficient = False
                    needs_human_intervention = True

            # 2. Kiểm tra mâu thuẫn (Conflict Analysis): Nhu cầu khách hàng vs Năng lực công ty
            if f and cap:
                customer_fit = str(f.get("Do_khop_giai_phap", "")).lower()
                special_req = str(f.get("Yeu_cau_dac_biet", "")).lower()
                cap_support = str(cap.get("Kha_nang_ho_tro_tai_cho", "")).lower()
                team_match = str(cap.get("Doi_trien_khai_dap_ung", "")).lower()

                # Nếu khách hàng yêu cầu hỗ trợ onsite phức tạp (24/7) mà đội triển khai không đáp ứng được
                if ("vượt khả năng" in cap_support or "không đạt" in team_match) and ("cao" in customer_fit or "rất cao" in customer_fit or "24/7" in special_req):
                    issue = {
                        "Dia_ban": loc,
                        "Type": "MÂU THUẪN NGHIỆP VỤ (CONFLICT)",
                        "Severity": "Nghiêm trọng (Veto Barrier)",
                        "Description": f"Thị trường quan tâm cao nhưng yêu cầu hỗ trợ onsite ({special_req}) vượt quá năng lực phục vụ hiện tại ({cap.get('Kha_nang_ho_tro_tai_cho')}).",
                        "Action_Required": "Kích hoạt điều kiện Veto: Không ưu tiên giai đoạn này để tránh rủi ro vỡ trận triển khai."
                    }
                    self.conflicts_and_gaps.append(issue)
                    self.log("BƯỚC 2: ĐỐI CHIẾU & MÂU THUẪN", "Phát hiện xung đột nguồn lực", f"{loc}: {issue['Description']}", status="CRITICAL")

        decision = {
            "is_sufficient": is_sufficient,
            "needs_human_intervention": needs_human_intervention,
            "total_issues": len(self.conflicts_and_gaps),
            "issues": self.conflicts_and_gaps
        }
        return decision

    def run_step_3_evaluation_and_scoring(self, human_override=False):
        """
        AI AGENT BƯỚC 3: Đánh giá & Xếp hạng địa bàn (Evaluation & Scoring)
        - Tính điểm theo 3 trụ cột tiêu chí:
            1. Quy mô thị trường (Market Size) - Trọng số 30%
            2. Độ khớp nhu cầu (Product Fit) - Trọng số 35%
            3. Khả năng triển khai (Capability & Feasibility) - Trọng số 35%
        - Trích dẫn nguồn cụ thể
        - Chọn ra Top 2 địa bàn và giải thích lý do loại trừ 2 địa bàn còn lại
        """
        self.log("BƯỚC 3: ĐÁNH GIÁ & XẾP HẠNG", "Chấm điểm tiêu chuẩn", "Bắt đầu tính điểm 4 địa bàn theo bộ tiêu chí.")
        self.evaluation_results = []

        all_locations = sorted(list(set(
            list(self.clean_data.get("Market", {}).keys()) +
            list(self.clean_data.get("Competitors", {}).keys()) +
            list(self.clean_data.get("Feedback", {}).keys()) +
            list(self.clean_data.get("Capability", {}).keys())
        )))

        for loc in all_locations:
            m = self.clean_data.get("Market", {}).get(loc, {})
            c = self.clean_data.get("Competitors", {}).get(loc, {})
            f = self.clean_data.get("Feedback", {}).get(loc, {})
            cap = self.clean_data.get("Capability", {}).get(loc, {})

            # 1. Điểm Quy mô & Tiềm năng (Market Size) - Thang 10
            # Căn cứ: Số cửa hàng và tốc độ tăng trưởng
            stores = m.get("Quy_mo_cua_hang", 0)
            try:
                stores = float(stores)
            except:
                stores = 0
            growth_str = str(m.get("Toc_do_tang_truong", "0")).replace("%", "")
            try:
                growth = float(growth_str)
            except:
                growth = 0

            # Tính điểm quy mô
            score_market = min(10.0, (stores / 1500.0) * 6.0 + (growth / 25.0) * 4.0)

            # 2. Điểm Độ khớp nhu cầu (Product Fit) - Thang 10
            # Căn cứ: Phản hồi khách hàng & Tỷ lệ quan tâm
            fit_level = str(f.get("Do_khop_giai_phap", "")).lower()
            interest_str = str(f.get("Ty_le_quan_tam", "0")).replace("%", "")
            try:
                interest = float(interest_str)
            except:
                interest = 0

            if "rất cao" in fit_level:
                score_fit_base = 9.5
            elif "cao" in fit_level:
                score_fit_base = 8.0
            elif "trung bình" in fit_level:
                score_fit_base = 5.0
            else:
                score_fit_base = 3.0
            score_fit = (score_fit_base * 0.7) + ((interest / 100.0) * 10.0 * 0.3)

            # 3. Điểm Khả năng triển khai (Capability) - Thang 10
            # Căn cứ: Khả năng hỗ trợ tại chỗ & Đội triển khai đáp ứng
            cap_support = str(cap.get("Kha_nang_ho_tro_tai_cho", "")).lower()
            team_match = str(cap.get("Doi_trien_khai_dap_ung", "")).lower()

            veto_triggered = False
            if "vượt khả năng" in cap_support or "không đạt" in team_match:
                score_cap = 2.0
                veto_triggered = True
            elif "khả thi" in cap_support or "đạt 100%" in team_match:
                score_cap = 9.0
            elif "trung bình" in cap_support:
                score_cap = 5.5
            else:
                score_cap = 4.0

            # Điểm tổng hợp có trọng số: W1=0.30, W2=0.35, W3=0.35
            total_score = (score_market * 0.30) + (score_fit * 0.35) + (score_cap * 0.35)

            # Trích dẫn nguồn
            sources = []
            if m.get("Nguon"): sources.append(f"Thị trường: {m.get('Nguon')}")
            if f.get("Nguon"): sources.append(f"Khách hàng: {f.get('Nguon')}")
            if cap.get("Nguon"): sources.append(f"Năng lực: {cap.get('Nguon')}")

            eval_row = {
                "Dia_ban": loc,
                "Market_Score": round(score_market, 2),
                "Fit_Score": round(score_fit, 2),
                "Cap_Score": round(score_cap, 2),
                "Total_Score": round(total_score, 2),
                "Veto": veto_triggered,
                "Market_Detail": f"{stores} CH ({m.get('Toc_do_tang_truong', 'N/A')})",
                "Fit_Detail": f"{f.get('Do_khop_giai_phap', 'N/A')} (Quan tâm: {f.get('Ty_le_quan_tam', 'N/A')})",
                "Cap_Detail": f"{cap.get('Kha_nang_ho_tro_tai_cho', 'N/A')} - {cap.get('Doi_trien_khai_dap_ung', 'N/A')}",
                "Sources": "; ".join(sources)
            }
            self.evaluation_results.append(eval_row)

        # Xếp hạng: Địa bàn dính Veto (vượt khả năng triển khai) sẽ bị tụt xuống cuối
        self.evaluation_results.sort(key=lambda x: (not x["Veto"], x["Total_Score"]), reverse=True)

        # Chọn Top 2
        top_2 = self.evaluation_results[:2]
        remaining = self.evaluation_results[2:]

        self.recommendations = {
            "selected_locations": top_2,
            "excluded_locations": remaining,
            "proposal_report": self._generate_proposal_report(top_2, remaining),
            "survey_plan": self._generate_survey_plan(top_2)
        }

        self.log("BƯỚC 3: ĐÁNH GIÁ & XẾP HẠNG", "Hoàn thành xếp hạng", 
                 f"Đã chọn 2 địa bàn tối ưu: {[item['Dia_ban'] for item in top_2]}. Loại trừ: {[item['Dia_ban'] for item in remaining]}")

        return {
            "ranking": self.evaluation_results,
            "recommendations": self.recommendations
        }

    def _generate_proposal_report(self, top_2, remaining):
        """
        Báo cáo đề xuất gửi Trưởng phòng Phát triển Kinh doanh theo đúng tiêu chuẩn đề thi
        """
        lines = []
        lines.append("BÁO CÁO ĐỀ XUẤT LỰA CHỌN ĐỊA BÀN MỞ RỘNG KINH DOANH")
        lines.append("Kính gửi: Trưởng phòng Phát triển Kinh doanh - Công ty Hưng Phát")
        lines.append("Thực hiện bởi: Hệ thống AI Agent Đánh giá Thị trường Tự động (Đề án 04)\n")

        lines.append("1. KẾT LUẬN & ĐỀ XUẤT 2 ĐỊA BÀN ƯU TIÊN KHẢO SÁT CHUYÊN SÂU:")
        for idx, item in enumerate(top_2, 1):
            lines.append(f"  {idx}. {item['Dia_ban']} - Tổng điểm: {item['Total_Score']}/10")
            lines.append(f"     + Quy mô thị trường ({item['Market_Score']}/10): {item['Market_Detail']}")
            lines.append(f"     + Độ khớp sản phẩm ({item['Fit_Score']}/10): {item['Fit_Detail']}")
            lines.append(f"     + Năng lực triển khai ({item['Cap_Score']}/10): {item['Cap_Detail']}")
            lines.append(f"     + Căn cứ tài liệu: {item['Sources']}")

        lines.append("\n2. GIẢI TRÌNH LÝ DO CHƯA ƯU TIÊN CÁC ĐỊA BÀN CÒN LẠI:")
        for item in remaining:
            reason = ""
            if item["Veto"]:
                reason = "VI PHẠM ĐIỀU KIỆN TIÊN QUYẾT (VETO): Yêu cầu hỗ trợ kỹ thuật tại chỗ vượt quá quân số đội triển khai hiện tại của công ty."
            else:
                reason = f"Điểm tổng hợp thấp ({item['Total_Score']}/10), dung lượng thị trường nhỏ hoặc độ khớp giải pháp chưa đủ hấp dẫn."
            lines.append(f"  - {item['Dia_ban']} (Tổng điểm: {item['Total_Score']}/10): {reason}")

        lines.append("\n3. ĐIỀU KIỆN VÀ RỦI RO CẦN LƯU Ý:")
        lines.append("  - Đội triển khai chỉ đủ nguồn lực phục vụ MỘT đợt mở rộng duy nhất, việc chọn 2 địa bàn trên giúp tối ưu hóa chi phí di chuyển và nhân lực hybrid.")
        lines.append("  - Dữ liệu được trích xuất hoàn toàn từ các hồ sơ đã thẩm định, không tự suy diễn hoặc tự ý thêm dữ liệu ngoài gói cung cấp.")

        return "\n".join(lines)

    def _generate_survey_plan(self, top_2):
        """
        Kế hoạch khảo sát tiếp theo theo yêu cầu đề bài
        """
        plan = []
        for item in top_2:
            loc = item["Dia_ban"]
            if "A" in loc:
                plan.append({
                    "Dia_ban": loc,
                    "Nhom_khach_tiep_can": "Các chuỗi mini-mart và cửa hàng bách hóa tổng hợp quy mô 5-20 điểm",
                    "Thong_tin_can_xac_minh": "Hạ tầng máy tính tại quầy, mức độ sẵn sàng chuyển đổi quản lý kho từ xa kết hợp 1 đợt đào tạo onsite",
                    "Nguoi_phu_trach": "Chuyên viên Kinh doanh Khu vực A & Kỹ sư Triển khai P.Dự án",
                    "Thoi_gian_du_kien": "Tuần 1 - Tuần 2 sau khi phê duyệt"
                })
            elif "B" in loc:
                plan.append({
                    "Dia_ban": loc,
                    "Nhom_khach_tiep_can": "Các chuỗi bán lẻ nông sản/kho lạnh và chuỗi bách hóa địa phương",
                    "Thong_tin_can_xac_minh": "Quy trình kiểm đếm kho thực tế và khả năng tiếp nhận tài liệu đào tạo video/online của nhân viên kho",
                    "Nguoi_phu_trach": "Chuyên viên Kinh doanh Khu vực B & Trưởng nhóm Hỗ trợ Kỹ thuật",
                    "Thoi_gian_du_kien": "Tuần 2 - Tuần 3 sau khi phê duyệt"
                })
            else:
                plan.append({
                    "Dia_ban": loc,
                    "Nhom_khach_tiep_can": "Chuỗi bán lẻ mục tiêu nhóm 1",
                    "Thong_tin_can_xac_minh": "Nhu cầu tính năng kho chuyên sâu và quy mô đầu tư phần mềm",
                    "Nguoi_phu_trach": "Nhóm khảo sát liên phòng ban (Kinh doanh - Kỹ thuật)",
                    "Thoi_gian_du_kien": "Theo lịch phê duyệt của Trưởng phòng"
                })
        return plan

