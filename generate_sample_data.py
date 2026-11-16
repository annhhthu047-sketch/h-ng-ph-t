import os
import sys
import pandas as pd

def create_sample_excel():
    excel_path = os.path.join(os.getcwd(), "hung_phat_sample_data.xlsx")
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        # Sheet 1: Tai_lieu_thi_truong
        df_market = pd.DataFrame([
            {"Dia_ban": "Địa bàn A", "Quy_mo_cua_hang": 1200, "Toc_do_tang_truong": "15%", "Ky_thong_ke": "2024", "Nguon": "Báo cáo Hiệp hội Bán lẻ 2024", "Ghi_chu": "Dung lượng thị trường lớn, tập trung chuỗi mini-mart"},
            {"Dia_ban": "Địa bàn A", "Quy_mo_cua_hang": 1150, "Toc_do_tang_truong": "14%", "Ky_thong_ke": "Q3/2023", "Nguon": "Báo cáo Hiệp hội Bán lẻ 2024", "Ghi_chu": "Số liệu cũ trùng nguồn khảo sát"},
            {"Dia_ban": "Địa bàn B", "Quy_mo_cua_hang": 850, "Toc_do_tang_truong": "22%", "Ky_thong_ke": "2024", "Nguon": "Sở Công Thương Địa bàn B", "Ghi_chu": "Thị trường mới nổi, tốc độ mở rộng chuỗi bán lẻ rất nhanh"},
            {"Dia_ban": "Địa bàn C", "Quy_mo_cua_hang": 1500, "Toc_do_tang_truong": "8%", "Ky_thong_ke": "2024", "Nguon": "Cục Thống kê Địa bàn C", "Ghi_chu": "Thị trường lớn nhất nhưng độ bão hòa cao, chuỗi lớn đã có ERP riêng"},
            {"Dia_ban": "Địa bàn D", "Quy_mo_cua_hang": 400, "Toc_do_tang_truong": "5%", "Ky_thong_ke": "2023", "Nguon": "Báo cáo nội bộ P.KD", "Ghi_chu": "Quy mô nhỏ, số cửa hàng bán lẻ phân tán"}
        ])
        df_market.to_excel(writer, sheet_name="Tai_lieu_thi_truong", index=False)

        # Sheet 2: Ho_so_doi_thu
        df_competitors = pd.DataFrame([
            {"Dia_ban": "Địa bàn A", "So_doi_thu_chinh": 3, "Doi_thu_tieu_bieu": "KhoViet, SmartLog", "Muc_do_canh_tranh": "Trung bình", "Diem_yeu_doi_thu": "Giá cao, dịch vụ hỗ trợ chậm trễ", "Nguon": "Khảo sát thực địa Q1/2024"},
            {"Dia_ban": "Địa bàn B", "So_doi_thu_chinh": 1, "Doi_thu_tieu_bieu": "Phần mềm địa phương X", "Muc_do_canh_tranh": "Thấp", "Diem_yeu_doi_thu": "Chỉ quản lý bán hàng, thiếu tính năng kho nâng cao (WMS)", "Nguon": "Báo cáo đối thủ P.Marketing 2024"},
            {"Dia_ban": "Địa bàn C", "So_doi_thu_chinh": 5, "Doi_thu_tieu_bieu": "SAP, Oracle, Fast, KiotViet", "Muc_do_canh_tranh": "Rất cao", "Diem_yeu_doi_thu": "Thị phần bị chia nhỏ, cạnh tranh gay gắt về chiết khấu", "Nguon": "Báo cáo đối thủ 2024"},
            {"Dia_ban": "Địa bàn D", "So_doi_thu_chinh": 2, "Doi_thu_tieu_bieu": "EasyPos, Sapo", "Muc_do_canh_tranh": "Thấp", "Diem_yeu_doi_thu": "Chưa có WMS chuyên sâu", "Nguon": "Khảo sát online 2023"}
        ])
        df_competitors.to_excel(writer, sheet_name="Ho_so_doi_thu", index=False)

        # Sheet 3: Phan_hoi_khach_hang
        df_feedback = pd.DataFrame([
            {"Dia_ban": "Địa bàn A", "So_khach_khao_sat": 35, "Ty_le_quan_tam": "78%", "Nhu_cau_chinh": "Quản lý tồn kho đa điểm, đồng bộ bán lẻ", "Do_khop_giai_phap": "Cao", "Yeu_cau_dac_biet": "Cần hỗ trợ onboarding từ xa và 1 đợt trực tiếp ban đầu", "Nguon": "Phỏng vấn KH mục tiêu 03/2024"},
            {"Dia_ban": "Địa bàn B", "So_khach_khao_sat": 28, "Ty_le_quan_tam": "85%", "Nhu_cau_chinh": "Chuyển đổi số quản lý kho lạnh & bách hóa bán lẻ", "Do_khop_giai_phap": "Rất cao", "Yeu_cau_dac_biet": "Đào tạo online kết hợp tài liệu video chuẩn hóa", "Nguon": "Hội thảo bán lẻ tỉnh B 04/2024"},
            {"Dia_ban": "Địa bàn C", "So_khach_khao_sat": 40, "Ty_le_quan_tam": "80%", "Nhu_cau_chinh": "Tích hợp robot tự động và kỹ thuật viên trực tại chỗ 24/7", "Do_khop_giai_phap": "Thấp", "Yeu_cau_dac_biet": "Bắt buộc có đội kỹ sư onsite 24/7 tại kho trung tâm", "Nguon": "Khảo sát nhu cầu Q2/2024"},
            {"Dia_ban": "Địa bàn D", "So_khach_khao_sat": 15, "Ty_le_quan_tam": "40%", "Nhu_cau_chinh": "Chỉ muốn sổ sách excel đơn giản", "Do_khop_giai_phap": "Thấp", "Yeu_cau_dac_biet": "Ngân sách rất thấp", "Nguon": "Phỏng vấn điện thoại 2023"}
        ])
        df_feedback.to_excel(writer, sheet_name="Phan_hoi_khach_hang", index=False)

        # Sheet 4: Kha_nang_trien_khai
        df_capability = pd.DataFrame([
            {"Dia_ban": "Địa bàn A", "Khoang_cach_tru_so": "80 km", "Kha_nang_ho_tro_tai_cho": "Khả thi", "Nhan_luc_trien_khai_can": "3 chuyên viên", "Doi_trien_khai_dap_ung": "Đạt 100% (trong bán kính hỗ trợ nhanh)", "Nguon": "Kế hoạch nhân sự Đội Triển khai 2024"},
            {"Dia_ban": "Địa bàn B", "Khoang_cach_tru_so": "120 km", "Kha_nang_ho_tro_tai_cho": "Khả thi", "Nhan_luc_trien_khai_can": "2 chuyên viên", "Doi_trien_khai_dap_ung": "Đạt 100% (mô hình hybrid từ xa + onsite)", "Nguon": "Kế hoạch nhân sự Đội Triển khai 2024"},
            {"Dia_ban": "Địa bàn C", "Khoang_cach_tru_so": "500 km", "Kha_nang_ho_tro_tai_cho": "Vượt khả năng", "Nhan_luc_trien_khai_can": "8 chuyên viên onsite 24/7", "Doi_trien_khai_dap_ung": "Không đạt (vượt quá quân số đội triển khai hiện tại)", "Nguon": "Đánh giá khả năng đáp ứng kỹ thuật Q2/2024"},
            {"Dia_ban": "Địa bàn D", "Khoang_cach_tru_so": "350 km", "Kha_nang_ho_tro_tai_cho": "Trung bình", "Nhan_luc_trien_khai_can": "2 chuyên viên", "Doi_trien_khai_dap_ung": "Có thể thu xếp nhưng hiệu quả doanh thu thấp", "Nguon": "Kế hoạch nhân sự Đội Triển khai 2024"}
        ])
        df_capability.to_excel(writer, sheet_name="Kha_nang_trien_khai", index=False)

        # Sheet 5: Bo_tieu_chi
        df_criteria = pd.DataFrame([
            {"Tieu_chi": "Quy mô & Tiềm năng thị trường", "Ma_tieu_chi": "Market_Size", "Trong_so": 0.30, "Mo_ta": "Số lượng cửa hàng chuỗi bán lẻ và tốc độ tăng trưởng hàng năm", "Dieu_kien_tien_quyet": "Kỳ thống kê 2024"},
            {"Tieu_chi": "Độ khớp nhu cầu khách hàng với giải pháp", "Ma_tieu_chi": "Product_Fit", "Trong_so": 0.35, "Mo_ta": "Tính năng kho WMS phù hợp nhu cầu, khách hàng sẵn sàng ứng dụng", "Dieu_kien_tien_quyet": "Tỷ lệ quan tâm > 60%"},
            {"Tieu_chi": "Khả năng triển khai & Hỗ trợ kỹ thuật", "Ma_tieu_chi": "Feasibility_Capability", "Trong_so": 0.35, "Mo_ta": "Đáp ứng nguồn lực của công ty Hưng Phát, không vượt quá giới hạn hỗ trợ tại chỗ", "Dieu_kien_tien_quyet": "Không được vượt khả năng (Veto criterion)"}
        ])
        df_criteria.to_excel(writer, sheet_name="Bo_tieu_chi", index=False)

    print(f"Sample data created successfully at {excel_path}")

if __name__ == "__main__":
    create_sample_excel()

