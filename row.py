import pandas as pd
from datetime import datetime, timedelta

def create_sample_excel():
    # Tạo dữ liệu mẫu
    data = {
        'STT': [1, 2],
        'Tên Nghiên cứu sinh': ['Nguyễn Văn A', 'Trần Thị B'],
        'Mã số NCS': ['NCS001', 'NCS002'],
        'Khóa': ['K35', 'K35'],
        'Chuyên ngành': ['Công nghệ thông tin', 'Khoa học máy tính'],
        'Ngày sinh': ['01/01/1990', '02/02/1991'],
        'Nơi sinh': ['Hà Nội', 'TP.HCM'],
        'Tên đề tài luận án': ['Nghiên cứu về AI', 'Phát triển phần mềm'],
        'Người hướng dẫn khoa học': ['PGS.TS Nguyễn Văn X', 'GS.TS Lê Văn Y'],
        'Thời gian xét duyệt đề cương': ['01/02/2024', '15/02/2024'],
        'Thời gian bảo vệ đề cương': ['01/03/2024', '15/03/2024'],
        'Địa điểm bảo vệ đề cương': ['Phòng họp A', 'Phòng họp B'],
        'Thời gian chuyên đề 1': ['01/04/2024', '15/04/2024'],
        'Người hướng dẫn chuyên đề 1': ['TS. Phạm Văn M', 'TS. Trần Văn N'],
        'Địa điểm chuyên đề 1': ['Phòng 101', 'Phòng 102'],
        'Thời gian chuyên đề 2': ['01/05/2024', '15/05/2024'],
        'Người hướng dẫn chuyên đề 2': ['TS. Hoàng Văn P', 'TS. Lý Văn Q'],
        'Địa điểm chuyên đề 2': ['Phòng 201', 'Phòng 202'],
        'Thời gian chuyên đề 3': ['01/06/2024', '15/06/2024'],
        'Người hướng dẫn chuyên đề 3': ['TS. Vũ Văn R', 'TS. Đinh Văn S'],
        'Địa điểm chuyên đề 3': ['Phòng 301', 'Phòng 302'],
        'Thời gian bảo vệ cấp Khoa': ['01/07/2024', '15/07/2024'],
        'Địa điểm bảo vệ cấp Khoa': ['Hội trường A', 'Hội trường B'],
        'Thời gian bảo vệ cấp Trường': ['01/08/2024', '15/08/2024'],
        'Địa điểm bảo vệ cấp Trường': ['Hội trường lớn', 'Hội trường lớn'],
        'email': ['hoatrodun@gmail.com', 'hungtrongdoang@gmail.com'],
        'Trạng thái xét duyệt đề cương': ['', ''],
        'Trạng thái bảo vệ đề cương': ['', ''],
        'Trạng thái chuyên đề 1': ['', ''],
        'Trạng thái chuyên đề 2': ['', ''],
        'Trạng thái chuyên đề 3': ['', ''],
        'Trạng thái bảo vệ cấp Khoa': ['', ''],
        'Trạng thái bảo vệ cấp Trường': ['', '']
    }
    
    # Tạo DataFrame
    df = pd.DataFrame(data)
    
    # Lưu thành file Excel
    try:
        df.to_excel('research_progress.xlsx', index=False)
        print("Đã tạo file Excel mẫu thành công!")
    except Exception as e:
        print(f"Lỗi khi tạo file Excel: {str(e)}")

if __name__ == "__main__":
    create_sample_excel()