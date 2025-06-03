import pdfplumber
import pandas as pd
import re
import os

def to_number(val):
    """
    Chuyển số kiểu Việt Nam (1.234,56 hoặc 1,234.56 hoặc 1234,56 hoặc 1234) về float.
    Luôn lấy phần thập phân sau dấu phẩy (nếu có).
    """
    if val is None: return None
    s = str(val).replace(' ', '').replace('\xa0', '')
    # Nếu cả dấu chấm và dấu phẩy, thì dấu phẩy là thập phân
    if '.' in s and ',' in s:
        s = s.replace('.', '').replace(',', '.')
    # Nếu chỉ có dấu phẩy, giả định là thập phân
    elif ',' in s:
        s = s.replace(',', '.')
    # Nếu chỉ có dấu chấm, kiểm tra xem đó là phân cách thập phân hay nghìn
    elif '.' in s:
        # Nếu chỉ một dấu chấm ở cuối (kiểu 1000.25), giữ nguyên
        if s.count('.') == 1 and len(s.split('.')[-1]) <= 3:
            pass
        else:
            # Có nhiều dấu chấm, giả định là phân cách nghìn, bỏ hết
            s = s.replace('.', '')
    try:
        return float(s)
    except:
        return None

def extract_date(text):
    match = re.search(r"Ngày\s+(\d{1,2})\s+tháng\s+(\d{1,2})\s+năm\s+(\d{4})", text, re.IGNORECASE)
    if match:
        return f"{int(match.group(1)):02d}/{int(match.group(2)):02d}/{match.group(3)}"
    return ""

def clean_item_name(name):
    if not name:
        return ''
    name = name.strip()
    name = re.sub(r'^(\d+\s+)?Hàng\s*h[oó]a[,\s\n]*d[iị]ch\s*v[uụ][:,\s\n]*', '', name, flags=re.IGNORECASE)
    return name

def extract_invoice_items(pdf_path, file_name_hint='UNKNOWN'):
    items = []
    known_units = ['Cái', 'Lít', 'm2', 'm', 'Bộ', 'Kg', 'Tấm', 'Ống', 'Phào', 'mét', 'Cặp', 'Chiếc', 'M']
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])
            serial = re.search(r"Ký hiệu.*?:\s*([A-Z0-9]+)", text)
            number = (
                re.search(r"Số[:：]?\s*(\d+)", text) or
                re.search(r"Số hóa đơn[:：]?\s*(\d+)", text) or
                re.search(r"Số HĐ[:：]?\s*(\d+)", text)
            )
            date_str = extract_date(text)
            seller_match = re.search(r"Tên người bán[:：]?\s*(.*)", text)
            seller = seller_match.group(1).strip() if seller_match else ''
            tax = re.search(r"Mã số thuế:?\s*([0-9\-\.]+)", text)

            col_map = {}
            found_header = False

            for page in pdf.pages:
                tables = page.extract_tables()
                for row in tables[0] if tables else []:
                    if not row: continue
                    row = [r.strip() if isinstance(r, str) else '' for r in row]

                    # Tìm header: dòng có >= 6 cột và chứa "tên hàng" hoặc "ten hang"
                    if not found_header and any('tên hàng' in c.lower() or 'ten hang' in c.lower() for c in row):
                        found_header = True
                        for i, col in enumerate(row):
                            c = col.lower()
                            if "tên hàng" in c or "ten hang" in c:
                                col_map["Tên hàng hóa, dịch vụ"] = i
                            elif "đơn vị" in c or "don vi" in c:
                                col_map["Đơn vị tính"] = i
                            elif "số lượng" in c or "so luong" in c:
                                col_map["Số lượng"] = i
                            elif "đơn giá" in c or "don gia" in c:
                                col_map["Đơn giá"] = i
                            elif "thuế suất" in c or "thue suat" in c:
                                col_map["Thuế suất (%)"] = i
                            elif "thành tiền" in c or "thanh tien" in c:
                                col_map["Giá trị HHDV mua vào chưa có thuế GTGT"] = i
                        continue

                    # Nếu đã tìm thấy header và row có đúng số lượng cột, lấy dữ liệu theo col_map
                    if found_header and len(col_map) > 0:
                        def get_col(name):
                            idx = col_map.get(name)
                            if idx is not None and idx < len(row): return row[idx]
                            return ''
                        ten_hh = clean_item_name(get_col("Tên hàng hóa, dịch vụ"))
                        dvt = get_col("Đơn vị tính")
                        so_luong = get_col("Số lượng")
                        don_gia = get_col("Đơn giá")
                        thue_suat = get_col("Thuế suất (%)")
                        thanh_tien = get_col("Giá trị HHDV mua vào chưa có thuế GTGT")

                        # Xử lý số đúng chuẩn
                        so_luong_num = to_number(so_luong)
                        don_gia_num = to_number(don_gia)
                        thanh_tien_num = to_number(thanh_tien)

                        # Chuẩn hóa thuế suất
                        if thue_suat:
                            ts_str = thue_suat.replace("%", "").replace(",", ".").strip()
                            if ts_str.upper() == "KCT":
                                thue_suat_val = "KCT"
                                vat_rate_num = 0
                            else:
                                try:
                                    vat_rate_num = float(ts_str)
                                    thue_suat_val = ts_str
                                except:
                                    vat_rate_num = 0
                                    thue_suat_val = ""
                        else:
                            vat_rate_num = 0
                            thue_suat_val = ""

                        # Tính tiền thuế GTGT
                        if thanh_tien_num is not None and vat_rate_num > 0:
                            vat_tax = round(thanh_tien_num * vat_rate_num / 100, 2)
                        else:
                            vat_tax = ""

                        if not ten_hh or (so_luong == '' and don_gia == ''):
                            continue

                        items.append({
                            'Mẫu số': '01GTKT0/001',
                            'Ký hiệu': serial.group(1).strip() if serial else '',
                            'Số': number.group(1).strip() if number else file_name_hint,
                            'Ngày, tháng, năm': date_str,
                            'Tên người bán': seller,
                            'Mã số thuế người bán': tax.group(1).strip() if tax else '',
                            'Tên hàng hóa, dịch vụ': ten_hh,
                            'Đơn vị tính': dvt,
                            'Số lượng': so_luong_num if so_luong != "" else "",
                            'Đơn giá': don_gia_num if don_gia != "" else "",
                            'Giá trị HHDV mua vào chưa có thuế GTGT': thanh_tien_num if thanh_tien != "" else "",
                            'Thuế suất (%)': thue_suat if thue_suat else "",
                            'Tiền thuế GTGT': vat_tax,
                            'Ghi chú': ''
                        })
    except Exception as e:
        print(f"Lỗi khi xử lý file {pdf_path}: {e}")
    return items

def main(pdf_dir, output_file):
    all_data = []
    for file in os.listdir(pdf_dir):
        if file.lower().endswith(".pdf"):
            file_path = os.path.join(pdf_dir, file)
            items = extract_invoice_items(file_path, file_name_hint=file)
            all_data.extend(items)

    # Đảm bảo đúng cột và thứ tự như yêu cầu
    cols = [
        'STT',
        'Mẫu số',
        'Ký hiệu',
        'Số',
        'Ngày, tháng, năm',
        'Tên người bán',
        'Mã số thuế người bán',
        'Tên hàng hóa, dịch vụ',
        'Đơn vị tính',
        'Số lượng',
        'Đơn giá',
        'Giá trị HHDV mua vào chưa có thuế GTGT',
        'Thuế suất (%)',
        'Tiền thuế GTGT',
        'Ghi chú'
    ]
    df = pd.DataFrame(all_data)
    for col in cols:
        if col not in df.columns:
            df[col] = ''
    df = df[cols]
    # Xóa dòng header hoặc dòng trống
    df = df[~df['Tên hàng hóa, dịch vụ'].str.lower().str.contains('tên hàng hóa|đơn vị tính|char', na=False)]
    df = df[~(
        (df['Tên hàng hóa, dịch vụ'].isna() | (df['Tên hàng hóa, dịch vụ'] == '')) &
        (df['Số lượng'].isna() | (df['Số lượng'] == '')) &
        (df['Đơn giá'].isna() | (df['Đơn giá'] == ''))
    )]
    df['STT'] = range(1, len(df) + 1)
    try:
        df.to_excel(output_file, index=False)
        print(f"✅ Đã xuất file: {output_file}")
    except PermissionError:
        print(f"⚠️ Không thể ghi đè file {output_file}. Đang ghi vào file dự phòng...")
        fallback_file = output_file.replace('.xlsx', '_v2.xlsx')
        df.to_excel(fallback_file, index=False)
        print(f"✅ Đã ghi vào: {fallback_file}")

if __name__ == "__main__":
    main('./pdfs', 'Ket_qua_hoa_don_final.xlsx')