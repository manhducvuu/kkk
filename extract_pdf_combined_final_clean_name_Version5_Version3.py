import pdfplumber
import pandas as pd
import re
import os

def to_number(val):
    if val is None: return None
    val = str(val).replace('.', '').replace(',', '.')
    try:
        return float(val)
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
                        # Tính tiền thuế GTGT
                        value = to_number(thanh_tien)
                        try:
                            vat_rate = thue_suat
                            vat_rate_num = int(vat_rate.replace("%", "").strip()) if vat_rate and vat_rate != 'KCT' else 0
                            vat_tax = round(value * vat_rate_num / 100, 0) if value and vat_rate_num > 0 else ''
                        except:
                            vat_tax = ''

                        if not ten_hh or (not so_luong and not don_gia):
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
                            'Số lượng': so_luong,
                            'Đơn giá': don_gia,
                            'Giá trị HHDV mua vào chưa có thuế GTGT': thanh_tien,
                            'Thuế suất (%)': thue_suat,
                            'Tiền thuế GTGT': vat_tax,
                            'Ghi chú': ''
                        })
                    # Nếu chưa có header hoặc row không mapping được, dùng fallback mẫu cũ (chỉ nếu đủ cột/thông tin)
                    elif not found_header and len(row) >= 6:
                        unit_idx = -1
                        for i in range(len(row)-1, 0, -1):
                            if any(row[i].lower() == u.lower() for u in known_units if row[i]):
                                unit_idx = i
                                break
                        if unit_idx == -1:
                            unit_idx = 2
                        numeric_fields = row[unit_idx+1:]
                        quantity, unit_price = '', ''
                        if len(numeric_fields) >= 2:
                            quantity = numeric_fields[0]
                            unit_price = numeric_fields[1]
                        tax_rate = ''
                        for val in reversed(numeric_fields):
                            if not val:
                                continue
                            stripped = val.strip().replace('%', '')
                            if stripped.isdigit():
                                num = int(stripped)
                                if num in [5, 8, 10]:
                                    tax_rate = str(num)
                                    break
                            elif val.strip().upper() == 'KCT':
                                tax_rate = 'KCT'
                                break
                        name_parts = row[1:unit_idx]
                        name = " ".join([p for p in name_parts if p])
                        qty_val = to_number(quantity) if quantity else None
                        price_val = to_number(unit_price) if unit_price else None
                        value = None
                        for v in reversed(row):
                            v_num = to_number(v)
                            if v_num is not None and v_num > 0:
                                value = v_num
                                break
                        if (value is None or value == 0) and (qty_val is not None and price_val is not None):
                            value = round(qty_val * price_val, 2)
                        if value is None or value == 0:
                            value = ''
                        try:
                            vat_rate_num = int(tax_rate) if tax_rate and tax_rate != 'KCT' else 0
                        except:
                            vat_rate_num = 0
                        vat_tax = round(value * vat_rate_num / 100, 0) if value != '' and vat_rate_num > 0 else ''
                        items.append({
                            'Mẫu số': '01GTKT0/001',
                            'Ký hiệu': serial.group(1).strip() if serial else '',
                            'Số': number.group(1).strip() if number else file_name_hint,
                            'Ngày, tháng, năm': date_str,
                            'Tên người bán': seller,
                            'Mã số thuế người bán': tax.group(1).strip() if tax else '',
                            'Tên hàng hóa, dịch vụ': clean_item_name(name),
                            'Đơn vị tính': row[unit_idx] if unit_idx < len(row) else '',
                            'Số lượng': quantity if quantity else '',
                            'Đơn giá': unit_price if unit_price else '',
                            'Giá trị HHDV mua vào chưa có thuế GTGT': value,
                            'Thuế suất (%)': tax_rate if tax_rate else '',
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
    except PermissionError:
        print(f"⚠️ Không thể ghi đè file {output_file}. Đang ghi vào file dự phòng...")
        fallback_file = output_file.replace('.xlsx', '_v2.xlsx')
        df.to_excel(fallback_file, index=False)
        print(f"✅ Đã ghi vào: {fallback_file}")

if __name__ == "__main__":
    main('./pdfs', 'Ket_qua_hoa_don_final.xlsx')