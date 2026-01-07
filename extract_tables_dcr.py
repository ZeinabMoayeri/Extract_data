import arabic_reshaper
from bidi.algorithm import get_display
import pdfplumber
import pandas as pd
import os
import json
from flatten import flatten_with_pikepdf
from persian_text import correct_persian_text
#from utils.load_coordinates_points import load_coordinates_points
import config

def load_coordinates_points(file_path):
    if not os.path.exists(file_path):
        return []

    try:
        with open(file_path, encoding="utf-8") as f:
            return json.load(f)
    except json.JSONDecodeError as e:
        print("Invalid JSON format")
    except Exception:
        print("Unexpected error while loading coordinates points")  
    return []
    
def convert_header_to_key_value(df):
    if df.empty:
        return {}
    
    # گرفتن اولین ردیف
    first_row = df.iloc[0].tolist()
    
    result = {}
    
    # برای تاریخ: مستقیماً column های 0, 1, 2, 3 را merge کن
    date_parts = []
    for col_idx in [0, 1, 2, 3]:
        if col_idx < len(first_row) and pd.notna(first_row[col_idx]):
            part = str(first_row[col_idx]).strip()
            if part:
                # حذف فاصله‌ها و "/" و ":" و ترکیب
                part_clean = part.replace(" ", "").replace("/", "").replace(":", "").strip()
                if part_clean:
                    date_parts.append(part_clean)
    
    # ترکیب همه قسمت‌ها برای تاریخ
    if date_parts:
        combined_date = "".join(date_parts)
        result["تاریخ"] = combined_date
    
    # حالا پردازش بقیه سلول‌ها
    i = 0
    while i < len(first_row):
        current = str(first_row[i]).strip() if pd.notna(first_row[i]) else ""
        next_val = str(first_row[i + 1]).strip() if i + 1 < len(first_row) and pd.notna(first_row[i + 1]) else ""
        
        # اگر سلول خالی است یا در column های 0-3 است (که برای تاریخ استفاده شده)، رد شو
        # یا اگر "تاریخ:" است (که قبلاً پردازش شده)، رد شو
        if not current or i < 4 or ("تاریخ" in current and ":" in current):
            i += 1
            continue
        
        # حالت 1: سلول فعلی به ":" ختم می‌شود و سلول بعدی مقدار دارد
        if current.endswith(":") and next_val and not next_val.endswith(":"):
            key = current.rstrip(":").strip()
            value = next_val
            result[key] = value
            i += 2
            continue
        
        # حالت 2: سلول فعلی مقدار دارد و سلول بعدی به ":" ختم می‌شود (مثل "O3" و "دستگاه حفاری:")
        if next_val.endswith(":") and current and not current.endswith(":"):
            key = next_val.rstrip(":").strip()
            value = current
            result[key] = value
            i += 2
            continue
        
        # حالت 3: سلول فعلی شامل ":" است (مثل "مسوول اردوگاه: اصلان ویسى")
        if ":" in current:
            parts = current.split(":", 1)
            if len(parts) == 2:
                key = parts[0].strip()
                value = parts[1].strip()
                if key and value:
                    result[key] = value
            i += 1
            continue
        
        # حالت 4: اگر هیچکدام از موارد بالا نبود، فقط مقدار را نگه دار
        i += 1
    
    return result

def convert_shift_to_structured(df):
    """
    تبدیل داده‌های یک شیفت به ساختار جدید
    df: DataFrame مربوط به یک شیفت
    """
    if df.empty or len(df) < 3:
        return {"persons": [], "TotalShift": {}}
    
    persons = []
    total_shift = {}
    
    # تبدیل به records برای پردازش راحت‌تر
    records = df.to_dict('records')
    
    # رد کردن دو ردیف اول (نام شیفت و هدر)
    data_rows = records[2:]
    
    # پیدا کردن ردیف Total (آخرین ردیفی که شامل "مجموع آمار شیفت" است)
    total_row_index = -1
    for i in range(len(data_rows) - 1, -1, -1):
        row = data_rows[i]
        # ستون 5 معمولاً شامل position است
        col5 = str(row.get(5, "")).strip() if 5 in row else ""
        if "مجموع آمار شیفت" in col5:
            total_row_index = i
            break
    
    # استخراج Total
    if total_row_index >= 0:
        total_row = data_rows[total_row_index]
        total_shift = {
            "ص": str(total_row.get(4, "")).strip() if 4 in total_row else "",
            "ن": str(total_row.get(3, "")).strip() if 3 in total_row else "",
            "ش": str(total_row.get(2, "")).strip() if 2 in total_row else "",
            "پ": str(total_row.get(1, "")).strip() if 1 in total_row else "",
            "خ": str(total_row.get(0, "")).strip() if 0 in total_row else ""
        }
        # حذف ردیف Total از لیست داده‌ها
        data_rows = data_rows[:total_row_index]
    
    # استخراج اطلاعات افراد
    for row in data_rows:
        name = str(row.get(6, "")).strip() if 6 in row else ""
        position = str(row.get(5, "")).strip() if 5 in row else ""
        
        # اگر نام یا پوزیشن خالی است، رد کن
        if not name and not position:
            continue
        
        person = {
            "name": name,
            "position": position,
            "ص": str(row.get(4, "")).strip() if 4 in row else "",
            "ن": str(row.get(3, "")).strip() if 3 in row else "",
            "ش": str(row.get(2, "")).strip() if 2 in row else "",
            "پ": str(row.get(1, "")).strip() if 1 in row else "",
            "خ": str(row.get(0, "")).strip() if 0 in row else ""
        }
        persons.append(person)
    
    return {"persons": persons, "TotalShift": total_shift}

def extract_tables_from_dcr(pdf_file, output_folder, tables_info):
    if not os.path.exists(pdf_file):
        return

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    # دیکشنری برای ذخیره همه جداول
    all_tables_data = {}
    # ایجاد ساختار Operation برای ذخیره شیفت‌ها
    all_tables_data["Operation"] = {}
        
    with pdfplumber.open(pdf_file) as pdf:
        for table_meta in tables_info:
            page_num = table_meta["page_number"]
            sheet_name = table_meta["sheet_name"]
            coords = table_meta["coordinates"]
            
            if page_num > len(pdf.pages):
                print(f"Not found page in PDF: {page_num}")
                continue
            
            bbox = (coords[1], coords[0], coords[3], coords[2])
            
            page = pdf.pages[page_num - 1]
            cropped_page = page.crop(bbox)
            
            df = None
            
            my_table_settings = {
                "vertical_strategy": "lines", 
                "horizontal_strategy": "lines",
                "snap_tolerance": 4,
                "join_tolerance": 4,
                "intersection_tolerance": 5,
            }
            
            table_data = cropped_page.extract_table(table_settings=my_table_settings)
            
            if table_data:
                # 1. ساخت دیتافریم با کل داده‌ها (بدون جدا کردن هدر)
                df = pd.DataFrame(table_data)
                
                # 2. نام‌گذاری ستون‌ها به صورت پیش‌فرض (Column 1, Column 2, ...)
                # اگر می‌خواهید فارسی باشد، داخل پرانتز f"ستون {i+1}" بنویسید
                #df.columns = [f"Column {i+1}" for i in range(df.shape[1])]

            if df is not None:
                df.dropna(axis=0, how='all', inplace=True)
                df.dropna(axis=1, how='all', inplace=True)
                df.fillna('', inplace=True)

                try:
                    df = df.applymap(lambda x: correct_persian_text(x) if isinstance(x, str) else x)
                except Exception as e:
                    print("Unexpected error while correcting Persian text")
                # تبدیل DataFrame به ساختار مناسب
                try:
                    # اگر Header است، به key-value تبدیل کن
                    if sheet_name == "Header":
                        key_value_dict = convert_header_to_key_value(df)
                        all_tables_data[sheet_name] = key_value_dict
                    # اگر مربوط به Operation است (شیفت‌ها)
                    elif sheet_name in ["ShiftAPage1", "ShiftBPage1", "ShiftCPage1", "ShiftDPage1"]:
                        shift_data = convert_shift_to_structured(df)
                        # تبدیل نام sheet به نام شیفت (مثلاً ShiftAPage1 -> ShiftA)
                        shift_name = sheet_name.replace("Page1", "")
                        # ایجاد لیست اشخاص
                        shift_list = shift_data["persons"].copy() if shift_data["persons"] else []
                        # اضافه کردن TotalShift در انتها
                        if shift_data["TotalShift"]:
                            total_key = f"Total{shift_name}"
                            shift_list.append({total_key: shift_data["TotalShift"]})
                        all_tables_data["Operation"][shift_name] = shift_list
                    # اگر ShiftTotalPage1 است
                    elif sheet_name == "ShiftTotalPage1":
                        records = df.to_dict('records')
                        if len(records) >= 3:
                            # ردیف سوم شامل Total است
                            total_row = records[2]
                            all_tables_data["Operation"]["ShiftTotalPage1"] = {
                                "ص": str(total_row.get(4, "")).strip() if 4 in total_row else "",
                                "ن": str(total_row.get(3, "")).strip() if 3 in total_row else "",
                                "ش": str(total_row.get(2, "")).strip() if 2 in total_row else "",
                                "پ": str(total_row.get(1, "")).strip() if 1 in total_row else "",
                                "خ": str(total_row.get(0, "")).strip() if 0 in total_row else ""
                            }
                        else:
                            all_tables_data["Operation"]["ShiftTotalPage1"] = {}
                    else:
                        # برای جداول دیگر، به records تبدیل کن
                        table_records = df.to_dict('records')
                        all_tables_data[sheet_name] = table_records
                except Exception as inner_e:
                    print("Unexpected error while extracting tables")
    
    # اطمینان از وجود بخش‌های خالی برای ساختار نهایی
    if "Operation" not in all_tables_data:
        all_tables_data["Operation"] = {}
    
    # اطمینان از وجود شیفت‌ها به صورت لیست خالی اگر وجود نداشتند
    if "ShiftA" not in all_tables_data["Operation"]:
        all_tables_data["Operation"]["ShiftA"] = []
    if "ShiftB" not in all_tables_data["Operation"]:
        all_tables_data["Operation"]["ShiftB"] = []
    if "ShiftC" not in all_tables_data["Operation"]:
        all_tables_data["Operation"]["ShiftC"] = []
    if "ShiftD" not in all_tables_data["Operation"]:
        all_tables_data["Operation"]["ShiftD"] = []
    if "ShiftTotalPage1" not in all_tables_data["Operation"]:
        all_tables_data["Operation"]["ShiftTotalPage1"] = {}
    
    # اضافه کردن بخش‌های خالی دیگر
    all_tables_data["Herasat"] = {}
    all_tables_data["Ordogahi"] = {}
    all_tables_data["employer"] = {}
    all_tables_data["Drilling"] = {}
    all_tables_data["Foods"] = {}
    all_tables_data["total"] = {}
    
    # ساخت ساختار نهایی JSON با ترتیب صحیح
    final_structure = {}
    
    # Header
    if "Header" in all_tables_data:
        final_structure["Header"] = all_tables_data["Header"]
    else:
        final_structure["Header"] = {
            "تاریخ": "",
            "روز هفته": "",
            "مسوول اردوگاه": "",
            "رییس دستگاه": "",
            "دستگاه حفاری": ""
        }
    
    # Operation
    final_structure["Operation"] = all_tables_data.get("Operation", {})
    
    # بقیه بخش‌ها
    final_structure["Herasat"] = all_tables_data.get("Herasat", {})
    final_structure["Ordogahi"] = all_tables_data.get("Ordogahi", {})
    final_structure["employer"] = all_tables_data.get("employer", {})
    final_structure["Drilling"] = all_tables_data.get("Drilling", {})
    final_structure["Foods"] = all_tables_data.get("Foods", {})
    final_structure["total"] = all_tables_data.get("total", {})
    
    # ذخیره همه جداول در یک فایل JSON
    if final_structure:
        pdf_basename = os.path.splitext(os.path.basename(pdf_file))[0]
        json_filename = f"{pdf_basename}_tables.json"
        json_path = os.path.join(output_folder, json_filename)
        
        try:
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(final_structure, f, ensure_ascii=False, indent=2)
            print(f"Complete Process for {os.path.basename(pdf_file)} - Saved to {json_filename}")
        except Exception as e:
            print("Unexpected error while saving JSON file")

    else:
        print("No tables extracted from PDF")
if __name__ == "__main__":
    output_folder = "./extract_tables_dcr"
    pdf_file = "DCR O3 1404 1007.pdf"
    
    output_file= "DCR O3 1404 1007_flatten.pdf"
    
    if not os.path.exists(output_file) and os.path.exists(pdf_file):
        flatten_with_pikepdf(pdf_file, output_file)
    elif not os.path.exists(pdf_file):
        print("Input file not found.")
    
    if os.path.exists(output_file):
        coordinates_points_file_path = config.COORDINATES_POINTS_PATH
        coordinates_points = load_coordinates_points(coordinates_points_file_path)
        extract_tables_from_dcr(output_file, output_folder, coordinates_points["tables_metadataـDCR"])