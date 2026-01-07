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

def convert_to_int(value):
    """
    تبدیل مقدار string به int
    اگر خالی است یا نمی‌تواند تبدیل شود، 0 برمی‌گرداند
    اعداد منفی را هم درست پردازش می‌کند
    """
    if not value or value == "":
        return 0
    try:
        # حذف فاصله‌ها و تبدیل به int (اعداد منفی را حفظ می‌کند)
        clean_val = str(value).strip().replace(" ", "").replace(",", "")
        if clean_val:
            return int(clean_val)
        return 0
    except:
        return 0
    
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
            "ص": convert_to_int(total_row.get(4, "") if 4 in total_row else ""),
            "ن": convert_to_int(total_row.get(3, "") if 3 in total_row else ""),
            "ش": convert_to_int(total_row.get(2, "") if 2 in total_row else ""),
            "پ": convert_to_int(total_row.get(1, "") if 1 in total_row else ""),
            "خ": convert_to_int(total_row.get(0, "") if 0 in total_row else "")
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
            "ص": convert_to_int(row.get(4, "") if 4 in row else ""),
            "ن": convert_to_int(row.get(3, "") if 3 in row else ""),
            "ش": convert_to_int(row.get(2, "") if 2 in row else ""),
            "پ": convert_to_int(row.get(1, "") if 1 in row else ""),
            "خ": convert_to_int(row.get(0, "") if 0 in row else "")
        }
        persons.append(person)
    
    return {"persons": persons, "TotalShift": total_shift}

def convert_employer_to_structured(df):
    """
    تبدیل داده‌های employer به ساختار جدید
    df: DataFrame مربوط به employer
    ترتیب ستون‌ها: نام، سمت، شرکت، ص، ن، ش، پ، خ
    """
    if df.empty or len(df) < 2:
        return []
    
    persons = []
    
    # تبدیل به records برای پردازش راحت‌تر
    records = df.to_dict('records')
    
    # رد کردن ردیف اول (هدر)
    data_rows = records[1:]
    
    # استخراج اطلاعات افراد
    # ترتیب ستون‌ها در employer (از راست به چپ): خ, پ, ش, ن, ص, company, position, name
    # یعنی: column 0=خ, 1=پ, 2=ش, 3=ن, 4=ص, 5=company, 6=position, 7=name
    for row in data_rows:
        name = str(row.get(7, "")).strip() if 7 in row else ""
        position = str(row.get(6, "")).strip() if 6 in row else ""
        company = str(row.get(5, "")).strip() if 5 in row else ""
        
        # اگر نام یا پوزیشن خالی است، رد کن
        if not name and not position:
            continue
        
        person = {
            "name": name,
            "position": position,
            "company": company,
            "ص": convert_to_int(row.get(4, "") if 4 in row else ""),
            "ن": convert_to_int(row.get(3, "") if 3 in row else ""),
            "ش": convert_to_int(row.get(2, "") if 2 in row else ""),
            "پ": convert_to_int(row.get(1, "") if 1 in row else ""),
            "خ": convert_to_int(row.get(0, "") if 0 in row else "")
        }
        persons.append(person)
    
    return persons

def convert_foods_to_structured(df):
    """
    تبدیل داده‌های Foods به ساختار جدید
    df: DataFrame مربوط به Foods
    ساختار: صبحانه، ناهار، شام، پس شام (value در ستون کناری) و توضیحات (value در row پایینی)
    """
    if df.empty:
        return {
            "صبحانه": "",
            "ناهار": "",
            "شام": "",
            "پس شام": "",
            "توضیحات": ""
        }
    
    foods_data = {
        "صبحانه": "",
        "ناهار": "",
        "شام": "",
        "پس شام": "",
        "توضیحات": ""
    }
    
    # تبدیل به records
    records = df.to_dict('records')
    
    # پیدا کردن صبحانه، ناهار، شام، پس شام (value در ستون قبلی)
    for row_idx, row in enumerate(records):
        # بررسی همه ستون‌ها برای پیدا کردن key-value
        max_col = max(row.keys()) if row else 0
        for i in range(max_col + 1):
            key = str(row.get(i, "")).strip() if i in row else ""
            prev_val = str(row.get(i - 1, "")).strip() if i - 1 >= 0 and (i - 1) in row else ""
            next_val = str(row.get(i + 1, "")).strip() if i + 1 in row else ""
            
            # صبحانه، ناهار، شام، پس شام: value در ستون قبلی (i-1)
            # توجه: "پس شام" را قبل از "شام" بررسی می‌کنیم چون "پس شام" شامل "شام" است
            if "صبحانه" in key and prev_val:
                foods_data["صبحانه"] = prev_val
            elif "ناهار" in key and prev_val:
                foods_data["ناهار"] = prev_val
            elif ("پس شام" in key or "پس‌شام" in key) and prev_val:
                foods_data["پس شام"] = prev_val
            elif "شام" in key and prev_val:
                foods_data["شام"] = prev_val
            elif "توضیحات" in key:
                # توضیحات: value در row پایینی یا در ستون کناری
                if next_val:
                    foods_data["توضیحات"] = next_val
                # یا ممکن است در همان سلول باشد (بعد از ":")
                elif ":" in key:
                    parts = key.split(":", 1)
                    if len(parts) == 2:
                        foods_data["توضیحات"] = parts[1].strip()
    
    # اگر توضیحات پیدا نشد، آخرین ردیف را بررسی کن (row پایینی)
    if not foods_data["توضیحات"] and len(records) > 0:
        last_row = records[-1]
        # بررسی اینکه آیا ردیف آخر شامل توضیحات است
        max_col = max(last_row.keys()) if last_row else 0
        for i in range(max_col + 1):
            val = str(last_row.get(i, "")).strip() if i in last_row else ""
            if val and "توضیحات" not in val and val not in ["صبحانه", "ناهار", "شام", "پس شام", "پس‌شام"]:
                # اگر مقدار طولانی است یا شامل متن است، احتمالاً توضیحات است
                if len(val) > 5:
                    foods_data["توضیحات"] = val
                    break
    
    return foods_data

def convert_total_to_structured(df):
    """
    تبدیل داده‌های Total به ساختار جدید
    df: DataFrame مربوط به Total
    ساختار: 7 key اصلی که هر کدام 5 value دارند (صبحانه، ناهار، شام، پس شام، خدمات)
    """
    if df.empty:
        return {}
    
    total_data = {}
    
    # تبدیل به records
    records = df.to_dict('records')
    
    if len(records) == 0:
        return {}
    
    # پیدا کردن هدر جدول (ردیف اول که شامل صبحانه، ناهار، شام، پس شام، خدمات است)
    header_cols = {}  # mapping از نام ستون به index
    
    first_row = records[0]
    max_col = max(first_row.keys()) if first_row else 0
    for i in range(max_col + 1):
        val = str(first_row.get(i, "")).strip() if i in first_row else ""
        # حذف فاصله‌ها و کاراکترهای اضافی برای مقایسه بهتر
        val_clean = val.replace(" ", "").replace("\t", "").replace("\n", "").replace("،", "")
        
        # بررسی با حذف فاصله‌ها (مثلاً "خدما ت" -> "خدمات", "شا م" -> "شام")
        if "صبحانه" in val_clean:
            header_cols["صبحانه"] = i
        elif "ناهار" in val_clean:
            header_cols["ناهار"] = i
        elif "پسشام" in val_clean or ("پس" in val_clean and "شام" in val_clean):
            header_cols["پس شام"] = i
        elif "شام" in val_clean:
            header_cols["شام"] = i
        elif "خدمات" in val_clean:
            header_cols["خدمات"] = i
    
    # اگر هدر پیدا نشد، از ردیف اول شروع کن، وگرنه از ردیف دوم
    start_idx = 1 if header_cols else 0
    
    # بررسی هر ردیف برای پیدا کردن کلیدهای اصلی (7 ردیف)
    for row_idx in range(start_idx, min(start_idx + 7, len(records))):
        row = records[row_idx]
        max_col = max(row.keys()) if row else 0
        
        # پیدا کردن کلید اصلی: اولین ستونی که متن دارد و عدد نیست
        main_key = None
        main_key_col = None
        
        # بررسی ستون‌ها از چپ به راست
        for i in range(max_col + 1):
            val = str(row.get(i, "")).strip() if i in row else ""
            
            if val:
                # بررسی کن که آیا این یک عدد نیست
                is_number = False
                try:
                    # حذف فاصله و کاما و بررسی عدد
                    clean_val = val.replace(",", "").replace(" ", "").replace("-", "").replace("+", "")
                    if clean_val:
                        float(clean_val)
                        is_number = True
                except:
                    pass
                
                # اگر عدد نیست و طول مناسبی دارد، احتمالاً کلید اصلی است
                if not is_number and len(val) > 1:
                    main_key = val
                    main_key_col = i
                    break
        
        if main_key:
            # استخراج 5 value: صبحانه، ناهار، شام، پس شام، خدمات
            values = {
                "صبحانه": 0,
                "ناهار": 0,
                "شام": 0,
                "پس شام": 0,
                "خدمات": 0
            }
            
            # اگر هدر پیدا کردیم، از mapping استفاده کن
            if header_cols:
                for key_name, col_idx in header_cols.items():
                    if col_idx in row:
                        val = str(row.get(col_idx, "")).strip()
                        values[key_name] = convert_to_int(val)
            else:
                # اگر هدر پیدا نکردیم، از ستون‌های بعد از key اصلی استفاده کن
                for idx, key_name in enumerate(["صبحانه", "ناهار", "شام", "پس شام", "خدمات"]):
                    col_idx = main_key_col + 1 + idx
                    if col_idx in row:
                        val = str(row.get(col_idx, "")).strip()
                        values[key_name] = convert_to_int(val)
            
            total_data[main_key] = values
    
    return total_data

def extract_tables_from_dcr(pdf_file, output_folder, tables_info):
    if not os.path.exists(pdf_file):
        return

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    # دیکشنری برای ذخیره همه جداول
    all_tables_data = {}
    # ایجاد ساختار Operation، Herasat، Ordogahi، employer، Drilling و Foods برای ذخیره داده‌ها
    all_tables_data["Operation"] = {}
    all_tables_data["Herasat"] = {}
    all_tables_data["Ordogahi"] = {}
    all_tables_data["employer"] = {}
    all_tables_data["Drilling"] = {}
    all_tables_data["Foods"] = {}
    all_tables_data["total"] = {}
        
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
                        total_row = None
                        # بررسی ردیف‌ها برای پیدا کردن Total
                        if len(records) >= 3:
                            # ردیف سوم شامل Total است
                            total_row = records[2]
                        elif len(records) >= 2:
                            # اگر فقط 2 ردیف داریم، ردیف دوم را بررسی کن
                            total_row = records[1]
                        elif len(records) >= 1:
                            # اگر فقط 1 ردیف داریم، همان را بررسی کن
                            total_row = records[0]
                        
                        if total_row:
                            all_tables_data["Operation"]["ShiftTotalPage1"] = {
                                "ص": convert_to_int(total_row.get(4, "") if 4 in total_row else ""),
                                "ن": convert_to_int(total_row.get(3, "") if 3 in total_row else ""),
                                "ش": convert_to_int(total_row.get(2, "") if 2 in total_row else ""),
                                "پ": convert_to_int(total_row.get(1, "") if 1 in total_row else ""),
                                "خ": convert_to_int(total_row.get(0, "") if 0 in total_row else "")
                            }
                        else:
                            all_tables_data["Operation"]["ShiftTotalPage1"] = {}
                    # اگر مربوط به Herasat است (شیفت‌ها)
                    elif sheet_name in ["HerasatShiftA", "HerasatShiftB", "HerasatShiftC", "HerasatShiftD"]:
                        shift_data = convert_shift_to_structured(df)
                        # تبدیل نام sheet به نام شیفت (مثلاً HerasatShiftA -> ShiftA)
                        shift_name = sheet_name.replace("Herasat", "")
                        # ایجاد لیست اشخاص
                        shift_list = shift_data["persons"].copy() if shift_data["persons"] else []
                        # اضافه کردن TotalShift در انتها
                        if shift_data["TotalShift"]:
                            total_key = f"Total{shift_name}"
                            shift_list.append({total_key: shift_data["TotalShift"]})
                        all_tables_data["Herasat"][shift_name] = shift_list
                    # اگر ShiftTotalHerasat است
                    elif sheet_name == "ShiftTotalHerasat":
                        records = df.to_dict('records')
                        total_row = None
                        # بررسی ردیف‌ها برای پیدا کردن Total
                        if len(records) >= 3:
                            # ردیف سوم شامل Total است
                            total_row = records[2]
                        elif len(records) >= 2:
                            # اگر فقط 2 ردیف داریم، ردیف دوم را بررسی کن
                            total_row = records[1]
                        elif len(records) >= 1:
                            # اگر فقط 1 ردیف داریم، همان را بررسی کن
                            total_row = records[0]
                        
                        if total_row:
                            all_tables_data["Herasat"]["ShiftTotalHerasat"] = {
                                "ص": convert_to_int(total_row.get(4, "") if 4 in total_row else ""),
                                "ن": convert_to_int(total_row.get(3, "") if 3 in total_row else ""),
                                "ش": convert_to_int(total_row.get(2, "") if 2 in total_row else ""),
                                "پ": convert_to_int(total_row.get(1, "") if 1 in total_row else ""),
                                "خ": convert_to_int(total_row.get(0, "") if 0 in total_row else "")
                            }
                        else:
                            all_tables_data["Herasat"]["ShiftTotalHerasat"] = {}
                    # اگر مربوط به Ordogahi است (شیفت‌ها)
                    elif sheet_name in ["OrdogahiShiftA", "OrdogahiShiftB", "OrdogahiShiftC", "OrdogahiShiftD"]:
                        shift_data = convert_shift_to_structured(df)
                        # تبدیل نام sheet به نام شیفت (مثلاً OrdogahiShiftA -> ShiftA)
                        shift_name = sheet_name.replace("Ordogahi", "")
                        # ایجاد لیست اشخاص
                        shift_list = shift_data["persons"].copy() if shift_data["persons"] else []
                        # اضافه کردن TotalShift در انتها
                        if shift_data["TotalShift"]:
                            total_key = f"Total{shift_name}"
                            shift_list.append({total_key: shift_data["TotalShift"]})
                        all_tables_data["Ordogahi"][shift_name] = shift_list
                    # اگر ShiftTotalOrdogahi است
                    elif sheet_name == "ShiftTotalOrdogahi":
                        records = df.to_dict('records')
                        total_row = None
                        # بررسی ردیف‌ها برای پیدا کردن Total
                        if len(records) >= 3:
                            # ردیف سوم شامل Total است
                            total_row = records[2]
                        elif len(records) >= 2:
                            # اگر فقط 2 ردیف داریم، ردیف دوم را بررسی کن
                            total_row = records[1]
                        elif len(records) >= 1:
                            # اگر فقط 1 ردیف داریم، همان را بررسی کن
                            total_row = records[0]
                        
                        if total_row:
                            all_tables_data["Ordogahi"]["ShiftTotalOrdogahi"] = {
                                "ص": convert_to_int(total_row.get(4, "") if 4 in total_row else ""),
                                "ن": convert_to_int(total_row.get(3, "") if 3 in total_row else ""),
                                "ش": convert_to_int(total_row.get(2, "") if 2 in total_row else ""),
                                "پ": convert_to_int(total_row.get(1, "") if 1 in total_row else ""),
                                "خ": convert_to_int(total_row.get(0, "") if 0 in total_row else "")
                            }
                        else:
                            all_tables_data["Ordogahi"]["ShiftTotalOrdogahi"] = {}
                    # اگر مربوط به employer است
                    elif sheet_name == "EmployerPage3":
                        persons_list = convert_employer_to_structured(df)
                        all_tables_data["employer"]["EmployerPage3"] = persons_list
                    # اگر EmployerTotal است
                    elif sheet_name == "EmployerTotal":
                        records = df.to_dict('records')
                        total_row = None
                        # بررسی ردیف‌ها برای پیدا کردن Total
                        if len(records) >= 3:
                            total_row = records[2]
                        elif len(records) >= 2:
                            total_row = records[1]
                        elif len(records) >= 1:
                            total_row = records[0]
                        
                        if total_row:
                            all_tables_data["employer"]["EmployerTotal"] = {
                                "ص": convert_to_int(total_row.get(4, "") if 4 in total_row else ""),
                                "ن": convert_to_int(total_row.get(3, "") if 3 in total_row else ""),
                                "ش": convert_to_int(total_row.get(2, "") if 2 in total_row else ""),
                                "پ": convert_to_int(total_row.get(1, "") if 1 in total_row else ""),
                                "خ": convert_to_int(total_row.get(0, "") if 0 in total_row else "")
                            }
                        else:
                            all_tables_data["employer"]["EmployerTotal"] = {}
                    # اگر EmployerSupervisor است
                    elif sheet_name == "EmployerSupervisor":
                        # EmployerSupervisor: ردیف دوم را می‌گیریم
                        records = df.to_dict('records')
                        supervisor_data = {}
                        # بررسی وجود ردیف دوم
                        if len(records) >= 2:
                            # ردیف دوم را بگیر
                            second_row = records[1]
                            # تبدیل به دیکشنری ساده
                            for key, value in second_row.items():
                                if value and str(value).strip():
                                    supervisor_data[str(key)] = str(value).strip()
                        elif len(records) >= 1:
                            # اگر فقط یک ردیف داریم، همان را بگیر
                            first_row = records[0]
                            for key, value in first_row.items():
                                if value and str(value).strip():
                                    supervisor_data[str(key)] = str(value).strip()
                        all_tables_data["employer"]["EmployerSupervisor"] = supervisor_data
                    # اگر مربوط به Drilling است
                    elif sheet_name == "DrillingPage4":
                        persons_list = convert_employer_to_structured(df)
                        all_tables_data["Drilling"]["DrillingPage4"] = persons_list
                    # اگر DrillingTotal است
                    elif sheet_name == "DrillingTotal":
                        records = df.to_dict('records')
                        total_row = None
                        # بررسی ردیف‌ها برای پیدا کردن Total
                        if len(records) >= 3:
                            total_row = records[2]
                        elif len(records) >= 2:
                            total_row = records[1]
                        elif len(records) >= 1:
                            total_row = records[0]
                        
                        if total_row:
                            all_tables_data["Drilling"]["DrillingTotal"] = {
                                "ص": convert_to_int(total_row.get(4, "") if 4 in total_row else ""),
                                "ن": convert_to_int(total_row.get(3, "") if 3 in total_row else ""),
                                "ش": convert_to_int(total_row.get(2, "") if 2 in total_row else ""),
                                "پ": convert_to_int(total_row.get(1, "") if 1 in total_row else ""),
                                "خ": convert_to_int(total_row.get(0, "") if 0 in total_row else "")
                            }
                        else:
                            all_tables_data["Drilling"]["DrillingTotal"] = {}
                    # اگر مربوط به Foods است
                    elif sheet_name == "Foods":
                        foods_data = convert_foods_to_structured(df)
                        all_tables_data["Foods"] = foods_data
                    # اگر مربوط به Total است
                    elif sheet_name == "Total":
                        total_data = convert_total_to_structured(df)
                        all_tables_data["total"] = total_data
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
    
    # اطمینان از وجود Herasat
    if "Herasat" not in all_tables_data:
        all_tables_data["Herasat"] = {}
    
    # اطمینان از وجود شیفت‌های Herasat به صورت لیست خالی اگر وجود نداشتند
    if "ShiftA" not in all_tables_data["Herasat"]:
        all_tables_data["Herasat"]["ShiftA"] = []
    if "ShiftB" not in all_tables_data["Herasat"]:
        all_tables_data["Herasat"]["ShiftB"] = []
    if "ShiftC" not in all_tables_data["Herasat"]:
        all_tables_data["Herasat"]["ShiftC"] = []
    if "ShiftD" not in all_tables_data["Herasat"]:
        all_tables_data["Herasat"]["ShiftD"] = []
    if "ShiftTotalHerasat" not in all_tables_data["Herasat"]:
        all_tables_data["Herasat"]["ShiftTotalHerasat"] = {}
    
    # اطمینان از وجود Ordogahi
    if "Ordogahi" not in all_tables_data:
        all_tables_data["Ordogahi"] = {}
    
    # اطمینان از وجود شیفت‌های Ordogahi به صورت لیست خالی اگر وجود نداشتند
    if "ShiftA" not in all_tables_data["Ordogahi"]:
        all_tables_data["Ordogahi"]["ShiftA"] = []
    if "ShiftB" not in all_tables_data["Ordogahi"]:
        all_tables_data["Ordogahi"]["ShiftB"] = []
    if "ShiftC" not in all_tables_data["Ordogahi"]:
        all_tables_data["Ordogahi"]["ShiftC"] = []
    if "ShiftD" not in all_tables_data["Ordogahi"]:
        all_tables_data["Ordogahi"]["ShiftD"] = []
    if "ShiftTotalOrdogahi" not in all_tables_data["Ordogahi"]:
        all_tables_data["Ordogahi"]["ShiftTotalOrdogahi"] = {}
    
    # اطمینان از وجود employer
    if "employer" not in all_tables_data:
        all_tables_data["employer"] = {}
    
    # اطمینان از وجود بخش‌های employer به صورت خالی اگر وجود نداشتند
    if "EmployerPage3" not in all_tables_data["employer"]:
        all_tables_data["employer"]["EmployerPage3"] = []
    if "EmployerTotal" not in all_tables_data["employer"]:
        all_tables_data["employer"]["EmployerTotal"] = {}
    if "EmployerSupervisor" not in all_tables_data["employer"]:
        all_tables_data["employer"]["EmployerSupervisor"] = {}
    
    # اطمینان از وجود Drilling
    if "Drilling" not in all_tables_data:
        all_tables_data["Drilling"] = {}
    
    # اطمینان از وجود بخش‌های Drilling به صورت خالی اگر وجود نداشتند
    if "DrillingPage4" not in all_tables_data["Drilling"]:
        all_tables_data["Drilling"]["DrillingPage4"] = []
    if "DrillingTotal" not in all_tables_data["Drilling"]:
        all_tables_data["Drilling"]["DrillingTotal"] = {}
    
    # اطمینان از وجود Foods
    if "Foods" not in all_tables_data:
        all_tables_data["Foods"] = {
            "صبحانه": "",
            "ناهار": "",
            "شام": "",
            "پس شام": "",
            "توضیحات": ""
        }
    
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