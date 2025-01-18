import pandas as pd

# 1. مسیر فایل‌ها
file1_path = "C:\\Users\\Bahar\\Desktop\\code_desc_old.xlsx"  # مسیر فایل اول
file2_path = "C:\\Users\\Bahar\\Desktop\\Stage.xlsx"  # مسیر فایل دوم
output_path = "C:\\Users\\Bahar\\Desktop\\Stage_missmatches.xlsx"  # مسیر فایل خروجی
# 2. خواندن فایل‌ها
df1 = pd.read_excel(file1_path)  # فایل اول
df2 = pd.read_excel(file2_path)  # فایل دوم

# 3. استخراج ستون‌ها
file1_data = df1[['DetailCode', 'DescDetailCode']]  # ستون‌های فایل اول
file2_data = df2[['DetailCode', 'DescDetailCode']]  # ستون‌های فایل دوم

# 4. تبدیل فایل دوم به دیکشنری برای مقایسه سریع
file2_dict = dict(zip(file2_data['DetailCode'], file2_data['DescDetailCode']))

# 5. پیدا کردن مغایرت‌ها
mismatches = []

for _, row in file1_data.iterrows():
    detail_code = row['DetailCode']
    desc_detail_code = row['DescDetailCode']
    
    if detail_code not in file2_dict:
        # اگر DetailCode در فایل دوم وجود نداشت
        mismatches.append({
            "DetailCode": detail_code,
            "DescDetailCode_File1": desc_detail_code,
            "Issue": "Not Found in File 2",
            "DescDetailCode_File2": None
        })
    elif file2_dict[detail_code] != desc_detail_code:
        # اگر DescDetailCode با فایل دوم متفاوت بود
        mismatches.append({
            "DetailCode": detail_code,
            "DescDetailCode_File1": desc_detail_code,
            "Issue": "Mismatch",
            "DescDetailCode_File2": file2_dict[detail_code]
        })

# 6. تبدیل مغایرت‌ها به DataFrame
mismatches_df = pd.DataFrame(mismatches)

# 7. ذخیره به اکسل
mismatches_df.to_excel(output_path, index=False)

print(f"مغایرت‌ها در فایل اکسل ذخیره شدند: {output_path}")