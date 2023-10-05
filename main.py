import openpyxl
from ics import Calendar, Event

# 读取Excel文件并选择Sheet1
workbook = openpyxl.load_workbook('birth2.xlsx')
sheet = workbook["Sheet1"]

# 创建一个空的日历对象
cal = Calendar()
# 跳过标题行，并遍历所有行
for row in sheet.iter_rows(min_row=2, max_col=9, values_only=True):
    person_count, month, full_date, birth_year, age_2023, age_2024, name, gender, phone = row
    if not birth_year:  # 检查birth_year是否为None或空
        print(f"Skipping row due to missing birth year: {row}")
        continue

    birth_year = int(birth_year)

    birth_year = int(birth_year)

    # 检查必要的字段是否为空
    if not all([month, full_date, name, gender, phone]):
        print(f"Skipping row due to missing values: {row}")
        continue

    # 处理整数格式的月和日
    month = str(month) if isinstance(month, int) else month
    if isinstance(full_date, int):
        full_date = openpyxl.utils.datetime.from_excel(full_date).strftime("%m月%d日")

    month_num = int(month.split("月")[0])
    day_num = int(full_date.split("月")[1].split("日")[0])

    # 创建从2023年到2026年的生日事件
    for year in range(2023, 2027):
        if year == 2023 and month_num < 10:
            continue  # 如果在2023年并且月份小于10月，则跳过
        if year == 2026 and month_num > 6:
            break  # 如果在2026年并且月份超过6月，则结束

        event = Event()
        event.name = f"{name}的生日"
        event.begin = f'{year}-{month_num:02}-{day_num:02}'
        age = year - birth_year
        event.description = f"{name} ({gender}, {phone}) 的 {age} 岁生日,记得去班群发生日祝福(!"

        # 将事件添加到日历
        cal.events.add(event)

# 将日历保存为ICS文件
with open('birthdays.ics', 'w', encoding='utf-8') as file:
    file.writelines(cal)

print("转换完成，已保存为birthdays.ics")
