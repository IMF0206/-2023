import pandas as pd

# 读取Excel文件
excel_file_path = 'D:\\highq\\October.xlsx'  # 请替换为你的Excel文件路径
df = pd.read_excel(excel_file_path, dtype={'借书时间': str, '还书时间': str})

# 遍历表格，确保还书时间不存在时使用虚假日期存储，并检查还书时间是否大于借书时间
for index, row in df.iterrows():
    borrow_date = row['借书时间']
    return_date = row['还书时间']
    borrower = row['借阅人']

    # 如果借书时间缺失，输出序号并用虚假日期 '0.0' 存储
    if pd.isna(borrow_date):
        print(f"错误：序号 {row['序号']} 的借书时间缺失。")
        df.at[index, '借书时间'] = '0.0'
    elif not borrow_date.startswith('10.'):
        continue  # 如果不是10月份的借书时间，跳过处理

    borrow_month, borrow_day = map(int, borrow_date.split('.'))

    # 如果还书时间存在
    if pd.notna(return_date) and return_date != '13.32':
        return_month, return_day = map(int, return_date.split('.'))
        
        # 检查还书时间是否大于借书时间
        if (return_month < borrow_month) or (return_month == borrow_month and return_day <= borrow_day):
            print(f"错误：序号 {row['序号']} 的还书时间不合法（小于等于借书时间）。")
            
        # 统计借书日期等于还书日期的次数
        if borrow_date == return_date:
            df.at[index, '日期相等'] = 1.0  # 修改此处，将 1.0 改为 1 以避免警告
    # 如果还书时间不存在，使用虚假日期存储，并输出提示信息
    elif pd.isna(return_date):
        df.at[index, '还书时间'] = '13.32'
        print(f"提示：序号 {row['序号']} 的还书时间未填写，已使用虚拟日期 '13.32' 存储.")

    # 如果借阅人缺失，输出序号并用 "空缺" 替代
    if pd.isna(borrower):
        print(f"错误：序号 {row['序号']} 的借阅人缺失。")
        df.at[index, '借阅人'] = '空缺'

# 给 '日期相等' 列的 NaN 值填充为 0
df['日期相等'].fillna(0, inplace=True)

# 统计借书日期等于还书日期的总次数
equal_dates_count = int(df['日期相等'].sum())  # 显式地转换为 int 类型

# 统计总次数
total_count = len(df)

# 如果借书日期等于还书日期的次数超过总次数的10%，打印这些序号
if equal_dates_count > total_count * 0.1:
    print(f"借书日期等于还书日期的次数超过总次数的10%，以下是这些序号：")
    print(df[df['日期相等'] == 1]['序号'].to_string(index=False))

# 输出统计信息
print(f"借书日期等于还书日期的总次数：{equal_dates_count}")
print(f"总次数：{total_count}")
