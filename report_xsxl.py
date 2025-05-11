import xlsxwriter

# 创建一个新的 Excel 文件
workbook = xlsxwriter.Workbook('example.xlsx')

# 添加一个工作表
worksheet = workbook.add_worksheet('Sheet1')

# 写入数据
data = [
    ["姓名", "年龄", "城市"],
    ["张三", 25, "北京"],
    ["李四", 30, "上海"],
    ["王五", 28, "广州"],
]

# 写入行数据（从 A1 开始）
for row_num, row_data in enumerate(data):
    worksheet.write_row(row_num, 0, row_data)

# 可选：设置标题行加粗
bold_format = workbook.add_format({'bold': True})
worksheet.set_row(0, None, bold_format)

# 关闭文件（保存）
workbook.close()
print("Excel 文件已生成：example.xlsx")