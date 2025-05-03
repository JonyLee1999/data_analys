import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime, timedelta

# 模拟获取销售数据
def get_sales_data():
    data = {
        'date': pd.date_range(start="2025-04-20", end="2025-05-01"),
        'sales': [100, 150, 200, 250, 120, 130, 240, 180, 220, 300, 400, 350]
    }
    return pd.DataFrame(data)

# 计算指标
def calculate_metrics(df):
    df['daily_growth'] = df['sales'].pct_change()  # 日增长率
    df['cumulative_sales'] = df['sales'].cumsum()  # 累计销售额
    return df

# 生成报表
def generate_report(df, filename="sales_report.xlsx"):
    # 保存为 Excel
    df.to_excel(filename, index=False)
    print(f"报表已生成：{filename}")

    # 可视化
    plt.figure(figsize=(10, 6))
    plt.plot(df['date'], df['sales'], label="Daily Sales")
    plt.title("Daily Sales Report")
    plt.xlabel("Date")
    plt.ylabel("Sales")
    plt.legend()
    plt.savefig("sales_report.png")
    print("图表已生成：sales_report.png")

# 主函数
def main():
    sales_data = get_sales_data()
    sales_data = calculate_metrics(sales_data)
    generate_report(sales_data)

if __name__ == "__main__":
    main()