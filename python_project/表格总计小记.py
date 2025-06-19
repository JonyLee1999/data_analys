import pandas as pd
import numpy as np
from IPython.display import display

def create_overdue_pivot_table(data):
    """
    创建超期状态透视表，包含小计和总计
    
    参数:
        data: 包含原始数据的DataFrame，需要有以下列:
            - bg: BG分组
            - ou: OU分组
            - overdue_days: 超期天数
            - invoice_amount: 发票金额
            
    返回:
        格式化的透视表DataFrame
    """
    # 确保 invoice_amount 是数值类型
    data['invoice_amount'] = pd.to_numeric(data['invoice_amount'], errors='coerce')
    data['overdue_days'] = pd.to_numeric(data['overdue_days'], errors='coerce')
    
    # 添加超期状态分类
    conditions = [
        (data['overdue_days'] < 0),
        (data['overdue_days'].between(0, 29)),
        (data['overdue_days'].between(30, 59)),
        (data['overdue_days'].between(60, 89)),
        (data['overdue_days'] >= 90)
    ]
    choices = ['not_overdue', 'overdue_0_29', 'overdue_30_59', 'overdue_60_89', 'overdue_90_plus']
    data['overdue_status'] = np.select(conditions, choices, default='unknown')
    
    # 创建透视表
    pivot = pd.pivot_table(
        data,
        values='invoice_amount',
        index=['bg', 'ou'],
        columns='overdue_status',
        aggfunc='sum',
        fill_value=0,
        margins=True,
        margins_name='总计'
    )
    
    # 确保所有数值列都是数值类型
    numeric_cols = ['not_overdue', 'overdue_0_29', 'overdue_30_59', 'overdue_60_89', 'overdue_90_plus']
    for col in numeric_cols:
        if col in pivot.columns:
            pivot[col] = pd.to_numeric(pivot[col], errors='coerce').fillna(0)
    
    # 计算行总计和超期占比
    pivot['row_total'] = pivot[numeric_cols].sum(axis=1)
    overdue_cols = ['overdue_0_29', 'overdue_30_59', 'overdue_60_89', 'overdue_90_plus']
    pivot['overdue_percentage'] = (pivot[overdue_cols].sum(axis=1) / pivot['row_total'] * 100).round(2)
    
    # 添加BG小计
    bg_subtotals = pd.pivot_table(
        data,
        values='invoice_amount',
        index=['bg'],
        columns='overdue_status',
        aggfunc='sum',
        fill_value=0
    )
    
    # 确保小计中的所有列都是数值类型
    for col in numeric_cols:
        if col in bg_subtotals.columns:
            bg_subtotals[col] = pd.to_numeric(bg_subtotals[col], errors='coerce').fillna(0)
    
    bg_subtotals['row_total'] = bg_subtotals[numeric_cols].sum(axis=1)
    bg_subtotals['overdue_percentage'] = (bg_subtotals[overdue_cols].sum(axis=1) / bg_subtotals['row_total'] * 100).round(2)
    bg_subtotals.insert(0, 'ou', '小计')
    
    # 合并所有数据
    pivot_reset = pivot.reset_index()
    bg_subtotals_reset = bg_subtotals.reset_index()
    
    # 确保列名匹配
    pivot_reset.columns = pivot_reset.columns.astype(str)
    bg_subtotals_reset.columns = bg_subtotals_reset.columns.astype(str)
    
    combined = pd.concat([pivot_reset, bg_subtotals_reset], ignore_index=True)
    
    # 重新排列列顺序
    cols_order = ['bg', 'ou', 'not_overdue', 'overdue_0_29', 'overdue_30_59', 'overdue_60_89', 'overdue_90_plus', 'row_total', 'overdue_percentage']
    combined = combined[cols_order]
    
    # 格式化输出
    styled = combined.style.format({
        'not_overdue': '{:,.2f}',
        'overdue_0_29': '{:,.2f}',
        'overdue_30_59': '{:,.2f}',
        'overdue_60_89': '{:,.2f}',
        'overdue_90_plus': '{:,.2f}',
        'row_total': '{:,.2f}',
        'overdue_percentage': '{:.2f}%'
    })
    
    # 添加样式
    styled = styled.applymap(
        lambda x: 'font-weight: bold' if isinstance(x, str) and ('小计' in str(x) or '总计' in str(x)) else ''
    )
    
    return styled

# 示例数据
np.random.seed(42)
data = pd.DataFrame({
    'bg': np.random.choice(['BG1', 'BG2', 'BG3'], 100),
    'ou': np.random.choice(['OU1', 'OU2', 'OU3'], 100),
    'overdue_days': np.random.choice([-10, 15, 45, 75, 100], 100),
    'invoice_amount': np.random.uniform(100, 10000, 100).round(2)
})

# 创建并显示透视表
pivot_table = create_overdue_pivot_table(data)
display(pivot_table)