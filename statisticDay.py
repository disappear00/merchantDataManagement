import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
from datetime import datetime

def select_input_folder():
    """选择包含分组文件的文件夹"""
    root = tk.Tk()
    root.withdraw()
    folder_path = filedialog.askdirectory(title="选择包含外卖组织分组文件的文件夹")
    return folder_path

def select_output_folder():
    """选择输出文件夹"""
    root = tk.Tk()
    root.withdraw()
    folder_path = filedialog.askdirectory(title="选择输出文件夹")
    return folder_path

def get_month_from_user():
    """获取用户输入的月份"""
    root = tk.Tk()
    root.withdraw()
    month = simpledialog.askstring("输入月份", "请输入月份（格式：3月）：")
    return month

def get_date_from_user(month):
    """获取用户输入的日期"""
    root = tk.Tk()
    root.withdraw()
    date = simpledialog.askstring("输入日期", f"请输入{month}月的日期（格式：3）：")
    return date

def find_column(df, target_column):
    """模糊匹配列名"""
    for col in df.columns:
        if target_column in col:
            return col
    return None

def create_monthly_folder(output_folder, month):
    """创建月份文件夹"""
    year = datetime.now().year
    month_folder = os.path.join(output_folder, f"{year}年{month}月")
    
    if not os.path.exists(month_folder):
        os.makedirs(month_folder)
    
    return month_folder

def update_monthly_summary(month_folder, month, date, summary_data):
    """更新每月总表"""
    monthly_file = os.path.join(month_folder, f"{month}月汇总表.xlsx")
    formatted_date = f"{month}月{date}日"
    
    if os.path.exists(monthly_file):
        # 读取现有月度汇总表
        with pd.ExcelFile(monthly_file) as xls:
            monthly_df = pd.read_excel(xls, sheet_name=None)
        
        # 更新每个区域的数据
        for area, df in summary_data.items():
            if area in monthly_df:
                # 检查是否已存在该日期的数据
                existing_dates = monthly_df[area]['日期'].tolist()
                if formatted_date not in existing_dates:
                    # 添加新数据
                    new_row = summary_data[area].copy()
                    new_row['日期'] = formatted_date
                    monthly_df[area] = pd.concat([monthly_df[area], new_row], ignore_index=True)
                    # 按日期排序
                    monthly_df[area].sort_values(by='日期', inplace=True)
            else:
                # 添加新区域
                monthly_df[area] = summary_data[area]
                monthly_df[area]['日期'] = formatted_date
    else:
        # 创建新的月度汇总表
        monthly_df = {}
        for area, df in summary_data.items():
            monthly_df[area] = df.copy()
            monthly_df[area]['日期'] = formatted_date
    
    # 保存月度汇总表
    with pd.ExcelWriter(monthly_file, engine='openpyxl') as writer:
        for area, df in monthly_df.items():
            df.to_excel(writer, sheet_name=area, index=False)

def main():
    # 选择包含分组文件的文件夹
    input_folder = select_input_folder()
    if not input_folder:
        print("未选择输入文件夹，程序退出")
        return

    # 选择输出文件夹
    output_folder = select_output_folder()
    if not output_folder:
        print("未选择输出文件夹，程序退出")
        return

    # 获取用户输入的月份
    month = get_month_from_user()
    if not month:
        print("未输入月份，程序退出")
        return

    # 获取用户输入的日期
    date = get_date_from_user(month)
    if not date:
        print("未输入日期，程序退出")
        return

    # 创建月份文件夹
    month_folder = create_monthly_folder(output_folder, month)

    # 定义外卖组织和对应区域
    organization_mapping = {
        '高碑店一组': '高碑店',
        '高碑店二组': '白沟',
        '高碑店三组': '新城',
        '霸州一组': '霸州',
        '霸州二组': '胜芳',
        '霸州三组': '霸州乡镇'
    }

    # 定义列映射关系
    column_mapping = {
        '商业支持服务费(元)': '收商家服务费(元)',
        '企客履约服务费': None, # 有问题
        '一口价服务费(元)': '一口价服务费(元)',
        '配送费(元)': '用户配送费(元)',
        '活动款(元)': '活动款(元)',
        '竞价考核': '竞价考核(元)',
        '罚款(元)': '罚款(元)',
        '雇主险(元)': '雇主险(元)', # 有问题
        '非雇主责任险(元)': '非雇主责任险(元)',
        '邀新奖励支出(元)': '邀新奖励支出(元)',
        '省钱包售卖合作商承担': '省钱包售卖合作商承担',
        '省钱包售卖合作商承担-退款': '省钱包售卖合作商承担-退款',
        '二次配送费付合作商': '二次配送费付合作商',
        '二次配送费付合作商-退款': '二次配送费付合作商-退款',
        '合作商广告分成': '合作商广告分成',
        '合作商奖励': '合作商奖励',
        '合作商服务费': '合作商服务费',
        '合作商服务费退款': '合作商服务费退款',
        '关爱基金': '关爱基金',
        '商户服务费返还激励': '商户服务费返还激励',
        '省钱包售卖返还': '省钱包售卖返还',
        '合作商售后赔付费用': '合作商售后赔付费用',
        '合作商成本调账': '合作商成本调账',
        '春节服务费': '春节服务费',
        '省钱包核销美团承担': '省钱包核销美团承担',
        '拼好饭拼单宝': '拼好饭拼单宝',
        '经营权交易服务费': '经营权交易服务费'
    }

    # 查找所有xlsx文件
    xlsx_files = [f for f in os.listdir(input_folder) if f.endswith('.xlsx') and not f.startswith('updated_')]

    # 准备汇总的Excel文件
    output_path = os.path.join(month_folder, f"{month}月{date}日外卖组织服务费汇总.xlsx")
    
    # 用于存储所有区域的汇总数据，用于更新月度总表
    all_summary_data = {}
    
    # 使用ExcelWriter
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        for org_group, area in organization_mapping.items():
            # 查找对应的文件
            matching_files = [f for f in xlsx_files if org_group in f]
            
            if not matching_files:
                print(f"未找到 {org_group} 的文件，跳过")
                continue
            
            file_path = os.path.join(input_folder, matching_files[0])
            print(f"处理文件：{matching_files[0]}")
            
            # 读取文件
            df = pd.read_excel(file_path)
            
            # 检查并替换列名
            new_columns = {}
            for target, replacement in column_mapping.items():
                if replacement:  # 只处理有替换值的列
                    found_column = find_column(df, target)
                    if found_column:
                        new_columns[found_column] = replacement
            
            # 重命名列
            df = df.rename(columns=new_columns)
            
            # 创建汇总DataFrame
            summary_df = pd.DataFrame()
            summary_df['日期'] = [f"{month}月{date}日"]
            
            # 求和
            for col in new_columns.values():
                if col in df.columns:
                    summary_df[col] = [df[col].sum()]
            
            # 计算合计
            sum_columns = [col for col in summary_df.columns if col != '日期']
            summary_df['合计'] = summary_df[sum_columns].sum(axis=1)
            
            # 写入Excel
            summary_df.to_excel(writer, sheet_name=area, index=False)
            print(f"已写入 {area} 工作表")
            
            # 保存汇总数据用于更新月度总表
            all_summary_data[area] = summary_df

    print(f"\n当日汇总文件已生成：{output_path}")
    
    # 更新月度汇总表
    update_monthly_summary(month_folder, month, date, all_summary_data)
    print(f"月度汇总表已更新：{os.path.join(month_folder, f'{month}月汇总表.xlsx')}")

if __name__ == "__main__":
    main()