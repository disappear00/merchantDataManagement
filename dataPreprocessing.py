import tkinter as tk
from tkinter import messagebox
import pandas as pd
from tkinter import filedialog
import os


def select_file(title):
    """允许用户选择一个文件并返回文件路径"""
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    file_path = filedialog.askopenfilename(title=title, filetypes=[("Excel files", "*.xlsx")])
    return file_path


def select_output_folder():
    """允许用户选择输出文件夹"""
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    output_folder = filedialog.askdirectory(title="选择输出文件夹")
    return output_folder


def main():
    # 创建主界面
    root = tk.Tk()
    root.title("文件处理程序")

    # 创建用于显示信息的Text组件
    text_widget = tk.Text(root, height=20, width=80)
    text_widget.pack()

    def append_text(message):
        text_widget.insert(tk.END, message + "\n")
        text_widget.see(tk.END)

    append_text("请选择第一个包含商家ID的文件(霸州乡镇商家明细)...")
    file1_path = select_file("选择第一个文件")
    if not file1_path:
        messagebox.showerror("错误", "未选择文件，程序退出")
        root.destroy()
        return

    append_text("请选择第二个包含商家ID和外卖组织结构的文件(海豚_合作商商家数据)...")
    file2_path = select_file("选择第二个文件")
    if not file2_path:
        messagebox.showerror("错误", "未选择文件，程序退出")
        root.destroy()
        return

    append_text("请选择第三个包含商家ID和一系列数据的文件(淮安卓美网络科技有限公司)...")
    file3_path = select_file("选择第三个文件")
    if not file3_path:
        messagebox.showerror("错误", "未选择文件，程序退出")
        root.destroy()
        return

    append_text("请选择输出文件夹(最好统一存在一个文件夹(用于Test2))")
    output_folder = select_output_folder()
    if not output_folder:
        messagebox.showerror("错误", "未选择输出文件夹，程序退出")
        root.destroy()
        return

    append_text(f"已选择输出文件夹：{output_folder}")

    # 固定列名
    merchant_id_col1 = "商家ID"  # 文件1中的商家ID列名
    merchant_id_col2 = "商家ID"  # 文件2中的商家ID列名
    org_structure_col = "外卖组织结构"  # 文件2中的外卖组织结构列名
    merchant_id_col3 = "商家ID"  # 文件3中的商家ID列名

    # 读取三个文件
    try:
        df1 = pd.read_excel(file1_path)
        df2 = pd.read_excel(file2_path)
        df3 = pd.read_excel(file3_path)

        append_text(f"\n文件1包含 {len(df1)} 行数据")
        append_text(f"文件2包含 {len(df2)} 行数据")
        append_text(f"文件3包含 {len(df3)} 行数据")

        # 检查输入的列名是否存在
        if merchant_id_col1 not in df1.columns:
            messagebox.showerror("错误", f"列 '{merchant_id_col1}' 在文件1中不存在")
            root.destroy()
            return
        if merchant_id_col2 not in df2.columns or org_structure_col not in df2.columns:
            messagebox.showerror("错误", "列名在文件2中不存在")
            root.destroy()
            return
        if merchant_id_col3 not in df3.columns:
            messagebox.showerror("错误", f"列 '{merchant_id_col3}' 在文件3中不存在")
            root.destroy()
            return

        # 第一步：根据文件1中的商家ID，更新文件2中对应行的外卖组织结构为"霸州三组"
        # 获取文件1中的商家ID列表
        merchants_in_file1 = set(df1[merchant_id_col1].astype(str))

        # 更新文件2中匹配的行
        mask = df2[merchant_id_col2].astype(str).isin(merchants_in_file1)
        original_structure_count = df2[mask][org_structure_col].value_counts().to_dict()
        append_text(f"\n更新前文件2中匹配商家ID的外卖组织结构分布: {original_structure_count}")

        # 统计要更改的行数
        rows_to_update = mask.sum()
        append_text(f"将更改 {rows_to_update} 行数据的外卖组织结构为'霸州三组'")

        # 执行更新
        df2.loc[mask, org_structure_col] = "霸州三组"

        # 第二步：根据更新后的文件2，向文件3添加外卖组织结构列
        # 创建商家ID到外卖组织结构的映射
        merchant_to_org = dict(zip(df2[merchant_id_col2].astype(str), df2[org_structure_col]))

        # 向文件3添加外卖组织结构列
        df3[org_structure_col] = df3[merchant_id_col3].astype(str).map(merchant_to_org)

        # 检查文件3中有多少行没有对应的组织结构
        missing_org_mask = df3[org_structure_col].isna()
        missing_org_count = missing_org_mask.sum()
        append_text(f"\n文件3中有 {missing_org_count} 行数据没有对应的外卖组织结构")

        # 打印前5个缺失组织结构的商家ID
        if missing_org_count > 0:
            missing_ids = df3.loc[missing_org_mask, merchant_id_col3].astype(str).tolist()
            append_text("\n前5个缺失外卖组织结构的商家ID:")
            for i, id_value in enumerate(missing_ids[:5], 1):
                append_text(f"{i}. {id_value}")

            # 额外检查：这些ID是否在文件2中存在
            missing_ids_set = set(missing_ids)
            file2_ids_set = set(df2[merchant_id_col2].astype(str))
            ids_not_in_file2 = missing_ids_set - file2_ids_set
            if ids_not_in_file2:
                append_text(f"\n在文件2中不存在的商家ID数量: {len(ids_not_in_file2)}")
                append_text("前5个在文件2中不存在的ID样例:")
                for i, id_value in enumerate(list(ids_not_in_file2)[:5], 1):
                    append_text(f"{i}. {id_value}")

            # 为缺失的组织结构设置默认值
            default_org = "未知组织结构"
            df3[org_structure_col].fillna(default_org, inplace=True)
            append_text(f"\n已将缺失的外卖组织结构设置为 '{default_org}'")

        # 统计结果
        org_counts = df3[org_structure_col].value_counts().to_dict()
        append_text(f"\n文件3中外卖组织结构的分布: {org_counts}")

        # 第三步：按外卖组织结构分组并保存为不同的Excel文件
        base_name = os.path.splitext(os.path.basename(file3_path))[0]

        # 记录处理的行数
        total_processed = 0

        # 按组织结构分组
        for org_name, group_data in df3.groupby(org_structure_col):
            output_name = f"{base_name}_{org_name}.xlsx"
            output_path = os.path.join(output_folder, output_name)
            group_data.to_excel(output_path, index=False)
            append_text(f"已保存 {len(group_data)} 行数据到 {output_name}")
            total_processed += len(group_data)

        # 验证所有数据都被处理
        append_text(f"\n总共处理了 {total_processed} 行数据，原始文件3有 {len(df3)} 行数据")
        if total_processed != len(df3):
            append_text(f"警告：处理的数据行数与原始文件不一致，差异为 {len(df3) - total_processed} 行")
        else:
            append_text("验证成功：所有数据都已正确处理并保存到各分组文件中")

        # 保存更新后的文件2
        updated_file2_path = os.path.join(output_folder, "updated_" + os.path.basename(file2_path))
        df2.to_excel(updated_file2_path, index=False)
        append_text(f"\n已保存更新后的文件2到 {updated_file2_path}")

        # 保存完整的处理后的文件3（包含组织结构列）
        updated_file3_path = os.path.join(output_folder, "updated_" + os.path.basename(file3_path))
        df3.to_excel(updated_file3_path, index=False)
        append_text(f"已保存更新后的文件3到 {updated_file3_path}")

        # 保存缺失组织结构的商家ID到单独文件
        if missing_org_count > 0:
            missing_ids_df = pd.DataFrame({merchant_id_col3: missing_ids})
            missing_ids_path = os.path.join(output_folder, f"{base_name}_缺失组织结构ID列表.xlsx")
            missing_ids_df.to_excel(missing_ids_path, index=False)
            append_text(f"已保存缺失组织结构的商家ID列表到 {missing_ids_path}")

        append_text("\n处理完成！")

    except Exception as e:
        messagebox.showerror("错误", f"处理过程中发生错误: {e}")
        root.destroy()

    root.mainloop()


if __name__ == "__main__":
    main()