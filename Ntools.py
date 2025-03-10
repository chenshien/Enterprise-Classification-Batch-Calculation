import os
import sys
import pandas as pd
from datetime import datetime
import re
from datetime import datetime
from industries_config import get_industries


def check_date():
    current_date = datetime.now()
    target_date1 = datetime(2025, 5, 1)
    target_date2 = datetime(2025, 2, 6)

    if current_date > target_date1 or current_date < target_date2:
        print("程序版本校验出错V1，程序退出！！")
        print("程序版本校验出错V2，程序退出！！")
        print("程序版本校验出错V3，程序退出！！")
        input("按任意键退出...")
        sys.exit(1)


def load_data_from_excel(file_path, shujudanwei):
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"文件 {file_path} 不存在")

    if not os.path.isfile(file_path):
        raise ValueError(f"{file_path} 不是一个有效的文件路径")

    try:
        if file_path.endswith('.xls'):
            data = pd.read_excel(file_path, sheet_name=0, engine='xlrd')
        elif file_path.endswith('.xlsx'):
            data = pd.read_excel(file_path, sheet_name=0)
        else:
            raise ValueError("仅支持 .xls 和 .xlsx 文件")

        data.fillna(value='', inplace=True)

        # 清理行业分类代码列
        code_columns = ['行业分类代码', '行业分类二级代码', '行业分类三级代码']
        for col in code_columns:
            if col not in data.columns:
                raise KeyError(f"列 '{col}' 不存在于输入文件中")
            data[col] = data[col].apply(lambda x: re.sub(r'[A-Z0-9\s]', '', str(x)))

        # 处理数值列
        if shujudanwei.upper() in ['Y', 'WY']:
            numeric_columns = {
                '全年营业收入': True,
                '资产总额': True,
                '从业人数': False
            }

            for col, need_conversion in numeric_columns.items():
                if col in data.columns:
                    data[col] = pd.to_numeric(data[col], errors='coerce')
                    if need_conversion and shujudanwei.upper() == 'Y':
                        data[col] = data[col] / 10000
        else:
            raise ValueError("数据单位输入错误，请输入 'Y' 或 'WY'")

        return data

    except Exception as e:
        raise Exception(f"加载数据失败: {str(e)}")


def update_company_category(data):
    # 创建数据副本并预处理数据
    data = data.copy()

    # 转换数值型字段
    numeric_fields = ['全年营业收入', '资产总额', '从业人数']
    for field in numeric_fields:
        if field in data.columns:
            data[field] = pd.to_numeric(data[field], errors='coerce')

    # 转换字符串字段，去除空格和特殊字符
    string_fields = ['行业分类代码', '行业分类二级代码', '行业分类三级代码']
    for field in string_fields:
        if field in data.columns:
            data[field] = data[field].astype(str).str.strip()

    # 初始化排查结果列
    data['排查结果'] = "未匹配"

    industries = get_industries(data)
    print(f"\n获取到的行业配置数量: {len(industries)}")

    # 企业规模等级顺序，用于下调处理
    scale_levels = ["大型企业", "中型企业", "小型企业", "微型企业"]

    for industry, match_level, conditions in industries:
        if '行业分类代码' not in data.columns:
            continue

        # 根据match_level选择匹配字段
        if match_level == 1:
            industry_field = '行业分类代码'
        elif match_level == 2:
            industry_field = '行业分类二级代码'
        elif match_level == 3:
            industry_field = '行业分类三级代码'
        else:
            continue

        # 创建行业匹配掩码
        industry_mask = data[industry_field].str.contains(industry, na=False)
        matched_count = industry_mask.sum()
        print(f"\n行业: {industry}, 匹配级别: {match_level}, 匹配到的行数: {matched_count}")

        # 对匹配到的行进行处理
        matched_rows = data[industry_mask]
        if matched_rows.empty:
            continue

        for index, row in matched_rows.iterrows():
            row_data = pd.DataFrame([row])
            matched_categories = []  # 存储所有匹配的类别

            # 检查所有可能匹配的类别
            for criteria, category in conditions:
                try:
                    if category == "微型企业":
                        # 微型企业只要满足任一条件
                        conditions_results = [cond(row_data).iloc[0] for cond in criteria]
                        if any(conditions_results):
                            matched_categories.append(category)
                    else:
                        # 其他企业需要满足所有条件
                        conditions_results = [cond(row_data).iloc[0] for cond in criteria]
                        if all(conditions_results):
                            matched_categories.append(category)
                except Exception as e:
                    print(f"条件评估错误 - {industry} - {category}: {e}")
                    continue

            # 处理匹配结果
            if matched_categories:
                # 找出匹配类别中最高等级的索引
                highest_level_idx = min(scale_levels.index(cat) for cat in matched_categories)

                # 如果有多个匹配且不是最低级别，则下调一级
                if len(matched_categories) > 1 and highest_level_idx < len(scale_levels) - 1:
                    selected_category = scale_levels[highest_level_idx + 1]
                else:
                    selected_category = scale_levels[highest_level_idx]

                # 只有当新的匹配结果优于现有结果时才更新
                current_result = data.loc[index, '排查结果']
                if current_result == "未匹配" or (
                        current_result in scale_levels and
                        scale_levels.index(selected_category) < scale_levels.index(current_result)
                ):
                    data.loc[index, '排查结果'] = selected_category

    # 打印最终结果统计
    result_counts = data['排查结果'].value_counts()
    print("\n最终匹配结果统计:")
    print(result_counts)

    return data


def export_data_to_excel(data, output_file):
    try:
        data.to_excel(output_file, index=False)
        print(f"\n导出成功：{output_file}\n")
    except Exception as e:
        print(f"\n导出失败：{e}\n")


def main(file_path, output_file, shujudanwei):
    try:
        if os.path.exists(output_file):
            os.remove(output_file)  # 删除已存在的目标文件
            print(f"\n已删除旧文件：{output_file}\n")

        data = load_data_from_excel(file_path, shujudanwei)
        try:
            updated_data = update_company_category(data)
            export_data_to_excel(updated_data, output_file)
        except FileNotFoundError as e:
            print(f"错误: {e}")
            print("\n请确保 industries_config.ini 配置文件与程序在同一目录下\n")
            return
        except Exception as e:
            print(f"\n处理数据时出错: {e}")
            return

    except Exception as e:
        print(f"执行出错: {e}")
    finally:
        print("\n如果所有企业都显示为 未匹配，请检查：")
        print("1. industries_config.ini 文件是否存在且内容正确")
        print("2. 行业分类代码是否与配置文件中的行业名称匹配")
        print("3. 数据格式是否正确（数值字段不能为空或非数字）\n")


if __name__ == "__main__":
    check_date()
    if len(sys.argv) != 3:
        print("使用方法: Ntools.exe <数据文件> <数据单位Y/WY>\n")
        input("按任意键退出...")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = os.path.join(os.path.dirname(input_file), '排查结果.xlsx')
    starttime = datetime.now()
    shujudanwei = sys.argv[2]
    main(input_file, output_file, shujudanwei)
    endtime = datetime.now()
    print("执行时间：", endtime - starttime)
    input("按任意键退出...")
    sys.exit(0)
