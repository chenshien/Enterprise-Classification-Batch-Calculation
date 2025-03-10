import configparser
import os
import sys
import pandas as pd

def parse_condition(condition_str):
    """解析条件字符串为条件列表"""
    conditions = []
    for cond in condition_str.split(','):
        field, op, value = _parse_single_condition(cond.strip())
        def make_condition(f=field, o=op, v=float(value)):
            def condition(data):
                try:
                    # 数据已经在update_company_category中预处理过，直接使用
                    return _evaluate_condition(data[f], o, v)
                except Exception as e:
                    print(f"条件评估错误 - 字段:{f}, 操作:{o}, 值:{v}: {e}")
                    return pd.Series([False])
            return condition
        conditions.append(make_condition())
    return conditions

def _parse_single_condition(cond_str):
    """解析单个条件"""
    if '>=' in cond_str:
        field, value = cond_str.split('>=')
        return field.strip(), '>=', float(value)
    elif '<=' in cond_str:
        field, value = cond_str.split('<=')
        return field.strip(), '<=', float(value)
    elif '>' in cond_str:
        field, value = cond_str.split('>')
        return field.strip(), '>', float(value)
    elif '<' in cond_str:
        field, value = cond_str.split('<')
        return field.strip(), '<', float(value)
    raise ValueError(f"不支持的条件格式: {cond_str}")

def _evaluate_condition(series, op, value):
    """评估条件"""
    try:
        if op == '>=':
            return series >= value
        elif op == '<=':
            return series <= value
        elif op == '>':
            return series > value
        elif op == '<':
            return series < value
        raise ValueError(f"不支持的操作符: {op}")
    except Exception as e:
        print(f"条件评估错误: {e}")
        return pd.Series([False])

def get_industries(data):
    """从配置文件读取行业配置"""
    config = configparser.ConfigParser()
    
    # 获取exe所在目录
    if getattr(sys, 'frozen', False):
        # 如果是打包后的exe
        application_path = os.path.dirname(sys.executable)
    else:
        # 如果是直接运行python脚本
        application_path = os.path.dirname(os.path.abspath(__file__))
    
    config_path = os.path.join(application_path, 'industries_config.ini')
    
    # 添加调试信息
    print(f"尝试读取配置文件: {config_path}")
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"配置文件不存在: {config_path}")
        
    # 使用utf-8编码读取配置文件
    try:
        config.read(config_path, encoding='utf-8')
        print(f"成功读取配置文件，包含的section数量: {len(config.sections())}")
        print("配置文件中的sections:", config.sections())
    except Exception as e:
        print(f"读取配置文件失败: {e}")
        raise
    
    industries = []
    for industry in config.sections():
        try:
            match_level = int(config[industry]['match_level'])
            conditions = []
            
            # 跳过match_level配置项
            for category, condition_str in config[industry].items():
                if category == 'match_level':
                    continue
                conditions.append((parse_condition(condition_str), category))
            
            industries.append((industry, match_level, conditions))
            print(f"成功解析行业配置: {industry}, 匹配级别: {match_level}, 条件数量: {len(conditions)}")
            
        except Exception as e:
            print(f"解析行业 {industry} 配置失败: {e}")
            continue
    
    print(f"总共解析到 {len(industries)} 个有效行业配置")
    return industries 