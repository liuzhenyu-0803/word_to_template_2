#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import json
import time
import sys
from collections import OrderedDict

# 添加项目根目录到系统路径
current_dir = os.path.dirname(os.path.abspath(__file__))
project_dir = os.path.dirname(current_dir)
if project_dir not in sys.path:
    sys.path.append(project_dir)

# 使用绝对导入
from models.model_manager import llm_manager

# ================ 基础工具函数 ================

def read_file_content(file_path):
    """读取文件内容"""
    with open(file_path, 'r', encoding='utf-8') as f:
        return f.read()

# ================ LLM 相关函数 ================

def prepare_system_message(prompt_path, key_path, content_path):
    """准备user消息
    
    Args:
        prompt_path: 提示词文件路径
        key_path: 关键信息描述文件路径
        content_path: 内容文件路径
    """
    # 读取prompt模板、key描述和内容文件
    prompt = read_file_content(prompt_path)
    key_description = read_file_content(key_path)
    content = read_file_content(content_path)
    
    # 替换占位符
    prompt = prompt.replace('placeholder_table_content', content)
    prompt = prompt.replace('placeholder_key_description', key_description)

    return prompt

# 注意：原call_llm函数已被移除，现在直接使用llm_manager.create_completion

def parse_llm_response(response_text):
    """
    从LLM响应中解析出JSON结果
    
    Args:
        response_text: LLM返回的文本内容
        
    Returns:
        解析后的结果列表，空列表([])表示有效的空数组结果
        
    Raises:
        ValueError: 当无法找到匹配结果标签或解析失败时抛出
    """
    # 使用正则表达式提取<match_result>标签内的内容
    import re
    pattern = r'<match_result>(.*?)</match_result>'
    matches = re.findall(pattern, response_text, re.DOTALL)
    
    if not matches:
        raise ValueError("未找到<match_result>标签")
    
    # 使用最后一个匹配结果（如果有多个）
    json_str = matches[-1].strip()

    json_str = f"[{json_str}]"
    
    # 解析JSON
    result_list = json.loads(json_str)
    
    # 验证每个结果项是否包含必需的字段
    valid_results = []
    for item in result_list:
        if all(k in item for k in ['key', 'value', 'matched_key']):
            valid_results.append(item)
        else:
            print(f"警告: 跳过缺少必填字段的项: {item}")
    
    print(f"成功提取和解析了 {len(valid_results)} 个匹配结果")
    return valid_results

# ================ 主要功能函数 ================

def match_table(key_description_path, table_content_path):
    """
    调用LLM分析表格内容，提取字段并与关键词进行匹配
    
    该函数会调用大语言模型对表格内容进行分析，提取字段名和值，
    并与关键词进行语义匹配。如果解析失败，会自动重试最多3次。
    
    Args:
        key_description_path: 关键词描述文件路径
        table_content_path: 表格内容文件路径
        
    Returns:
        list: 包含字段匹配结果的字典列表，每个字典包含key、
              value、matched_key
    """

    # 构造所需文件路径
    current_dir = os.path.dirname(os.path.abspath(__file__))
    project_dir = os.path.dirname(current_dir)
    
    # 准备消息
    system_message = prepare_system_message(os.path.join(current_dir, 'table_system_prompt.txt'), key_description_path, table_content_path)
    messages = [
        {"role": "user", "content": system_message}
    ]

    # LLM调用和JSON解析，JSON解析失败时支持重试
    print("调用LLM进行表格匹配...")
    
    max_retries = 3
    analysis_result = None
    
    for attempt in range(max_retries):
        # 首次调用或重试时调用LLM
        if attempt == 0 or analysis_result is None:
            try:
                if attempt == 0:
                    print("首次调用LLM...")
                else:
                    print("重新调用LLM...")
                    
                analysis_result = llm_manager.create_completion(messages, temperature=0)
                
                if not analysis_result:
                    print("LLM返回空结果，终止处理")
                    return []
                
                # 打印模型输出结果
                print(f"LLM调用成功，输出内容：\n{'-'*50}\n{analysis_result}\n{'-'*50}")
                
            except Exception as e:
                print(f"LLM调用失败: {str(e)}")
                return []
        # 尝试解析JSON
        try:
            result_list = parse_llm_response(analysis_result)
            print(f"表格解析成功，返回 {len(result_list)} 条结果")
            return result_list
                
        except Exception as e:
            print(f"尝试 {attempt+1}/{max_retries} 解析失败: {str(e)}")
            
            # 如果还有重试机会，则添加提示准备重新调用LLM
            if attempt < max_retries - 1:
                print("添加提示并准备重新调用LLM...")
                messages.append({"role": "assistant", "content": analysis_result})
                messages.append({
                    "role": "user", 
                    "content": "输出格式错误，请检查"
                })
                # 设置 analysis_result 为 None，让下次循环重新调用LLM
                analysis_result = None
    
    print(f"所有 {max_retries} 次解析尝试都失败了，返回空结果")
    return []
    
    # 如果所有尝试都失败，返回空列表
    return []

def match_tables(document_parts_dir, match_results_dir, key_description_path):
    """
    批量处理多个表格文件，提取字段并与关键词进行匹配
    
    Args:
        document_parts_dir: 包含表格CSV文件的文档元素目录
        match_results_dir: 保存匹配结果的目录
        key_description_path: 关键词描述文件路径
        
    Returns:
        dict: 匹配结果统计信息，包含以下字段:
            - total_tables_processed: 处理的表格总数
            - total_keys_matched: 匹配的键值对总数
            - tables_with_matches: 有匹配结果的表格数
    """
    import os
    import json
    
    # 初始化统计信息
    stats = {
        "total_tables_processed": 0,
        "total_keys_matched": 0,
        "tables_with_matches": 0
    }
    
    # 获取所有表格文件
    table_files = [f for f in os.listdir(document_parts_dir) if f.startswith("table_") and f.endswith(".csv")]
    
    if not table_files:
        print(f"警告：在目录 {document_parts_dir} 中未找到表格文件")
        return stats
    
    # 确保输出目录存在
    os.makedirs(match_results_dir, exist_ok=True)
    
    # 处理每个表格文件
    for table_file in table_files:
        stats["total_tables_processed"] += 1
        table_path = os.path.join(document_parts_dir, table_file)
        
        # 调用表格匹配器
        results = match_table(key_description_path, table_path)
        
        # 保存匹配结果
        if results:
            stats["tables_with_matches"] += 1
            stats["total_keys_matched"] += len(results)
            
            # 构建输出文件名
            output_filename = table_file.replace(".csv", "_matches.json")
            output_path = os.path.join(match_results_dir, output_filename)
            
            # 保存匹配结果
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(results, f, ensure_ascii=False, indent=2)
            print(f"已为 {table_file} 保存 {len(results)} 个匹配结果")
    
    print(f"表格匹配处理完成: 共处理 {stats['total_tables_processed']} 个表格，"
          f"{stats['tables_with_matches']} 个表格有匹配结果，"
          f"总共匹配了 {stats['total_keys_matched']} 个键值对")
    
    return stats

# ================ 主入口 ================

if __name__ == "__main__":
    llm_manager.init_local_model()
    # llm_manager.init_remote_model()
    
    current_dir = os.path.dirname(os.path.abspath(__file__))
    project_dir = os.path.dirname(current_dir)
    key_description_path = os.path.join(project_dir, 'document/key_descriptions/table_key_description.txt')
    table_content_path = os.path.join(project_dir, 'document/document_parts/table_6.csv')
    results = match_table(key_description_path, table_content_path)
    
    # 显示匹配结果
    if results:
        print("\n匹配结果汇总:")
        for i, result in enumerate(results, 1):
            print(f"{i}. key: {result['key']}")
            print(f"   value: {result['value']}")
            print(f"   matched_key: {result['matched_key'] or '无匹配'}")
            print()
    else:
        print("未获取到有效的匹配结果")
