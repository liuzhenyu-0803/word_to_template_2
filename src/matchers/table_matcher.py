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

def prepare_extract_message(prompt_path, content_path):
    """准备第一阶段提取key-value的消息
    
    Args:
        prompt_path: 提示词文件路径
        content_path: 内容文件路径
    """
    # 读取prompt模板和内容文件
    prompt = read_file_content(prompt_path)
    content = read_file_content(content_path)
    
    # 替换占位符
    prompt = prompt.replace('placeholder_table_content', content)

    return prompt

def prepare_match_message(prompt_path, key_path, key_value_array):
    """准备第二阶段匹配key的消息
    
    Args:
        prompt_path: 提示词文件路径
        key_path: 关键信息描述文件路径
        key_value_array: 第一阶段提取的key-value数组的JSON字符串
    """
    # 读取prompt模板和key描述文件
    prompt = read_file_content(prompt_path)
    key_description = read_file_content(key_path)
    
    # 替换占位符
    prompt = prompt.replace('placeholder_key_value_array', key_value_array)
    prompt = prompt.replace('placeholder_key_description', key_description)

    return prompt

# 注意：原call_llm函数已被移除，现在直接使用llm_manager.create_completion

def parse_extract_response(response_text):
    """解析第一阶段的key-value提取结果"""
    try:
        # 直接解析JSON或提取JSON部分
        clean_text = response_text.strip()
        try:
            result_list = json.loads(clean_text)
        except json.JSONDecodeError:
            # 提取JSON数组部分
            start_pos = response_text.find('[')
            end_pos = response_text.rfind(']')
            if start_pos != -1 and end_pos != -1 and end_pos > start_pos:
                result_list = json.loads(response_text[start_pos:end_pos+1])
            else:
                raise ValueError("无法找到有效的JSON数组")
        
        # 验证格式：[{"key": "...", "value": "..."}]
        if not isinstance(result_list, list):
            raise ValueError("结果不是数组格式")
        
        valid_results = []
        for item in result_list:
            if isinstance(item, dict) and 'key' in item and 'value' in item:
                valid_results.append(item)
            else:
                print(f"警告: 跳过格式不正确的项: {item}")
        
        print(f"第一阶段提取了 {len(valid_results)} 个key-value对")
        return valid_results
        
    except Exception as e:
        raise ValueError(f"第一阶段解析失败: {str(e)}")

def parse_match_response(response_text):
    """解析第二阶段的匹配结果"""
    try:
        # 直接解析JSON或提取JSON部分
        clean_text = response_text.strip()
        try:
            result_list = json.loads(clean_text)
        except json.JSONDecodeError:
            # 提取JSON数组部分
            start_pos = response_text.find('[')
            end_pos = response_text.rfind(']')
            if start_pos != -1 and end_pos != -1 and end_pos > start_pos:
                result_list = json.loads(response_text[start_pos:end_pos+1])
            else:
                raise ValueError("无法找到有效的JSON数组")
        
        # 验证格式：[{"old_key": "...", "value": "...", "new_key": "..."}]
        if not isinstance(result_list, list):
            raise ValueError("结果不是数组格式")
        
        valid_results = []
        for item in result_list:
            if isinstance(item, dict) and all(k in item for k in ['old_key', 'value', 'new_key']):
                valid_results.append(item)
            else:
                print(f"警告: 跳过格式不正确的项: {item}")
        
        print(f"第二阶段匹配了 {len(valid_results)} 个结果")
        return valid_results
        
    except Exception as e:
        raise ValueError(f"第二阶段解析失败: {str(e)}")

# ================ 主要功能函数 ================

def match_table(key_description_path, table_content_path):
    """
    两阶段表格匹配：
    1. 提取key-value对
    2. 进行key语义匹配
    
    Returns:
        list: [{"old_key": "...", "value": "...", "new_key": "..."}]
    """
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    print("开始两阶段表格匹配...")
    
    # 第一阶段：提取key-value对
    print("第一阶段：提取key-value对...")
    extract_message = prepare_extract_message(
        os.path.join(current_dir, 'table_system_prompt_1.md'), 
        table_content_path
    )
    
    for attempt in range(2):
        try:
            print(f"第一阶段第{attempt+1}次调用LLM...")
            extract_result = llm_manager.create_completion([{"role": "user", "content": extract_message}], temperature=0)
            
            if not extract_result:
                print("第一阶段LLM返回空结果")
                return []
            
            print(f"第一阶段输出：\n{'-'*30}\n{extract_result}\n{'-'*30}")
            key_value_pairs = parse_extract_response(extract_result)
            break
                
        except Exception as e:
            if attempt == 0:
                print(f"第一阶段失败: {str(e)}，重试中...")
                continue
            else:
                print(f"第一阶段最终失败: {str(e)}")
                return []
    else:
        return []
    
    if not key_value_pairs:
        print("第一阶段未提取到key-value对")
        return []
    
    # 第二阶段：key匹配
    print("第二阶段：key匹配...")
    key_value_json = json.dumps(key_value_pairs, ensure_ascii=False, indent=2)
    match_message = prepare_match_message(
        os.path.join(current_dir, 'table_system_prompt_2.md'),
        key_description_path,
        key_value_json
    )
    
    for attempt in range(2):
        try:
            print(f"第二阶段第{attempt+1}次调用LLM...")
            match_result = llm_manager.create_completion([{"role": "user", "content": match_message}], temperature=0)
            
            if not match_result:
                print("第二阶段LLM返回空结果")
                return []
            
            print(f"第二阶段输出：\n{'-'*30}\n{match_result}\n{'-'*30}")
            final_results = parse_match_response(match_result)
            print(f"匹配完成，返回 {len(final_results)} 个结果")
            return final_results
                
        except Exception as e:
            if attempt == 0:
                print(f"第二阶段失败: {str(e)}，重试中...")
                continue
            else:
                print(f"第二阶段最终失败: {str(e)}")
                return []
    
    return []

def match_tables(document_parts_dir, match_results_dir, key_description_path):
    """批量处理表格文件进行两阶段匹配"""
    import os
    import json
    
    stats = {
        "total_tables_processed": 0,
        "total_keys_matched": 0,
        "tables_with_matches": 0
    }
    
    # 获取所有表格文件
    table_files = [f for f in os.listdir(document_parts_dir) if f.startswith("table_") and f.endswith(".html")]
    
    if not table_files:
        print(f"警告：在目录 {document_parts_dir} 中未找到表格文件")
        return stats
    
    os.makedirs(match_results_dir, exist_ok=True)
    
    # 处理每个表格文件
    for table_file in table_files:
        stats["total_tables_processed"] += 1
        table_path = os.path.join(document_parts_dir, table_file)
        
        # 调用两阶段表格匹配
        results = match_table(key_description_path, table_path)
        
        if results:
            stats["tables_with_matches"] += 1
            stats["total_keys_matched"] += len(results)
            
            # 保存结果
            output_filename = table_file.replace(".html", "_matches.json")
            output_path = os.path.join(match_results_dir, output_filename)
            
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(results, f, ensure_ascii=False, indent=2)
            print(f"已为 {table_file} 保存 {len(results)} 个匹配结果")
    
    print(f"表格匹配完成: 处理了 {stats['total_tables_processed']} 个表格，"
          f"{stats['tables_with_matches']} 个有匹配结果，"
          f"总共匹配了 {stats['total_keys_matched']} 个键值对")
    
    return stats

# ================ 主入口 ================

if __name__ == "__main__":
    import time
    start_time = time.time()
    
    llm_manager.init_local_model()
    # llm_manager.init_remote_model()
    
    current_dir = os.path.dirname(os.path.abspath(__file__))
    project_dir = os.path.dirname(os.path.dirname(current_dir))
    key_description_path = os.path.join(project_dir, 'document', 'key_descriptions', 'table_key_description.txt')
    table_content_path = os.path.join(project_dir, 'document', 'document_parts', 'table_6.html')
    results = match_table(key_description_path, table_content_path)
    
    # 显示匹配结果
    if results:
        print("\n最终匹配结果:")
        for i, result in enumerate(results, 1):
            print(f"{i}. old_key: {result['old_key']}")
            print(f"   value: {result['value']}")
            print(f"   new_key: {result['new_key'] or '无匹配'}")
            print()
    else:
        print("未获取到有效的匹配结果")
    
    # 总耗时
    end_time = time.time()
    print(f"\n总耗时: {end_time - start_time:.2f} 秒")
