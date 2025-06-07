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

def prepare_system_prompt_1(prompt_path, content_path):
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

def prepare_system_prompt_2(prompt_path, key_path, key_value_array):
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

def parse_response_1(response_text):
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
        
        # 验证格式：[{"key": "...", "value": "...", "valuePos": "..."}]
        if not isinstance(result_list, list):
            raise ValueError("结果不是数组格式")
        
        valid_results = []
        for item in result_list:
            if isinstance(item, dict) and 'key' in item and 'value' in item and 'valuePos' in item:
                valid_results.append(item)
            else:
                print(f"警告: 跳过格式不正确的项: {item}")
        
        print(f"第一阶段提取了 {len(valid_results)} 个key-value对")
        return valid_results
        
    except Exception as e:
        raise ValueError(f"第一阶段解析失败: {str(e)}")

def parse_response_2(response_text):
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

def match_table(table_content_path, key_description_path):
    """
    两阶段表格匹配：
    1. 提取key-value对
    2. 进行key语义匹配
    
    Returns:
        list: [{"old_key": "...", "value": "...", "new_key": "...", "valuePos": "..."}]
    """
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    print("开始两阶段表格匹配...")
      # 第一阶段：提取key-value对
    print("第一阶段：提取key-value对...")
    system_prompt_1 = prepare_system_prompt_1(
        os.path.join(current_dir, 'table_system_prompt_1.md'), 
        table_content_path
    )
    
    for attempt in range(2):
        try:
            print(f"第一阶段第{attempt+1}次调用LLM...")
            response_1 = llm_manager.create_completion([{"role": "user", "content": system_prompt_1}], temperature=0)
            
            if not response_1:
                print("第一阶段LLM返回空结果")
                return []
            
            print(f"第一阶段输出：\n{'-'*30}\n{response_1}\n{'-'*30}")
            key_value_pairs = parse_response_1(response_1)
            
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
    
    # 暂存第一阶段的valuePos信息，创建key到valuePos的映射
    key_to_valuepos = {}
    for item in key_value_pairs:
        key_to_valuepos[item['key']] = item['valuePos']
      # 第二阶段：key匹配
    print("第二阶段：key匹配...")
    # 准备第二阶段输入时，只包含key和value，不包含valuePos
    key_value_for_matching = [{"key": item["key"], "value": item["value"]} for item in key_value_pairs]
    key_value_json = json.dumps(key_value_for_matching, ensure_ascii=False, indent=2)
    system_prompt_2 = prepare_system_prompt_2(
        os.path.join(current_dir, 'table_system_prompt_2.md'),
        key_description_path,
        key_value_json
    )
    
    for attempt in range(2):
        try:
            print(f"第二阶段第{attempt+1}次调用LLM...")
            response_2 = llm_manager.create_completion([{"role": "user", "content": system_prompt_2}], temperature=0)
            
            if not response_2:
                print("第二阶段LLM返回空结果")
                return []
            
            print(f"第二阶段输出：\n{'-'*30}\n{response_2}\n{'-'*30}")
            match_results = parse_response_2(response_2)
            
            # 将第一阶段的valuePos字段合并到第二阶段结果中
            final_results = []
            for item in match_results:
                old_key = item['old_key']
                final_item = {
                    "old_key": old_key,
                    "value": item['value'],
                    "new_key": item['new_key'],
                    "valuePos": key_to_valuepos.get(old_key, "")
                }
                final_results.append(final_item)
            
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

def match_tables(table_files_paths: list[str], key_description_path: str, match_results_dir: str):
    """批量处理表格文件进行两阶段匹配"""
    import os
    import json
    
    stats = {
        "total_tables_processed": 0,
        "total_keys_matched": 0,
        "tables_with_matches": 0
    }
    
    if not table_files_paths:
        print(f"警告：传入的表格文件列表为空")
        return stats
    
    os.makedirs(match_results_dir, exist_ok=True)
    
    # 处理每个表格文件
    for table_path in table_files_paths:
        if not os.path.exists(table_path):
            print(f"警告：文件 {table_path} 不存在，跳过处理。")
            continue
            
        stats["total_tables_processed"] += 1
        table_file_name = os.path.basename(table_path) # 获取文件名用于输出
          # 调用两阶段表格匹配
        results = match_table(table_path, key_description_path)
        
        if results:
            stats["tables_with_matches"] += 1
            stats["total_keys_matched"] += len(results)
            
            # 保存结果
            output_filename = table_file_name.replace(".html", "_matches.json")
            output_path = os.path.join(match_results_dir, output_filename)
            
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(results, f, ensure_ascii=False, indent=2)
            print(f"已为 {table_file_name} 保存 {len(results)} 个匹配结果到 {output_path}")
        else:
            print(f"文件 {table_file_name} 未匹配到结果。")
    
    print(f"表格匹配完成: 处理了 {stats['total_tables_processed']} 个表格，"
          f"{stats['tables_with_matches']} 个有匹配结果，"
          f"总共匹配了 {stats['total_keys_matched']} 个键值对")
    
    return stats

# ================ 主入口 ================

if __name__ == "__main__":
    import time
    start_time = time.time()
    
    # llm_manager.init_local_model()
    llm_manager.init_remote_model()
    
    # 使用项目标准路径结构，但只测试单个表格
    current_dir = os.path.dirname(os.path.abspath(__file__))
    project_dir = os.path.dirname(os.path.dirname(current_dir))
    
    match_results_dir = os.path.join(project_dir, 'document', 'match_results')
    key_description_path = os.path.join(project_dir, 'document', 'key_descriptions', 'table_key_description.txt')
    table_content_path = os.path.join(project_dir, 'document', 'document_extract', 'table_6.html')
    
    print("开始单个表格匹配测试 (table_6)...")
    results = match_table(table_content_path, key_description_path)
    
    # 显示匹配结果
    if results:
        print("\n最终匹配结果:")
        for i, result in enumerate(results, 1):
            print(f"{i}. old_key: {result['old_key']}")
            print(f"   value: {result['value']}")
            print(f"   new_key: {result['new_key'] or '无匹配'}")
            print(f"   valuePos: {result['valuePos']}")
            print()
        
        # 保存匹配结果到文件（与batch处理保持一致的格式）
        os.makedirs(match_results_dir, exist_ok=True)
        output_path = os.path.join(match_results_dir, 'table_6_matches.json')
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(results, f, ensure_ascii=False, indent=2)
        print(f"已保存匹配结果到: {output_path}")
    else:
        print("未获取到有效的匹配结果")
    
    # 总耗时
    end_time = time.time()
    print(f"\n总耗时: {end_time - start_time:.2f} 秒")
