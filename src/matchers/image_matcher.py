"""  
图片匹配模块
负责分析图片的标题和描述，匹配对应的关键字
"""  

import os
import json
import time
import sys

# 添加项目根目录到系统路径
current_dir = os.path.dirname(os.path.abspath(__file__))
project_dir = os.path.dirname(current_dir)
if project_dir not in sys.path:
    sys.path.append(project_dir)

# 使用绝对导入
from models.model_manager import llm_manager

def read_file_content(file_path):
    """读取文件内容"""
    with open(file_path, 'r', encoding='utf-8') as f:
        return f.read()

def prepare_system_message(prompt_path, key_path, content_path):
    """准备user消息
    
    Args:
        prompt_path: 提示词文件路径
        key_path: 关键信息描述文件路径
        content_path: 图片信息文件路径
    """
    # 读取prompt模板、key描述和内容文件
    prompt = read_file_content(prompt_path)
    key_description = read_file_content(key_path)
    content = read_file_content(content_path)
    
    # 替换占位符
    prompt = prompt.replace('placeholder_image_info', content)
    prompt = prompt.replace('placeholder_key_description', key_description)

    return prompt

def parse_llm_response(response_text):
    """从LLM响应中解析出JSON结果"""
    try:
        # 提取<match_result>标签内的内容
        import re
        pattern = r'<match_result>(.*?)</match_result>'
        matches = re.findall(pattern, response_text, re.DOTALL)
        
        if not matches:
            print("警告: 未找到<match_result>标签")
            print(f"模型输出内容: \n{'-'*50}\n{response_text}\n{'-'*50}")
            return None
        
        # 使用最后一个匹配结果（如果有多个）
        json_str = matches[-1].strip()
        # 解析JSON
        result = json.loads(json_str)
        
        # 验证结果是否包含必需的字段
        required_fields = ['name', 'title', 'descr', 'new_key', 'match_reason']
        if all(k in result for k in required_fields):
            print(f"成功解析匹配结果")
            return [result]  # 返回包含单个结果的列表，以保持与其他代码的兼容性
        else:
            print(f"警告: 结果缺少必填字段: {result}")
            return None
        
    except Exception as e:
        print(f"解析LLM响应时出错: {str(e)}")
        print(f"模型原始输出: \n{'-'*50}\n{response_text}\n{'-'*50}")
        return None

def match_image(key_description_path, image_info_path):
    """分析图片信息并匹配关键词"""
    # 构造所需文件路径
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 准备消息
    system_message = prepare_system_message(
        os.path.join(current_dir, 'image_system_prompt.md'),
        key_description_path, 
        image_info_path
    )
    messages = [
        {"role": "user", "content": system_message}
    ]

    # LLM调用和JSON解析分开处理
    print("调用LLM进行图片匹配...")
    
    try:
        analysis_result = llm_manager.create_completion(messages, temperature=0)
        print("LLM调用成功：" + analysis_result)
        if not analysis_result:
            print("LLM返回空结果，终止处理")
            return []
    except Exception as e:
        print(f"LLM调用失败: {str(e)}")
        return []
    
    print("获取到LLM响应，开始解析JSON...")
    
    # 解析JSON可以重试多次
    max_retries = 3
    for attempt in range(max_retries):
        try:
            result_list = parse_llm_response(analysis_result)
            
            if result_list is None:
                print(f"解析失败，模型输出内容: \n{'-'*50}\n{analysis_result}\n{'-'*50}")
                raise ValueError("JSON解析失败")
            
            print(f"图片解析成功，返回 {len(result_list)} 条结果")
            return result_list
                
        except Exception as e:
            print(f"尝试 {attempt+1}/{max_retries} 解析失败: {str(e)}")
            
            if attempt < max_retries - 1:
                print("添加提示并重新调用LLM...")
                messages.append({"role": "assistant", "content": analysis_result})
                messages.append({
                    "role": "user", 
                    "content": "无法解析你的回答。请按prompt中指定的输出格式返回结果"
                })
                
                try:
                    print("重新调用LLM并要求正确格式...")
                    analysis_result = llm_manager.create_completion(messages, temperature=0)
                    if not analysis_result:
                        print("重试时LLM返回空结果")
                        continue
                except Exception as e:
                    print(f"重试时LLM调用失败: {str(e)}")
                    continue
            else:
                print(f"所有 {max_retries} 次解析尝试都失败了，返回空结果")
                return []
    
    return []

def match_images(document_parts_dir, match_results_dir, key_description_path):
    """批量处理图片匹配"""
    # 初始化统计信息
    stats = {
        "total_images_processed": 0,
        "images_with_matches": 0,
        "total_matches": 0
    }
    
    # 获取所有图片信息文件
    image_files = [f for f in os.listdir(document_parts_dir) if f.startswith("image_") and f.endswith(".json")]
    
    if not image_files:
        print(f"警告：在目录 {document_parts_dir} 中未找到图片信息文件")
        return stats
        
    # 确保输出目录存在
    os.makedirs(match_results_dir, exist_ok=True)
    
    # 处理每个图片文件
    for image_file in image_files:
        stats["total_images_processed"] += 1
        image_path = os.path.join(document_parts_dir, image_file)
        
        # 调用图片匹配器
        results = match_image(key_description_path, image_path)
        
        # 保存匹配结果
        if results:
            stats["images_with_matches"] += 1
            stats["total_matches"] += len(results)
            
            # 构建输出文件名
            output_filename = image_file.replace(".json", "_matches.json")
            output_path = os.path.join(match_results_dir, output_filename)
            
            # 保存匹配结果
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(results, f, ensure_ascii=False, indent=2)
            print(f"已为 {image_file} 保存 {len(results)} 个匹配结果")
    
    print(f"图片匹配处理完成: 共处理 {stats['total_images_processed']} 个图片，"
          f"{stats['images_with_matches']} 个图片有匹配结果，"
          f"总共匹配了 {stats['total_matches']} 个键值对")
    
    return stats

# 直接运行时的入口点
if __name__ == "__main__":
    llm_manager.init_local_model()
    # llm_manager.init_remote_model()
    
    current_dir = os.path.dirname(os.path.abspath(__file__))
    project_dir = os.path.dirname(current_dir)
    key_description_path = os.path.join(project_dir, 'document/key_descriptions/image_key_description.txt')
    image_info_path = os.path.join(project_dir, 'document/document_parts/image_4.json')
    
    results = match_image(key_description_path, image_info_path)
    
    # 显示匹配结果
    if results:
        print("\n匹配结果汇总:")
        for i, result in enumerate(results, 1):
            print(f"{i}. 图片名称: {result['name']}")
            print(f"   标题: {result['title']}")
            print(f"   描述: {result['descr']}")
            print(f"   匹配键名: {result['new_key'] or '无匹配'}")
            print(f"   匹配理由: {result['match_reason']}")
            print()
    else:
        print("未获取到有效的匹配结果")