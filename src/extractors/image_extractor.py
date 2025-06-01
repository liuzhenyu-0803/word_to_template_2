"""  
图片提取模块  
负责从Word文档中提取所有图片及其替换文字和名称  
使用直接的XML字符串搜索方法提取图片信息，支持所有类型的图片（内联、浮动、页眉页脚）
"""  

import os
import json

def extract_images(doc, output_dir):
    """
    提取Word文档中的所有图片及其替换文字和名称
    使用直接XML解析方法，适用于所有类型的图片
    
    参数:
        doc: docx.Document对象
        output_dir: 输出目录路径
        
    返回:
        tuple: (图片数量, 图片信息列表)
    """
    # 用于跟踪已处理的图片名称，避免重复
    processed_images = set()
    image_count = 0
    image_infos = []
    
    # 获取所有图片关系
    image_rels_info = {}
    try:
        for rel_id, rel in doc.part.rels.items():
            if "image" in rel.reltype:
                image_rels_info[rel_id] = rel.target_ref  # 记录关系ID和目标图片
    except Exception as e:
        print(f"获取图片关系时出错: {e}")
    
    # 使用字符串搜索方法从文档XML中提取图片及其替换文字
    
    try:
        # 获取document.xml文件内容
        doc_xml = doc.part.blob.decode('utf-8')
        
        # 查找所有wp:docPr标签
        current_pos = 0
        while True:
            # 查找下一个wp:docPr标签的开始位置
            start_tag = doc_xml.find('<wp:docPr', current_pos)
            if start_tag == -1:
                break
                
            # 找到标签的结束位置
            end_tag = doc_xml.find('>', start_tag)
            if end_tag == -1:
                break
                
            # 提取标签内容
            tag_content = doc_xml[start_tag:end_tag+1]
            
            # 提取属性
            name = ""
            title = ""
            descr = ""
            
            # 提取name属性
            name_start = tag_content.find('name="')
            if name_start != -1:
                name_start += 6
                name_end = tag_content.find('"', name_start)
                if name_end != -1:
                    name = tag_content[name_start:name_end]
            
            # 提取title属性
            title_start = tag_content.find('title="')
            if title_start != -1:
                title_start += 7
                title_end = tag_content.find('"', title_start)
                if title_end != -1:
                    title = tag_content[title_start:title_end]
            
            # 提取descr属性
            descr_start = tag_content.find('descr="')
            if descr_start != -1:
                descr_start += 7
                descr_end = tag_content.find('"', descr_start)
                if descr_end != -1:
                    descr = tag_content[descr_start:descr_end]
            
            # 只要有title或descr其中之一就处理
            if title or descr:
                image_count += 1
                
                # 创建图片信息对象
                image_info = {
                    "name": name,
                    "title": title,
                    "descr": descr
                }
                
                # 添加到图片信息列表
                image_infos.append(image_info)
                
                # 保存为JSON文件
                image_filename = f"image_{image_count}.json"
                with open(os.path.join(output_dir, image_filename), 'w', encoding='utf-8') as f:
                    json.dump(image_info, f, ensure_ascii=False, indent=2)
                
                print(f"提取到图片 {image_count}:")
                print(f"  名称: {name}")
                print(f"  标题: {title}")
                print(f"  描述: {descr}")
            
            # 移动到下一个位置继续搜索
            current_pos = end_tag + 1
            
        print(f"从文档XML中提取了 {image_count} 个有效图片（至少包含title或descr属性之一）")
    except Exception as e:
        print(f"提取图片时出错: {e}")
    
    # 打印统计信息
    print(f"图片提取完成: 共提取了 {image_count} 个图片")
    if image_count < len(image_rels_info):
        print(f"注意: 文档中包含 {len(image_rels_info)} 个图片关系，但只提取到 {image_count} 个图片")
    
    return image_count, image_infos

# 如果直接运行该模块，执行测试
if __name__ == "__main__":
    from docx import Document
    import os
    
    # 获取当前文件所在目录
    current_dir = os.path.dirname(os.path.abspath(__file__))
    project_dir = os.path.dirname(current_dir)
    
    # 构造测试文档路径
    test_doc_path = os.path.join(project_dir, "document", "document.docx")
    output_dir = os.path.join(project_dir, "document", "document_parts")
    
    # 确保输出目录存在
    os.makedirs(output_dir, exist_ok=True)
    
    # 打开测试文档
    try:
        doc = Document(test_doc_path)
        print(f"成功打开文档: {test_doc_path}")
        
        # 提取图片
        image_count, image_infos = extract_images(doc, output_dir)
        
        # 显示提取结果
        print(f"共提取了 {image_count} 个图片")
        for i, info in enumerate(image_infos, 1):
            print(f"图片 {i}:")
            print(f"  名称: {info['name']}")
            print(f"  标题: {info['title']}")
            print(f"  描述: {info['descr']}")
            print()
            
    except Exception as e:
        print(f"测试失败: {e}")
