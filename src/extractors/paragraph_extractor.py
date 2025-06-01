"""  
段落提取模块  
负责从Word文档中提取所有段落内容  
"""  

import os  

def extract_paragraphs(doc, output_dir):  
    """提取所有段落并保存"""  
    paragraph_count = 0  
    
    for para in doc.paragraphs:  
        if not para.text.strip():  
            continue  
            
        paragraph_count += 1  
        para_filename = f"paragraph_{paragraph_count}.txt"  
        para_text = para.text.strip()  
        
        # 保存段落内容  
        with open(os.path.join(output_dir, para_filename), 'w', encoding='utf-8') as f:  
            f.write(para_text)  
    
    return paragraph_count  