import win32com.client as win32
import os

def word_to_html(word_file, html_file):
    """使用win32com调用Word应用将Word文件转换为HTML"""
    # 启动Word应用
    word = win32.Dispatch("Word.Application")
    word.Visible = False
    
    try:
        # 打开文档
        doc = word.Documents.Open(word_file)
        
        # 导出为筛选过的HTML
        # wdFormatFilteredHTML = 10
        doc.SaveAs2(html_file, FileFormat=10)
        
        doc.Close()
        print(f"成功导出: {html_file}")
        
    except Exception as e:
        print(f"导出失败: {e}")
    finally:
        word.Quit()

if __name__ == "__main__":
    # 获取项目根目录
    base_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    docx_path = os.path.join(base_dir, "document", "document.docx")
    html_path = os.path.join(base_dir, "document", "document.html")
    print(f"将 {docx_path} 转换为 {html_path}")
    word_to_html(docx_path, html_path)
    print("转换完成！")
