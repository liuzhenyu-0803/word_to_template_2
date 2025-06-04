"""
HTML文件替换模块
专门处理直接修改HTML文件的替换操作
"""

from .table_replacer import replace_tables_in_html, convert_html_to_word
from .paragraph_replacer import replace_paragraphs_in_html  
from .replacer import replace_html_document, create_template_from_original_html

__all__ = [
    'replace_tables_in_html',
    'convert_html_to_word',
    'replace_paragraphs_in_html', 
    'replace_html_document',
    'create_template_from_original_html'
]
