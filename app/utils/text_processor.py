"""
文本处理工具
"""
import re
import markdown
import logging
from typing import List, Dict, Any

logger = logging.getLogger(__name__)

class TextProcessor:
    """文本处理器"""
    
    def __init__(self):
        """初始化文本处理器"""
        # 初始化markdown处理器
        self.md = markdown.Markdown(extensions=['tables', 'toc'])
    
    def clean_markdown_content(self, content: str, title: str, issuing_department: str, 
                             issue_date: str, receiving_department: str = None) -> str:
        """
        清理markdown内容，去除重复的标题、部门、日期信息
        
        Args:
            content: 原始markdown内容
            title: 文档标题
            issuing_department: 发文部门
            issue_date: 发文日期
            receiving_department: 收文部门
            
        Returns:
            str: 清理后的内容
        """
        try:
            cleaned_content = content
            
            # 1. 去除可能重复的标题
            title_patterns = [
                rf'^#\s*{re.escape(title)}\s*$',  # # 标题
                rf'^##\s*{re.escape(title)}\s*$',  # ## 标题
                rf'^###\s*{re.escape(title)}\s*$',  # ### 标题
                rf'^{re.escape(title)}\s*$',  # 纯标题
            ]
            
            for pattern in title_patterns:
                cleaned_content = re.sub(pattern, '', cleaned_content, flags=re.MULTILINE)
            
            # 2. 去除可能重复的部门信息
            dept_patterns = [
                rf'发文部门[:：]\s*{re.escape(issuing_department)}',
                rf'发文单位[:：]\s*{re.escape(issuing_department)}',
                rf'{re.escape(issuing_department)}\s*发文',
            ]
            
            for pattern in dept_patterns:
                cleaned_content = re.sub(pattern, '', cleaned_content, flags=re.IGNORECASE)
            
            # 3. 去除可能重复的日期信息
            date_patterns = [
                rf'发文日期[:：]\s*{re.escape(issue_date)}',
                rf'日期[:：]\s*{re.escape(issue_date)}',
                rf'{re.escape(issue_date)}',
            ]
            
            for pattern in date_patterns:
                cleaned_content = re.sub(pattern, '', cleaned_content, flags=re.IGNORECASE)
            
            # 4. 去除可能重复的收文部门信息
            if receiving_department:
                recv_patterns = [
                    rf'收文部门[:：]\s*{re.escape(receiving_department)}',
                    rf'收文单位[:：]\s*{re.escape(receiving_department)}',
                ]
                
                for pattern in recv_patterns:
                    cleaned_content = re.sub(pattern, '', cleaned_content, flags=re.IGNORECASE)
            
            # 5. 清理多余的空行
            cleaned_content = re.sub(r'\n\s*\n\s*\n+', '\n\n', cleaned_content)
            
            # 6. 去除开头和结尾的空白
            cleaned_content = cleaned_content.strip()
            
            logger.info("文本内容清理完成")
            return cleaned_content
            
        except Exception as e:
            logger.error(f"清理文本内容时发生错误: {str(e)}")
            return content  # 如果处理失败，返回原始内容
    
    def convert_markdown_to_structured_text(self, markdown_content: str) -> Dict[str, Any]:
        """
        将markdown内容转换为结构化文本
        
        Args:
            markdown_content: markdown格式的内容
            
        Returns:
            Dict[str, Any]: 结构化的文本数据
        """
        try:
            # 解析markdown内容
            lines = markdown_content.split('\n')
            structured_data = {
                'paragraphs': [],
                'headers': {
                    'level1': [],  # 一级标题 (一、)
                    'level2': [],  # 二级标题 (（一）)
                    'level3': []   # 三级标题 (1.)
                }
            }
            
            current_paragraph = ""
            
            for line in lines:
                line = line.strip()
                
                if not line:
                    # 空行，结束当前段落
                    if current_paragraph:
                        structured_data['paragraphs'].append(current_paragraph.strip())
                        current_paragraph = ""
                    continue
                
                # 检查是否是标题
                if line.startswith('# '):
                    # 一级标题
                    header_text = line[2:].strip()
                    structured_data['headers']['level1'].append(header_text)
                    if current_paragraph:
                        structured_data['paragraphs'].append(current_paragraph.strip())
                        current_paragraph = ""
                elif line.startswith('## '):
                    # 二级标题
                    header_text = line[3:].strip()
                    structured_data['headers']['level2'].append(header_text)
                    if current_paragraph:
                        structured_data['paragraphs'].append(current_paragraph.strip())
                        current_paragraph = ""
                elif line.startswith('### '):
                    # 三级标题
                    header_text = line[4:].strip()
                    structured_data['headers']['level3'].append(header_text)
                    if current_paragraph:
                        structured_data['paragraphs'].append(current_paragraph.strip())
                        current_paragraph = ""
                else:
                    # 普通文本
                    if current_paragraph:
                        current_paragraph += " " + line
                    else:
                        current_paragraph = line
            
            # 添加最后一个段落
            if current_paragraph:
                structured_data['paragraphs'].append(current_paragraph.strip())
            
            return structured_data
            
        except Exception as e:
            logger.error(f"转换markdown内容时发生错误: {str(e)}")
            return {'paragraphs': [markdown_content], 'headers': {'level1': [], 'level2': [], 'level3': []}}
    
    def format_content_for_document(self, content: str) -> List[Dict[str, Any]]:
        """
        将内容格式化为适合Word文档的结构
        
        Args:
            content: 清理后的内容
            
        Returns:
            List[Dict[str, Any]]: 格式化后的内容结构
        """
        try:
            lines = content.split('\n')
            formatted_content = []
            
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                
                # 识别标题级别并格式化
                # 一级标题：以中文数字加顿号开头
                if re.match(r'^[一二三四五六七八九十]+、', line):
                    formatted_content.append({
                        'type': 'header1',
                        'text': re.sub(r'^[一二三四五六七八九十]+、', '', line).strip(),
                        'style': 'heading1'
                    })

                # 二级标题：以（中文数字）开头
                elif re.match(r'^（[一二三四五六七八九十]+）', line):
                    formatted_content.append({
                        'type': 'header2',
                        'text': re.sub(r'^（[一二三四五六七八九十]+）', '', line).strip(),
                        'style': 'heading2'
                    })

                # 三级标题：数字 + 点 + 标题内容，遇到第一个中文标点符号结束
                elif re.match(r'^(\d+\.)\s*([^，。；：？！,.!?]+)([，。；：？！,.!?])(.*)', line):
                    match = re.match(r'^(\d+\.)\s*([^，。；：？！,.!?]+)([，。；：？！,.!?])(.*)', line)
                    title = match.group(2).strip()
                    rest = match.group(4).strip()
                    formatted_content.append({
                        'type': 'header3',
                        'text': f"{title}{match.group(3)}",  # 保留标点
                        'style': 'heading3'
                    })
                    if rest:
                        formatted_content.append({
                            'type': 'paragraph',
                            'text': rest,
                            'style': 'normal'
                        })
                elif line.startswith('- ') or line.startswith('* '):
                    # 列表项
                    text = line[2:].strip()
                    formatted_content.append({
                        'type': 'list_item',
                        'text': text,
                        'style': 'list'
                    })
                else:
                    # 普通段落
                    formatted_content.append({
                        'type': 'paragraph',
                        'text': line,
                        'style': 'normal'
                    })
            
            return formatted_content
            
        except Exception as e:
            logger.error(f"格式化内容时发生错误: {str(e)}")
            return [{'type': 'paragraph', 'text': content, 'style': 'normal'}]

# 全局文本处理器实例
text_processor = TextProcessor() 