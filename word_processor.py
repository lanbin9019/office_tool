import os
import re
from typing import List, Dict, Any
from collections import Counter
from datetime import datetime

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

import docx2txt

from office_tool.utils import generate_output_filename


class WordProcessor:
    """Word处理器类"""
    
    def __init__(self):
        self._stop_words = {
            '的', '了', '是', '在', '我', '有', '和', '就', '不', '人', '都', '一', '一个',
            '上', '也', '很', '到', '说', '要', '去', '你', '会', '着', '没有', '看', '好',
            '自己', '这', '那', '他', '她', '它', '们', '这个', '那个', '什么', '怎么',
            '为什么', '哪', '哪里', '谁', '多少', '几', '啊', '吧', '呢', '吗', '呀', '哦',
            '嗯', '哈', 'the', 'a', 'an', 'is', 'are', 'was', 'were', 'be', 'been', 'being',
            'in', 'on', 'at', 'to', 'for', 'of', 'and', 'or', 'but', 'so', 'if', 'then',
            'this', 'that', 'these', 'those', 'it', 'he', 'she', 'they', 'we', 'you', 'i',
            'my', 'your', 'his', 'her', 'its', 'our', 'their', 'me', 'him', 'them', 'us',
            'have', 'has', 'had', 'do', 'does', 'did', 'will', 'would', 'could', 'should',
            'may', 'might', 'must', 'can', 'need', 'dare', 'ought', 'used', 'better',
            'with', 'by', 'from', 'as', 'into', 'through', 'during', 'before', 'after',
            'above', 'below', 'between', 'under', 'again', 'further', 'then', 'once',
            'here', 'there', 'when', 'where', 'why', 'how', 'all', 'each', 'few', 'more',
            'most', 'other', 'some', 'such', 'no', 'nor', 'not', 'only', 'own', 'same',
            'just', 'than', 'too', 'very', 'also', 'well', 'still', 'even', 'up', 'down'
        }
    
    def merge_files(self, file_paths: List[str], log_callback=None) -> str:
        """
        合并多个Word文件为一个文件
        
        Args:
            file_paths: Word文件路径列表
            log_callback: 日志回调函数
            
        Returns:
            输出文件路径
        """
        if len(file_paths) < 1:
            raise ValueError("至少需要一个Word文件")
        
        output_path = generate_output_filename(file_paths[0], "merged", ".docx")
        
        merged_doc = Document()
        
        for i, file_path in enumerate(file_paths):
            file_name = os.path.basename(file_path)
            
            if log_callback:
                log_callback(f"正在处理: {file_name}")
            
            if i > 0:
                merged_doc.add_page_break()
            
            title_para = merged_doc.add_paragraph()
            title_run = title_para.add_run(f"=== {file_name} ===")
            title_run.bold = True
            title_run.font.size = Pt(14)
            title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            try:
                doc = Document(file_path)
                
                for element in doc.element.body:
                    merged_doc.element.body.append(element)
                    
            except Exception as e:
                if log_callback:
                    log_callback(f"警告: 处理 {file_name} 时出错: {str(e)}")
        
        merged_doc.save(output_path)
        
        if log_callback:
            log_callback(f"合并完成，共 {len(file_paths)} 个文件")
        
        return output_path
    
    def split_file(self, file_path: str, pages_per_file: int, log_callback=None) -> List[str]:
        """
        按段落数拆分Word文件（注：python-docx无法直接获取页码，按段落数估算）
        
        Args:
            file_path: Word文件路径
            pages_per_file: 每N个段落为一个文件（近似页数）
            log_callback: 日志回调函数
            
        Returns:
            输出文件路径列表
        """
        output_paths = []
        base_dir = os.path.dirname(file_path)
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        doc = Document(file_path)
        
        paragraphs_per_page = max(1, pages_per_file * 10)
        all_paragraphs = doc.paragraphs
        total_paragraphs = len(all_paragraphs)
        
        if log_callback:
            log_callback(f"文档共 {total_paragraphs} 个段落，每 {paragraphs_per_page} 个段落拆分为一个文件")
        
        for i in range(0, total_paragraphs, paragraphs_per_page):
            start_idx = i
            end_idx = min(i + paragraphs_per_page, total_paragraphs)
            
            if log_callback:
                log_callback(f"正在拆分段落 {start_idx+1} - {end_idx}")
            
            new_doc = Document()
            
            for para_idx in range(start_idx, end_idx):
                para = all_paragraphs[para_idx]
                new_para = new_doc.add_paragraph()
                
                for run in para.runs:
                    new_run = new_para.add_run(run.text)
                    new_run.bold = run.bold
                    new_run.italic = run.italic
                    new_run.underline = run.underline
                    if run.font.size:
                        new_run.font.size = run.font.size
                    if run.font.name:
                        new_run.font.name = run.font.name
            
            file_num = (i // paragraphs_per_page) + 1
            output_name = f"{base_name}_part{file_num}_{timestamp}.docx"
            output_path = os.path.join(base_dir, output_name)
            new_doc.save(output_path)
            output_paths.append(output_path)
        
        if log_callback:
            log_callback(f"拆分完成，共生成 {len(output_paths)} 个文件")
        
        return output_paths
    
    def generate_summary(self, file_path: str, summary_length: str, log_callback=None) -> str:
        """
        生成文档内容总结
        
        Args:
            file_path: Word文件路径
            summary_length: 总结长度（简短/中等/详细）
            log_callback: 日志回调函数
            
        Returns:
            输出文件路径
        """
        if log_callback:
            log_callback("正在分析文档内容...")
        
        text = self._extract_text(file_path)
        
        if not text.strip():
            raise ValueError("文档内容为空")
        
        if log_callback:
            log_callback("正在生成总结...")
        
        length_config = {
            '简短': {'sentences': 3, 'keywords': 5},
            '中等': {'sentences': 5, 'keywords': 10},
            '详细': {'sentences': 10, 'keywords': 15}
        }
        
        config = length_config.get(summary_length, length_config['中等'])
        
        sentences = self._split_sentences(text)
        keywords = self._extract_keywords(text, config['keywords'])
        
        important_sentences = self._select_important_sentences(
            sentences, keywords, config['sentences']
        )
        
        summary = self._build_summary(important_sentences, keywords, summary_length)
        
        output_path = generate_output_filename(file_path, "summary", ".docx")
        self._save_summary_to_docx(summary, keywords, output_path, file_path)
        
        if log_callback:
            log_callback(f"总结生成完成，共 {len(important_sentences)} 个关键句子")
        
        return output_path
    
    def analyze_document(self, file_path: str, log_callback=None) -> Dict[str, Any]:
        """
        分析文档统计信息
        
        Args:
            file_path: Word文件路径
            log_callback: 日志回调函数
            
        Returns:
            分析结果字典
        """
        if log_callback:
            log_callback("正在分析文档...")
        
        doc = Document(file_path)
        text = self._extract_text(file_path)
        
        paragraph_count = len(doc.paragraphs)
        table_count = len(doc.tables)
        
        word_count = len(re.findall(r'\b\w+\b', text))
        char_count = len(text)
        char_count_no_space = len(text.replace(' ', '').replace('\n', ''))
        
        sentences = self._split_sentences(text)
        sentence_count = len(sentences)
        
        sections = self._extract_sections(doc)
        
        keywords = self._extract_keywords(text, 10)
        
        result = {
            '段落数': paragraph_count,
            '表格数': table_count,
            '总字数（含空格）': char_count,
            '总字数（不含空格）': char_count_no_space,
            '词数': word_count,
            '句子数': sentence_count,
            '主要章节': sections if sections else ['无明显章节'],
            '关键词': keywords
        }
        
        if log_callback:
            log_callback(f"分析完成: 段落={paragraph_count}, 表格={table_count}, 词数={word_count}")
        
        return result
    
    def _extract_text(self, file_path: str) -> str:
        """提取文档文本"""
        try:
            text = docx2txt.process(file_path)
            return text
        except:
            try:
                doc = Document(file_path)
                text_parts = []
                for para in doc.paragraphs:
                    text_parts.append(para.text)
                return '\n'.join(text_parts)
            except:
                return ''
    
    def _split_sentences(self, text: str) -> List[str]:
        """分割句子"""
        text = text.replace('\n', ' ')
        text = re.sub(r'\s+', ' ', text).strip()
        
        sentences = re.split(r'(?<=[。！？.!?])\s*', text)
        sentences = [s.strip() for s in sentences if s.strip() and len(s.strip()) > 5]
        
        return sentences
    
    def _extract_keywords(self, text: str, top_n: int = 10) -> List[str]:
        """提取关键词"""
        text = text.lower()
        
        words = re.findall(r'\b[\u4e00-\u9fa5]{2,}\b|\b[a-zA-Z]{3,}\b', text)
        
        words = [w for w in words if w not in self._stop_words]
        
        word_counts = Counter(words)
        
        keywords = [word for word, count in word_counts.most_common(top_n)]
        
        return keywords
    
    def _select_important_sentences(self, sentences: List[str], keywords: List[str], 
                                     num_sentences: int) -> List[str]:
        """选择重要句子"""
        if not sentences:
            return []
        
        scored_sentences = []
        
        for sentence in sentences:
            score = 0
            
            for keyword in keywords:
                if keyword.lower() in sentence.lower():
                    score += 1
            
            if len(sentence) > 50:
                score += 1
            
            if score > 0:
                scored_sentences.append((sentence, score))
        
        scored_sentences.sort(key=lambda x: x[1], reverse=True)
        
        selected = [s[0] for s in scored_sentences[:num_sentences]]
        
        original_order = []
        for sentence in sentences:
            if sentence in selected and sentence not in original_order:
                original_order.append(sentence)
        
        return original_order if original_order else sentences[:num_sentences]
    
    def _build_summary(self, sentences: List[str], keywords: List[str], 
                       summary_length: str) -> str:
        """构建总结文本"""
        if not sentences:
            return "文档内容不足，无法生成有意义的总结。"
        
        summary_parts = []
        
        summary_parts.append(f"【文档总结 - {summary_length}版】")
        summary_parts.append("=" * 50)
        summary_parts.append("")
        
        summary_parts.append("### 核心内容：")
        for i, sentence in enumerate(sentences, 1):
            summary_parts.append(f"{i}. {sentence}")
        
        summary_parts.append("")
        summary_parts.append("### 关键词：")
        summary_parts.append("、".join(keywords) if keywords else "无")
        
        return "\n".join(summary_parts)
    
    def _save_summary_to_docx(self, summary: str, keywords: List[str], 
                               output_path: str, original_path: str):
        """保存总结到Word文档"""
        doc = Document()
        
        title = doc.add_paragraph()
        title_run = title.add_run("文档内容总结")
        title_run.bold = True
        title_run.font.size = Pt(18)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        doc.add_paragraph(f"原文档：{os.path.basename(original_path)}")
        doc.add_paragraph(f"生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        doc.add_paragraph("-" * 50)
        
        lines = summary.split('\n')
        for line in lines:
            if line.startswith('###'):
                para = doc.add_paragraph()
                run = para.add_run(line.replace('### ', ''))
                run.bold = True
                run.font.size = Pt(12)
            elif line.strip() and not line.startswith('='):
                doc.add_paragraph(line)
        
        doc.save(output_path)
    
    def _extract_sections(self, doc) -> List[str]:
        """提取文档章节"""
        sections = []
        
        for para in doc.paragraphs:
            if para.style and 'Heading' in para.style.name:
                if para.text.strip():
                    sections.append(para.text.strip())
            
            if para.runs:
                for run in para.runs:
                    if run.bold and len(para.text.strip()) > 0 and len(para.text.strip()) < 50:
                        if para.text.strip() not in sections:
                            sections.append(para.text.strip())
                        break
        
        return sections[:10]
