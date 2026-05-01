import os
import re
from typing import List, Dict, Any
from collections import Counter
from datetime import datetime

import PyPDF2
import pdfplumber

from office_tool.utils import generate_output_filename


class PDFProcessor:
    """PDF处理器类"""
    
    def __init__(self):
        self._stop_words = {
            '的', '了', '是', '在', '我', '有', '和', '就', '不', '人', '都', '一', '一个',
            '上', '也', '很', '到', '说', '要', '去', '你', '会', '着', '没有', '看', '好',
            'the', 'a', 'an', 'is', 'are', 'was', 'were', 'be', 'been', 'being',
            'in', 'on', 'at', 'to', 'for', 'of', 'and', 'or', 'but', 'so', 'if', 'then',
            'this', 'that', 'these', 'those', 'it', 'he', 'she', 'they', 'we', 'you', 'i'
        }
    
    def merge_files(self, file_paths: List[str], log_callback=None) -> str:
        """
        合并多个PDF文件为一个文件
        
        Args:
            file_paths: PDF文件路径列表
            log_callback: 日志回调函数
            
        Returns:
            输出文件路径
        """
        if len(file_paths) < 1:
            raise ValueError("至少需要一个PDF文件")
        
        output_path = generate_output_filename(file_paths[0], "merged", ".pdf")
        
        merger = PyPDF2.PdfMerger()
        
        try:
            for i, file_path in enumerate(file_paths):
                file_name = os.path.basename(file_path)
                
                if log_callback:
                    log_callback(f"正在合并: {file_name}")
                
                with open(file_path, 'rb') as f:
                    merger.append(f)
            
            with open(output_path, 'wb') as f:
                merger.write(f)
            
            if log_callback:
                log_callback(f"合并完成，共 {len(file_paths)} 个文件")
            
            return output_path
            
        finally:
            merger.close()
    
    def split_file(self, file_path: str, start_page: int, end_page: int, log_callback=None) -> str:
        """
        拆分PDF文件（提取指定页码范围）
        
        Args:
            file_path: PDF文件路径
            start_page: 起始页码（从1开始）
            end_page: 结束页码
            log_callback: 日志回调函数
            
        Returns:
            输出文件路径
        """
        output_path = generate_output_filename(file_path, f"split_p{start_page}-{end_page}", ".pdf")
        
        with open(file_path, 'rb') as f:
            reader = PyPDF2.PdfReader(f)
            total_pages = len(reader.pages)
            
            if start_page < 1 or end_page > total_pages:
                raise ValueError(f"页码范围无效。文档共 {total_pages} 页。")
            
            if log_callback:
                log_callback(f"正在提取页码 {start_page}-{end_page}（共 {total_pages} 页）")
            
            writer = PyPDF2.PdfWriter()
            
            for page_num in range(start_page - 1, end_page):
                page = reader.pages[page_num]
                writer.add_page(page)
            
            with open(output_path, 'wb') as out_f:
                writer.write(out_f)
        
        if log_callback:
            log_callback(f"拆分完成，共 {end_page - start_page + 1} 页")
        
        return output_path
    
    def generate_summary(self, file_path: str, log_callback=None) -> str:
        """
        生成PDF内容总结
        
        Args:
            file_path: PDF文件路径
            log_callback: 日志回调函数
            
        Returns:
            输出文件路径（.txt格式）
        """
        if log_callback:
            log_callback("正在提取PDF文本内容...")
        
        text = self._extract_text(file_path)
        
        if not text.strip():
            raise ValueError("PDF内容为空或无法提取文本")
        
        if log_callback:
            log_callback("正在分析内容并生成总结...")
        
        sentences = self._split_sentences(text)
        keywords = self._extract_keywords(text, 15)
        
        important_sentences = self._select_important_sentences(
            sentences, keywords, min(10, len(sentences))
        )
        
        summary = self._build_summary(important_sentences, keywords, file_path)
        
        output_path = generate_output_filename(file_path, "summary", ".txt")
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(summary)
        
        if log_callback:
            log_callback(f"总结生成完成，共 {len(important_sentences)} 个关键句子")
        
        return output_path
    
    def extract_text(self, file_path: str, log_callback=None) -> str:
        """
        提取PDF文本内容
        
        Args:
            file_path: PDF文件路径
            log_callback: 日志回调函数
            
        Returns:
            输出文件路径（.txt格式）
        """
        if log_callback:
            log_callback("正在提取PDF文本内容...")
        
        text = self._extract_text(file_path)
        
        if not text.strip():
            raise ValueError("PDF内容为空或无法提取文本")
        
        output_path = generate_output_filename(file_path, "extracted", ".txt")
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(text)
        
        if log_callback:
            stats = self._get_text_stats(text)
            log_callback(f"文本提取完成: {stats}")
        
        return output_path
    
    def _extract_text(self, file_path: str) -> str:
        """提取PDF文本"""
        try:
            all_text = []
            with pdfplumber.open(file_path) as pdf:
                for i, page in enumerate(pdf.pages):
                    text = page.extract_text()
                    if text:
                        all_text.append(f"\n--- 第 {i+1} 页 ---\n{text}")
            
            if all_text:
                return '\n'.join(all_text)
        except:
            pass
        
        try:
            all_text = []
            with open(file_path, 'rb') as f:
                reader = PyPDF2.PdfReader(f)
                for i, page in enumerate(reader.pages):
                    text = page.extract_text()
                    if text:
                        all_text.append(f"\n--- 第 {i+1} 页 ---\n{text}")
            
            if all_text:
                return '\n'.join(all_text)
        except:
            pass
        
        return ''
    
    def _get_text_stats(self, text: str) -> str:
        """获取文本统计信息"""
        char_count = len(text)
        word_count = len(re.findall(r'\b\w+\b', text))
        line_count = text.count('\n')
        
        return f"字符数: {char_count}, 词数: {word_count}, 行数: {line_count}"
    
    def _split_sentences(self, text: str) -> List[str]:
        """分割句子"""
        text = text.replace('\n', ' ')
        text = re.sub(r'\s+', ' ', text).strip()
        
        sentences = re.split(r'(?<=[。！？.!?])\s*', text)
        sentences = [s.strip() for s in sentences if s.strip() and len(s.strip()) > 10]
        
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
            
            if len(sentence) > 30:
                score += 1
            
            if score > 0:
                scored_sentences.append((sentence, score))
        
        scored_sentences.sort(key=lambda x: x[1], reverse=True)
        
        selected = [s[0] for s in scored_sentences[:num_sentences]]
        
        original_order = []
        for sentence in sentences:
            if sentence in selected and sentence not in original_order:
                original_order.append(sentence)
        
        return original_order if original_order else sentences[:min(num_sentences, len(sentences))]
    
    def _build_summary(self, sentences: List[str], keywords: List[str], 
                       file_path: str) -> str:
        """构建总结文本"""
        summary_parts = []
        
        summary_parts.append("=" * 60)
        summary_parts.append("PDF文档内容总结")
        summary_parts.append("=" * 60)
        summary_parts.append(f"文档: {os.path.basename(file_path)}")
        summary_parts.append(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        summary_parts.append("")
        
        if keywords:
            summary_parts.append("【关键词】")
            summary_parts.append("、".join(keywords))
            summary_parts.append("")
        
        if sentences:
            summary_parts.append("【核心内容摘要】")
            summary_parts.append("-" * 40)
            for i, sentence in enumerate(sentences, 1):
                summary_parts.append(f"{i}. {sentence}")
        else:
            summary_parts.append("【核心内容摘要】")
            summary_parts.append("文档内容较短，无法提取有效摘要。")
        
        summary_parts.append("")
        summary_parts.append("=" * 60)
        summary_parts.append("总结结束")
        summary_parts.append("=" * 60)
        
        return "\n".join(summary_parts)
