import sys
import os
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QListWidget, QListWidgetItem, QPushButton, QLabel, QSplitter,
    QTextEdit, QScrollArea, QFrame, QSizePolicy, QFileDialog,
    QMessageBox, QTabWidget, QGroupBox, QSpinBox, QComboBox,
    QCheckBox, QLineEdit, QProgressBar
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QIcon

from office_tool.config import Config
from office_tool.utils import get_file_extension, show_message, format_file_size
from office_tool.file_manager import FileManager
from office_tool.excel_processor import ExcelProcessor
from office_tool.word_processor import WordProcessor
from office_tool.pdf_processor import PDFProcessor
from office_tool.finance_analyzer import FinanceAnalyzer
from office_tool.resume_filter import ResumeFilter


class WorkerThread(QThread):
    """后台工作线程"""
    log_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(bool, str)
    progress_signal = pyqtSignal(int)
    
    def __init__(self, func, *args, **kwargs):
        super().__init__()
        self.func = func
        self.args = args
        self.kwargs = kwargs
        self.kwargs['log_callback'] = self.log_signal.emit
        self.kwargs['progress_callback'] = self.progress_signal.emit
    
    def run(self):
        try:
            result = self.func(*self.args, **self.kwargs)
            self.finished_signal.emit(True, result if isinstance(result, str) else "操作完成")
        except Exception as e:
            self.finished_signal.emit(False, str(e))


class MainWindow(QMainWindow):
    """主窗口类"""
    
    def __init__(self):
        super().__init__()
        self.config = Config()
        self.file_manager = FileManager()
        self.excel_processor = ExcelProcessor()
        self.word_processor = WordProcessor()
        self.pdf_processor = PDFProcessor()
        self.finance_analyzer = FinanceAnalyzer()
        self.resume_filter = ResumeFilter()
        
        self.worker_thread = None
        self.init_ui()
        self.connect_signals()
    
    def init_ui(self):
        """初始化UI"""
        self.setWindowTitle(f"{self.config.APP_NAME} v{self.config.APP_VERSION}")
        self.setMinimumSize(self.config.WINDOW_WIDTH, self.config.WINDOW_HEIGHT)
        self.setStyleSheet(self.get_stylesheet())
        
        # 中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 主布局
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(10)
        
        # 标题栏
        title_label = QLabel(self.config.APP_NAME)
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setFont(QFont(self.config.FONT_FAMILY, 18, QFont.Bold))
        title_label.setStyleSheet(f"color: {self.config.PRIMARY_COLOR}; padding: 10px;")
        main_layout.addWidget(title_label)
        
        # 主分割器
        main_splitter = QSplitter(Qt.Horizontal)
        
        # 左侧：文件管理区
        left_widget = self.create_file_manager_widget()
        main_splitter.addWidget(left_widget)
        
        # 右侧：功能操作区
        right_widget = self.create_function_widget()
        main_splitter.addWidget(right_widget)
        
        # 设置分割器比例
        main_splitter.setSizes([400, 800])
        main_layout.addWidget(main_splitter)
        
        # 底部：日志区域
        log_group = QGroupBox("操作日志")
        log_layout = QVBoxLayout(log_group)
        log_layout.setContentsMargins(10, 10, 10, 10)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(150)
        self.log_text.setFont(QFont(self.config.FONT_FAMILY, self.config.FONT_SIZE_SMALL))
        log_layout.addWidget(self.log_text)
        
        main_layout.addWidget(log_group)
        
        # 状态栏
        self.status_bar = self.statusBar()
        self.status_bar.showMessage("就绪")
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.status_bar.addPermanentWidget(self.progress_bar)
    
    def create_file_manager_widget(self):
        """创建文件管理区部件"""
        group = QGroupBox("文件管理")
        layout = QVBoxLayout(group)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)
        
        # 文件列表
        file_label = QLabel("已选择文件:")
        layout.addWidget(file_label)
        
        self.file_list = QListWidget()
        self.file_list.setSelectionMode(QListWidget.ExtendedSelection)
        self.file_list.setMinimumHeight(200)
        layout.addWidget(self.file_list)
        
        # 文件信息标签
        self.file_info_label = QLabel("共 0 个文件")
        self.file_info_label.setStyleSheet(f"color: {self.config.TEXT_SECONDARY_COLOR};")
        layout.addWidget(self.file_info_label)
        
        # 文件操作按钮
        btn_layout = QHBoxLayout()
        
        self.add_file_btn = QPushButton("添加文件")
        self.add_file_btn.setMinimumHeight(35)
        btn_layout.addWidget(self.add_file_btn)
        
        self.add_folder_btn = QPushButton("添加文件夹")
        self.add_folder_btn.setMinimumHeight(35)
        btn_layout.addWidget(self.add_folder_btn)
        
        self.remove_file_btn = QPushButton("删除选中")
        self.remove_file_btn.setMinimumHeight(35)
        btn_layout.addWidget(self.remove_file_btn)
        
        self.clear_files_btn = QPushButton("清空列表")
        self.clear_files_btn.setMinimumHeight(35)
        btn_layout.addWidget(self.clear_files_btn)
        
        layout.addLayout(btn_layout)
        
        return group
    
    def create_function_widget(self):
        """创建功能操作区部件"""
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(15)
        
        # 标签页
        self.tab_widget = QTabWidget()
        self.tab_widget.setMinimumHeight(400)
        
        # Excel处理标签
        excel_tab = self.create_excel_tab()
        self.tab_widget.addTab(excel_tab, "Excel处理")
        
        # Word处理标签
        word_tab = self.create_word_tab()
        self.tab_widget.addTab(word_tab, "Word处理")
        
        # PDF处理标签
        pdf_tab = self.create_pdf_tab()
        self.tab_widget.addTab(pdf_tab, "PDF处理")
        
        # 财务统计标签
        finance_tab = self.create_finance_tab()
        self.tab_widget.addTab(finance_tab, "财务统计")
        
        # 简历筛选标签
        resume_tab = self.create_resume_tab()
        self.tab_widget.addTab(resume_tab, "简历筛选")
        
        layout.addWidget(self.tab_widget)
        
        scroll.setWidget(widget)
        return scroll
    
    def create_excel_tab(self):
        """创建Excel处理标签页"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)
        
        # 合并功能
        merge_group = QGroupBox("文件合并")
        merge_layout = QVBoxLayout(merge_group)
        merge_layout.setSpacing(8)
        
        merge_desc = QLabel("将多个Excel文件合并为一个文件（每个文件作为一个工作表）")
        merge_desc.setStyleSheet(f"color: {self.config.TEXT_SECONDARY_COLOR};")
        merge_desc.setWordWrap(True)
        merge_layout.addWidget(merge_desc)
        
        merge_btn_layout = QHBoxLayout()
        self.excel_merge_btn = QPushButton("开始合并")
        self.excel_merge_btn.setMinimumHeight(35)
        self.excel_merge_btn.setMaximumWidth(150)
        merge_btn_layout.addWidget(self.excel_merge_btn)
        merge_btn_layout.addStretch()
        merge_layout.addLayout(merge_btn_layout)
        
        layout.addWidget(merge_group)
        
        # 拆分功能
        split_group = QGroupBox("文件拆分")
        split_layout = QVBoxLayout(split_group)
        split_layout.setSpacing(8)
        
        split_desc = QLabel("将Excel文件按工作表拆分为多个独立文件")
        split_desc.setStyleSheet(f"color: {self.config.TEXT_SECONDARY_COLOR};")
        split_desc.setWordWrap(True)
        split_layout.addWidget(split_desc)
        
        split_btn_layout = QHBoxLayout()
        self.excel_split_btn = QPushButton("开始拆分")
        self.excel_split_btn.setMinimumHeight(35)
        self.excel_split_btn.setMaximumWidth(150)
        split_btn_layout.addWidget(self.excel_split_btn)
        split_btn_layout.addStretch()
        split_layout.addLayout(split_btn_layout)
        
        layout.addWidget(split_group)
        
        # 数据提取
        extract_group = QGroupBox("数据提取")
        extract_layout = QVBoxLayout(extract_group)
        extract_layout.setSpacing(8)
        
        extract_desc = QLabel("从Excel文件中提取指定范围的数据到新文件")
        extract_desc.setStyleSheet(f"color: {self.config.TEXT_SECONDARY_COLOR};")
        extract_desc.setWordWrap(True)
        extract_layout.addWidget(extract_desc)
        
        range_layout = QHBoxLayout()
        range_layout.addWidget(QLabel("工作表:"))
        self.excel_sheet_combo = QComboBox()
        self.excel_sheet_combo.setMinimumWidth(100)
        range_layout.addWidget(self.excel_sheet_combo)
        
        range_layout.addWidget(QLabel("起始行:"))
        self.excel_start_row = QSpinBox()
        self.excel_start_row.setMinimum(1)
        self.excel_start_row.setMaximum(100000)
        range_layout.addWidget(self.excel_start_row)
        
        range_layout.addWidget(QLabel("结束行:"))
        self.excel_end_row = QSpinBox()
        self.excel_end_row.setMinimum(1)
        self.excel_end_row.setMaximum(100000)
        self.excel_end_row.setValue(100)
        range_layout.addWidget(self.excel_end_row)
        range_layout.addStretch()
        extract_layout.addLayout(range_layout)
        
        extract_btn_layout = QHBoxLayout()
        self.excel_extract_btn = QPushButton("开始提取")
        self.excel_extract_btn.setMinimumHeight(35)
        self.excel_extract_btn.setMaximumWidth(150)
        extract_btn_layout.addWidget(self.excel_extract_btn)
        extract_btn_layout.addStretch()
        extract_layout.addLayout(extract_btn_layout)
        
        layout.addWidget(extract_group)
        
        # 报表生成
        report_group = QGroupBox("自动报表生成")
        report_layout = QVBoxLayout(report_group)
        report_layout.setSpacing(8)
        
        report_desc = QLabel("根据Excel数据自动生成统计报表（包含汇总、平均值等）")
        report_desc.setStyleSheet(f"color: {self.config.TEXT_SECONDARY_COLOR};")
        report_desc.setWordWrap(True)
        report_layout.addWidget(report_desc)
        
        report_btn_layout = QHBoxLayout()
        self.excel_report_btn = QPushButton("生成报表")
        self.excel_report_btn.setMinimumHeight(35)
        self.excel_report_btn.setMaximumWidth(150)
        report_btn_layout.addWidget(self.excel_report_btn)
        report_btn_layout.addStretch()
        report_layout.addLayout(report_btn_layout)
        
        layout.addWidget(report_group)
        layout.addStretch()
        
        return widget
    
    def create_word_tab(self):
        """创建Word处理标签页"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)
        
        # 合并功能
        merge_group = QGroupBox("文件合并")
        merge_layout = QVBoxLayout(merge_group)
        merge_layout.setSpacing(8)
        
        merge_desc = QLabel("将多个Word文件按顺序合并为一个文件")
        merge_desc.setStyleSheet(f"color: {self.config.TEXT_SECONDARY_COLOR};")
        merge_desc.setWordWrap(True)
        merge_layout.addWidget(merge_desc)
        
        merge_btn_layout = QHBoxLayout()
        self.word_merge_btn = QPushButton("开始合并")
        self.word_merge_btn.setMinimumHeight(35)
        self.word_merge_btn.setMaximumWidth(150)
        merge_btn_layout.addWidget(self.word_merge_btn)
        merge_btn_layout.addStretch()
        merge_layout.addLayout(merge_btn_layout)
        
        layout.addWidget(merge_group)
        
        # 拆分功能
        split_group = QGroupBox("文件拆分")
        split_layout = QVBoxLayout(split_group)
        split_layout.setSpacing(8)
        
        split_desc = QLabel("按指定页数将Word文件拆分为多个文件")
        split_desc.setStyleSheet(f"color: {self.config.TEXT_SECONDARY_COLOR};")
        split_desc.setWordWrap(True)
        split_layout.addWidget(split_desc)
        
        split_setting_layout = QHBoxLayout()
        split_setting_layout.addWidget(QLabel("每"))
        self.word_split_pages = QSpinBox()
        self.word_split_pages.setMinimum(1)
        self.word_split_pages.setMaximum(1000)
        self.word_split_pages.setValue(10)
        split_setting_layout.addWidget(self.word_split_pages)
        split_setting_layout.addWidget(QLabel("页为一个文件"))
        split_setting_layout.addStretch()
        split_layout.addLayout(split_setting_layout)
        
        split_btn_layout = QHBoxLayout()
        self.word_split_btn = QPushButton("开始拆分")
        self.word_split_btn.setMinimumHeight(35)
        self.word_split_btn.setMaximumWidth(150)
        split_btn_layout.addWidget(self.word_split_btn)
        split_btn_layout.addStretch()
        split_layout.addLayout(split_btn_layout)
        
        layout.addWidget(split_group)
        
        # 内容总结
        summary_group = QGroupBox("内容总结")
        summary_layout = QVBoxLayout(summary_group)
        summary_layout.setSpacing(8)
        
        summary_desc = QLabel("提取Word文档的关键内容进行总结（基于关键词和段落分析）")
        summary_desc.setStyleSheet(f"color: {self.config.TEXT_SECONDARY_COLOR};")
        summary_desc.setWordWrap(True)
        summary_layout.addWidget(summary_desc)
        
        summary_setting_layout = QHBoxLayout()
        summary_setting_layout.addWidget(QLabel("总结长度:"))
        self.word_summary_length = QComboBox()
        self.word_summary_length.addItems(["简短", "中等", "详细"])
        self.word_summary_length.setCurrentIndex(1)
        summary_setting_layout.addWidget(self.word_summary_length)
        summary_setting_layout.addStretch()
        summary_layout.addLayout(summary_setting_layout)
        
        summary_btn_layout = QHBoxLayout()
        self.word_summary_btn = QPushButton("生成总结")
        self.word_summary_btn.setMinimumHeight(35)
        self.word_summary_btn.setMaximumWidth(150)
        summary_btn_layout.addWidget(self.word_summary_btn)
        summary_btn_layout.addStretch()
        summary_layout.addLayout(summary_btn_layout)
        
        layout.addWidget(summary_group)
        
        # 文档分析
        analyze_group = QGroupBox("文档分析")
        analyze_layout = QVBoxLayout(analyze_group)
        analyze_layout.setSpacing(8)
        
        analyze_desc = QLabel("分析Word文档的字数、段落数、表格数等统计信息")
        analyze_desc.setStyleSheet(f"color: {self.config.TEXT_SECONDARY_COLOR};")
        analyze_desc.setWordWrap(True)
        analyze_layout.addWidget(analyze_desc)
        
        analyze_btn_layout = QHBoxLayout()
        self.word_analyze_btn = QPushButton("开始分析")
        self.word_analyze_btn.setMinimumHeight(35)
        self.word_analyze_btn.setMaximumWidth(150)
        analyze_btn_layout.addWidget(self.word_analyze_btn)
        analyze_btn_layout.addStretch()
        analyze_layout.addLayout(analyze_btn_layout)
        
        layout.addWidget(analyze_group)
        layout.addStretch()
        
        return widget
    
    def create_pdf_tab(self):
        """创建PDF处理标签页"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)
        
        # 合并功能
        merge_group = QGroupBox("文件合并")
        merge_layout = QVBoxLayout(merge_group)
        merge_layout.setSpacing(8)
        
        merge_desc = QLabel("将多个PDF文件按顺序合并为一个文件")
        merge_desc.setStyleSheet(f"color: {self.config.TEXT_SECONDARY_COLOR};")
        merge_desc.setWordWrap(True)
        merge_layout.addWidget(merge_desc)
        
        merge_btn_layout = QHBoxLayout()
        self.pdf_merge_btn = QPushButton("开始合并")
        self.pdf_merge_btn.setMinimumHeight(35)
        self.pdf_merge_btn.setMaximumWidth(150)
        merge_btn_layout.addWidget(self.pdf_merge_btn)
        merge_btn_layout.addStretch()
        merge_layout.addLayout(merge_btn_layout)
        
        layout.addWidget(merge_group)
        
        # 拆分功能
        split_group = QGroupBox("文件拆分")
        split_layout = QVBoxLayout(split_group)
        split_layout.setSpacing(8)
        
        split_desc = QLabel("按指定页数将PDF文件拆分为多个文件")
        split_desc.setStyleSheet(f"color: {self.config.TEXT_SECONDARY_COLOR};")
        split_desc.setWordWrap(True)
        split_layout.addWidget(split_desc)
        
        split_setting_layout = QHBoxLayout()
        split_setting_layout.addWidget(QLabel("起始页:"))
        self.pdf_start_page = QSpinBox()
        self.pdf_start_page.setMinimum(1)
        self.pdf_start_page.setMaximum(10000)
        split_setting_layout.addWidget(self.pdf_start_page)
        
        split_setting_layout.addWidget(QLabel("结束页:"))
        self.pdf_end_page = QSpinBox()
        self.pdf_end_page.setMinimum(1)
        self.pdf_end_page.setMaximum(10000)
        self.pdf_end_page.setValue(10)
        split_setting_layout.addWidget(self.pdf_end_page)
        split_setting_layout.addStretch()
        split_layout.addLayout(split_setting_layout)
        
        split_btn_layout = QHBoxLayout()
        self.pdf_split_btn = QPushButton("开始拆分")
        self.pdf_split_btn.setMinimumHeight(35)
        self.pdf_split_btn.setMaximumWidth(150)
        split_btn_layout.addWidget(self.pdf_split_btn)
        split_btn_layout.addStretch()
        split_layout.addLayout(split_btn_layout)
        
        layout.addWidget(split_group)
        
        # 内容总结
        summary_group = QGroupBox("内容总结")
        summary_layout = QVBoxLayout(summary_group)
        summary_layout.setSpacing(8)
        
        summary_desc = QLabel("提取PDF文档的文本内容进行总结")
        summary_desc.setStyleSheet(f"color: {self.config.TEXT_SECONDARY_COLOR};")
        summary_desc.setWordWrap(True)
        summary_layout.addWidget(summary_desc)
        
        summary_btn_layout = QHBoxLayout()
        self.pdf_summary_btn = QPushButton("生成总结")
        self.pdf_summary_btn.setMinimumHeight(35)
        self.pdf_summary_btn.setMaximumWidth(150)
        summary_btn_layout.addWidget(self.pdf_summary_btn)
        summary_btn_layout.addStretch()
        summary_layout.addLayout(summary_btn_layout)
        
        layout.addWidget(summary_group)
        
        # 数据提取
        extract_group = QGroupBox("数据提取")
        extract_layout = QVBoxLayout(extract_group)
        extract_layout.setSpacing(8)
        
        extract_desc = QLabel("从PDF中提取文本内容并保存为文本文件")
        extract_desc.setStyleSheet(f"color: {self.config.TEXT_SECONDARY_COLOR};")
        extract_desc.setWordWrap(True)
        extract_layout.addWidget(extract_desc)
        
        extract_btn_layout = QHBoxLayout()
        self.pdf_extract_btn = QPushButton("开始提取")
        self.pdf_extract_btn.setMinimumHeight(35)
        self.pdf_extract_btn.setMaximumWidth(150)
        extract_btn_layout.addWidget(self.pdf_extract_btn)
        extract_btn_layout.addStretch()
        extract_layout.addLayout(extract_btn_layout)
        
        layout.addWidget(extract_group)
        layout.addStretch()
        
        return widget
    
    def create_finance_tab(self):
        """创建财务统计标签页"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)
        
        # 财务统计说明
        desc_group = QGroupBox("功能说明")
        desc_layout = QVBoxLayout(desc_group)
        desc_label = QLabel(
            "财务统计功能支持：\n"
            "1. 从Excel文件中读取财务数据\n"
            "2. 自动计算收入、支出、利润等统计指标\n"
            "3. 生成财务统计报表\n"
            "4. 支持按月份、季度进行分组统计"
        )
        desc_label.setStyleSheet(f"color: {self.config.TEXT_SECONDARY_COLOR};")
        desc_layout.addWidget(desc_label)
        layout.addWidget(desc_group)
        
        # 统计设置
        setting_group = QGroupBox("统计设置")
        setting_layout = QVBoxLayout(setting_group)
        setting_layout.setSpacing(8)
        
        # 数据列设置
        col_layout1 = QHBoxLayout()
        col_layout1.addWidget(QLabel("日期列名:"))
        self.finance_date_col = QLineEdit("日期")
        self.finance_date_col.setMaximumWidth(150)
        col_layout1.addWidget(self.finance_date_col)
        
        col_layout1.addWidget(QLabel("收入列名:"))
        self.finance_income_col = QLineEdit("收入")
        self.finance_income_col.setMaximumWidth(150)
        col_layout1.addWidget(self.finance_income_col)
        col_layout1.addStretch()
        setting_layout.addLayout(col_layout1)
        
        col_layout2 = QHBoxLayout()
        col_layout2.addWidget(QLabel("支出列名:"))
        self.finance_expense_col = QLineEdit("支出")
        self.finance_expense_col.setMaximumWidth(150)
        col_layout2.addWidget(self.finance_expense_col)
        
        col_layout2.addWidget(QLabel("分组方式:"))
        self.finance_group_by = QComboBox()
        self.finance_group_by.addItems(["按月份", "按季度", "按年份"])
        col_layout2.addWidget(self.finance_group_by)
        col_layout2.addStretch()
        setting_layout.addLayout(col_layout2)
        
        layout.addWidget(setting_group)
        
        # 统计按钮
        btn_layout = QHBoxLayout()
        self.finance_analyze_btn = QPushButton("开始统计分析")
        self.finance_analyze_btn.setMinimumHeight(40)
        self.finance_analyze_btn.setMaximumWidth(200)
        btn_layout.addWidget(self.finance_analyze_btn)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)
        
        layout.addStretch()
        
        return widget
    
    def create_resume_tab(self):
        """创建简历筛选标签页"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)
        
        # 功能说明
        desc_group = QGroupBox("功能说明")
        desc_layout = QVBoxLayout(desc_group)
        desc_label = QLabel(
            "简历筛选功能支持：\n"
            "1. 批量读取Word/PDF格式的简历文件\n"
            "2. 根据关键词自动筛选符合条件的简历\n"
            "3. 提取简历中的关键信息（学历、工作经验、技能等）\n"
            "4. 生成筛选结果报告"
        )
        desc_label.setStyleSheet(f"color: {self.config.TEXT_SECONDARY_COLOR};")
        desc_layout.addWidget(desc_label)
        layout.addWidget(desc_group)
        
        # 筛选条件
        filter_group = QGroupBox("筛选条件")
        filter_layout = QVBoxLayout(filter_group)
        filter_layout.setSpacing(8)
        
        # 关键词
        keyword_layout = QHBoxLayout()
        keyword_layout.addWidget(QLabel("必备关键词 (逗号分隔):"))
        self.resume_keywords = QLineEdit()
        self.resume_keywords.setPlaceholderText("例如: Python, 数据分析, 项目管理")
        keyword_layout.addWidget(self.resume_keywords)
        filter_layout.addLayout(keyword_layout)
        
        # 排除关键词
        exclude_layout = QHBoxLayout()
        exclude_layout.addWidget(QLabel("排除关键词 (逗号分隔):"))
        self.resume_exclude_keywords = QLineEdit()
        self.resume_exclude_keywords.setPlaceholderText("例如: 应届生, 无经验")
        exclude_layout.addWidget(self.resume_exclude_keywords)
        filter_layout.addLayout(exclude_layout)
        
        # 最低学历要求
        edu_layout = QHBoxLayout()
        edu_layout.addWidget(QLabel("最低学历要求:"))
        self.resume_min_education = QComboBox()
        self.resume_min_education.addItems(["不限", "大专", "本科", "硕士", "博士"])
        edu_layout.addWidget(self.resume_min_education)
        
        edu_layout.addWidget(QLabel("最低工作年限:"))
        self.resume_min_experience = QSpinBox()
        self.resume_min_experience.setMinimum(0)
        self.resume_min_experience.setMaximum(30)
        self.resume_min_experience.setSuffix(" 年")
        edu_layout.addWidget(self.resume_min_experience)
        edu_layout.addStretch()
        filter_layout.addLayout(edu_layout)
        
        layout.addWidget(filter_group)
        
        # 筛选按钮
        btn_layout = QHBoxLayout()
        self.resume_filter_btn = QPushButton("开始筛选简历")
        self.resume_filter_btn.setMinimumHeight(40)
        self.resume_filter_btn.setMaximumWidth(200)
        btn_layout.addWidget(self.resume_filter_btn)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)
        
        layout.addStretch()
        
        return widget
    
    def connect_signals(self):
        """连接信号槽"""
        # 文件管理信号
        self.add_file_btn.clicked.connect(self.add_files)
        self.add_folder_btn.clicked.connect(self.add_folder)
        self.remove_file_btn.clicked.connect(self.remove_selected_files)
        self.clear_files_btn.clicked.connect(self.clear_files)
        self.file_list.itemSelectionChanged.connect(self.on_file_selection_changed)
        
        # Excel处理信号
        self.excel_merge_btn.clicked.connect(lambda: self.run_task(self.excel_merge_task))
        self.excel_split_btn.clicked.connect(lambda: self.run_task(self.excel_split_task))
        self.excel_extract_btn.clicked.connect(lambda: self.run_task(self.excel_extract_task))
        self.excel_report_btn.clicked.connect(lambda: self.run_task(self.excel_report_task))
        
        # Word处理信号
        self.word_merge_btn.clicked.connect(lambda: self.run_task(self.word_merge_task))
        self.word_split_btn.clicked.connect(lambda: self.run_task(self.word_split_task))
        self.word_summary_btn.clicked.connect(lambda: self.run_task(self.word_summary_task))
        self.word_analyze_btn.clicked.connect(lambda: self.run_task(self.word_analyze_task))
        
        # PDF处理信号
        self.pdf_merge_btn.clicked.connect(lambda: self.run_task(self.pdf_merge_task))
        self.pdf_split_btn.clicked.connect(lambda: self.run_task(self.pdf_split_task))
        self.pdf_summary_btn.clicked.connect(lambda: self.run_task(self.pdf_summary_task))
        self.pdf_extract_btn.clicked.connect(lambda: self.run_task(self.pdf_extract_task))
        
        # 财务统计信号
        self.finance_analyze_btn.clicked.connect(lambda: self.run_task(self.finance_analyze_task))
        
        # 简历筛选信号
        self.resume_filter_btn.clicked.connect(lambda: self.run_task(self.resume_filter_task))
    
    def add_files(self):
        """添加文件"""
        file_filter = "支持的文件 (*.xlsx *.xls *.docx *.doc *.pdf);;Excel文件 (*.xlsx *.xls);;Word文件 (*.docx *.doc);;PDF文件 (*.pdf);;所有文件 (*.*)"
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择文件", "", file_filter
        )
        if files:
            for file_path in files:
                if self.file_manager.add_file(file_path):
                    self.update_file_list()
            self.log(f"已添加 {len(files)} 个文件")
    
    def add_folder(self):
        """添加文件夹"""
        folder_path = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if folder_path:
            count = 0
            for root, dirs, files in os.walk(folder_path):
                for file_name in files:
                    file_path = os.path.join(root, file_name)
                    if get_file_extension(file_path) in self.config.SUPPORTED_EXTENSIONS:
                        if self.file_manager.add_file(file_path):
                            count += 1
            self.update_file_list()
            self.log(f"从文件夹添加了 {count} 个文件")
    
    def remove_selected_files(self):
        """删除选中的文件"""
        selected_items = self.file_list.selectedItems()
        if not selected_items:
            show_message(self, "提示", "请先选择要删除的文件")
            return
        
        reply = QMessageBox.question(
            self, "确认删除",
            f"确定要删除选中的 {len(selected_items)} 个文件吗？",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            for item in selected_items:
                file_path = item.data(Qt.UserRole)
                self.file_manager.remove_file(file_path)
            self.update_file_list()
            self.log(f"已删除 {len(selected_items)} 个文件")
    
    def clear_files(self):
        """清空文件列表"""
        if not self.file_manager.get_files():
            show_message(self, "提示", "文件列表已经为空")
            return
        
        reply = QMessageBox.question(
            self, "确认清空",
            "确定要清空文件列表吗？",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            self.file_manager.clear_files()
            self.update_file_list()
            self.log("已清空文件列表")
    
    def update_file_list(self):
        """更新文件列表显示"""
        self.file_list.clear()
        files = self.file_manager.get_files()
        
        for file_path in files:
            file_name = os.path.basename(file_path)
            file_size = os.path.getsize(file_path) if os.path.exists(file_path) else 0
            size_str = format_file_size(file_size)
            
            item = QListWidgetItem(f"{file_name} ({size_str})")
            item.setData(Qt.UserRole, file_path)
            self.file_list.addItem(item)
        
        self.file_info_label.setText(f"共 {len(files)} 个文件")
    
    def on_file_selection_changed(self):
        """文件选择改变时更新工作表列表"""
        selected_items = self.file_list.selectedItems()
        if selected_items:
            file_path = selected_items[0].data(Qt.UserRole)
            ext = get_file_extension(file_path)
            
            if ext in self.config.EXCEL_EXTENSIONS:
                try:
                    sheets = self.excel_processor.get_sheet_names(file_path)
                    self.excel_sheet_combo.clear()
                    self.excel_sheet_combo.addItems(sheets)
                except:
                    pass
    
    def log(self, message):
        """添加日志"""
        from datetime import datetime
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_message = f"[{timestamp}] {message}"
        self.log_text.append(log_message)
        self.status_bar.showMessage(message)
    
    def run_task(self, task_func):
        """运行后台任务"""
        if self.worker_thread and self.worker_thread.isRunning():
            show_message(self, "提示", "请等待当前操作完成")
            return
        
        if not self.file_manager.get_files():
            show_message(self, "提示", "请先添加文件")
            return
        
        self.worker_thread = WorkerThread(task_func)
        self.worker_thread.log_signal.connect(self.log)
        self.worker_thread.finished_signal.connect(self.on_task_finished)
        self.worker_thread.progress_signal.connect(self.on_progress_changed)
        
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.set_buttons_enabled(False)
        
        self.worker_thread.start()
    
    def on_task_finished(self, success, message):
        """任务完成回调"""
        self.progress_bar.setVisible(False)
        self.set_buttons_enabled(True)
        
        if success:
            show_message(self, "成功", message, QMessageBox.Information)
        else:
            show_message(self, "错误", message, QMessageBox.Critical)
    
    def on_progress_changed(self, value):
        """进度改变回调"""
        self.progress_bar.setValue(value)
    
    def set_buttons_enabled(self, enabled):
        """设置按钮启用状态"""
        buttons = [
            self.add_file_btn, self.add_folder_btn, self.remove_file_btn,
            self.clear_files_btn, self.excel_merge_btn, self.excel_split_btn,
            self.excel_extract_btn, self.excel_report_btn, self.word_merge_btn,
            self.word_split_btn, self.word_summary_btn, self.word_analyze_btn,
            self.pdf_merge_btn, self.pdf_split_btn, self.pdf_summary_btn,
            self.pdf_extract_btn, self.finance_analyze_btn, self.resume_filter_btn
        ]
        for btn in buttons:
            btn.setEnabled(enabled)
    
    # 任务函数
    def excel_merge_task(self, log_callback, progress_callback):
        """Excel合并任务"""
        files = self.file_manager.get_files_by_extensions(self.config.EXCEL_EXTENSIONS)
        if not files:
            raise Exception("没有找到Excel文件")
        
        log_callback(f"开始合并 {len(files)} 个Excel文件...")
        progress_callback(10)
        
        output_path = self.excel_processor.merge_files(files, log_callback)
        progress_callback(100)
        
        return f"合并完成！输出文件: {output_path}"
    
    def excel_split_task(self, log_callback, progress_callback):
        """Excel拆分任务"""
        files = self.file_manager.get_files_by_extensions(self.config.EXCEL_EXTENSIONS)
        if not files:
            raise Exception("没有找到Excel文件")
        
        log_callback(f"开始拆分 {len(files)} 个Excel文件...")
        progress_callback(10)
        
        output_paths = []
        for i, file_path in enumerate(files):
            paths = self.excel_processor.split_file(file_path, log_callback)
            output_paths.extend(paths)
            progress_callback(int((i + 1) / len(files) * 100))
        
        return f"拆分完成！共生成 {len(output_paths)} 个文件"
    
    def excel_extract_task(self, log_callback, progress_callback):
        """Excel数据提取任务"""
        files = self.file_manager.get_files_by_extensions(self.config.EXCEL_EXTENSIONS)
        if not files:
            raise Exception("没有找到Excel文件")
        
        sheet_name = self.excel_sheet_combo.currentText()
        start_row = self.excel_start_row.value()
        end_row = self.excel_end_row.value()
        
        if start_row > end_row:
            raise Exception("起始行不能大于结束行")
        
        log_callback(f"开始提取数据: 工作表={sheet_name}, 行范围={start_row}-{end_row}")
        progress_callback(10)
        
        output_path = self.excel_processor.extract_data(
            files[0], sheet_name, start_row, end_row, log_callback
        )
        progress_callback(100)
        
        return f"数据提取完成！输出文件: {output_path}"
    
    def excel_report_task(self, log_callback, progress_callback):
        """Excel报表生成任务"""
        files = self.file_manager.get_files_by_extensions(self.config.EXCEL_EXTENSIONS)
        if not files:
            raise Exception("没有找到Excel文件")
        
        log_callback(f"开始为 {len(files)} 个文件生成报表...")
        progress_callback(10)
        
        output_paths = []
        for i, file_path in enumerate(files):
            path = self.excel_processor.generate_report(file_path, log_callback)
            output_paths.append(path)
            progress_callback(int((i + 1) / len(files) * 100))
        
        return f"报表生成完成！共生成 {len(output_paths)} 个报表文件"
    
    def word_merge_task(self, log_callback, progress_callback):
        """Word合并任务"""
        files = self.file_manager.get_files_by_extensions(self.config.WORD_EXTENSIONS)
        if not files:
            raise Exception("没有找到Word文件")
        
        log_callback(f"开始合并 {len(files)} 个Word文件...")
        progress_callback(10)
        
        output_path = self.word_processor.merge_files(files, log_callback)
        progress_callback(100)
        
        return f"合并完成！输出文件: {output_path}"
    
    def word_split_task(self, log_callback, progress_callback):
        """Word拆分任务"""
        files = self.file_manager.get_files_by_extensions(self.config.WORD_EXTENSIONS)
        if not files:
            raise Exception("没有找到Word文件")
        
        pages_per_file = self.word_split_pages.value()
        
        log_callback(f"开始拆分Word文件，每 {pages_per_file} 页为一个文件...")
        progress_callback(10)
        
        output_paths = self.word_processor.split_file(files[0], pages_per_file, log_callback)
        progress_callback(100)
        
        return f"拆分完成！共生成 {len(output_paths)} 个文件"
    
    def word_summary_task(self, log_callback, progress_callback):
        """Word内容总结任务"""
        files = self.file_manager.get_files_by_extensions(self.config.WORD_EXTENSIONS)
        if not files:
            raise Exception("没有找到Word文件")
        
        summary_length = self.word_summary_length.currentText()
        
        log_callback(f"开始生成内容总结（{summary_length}）...")
        progress_callback(10)
        
        output_path = self.word_processor.generate_summary(files[0], summary_length, log_callback)
        progress_callback(100)
        
        return f"总结生成完成！输出文件: {output_path}"
    
    def word_analyze_task(self, log_callback, progress_callback):
        """Word文档分析任务"""
        files = self.file_manager.get_files_by_extensions(self.config.WORD_EXTENSIONS)
        if not files:
            raise Exception("没有找到Word文件")
        
        log_callback("开始分析文档...")
        progress_callback(10)
        
        analysis = self.word_processor.analyze_document(files[0], log_callback)
        progress_callback(100)
        
        analysis_str = "\n".join([f"{k}: {v}" for k, v in analysis.items()])
        log_callback(f"分析结果:\n{analysis_str}")
        
        return f"文档分析完成！\n{analysis_str}"
    
    def pdf_merge_task(self, log_callback, progress_callback):
        """PDF合并任务"""
        files = self.file_manager.get_files_by_extensions(self.config.PDF_EXTENSIONS)
        if not files:
            raise Exception("没有找到PDF文件")
        
        log_callback(f"开始合并 {len(files)} 个PDF文件...")
        progress_callback(10)
        
        output_path = self.pdf_processor.merge_files(files, log_callback)
        progress_callback(100)
        
        return f"合并完成！输出文件: {output_path}"
    
    def pdf_split_task(self, log_callback, progress_callback):
        """PDF拆分任务"""
        files = self.file_manager.get_files_by_extensions(self.config.PDF_EXTENSIONS)
        if not files:
            raise Exception("没有找到PDF文件")
        
        start_page = self.pdf_start_page.value()
        end_page = self.pdf_end_page.value()
        
        if start_page > end_page:
            raise Exception("起始页不能大于结束页")
        
        log_callback(f"开始拆分PDF文件，页码范围: {start_page}-{end_page}")
        progress_callback(10)
        
        output_path = self.pdf_processor.split_file(files[0], start_page, end_page, log_callback)
        progress_callback(100)
        
        return f"拆分完成！输出文件: {output_path}"
    
    def pdf_summary_task(self, log_callback, progress_callback):
        """PDF内容总结任务"""
        files = self.file_manager.get_files_by_extensions(self.config.PDF_EXTENSIONS)
        if not files:
            raise Exception("没有找到PDF文件")
        
        log_callback("开始生成PDF内容总结...")
        progress_callback(10)
        
        output_path = self.pdf_processor.generate_summary(files[0], log_callback)
        progress_callback(100)
        
        return f"总结生成完成！输出文件: {output_path}"
    
    def pdf_extract_task(self, log_callback, progress_callback):
        """PDF数据提取任务"""
        files = self.file_manager.get_files_by_extensions(self.config.PDF_EXTENSIONS)
        if not files:
            raise Exception("没有找到PDF文件")
        
        log_callback("开始提取PDF文本内容...")
        progress_callback(10)
        
        output_path = self.pdf_processor.extract_text(files[0], log_callback)
        progress_callback(100)
        
        return f"文本提取完成！输出文件: {output_path}"
    
    def finance_analyze_task(self, log_callback, progress_callback):
        """财务统计任务"""
        files = self.file_manager.get_files_by_extensions(self.config.EXCEL_EXTENSIONS)
        if not files:
            raise Exception("没有找到Excel文件")
        
        settings = {
            'date_col': self.finance_date_col.text(),
            'income_col': self.finance_income_col.text(),
            'expense_col': self.finance_expense_col.text(),
            'group_by': self.finance_group_by.currentText()
        }
        
        log_callback("开始进行财务统计分析...")
        progress_callback(10)
        
        output_path = self.finance_analyzer.analyze(files[0], settings, log_callback)
        progress_callback(100)
        
        return f"财务统计完成！输出文件: {output_path}"
    
    def resume_filter_task(self, log_callback, progress_callback):
        """简历筛选任务"""
        files = self.file_manager.get_files()
        if not files:
            raise Exception("没有找到文件")
        
        keywords_text = self.resume_keywords.text().strip()
        exclude_keywords_text = self.resume_exclude_keywords.text().strip()
        
        settings = {
            'keywords': [k.strip() for k in keywords_text.split(',') if k.strip()] if keywords_text else [],
            'exclude_keywords': [k.strip() for k in exclude_keywords_text.split(',') if k.strip()] if exclude_keywords_text else [],
            'min_education': self.resume_min_education.currentText(),
            'min_experience': self.resume_min_experience.value()
        }
        
        log_callback(f"开始筛选 {len(files)} 份简历...")
        progress_callback(10)
        
        result = self.resume_filter.filter_resumes(files, settings, log_callback)
        progress_callback(100)
        
        return f"简历筛选完成！符合条件: {result['matched_count']} 份，不符合条件: {result['unmatched_count']} 份"
    
    def get_stylesheet(self):
        """获取样式表"""
        return f"""
            QMainWindow {{
                background-color: {self.config.BACKGROUND_COLOR};
            }}
            QWidget {{
                font-family: "{self.config.FONT_FAMILY}";
                font-size: {self.config.FONT_SIZE}pt;
                color: {self.config.TEXT_COLOR};
            }}
            QGroupBox {{
                background-color: {self.config.SURFACE_COLOR};
                border: 1px solid #E0E0E0;
                border-radius: 5px;
                margin-top: 10px;
                padding-top: 10px;
                font-weight: bold;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }}
            QPushButton {{
                background-color: {self.config.PRIMARY_COLOR};
                color: white;
                border: none;
                border-radius: 5px;
                padding: 8px 16px;
                font-weight: bold;
            }}
            QPushButton:hover {{
                background-color: #1976D2;
            }}
            QPushButton:pressed {{
                background-color: #1565C0;
            }}
            QPushButton:disabled {{
                background-color: #BDBDBD;
                color: #757575;
            }}
            QListWidget {{
                background-color: {self.config.SURFACE_COLOR};
                border: 1px solid #E0E0E0;
                border-radius: 5px;
                padding: 5px;
            }}
            QListWidget::item {{
                padding: 5px;
                border-radius: 3px;
            }}
            QListWidget::item:selected {{
                background-color: {self.config.PRIMARY_COLOR};
                color: white;
            }}
            QListWidget::item:hover {{
                background-color: #E3F2FD;
            }}
            QTextEdit {{
                background-color: {self.config.SURFACE_COLOR};
                border: 1px solid #E0E0E0;
                border-radius: 5px;
            }}
            QTabWidget::pane {{
                border: 1px solid #E0E0E0;
                border-radius: 5px;
                background-color: {self.config.SURFACE_COLOR};
            }}
            QTabBar::tab {{
                background-color: #EEEEEE;
                border: 1px solid #E0E0E0;
                border-bottom: none;
                border-top-left-radius: 5px;
                border-top-right-radius: 5px;
                padding: 8px 16px;
                margin-right: 2px;
            }}
            QTabBar::tab:selected {{
                background-color: {self.config.SURFACE_COLOR};
                border-bottom: 1px solid {self.config.SURFACE_COLOR};
            }}
            QTabBar::tab:hover:!selected {{
                background-color: #E0E0E0;
            }}
            QSpinBox, QComboBox, QLineEdit {{
                background-color: {self.config.SURFACE_COLOR};
                border: 1px solid #E0E0E0;
                border-radius: 3px;
                padding: 5px;
            }}
            QSpinBox:focus, QComboBox:focus, QLineEdit:focus {{
                border-color: {self.config.PRIMARY_COLOR};
            }}
            QScrollBar:vertical {{
                background-color: #F5F5F5;
                width: 12px;
                border-radius: 6px;
            }}
            QScrollBar::handle:vertical {{
                background-color: #BDBDBD;
                border-radius: 6px;
                min-height: 20px;
            }}
            QScrollBar::handle:vertical:hover {{
                background-color: #9E9E9E;
            }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
                height: 0px;
            }}
            QScrollBar:horizontal {{
                background-color: #F5F5F5;
                height: 12px;
                border-radius: 6px;
            }}
            QScrollBar::handle:horizontal {{
                background-color: #BDBDBD;
                border-radius: 6px;
                min-width: 20px;
            }}
            QScrollBar::handle:horizontal:hover {{
                background-color: #9E9E9E;
            }}
            QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{
                width: 0px;
            }}
            QSplitter::handle {{
                background-color: #E0E0E0;
            }}
            QSplitter::handle:horizontal {{
                width: 2px;
            }}
            QSplitter::handle:vertical {{
                height: 2px;
            }}
            QStatusBar {{
                background-color: {self.config.SURFACE_COLOR};
                border-top: 1px solid #E0E0E0;
            }}
            QProgressBar {{
                border: none;
                background-color: #E0E0E0;
                height: 8px;
                text-align: center;
            }}
            QProgressBar::chunk {{
                background-color: {self.config.SUCCESS_COLOR};
            }}
        """
