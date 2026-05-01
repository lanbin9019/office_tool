class Config:
    """应用程序配置类"""
    
    APP_NAME = "智能办公自动化工具"
    APP_VERSION = "1.0.0"
    WINDOW_WIDTH = 1200
    WINDOW_HEIGHT = 800
    
    SUPPORTED_EXTENSIONS = {
        '.xlsx', '.xls', '.docx', '.doc', '.pdf'
    }
    
    EXCEL_EXTENSIONS = {'.xlsx', '.xls'}
    WORD_EXTENSIONS = {'.docx', '.doc'}
    PDF_EXTENSIONS = {'.pdf'}
    
    DEFAULT_OUTPUT_DIR = "output"
    
    # 样式配置
    PRIMARY_COLOR = "#2196F3"
    SECONDARY_COLOR = "#757575"
    SUCCESS_COLOR = "#4CAF50"
    WARNING_COLOR = "#FF9800"
    ERROR_COLOR = "#F44336"
    BACKGROUND_COLOR = "#FAFAFA"
    SURFACE_COLOR = "#FFFFFF"
    TEXT_COLOR = "#212121"
    TEXT_SECONDARY_COLOR = "#757575"
    
    # 字体配置
    FONT_FAMILY = "Microsoft YaHei"
    FONT_SIZE = 10
    FONT_SIZE_LARGE = 12
    FONT_SIZE_SMALL = 9
