import os
from datetime import datetime
from PyQt5.QtWidgets import QMessageBox


def get_file_extension(file_path):
    """获取文件扩展名"""
    return os.path.splitext(file_path)[1].lower()


def generate_output_filename(input_path, suffix, extension=None):
    """生成输出文件名"""
    dir_name = os.path.dirname(input_path)
    base_name = os.path.splitext(os.path.basename(input_path))[0]
    if extension is None:
        extension = get_file_extension(input_path)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return os.path.join(dir_name, f"{base_name}_{suffix}_{timestamp}{extension}")


def show_message(parent, title, message, icon=QMessageBox.Information):
    """显示消息框"""
    msg = QMessageBox(parent)
    msg.setWindowTitle(title)
    msg.setText(message)
    msg.setIcon(icon)
    msg.exec_()


def format_file_size(size_bytes):
    """格式化文件大小"""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.2f} KB"
    elif size_bytes < 1024 * 1024 * 1024:
        return f"{size_bytes / (1024 * 1024):.2f} MB"
    else:
        return f"{size_bytes / (1024 * 1024 * 1024):.2f} GB"
