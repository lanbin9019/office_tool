import os
from typing import List, Dict, Any, Optional
from datetime import datetime

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

from office_tool.utils import generate_output_filename, get_file_extension


class ExcelProcessor:
    """Excel处理器类"""
    
    def __init__(self):
        self._header_font = Font(bold=True, size=11)
        self._header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        self._header_font_white = Font(bold=True, size=11, color="FFFFFF")
        self._alignment = Alignment(horizontal='center', vertical='center')
        self._thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
    
    def get_sheet_names(self, file_path: str) -> List[str]:
        """
        获取Excel文件中的所有工作表名称
        
        Args:
            file_path: Excel文件路径
            
        Returns:
            工作表名称列表
        """
        ext = get_file_extension(file_path)
        if ext == '.xls':
            df = pd.ExcelFile(file_path)
            return df.sheet_names
        else:
            wb = load_workbook(file_path, read_only=True)
            sheet_names = wb.sheetnames
            wb.close()
            return sheet_names
    
    def merge_files(self, file_paths: List[str], log_callback=None) -> str:
        """
        合并多个Excel文件为一个文件（每个文件作为一个工作表）
        
        Args:
            file_paths: Excel文件路径列表
            log_callback: 日志回调函数
            
        Returns:
            输出文件路径
        """
        if len(file_paths) < 1:
            raise ValueError("至少需要一个Excel文件")
        
        output_path = generate_output_filename(file_paths[0], "merged")
        
        wb = Workbook()
        default_sheet = wb.active
        wb.remove(default_sheet)
        
        for i, file_path in enumerate(file_paths):
            file_name = os.path.basename(file_path)
            sheet_name = f"Sheet_{i+1}_{os.path.splitext(file_name)[0][:20]}"
            
            if log_callback:
                log_callback(f"正在处理: {file_name}")
            
            ext = get_file_extension(file_path)
            if ext == '.xls':
                xl = pd.ExcelFile(file_path)
                for sheet_idx, sheet_name_orig in enumerate(xl.sheet_names):
                    df = pd.read_excel(file_path, sheet_name=sheet_name_orig)
                    final_sheet_name = f"{sheet_name}_{sheet_idx+1}" if len(xl.sheet_names) > 1 else sheet_name
                    self._write_df_to_sheet(wb, final_sheet_name, df)
            else:
                wb_src = load_workbook(file_path, data_only=True)
                for sheet_idx, sheet_name_orig in enumerate(wb_src.sheetnames):
                    ws_src = wb_src[sheet_name_orig]
                    df = self._sheet_to_dataframe(ws_src)
                    final_sheet_name = f"{sheet_name}_{sheet_idx+1}" if len(wb_src.sheetnames) > 1 else sheet_name
                    self._write_df_to_sheet(wb, final_sheet_name, df)
                wb_src.close()
        
        wb.save(output_path)
        wb.close()
        
        if log_callback:
            log_callback(f"合并完成，共 {len(file_paths)} 个文件")
        
        return output_path
    
    def split_file(self, file_path: str, log_callback=None) -> List[str]:
        """
        按工作表拆分Excel文件
        
        Args:
            file_path: Excel文件路径
            log_callback: 日志回调函数
            
        Returns:
            输出文件路径列表
        """
        output_paths = []
        base_dir = os.path.dirname(file_path)
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        ext = get_file_extension(file_path)
        if ext == '.xls':
            xl = pd.ExcelFile(file_path)
            sheet_names = xl.sheet_names
        else:
            wb = load_workbook(file_path, read_only=True)
            sheet_names = wb.sheetnames
            wb.close()
        
        for i, sheet_name in enumerate(sheet_names):
            if log_callback:
                log_callback(f"正在拆分工作表: {sheet_name}")
            
            if ext == '.xls':
                df = pd.read_excel(file_path, sheet_name=sheet_name)
            else:
                wb_src = load_workbook(file_path, data_only=True)
                ws_src = wb_src[sheet_name]
                df = self._sheet_to_dataframe(ws_src)
                wb_src.close()
            
            safe_sheet_name = "".join(c for c in sheet_name if c.isalnum() or c in (' ', '_', '-')).strip()
            output_name = f"{base_name}_{safe_sheet_name}_{timestamp}{ext if ext == '.xlsx' else '.xlsx'}"
            output_path = os.path.join(base_dir, output_name)
            
            wb_out = Workbook()
            ws_out = wb_out.active
            ws_out.title = sheet_name if len(sheet_name) <= 31 else sheet_name[:31]
            
            for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
                for c_idx, value in enumerate(row, 1):
                    cell = ws_out.cell(row=r_idx, column=c_idx, value=value)
                    if r_idx == 1:
                        cell.font = self._header_font
                        cell.alignment = self._alignment
                        cell.fill = self._header_fill
                        cell.font = self._header_font_white
            
            for col in ws_out.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                ws_out.column_dimensions[column].width = adjusted_width
            
            wb_out.save(output_path)
            wb_out.close()
            output_paths.append(output_path)
        
        if log_callback:
            log_callback(f"拆分完成，共生成 {len(output_paths)} 个文件")
        
        return output_paths
    
    def extract_data(self, file_path: str, sheet_name: str, start_row: int, end_row: int, log_callback=None) -> str:
        """
        从Excel中提取指定范围的数据
        
        Args:
            file_path: Excel文件路径
            sheet_name: 工作表名称
            start_row: 起始行（从1开始）
            end_row: 结束行
            log_callback: 日志回调函数
            
        Returns:
            输出文件路径
        """
        if log_callback:
            log_callback(f"正在从 {sheet_name} 提取行 {start_row}-{end_row}")
        
        ext = get_file_extension(file_path)
        if ext == '.xls':
            df = pd.read_excel(file_path, sheet_name=sheet_name)
        else:
            wb = load_workbook(file_path, data_only=True)
            ws = wb[sheet_name]
            df = self._sheet_to_dataframe(ws)
            wb.close()
        
        start_idx = max(0, start_row - 1)
        end_idx = min(len(df), end_row)
        
        if start_idx >= len(df):
            raise ValueError(f"起始行 {start_row} 超出数据范围")
        
        extracted_df = df.iloc[start_idx:end_idx]
        
        output_path = generate_output_filename(file_path, "extracted")
        
        wb_out = Workbook()
        ws_out = wb_out.active
        ws_out.title = "ExtractedData"
        
        for r_idx, row in enumerate(dataframe_to_rows(extracted_df, index=False, header=True), 1):
            for c_idx, value in enumerate(row, 1):
                cell = ws_out.cell(row=r_idx, column=c_idx, value=value)
                if r_idx == 1:
                    cell.font = self._header_font
                    cell.alignment = self._alignment
        
        wb_out.save(output_path)
        wb_out.close()
        
        if log_callback:
            log_callback(f"数据提取完成，共 {len(extracted_df)} 行数据")
        
        return output_path
    
    def generate_report(self, file_path: str, log_callback=None) -> str:
        """
        自动生成统计报表
        
        Args:
            file_path: Excel文件路径
            log_callback: 日志回调函数
            
        Returns:
            输出文件路径
        """
        if log_callback:
            log_callback("正在生成统计报表...")
        
        ext = get_file_extension(file_path)
        
        if ext == '.xls':
            xl = pd.ExcelFile(file_path)
            dfs = {name: pd.read_excel(file_path, sheet_name=name) for name in xl.sheet_names}
        else:
            wb = load_workbook(file_path, data_only=True)
            dfs = {}
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                dfs[sheet_name] = self._sheet_to_dataframe(ws)
            wb.close()
        
        output_path = generate_output_filename(file_path, "report")
        
        wb_out = Workbook()
        default_sheet = wb_out.active
        wb_out.remove(default_sheet)
        
        for sheet_name, df in dfs.items():
            if log_callback:
                log_callback(f"分析工作表: {sheet_name}")
            
            self._create_report_sheet(wb_out, sheet_name, df)
        
        wb_out.save(output_path)
        wb_out.close()
        
        if log_callback:
            log_callback(f"报表生成完成")
        
        return output_path
    
    def _create_report_sheet(self, wb: Workbook, sheet_name: str, df: pd.DataFrame):
        """创建报表工作表"""
        ws_summary = wb.create_sheet(title=f"{sheet_name}_Summary"[:31])
        
        row_num = 1
        
        ws_summary.cell(row=row_num, column=1, value=f"工作表: {sheet_name}")
        ws_summary.cell(row=row_num, column=1).font = Font(bold=True, size=14)
        row_num += 2
        
        ws_summary.cell(row=row_num, column=1, value="基本信息")
        ws_summary.cell(row=row_num, column=1).font = self._header_font
        row_num += 1
        
        basic_info = [
            ("总行数", len(df)),
            ("总列数", len(df.columns)),
            ("列名", ", ".join(map(str, df.columns.tolist()))),
            ("缺失值总数", int(df.isnull().sum().sum())),
        ]
        
        for name, value in basic_info:
            ws_summary.cell(row=row_num, column=1, value=name)
            ws_summary.cell(row=row_num, column=2, value=str(value))
            row_num += 1
        
        row_num += 1
        
        numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
        if numeric_cols:
            ws_summary.cell(row=row_num, column=1, value="数值列统计")
            ws_summary.cell(row=row_num, column=1).font = self._header_font
            row_num += 1
            
            stats_headers = ["列名", "非空值数", "平均值", "中位数", "标准差", "最小值", "最大值", "总和"]
            for col_idx, header in enumerate(stats_headers, 1):
                cell = ws_summary.cell(row=row_num, column=col_idx, value=header)
                cell.font = self._header_font_white
                cell.fill = self._header_fill
                cell.alignment = self._alignment
            row_num += 1
            
            for col in numeric_cols:
                col_data = df[col].dropna()
                if len(col_data) > 0:
                    stats = [
                        col,
                        len(col_data),
                        round(col_data.mean(), 4),
                        round(col_data.median(), 4),
                        round(col_data.std(), 4) if len(col_data) > 1 else 0,
                        round(col_data.min(), 4),
                        round(col_data.max(), 4),
                        round(col_data.sum(), 4),
                    ]
                    for col_idx, value in enumerate(stats, 1):
                        ws_summary.cell(row=row_num, column=col_idx, value=value)
                    row_num += 1
        
        for col in ws_summary.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 30)
            ws_summary.column_dimensions[column].width = adjusted_width
    
    def _sheet_to_dataframe(self, ws) -> pd.DataFrame:
        """将工作表转换为DataFrame"""
        data = []
        for row in ws.iter_rows(values_only=True):
            data.append(row)
        
        if not data:
            return pd.DataFrame()
        
        headers = data[0]
        rows = data[1:]
        
        return pd.DataFrame(rows, columns=headers)
    
    def _write_df_to_sheet(self, wb: Workbook, sheet_name: str, df: pd.DataFrame):
        """将DataFrame写入工作表"""
        ws = wb.create_sheet(title=sheet_name[:31])
        
        for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
            for c_idx, value in enumerate(row, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=value)
                if r_idx == 1:
                    cell.font = self._header_font_white
                    cell.fill = self._header_fill
                    cell.alignment = self._alignment
        
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            ws.column_dimensions[column].width = adjusted_width
