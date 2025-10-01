import os
import glob
from datetime import datetime, timedelta
import xlwings as xw
import threading
import sys
from typing import List, Dict, Optional, Tuple, Any

class ExcelProcessor:
    def __init__(self):
        self.work_dir = os.path.join(os.path.expanduser("~"), "Downloads")
        self.output_dir = r"H:\0、工作\0、每日纯销统计\实时下载流向数据"
        
        self.app = None
        self.summary_wb = None
        self.total_rows = 0
        self.processed_rows = 0
        self._lock = threading.Lock()
        
        # 模式定义
        self.pattern_definitions = {
            "pattern1": {
                "headers": {1: "销售日期", 2: "单位名称", 3: "商品名称", 4: "商品规格", 6: "销售数量", 8: "销售批号"},
                "mapping": {1: 1, 3: 2, 4: 3, 8: 4, 2: 5, 6: 6}
            },
            "pattern2": {
                "headers": {1: "日期", 5: "销往单位", 9: "药品名称", 11: "规格", 12: "数量", 14: "批号"},
                "mapping": {1: 1, 9: 2, 11: 3, 14: 4, 5: 5, 12: 6}
            },
            "pattern3": {
                "headers": {1: "销售日期", 4: "销售商", 6: "商品名称", 7: "商品规格", 10: "数量", 13: "批号"},
                "mapping": {1: 1, 6: 2, 7: 3, 13: 4, 4: 5, 10: 6}
            },
            "pattern4": {
                "headers": {1: "销售日期", 4: "销售商", 6: "商品名称", 7: "商品规格", 10: "批号", 11: "数量"},
                "mapping": {1: 1, 6: 2, 7: 3, 10: 4, 4: 5, 11: 6}
            },
            "pattern5": {
                "headers": {3: "销售时间", 5: "客户名称", 10: "通用名", 14: "规格", 16: "供应商批次", 18: "销售数量"},
                "mapping": {3: 1, 10: 2, 14: 3, 16: 4, 5: 5, 18: 6},
                "date_format": True, "date_col": 3
            },
            "pattern6": {
                "headers": {3: "发票日期", 5: "客户", 9: "商品名称", 10: "商品规格", 12: "开票数量", 16: "批号"},
                "mapping": {3: 1, 9: 2, 10: 3, 16: 4, 5: 5, 12: 6}
            },
            "pattern7": {
                "headers": {3: "出库日期", 12: "客户名称", 6: "商品名称", 7: "品种规格", 9: "数量", 8: "批号"},
                "mapping": {3: 1, 6: 2, 7: 3, 8: 4, 12: 5, 9: 6}
            },
            "pattern8": {
                "headers": {2: "制单时间", 5: "客户名称", 7: "品名", 8: "品规", 12: "订单数量", 15: "批号"},
                "mapping": {2: 1, 7: 2, 8: 3, 15: 4, 5: 5, 12: 6}
            },
            "pattern9": {
                "headers": {7: "出库日期", 20: "下游收货方名称", 21: "产品名称", 22: "产品规格", 24: "数量", 25: "原始批号"},
                "mapping": {7: 1, 21: 2, 22: 3, 25: 4, 20: 5, 24: 6}
            }
        }
        
    def initialize_excel(self) -> bool:
        """初始化Excel应用"""
        try:
            # 使用更兼容的方式初始化Excel
            self.app = xw.App(visible=False, add_book=False)
            # 设置Excel应用程序属性，提高稳定性
            if self.app:
                self.app.display_alerts = False
                self.app.screen_updating = False
            return True
        except Exception as e:
            print(f"初始化Excel失败: {e}")
            # 尝试使用备用方法初始化
            try:
                print("尝试使用备用方法初始化Excel...")
                # 使用win32com直接初始化
                import win32com.client
                self.app = win32com.client.DispatchEx("Excel.Application")
                self.app.Visible = False
                self.app.DisplayAlerts = False
                # 将win32com对象包装为xlwings对象
                self.app = xw.App(impl=self.app)
                return True
            except Exception as e2:
                print(f"备用初始化方法也失败: {e2}")
                return False
    
    def create_summary_file(self) -> Optional[str]:
        """创建汇总文件"""
        try:
            yesterday = datetime.now() - timedelta(days=1)
            date_str = yesterday.strftime("%Y-%m-%d")
            filename = f"湖北区域每日网上下载出库汇总{date_str}.xlsx"
            filepath = os.path.join(self.output_dir, filename)
            
            self.summary_wb = self.app.books.add()
            ws = self.summary_wb.sheets[0]
            
            # 设置表头
            headers = ["日期", "品种", "规格", "批号", "流向单位", "数量"]
            ws.range("A1:F1").value = headers
            
            self.summary_wb.save(filepath)
            print(f"创建汇总文件: {filepath}")
            return filepath
            
        except Exception as e:
            print(f"创建汇总文件失败: {e}")
            return None
    
    def get_file_pattern(self, wb) -> str:
        """识别文件格式模式"""
        try:
            ws = wb.sheets[0]
            
            # 读取表头行，限制范围避免过大
            header_row = ws.range("A1:Z1").value
            if not header_row:
                return "unknown"
            
            # 创建headers字典
            headers = {}
            for col, value in enumerate(header_row, 1):
                if value:
                    headers[col] = str(value).strip()
            
            # 匹配模式
            for pattern_name, pattern_info in self.pattern_definitions.items():
                pattern_headers = pattern_info["headers"]
                if all(headers.get(col) == val for col, val in pattern_headers.items()):
                    return pattern_name
            
            return "unknown"
            
        except Exception as e:
            print(f"识别文件格式失败: {e}")
            return "unknown"
    
    def count_file_rows(self, wb) -> int:
        """统计文件中除第一行外的行数"""
        try:
            ws = wb.sheets[0]
            last_row = ws.used_range.last_cell.row
            return max(0, last_row - 1)
        except Exception as e:
            print(f"统计文件行数时出错: {e}")
            return 0
    
    def normalize_date(self, value: Any, pattern: str, src_col: int) -> str:
        """标准化日期格式"""
        if not value:
            return value
            
        # 处理字符串中的空格
        if isinstance(value, str):
            if pattern == "pattern8" and src_col == 2:
                if " " in value:
                    value = value.split(" ")[0]
            else:
                value = value.replace(" ", "")
        
        # 日期格式化
        if isinstance(value, str):
            try:
                if "/" in value or "-" in value:
                    date_parts = value.replace("/", "-").split("-")
                    if len(date_parts) == 3:
                        return f"{date_parts[0]}-{date_parts[1].zfill(2)}-{date_parts[2].zfill(2)}"
                elif len(value.strip()) == 8 and value.strip().isdigit():
                    return f"{value[:4]}-{value[4:6]}-{value[6:8]}"
            except:
                pass
        elif isinstance(value, (int, float)):
            try:
                value_str = str(int(value))
                if len(value_str) == 8:
                    return f"{value_str[:4]}-{value_str[4:6]}-{value_str[6:8]}"
            except:
                pass
        
        return value
    
    def process_file(self, wb, pattern: str) -> Optional[List[List[Any]]]:
        """根据识别的模式处理文件"""
        if pattern == "unknown" or pattern not in self.pattern_definitions:
            return None
        
        try:
            ws = wb.sheets[0]
            pattern_info = self.pattern_definitions[pattern]
            mapping = pattern_info["mapping"]
            
            # 获取数据范围
            last_row = ws.used_range.last_cell.row
            if last_row <= 1:
                return []
            
            # 分批处理大文件，避免内存问题
            batch_size = 1000
            result = []
            
            for start_row in range(2, last_row + 1, batch_size):
                end_row = min(start_row + batch_size - 1, last_row)
                
                # 只读取需要的列
                max_col = max(mapping.keys())
                data_range = f"A{start_row}:{chr(64+max_col)}{end_row}"
                
                try:
                    batch_data = ws.range(data_range).value
                    
                    # 处理单行数据的情况
                    if not isinstance(batch_data, list):
                        batch_data = [batch_data] if batch_data is not None else []
                    elif batch_data and not isinstance(batch_data[0], (list, tuple)):
                        batch_data = [batch_data]
                    
                    # 处理批次数据
                    for row_data in batch_data:
                        if not isinstance(row_data, (list, tuple)) or row_data is None:
                            continue
                            
                        # 跳过空行
                        if all(v is None or v == "" for v in row_data):
                            continue
                        
                        # 创建新行数据
                        new_row = [None] * 6
                        for src_col, dst_col in mapping.items():
                            value = row_data[src_col-1] if src_col-1 < len(row_data) else None
                            
                            # 日期列特殊处理
                            if dst_col == 1:
                                value = self.normalize_date(value, pattern, src_col)
                            elif isinstance(value, str):
                                # 其他列的字符串处理
                                if pattern == "pattern8" and src_col == 2:
                                    if " " in value:
                                        value = value.split(" ")[0]
                                else:
                                    value = value.replace(" ", "")
                            
                            new_row[dst_col-1] = value
                        
                        result.append(new_row)
                        
                except Exception as e:
                    print(f"处理批次数据时出错 (行 {start_row}-{end_row}): {e}")
                    continue
            
            self.processed_rows += len(result)
            return result
            
        except Exception as e:
            print(f"处理文件数据时出错: {e}")
            import traceback
            traceback.print_exc()
            return None
    
    def append_data_to_summary(self, result: List[List[Any]]) -> None:
        """将数据追加到汇总文件"""
        if not result:
            return
        
        try:
            ws = self.summary_wb.sheets[0]
            
            # 找到下一个空行
            last_row = ws.used_range.last_cell.row
            next_row = last_row + 1
            
            # 分批写入大量数据，避免Excel崩溃
            batch_size = 500
            for i in range(0, len(result), batch_size):
                batch = result[i:i+batch_size]
                current_row = next_row + i
                
                try:
                    # 写入批次数据
                    end_row = current_row + len(batch) - 1
                    ws.range(f"A{current_row}:F{end_row}").value = batch
                except Exception as e:
                    print(f"写入批次数据失败: {e}")
                    # 如果批量写入失败，尝试逐行写入
                    for j, row in enumerate(batch):
                        try:
                            ws.range(f"A{current_row + j}:F{current_row + j}").value = [row]
                        except:
                            continue
            
            # 保存文件
            self.summary_wb.save()
            print(f"已追加 {len(result)} 行数据到汇总文件")
            
        except Exception as e:
            print(f"追加数据失败: {e}")
    
    def process_excel_files(self) -> None:
        """处理所有Excel文件 - 串行版本避免并发问题"""
        # 获取所有Excel文件
        excel_files = []
        for ext in ['*.xls', '*.xlsx']:
            excel_files.extend(glob.glob(os.path.join(self.work_dir, ext)))
        
        print(f"找到 {len(excel_files)} 个Excel文件")
        
        if not excel_files:
            print("没有找到Excel文件")
            return
        
        # 统计总行数
        print("正在统计文件行数...")
        for file_path in excel_files:
            wb = None
            try:
                wb = self.app.books.open(file_path)
                rows = self.count_file_rows(wb)
                self.total_rows += rows
                print(f"文件 {os.path.basename(file_path)} 包含 {rows} 行数据")
            except Exception as e:
                print(f"统计文件 {os.path.basename(file_path)} 行数时出错: {e}")
            finally:
                if wb:
                    try:
                        wb.close()
                    except:
                        pass
        
        print(f"所有文件共包含 {self.total_rows} 行数据")
        
        # 串行处理文件
        processed_files = 0
        for file_path in excel_files:
            wb = None
            try:
                print(f"正在处理: {os.path.basename(file_path)} ({processed_files + 1}/{len(excel_files)})")
                
                wb = self.app.books.open(file_path)
                
                # 识别文件格式
                pattern = self.get_file_pattern(wb)
                print(f"识别为格式: {pattern}")
                
                if pattern != "unknown":
                    data = self.process_file(wb, pattern)
                    if data:
                        self.append_data_to_summary(data)
                        print(f"成功处理 {len(data)} 行数据")
                else:
                    print(f"未识别的文件格式，跳过文件: {os.path.basename(file_path)}")
                
                processed_files += 1
                
            except Exception as e:
                print(f"处理文件 {os.path.basename(file_path)} 时出错: {e}")
                import traceback
                traceback.print_exc()
            finally:
                if wb:
                    try:
                        wb.close()
                    except:
                        pass
    
    def run(self) -> None:
        """运行主程序"""
        print("开始流向整理程序...")
        
        # 1. 初始化Excel
        if not self.initialize_excel():
            return
        
        try:
            # 2. 创建汇总文件
            summary_file = self.create_summary_file()
            if not summary_file:
                return
            
            # 3. 处理所有Excel文件
            start_time = datetime.now()
            self.process_excel_files()
            end_time = datetime.now()
            
            # 4. 打印汇总信息
            processing_time = (end_time - start_time).total_seconds()
            print(f"处理耗时: {processing_time:.2f} 秒")
            print(f"总计处理了 {self.processed_rows} 行数据")
            
            if self.total_rows > 0:
                print(f"处理率: {self.processed_rows/self.total_rows*100:.2f}%")
                if self.processed_rows < self.total_rows:
                    print(f"警告：有 {self.total_rows - self.processed_rows} 行数据未被处理")
            
            # 5. 打开输出目录
            try:
                os.startfile(self.output_dir)
            except:
                print(f"无法打开输出目录: {self.output_dir}")
            
            print("流向整理完成！")
            
        except Exception as e:
            print(f"程序运行出错: {e}")
            import traceback
            traceback.print_exc()
        
        finally:
            # 确保Excel应用关闭
            if self.app:
                try:
                    if self.summary_wb:
                        self.summary_wb.close()
                    self.app.quit()
                except:
                    pass

if __name__ == "__main__":
    try:
        # 检查Excel是否已安装
        try:
            import win32com.client
            try:
                excel_check = win32com.client.GetActiveObject("Excel.Application")
                del excel_check  # 释放对象
                print("检测到Excel已运行")
            except:
                # 尝试创建新实例检查是否可用
                excel_check = win32com.client.Dispatch("Excel.Application")
                excel_check.Quit()
                del excel_check
                print("Excel已安装但未运行")
        except Exception as e:
            print(f"警告: 无法检测Excel: {e}")
            print("请确保Excel已正确安装")
            input("按回车键继续...")
        
        processor = ExcelProcessor()
        processor.run()
    except Exception as e:
        print(f"程序启动失败: {e}")
        import traceback
        traceback.print_exc()
        input("按回车键退出...")
