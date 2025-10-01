import os
import glob
from datetime import datetime, timedelta
import xlwings as xw

class ExcelProcessor:
    def __init__(self):
        self.work_dir = r"C:\Users\Monarch\Downloads"
        self.output_dir = r"C:\Users\Monarch\Desktop"
        self.app = None
        self.summary_wb = None
        self.total_rows = 0  # 添加总行数计数器
        self.processed_rows = 0  # 添加已处理行数计数器
        
        # 定义模式映射，用于快速识别文件格式
        self.pattern_definitions = {
            "pattern1": {"headers": {1: "销售日期", 2: "单位名称", 3: "商品名称", 4: "商品规格", 6: "销售数量", 8: "销售批号"},
                       "mapping": {1: 1, 3: 2, 4: 3, 8: 4, 2: 5, 6: 6}},
            "pattern2": {"headers": {1: "日期", 5: "销往单位", 9: "药品名称", 11: "规格", 12: "数量", 14: "批号"},
                       "mapping": {1: 1, 9: 2, 11: 3, 14: 4, 5: 5, 12: 6}},
            "pattern3": {"headers": {1: "销售日期", 4: "销售商", 6: "商品名称", 7: "商品规格", 10: "数量", 13: "批号"},
                       "mapping": {1: 1, 6: 2, 7: 3, 13: 4, 4: 5, 10: 6}},
            "pattern4": {"headers": {1: "销售日期", 4: "销售商", 6: "商品名称", 7: "商品规格", 10: "批号", 11: "数量"},
                       "mapping": {1: 1, 6: 2, 7: 3, 10: 4, 4: 5, 11: 6}},
            "pattern5": {"headers": {3: "销售时间", 5: "客户名称", 10: "通用名", 14: "规格", 16: "供应商批次", 18: "销售数量"},
                       "mapping": {3: 1, 10: 2, 14: 3, 16: 4, 5: 5, 18: 6}, "date_format": True, "date_col": 3},
            "pattern6": {"headers": {3: "发票日期", 5: "客户", 9: "商品名称", 10: "商品规格", 12: "开票数量", 16: "批号"},
                       "mapping": {3: 1, 9: 2, 10: 3, 16: 4, 5: 5, 12: 6}},
            "pattern7": {"headers": {3: "出库日期", 12: "客户名称", 6: "商品名称", 7: "品种规格", 9: "数量", 8: "批号"},
                       "mapping": {3: 1, 6: 2, 7: 3, 8: 4, 12: 5, 9: 6}},
            "pattern8": {"headers": {2: "制单时间", 5: "客户名称", 7: "品名", 8: "品规", 12: "订单数量", 15: "批号"},
                       "mapping": {2: 1, 7: 2, 8: 3, 15: 4, 5: 5, 12: 6}},
            "pattern9": {"headers": {7: "出库日期", 20: "下游收货方名称", 21: "产品名称", 22: "产品规格", 24: "数量", 25: "原始批号"},
                       "mapping": {7: 1, 21: 2, 22: 3, 25: 4, 20: 5, 24: 6}}
        }
        
    def initialize_excel(self):
        """初始化Excel应用"""
        try:
            self.app = xw.App(visible=False, add_book=False)  # 设置为不可见以提高性能
            return True
        except Exception as e:
            print(f"初始化Excel失败: {e}")
            return False
    
    def create_summary_file(self):
        """创建汇总文件"""
        try:
            # 获取昨天的日期
            yesterday = datetime.now() - timedelta(days=1)
            date_str = yesterday.strftime("%Y-%m-%d")
            
            # 创建文件名
            filename = f"湖北区域每日网上下载出库汇总{date_str}.xlsx"
            filepath = os.path.join(self.output_dir, filename)
            
            # 创建新工作簿
            self.summary_wb = self.app.books.add()
            ws = self.summary_wb.sheets[0]
            
            # 设置表头 - 使用一次性写入而不是循环
            headers = ["日期", "品种", "规格", "批号", "流向单位", "数量"]
            ws.range("A1").value = [headers]  # 一次性写入表头
            
            # 保存文件
            self.summary_wb.save(filepath)
            print(f"创建汇总文件: {filepath}")
            
            return filepath
            
        except Exception as e:
            print(f"创建汇总文件失败: {e}")
            return None
    
    def get_file_pattern(self, wb):
        """识别文件格式模式"""
        try:
            ws = wb.sheets[0]
            
            # 一次性读取第一行的所有值
            header_row = ws.range("A1:AA1").value
            
            # 创建headers字典
            headers = {}
            for col, value in enumerate(header_row, 1):
                if value:
                    headers[col] = str(value).strip()
            
            # 使用模式定义进行匹配
            for pattern_name, pattern_info in self.pattern_definitions.items():
                pattern_headers = pattern_info["headers"]
                if all(headers.get(col) == val for col, val in pattern_headers.items()):
                    return pattern_name
            
            return "unknown"
            
        except Exception as e:
            print(f"识别文件格式失败: {e}")
            return "unknown"
    
    def count_file_rows(self, wb):
        """统计文件中除第一行外的行数"""
        try:
            ws = wb.sheets[0]
            last_row = ws.used_range.last_cell.row
            if last_row <= 1:  # 只有表头，没有数据
                return 0
            return last_row - 1  # 减去表头行
        except Exception as e:
            print(f"统计文件行数时出错: {e}")
            return 0
    
    def process_file(self, wb, pattern):
        """根据识别的模式处理文件"""
        if pattern == "unknown" or pattern not in self.pattern_definitions:
            return None
        
        try:
            ws = wb.sheets[0]
            pattern_info = self.pattern_definitions[pattern]
            mapping = pattern_info["mapping"]
            
            # 获取数据范围
            last_row = ws.used_range.last_cell.row
            if last_row <= 1:  # 只有表头，没有数据
                return []
            
            # 一次性读取所有数据
            data_range = ws.range(f"A2:{chr(64+max(mapping.keys()))}" + str(last_row))
            all_data = data_range.value
            
            # 处理数据
            result = []
            
            # 特殊处理只有一行数据的情况
            if last_row == 2:
                # 如果只有一行数据，直接读取该行并进行列映射
                row_data = []
                for col in range(1, max(mapping.keys()) + 1):
                    cell_value = ws.range(f"{chr(64+col)}2").value
                    row_data.append(cell_value)
                
                # 创建新行数据
                new_row = [None] * 6
                for src_col, dst_col in mapping.items():
                    value = row_data[src_col-1] if src_col-1 < len(row_data) else None
                    
                    # 处理字符串中的空格
                    if isinstance(value, str):
                        # 对于pattern8，制单日期列（src_col=2）将空格及后面的内容替换为空字符串
                        if pattern == "pattern8" and src_col == 2:
                            if " " in value:
                                value = value.split(" ")[0]  # 只保留空格前的内容
                        else:
                            value = value.replace(" ", "")
                    
                    # 处理日期格式化
                    if pattern_info.get("date_format") and src_col == pattern_info.get("date_col") and isinstance(value, (int, float)):
                        value_str = str(int(value))
                        if len(value_str) == 8:
                            value = f"{value_str[:4]}/{value_str[4:6]}/{value_str[6:8]}"
                    
                    new_row[dst_col-1] = value
                
                result.append(new_row)
            else:
                # 处理多行数据的情况
                # 确保all_data是列表的列表
                if not isinstance(all_data, list):
                    all_data = [all_data]
                elif all_data and not isinstance(all_data[0], (list, tuple)):
                    all_data = [all_data]
                
                for row_data in all_data:
                    # 确保row_data是可迭代的
                    if not isinstance(row_data, (list, tuple)) or row_data is None:
                        continue
                        
                    # 跳过空行
                    if all(v is None or v == "" for v in row_data):
                        continue
                        
                    # 创建新行数据
                    new_row = [None] * 6
                    for src_col, dst_col in mapping.items():
                        value = row_data[src_col-1] if src_col-1 < len(row_data) else None
                        
                        # 处理字符串中的空格
                        if isinstance(value, str):
                            # 对于pattern8，制单日期列（src_col=2）将空格及后面的内容替换为空字符串
                            if pattern == "pattern8" and src_col == 2:
                                if " " in value:
                                    value = value.split(" ")[0]  # 只保留空格前的内容
                            else:
                                value = value.replace(" ", "")
                        
                        # 处理日期格式化
                        if dst_col == 1:  # 第一列是日期列
                            if isinstance(value, str):
                                # 尝试处理各种可能的日期格式
                                try:
                                    # 处理 YYYY/MM/DD 或 YYYY-MM-DD 格式
                                    if "/" in value or "-" in value:
                                        date_parts = value.replace("/", "-").split("-")
                                        if len(date_parts) == 3:
                                            value = f"{date_parts[0]}-{date_parts[1].zfill(2)}-{date_parts[2].zfill(2)}"
                                    # 处理 YYYYMMDD 格式
                                    elif len(value.strip()) == 8 and value.strip().isdigit():
                                        value = f"{value[:4]}-{value[4:6]}-{value[6:8]}"
                                except:
                                    pass  # 如果转换失败，保持原值
                            elif isinstance(value, (int, float)):
                                # 处理数字格式的日期
                                try:
                                    value_str = str(int(value))
                                    if len(value_str) == 8:
                                        value = f"{value_str[:4]}-{value_str[4:6]}-{value_str[6:8]}"
                                except:
                                    pass  # 如果转换失败，保持原值
                        
                        new_row[dst_col-1] = value
                    
                    result.append(new_row)
            
            # 更新已处理行数
            self.processed_rows += len(result)
            
            return result
            
        except Exception as e:
            print(f"处理文件数据时出错: {e}")
            import traceback
            print(traceback.format_exc())
            return None
    
    def append_data_to_summary(self, result):
        """将数据追加到汇总文件"""
        if not result:
            return
        
        try:
            ws = self.summary_wb.sheets[0]
            
            # 找到下一个空行
            last_row = ws.used_range.last_cell.row
            next_row = last_row + 1
            
            # 一次性写入所有数据
            if result:
                ws.range(f"A{next_row}").value = result
            
            # 保存文件
            self.summary_wb.save()
            print(f"已追加 {len(result)} 行数据到汇总文件")
            
        except Exception as e:
            print(f"追加数据失败: {e}")
    
    def process_excel_files(self):
        """处理所有Excel文件"""
        # 获取工作目录下的所有Excel文件
        excel_files = []
        for ext in ['*.xls', '*.xlsx']:
            excel_files.extend(glob.glob(os.path.join(self.work_dir, ext)))
        
        print(f"找到 {len(excel_files)} 个Excel文件")
        
        # 首先统计所有文件的总行数
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
        
        # 然后处理所有文件
        for file_path in excel_files:
            wb = None
            try:
                print(f"正在处理: {os.path.basename(file_path)}")
                
                # 打开文件
                wb = self.app.books.open(file_path)
                
                # 识别文件格式
                pattern = self.get_file_pattern(wb)
                print(f"识别为格式: {pattern}")
                
                # 处理数据
                if pattern != "unknown":
                    data = self.process_file(wb, pattern)
                    if data:
                        self.append_data_to_summary(data)
                else:
                    print(f"未识别的文件格式，跳过文件: {os.path.basename(file_path)}")
                
            except Exception as e:
                print(f"处理文件 {os.path.basename(file_path)} 时出错: {e}")
            finally:
                # 确保文件关闭
                if wb:
                    try:
                        wb.close()
                    except:
                        pass
    
    def run(self):
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
            self.process_excel_files()
            
            # 4. 打印汇总信息
            print(f"总计处理了 {self.processed_rows} 行数据，占总行数的 {self.processed_rows/self.total_rows*100:.2f}%")
            if self.processed_rows < self.total_rows:
                print(f"警告：有 {self.total_rows - self.processed_rows} 行数据未被处理，可能是因为格式不匹配或数据有问题")
            
            # 5. 打开输出目录
            os.startfile(self.output_dir)
            
            print("流向整理完成！")
            
        except Exception as e:
            print(f"程序运行出错: {e}")
        
        finally:
            # 确保Excel应用关闭
            if self.app:
                try:
                    self.app.quit()
                except:
                    pass

if __name__ == "__main__":
    processor = ExcelProcessor()
    processor.run()