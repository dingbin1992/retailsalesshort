import os
import glob
from datetime import datetime, timedelta
import xlwings as xw

class ExcelProcessor:
    def __init__(self):
        self.work_dir = r"C:\Users\DingBin\Downloads"
        self.output_dir = r"C:\Users\DingBin\Desktop"
        self.app = None
        self.summary_wb = None
        
    def initialize_excel(self):
        """l应用"""
        try:
            self.app = xw.App(visible=True, add_book=False)
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
            filename = f"湖北区域每日纯销统计{date_str}.xlsx"
            filepath = os.path.join(self.output_dir, filename)
            
            # 创建新工作簿
            self.summary_wb = self.app.books.add()
            ws = self.summary_wb.sheets[0]
            
            # 设置表头
            headers = ["日期", "品种", "规格", "批号", "单位名称", "数量"]
            for i, header in enumerate(headers, 1):
                ws.cells(1, i).value = header
            
            # 保存文件
            self.summary_wb.save(filepath)
            print(f"创建汇总文件: {filepath}")
            
            return filepath
            
        except Exception as e:
            print(f"创建汇总文件失败: {e}")
            return None
    
    def remove_spaces_from_cells(self, ws, max_row, max_col):
        """移除单元格中的所有空格"""
        for row in range(1, max_row + 1):
            for col in range(1, max_col + 1):
                cell_value = ws.cells(row, col).value
                if isinstance(cell_value, str):
                    ws.cells(row, col).value = cell_value.replace(" ", "")
    
    def format_date_column(self, ws, col_index, max_row):
        """格式化日期列，将20250631格式转换为2025/06/31"""
        for row in range(2, max_row + 1):
            cell_value = ws.cells(row, col_index).value
            if isinstance(cell_value, (int, float)) and len(str(int(cell_value))) == 8:
                date_str = str(int(cell_value))
                formatted_date = f"{date_str[:4]}/{date_str[4:6]}/{date_str[6:8]}"
                ws.cells(row, col_index).value = formatted_date
    
    def get_file_pattern(self, wb):
        """识别文件格式模式"""
        try:
            ws = wb.sheets[0]
            
            # 读取第一行的值来判断格式
            headers = {}
            for col in range(1, 21):  # 检查前20列
                cell_value = ws.cells(1, col).value
                if cell_value:
                    headers[col] = str(cell_value).strip()
            
            # 模式1: A1=销售日期, B1=单位名称, C1=商品名称, D1=商品规格, F1=销售数量, H1=销售批号
            if (headers.get(1) == "销售日期" and headers.get(2) == "单位名称" and headers.get(3) == "商品名称" and headers.get(4) == "商品规格" and  headers.get(6) == "销售数量" and headers.get(8) == "销售批号"):
                return "pattern1"
            
            # 模式2: A1=日期, E1=销往单位, I1=药品名称, K1=规格, L1=数量, N1=批号
            elif (headers.get(1) == "日期" and headers.get(5) == "销往单位" and headers.get(9) == "药品名称" and headers.get(11) == "规格" and headers.get(12) == "数量" and headers.get(14) == "批号"):
                return "pattern2"
            
            # 模式3: A1=销售日期, D1=销售商, F1=商品名称, G1=商品规格, J1=数量, M1=批号
            elif (headers.get(1) == "销售日期" and headers.get(4) == "销售商" and headers.get(6) == "商品名称" and headers.get(7) == "商品规格" and headers.get(10) == "数量" and headers.get(13) == "批号"):
                return "pattern3"
            
            # 模式4: A1=销售日期, D1=销售商, F1=商品名称, G1=商品规格, J1=批号, K1=数量
            elif (headers.get(1) == "销售日期" and headers.get(4) == "销售商" and headers.get(6) == "商品名称" and headers.get(7) == "商品规格" and headers.get(10) == "批号" and headers.get(11) == "数量"):
                return "pattern4"
            
            # 模式5: C1=销售时间, E1=客户名称, J1=通用名, N1=规格, P1=供应商批次, R1=销售数量
            elif (headers.get(3) == "销售时间" and headers.get(5) == "客户名称" and headers.get(10) == "通用名" and headers.get(14) == "规格" and headers.get(16) == "供应商批次" and headers.get(18) == "销售数量"):
                return "pattern5"
            
            # 模式6: C1=发票日期, E1=客户, I1=商品名称, J1=商品规格, L1=开票数量, P1=批号
            elif (headers.get(3) == "发票日期" and headers.get(5) == "客户" and headers.get(9) == "商品名称" and headers.get(10) == "商品规格" and headers.get(12) == "开票数量" and headers.get(16) == "批号"):
                return "pattern6"
            
            # 模式7: C1=出库日期, L1=客户名称, F1=商品名称, G1=品种规格, I1=数量, H1=批号
            elif (headers.get(3) == "出库日期" and headers.get(12) == "客户名称" and headers.get(6) == "商品名称" and headers.get(7) == "品种规格" and headers.get(9) == "数量" and headers.get(8) == "批号"):
                return "pattern7"
            
            # 模式8: B1=制单时间, E1=客户名称, G1=品名, H1=品规, L1=订单数量, O1=批号
            elif (headers.get(2) == "制单时间" and headers.get(5) == "客户名称" and headers.get(7) == "品名" and headers.get(8) == "品规" and headers.get(12) == "订单数量" and headers.get(15) == "批号"):
                return "pattern8"
            
            return "unknown"
            
        except Exception as e:
            print(f"识别文件格式失败: {e}")
            return "unknown"  

    def process_pattern1(self, wb):
        """处理模式1: A1=销售日期, B1=单位名称, C1=商品名称, D1=商品规格, F1=销售数量, H1=销售批号"""
        ws = wb.sheets[0]
        max_row = ws.used_range.last_cell.row
        
        # 1. 移除空格
        self.remove_spaces_from_cells(ws, max_row, 8)
        
        # 2. 重新排列列: 销售日期(A)->1, 商品名称(C)->2, 商品规格(D)->3, 销售批号(H)->4, 单位名称(B)->5, 销售数量(F)->6
        data = []
        for row in range(2, max_row + 1):
            row_data = [
                ws.cells(row, 1).value,  # 销售日期
                ws.cells(row, 3).value,  # 商品名称
                ws.cells(row, 4).value,  # 商品规格
                ws.cells(row, 8).value,  # 销售批号
                ws.cells(row, 2).value,  # 单位名称
                ws.cells(row, 6).value   # 销售数量
            ]
            data.append(row_data)
        
        return data
    
    def process_pattern2(self, wb):
        """处理模式2: A1=日期, E1=销往单位, I1=药品名称, K1=规格, L1=数量, N1=批号"""
        ws = wb.sheets[0]
        max_row = ws.used_range.last_cell.row
        
        # 1. 移除空格
        self.remove_spaces_from_cells(ws, max_row, 14)
        
        # 2. 重新排列列: 日期(A)->1, 药品名称(I)->2, 规格(K)->3, 批号(N)->4, 销往单位(E)->5, 数量(L)->6
        data = []
        for row in range(2, max_row + 1):
            row_data = [
                ws.cells(row, 1).value,  # 日期
                ws.cells(row, 9).value,  # 药品名称
                ws.cells(row, 11).value, # 规格
                ws.cells(row, 14).value, # 批号
                ws.cells(row, 5).value,  # 销往单位
                ws.cells(row, 12).value  # 数量
            ]
            data.append(row_data)
        
        return data
    
    def process_pattern3(self, wb):
        """处理模式3: A1=销售日期, D1=销售商, F1=商品名称, G1=商品规格, J1=数量, M1=批号"""
        ws = wb.sheets[0]
        max_row = ws.used_range.last_cell.row
        
        # 1. 移除空格
        self.remove_spaces_from_cells(ws, max_row, 13)
        
        # 2. 重新排列列: 销售日期(A)->1, 商品名称(F)->2, 商品规格(G)->3, 批号(M)->4, 销售商(D)->5, 数量(J)->6
        data = []
        for row in range(2, max_row + 1):
            row_data = [
                ws.cells(row, 1).value,  # 销售日期
                ws.cells(row, 6).value,  # 商品名称
                ws.cells(row, 7).value,  # 商品规格
                ws.cells(row, 13).value, # 批号
                ws.cells(row, 4).value,  # 销售商
                ws.cells(row, 10).value  # 数量
            ]
            data.append(row_data)
        
        return data
    
    def process_pattern4(self, wb):
        """处理模式4: A1=销售日期, D1=销售商, F1=商品名称, G1=商品规格, J1=批号, K1=数量"""
        ws = wb.sheets[0]
        max_row = ws.used_range.last_cell.row
        
        # 1. 移除空格
        self.remove_spaces_from_cells(ws, max_row, 11)
        
        # 2. 重新排列列: 销售日期(A)->1, 商品名称(F)->2, 商品规格(G)->3, 批号(J)->4, 销售商(D)->5, 数量(K)->6
        data = []
        for row in range(2, max_row + 1):
            row_data = [
                ws.cells(row, 1).value,  # 销售日期
                ws.cells(row, 6).value,  # 商品名称
                ws.cells(row, 7).value,  # 商品规格
                ws.cells(row, 10).value, # 批号
                ws.cells(row, 4).value,  # 销售商
                ws.cells(row, 11).value  # 数量
            ]
            data.append(row_data)
        
        return data
    
    def process_pattern5(self, wb):
        """处理模式5: C1=销售时间, E1=客户名称, J1=通用名, N1=规格, P1=供应商批次, R1=销售数量"""
        ws = wb.sheets[0]
        max_row = ws.used_range.last_cell.row
        
        # 1. 格式化日期列(C列)
        self.format_date_column(ws, 3, max_row)
        
        # 2. 重新排列列: 销售时间(C)->1, 通用名(J)->2, 规格(N)->3, 供应商批次(P)->4, 客户名称(E)->5, 销售数量(R)->6
        data = []
        for row in range(2, max_row + 1):
            row_data = [
                ws.cells(row, 3).value,  # 销售时间
                ws.cells(row, 10).value, # 通用名
                ws.cells(row, 14).value, # 规格
                ws.cells(row, 16).value, # 供应商批次
                ws.cells(row, 5).value,  # 客户名称
                ws.cells(row, 18).value  # 销售数量
            ]
            data.append(row_data)
        
        return data    

    def process_pattern6(self, wb):
        """处理模式6: C1=发票日期, E1=客户, I1=商品名称, J1=商品规格, L1=开票数量, P1=批号"""
        ws = wb.sheets[0]
        max_row = ws.used_range.last_cell.row
        
        # 1. 移除空格
        self.remove_spaces_from_cells(ws, max_row, 16)
        
        # 2. 重新排列列: 发票日期(C)->1, 商品名称(I)->2, 商品规格(J)->3, 批号(P)->4, 客户(E)->5, 开票数量(L)->6
        data = []
        for row in range(2, max_row + 1):
            row_data = [
                ws.cells(row, 3).value,  # 发票日期
                ws.cells(row, 9).value,  # 商品名称
                ws.cells(row, 10).value, # 商品规格
                ws.cells(row, 16).value, # 批号
                ws.cells(row, 5).value,  # 客户
                ws.cells(row, 12).value  # 开票数量
            ]
            data.append(row_data)
        
        return data
    
    def process_pattern7(self, wb):
        """处理模式7: C1=出库日期, L1=客户名称, F1=商品名称, G1=品种规格, I1=数量, H1=批号"""
        ws = wb.sheets[0]
        max_row = ws.used_range.last_cell.row
        
        # 1. 移除空格
        self.remove_spaces_from_cells(ws, max_row, 12)
        
        # 2. 重新排列列: 出库日期(C)->1, 商品名称(F)->2, 品种规格(G)->3, 批号(H)->4, 客户名称(L)->5, 数量(I)->6
        data = []
        for row in range(2, max_row + 1):
            row_data = [
                ws.cells(row, 3).value,  # 出库日期
                ws.cells(row, 6).value,  # 商品名称
                ws.cells(row, 7).value,  # 品种规格
                ws.cells(row, 8).value,  # 批号
                ws.cells(row, 12).value, # 客户名称
                ws.cells(row, 9).value   # 数量
            ]
            data.append(row_data)
        
        return data
    
    def process_pattern8(self, wb):
        """处理模式8: B1=制单时间, E1=客户名称, G1=品名, H1=品规, L1=订单数量, O1=批号"""
        ws = wb.sheets[0]
        max_row = ws.used_range.last_cell.row
        
        # 1. 移除空格
        self.remove_spaces_from_cells(ws, max_row, 15)
        
        # 2. 重新排列列: 制单时间(B)->1, 品名(G)->2, 品规(H)->3, 批号(O)->4, 客户名称(E)->5, 订单数量(L)->6
        data = []
        for row in range(2, max_row + 1):
            row_data = [
                ws.cells(row, 2).value,  # 制单时间
                ws.cells(row, 7).value,  # 品名
                ws.cells(row, 8).value,  # 品规
                ws.cells(row, 15).value, # 批号
                ws.cells(row, 5).value,  # 客户名称
                ws.cells(row, 12).value  # 订单数量
            ]
            data.append(row_data)
        
        return data
    
    def append_data_to_summary(self, data):
        """将数据追加到汇总文件"""
        try:
            if not data:
                return
            
            ws = self.summary_wb.sheets[0]
            
            # 找到下一个空行
            last_row = ws.used_range.last_cell.row if ws.used_range else 1
            next_row = last_row + 1
            
            # 追加数据
            for i, row_data in enumerate(data):
                for j, value in enumerate(row_data):
                    ws.cells(next_row + i, j + 1).value = value
            
            # 保存文件
            self.summary_wb.save()
            print(f"已追加 {len(data)} 行数据到汇总文件")
            
        except Exception as e:
            print(f"追加数据失败: {e}") 
    def process_excel_files(self):
        """处理所有Excel文件"""
        try:
            # 获取工作目录下的所有Excel文件
            excel_files = []
            for ext in ['*.xls', '*.xlsx']:
                excel_files.extend(glob.glob(os.path.join(self.work_dir, ext)))
            
            print(f"找到 {len(excel_files)} 个Excel文件")
            
            for file_path in excel_files:
                try:
                    print(f"正在处理: {os.path.basename(file_path)}")
                    
                    # 打开文件
                    wb = self.app.books.open(file_path)
                    
                    # 识别文件格式
                    pattern = self.get_file_pattern(wb)
                    print(f"识别为格式: {pattern}")
                    
                    # 根据格式处理数据
                    data = None
                    if pattern == "pattern1":
                        data = self.process_pattern1(wb)
                    elif pattern == "pattern2":
                        data = self.process_pattern2(wb)
                    elif pattern == "pattern3":
                        data = self.process_pattern3(wb)
                    elif pattern == "pattern4":
                        data = self.process_pattern4(wb)
                    elif pattern == "pattern5":
                        data = self.process_pattern5(wb)
                    elif pattern == "pattern6":
                        data = self.process_pattern6(wb)
                    elif pattern == "pattern7":
                        data = self.process_pattern7(wb)
                    elif pattern == "pattern8":
                        data = self.process_pattern8(wb)
                    else:
                        print(f"未识别的文件格式，跳过文件: {os.path.basename(file_path)}")
                    
                    # 追加数据到汇总文件
                    if data:
                        self.append_data_to_summary(data)
                    
                    # 关闭文件
                    wb.close()
                    
                except Exception as e:
                    print(f"处理文件 {os.path.basename(file_path)} 时出错: {e}")
                    try:
                        wb.close()
                    except:
                        pass
                    continue
            
        except Exception as e:
            print(f"处理Excel文件时出错: {e}")
    
    def cleanup_and_open_output_dir(self):
        """清理并打开输出目录"""
        try:
            # 关闭所有工作簿
            for wb in self.app.books:
                try:
                    wb.close()
                except:
                    pass
            
            # 打开输出目录
            os.startfile(self.output_dir)
            
        except Exception as e:
            print(f"清理和打开目录时出错: {e}")
    
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
            
            # 4. 清理并打开输出目录
            self.cleanup_and_open_output_dir()
            
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