import sys
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QTableWidget, QTableWidgetItem, QPushButton,
                            QLabel, QLineEdit, QFileDialog, QGroupBox, QMessageBox)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
import os
from os.path import abspath
from docx import Document
from PyPDF2 import PdfReader
from pptx import Presentation
import platform
import fitz
from pathlib import Path
from PyQt6.QtGui import QFont, QIcon

# Windows 系统才导入 win32com
if platform.system() == 'Windows':
    from win32com import client

class CounterThread(QThread):
    progress = pyqtSignal(int, str, int, int, str)  # 行号, 文件路径, 单面页数, 双面页数, 状态
    finished = pyqtSignal(dict)  # 统计结果信号
    
    def __init__(self, files, counter):
        super().__init__()
        self.files = files
        self.counter = counter
        
    def run(self):
        stats = {
            '.pdf': {'pages': 0, 'files': 0},
            '.docx': {'pages': 0, 'files': 0},
            '.pptx': {'pages': 0, 'files': 0},
            '.xlsx': {'pages': 0, 'files': 0},
            '.xls': {'pages': 0, 'files': 0}
        }
        
        for i, file in enumerate(self.files):
            try:
                self.progress.emit(i, file, 0, 0, "处理中...")
                pages = self.counter.get_page_count(file)
                ext = os.path.splitext(file)[1].lower()
                
                if ext in stats:
                    stats[ext]['pages'] += pages
                    stats[ext]['files'] += 1
                
                self.progress.emit(i, file, pages, (pages + 1) // 2, "已完成")
                
            except Exception as e:
                self.progress.emit(i, file, 0, 0, f"错误: {str(e)}")
                
        self.finished.emit(stats)

class DocCounter(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("打印店文档页数统计 V1.0 by LittleBen")
        self.setGeometry(100, 100, 1000, 600)
        
        # 设置全局字体大小
        font = self.font()
        font.setPointSize(font.pointSize() + 2)
        self.setFont(font)
        
        # 创建主widget和布局
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QHBoxLayout()
        main_widget.setLayout(layout)
        
        # 左侧部分
        left_widget = QWidget()
        left_layout = QVBoxLayout()
        
        # 工具栏按钮
        toolbar = QHBoxLayout()
        self.add_folder_btn = QPushButton("导入文件夹")
        self.add_file_btn = QPushButton("导入文件")
        self.start_btn = QPushButton("开始统计")
        self.single_count_btn = QPushButton("单面计算器")
        self.double_count_btn = QPushButton("双面计算器")
        self.export_btn = QPushButton("导出excel")
        self.exit_btn = QPushButton("清空")
        
        toolbar.addWidget(self.add_folder_btn)
        toolbar.addWidget(self.add_file_btn)
        toolbar.addWidget(self.start_btn)
        toolbar.addWidget(self.single_count_btn)
        toolbar.addWidget(self.double_count_btn)
        toolbar.addWidget(self.export_btn)
        toolbar.addWidget(self.exit_btn)
        
        # 表格
        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels([ "文件路径", "类型", "单面", "双面", "状态"])
        # 设置表格自动调整大小
        self.table.horizontalHeader().setStretchLastSection(False)
        self.table.resizeEvent = lambda event: self.update_table()
        
        left_layout.addLayout(toolbar)
        left_layout.addWidget(self.table)
        left_widget.setLayout(left_layout)
        
        # 右侧统计区
        right_widget = QWidget()
        right_layout = QVBoxLayout()
        right_widget.setLayout(right_layout)
        right_widget.setMaximumWidth(300)

        # 总页数显示
        total_pages_group = QGroupBox("统计信息")
        total_pages_layout = QHBoxLayout()
        total_pages_label = QLabel("总页数：")
        self.total_pages = QLabel("")
        total_pages_layout.addWidget(total_pages_label)
        total_pages_layout.addWidget(self.total_pages)
        total_pages_layout.addStretch()
        total_pages_group.setLayout(total_pages_layout)
        right_layout.addWidget(total_pages_group)

        # 黑白打印区域
        bw_group = QGroupBox("黑白打印")
        bw_layout = QVBoxLayout()
        self.price_inputs = {}

        # 黑白单双面
        for label_text, name in [("单面", "bw_single"), ("双面", "bw_double")]:
            group = QHBoxLayout()
            quantity_input = QLineEdit()
            quantity_input.setPlaceholderText("数量")
            quantity_input.setMaximumWidth(60)
            price_input = QLineEdit()
            price_input.setPlaceholderText("单价")
            price_input.setMaximumWidth(60)
            amount_input = QLineEdit()
            amount_input.setReadOnly(True)
            amount_input.setMaximumWidth(60)
            amount_input.setText("0")
            
            group.addWidget(QLabel(f"{label_text}"))
            group.addWidget(quantity_input)
            group.addWidget(QLabel("X"))
            group.addWidget(price_input)
            group.addWidget(amount_input)
            group.addStretch()
            
            self.price_inputs[name] = {
                'quantity': quantity_input,
                'price': price_input,
                'amount': amount_input
            }
            bw_layout.addLayout(group)

            # 添加输入变化事件处理
            def create_change_handler(q_input, p_input, a_input):
                def on_change():
                    try:
                        quantity = float(q_input.text() or 0)
                        price = float(p_input.text() or 0)
                        amount = quantity * price
                        a_input.setText(f"{amount:.2f}")
                        # 更新总金额
                        self.calculate_amount()
                    except ValueError:
                        a_input.setText("0")
                return on_change
            
            change_handler = create_change_handler(quantity_input, price_input, amount_input)
            quantity_input.textChanged.connect(change_handler)
            price_input.textChanged.connect(change_handler)

        bw_group.setLayout(bw_layout)
        right_layout.addWidget(bw_group)

        # 彩色打印区域
        color_group = QGroupBox("彩色打印")
        color_layout = QVBoxLayout()

        # 彩色单双面
        for label_text, name in [("单面", "color_single"), ("双面", "color_double")]:
            group = QHBoxLayout()
            quantity_input = QLineEdit()
            quantity_input.setPlaceholderText("数量")
            quantity_input.setMaximumWidth(60)
            price_input = QLineEdit()
            price_input.setPlaceholderText("单价")
            price_input.setMaximumWidth(60)
            amount_input = QLineEdit()
            amount_input.setReadOnly(True)
            amount_input.setMaximumWidth(60)
            amount_input.setText("0")
            
            group.addWidget(QLabel(f"{label_text}"))
            group.addWidget(quantity_input)
            group.addWidget(QLabel("X"))
            group.addWidget(price_input)
            group.addWidget(amount_input)
            group.addStretch()
            
            self.price_inputs[name] = {
                'quantity': quantity_input,
                'price': price_input,
                'amount': amount_input
            }
            color_layout.addLayout(group)

            # 添加输入变化事件处理
            def create_change_handler(q_input, p_input, a_input):
                def on_change():
                    try:
                        quantity = float(q_input.text() or 0)
                        price = float(p_input.text() or 0)
                        amount = quantity * price
                        a_input.setText(f"{amount:.2f}")
                        # 更新总金额
                        self.calculate_amount()
                    except ValueError:
                        a_input.setText("0")
                return on_change
            
            change_handler = create_change_handler(quantity_input, price_input, amount_input)
            quantity_input.textChanged.connect(change_handler)
            price_input.textChanged.connect(change_handler)

        color_group.setLayout(color_layout)
        right_layout.addWidget(color_group)

        # 总金额显示
        total_amount_group = QGroupBox("金额统计")
        total_amount_layout = QHBoxLayout()
        total_amount_label = QLabel("总金额：")
        self.total_amount = QLabel("")
        total_amount_layout.addWidget(total_amount_label)
        total_amount_layout.addWidget(self.total_amount)
        total_amount_layout.addStretch()
        total_amount_group.setLayout(total_amount_layout)

        # 计算按钮
        calc_btn = QPushButton("计算")
        calc_btn.clicked.connect(self.calculate_amount)

        # 添加总金额和计算按钮
        right_layout.addWidget(total_amount_group)
        right_layout.addWidget(calc_btn)
        right_layout.addStretch()
        
        # 添加左右部分到主布局
        layout.addWidget(left_widget, stretch=7)
        layout.addWidget(right_widget, stretch=3)
        
        # 连接信号
        self.add_folder_btn.clicked.connect(self.add_folder)
        self.add_file_btn.clicked.connect(self.add_files)
        self.start_btn.clicked.connect(self.start_counting)
        self.exit_btn.clicked.connect(self.clear_table)
        
        self.files = []
        self.counter_thread = None  # 添加线程属性
        
        # 启用拖拽
        self.setAcceptDrops(True)
        self.table.setAcceptDrops(True)

        self.setWindowIcon(QIcon('app.ico'))

    def _count_word_pages(self, file_path):
        """Windows 系统下获取 Word 文档页数"""
        word = None
        try:
            word = client.Dispatch("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(abspath(file_path))
            word.ActiveDocument.Repaginate()
            pages = word.ActiveDocument.ComputeStatistics(2)  # 2 表示统计页数
            doc.Close()
            return pages
        except Exception as e:
            if "无法打开" in str(e):
                raise Exception("文件可能被占用或损坏")
            elif "RPC 服务器不可用" in str(e):
                raise Exception("Word服务未响应，请重试")
            elif "ActiveX 组件不能创建" in str(e):
                raise Exception("未安装Word或Word组件异常")
            else:
                raise Exception(f"Word处理失败: {str(e)}")
        finally:
            if word:
                try:
                    word.Application.Quit()
                except:
                    pass  # 忽略关闭Word时的错误

    def _count_ppt_pages(self, file_path):
        try:
            presentation = Presentation(file_path)
            return len(presentation.slides)
        except Exception as e:
            print(f"Error counting PPT pages: {str(e)}")
            return 0    

    def _count_excel_pages(self, excel_path):
        # 创建 Excel 应用实例
        excel = client.Dispatch("Excel.Application")
        excel.Visible = False  # 设置为不可见模式（后台运行）
        wb = None  # 初始化 wb 变量
        pdf_path = os.path.join(os.path.dirname(excel_path), "temp_output.pdf")  # 定义临时 PDF 路径

        try:
            wb = excel.Workbooks.Open(excel_path)
            
            # 获取所有工作表
            sheets = wb.Sheets
            
            # 先消所有选择
            excel.DisplayAlerts = False
            
            # 选择所有工作表
            sheets.Select()  # 直接选择所有工作表
            
            # 导出为PDF
            wb.ActiveSheet.ExportAsFixedFormat(
                Type=0,  # PDF格式
                Filename=pdf_path,
                Quality=0,  # 标准质量
                IncludeDocProperties=True,
                IgnorePrintAreas=False,
                OpenAfterPublish=False
            )
            
            # 统计PDF页数
            with open(pdf_path, 'rb') as pdf_file:
                pdf_reader = PdfReader(pdf_file)
                page_count = len(pdf_reader.pages)
                # print(f'PDF页数: {page_count}')
            
            # 关闭PDF文件后删除
            os.remove(pdf_path)
            # //print('PDF文件已删除')
            
            return page_count

        except Exception as e:
            print(f'失败: {str(e)}')
            return 0

        finally:
            if wb:
                wb.Close(SaveChanges=False)
            excel.Quit()

    def get_page_count(self, file_path):
        ext = Path(file_path).suffix.lower()
        
        try:
            if ext == '.pdf':
                doc = fitz.open(file_path)
                return len(doc)
            elif ext in ['.doc', '.docx', '.wps']:
                return self._count_word_pages(file_path)
            elif ext in ['.xls', '.xlsx', '.et']:
                return self._count_excel_pages(file_path)
            elif ext in ['.ppt', '.pptx', '.dps']:
                return self._count_ppt_pages(file_path)
        finally:
            if 'doc' in locals():
                doc.close()

        
    def add_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if folder:
            for root, _, files in os.walk(folder):
                for file in files:
                    if file.endswith(('.pdf', '.docx', '.doc', '.pptx', '.ppt', '.xlsx', '.xls')):
                        self.files.append(os.path.join(root, file))
            self.update_table()
    
    def add_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "选择文件",
            "",
            "Documents (*.pdf *.docx *.doc *.pptx *.ppt *.xlsx *.xls *.wps *.et *.dps)"
        )
        if files:
            self.files.extend(files)
            self.update_table()
            
    def update_table(self):
        self.table.setRowCount(len(self.files))
        self.table.setColumnCount(5)  # 确保列数为6

        # 设置列宽
        total_width = self.table.width()
        self.table.setColumnWidth(0, int(total_width * 0.5))  # 件路径列70%
        remaining_width = int(total_width * 0.5 / 4)  # 其他5列平均分配剩余30%
        for col in [1, 2, 3, 4]:
            self.table.setColumnWidth(col, remaining_width)
        
        # 填充数据
        for i, file in enumerate(self.files):
            # self.table.setItem(i, 0, QTableWidgetItem(str(i+1)))  # 只在这里设置序号
            self.table.setItem(i, 0, QTableWidgetItem(file))
            self.table.setItem(i, 1, QTableWidgetItem(os.path.splitext(file)[1]))
            self.table.setItem(i, 2, QTableWidgetItem(""))
            self.table.setItem(i, 3, QTableWidgetItem(""))
            self.table.setItem(i, 4, QTableWidgetItem("待统计"))
            
    def update_progress(self, row, file, single_pages, double_pages, status):
        """更新单个文件的处理进度"""
        self.table.setItem(row, 2, QTableWidgetItem(str(single_pages)))
        self.table.setItem(row, 3, QTableWidgetItem(str(double_pages)))
        self.table.setItem(row, 4, QTableWidgetItem(status))
    
    def start_counting(self):
        # 禁用开始按钮，避免重复点击
        self.start_btn.setEnabled(False)
        
        # 创建并启动计数线程
        self.counter_thread = CounterThread(self.files, self)
        self.counter_thread.progress.connect(self.update_progress)
        self.counter_thread.finished.connect(self.counting_finished)
        self.counter_thread.start()
    

    
    def counting_finished(self, stats):
        """处理完成后更新统计信息"""
        total_pages = sum(s['pages'] for s in stats.values())
        total_files = len(self.files)
        
        # 更新总页数显示
        self.total_pages.setText(str(total_pages))
        
        # 显示完成对话框
        QMessageBox.information(
            self,
            "处理完成",
            f"共处理 {total_files} 个文件\n"
            f"https://blog.csdn.net/vip"
        )
        
        # 重新启用开始按钮
        self.start_btn.setEnabled(True)

    def clear_table(self):
        """清空表格内容"""
        self.files.clear()
        self.update_table()
        
        
        # 重新启用开始按钮
        self.start_btn.setEnabled(True)

    def dragEnterEvent(self, event):
        """当拖拽进入窗口时触发"""
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()
            
    def dragMoveEvent(self, event):
        """当拖拽在窗口内移动时触发"""
        if event.mimeData().hasUrls():
            event.setDropAction(Qt.DropAction.CopyAction)
            event.accept()
        else:
            event.ignore()
            
    def dropEvent(self, event):
        """当放下拖拽项时触发"""
        files = []
        supported_extensions = ('.pdf', '.docx', '.doc', '.pptx', '.ppt', 
                              '.xlsx', '.xls', '.wps', '.et', '.dps')
        
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            if os.path.isfile(file_path):
                if file_path.lower().endswith(supported_extensions):
                    files.append(file_path)
            elif os.path.isdir(file_path):
                for root, _, filenames in os.walk(file_path):
                    for filename in filenames:
                        if filename.lower().endswith(supported_extensions):
                            files.append(os.path.join(root, filename))
        
        if files:
            self.files.extend(files)
            self.update_table()

    def calculate_amount(self):
        """计算总金额"""
        total = 0
        
        # 计算每种类型的金额
        for name, inputs in self.price_inputs.items():
            try:
                quantity = float(inputs['quantity'].text() or 0)
                price = float(inputs['price'].text() or 0)
                amount = quantity * price
                inputs['amount'].setText(f"={amount:.1f}")
                total += amount
            except ValueError:
                inputs['amount'].setText("=0")
        
        # 更新总金额
        self.total_amount.setText(f"{total:.1f}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = DocCounter()
    window.show()
    sys.exit(app.exec())