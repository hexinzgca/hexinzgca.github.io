import sys
import re
import os
from pathlib import Path
import pypandoc
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QTextEdit, QLabel, 
                             QFileDialog, QMessageBox, QComboBox, QProgressBar)
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QFont


class ConversionWorker(QThread):
    """转换工作线程"""
    progress = Signal(str)
    finished = Signal(str)
    error = Signal(str)
    
    def __init__(self, input_file, output_file, format_type):
        super().__init__()
        self.input_file = input_file
        self.output_file = output_file
        self.format_type = format_type
    
    def run(self):
        try:
            self.progress.emit("正在读取文件...")
            
            # 读取输入文件
            with open(self.input_file, 'r', encoding='utf-8') as f:
                content = f.read()
            
            self.progress.emit("正在分析文档格式...")
            
            # 检测文档格式
            detected_format = self.detect_math_format(content)
            
            if self.format_type == "auto":
                format_to_use = detected_format
            else:
                format_to_use = self.format_type
            
            self.progress.emit(f"检测到格式: {format_to_use}")
            
            # 转换数学公式格式
            if format_to_use == "latex":
                self.progress.emit("正在转换LaTeX格式数学公式...")
                converted_content = self.convert_latex_to_dollar(content)
                
                # 创建临时文件
                temp_file = self.input_file + ".converted.md"
                with open(temp_file, 'w', encoding='utf-8') as f:
                    f.write(converted_content)
                
                input_for_pandoc = temp_file
            else:
                input_for_pandoc = self.input_file
            
            self.progress.emit("正在转换为DOCX格式...")
            
            # 使用pypandoc转换
            pypandoc.convert_file(
                input_for_pandoc,
                'docx',
                format='markdown',  # 明确指定输入格式为markdown
                outputfile=self.output_file,
                extra_args=[
                    '--standalone',
                    '--mathml'
                ]
            )
            
            # 清理临时文件
            if format_to_use == "latex" and os.path.exists(temp_file):
                os.remove(temp_file)
            
            self.finished.emit(f"转换完成！输出文件：{self.output_file}")
            
        except Exception as e:
            self.error.emit(f"转换失败：{str(e)}")
    
    def detect_math_format(self, content):
        """检测数学公式格式"""
        # 检查是否有 \( \) 或 \[ \] 格式
        latex_inline_pattern = r'\\\('
        latex_block_pattern = r'\\\['
        
        # 检查是否有 $ $$ 格式
        dollar_inline_pattern = r'(?<!\$)\$(?!\$)[^$]+?\$(?!\$)'
        dollar_block_pattern = r'\$\$[^$]+?\$\$'
        
        if re.search(latex_inline_pattern, content) or re.search(latex_block_pattern, content):
            return "latex"
        elif re.search(dollar_inline_pattern, content) or re.search(dollar_block_pattern, content):
            return "dollar"
        else:
            return "unknown"
    
    def convert_latex_to_dollar(self, content):
        """将LaTeX格式的数学公式转换为dollar格式"""
        # 1. 转换行间公式：分别替换 \[ 和 \] 为 $$
        content = re.sub(r'\\\[', '$$', content)
        content = re.sub(r'\\\]', '$$', content)
        
        # 2. 转换行内公式：\( ... \) 替换为 $...$，并确保两侧有空格
        def replace_inline_math(match):
            expr = match.group(1).strip()
            return f' ${expr}$ '
        
        content = re.sub(r'\\\(\s*(.+?)\s*\\\)', replace_inline_math, content, flags=re.DOTALL)
        
        # 3. 清理多余的连续空格
        content = re.sub(r' +', ' ', content)
        
        return content


class MarkdownConverter(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Markdown数学公式转换器")
        self.setGeometry(100, 100, 800, 600)
        
        # 创建中心widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 创建布局
        layout = QVBoxLayout(central_widget)
        
        # 标题
        title_label = QLabel("Markdown数学公式转换器")
        title_label.setAlignment(Qt.AlignCenter)
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title_label.setFont(title_font)
        layout.addWidget(title_label)
        
        # 说明文本
        info_label = QLabel(
            "支持两种格式:\n"
            "1. LaTeX格式: 行内公式 \\( ... \\), 行间公式 \\[ ... \\]\n"
            "2. Dollar格式: 行内公式 $ ... $, 行间公式 $$ ... $$"
        )
        info_label.setStyleSheet("background-color: #f0f0f0; padding: 10px; border-radius: 5px;")
        layout.addWidget(info_label)
        
        # 文件选择部分
        file_layout = QHBoxLayout()
        
        self.file_path_label = QLabel("未选择文件")
        self.file_path_label.setStyleSheet("border: 1px solid #ccc; padding: 5px; background-color: white;")
        file_layout.addWidget(self.file_path_label)
        
        self.select_file_btn = QPushButton("选择Markdown文件")
        self.select_file_btn.clicked.connect(self.select_input_file)
        file_layout.addWidget(self.select_file_btn)
        
        layout.addLayout(file_layout)
        
        # 格式选择
        format_layout = QHBoxLayout()
        format_layout.addWidget(QLabel("格式检测:"))
        
        self.format_combo = QComboBox()
        self.format_combo.addItems(["自动检测", "强制LaTeX格式", "强制Dollar格式"])
        format_layout.addWidget(self.format_combo)
        
        layout.addLayout(format_layout)
        
        # 预览区域
        preview_label = QLabel("文件预览:")
        layout.addWidget(preview_label)
        
        self.preview_text = QTextEdit()
        self.preview_text.setMaximumHeight(200)
        self.preview_text.setReadOnly(True)
        layout.addWidget(self.preview_text)
        
        # 转换按钮
        self.convert_btn = QPushButton("转换为DOCX")
        self.convert_btn.clicked.connect(self.convert_file)
        self.convert_btn.setEnabled(False)
        layout.addWidget(self.convert_btn)
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)
        
        # 状态标签
        self.status_label = QLabel("请选择Markdown文件")
        self.status_label.setStyleSheet("color: #666; font-style: italic;")
        layout.addWidget(self.status_label)
        
        # 存储选择的文件路径
        self.input_file_path = None
    
    def select_input_file(self):
        """选择输入文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择Markdown文件",
            "",
            "Markdown files (*.md *.txt);;All files (*.*)"
        )
        
        if file_path:
            self.input_file_path = file_path
            self.file_path_label.setText(os.path.basename(file_path))
            self.convert_btn.setEnabled(True)
            self.load_preview()
            self.status_label.setText("文件已选择，可以开始转换")
    
    def load_preview(self):
        """加载文件预览"""
        if self.input_file_path:
            try:
                with open(self.input_file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                # 只显示前500个字符
                preview_content = content[:500]
                if len(content) > 500:
                    preview_content += "\\n\\n... (文件内容较长，仅显示前500个字符)"
                
                self.preview_text.setPlainText(preview_content)
                
            except Exception as e:
                self.preview_text.setPlainText(f"无法预览文件: {str(e)}")
    
    def convert_file(self):
        """转换文件"""
        if not self.input_file_path:
            QMessageBox.warning(self, "警告", "请先选择输入文件")
            return
        
        # 选择输出文件
        output_file, _ = QFileDialog.getSaveFileName(
            self,
            "保存DOCX文件",
            os.path.splitext(self.input_file_path)[0] + ".docx",
            "Word documents (*.docx);;All files (*.*)"
        )
        
        if not output_file:
            return
        
        # 获取格式类型
        format_mapping = {
            "自动检测": "auto",
            "强制LaTeX格式": "latex", 
            "强制Dollar格式": "dollar"
        }
        format_type = format_mapping[self.format_combo.currentText()]
        
        # 开始转换
        self.convert_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)  # 无确定进度
        
        # 创建工作线程
        self.worker = ConversionWorker(self.input_file_path, output_file, format_type)
        self.worker.progress.connect(self.update_status)
        self.worker.finished.connect(self.conversion_finished)
        self.worker.error.connect(self.conversion_error)
        self.worker.start()
    
    def update_status(self, message):
        """更新状态"""
        self.status_label.setText(message)
    
    def conversion_finished(self, message):
        """转换完成"""
        self.progress_bar.setVisible(False)
        self.convert_btn.setEnabled(True)
        self.status_label.setText(message)
        QMessageBox.information(self, "成功", message)
    
    def conversion_error(self, error_message):
        """转换错误"""
        self.progress_bar.setVisible(False)
        self.convert_btn.setEnabled(True)
        self.status_label.setText(f"转换失败: {error_message}")
        QMessageBox.critical(self, "错误", error_message)


def main():
    app = QApplication(sys.argv)
    
    # 检查pypandoc是否可用
    try:
        pypandoc.get_pandoc_version()
    except OSError:
        QMessageBox.critical(
            None, 
            "错误", 
            "未找到pandoc！请先安装pandoc:\\n"
            "1. 访问 https://pandoc.org/installing.html\\n"
            "2. 下载并安装pandoc\\n"
            "3. 重启应用程序"
        )
        return
    
    window = MarkdownConverter()
    window.show()
    
    sys.exit(app.exec())


if __name__ == "__main__":
    main() 
