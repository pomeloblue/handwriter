from PIL import Image, ImageDraw, ImageFont
import docx2txt
import textwrap
import os
import random
import json
from pathlib import Path
from PyQt6.QtWidgets import (QApplication, QMainWindow, QPushButton, QLabel,
                            QFileDialog, QVBoxLayout, QHBoxLayout, QWidget, 
                            QProgressBar, QMessageBox, QLineEdit, QSpinBox,
                            QComboBox, QDoubleSpinBox, QGridLayout)  
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QPixmap, QImage, QDragEnterEvent, QDropEvent

class StyleSheet:
    """样式表"""
    MAIN_WINDOW = """
        QMainWindow {
            background-color: #f5f5f5;
        }
    """
    
    WIDGET = """
        QWidget {
            background-color: #f5f5f5;
            font-family: "Microsoft YaHei";
        }
    """
    
    BUTTON = """
        QPushButton {
            background-color: #2196F3;
            color: white;
            border: none;
            padding: 8px 20px;
            border-radius: 4px;
            font-weight: bold;
            font-size: 14px;
        }
        QPushButton:hover {
            background-color: #1976D2;
        }
        QPushButton:pressed {
            background-color: #0D47A1;
        }
        QPushButton:disabled {
            background-color: #BDBDBD;
        }
    """
    
    ACTION_BUTTON = """
        QPushButton {
            background-color: #4CAF50;
            color: white;
            border: none;
            padding: 8px 20px;
            border-radius: 4px;
            font-weight: bold;
            font-size: 14px;
        }
        QPushButton:hover {
            background-color: #388E3C;
        }
        QPushButton:pressed {
            background-color: #1B5E20;
        }
        QPushButton:disabled {
            background-color: #BDBDBD;
        }
    """
    
    SPIN_BOX = """
        QSpinBox, QDoubleSpinBox {
            border: 1px solid #BDBDBD;
            border-radius: 4px;
            padding: 4px;
            background: white;
            min-width: 80px;
        }
        QSpinBox:hover, QDoubleSpinBox:hover {
            border: 1px solid #2196F3;
        }
        QSpinBox:focus, QDoubleSpinBox:focus {
            border: 2px solid #2196F3;
        }
    """
    
    COMBO_BOX = """
        QComboBox {
            border: 1px solid #BDBDBD;
            border-radius: 4px;
            padding: 1px 4px;
            background: white;
            min-width: 150px;
            
        }
        QComboBox:hover {
            border: 1px solid #2196F3;
        }
        QComboBox:focus {
            border: 2px solid #2196F3;
        }
        QComboBox::drop-down {
            border: none;
        }
        QComboBox::down-arrow {
            image: url(down_arrow.png);
            width: 12px;
            height: 12px;
        }
    """
    
    LINE_EDIT = """
        QLineEdit {
            border: 1px solid #BDBDBD;
            border-radius: 4px;
            padding: 6px;
            background: white;
        }
        QLineEdit:hover {
            border: 1px solid #2196F3;
        }
        QLineEdit:focus {
            border: 2px solid #2196F3;
        }
    """
    
    PREVIEW_LABEL = """
        QLabel {
            background-color: white;
            border: 1px solid #BDBDBD;
            border-radius: 4px;
        }
    """

    PROGRESS_BAR = """
        QProgressBar {
            border: 1px solid #BDBDBD;
            border-radius: 4px;
            text-align: center;
            background-color: #f5f5f5;
            height: 20px;
        }
        QProgressBar::chunk {
            background-color: #2196F3;
            border-radius: 3px;
        }
    """

    LABEL = """
        QLabel {
            margin: 0;
            padding: 0;
        }
    """

class PreviewWidget(QLabel):
    """预览窗口"""
    def __init__(self):
        super().__init__()
        self.setMinimumSize(400, 500)
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setText('预览区域\n\n拖拽文件到此处或点击"选择文件"')
        self.setStyleSheet(StyleSheet.PREVIEW_LABEL)
        self.setAcceptDrops(True)
  
    def update_preview(self, background_path, font_path, params):
        """更新预览图像"""
        try:
            # 创建预览图像
            background = Image.open(background_path)
            # 调整预览图像大小
            preview_width = 400
            ratio = preview_width / background.width
            preview_height = int(background.height * ratio)
            background = background.resize((preview_width, preview_height))
            
            draw = ImageDraw.Draw(background)
            font = ImageFont.truetype(font_path, int(params['font_size'] * ratio))
            
            # 使用参数中的预览文本
            preview_text = params.get('preview_text', "预览文本\n第二行文本")
            
            # 绘制文本
            x = params['left_margin'] * ratio
            y = params['top_margin'] * ratio
            
            # 添加随机扰动
            x += random.gauss(0, params['perturb_x_sigma'] * ratio)
            y += random.gauss(0, params['perturb_y_sigma'] * ratio)
            
            draw.text((x, y), preview_text, font=font, fill=(0, 0, 0))
            
            # 转换为QPixmap并显示
            img = background.convert('RGB')
            data = img.tobytes("raw", "RGB")
            qimg = QImage(data, img.width, img.height, QImage.Format.Format_RGB888)
            pixmap = QPixmap.fromImage(qimg)
            self.setPixmap(pixmap)
        except Exception as e:
            self.setText(f"预览失败: {str(e)}")

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        if files:
            file_path = files[0]
            if file_path.lower().endswith(('.txt', '.doc', '.docx')):
                self.parent().parent().input_path.setText(file_path)
                self.parent().parent().update_preview()
            else:
                QMessageBox.warning(self, "错误", "不支持的文件格式！")

class HandwritingConverter(QThread):
    """后台转换线程"""
    progress = pyqtSignal(int, str)  # 进度值和进度信息
    finished = pyqtSignal(bool, str)

    def __init__(self, input_file, output_dir, params):
        super().__init__()
        self.input_file = input_file
        self.output_dir = output_dir
        self.params = params
        self.is_running = True

    def run(self):
        try:
            # 读取文本内容
            self.progress.emit(10, "正在读取文件...")
            text_content = self.read_text_from_file(self.input_file)
            
            # 分割文本为段落
            paragraphs = text_content.split('\n')
            total_paragraphs = len(paragraphs)
            
            # 创建输出目录
            output_base = os.path.join(
                self.output_dir,
                f"handwritten_{os.path.splitext(os.path.basename(self.input_file))[0]}"
            )
            os.makedirs(output_base, exist_ok=True)
            
            # 开始转换
            self.progress.emit(20, "正在转换...")
            current_page = 1
            current_y = self.params['top_margin']
            current_text = []
            
            background = Image.open(self.params['background_path'])
            draw = ImageDraw.Draw(background)
            font = ImageFont.truetype(self.params['font_path'], self.params['font_size'])
            
            for i, paragraph in enumerate(paragraphs):
                if not self.is_running:
                    raise InterruptedError("转换已取消")
                
                # 处理段落文本换行
                max_width = background.width - self.params['left_margin'] - self.params['right_margin']
                wrapped_lines = textwrap.wrap(paragraph, width=int(max_width / (self.params['font_size'] * 0.5)))
                
                for line in wrapped_lines:
                    # 检查是否需要新页
                    if current_y + self.params['font_size'] > background.height - self.params['bottom_margin']:
                        # 保存当前页
                        output_path = os.path.join(output_base, f"page_{current_page:03d}.png")
                        background.save(output_path)
                        current_page += 1
                        
                        # 创建新页
                        background = Image.open(self.params['background_path'])
                        draw = ImageDraw.Draw(background)
                        current_y = self.params['top_margin']
                    
                    # 写入当前行
                    x = self.params['left_margin']
                    for char in line:
                        # 添加随机扰动
                        dx = random.gauss(0, self.params['perturb_x_sigma'])
                        dy = random.gauss(0, self.params['perturb_y_sigma'])
                        theta = random.gauss(0, self.params['perturb_theta_sigma'])
                        
                        # 绘制字符
                        draw.text(
                            (x + dx, current_y + dy),
                            char,
                            font=font,
                            fill=(0, 0, 0)
                        )
                        
                        # 计算字符宽度并添加随机间距
                        char_width = draw.textlength(char, font=font)
                        x += char_width + self.params['word_spacing'] + random.gauss(0, self.params['word_spacing_sigma'])
                    
                    current_y += self.params['line_spacing'] + random.gauss(0, self.params['line_spacing_sigma'])
                
                # 段落间距
                if i < len(paragraphs) - 1:
                    current_y += self.params['line_spacing'] * 1.5
                
                # 更新进度
                progress = int(20 + (i / total_paragraphs) * 70)
                self.progress.emit(progress, f"正在处理第 {current_page} 页...")
            
            # 保存最后一页
            output_path = os.path.join(output_base, f"page_{current_page:03d}.png")
            background.save(output_path)
            
            self.progress.emit(100, "转换完成！")
            self.finished.emit(True, f"转换完成！共生成 {current_page} 页")
            
        except InterruptedError as e:
            self.finished.emit(False, str(e))
        except Exception as e:
            self.finished.emit(False, f"错误: {str(e)}")

    def read_text_from_file(self, file_path):
        """读取文件内容"""
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"找不到文件: {file_path}")
            
        file_ext = os.path.splitext(file_path)[1].lower()
        
        if file_ext == '.txt':
            with open(file_path, 'r', encoding='utf-8') as f:
                return f.read()
        elif file_ext in ['.docx', '.doc']:
            return docx2txt.process(file_path)
        else:
            raise ValueError(f"不支持的文件格式: {file_ext}")

    def stop(self):
        """停止转换"""
        self.is_running = False

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.converter = None
        self.preview_pages = []  # 存储预览页面的内容
        self.current_preview_page = 0  # 当前预览页码
        self.initUI()
        self.create_required_directories()
        self.load_default_params()

    def initUI(self):
        """初始化UI"""
        self.setWindowTitle('手写模拟器 ')
        self.setMinimumSize(900, 600)
        self.setStyleSheet(StyleSheet.MAIN_WINDOW)

        # 创建中心部件和主布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)
        main_layout.setSpacing(20)

        # 左侧控制面板
        control_panel = QWidget()
        control_layout = QVBoxLayout(control_panel)
        control_layout.setSpacing(10)
        main_layout.addWidget(control_panel)

        # 文件选择区域
        file_layout = QHBoxLayout()
        self.input_path = QLineEdit()
        self.input_path.setStyleSheet(StyleSheet.LINE_EDIT)
        self.input_path.setPlaceholderText('选择要转换的文件...')
        file_layout.addWidget(self.input_path)
        
        self.select_file_btn = QPushButton('选择文件')
        self.select_file_btn.setStyleSheet(StyleSheet.BUTTON) 
        self.select_file_btn.clicked.connect(self.select_input_file)
        file_layout.addWidget(self.select_file_btn)
        control_layout.addLayout(file_layout)

        # 保存路径
        save_layout = QHBoxLayout()
        self.output_path = QLineEdit()
        self.output_path.setStyleSheet(StyleSheet.LINE_EDIT)
        self.output_path.setPlaceholderText('选择保存位置...')
        save_layout.addWidget(self.output_path)
        
        self.select_output_btn = QPushButton('保存路径')
        self.select_output_btn.setStyleSheet(StyleSheet.ACTION_BUTTON)
        self.select_output_btn.clicked.connect(self.select_output_dir)
        save_layout.addWidget(self.select_output_btn)
        control_layout.addLayout(save_layout)
 
        # 字体和背景选择
        font_bg_layout = QGridLayout()
        font_bg_layout.setHorizontalSpacing(20)  # 列之间的间距
        font_bg_layout.setVerticalSpacing(3)  # 行之间的间距
        # 字体选择
        font_label = QLabel('字体')
        font_label.setAlignment(Qt.AlignmentFlag.AlignBottom) 
        font_bg_layout.addWidget(font_label, 0, 0)  # 第0行，第0列
        self.font_combo = QComboBox()
        self.font_combo.setStyleSheet(StyleSheet.COMBO_BOX)
        self.load_fonts()
        font_bg_layout.addWidget(self.font_combo, 1, 0)  # 第1行，第0列
        
        # 背景选择
        bg_label = QLabel('背景')
        bg_label.setAlignment(Qt.AlignmentFlag.AlignBottom)
        font_bg_layout.addWidget(bg_label, 0, 1)  # 第0行，第1列
        self.bg_combo = QComboBox()
        self.bg_combo.setStyleSheet(StyleSheet.COMBO_BOX)
        self.load_backgrounds()
        font_bg_layout.addWidget(self.bg_combo, 1, 1)  # 第1行，第1列
        
        # 设置列宽度
        font_bg_layout.setColumnMinimumWidth(0, 150)  # 第一列最小宽度
        font_bg_layout.setColumnMinimumWidth(1, 150)  # 第二列最小宽度

        control_layout.addLayout(font_bg_layout)
        # 参数设置区域
        params_layout = QVBoxLayout()
        params_layout.setSpacing(10)
        
        # 字体大小
        font_size_layout = QHBoxLayout()
        font_size_layout.addWidget(QLabel('字体大小'))
        self.font_size_spin = QSpinBox()
        self.font_size_spin.setRange(20, 600)
        self.font_size_spin.setValue(50)
        self.font_size_spin.setStyleSheet(StyleSheet.SPIN_BOX)
        font_size_layout.addWidget(self.font_size_spin)
        params_layout.addLayout(font_size_layout)

        # 边距设置
        margins_layout = QGridLayout()
        margins_layout.setSpacing(10)
        
        # 上边距
        margins_layout.addWidget(QLabel('上边距'), 0, 0)
        self.top_margin_spin = QSpinBox()
        self.top_margin_spin.setRange(0, 500)
        self.top_margin_spin.setValue(140)
        self.top_margin_spin.setStyleSheet(StyleSheet.SPIN_BOX)
        margins_layout.addWidget(self.top_margin_spin, 0, 1)
        
        # 下边距
        margins_layout.addWidget(QLabel('下边距'), 1, 0)
        self.bottom_margin_spin = QSpinBox()
        self.bottom_margin_spin.setRange(0, 500)
        self.bottom_margin_spin.setValue(70)
        self.bottom_margin_spin.setStyleSheet(StyleSheet.SPIN_BOX)
        margins_layout.addWidget(self.bottom_margin_spin, 1, 1)
        
        # 左边距
        margins_layout.addWidget(QLabel('左边距'), 0, 2)
        self.left_margin_spin = QSpinBox()
        self.left_margin_spin.setRange(0, 500)
        self.left_margin_spin.setValue(100)
        self.left_margin_spin.setStyleSheet(StyleSheet.SPIN_BOX)
        margins_layout.addWidget(self.left_margin_spin, 0, 3)
        
        # 右边距
        margins_layout.addWidget(QLabel('右边距'), 1, 2)
        self.right_margin_spin = QSpinBox()
        self.right_margin_spin.setRange(0, 500)
        self.right_margin_spin.setValue(100)
        self.right_margin_spin.setStyleSheet(StyleSheet.SPIN_BOX)
        margins_layout.addWidget(self.right_margin_spin, 1, 3)
        
        params_layout.addLayout(margins_layout)

        # 间距设置
        spacing_layout = QGridLayout()
        spacing_layout.setSpacing(10)
        
        # 字间距
        spacing_layout.addWidget(QLabel('字间距'), 0, 0)
        self.word_spacing_spin = QSpinBox()
        self.word_spacing_spin.setRange(0, 50)
        self.word_spacing_spin.setValue(5)
        self.word_spacing_spin.setStyleSheet(StyleSheet.SPIN_BOX)
        spacing_layout.addWidget(self.word_spacing_spin, 0, 1)
        
        # 行间距
        spacing_layout.addWidget(QLabel('行间距'), 0, 2)
        self.line_spacing_spin = QSpinBox()
        self.line_spacing_spin.setRange(0, 500)
        self.line_spacing_spin.setValue(143)
        self.line_spacing_spin.setStyleSheet(StyleSheet.SPIN_BOX)
        spacing_layout.addWidget(self.line_spacing_spin, 0, 3)
        
        params_layout.addLayout(spacing_layout)

        # 扰动参数
        perturb_layout = QGridLayout()
        perturb_layout.setSpacing(10)
        
        # 水平扰动
        perturb_layout.addWidget(QLabel('水平扰动'), 0, 0)
        self.perturb_x_spin = QDoubleSpinBox()
        self.perturb_x_spin.setRange(0, 10)
        self.perturb_x_spin.setValue(3)
        self.perturb_x_spin.setStyleSheet(StyleSheet.SPIN_BOX)
        perturb_layout.addWidget(self.perturb_x_spin, 0, 1)
        
        # 垂直扰动
        perturb_layout.addWidget(QLabel('垂直扰动'), 0, 2)
        self.perturb_y_spin = QDoubleSpinBox()
        self.perturb_y_spin.setRange(0, 10)
        self.perturb_y_spin.setValue(3)
        self.perturb_y_spin.setStyleSheet(StyleSheet.SPIN_BOX)
        perturb_layout.addWidget(self.perturb_y_spin, 0, 3)
        
        params_layout.addLayout(perturb_layout)
        control_layout.addLayout(params_layout)

        # 预览控制按钮
        preview_control_layout = QHBoxLayout()
        
        self.prev_page_btn = QPushButton('上一页')
        self.prev_page_btn.setStyleSheet(StyleSheet.BUTTON)
        self.prev_page_btn.clicked.connect(self.prev_preview_page)
        self.prev_page_btn.setEnabled(False)
        preview_control_layout.addWidget(self.prev_page_btn)
        
        self.page_label = QLabel('第 1 页')
        self.page_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.page_label.setStyleSheet(StyleSheet.LABEL)
        preview_control_layout.addWidget(self.page_label)
        
        self.next_page_btn = QPushButton('下一页')
        self.next_page_btn.setStyleSheet(StyleSheet.BUTTON)
        self.next_page_btn.clicked.connect(self.next_preview_page)
        self.next_page_btn.setEnabled(False)
        preview_control_layout.addWidget(self.next_page_btn)
        
        control_layout.addLayout(preview_control_layout)

        # 预览和转换按钮
        button_layout = QHBoxLayout()
        
        self.preview_btn = QPushButton('预览')
        self.preview_btn.setStyleSheet(StyleSheet.BUTTON)
        self.preview_btn.clicked.connect(self.update_preview)
        button_layout.addWidget(self.preview_btn)
        
        self.convert_btn = QPushButton('开始转换')
        self.convert_btn.setStyleSheet(StyleSheet.ACTION_BUTTON)
        self.convert_btn.clicked.connect(self.start_conversion)
        button_layout.addWidget(self.convert_btn)
        
        control_layout.addLayout(button_layout)

        # 进度条
        self.progress = QProgressBar()
        self.progress.setStyleSheet(StyleSheet.PROGRESS_BAR)
        control_layout.addWidget(self.progress)

        # 右侧预览区域
        self.preview = PreviewWidget()
        main_layout.addWidget(self.preview)

        # 连接信号
        self.font_combo.currentIndexChanged.connect(self.update_preview)
        self.bg_combo.currentIndexChanged.connect(self.update_preview)
        self.font_size_spin.valueChanged.connect(self.update_preview)
        
        # 设置布局比例
        main_layout.setStretch(0, 1)  # 控制面板
        main_layout.setStretch(1, 2)  # 预览区域
    def create_required_directories(self):
        """创建必要的目录"""
        current_dir = os.path.dirname(os.path.abspath(__file__))
        
        # 创建字体目录
        fonts_dir = os.path.join(current_dir, "fonts")
        os.makedirs(fonts_dir, exist_ok=True)
        
        # 创建背景图片目录
        bg_dir = os.path.join(current_dir, "Background")
        os.makedirs(bg_dir, exist_ok=True)
        
        # 创建参数配置目录
        param_dir = os.path.join(current_dir, "Parameter")
        os.makedirs(param_dir, exist_ok=True)

    def load_default_params(self):
        """加载默认参数"""
        try:
            current_dir = os.path.dirname(os.path.abspath(__file__))
            default_param_path = os.path.join(current_dir, "Parameter", "default.json")
            
            if os.path.exists(default_param_path):
                with open(default_param_path, 'r', encoding='utf-8') as f:
                    params = json.load(f)
                self.update_params_from_config(params)
            else:
                # 创建默认参数文件
                default_params = {
                    'font_size': 40,
                    'line_spacing': 143,
                    'word_spacing': 5,
                    'left_margin': 180,
                    'right_margin': 100,
                    'top_margin': 140,
                    'bottom_margin': 70,
                    'word_spacing_sigma': 2,
                    'line_spacing_sigma': 0,
                    'perturb_x_sigma': 3,
                    'perturb_y_sigma': 3,
                    'perturb_theta_sigma': 0.05
                }
                
                with open(default_param_path, 'w', encoding='utf-8') as f:
                    json.dump(default_params, f, indent=4, ensure_ascii=False)
        
        except Exception as e:
            QMessageBox.warning(self, "警告", f"加载默认参数失败: {str(e)}")

    def save_current_params(self):
        """保存当前参数为默认值"""
        try:
            current_dir = os.path.dirname(os.path.abspath(__file__))
            default_param_path = os.path.join(current_dir, "Parameter", "default.json")
            
            params = self.get_current_params()
            with open(default_param_path, 'w', encoding='utf-8') as f:
                json.dump(params, f, indent=4, ensure_ascii=False)
            
            QMessageBox.information(self, "成功", "参数已保存为默认值")
        
        except Exception as e:
            QMessageBox.warning(self, "错误", f"保存参数失败: {str(e)}")

    def select_input_file(self):
        """选择输入文件"""
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            "选择文件",
            "",
            "文本文件 (*.txt);;Word文档 (*.docx *.doc)"
        )
        if file_name:
            self.input_path.setText(file_name)
            self.update_preview()

    def select_output_dir(self):
        """选择输出目录"""
        dir_name = QFileDialog.getExistingDirectory(
            self,
            "选择输出目录",
            ""
        )
        if dir_name:
            self.output_path.setText(dir_name)

    def update_preview(self):
        """更新预览图"""
        if not self.bg_combo.currentData() or not self.font_combo.currentData():
            return
                
        try:
            # 获取当前文件内容
            preview_text = ""
            if self.input_path.text() and os.path.exists(self.input_path.text()):
                file_path = self.input_path.text()
                if file_path.lower().endswith('.txt'):
                    with open(file_path, 'r', encoding='utf-8') as f:
                        content = f.read()
                elif file_path.lower().endswith(('.docx', '.doc')):
                    content = docx2txt.process(file_path)
                else:
                    content = "不支持的文件格式"
                
             # 分页处理
                self.preview_pages = self.split_content_to_pages(content)
                self.current_preview_page = 0
                
                # 更新翻页按钮状态
                self.update_page_controls()
                
                # 显示当前页
                self.show_current_preview_page()
            else:
                # 无文件时显示默认预览
                preview_text = "预览文本\n第二行文本"
                params = self.get_current_params()
                params['preview_text'] = preview_text
                self.preview.update_preview(
                    background_path=self.bg_combo.currentData(),
                    font_path=self.font_combo.currentData(),
                    params=params
                )
                
        except Exception as e:
            QMessageBox.warning(self, "预览失败", str(e))


    def start_conversion(self):
        """开始转换"""
        if not self.input_path.text():
            QMessageBox.warning(self, "警告", "请选择输入文件！")
            return
        if not self.output_path.text():
            QMessageBox.warning(self, "警告", "请选择输出目录！")
            return
        if not self.font_combo.currentData():
            QMessageBox.warning(self, "警告", "请选择字体！")
            return
        if not self.bg_combo.currentData():
            QMessageBox.warning(self, "警告", "请选择背景图片！")
            return

        # 禁用转换按钮
        self.convert_btn.setEnabled(False)
        self.progress.setValue(0)

        # 准备参数
        params = {
            'font_path': self.font_combo.currentData(),
            'background_path': self.bg_combo.currentData(),
            **self.get_current_params()
        }

        # 创建并启动转换线程
        self.converter = HandwritingConverter(
            self.input_path.text(),
            self.output_path.text(),
            params
        )
        self.converter.progress.connect(self.update_conversion_progress)
        self.converter.finished.connect(self.conversion_finished)
        self.converter.start()

    def update_conversion_progress(self, value, message):
        """更新转换进度"""
        self.progress.setValue(value)
        self.progress.setFormat(f"{message} ({value}%)")

    def conversion_finished(self, success, message):
        """转换完成处理"""
        self.convert_btn.setEnabled(True)
        
        if success:
            QMessageBox.information(self, "完成", message)
            # 打开输出目录
            os.startfile(self.output_path.text())
        else:
            QMessageBox.warning(self, "错误", message)
    def load_fonts(self):
        """加载字体列表"""
        self.font_combo.clear()
        current_dir = os.path.dirname(os.path.abspath(__file__))
        fonts_dir = os.path.join(current_dir, "fonts")
        
        if os.path.exists(fonts_dir):
            for font_file in os.listdir(fonts_dir):
                if font_file.lower().endswith('.ttf'):
                    font_path = os.path.join(fonts_dir, font_file)
                    self.font_combo.addItem(font_file, font_path)
    def get_current_params(self):
        """获取当前参数设置"""
        return {
            'font_size': self.font_size_spin.value(),
            'line_spacing': self.line_spacing_spin.value(),
            'word_spacing': self.word_spacing_spin.value(),
            'left_margin': self.left_margin_spin.value(),
            'right_margin': self.right_margin_spin.value(),
            'top_margin': self.top_margin_spin.value(),
            'bottom_margin': self.bottom_margin_spin.value(),
            'perturb_x_sigma': self.perturb_x_spin.value(),
            'perturb_y_sigma': self.perturb_y_spin.value(),
            'perturb_theta_sigma': 0.05,  # 角度扰动固定值
            'word_spacing_sigma': 2,  # 字间距扰动固定值
            'line_spacing_sigma': 0,  # 行间距扰动固定值
        }

    def update_params_from_config(self, params):
        """从配置更新参数"""
        try:
            self.font_size_spin.setValue(params.get('font_size', 40))
            self.line_spacing_spin.setValue(params.get('line_spacing', 143))
            self.word_spacing_spin.setValue(params.get('word_spacing', 5))
            self.left_margin_spin.setValue(params.get('left_margin', 180))
            self.right_margin_spin.setValue(params.get('right_margin', 100))
            self.top_margin_spin.setValue(params.get('top_margin', 140))
            self.bottom_margin_spin.setValue(params.get('bottom_margin', 70))
            self.perturb_x_spin.setValue(params.get('perturb_x_sigma', 3))
            self.perturb_y_spin.setValue(params.get('perturb_y_sigma', 3))
        except Exception as e:
            QMessageBox.warning(self, "警告", f"加载参数失败: {str(e)}")
    def load_backgrounds(self):
        """加载背景图片列表"""
        self.bg_combo.clear()
        current_dir = os.path.dirname(os.path.abspath(__file__))
        bg_dir = os.path.join(current_dir, "Background")
        
        if os.path.exists(bg_dir):
            for bg_file in os.listdir(bg_dir):
                if bg_file.lower().endswith(('.png', '.jpg', '.jpeg')):
                    bg_path = os.path.join(bg_dir, bg_file)
                    self.bg_combo.addItem(bg_file, bg_path)
    def closeEvent(self, event):
        """关闭窗口时保存参数"""
        self.save_current_params()
        if self.converter and self.converter.isRunning():
            self.converter.stop()
            self.converter.wait()
        event.accept()

    def split_content_to_pages(self, content):
        """将内容分割成页"""
        # 获取当前参数
        params = self.get_current_params()
        
        # 计算每页能容纳的行数
        background = Image.open(self.bg_combo.currentData())
        available_height = background.height - params['top_margin'] - params['bottom_margin']
        line_height = params['font_size'] + params['line_spacing']
        lines_per_page = int(available_height / line_height)
        
        # 分割内容
        lines = content.split('\n')
        pages = []
        current_page = []
        line_count = 0
        
        for line in lines:
            if line_count >= lines_per_page:
                pages.append('\n'.join(current_page))
                current_page = []
                line_count = 0
            current_page.append(line)
            line_count += 1
        
        if current_page:
            pages.append('\n'.join(current_page))
        
        return pages

    def show_current_preview_page(self):
        """显示当前预览页"""
        if not self.preview_pages:
            return
            
        params = self.get_current_params()
        params['preview_text'] = self.preview_pages[self.current_preview_page]
        
        self.preview.update_preview(
            background_path=self.bg_combo.currentData(),
            font_path=self.font_combo.currentData(),
            params=params
        )
        
        # 更新页码显示
        self.page_label.setText(f'第 {self.current_preview_page + 1} 页 / 共 {len(self.preview_pages)} 页')

    def update_page_controls(self):
        """更新翻页按钮状态"""
        has_pages = len(self.preview_pages) > 0
        self.prev_page_btn.setEnabled(has_pages and self.current_preview_page > 0)
        self.next_page_btn.setEnabled(has_pages and self.current_preview_page < len(self.preview_pages) - 1)

    def prev_preview_page(self):
        """显示上一页"""
        if self.current_preview_page > 0:
            self.current_preview_page -= 1
            self.show_current_preview_page()
            self.update_page_controls()

    def next_preview_page(self):
        """显示下一页"""
        if self.current_preview_page < len(self.preview_pages) - 1:
            self.current_preview_page += 1
            self.show_current_preview_page()
            self.update_page_controls()
def main():
    import sys
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()