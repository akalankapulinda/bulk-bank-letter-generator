import sys
import os
import platform
import subprocess
from pathlib import Path
from typing import Optional, List
import pandas as pd
from docxtpl import DocxTemplate
from threading import Thread
from queue import Queue

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QHBoxLayout, QVBoxLayout,
    QLabel, QPushButton, QScrollArea, QFileDialog, QMessageBox,
    QProgressBar, QFrame, QGridLayout
)
from PyQt5.QtCore import Qt, pyqtSignal, QObject, QTimer, QSize, QRect
from PyQt5.QtGui import QColor, QFont, QIcon, QPixmap, QCursor, QPainter, QLinearGradient, QRadialGradient
from PyQt5.QtCore import QMimeData


class ProcessingSignals(QObject):
    """Signal emitter for background processing"""
    progress_updated = pyqtSignal(int)
    status_updated = pyqtSignal(str)
    file_processed = pyqtSignal(dict)
    processing_complete = pyqtSignal(int, int)
    error_occurred = pyqtSignal(str)


class CyberpunkTheme:
    """Enhanced cyberpunk color scheme with gradients"""
    BACKGROUND = "#0a0e27"
    DARK_BG = "#0f1535"
    PANEL_BG = "#1a1f3a"
    CYAN = "#00d4ff"
    MAGENTA = "#ff006e"
    YELLOW = "#ffbe0b"
    GREEN = "#00ff41"
    LIGHT_TEXT = "#e0e0e0"
    DARK_TEXT = "#808080"
    ACCENT_CYAN = "#00ffd4"
    PURPLE = "#9d00ff"
    DEEP_BLUE = "#001a4d"

    @staticmethod
    def get_stylesheet():
        return f"""
        QMainWindow {{
            background-color: {CyberpunkTheme.BACKGROUND};
            color: {CyberpunkTheme.LIGHT_TEXT};
        }}
        
        QWidget {{
            background-color: {CyberpunkTheme.BACKGROUND};
            color: {CyberpunkTheme.LIGHT_TEXT};
        }}
        
        QLabel {{
            color: {CyberpunkTheme.LIGHT_TEXT};
        }}
        
        QFrame {{
            background-color: {CyberpunkTheme.PANEL_BG};
            border: 2px solid {CyberpunkTheme.CYAN};
            border-radius: 12px;
        }}
        
        QPushButton {{
            background-color: {CyberpunkTheme.PANEL_BG};
            color: {CyberpunkTheme.CYAN};
            border: 2px solid {CyberpunkTheme.CYAN};
            border-radius: 8px;
            padding: 8px 16px;
            font-weight: bold;
            font-size: 12px;
        }}
        
        QPushButton:hover {{
            background-color: {CyberpunkTheme.CYAN};
            color: {CyberpunkTheme.BACKGROUND};
            border: 2px solid {CyberpunkTheme.CYAN};
        }}
        
        QPushButton:pressed {{
            background-color: {CyberpunkTheme.ACCENT_CYAN};
        }}
        
        QProgressBar {{
            background-color: {CyberpunkTheme.PANEL_BG};
            border: 2px solid {CyberpunkTheme.CYAN};
            border-radius: 8px;
            text-align: center;
            color: {CyberpunkTheme.YELLOW};
            height: 20px;
        }}
        
        QProgressBar::chunk {{
            background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                       stop:0 {CyberpunkTheme.GREEN},
                                       stop:0.5 {CyberpunkTheme.YELLOW},
                                       stop:1 {CyberpunkTheme.CYAN});
            border-radius: 6px;
        }}
        
        QScrollArea {{
            background-color: {CyberpunkTheme.BACKGROUND};
            border: none;
        }}
        
        QScrollBar:vertical {{
            background-color: {CyberpunkTheme.BACKGROUND};
            width: 12px;
            border: none;
        }}
        
        QScrollBar::handle:vertical {{
            background-color: {CyberpunkTheme.CYAN};
            border-radius: 6px;
            min-height: 20px;
        }}
        
        QScrollBar::handle:vertical:hover {{
            background-color: {CyberpunkTheme.ACCENT_CYAN};
        }}
        """


class FileInfoFrame(QFrame):
    """Display loaded file information prominently"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()
        self.file_path = None
        self.hide()

    def setup_ui(self):
        """Setup file info display"""
        self.setStyleSheet(
            f"""
            QFrame {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                           stop:0 rgba(0, 212, 255, 0.15),
                                           stop:1 rgba(0, 212, 255, 0.05));
                border: 2px solid {CyberpunkTheme.CYAN};
                border-radius: 10px;
            }}
            """
        )
        
        layout = QVBoxLayout()
        layout.setContentsMargins(15, 12, 15, 12)
        layout.setSpacing(8)

        # Title
        title_label = QLabel("📁 LOADED FILE:")
        title_label.setFont(QFont("Courier New", 11, QFont.Bold))
        title_label.setStyleSheet(f"color: {CyberpunkTheme.DARK_TEXT};")
        layout.addWidget(title_label)

        # File name display
        self.file_name_label = QLabel("")
        self.file_name_label.setFont(QFont("Courier New", 14, QFont.Bold))
        self.file_name_label.setStyleSheet(
            f"""
            color: {CyberpunkTheme.GREEN};
            background-color: rgba(0, 0, 0, 0.3);
            border: 1px solid {CyberpunkTheme.GREEN};
            border-radius: 6px;
            padding: 10px 12px;
            """
        )
        self.file_name_label.setWordWrap(True)
        self.file_name_label.setMinimumHeight(40)
        layout.addWidget(self.file_name_label)

        # File details row
        details_layout = QHBoxLayout()

        self.file_size_label = QLabel("")
        self.file_size_label.setFont(QFont("Courier New", 10))
        self.file_size_label.setStyleSheet(f"color: {CyberpunkTheme.YELLOW};")
        details_layout.addWidget(self.file_size_label)

        details_layout.addSpacing(30)

        self.record_count_label = QLabel("")
        self.record_count_label.setFont(QFont("Courier New", 10))
        self.record_count_label.setStyleSheet(f"color: {CyberpunkTheme.CYAN};")
        details_layout.addWidget(self.record_count_label)

        details_layout.addStretch()

        layout.addLayout(details_layout)

        # Action button
        change_btn = QPushButton("CHANGE FILE")
        change_btn.setMinimumHeight(32)
        change_btn.setStyleSheet(
            f"""
            QPushButton {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                           stop:0 rgba(0, 212, 255, 0.1),
                                           stop:1 rgba(0, 212, 255, 0.05));
                color: {CyberpunkTheme.CYAN};
                border: 2px solid {CyberpunkTheme.CYAN};
                border-radius: 6px;
                font-weight: bold;
                font-size: 11px;
            }}
            QPushButton:hover {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                           stop:0 rgba(0, 212, 255, 0.2),
                                           stop:1 rgba(0, 212, 255, 0.1));
                border: 2px solid {CyberpunkTheme.ACCENT_CYAN};
            }}
            """
        )
        self.change_btn = change_btn
        layout.addWidget(change_btn)

        self.setLayout(layout)
        self.setMinimumHeight(160)

    def set_file_info(self, file_path: str, record_count: int):
        """Update file information display"""
        self.file_path = file_path
        file_name = Path(file_path).name
        file_size = os.path.getsize(file_path) / 1024  # KB

        self.file_name_label.setText(f"✓ {file_name}")
        self.file_size_label.setText(f"📊 Size: {file_size:.2f} KB")
        self.record_count_label.setText(f"📋 Records: {record_count}")
        
        self.show()


class DragDropZone(QFrame):
    """Custom drag and drop zone with enhanced styling"""
    file_dropped = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setFrameStyle(QFrame.StyledPanel | QFrame.Raised)
        self.setup_ui()
        self.setStyleSheet(
            f"""
            QFrame {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                                           stop:0 rgba(0, 212, 255, 0.05),
                                           stop:1 rgba(0, 212, 255, 0.02));
                border: 2px dashed {CyberpunkTheme.CYAN};
                border-radius: 12px;
            }}
            """
        )

    def setup_ui(self):
        layout = QVBoxLayout()
        layout.setSpacing(16)
        layout.setContentsMargins(30, 30, 30, 30)

        # Upload icon
        icon_label = QLabel("📤")
        icon_label.setFont(QFont("Arial", 56))
        icon_label.setAlignment(Qt.AlignCenter)
        icon_label.setStyleSheet(f"color: {CyberpunkTheme.CYAN};")

        # Title
        title = QLabel("FILE UPLOAD")
        title_font = QFont("Courier New", 16, QFont.Bold)
        title.setFont(title_font)
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet(
            f"""
            color: {CyberpunkTheme.CYAN};
            font-weight: bold;
            """
        )

        # Description
        desc = QLabel("DRAG & DROP\ncustomer_data.xlsx\nHERE")
        desc_font = QFont("Courier New", 13)
        desc.setFont(desc_font)
        desc.setAlignment(Qt.AlignCenter)
        desc.setStyleSheet(
            f"""
            color: {CyberpunkTheme.CYAN};
            font-weight: bold;
            """
        )

        # Browse button
        browse_btn = QPushButton("BROWSE FILE")
        browse_btn.clicked.connect(self.browse_file)
        browse_btn.setCursor(QCursor(Qt.PointingHandCursor))
        browse_btn.setMinimumHeight(50)
        browse_btn.setMinimumWidth(200)
        browse_btn.setFont(QFont("Courier New", 12, QFont.Bold))
        browse_btn.setStyleSheet(
            f"""
            QPushButton {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                           stop:0 {CyberpunkTheme.PANEL_BG},
                                           stop:1 rgba(0, 212, 255, 0.1));
                color: {CyberpunkTheme.CYAN};
                border: 2px solid {CyberpunkTheme.CYAN};
                border-radius: 8px;
                padding: 15px;
                font-weight: bold;
                font-size: 13px;
            }}
            QPushButton:hover {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                           stop:0 rgba(0, 212, 255, 0.2),
                                           stop:1 rgba(0, 212, 255, 0.1));
                border: 2px solid {CyberpunkTheme.ACCENT_CYAN};
            }}
            """
        )

        layout.addStretch()
        layout.addWidget(icon_label, alignment=Qt.AlignCenter)
        layout.addWidget(title, alignment=Qt.AlignCenter)
        layout.addWidget(desc, alignment=Qt.AlignCenter)
        layout.addWidget(browse_btn, alignment=Qt.AlignCenter)
        layout.addStretch()

        self.setLayout(layout)
        self.setMinimumHeight(350)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.setStyleSheet(
                f"""
                QFrame {{
                    background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                                               stop:0 rgba(0, 255, 65, 0.15),
                                               stop:1 rgba(0, 255, 65, 0.05));
                    border: 3px solid {CyberpunkTheme.GREEN};
                    border-radius: 12px;
                }}
                """
            )
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        self.setStyleSheet(
            f"""
            QFrame {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                                           stop:0 rgba(0, 212, 255, 0.05),
                                           stop:1 rgba(0, 212, 255, 0.02));
                border: 2px dashed {CyberpunkTheme.CYAN};
                border-radius: 12px;
            }}
            """
        )

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            file_path = event.mimeData().urls()[0].toLocalFile()
            if file_path.endswith(('.xlsx', '.xls')):
                self.file_dropped.emit(file_path)
            else:
                QMessageBox.warning(self, "Invalid File", 
                                  "Please drop an Excel file (.xlsx or .xls)")
            self.setStyleSheet(
                f"""
                QFrame {{
                    background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                                               stop:0 rgba(0, 212, 255, 0.05),
                                               stop:1 rgba(0, 212, 255, 0.02));
                    border: 2px dashed {CyberpunkTheme.CYAN};
                    border-radius: 12px;
                }}
                """
            )

    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)"
        )
        if file_path:
            self.file_dropped.emit(file_path)


class CustomerCard(QFrame):
    """Enhanced individual customer result card with prominent name display"""
    print_toggled = pyqtSignal(str, bool)

    def __init__(self, customer_data: dict, index: int = 0, parent=None):
        super().__init__(parent)
        self.customer_data = customer_data
        self.print_enabled = False
        self.index = index
        self.setup_ui()

    def setup_ui(self):
        self.setFrameStyle(QFrame.StyledPanel | QFrame.Raised)
        
        # Determine colors based on status
        status = self.customer_data.get('status', 'PENDING')
        if status == 'COMPLETE':
            border_color = CyberpunkTheme.CYAN
            grad_start = "rgba(0, 212, 255, 0.1)"
            grad_end = "rgba(0, 212, 255, 0.02)"
            status_color = CyberpunkTheme.GREEN
        elif status == 'ERROR':
            border_color = CyberpunkTheme.MAGENTA
            grad_start = "rgba(255, 0, 110, 0.1)"
            grad_end = "rgba(255, 0, 110, 0.02)"
            status_color = CyberpunkTheme.MAGENTA
        else:
            border_color = CyberpunkTheme.DARK_TEXT
            grad_start = "rgba(128, 128, 128, 0.05)"
            grad_end = "rgba(128, 128, 128, 0.02)"
            status_color = CyberpunkTheme.YELLOW

        self.setStyleSheet(
            f"""
            QFrame {{
                border: 2px solid {border_color};
                border-radius: 10px;
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                                           stop:0 {grad_start},
                                           stop:1 {grad_end});
            }}
            """
        )

        layout = QVBoxLayout()
        layout.setSpacing(12)
        layout.setContentsMargins(15, 15, 15, 15)

        # Row number and index at top
        header_layout = QHBoxLayout()
        
        row_label = QLabel(f"#{self.index + 1}")
        row_label.setFont(QFont("Courier New", 10, QFont.Bold))
        row_label.setStyleSheet(f"color: {CyberpunkTheme.YELLOW};")
        row_label.setAlignment(Qt.AlignLeft)
        
        header_layout.addWidget(row_label)
        header_layout.addStretch()
        
        layout.addLayout(header_layout)

        # **PROMINENT CUSTOMER NAME SECTION**
        name_section = QFrame()
        name_section.setStyleSheet(
            f"""
            QFrame {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                           stop:0 rgba(255, 0, 110, 0.2),
                                           stop:1 rgba(255, 0, 110, 0.05));
                border: 2px solid {CyberpunkTheme.MAGENTA};
                border-radius: 8px;
            }}
            """
        )
        name_section.setMinimumHeight(70)
        
        name_layout = QVBoxLayout()
        name_layout.setContentsMargins(12, 10, 12, 10)
        name_layout.setSpacing(4)
        
        # Label
        name_label_text = QLabel("👤 CUSTOMER NAME:")
        name_label_text.setFont(QFont("Courier New", 9, QFont.Bold))
        name_label_text.setStyleSheet(f"color: {CyberpunkTheme.DARK_TEXT};")
        name_layout.addWidget(name_label_text)
        
        # Customer name - LARGE AND BOLD
        customer_name = QLabel(self.customer_data.get('Customer_Name', 'Unknown'))
        customer_name.setFont(QFont("Courier New", 16, QFont.Bold))
        customer_name.setStyleSheet(
            f"""
            color: {CyberpunkTheme.MAGENTA};
            background-color: rgba(0, 0, 0, 0.3);
            border: 1px solid {CyberpunkTheme.MAGENTA};
            border-radius: 5px;
            padding: 8px 10px;
            """
        )
        customer_name.setWordWrap(True)
        customer_name.setAlignment(Qt.AlignCenter)
        name_layout.addWidget(customer_name)
        
        name_section.setLayout(name_layout)
        layout.addWidget(name_section)

        # File info row
        info_layout = QHBoxLayout()
        info_layout.setSpacing(15)

        # File size
        file_size_label = QLabel("📊 File size")
        file_size_label.setStyleSheet(f"color: {CyberpunkTheme.DARK_TEXT}; font-size: 10px;")
        file_size_label.setFont(QFont("Courier New", 9))
        info_layout.addWidget(file_size_label)

        info_layout.addStretch()

        # Status
        status_label = QLabel(f"[{status}]")
        status_label.setFont(QFont("Courier New", 11, QFont.Bold))
        status_label.setStyleSheet(
            f"""
            color: {status_color};
            background-color: rgba(0, 0, 0, 0.2);
            border: 1px solid {status_color};
            border-radius: 4px;
            padding: 4px 8px;
            """
        )
        info_layout.addWidget(status_label)

        layout.addLayout(info_layout)

        # Print button - Full width
        print_btn = QPushButton("🖨️  PRINT NOW")
        print_btn.setCheckable(True)
        print_btn.setMinimumHeight(36)
        print_btn.setFont(QFont("Courier New", 11, QFont.Bold))
        print_btn.setStyleSheet(
            f"""
            QPushButton {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                           stop:0 {CyberpunkTheme.PANEL_BG},
                                           stop:1 rgba(0, 212, 255, 0.05));
                color: {CyberpunkTheme.CYAN};
                border: 2px solid {CyberpunkTheme.CYAN};
                border-radius: 6px;
                padding: 8px 12px;
                font-weight: bold;
                font-size: 11px;
            }}
            QPushButton:hover {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                           stop:0 rgba(0, 212, 255, 0.15),
                                           stop:1 rgba(0, 212, 255, 0.05));
                border: 2px solid {CyberpunkTheme.ACCENT_CYAN};
            }}
            QPushButton:checked {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                           stop:0 {CyberpunkTheme.CYAN},
                                           stop:1 rgba(0, 212, 255, 0.3));
                color: {CyberpunkTheme.BACKGROUND};
                border: 2px solid {CyberpunkTheme.CYAN};
            }}
            """
        )
        print_btn.toggled.connect(self.on_print_toggled)
        
        layout.addWidget(print_btn)
        
        self.setLayout(layout)
        self.setMinimumHeight(200)

    def on_print_toggled(self, checked):
        self.print_enabled = checked
        name = self.customer_data.get('Customer_Name', '')
        self.print_toggled.emit(name, checked)


class LeftSidebar(QFrame):
    """Left navigation sidebar with gradient"""
    nav_clicked = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()

    def setup_ui(self):
        self.setFrameStyle(QFrame.NoFrame)
        self.setStyleSheet(
            f"""
            background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                                       stop:0 {CyberpunkTheme.DARK_BG},
                                       stop:1 rgba(15, 21, 53, 0.8));
            border-right: 2px solid {CyberpunkTheme.CYAN};
            """
        )
        self.setMaximumWidth(100)

        layout = QVBoxLayout()
        layout.setSpacing(20)
        layout.setContentsMargins(10, 20, 10, 20)

        # Navigation buttons
        nav_items = [
            ("📤", "UPLOAD", "upload"),
            ("🖨️", "PRINTER", "printer"),
            ("⚙️", "SETTINGS", "settings"),
            ("🕐", "HISTORY", "history"),
            ("❓", "HELP", "help"),
        ]

        for icon, label, action in nav_items:
            btn = QPushButton(f"{icon}\n{label}")
            btn.setFont(QFont("Arial", 9, QFont.Bold))
            btn.setMinimumHeight(80)
            btn.setCursor(QCursor(Qt.PointingHandCursor))
            btn.setStyleSheet(
                f"""
                QPushButton {{
                    background-color: transparent;
                    color: {CyberpunkTheme.DARK_TEXT};
                    border: none;
                    padding: 8px;
                    text-align: center;
                }}
                QPushButton:hover {{
                    color: {CyberpunkTheme.CYAN};
                }}
                """
            )
            btn.clicked.connect(lambda checked, a=action: self.nav_clicked.emit(a))
            layout.addWidget(btn)

        layout.addStretch()
        self.setLayout(layout)


class TopBar(QFrame):
    """Top status and info bar with gradient"""
    settings_clicked = pyqtSignal()
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()

    def setup_ui(self):
        self.setFrameStyle(QFrame.NoFrame)
        self.setStyleSheet(
            f"""
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                       stop:0 rgba(10, 14, 39, 0.95),
                                       stop:1 rgba(26, 31, 58, 0.5));
            border-bottom: 2px solid {CyberpunkTheme.CYAN};
            """
        )
        self.setMaximumHeight(60)

        layout = QHBoxLayout()
        layout.setContentsMargins(20, 10, 20, 10)
        layout.setSpacing(15)

        # Title
        title = QLabel("BULK BANK CONFIRMATION ENGINE")
        title.setFont(QFont("Courier New", 16, QFont.Bold))
        title.setStyleSheet(
            f"""
            color: {CyberpunkTheme.CYAN};
            font-weight: bold;
            """
        )

        # Status indicator
        status_layout = QHBoxLayout()
        status_dot = QLabel("●")
        status_dot.setStyleSheet(
            f"""
            color: {CyberpunkTheme.GREEN};
            font-size: 16px;
            """
        )
        status_label = QLabel("Ready")
        status_label.setStyleSheet(
            f"""
            color: {CyberpunkTheme.GREEN};
            font-weight: bold;
            """
        )
        status_layout.addWidget(status_dot)
        status_layout.addWidget(status_label)

        # Settings button
        settings_btn = QPushButton("⚙️")
        settings_btn.setMaximumWidth(40)
        settings_btn.setStyleSheet(
            f"""
            QPushButton {{
                background-color: transparent;
                color: {CyberpunkTheme.DARK_TEXT};
                border: 2px solid {CyberpunkTheme.DARK_TEXT};
                border-radius: 6px;
                padding: 5px;
            }}
            QPushButton:hover {{
                color: {CyberpunkTheme.CYAN};
                border: 2px solid {CyberpunkTheme.CYAN};
            }}
            """
        )
        settings_btn.clicked.connect(self.settings_clicked.emit)

        # User label
        user_label = QLabel("👤 KKP_TECHNOLOGY ▼")
        user_label.setStyleSheet(
            f"""
            color: {CyberpunkTheme.DARK_TEXT};
            font-size: 10px;
            font-weight: bold;
            """
        )

        layout.addWidget(title)
        layout.addStretch()
        layout.addLayout(status_layout)
        layout.addWidget(settings_btn)
        layout.addWidget(user_label)

        self.setLayout(layout)


class GenerateButton(QPushButton):
    """Large circular generate button with glow"""
    def __init__(self, parent=None):
        super().__init__("GENERATE\nFILES", parent)
        self.setFixedSize(QSize(200, 200))
        self.setStyleSheet(
            f"""
            QPushButton {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                                           stop:0 rgba(0, 255, 212, 0.3),
                                           stop:1 {CyberpunkTheme.PANEL_BG});
                color: {CyberpunkTheme.CYAN};
                border: 3px solid {CyberpunkTheme.CYAN};
                border-radius: 100px;
                font-weight: bold;
                font-size: 14px;
                font-family: 'Courier New';
            }}
            QPushButton:hover {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                                           stop:0 rgba(0, 255, 212, 0.5),
                                           stop:1 rgba(0, 255, 212, 0.1));
                border: 3px solid {CyberpunkTheme.ACCENT_CYAN};
            }}
            QPushButton:pressed {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                                           stop:0 {CyberpunkTheme.CYAN},
                                           stop:1 {CyberpunkTheme.PANEL_BG});
                color: {CyberpunkTheme.BACKGROUND};
            }}
            """
        )
        self.setCursor(QCursor(Qt.PointingHandCursor))


class BulkBankConfirmationGUI(QMainWindow):
    """Main application window"""
    
    def __init__(self):
        super().__init__()
        self.excel_file: Optional[str] = None
        self.template_file: Optional[str] = None
        self.output_folder: Optional[str] = None
        self.customer_data: List[dict] = []
        self.signals = ProcessingSignals()
        self.setup_paths()
        self.init_ui()
        self.connect_signals()

    def setup_paths(self):
        """Setup file paths relative to script location"""
        script_dir = Path(__file__).parent
        self.template_file = script_dir / 'Bank_Confirmation_Template.docx'
        self.output_folder = script_dir / 'Generated_Letters'
        self.output_folder.mkdir(exist_ok=True)

    def init_ui(self):
        """Initialize the user interface"""
        self.setWindowTitle("Bulk Bank Confirmation Engine")
        self.setGeometry(100, 100, 1400, 900)
        self.setStyleSheet(CyberpunkTheme.get_stylesheet())

        # Main container
        main_widget = QWidget()
        main_layout = QHBoxLayout()
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # Left sidebar
        self.sidebar = LeftSidebar()
        main_layout.addWidget(self.sidebar)

        # Central content
        central_layout = QVBoxLayout()
        central_layout.setContentsMargins(0, 0, 0, 0)
        central_layout.setSpacing(0)

        # Top bar
        self.topbar = TopBar()
        central_layout.addWidget(self.topbar)

        # Main content area
        content_frame = QFrame()
        content_frame.setFrameStyle(QFrame.NoFrame)
        content_layout = QHBoxLayout()
        content_layout.setContentsMargins(20, 20, 20, 20)
        content_layout.setSpacing(20)

        # Left panel: Drag and drop + File info
        left_panel_layout = QVBoxLayout()
        left_panel_layout.setSpacing(15)

        # File info frame (initially hidden)
        self.file_info_frame = FileInfoFrame()
        self.file_info_frame.change_btn.clicked.connect(self.on_change_file)
        left_panel_layout.addWidget(self.file_info_frame)

        # Drag and drop zone
        self.drag_drop_zone = DragDropZone()
        self.drag_drop_zone.file_dropped.connect(self.on_file_dropped)
        left_panel_layout.addWidget(self.drag_drop_zone)

        content_layout.addLayout(left_panel_layout)

        # Right panel: Results
        right_panel_layout = QVBoxLayout()
        right_panel_layout.setSpacing(15)

        # Customer cards area
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setStyleSheet(
            f"""
            QScrollArea {{
                background-color: {CyberpunkTheme.BACKGROUND};
                border: none;
            }}
            """
        )

        self.cards_container = QWidget()
        self.cards_layout = QVBoxLayout()
        self.cards_layout.setSpacing(12)
        self.cards_container.setLayout(self.cards_layout)
        scroll_area.setWidget(self.cards_container)

        right_panel_layout.addWidget(scroll_area)

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setMaximumHeight(25)
        self.progress_bar.setVisible(False)
        self.progress_label = QLabel("GENERATING FILES - 0%")
        self.progress_label.setStyleSheet(
            f"""
            color: {CyberpunkTheme.YELLOW};
            font-weight: bold;
            font-size: 11px;
            """
        )
        self.progress_label.setVisible(False)

        right_panel_layout.addWidget(self.progress_label)
        right_panel_layout.addWidget(self.progress_bar)

        # Generate buttons area
        buttons_layout = QHBoxLayout()
        buttons_layout.addStretch()

        self.print_all_btn = QPushButton("GENERATE &\nPRINT ALL")
        self.print_all_btn.setMinimumHeight(80)
        self.print_all_btn.setMinimumWidth(180)
        self.print_all_btn.setStyleSheet(
            f"""
            QPushButton {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                           stop:0 {CyberpunkTheme.PANEL_BG},
                                           stop:1 rgba(255, 0, 110, 0.1));
                color: {CyberpunkTheme.MAGENTA};
                border: 2px solid {CyberpunkTheme.MAGENTA};
                border-radius: 8px;
                font-weight: bold;
                font-size: 11px;
                font-family: 'Courier New';
            }}
            QPushButton:hover {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                           stop:0 rgba(255, 0, 110, 0.2),
                                           stop:1 rgba(255, 0, 110, 0.1));
                border: 2px solid {CyberpunkTheme.MAGENTA};
            }}
            QPushButton:pressed {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                           stop:0 {CyberpunkTheme.MAGENTA},
                                           stop:1 rgba(255, 0, 110, 0.3));
                color: {CyberpunkTheme.BACKGROUND};
            }}
            """
        )
        self.print_all_btn.clicked.connect(self.on_generate_and_print_all)
        self.print_all_btn.setEnabled(False)
        buttons_layout.addWidget(self.print_all_btn)

        self.generate_btn = GenerateButton()
        self.generate_btn.clicked.connect(self.on_generate_files)
        self.generate_btn.setEnabled(False)
        buttons_layout.addWidget(self.generate_btn)

        buttons_layout.addStretch()

        right_panel_layout.addLayout(buttons_layout)

        content_layout.addLayout(right_panel_layout, 1)

        content_frame.setLayout(content_layout)
        central_layout.addWidget(content_frame, 1)

        main_layout.addLayout(central_layout, 1)
        main_widget.setLayout(main_layout)

        self.setCentralWidget(main_widget)

    def connect_signals(self):
        """Connect processing signals"""
        self.signals.progress_updated.connect(self.update_progress)
        self.signals.file_processed.connect(self.add_customer_card)
        self.signals.processing_complete.connect(self.on_processing_complete)
        self.signals.error_occurred.connect(self.show_error)

    def on_file_dropped(self, file_path: str):
        """Handle Excel file dropped or selected"""
        self.excel_file = file_path
        self.load_customer_data()

    def on_change_file(self):
        """Open file browser to change file"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)"
        )
        if file_path:
            self.on_file_dropped(file_path)

    def load_customer_data(self):
        """Load and display customer data from Excel"""
        if not self.excel_file or not Path(self.excel_file).exists():
            self.show_error("File not found")
            return

        try:
            df = pd.read_excel(self.excel_file)
            self.customer_data = df.to_dict('records')

            # Update file info display
            self.file_info_frame.set_file_info(self.excel_file, len(self.customer_data))

            # Clear existing cards
            while self.cards_layout.count():
                self.cards_layout.takeAt(0).widget().deleteLater()

            # Add customer cards with index
            for idx, row in df.iterrows():
                card_data = {
                    'Customer_Name': row.get('Customer_Name', f'Customer {idx+1}'),
                    'status': 'PENDING'
                }
                card = CustomerCard(card_data, index=idx)
                card.print_toggled.connect(self.on_print_toggled)
                self.cards_layout.addWidget(card)

            self.cards_layout.addStretch()

            # Enable generate buttons
            self.generate_btn.setEnabled(True)
            self.print_all_btn.setEnabled(True)

            self.show_message(f"✓ Loaded {len(self.customer_data)} customer records")

        except Exception as e:
            self.show_error(f"Error loading Excel file: {str(e)}")

    def add_customer_card(self, customer_data: dict):
        """Add or update customer card with results"""
        # Remove stretch and update card
        if self.cards_layout.count() > 0:
            item = self.cards_layout.takeAt(self.cards_layout.count() - 1)
            if item and item.widget():
                item.widget().deleteLater()

        # Find current index
        current_count = self.cards_layout.count()
        
        card = CustomerCard(customer_data, index=current_count)
        card.print_toggled.connect(self.on_print_toggled)
        self.cards_layout.addWidget(card)
        self.cards_layout.addStretch()

    def on_print_toggled(self, customer_name: str, enabled: bool):
        """Handle print toggle for a customer"""
        pass  # Implementation for print selection

    def on_generate_files(self):
        """Generate files only (no printing)"""
        self.generate_files(print_files=False)

    def on_generate_and_print_all(self):
        """Generate and print all files"""
        self.generate_files(print_files=True)

    def generate_files(self, print_files: bool = False):
        """Generate files in background thread"""
        if not self.customer_data or not self.template_file.exists():
            self.show_error("Please load a valid Excel file first")
            return

        # Disable buttons during processing
        self.generate_btn.setEnabled(False)
        self.print_all_btn.setEnabled(False)

        # Show progress
        self.progress_bar.setVisible(True)
        self.progress_label.setVisible(True)
        self.progress_bar.setValue(0)

        # Start processing in background
        thread = Thread(
            target=self._process_customers,
            args=(print_files,),
            daemon=True
        )
        thread.start()

    def _process_customers(self, print_files: bool):
        """Background thread: Process customer data"""
        success_count = 0
        total = len(self.customer_data)

        for idx, row in enumerate(self.customer_data):
            try:
                # Load template
                doc = DocxTemplate(str(self.template_file))

                # Format dates
                date_columns = ['Date', 'Start_Date', 'Balance_Date']
                for col in date_columns:
                    if col in row and pd.notna(row[col]):
                        try:
                            row[col] = pd.to_datetime(row[col]).strftime('%Y-%m-%d')
                        except:
                            pass

                # Build context
                context = {
                    'Reference_Number': row.get('Reference_Number', ''),
                    'Date': row.get('Date', ''),
                    'Customer_Name': row.get('Customer_Name', ''),
                    'Account_Type': row.get('Account_Type', ''),
                    'Bank_Name': row.get('Bank_Name', ''),
                    'Customer_ID': row.get('Customer_ID', ''),
                    'Start_Date': row.get('Start_Date', ''),
                    'Balance_Date': row.get('Balance_Date', ''),
                    'Balance_Local': f"{float(row.get('Balance_Local', 0)):,.2f}",
                    'Balance_Foreign': f"{float(row.get('Balance_Foreign', 0)):,.2f}",
                    'Conversion_Rate': row.get('Conversion_Rate', ''),
                    'Customer_Address': row.get('Customer_Address', '')
                }

                # Render document
                doc.render(context)

                # Save file
                safe_name = str(row.get('Customer_Name', f'Customer_{idx+1}')).replace(" ", "_").replace("/", "_")
                output_file = self.output_folder / f"Bank_Letter_{safe_name}.docx"
                doc.save(str(output_file))

                success_count += 1

                # Print if requested
                if print_files:
                    self._print_file(str(output_file))

                # Emit signals
                self.signals.file_processed.emit({
                    'Customer_Name': row.get('Customer_Name', 'Unknown'),
                    'status': 'COMPLETE'
                })

                # Update progress
                progress = int((idx + 1) / total * 100)
                self.signals.progress_updated.emit(progress)

            except Exception as e:
                self.signals.file_processed.emit({
                    'Customer_Name': row.get('Customer_Name', 'Unknown'),
                    'status': 'ERROR'
                })
                self.signals.error_occurred.emit(
                    f"Error processing {row.get('Customer_Name', 'Unknown')}: {str(e)}"
                )

        self.signals.processing_complete.emit(success_count, total)

    def _print_file(self, filepath: str):
        """Print a file"""
        try:
            system = platform.system()
            if system == "Windows":
                os.startfile(filepath, "print")
            elif system == "Darwin":  # macOS
                subprocess.run(['lpr', filepath], check=True)
            else:  # Linux
                subprocess.run(['lp', filepath], check=True)
        except Exception as e:
            self.signals.error_occurred.emit(f"Print error: {str(e)}")

    def update_progress(self, value: int):
        """Update progress bar"""
        self.progress_bar.setValue(value)
        self.progress_label.setText(f"GENERATING FILES - {value}%")

    def on_processing_complete(self, success_count: int, total: int):
        """Processing complete"""
        self.progress_bar.setVisible(False)
        self.progress_label.setVisible(False)
        self.generate_btn.setEnabled(True)
        self.print_all_btn.setEnabled(True)
        self.show_message(f"✓ Successfully generated {success_count}/{total} files")

    def show_error(self, message: str):
        """Show error dialog"""
        QMessageBox.critical(self, "Error", message)

    def show_message(self, message: str):
        """Show info dialog"""
        QMessageBox.information(self, "Info", message)


def main():
    app = QApplication(sys.argv)
    window = BulkBankConfirmationGUI()
    window.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()