import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QLabel, 
                            QVBoxLayout, QWidget, QFileDialog, QProgressBar, QListWidget, QTextEdit, QHBoxLayout, QStyle, QListWidgetItem, QFrame, QSplitter)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QIcon, QPalette, QColor
import os
import datetime
import xml.etree.ElementTree as ET
import pandas as pd

class DropZone(QWidget):
    filesDropped = pyqtSignal(list)
    
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout()
        self.setLayout(layout)
        
        self.label = QLabel("XML dosyalarını buraya sürükleyin\nveya dosya seçmek için tıklayın")
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setStyleSheet("""
            QWidget {
                border: 2px dashed #aaa;
                border-radius: 5px;
                padding: 20px;
                background: #f8f8f8;
            }
            QWidget:hover {
                border-color: #666;
                background: #f0f0f0;
            }
        """)
        layout.addWidget(self.label)
        
    def mousePressEvent(self, event):
        files, _ = QFileDialog.getOpenFileNames(self, "XML Dosyaları Seç", "", "XML Dosyaları (*.xml)")
        if files:
            self.filesDropped.emit(files)
            
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()
            
    def dropEvent(self, event):
        files = [url.toLocalFile() for url in event.mimeData().urls() if url.toLocalFile().endswith('.xml')]
        if files:
            self.filesDropped.emit(files)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("XML Fatura Dönüştürücü")
        self.setMinimumSize(800, 600)
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f0f0f0;
            }
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                padding: 8px 15px;
                border-radius: 4px;
                font-weight: bold;
                min-width: 100px;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
            QPushButton:pressed {
                background-color: #0D47A1;
            }
            QLabel {
                color: #333;
                font-weight: bold;
            }
            QListWidget {
                background-color: white;
                border: 1px solid #ddd;
                border-radius: 4px;
                padding: 5px;
            }
            QTextEdit {
                background-color: white;
                border: 1px solid #ddd;
                border-radius: 4px;
                padding: 5px;
            }
            QProgressBar {
                border: 1px solid #ddd;
                border-radius: 4px;
                text-align: center;
            }
            QProgressBar::chunk {
                background-color: #2196F3;
            }
        """)
        
        # Ana widget ve layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setSpacing(10)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Üst bölüm - Butonlar
        button_layout = QHBoxLayout()
        
        select_files_btn = QPushButton("Dosya Seç")
        select_files_btn.setIcon(self.style().standardIcon(QStyle.SP_FileIcon))
        select_files_btn.clicked.connect(self.select_files)
        button_layout.addWidget(select_files_btn)
        
        select_folder_btn = QPushButton("Klasör Seç")
        select_folder_btn.setIcon(self.style().standardIcon(QStyle.SP_DirIcon))
        select_folder_btn.clicked.connect(self.select_folder)
        button_layout.addWidget(select_folder_btn)
        
        button_layout.addStretch()
        
        view_report_btn = QPushButton("Raporu Görüntüle")
        view_report_btn.setIcon(self.style().standardIcon(QStyle.SP_FileDialogInfoView))
        view_report_btn.clicked.connect(self.view_report)
        button_layout.addWidget(view_report_btn)
        
        layout.addLayout(button_layout)
        
        # Ayırıcı çizgi
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        layout.addWidget(line)
        
        # Splitter oluştur
        splitter = QSplitter(Qt.Vertical)
        
        # İşlem durumu widget'ı
        status_widget = QWidget()
        status_layout = QVBoxLayout(status_widget)
        status_layout.setContentsMargins(0, 0, 0, 0)
        
        list_label = QLabel("İşlem Durumu:")
        status_layout.addWidget(list_label)
        
        self.file_list = QListWidget()
        status_layout.addWidget(self.file_list)
        
        splitter.addWidget(status_widget)
        
        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFormat("%p% Tamamlandı")
        layout.addWidget(self.progress_bar)
        
        # Log widget'ı
        log_widget = QWidget()
        log_layout = QVBoxLayout(log_widget)
        log_layout.setContentsMargins(0, 0, 0, 0)
        
        log_label = QLabel("İşlem Detayları:")
        log_layout.addWidget(log_label)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        log_layout.addWidget(self.log_text)
        
        splitter.addWidget(log_widget)
        
        # Splitter'ı ana layout'a ekle
        layout.addWidget(splitter)
        
        # Başlangıç boyutlarını ayarla
        splitter.setSizes([300, 300])  # Her iki bölüm eşit boyutta
        
        self.worker = None
        self.error_report = []

    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "XML Dosyaları Seç",
            "",
            "XML Dosyaları (*.xml)"
        )
        if files:
            self.process_files(files)

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Klasör Seç")
        if folder:
            files = [os.path.join(folder, f) for f in os.listdir(folder) if f.endswith('.xml')]
            if files:
                self.process_files(files)
            else:
                self.log_message("Seçilen klasörde XML dosyası bulunamadı!")

    def process_files(self, files):
        self.file_list.clear()
        self.error_report.clear()
        self.progress_bar.setValue(0)
        
        for file in files:
            item = QListWidgetItem(os.path.basename(file))
            item.setIcon(self.style().standardIcon(QStyle.SP_FileIcon))
            self.file_list.addItem(item)
        
        self.worker = XMLConverter(files)
        self.worker.progress.connect(self.update_progress)
        self.worker.file_status.connect(self.update_file_status)
        self.worker.error.connect(self.handle_error)
        self.worker.finished.connect(self.process_completed)
        self.worker.start()

    def update_progress(self, value):
        self.progress_bar.setValue(value)

    def update_file_status(self, file_name, status, is_error=False):
        items = self.file_list.findItems(file_name, Qt.MatchFlag.MatchExactly)
        if items:
            item = items[0]
            item.setText(f"{file_name} - {status}")
            if is_error:
                item.setForeground(Qt.GlobalColor.red)
                self.error_report.append((file_name, status))
            else:
                item.setForeground(Qt.GlobalColor.green)
        self.log_message(f"{file_name}: {status}")

    def handle_error(self, error_msg):
        self.log_message(f"HATA: {error_msg}")

    def log_message(self, message):
        self.log_text.append(f"{datetime.datetime.now().strftime('%H:%M:%S')}: {message}")

    def process_completed(self, summary):
        self.log_message("\n=== İşlem Tamamlandı ===")
        self.log_message(summary)
        if self.error_report:
            self.create_error_report()

    def create_error_report(self):
        report_path = os.path.join(os.getcwd(), "hata_raporu.txt")
        with open(report_path, "w", encoding="utf-8") as f:
            f.write("=== XML Fatura Dönüştürme Hata Raporu ===\n")
            f.write(f"Tarih: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
            for file_name, error in self.error_report:
                f.write(f"Dosya: {file_name}\n")
                f.write(f"Hata: {error}\n\n")
        self.log_message(f"Hata raporu oluşturuldu: {report_path}")

    def view_report(self):
        if self.error_report:
            report_path = os.path.join(os.getcwd(), "hata_raporu.txt")
            if os.path.exists(report_path):
                os.startfile(report_path)
        else:
            self.log_message("Henüz hata raporu oluşturulmadı.")

class XMLConverter(QThread):
    progress = pyqtSignal(int)
    file_status = pyqtSignal(str, str, bool)
    error = pyqtSignal(str)
    finished = pyqtSignal(str)

    def __init__(self, files):
        super().__init__()
        self.files = files
        self.namespaces = {
            'cac': 'urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2',
            'cbc': 'urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2'
        }

    def run(self):
        total_files = len(self.files)
        processed_files = 0
        error_count = 0
        
        for i, file_path in enumerate(self.files):
            try:
                file_name = os.path.basename(file_path)
                self.file_status.emit(file_name, "İşleniyor...", False)
                
                # XML'i işle ve Excel'e dönüştür
                tree = ET.parse(file_path)
                root = tree.getroot()
                
                # Fatura satırlarını işle
                invoice_lines = root.findall('.//cac:InvoiceLine', self.namespaces)
                if not invoice_lines:
                    raise Exception("Fatura satırı bulunamadı")
                
                data = self.process_invoice_lines(invoice_lines)
                
                # Excel dosyasını oluştur
                excel_path = os.path.splitext(file_path)[0] + '.xlsx'
                df = pd.DataFrame(data)
                df.to_excel(excel_path, index=False)
                
                self.file_status.emit(file_name, "Tamamlandı", False)
                processed_files += 1
                
            except Exception as e:
                error_msg = str(e)
                self.file_status.emit(file_name, f"Hata: {error_msg}", True)
                error_count += 1
            
            self.progress.emit(int((i + 1) / total_files * 100))
        
        summary = (f"Toplam {total_files} dosya işlendi.\n"
                  f"Başarılı: {processed_files}\n"
                  f"Hatalı: {error_count}")
        self.finished.emit(summary)

    def process_invoice_lines(self, invoice_lines):
        """Fatura satırlarını işler"""
        # Faturadaki sütun sıralaması ve XML yolları
        field_mappings = {
            'Sıra No': './/cbc:ID',
            'Mal Hizmet': './/cac:Item/cbc:Name',
            'Miktar': './/cbc:InvoicedQuantity',
            'Birim Fiyat': './/cac:Price/cbc:PriceAmount',
            'İskonto Oranı': './/cac:AllowanceCharge/cbc:MultiplierFactorNumeric',
            'İskonto Tutarı': './/cac:AllowanceCharge/cbc:Amount',
            'KDV Oranı': './/cac:TaxTotal/cac:TaxSubtotal/cbc:Percent',
            'KDV Tutarı': './/cac:TaxTotal/cac:TaxSubtotal/cbc:TaxAmount',
            'Diğer Vergiler': './/cac:WithholdingTaxTotal/cbc:TaxAmount',
            'Mal Hizmet Tutarı': './/cbc:LineExtensionAmount',
            'Teslim Şartı': './/cac:Delivery/cac:DeliveryTerms/cbc:ID',
            'Eşya Kap Cinsi': './/cac:Delivery/cac:Shipment/cac:TransportHandlingUnit/cbc:PackagingTypeCode',
            'Kap No': './/cac:Delivery/cac:Shipment/cac:TransportHandlingUnit/cbc:ID',
            'Kap Adet': './/cac:Delivery/cac:Shipment/cac:TransportHandlingUnit/cbc:Quantity',
            'Teslim/Bedel Ödeme Yeri': './/cac:Delivery/cac:DeliveryAddress/cbc:BuildingName',
            'Gönderilme Şekli': './/cac:Delivery/cac:Shipment/cac:ShipmentStage/cbc:TransportModeCode',
            'GTİP': './/cac:Delivery/cac:Shipment/cac:GoodsItem/cbc:RequiredCustomsID'
        }

        data = []
        
        for line in invoice_lines:
            row_data = {}
            for field_name, xpath in field_mappings.items():
                try:
                    element = line.find(xpath, self.namespaces)
                    value = element.text.strip() if element is not None and element.text else None
                    
                    # Sayısal değerleri dönüştür
                    if value:
                        try:
                            if '.' in str(value):
                                value = float(value)
                            elif str(value).isdigit():
                                value = int(value)
                        except:
                            pass
                            
                    row_data[field_name] = value
                    
                except Exception as e:
                    print(f"  ! {field_name} hatası: {str(e)}")
                    row_data[field_name] = None
                    
            data.append(row_data)
        
        return data

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec()) 