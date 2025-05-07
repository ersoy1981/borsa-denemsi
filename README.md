# borsa-denemesi
bist 100 deneme yazılımı
import pandas as pd
import numpy as np
import yfinance as yf
import os
from datetime import datetime, timedelta
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton, 
                             QMessageBox, QTableWidget, QTableWidgetItem, QHBoxLayout, QHeaderView, 
                             QProgressBar, QFileDialog)
from PyQt5.QtCore import QObject, pyqtSignal, QThread, Qt

# Thread ile GUI arasında iletişim için sinyal sınıfı
class WorkerSignals(QObject):
    finished = pyqtSignal(str, str, float)  # symbol, decision, price
    error = pyqtSignal(str)
    progress = pyqtSignal(int)

# Hisse analizi yapan thread
class StockAnalysisWorker(QThread):
    def __init__(self, stock_symbol):
        super().__init__()
        self.stock_symbol = stock_symbol
        self.signals = WorkerSignals()
    
    def run(self):
        try:
            decision, current_price = self.moving_average_strategy(self.stock_symbol)
            if decision is None:
                self.signals.error.emit(f'{self.stock_symbol} için veri alınırken bir hata oluştu.')
            else:
                self.signals.finished.emit(self.stock_symbol, decision, current_price)
        except Exception as e:
            self.signals.error.emit(f'Hata: {str(e)}')
    
    def get_stock_data(self, stock_symbol):
        try:
            # Bugünün tarihini al
            end_date = datetime.now()
            # 3 ay önceki tarihi al
            start_date = end_date - timedelta(days=90)
            
            # Türkiye hisseleri için .IS ekleyerek Yahoo Finance'den veri alıyoruz
            ticker = f"{stock_symbol}.IS"
            stock_data = yf.download(ticker, start=start_date.strftime('%Y-%m-%d'), end=end_date.strftime('%Y-%m-%d'))
            
            if stock_data.empty:
                raise ValueError(f"{stock_symbol} için veri bulunamadı.")
            return stock_data
        except Exception as e:
            print(f"Hata ({stock_symbol}): {str(e)}")
            return None
    
    def moving_average_strategy(self, stock_symbol):
        # Hisse verilerini al
        data = self.get_stock_data(stock_symbol)
        if data is None:
            return None, None
        
        # Eğer yeterli veri yoksa, MA20 ve MA50 dizileri boş olabilir.
        if len(data) < 50:
            return "Yeterli veri yok", None
        
        # Son kapanış fiyatını al
        current_price = data['Close'].iloc[-1]
        
        # Burada current_price'ın float olduğundan emin oluyoruz
        if isinstance(current_price, pd.Series):
            current_price = float(current_price.iloc[-1])
        else:
            current_price = float(current_price)
        
        # 20 gün ve 50 gün hareketli ortalama hesaplama
        data['MA20'] = data['Close'].rolling(window=20).mean()
        data['MA50'] = data['Close'].rolling(window=50).mean()
        
        # Al/Sat sinyali oluşturma (kısa MA uzun MA'nın üzerine çıkarsa 'Al', aksi takdirde 'Sat')
        if data['MA20'].iloc[-1] > data['MA50'].iloc[-1]:
            return 'AL', current_price
        else:
            return 'SAT', current_price

# Toplu hisse analizi için thread
class BatchAnalysisWorker(QThread):
    def __init__(self, symbols):
        super().__init__()
        self.symbols = symbols
        self.signals = WorkerSignals()
        self.results = {}  # {symbol: (decision, price)}
        
    def run(self):
        total = len(self.symbols)
        for i, symbol in enumerate(self.symbols):
            try:
                worker = StockAnalysisWorker(symbol)
                decision, price = worker.moving_average_strategy(symbol)
                
                if decision is not None and price is not None:
                    self.results[symbol] = (decision, price)
                
                # İlerleme durumunu bildir
                self.signals.progress.emit(int((i + 1) / total * 100))
            except Exception as e:
                print(f"Batch Analysis Error ({symbol}): {str(e)}")
        
        self.signals.finished.emit("batch", "completed", 0.0)

# GUI uygulaması
class StockApp(QWidget):
    def __init__(self):
        super().__init__()
        
        # GUI bileşenlerini oluştur
        self.setWindowTitle('Borsa Al/Sat Tavsiyesi - BIST 100')
        self.setGeometry(100, 100, 800, 600)
        
        self.setup_ui()
        
        # Worker thread referansını sakla
        self.worker = None
        self.batch_worker = None
        
        # Porföy takibi için verileri saklama
        self.portfolio_data = {}  # {date: {symbol: {'price': price, 'decision': decision, 'shares': shares, 'cost': cost}}}
        
        # Excel dosya yolu
        self.excel_file_path = None
    
    def setup_ui(self):
        main_layout = QVBoxLayout()
        
        # Hisse senedi girişi bölümü
        input_layout = QHBoxLayout()
        
        self.label = QLabel("Hisse Senedi Kodu (örneğin: 'AKBNK'):")
        input_layout.addWidget(self.label)
        
        self.entry = QLineEdit(self)
        input_layout.addWidget(self.entry)
        
        # Tekli analiz butonu
        self.button = QPushButton("Al/Sat Tavsiyesi", self)
        self.button.clicked.connect(self.on_button_click)
        input_layout.addWidget(self.button)
        
        main_layout.addLayout(input_layout)
        
        # Butonlar için yatay düzen
        buttons_layout = QHBoxLayout()
        
        # Toplu analiz butonu
        self.batch_button = QPushButton("Tüm BIST 100 Hisselerini Analiz Et", self)
        self.batch_button.clicked.connect(self.on_batch_button_click)
        buttons_layout.addWidget(self.batch_button)
        
        # Excel dosyası seçme butonu
        self.select_excel_button = QPushButton("Excel Dosyası Seç/Oluştur", self)
        self.select_excel_button.clicked.connect(self.select_excel_file)
        buttons_layout.addWidget(self.select_excel_button)
        
        # Excel kaydetme butonu
        self.save_button = QPushButton("Sonuçları Excel'e Kaydet", self)
        self.save_button.clicked.connect(self.save_to_excel)
        self.save_button.setEnabled(False)  # Başlangıçta devre dışı
        buttons_layout.addWidget(self.save_button)
        
        main_layout.addLayout(buttons_layout)
        
        # İlerleme çubuğu
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)
        
        # Sonuç tablosu
        self.table = QTableWidget(0, 6)  # 6 sütun: Hisse Kodu, Fiyat, Öneri, Günlük Alım Adedi, Günlük Yatırım, Kar/Zarar
        self.table.setHorizontalHeaderLabels(["Hisse Kodu", "Güncel Fiyat", "Öneri", "Günlük Alım Adedi", "Günlük Yatırım", "Kar/Zarar"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        main_layout.addWidget(self.table)
        
        self.setLayout(main_layout)
    
    def on_button_click(self):
        stock_symbol = self.entry.text().strip().upper()
        if stock_symbol == '':
            QMessageBox.warning(self, "Hata", "Lütfen geçerli bir hisse senedi kodu girin.")
            return
        
        # Daha önce bir işlem yapılıyorsa butonu devre dışı bırak
        self.button.setEnabled(False)
        self.update_status_bar(f"'{stock_symbol}' için veri alınıyor ve analiz ediliyor...")
        
        # Worker thread'i oluştur ve başlat
        self.worker = StockAnalysisWorker(stock_symbol)
        self.worker.signals.finished.connect(self.update_result)
        self.worker.signals.error.connect(self.show_error)
        self.worker.finished.connect(self.enable_button)
        self.worker.start()
    
    def on_batch_button_click(self):
        self.batch_button.setEnabled(False)
        self.button.setEnabled(False)
        self.table.setRowCount(0)
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(True)
        self.update_status_bar("Tüm BIST 100 hisseleri analiz ediliyor...")
        
        # Batch worker'ı başlat
        self.batch_worker = BatchAnalysisWorker(bist100_symbols)
        self.batch_worker.signals.progress.connect(self.update_progress)
        self.batch_worker.signals.finished.connect(self.batch_analysis_completed)
        self.batch_worker.start()
    
    def update_progress(self, value):
        self.progress_bar.setValue(value)
    
    def batch_analysis_completed(self, _, __, ___):
        # Tabloyu temizle
        self.table.setRowCount(0)
        
        # Sonuçları tabloya ekle
        results = self.batch_worker.results
        self.table.setRowCount(len(results))
        
        # Bugünün tarihini al
        today = datetime.now().strftime("%Y-%m-%d")
        
        # Portföy verilerini güncelle
        if today not in self.portfolio_data:
            self.portfolio_data[today] = {}
        
        # Önceki tarih verilerini al (karşılaştırma için)
        previous_data = self.load_previous_data()
        
        row = 0
        for symbol, (decision, price) in results.items():
            # Fiyatın float olduğundan emin ol
            price_float = float(price) if not isinstance(price, float) else price
            
            # Günlük yatırım ve hisse adedi hesapla
            daily_investment = 100.0  # Her gün 100 TL yatırım
            shares_to_buy = np.ceil(daily_investment / price_float)  # Yukarı yuvarla
            actual_cost = shares_to_buy * price_float  # Gerçek maliyet
            
            # Kar/zarar hesapla - önceki veriler varsa
            profit_loss = 0.0
            total_shares = shares_to_buy
            total_cost = actual_cost
            
            if symbol in previous_data:
                total_shares += previous_data[symbol]['shares']
                total_cost += previous_data[symbol]['cost']
                # Güncel toplam değer
                current_value = total_shares * price_float
                # Kar/zarar
                profit_loss = current_value - total_cost
            
            # Portföy verisini güncelle
            self.portfolio_data[today][symbol] = {
                'price': price_float,
                'decision': decision,
                'shares': total_shares,
                'cost': total_cost,
                'daily_shares': shares_to_buy,
                'daily_cost': actual_cost,
                'profit_loss': profit_loss
            }
            
            # Tabloya ekle
            # Hisse kodu
            self.table.setItem(row, 0, QTableWidgetItem(symbol))
            
            # Fiyat
            price_item = QTableWidgetItem(f"{price_float:.2f} TL")
            price_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.table.setItem(row, 1, price_item)
            
            # Öneri
            item = QTableWidgetItem(decision)
            if decision == 'AL':
                item.setBackground(Qt.green)
            elif decision == 'SAT':
                item.setBackground(Qt.red)
            self.table.setItem(row, 2, item)
            
            # Günlük alım adedi
            shares_item = QTableWidgetItem(f"{shares_to_buy:.0f}")
            shares_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.table.setItem(row, 3, shares_item)
            
            # Günlük yatırım miktarı
            cost_item = QTableWidgetItem(f"{actual_cost:.2f} TL")
            cost_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.table.setItem(row, 4, cost_item)
            
            # Kar/Zarar
            profit_item = QTableWidgetItem(f"{profit_loss:.2f} TL")
            profit_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            if profit_loss > 0:
                profit_item.setBackground(Qt.green)
            elif profit_loss < 0:
                profit_item.setBackground(Qt.red)
            self.table.setItem(row, 5, profit_item)
            
            row += 1
        
        # Sıralama için sinyal bağla
        self.table.horizontalHeader().sectionClicked.connect(self.sort_table)
        
        # UI'ı normale döndür
        self.batch_button.setEnabled(True)
        self.button.setEnabled(True)
        self.save_button.setEnabled(True)  # Excel'e kaydetme butonunu etkinleştir
        self.progress_bar.setVisible(False)
        self.update_status_bar("Analiz tamamlandı.")
    
    def load_previous_data(self):
        """
        Önceki günlerin portföy verilerini birleştirerek toplam hisse ve maliyet bilgilerini döndürür
        """
        previous_data = {}
        
        # Tarihleri sırala (en eskiden en yeniye)
        sorted_dates = sorted(self.portfolio_data.keys())
        
        # Bugünün tarihini çıkar (henüz kaydedilmedi)
        today = datetime.now().strftime("%Y-%m-%d")
        if today in sorted_dates:
            sorted_dates.remove(today)
        
        # Önceki tüm günlerin verilerini birleştir
        for date in sorted_dates:
            for symbol, data in self.portfolio_data[date].items():
                if symbol not in previous_data:
                    previous_data[symbol] = {
                        'shares': 0,
                        'cost': 0.0
                    }
                
                # Günlük alımları ekle
                previous_data[symbol]['shares'] += data['daily_shares']
                previous_data[symbol]['cost'] += data['daily_cost']
        
        return previous_data
    
    def sort_table(self, column_index):
        self.table.sortItems(column_index)
    
    def update_result(self, symbol, decision, price):
        # Tek bir hisse için sonuçları tabloya ekle
        matching_items = self.table.findItems(symbol, Qt.MatchExactly)
        
        # Fiyatın float olduğundan emin ol
        price_float = float(price) if not isinstance(price, float) else price
        
        # Günlük yatırım ve hisse adedi hesapla
        daily_investment = 100.0  # Her gün 100 TL yatırım
        shares_to_buy = np.ceil(daily_investment / price_float)  # Yukarı yuvarla
        actual_cost = shares_to_buy * price_float  # Gerçek maliyet
        
        # Bugünün tarihini al
        today = datetime.now().strftime("%Y-%m-%d")
        
        # Portföy verilerini güncelle
        if today not in self.portfolio_data:
            self.portfolio_data[today] = {}
            
        # Önceki tarih verilerini al (karşılaştırma için)
        previous_data = self.load_previous_data()
        
        # Kar/zarar hesapla - önceki veriler varsa
        profit_loss = 0.0
        total_shares = shares_to_buy
        total_cost = actual_cost
        
        if symbol in previous_data:
            total_shares += previous_data[symbol]['shares']
            total_cost += previous_data[symbol]['cost']
            # Güncel toplam değer
            current_value = total_shares * price_float
            # Kar/zarar
            profit_loss = current_value - total_cost
        
        # Portföy verisini güncelle
        self.portfolio_data[today][symbol] = {
            'price': price_float,
            'decision': decision,
            'shares': total_shares,
            'cost': total_cost,
            'daily_shares': shares_to_buy,
            'daily_cost': actual_cost,
            'profit_loss': profit_loss
        }
        
        if matching_items:
            # Hisse zaten tabloda var, güncelle
            row = matching_items[0].row()
        else:
            # Yeni satır ekle
            row = self.table.rowCount()
            self.table.setRowCount(row + 1)
            self.table.setItem(row, 0, QTableWidgetItem(symbol))
        
        # Fiyat bilgisi
        price_item = QTableWidgetItem(f"{price_float:.2f} TL")
        price_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.table.setItem(row, 1, price_item)
        
        # Al/Sat önerisi
        item = QTableWidgetItem(decision)
        if decision == 'AL':
            item.setBackground(Qt.green)
        elif decision == 'SAT':
            item.setBackground(Qt.red)
        self.table.setItem(row, 2, item)
        
        # Günlük alım adedi
        shares_item = QTableWidgetItem(f"{shares_to_buy:.0f}")
        shares_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.table.setItem(row, 3, shares_item)
        
        # Günlük yatırım miktarı
        cost_item = QTableWidgetItem(f"{actual_cost:.2f} TL")
        cost_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.table.setItem(row, 4, cost_item)
        
        # Kar/Zarar
        profit_item = QTableWidgetItem(f"{profit_loss:.2f} TL")
        profit_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
        if profit_loss > 0:
            profit_item.setBackground(Qt.green)
        elif profit_loss < 0:
            profit_item.setBackground(Qt.red)
        self.table.setItem(row, 5, profit_item)
        
        self.update_status_bar(f"'{symbol}' için işlem önerisi: {decision} (Fiyat: {price_float:.2f} TL)")
        
        # Excel kaydetme butonunu etkinleştir
        self.save_button.setEnabled(True)
    
    def show_error(self, error_message):
        self.update_status_bar(error_message)
        QMessageBox.warning(self, "Hata", error_message)
    
    def enable_button(self):
        self.button.setEnabled(True)
    
    def update_status_bar(self, message):
        # Burada bir status bar oluşturup mesaj gösterilebilir
        print(message)
    
    def select_excel_file(self):
        """Excel dosyası seç veya yeni oluştur"""
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getSaveFileName(self, "Excel Dosyası Seç/Oluştur", 
                                                  "BIST100_Analiz.xlsx", 
                                                  "Excel Files (*.xlsx);;All Files (*)", 
                                                  options=options)
        
        if file_name:
            self.excel_file_path = file_name
            
            # Eğer dosya varsa, önceki verileri yükle
            if os.path.exists(file_name):
                self.load_excel_data(file_name)
                QMessageBox.information(self, "Başarılı", f"Excel dosyası seçildi: {file_name}")
            else:
                QMessageBox.information(self, "Bilgi", f"Yeni Excel dosyası oluşturulacak: {file_name}")
    
    def load_excel_data(self, file_path):
        """Excel dosyasından önceki portföy verilerini yükle"""
        try:
            # Excel dosyasını oku
            excel_data = pd.read_excel(file_path, sheet_name='Özet', index_col=0)
            
            # Sütun isimlerinden tarihleri ayır
            for col in excel_data.columns:
                if '_' in col:
                    parts = col.split('_')
                    if len(parts) >= 3:  # Tarih_Veri_Türü formatında
                        date_str = parts[0]
                        data_type = '_'.join(parts[1:])
                        
                        # Tarihi kontrol et
                        try:
                            date_obj = datetime.strptime(date_str, "%Y-%m-%d")
                            date_key = date_obj.strftime("%Y-%m-%d")
                            
                            # Portföy verilerini oluştur
                            if date_key not in self.portfolio_data:
                                self.portfolio_data[date_key] = {}
                            
                            # Her bir hisse için verileri doldur
                            for index, value in excel_data[col].items():
                                symbol = index
                                
                                if symbol not in self.portfolio_data[date_key]:
                                    self.portfolio_data[date_key][symbol] = {
                                        'price': 0.0,
                                        'decision': '',
                                        'shares': 0,
                                        'cost': 0.0,
                                        'daily_shares': 0,
                                        'daily_cost': 0.0,
                                        'profit_loss': 0.0
                                    }
                                
                                # NaN değerleri kontrol et
                                if pd.notna(value):
                                    # Veri tipine göre değeri ata
                                    if data_type == 'Fiyat':
                                        self.portfolio_data[date_key][symbol]['price'] = float(value)
                                    elif data_type == 'Öneri':
                                        self.portfolio_data[date_key][symbol]['decision'] = str(value)
                                    elif data_type == 'Günlük_Hisse_Adedi':
                                        self.portfolio_data[date_key][symbol]['daily_shares'] = float(value)
                                    elif data_type == 'Günlük_Maliyet':
                                        self.portfolio_data[date_key][symbol]['daily_cost'] = float(value)
                                    elif data_type == 'Toplam_Hisse_Adedi':
                                        self.portfolio_data[date_key][symbol]['shares'] = float(value)
                                    elif data_type == 'Toplam_Maliyet':
                                        self.portfolio_data[date_key][symbol]['cost'] = float(value)
                                    elif data_type == 'Kar_Zarar':
                                        self.portfolio_data[date_key][symbol]['profit_loss'] = float(value)
                        except ValueError:
                            continue
        except Exception as e:
            print(f"Excel yükleme hatası: {str(e)}")
            QMessageBox.warning(self, "Hata", f"Excel dosyası yüklenirken bir hata oluştu: {str(e)}")
    
    def save_to_excel(self):
        """Analiz sonuçlarını Excel dosyasına kaydet"""
        if not self.excel_file_path:
            # Eğer dosya seçilmemişse, kullanıcıya seçtir
            self.select_excel_file()
            if not self.excel_file_path:
                return
        
        today = datetime.now().strftime("%Y-%m-%d")
        
        try:
            # Eğer dosya varsa, var olan dataframe'i yükle
            if os.path.exists(self.excel_file_path):
                try:
                    existing_df = pd.read_excel(self.excel_file_path, sheet_name='Özet', index_col=0)
                except:
                    # Dosya var ama içeriği okunamıyorsa yeni DataFrame oluştur
                    existing_df = pd.DataFrame(index=bist100_symbols)
            else:
                # Yeni DataFrame oluştur
                existing_df = pd.DataFrame(index=bist100_symbols)
            
            # Bugünün verilerini yeni sütunlar olarak ekle
            if today in self.portfolio_data:
                # Fiyat sütunu
                existing_df[f'{today}_Fiyat'] = pd.Series({symbol: data['price'] 
                                                       for symbol, data in self.portfolio_data[today].items()})
                
                # Öneri sütunu
                existing_df[f'{today}_Öneri'] = pd.Series({symbol: data['decision'] 
                                                      for symbol, data in self.portfolio_data[today].items()})
                
                # Günlük hisse adedi
                existing_df[f'{today}_Günlük_Hisse_Adedi'] = pd.Series({symbol: data['daily_shares'] 
                                                                   for symbol, data in self.portfolio_data[today].items()})
                
                # Günlük maliyet
                existing_df[f'{today}_Günlük_Maliyet'] = pd.Series({symbol: data['daily_cost'] 
                                                               for symbol, data in self.portfolio_data[today].items()})
                
                # Toplam hisse adedi
                existing_df[f'{today}_Toplam_Hisse_Adedi'] = pd.Series({symbol: data['shares'] 
                                                                   for symbol, data in self.portfolio_data[today].items()})
                
                # Toplam maliyet
                existing_df[f'{today}_Toplam_Maliyet'] = pd.Series({symbol: data['cost'] 
                                                               for symbol, data in self.portfolio_data[today].items()})
                
                # Kar/Zarar
                existing_df[f'{today}_Kar_Zarar'] = pd.Series({symbol: data['profit_loss'] 
                                                          for symbol, data in self.portfolio_data[today].items()})
            
            # DataFrame'i Excel'e kaydet
            with pd.ExcelWriter(self.excel_file_path, engine='openpyxl') as writer:
                existing_df.to_excel(writer, sheet_name='Özet')
            
            QMessageBox.information(self, "Başarılı", f"Analiz sonuçları {self.excel_file_path} dosyasına kaydedildi.")
            
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Excel dosyası kaydedilirken bir hata oluştu: {str(e)}")
            print(f"Excel kaydetme hatası: {str(e)}")

# BIST 100 Hisse Senedi Listesi
bist100_symbols = [
    'ACSEL', 'AEFES', 'AFYON', 'AGESA', 'AKBNK', 'AKCNS', 'AKFGY', 'AKSEN', 'ALARK', 'ALBRK', 
    'ALFAS', 'ALGYO', 'ALKIM', 'ARCLK', 'ARDYZ', 'ASELS', 'ASUZU', 'AYGAZ', 'BAGFS', 
    'BANVT', 'BERA', 'BIMAS', 'BIOEN', 'BRISA', 'BRSAN', 'BUCIM', 'CCOLA', 'CEMAS', 'CIMSA', 
    'DEVA', 'DOAS', 'DOHOL',  'ECILC', 'ECZYT', 'EGEEN', 'EKGYO', 'ENJSA', 'ENKAI', 'ERBOS', 
    'EREGL', 'FROTO', 'GARAN', 'GESAN', 'GLYHO', 'GOZDE', 'GSDHO', 'GUBRF', 'HALKB','HEKTS',
     'HLGYO', 'INDES', 'IPEKE', 'ISCTR', 'ISFIN', 'ISGYO', 'ISMEN', 'KARSN', 
    'KARTN', 'KCHOL', 'KLNMA', 'KMPUR', 'KONTR', 'KONYA', 'KORDS', 'KOZAA', 'KOZAL', 'KRDMD', 
    'KRVGD', 'LOGO', 'MAVI', 'MGROS', 'MPARK', 'NETAS', 'ODAS', 'OTKAR', 'OYAKC', 'PENTA', 
    'PETKM', 'PGSUS', 'PKENT', 'QUAGR', 'SAHOL', 'SASA', 'SELEC', 'SISE', 'SKBNK', 'SMRTG', 
    'SNGYO', 'SOKM', 'TAVHL', 'TCELL', 'TEKTU', 'THYAO', 'TOASO', 'TSKB', 'TTKOM', 
    'TTRAK', 'TUPRS', 'ULKER', 'VAKBN', 'VERUS', 'VESBE', 'VESTL', 'YKBNK', 'YYLGD', 'ZOREN'
]
# Mevcut dosyadan eski verileri yükle
try:
    if os.path.exists(self.excel_file_path) and os.path.getsize(self.excel_file_path) > 0:
        existing_df = pd.read_excel(self.excel_file_path, sheet_name='Özet', index_col=0)
                    
# Yeni sütunları ekle
    for col in df.columns:
        existing_df[col] = df[col]
                    
# Güncellenen veriyi kullan
        df = existing_df
    except Exception as e:
    print(f"Eski veri okuma hatası: {str(e)}")
    # Devam et, yeni bir Excel dosyası oluşacak
    pass
            
# Excel Writer'ı başlat
    try:
# Excel dosyasını başka bir süreç kullanıyor olabilir, bu yüzden önce bir kontrol yapalım
    is_file_open = False
    try:
    with open(self.excel_file_path, 'a') as f:
        pass
    except PermissionError:
    is_file_open = True
                
    if is_file_open:
    QMessageBox.warning(self, "Uyarı", "Excel dosyası başka bir program tarafından açık. Lütfen kapatıp tekrar deneyin.")
    return
                
# Excel'e kaydet
    with pd.ExcelWriter(self.excel_file_path, engine='openpyxl', mode='w') as writer:
                    df.to_excel(writer, sheet_name='Özet')
# Uygulama başlatma
if __name__ == "__main__":
    app = QApplication([])
    window = StockApp()
    window.show()
    app.exec_()
