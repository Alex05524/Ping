import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QTableWidget, \
    QTableWidgetItem, QPushButton, QPlainTextEdit, QFileDialog, QGroupBox, QHeaderView
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtGui import QPainter, QColor
import pandas as pd
from ping3 import ping

class RoundedCornersWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.radius = 15

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setBrush(QColor("#f0f0f0"))  # Background color
        painter.setPen(QColor("#f0f0f0"))
        painter.drawRoundedRect(self.rect(), self.radius, self.radius)

class PingThread(QThread):
    update_table = pyqtSignal(str, str, str)

    def __init__(self, ip_addresses, ping_count=4):
        super().__init__()
        self.ip_addresses = ip_addresses
        self.ping_count = ping_count

    def run(self):
        for ip in self.ip_addresses:
            rtt_sum = 0  # To calculate average RTT
            status = 'Не отвечает'

            for _ in range(self.ping_count):
                try:
                    rtt = ping(ip)
                    if rtt is not None:
                        status = 'Да'
                        rtt_sum += rtt
                except Exception as e:
                    print(f"An error occurred: {e}")

            if status == 'Да':
                rtt_average = rtt_sum / self.ping_count
            else:
                rtt_average = ''

            self.update_table.emit(ip, status, str(rtt_average))

class App(RoundedCornersWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setGeometry(300, 300, 1200, 600)
        self.setWindowTitle('Ping IP Addresses')

        layout = QHBoxLayout(self)

        # Server Section
        server_group = QGroupBox("Сервера")
        server_layout = QVBoxLayout()

        self.server_table = QTableWidget(self)
        self.server_table.setColumnCount(3)
        self.server_table.setHorizontalHeaderLabels(['IP Address', 'Status', 'RTT (ms)'])
        self.server_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        server_layout.addWidget(self.server_table)

        server_buttons_layout = QHBoxLayout()
        server_buttons_layout.addWidget(QPushButton('Старт мониторинга (.3)', clicked=self.start_server_monitoring))
        server_buttons_layout.addWidget(QPushButton('Стоп мониторинга', clicked=self.stop_server_monitoring))
        server_buttons_layout.addWidget(QPushButton('Сохранить в Excel', clicked=self.save_server_to_excel))
        server_layout.addLayout(server_buttons_layout)

        server_group.setLayout(server_layout)

        # Spacer
        layout.addSpacing(20)

        # Network Section
        network_group = QGroupBox("Общая сеть")
        network_layout = QVBoxLayout()

        self.network_table = QTableWidget(self)
        self.network_table.setColumnCount(3)
        self.network_table.setHorizontalHeaderLabels(['IP Address', 'Status', 'RTT (ms)'])
        self.network_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        network_layout.addWidget(self.network_table)

        network_buttons_layout = QHBoxLayout()
        network_buttons_layout.addWidget(QPushButton('Старт мониторинга (.1)', clicked=self.start_network_monitoring))
        network_buttons_layout.addWidget(QPushButton('Стоп мониторинга', clicked=self.stop_network_monitoring))
        network_buttons_layout.addWidget(QPushButton('Сохранить в Excel', clicked=self.save_network_to_excel))
        network_layout.addLayout(network_buttons_layout)

        network_group.setLayout(network_layout)

        layout.addWidget(server_group)
        layout.addWidget(network_group)

        # Applying a simple stylesheet for a modern Windows 11-like style
        stylesheet = """
            QPushButton {
                background-color: #2f80ed;
                color: #ffffff;
                padding: 10px;
                border: none;
                border-radius: 5px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2a66b5;
            }
            QTableWidget {
                background-color: #ffffff;
                border-radius: 10px;
            }
            QTableWidget::item {
                padding: 5px;
            }
            QGroupBox {
                border: 2px solid #2f80ed;
                border-radius: 10px;
                margin-top: 10px;
            }
            QGroupBox::title {
                color: #2f80ed;
                subcontrol-origin: margin;
                left: 10px;
            }
        """
        self.setStyleSheet(stylesheet)

        self.show()

        self.server_thread = None
        self.network_thread = None

    def start_server_monitoring(self):
        ip_addresses = ["192.168.30.3", "192.168.54.3", 
                        "192.168.88.3", "192.168.89.3", 
                        "192.168.90.3", "192.168.91.3", 
                        "192.168.92.3", "192.168.93.3", 
                        "192.168.94.3", "192.168.95.3", 
                        "192.168.96.3", "192.168.98.3", 
                        "192.168.99.3", "192.168.101.3", 
                        "192.168.102.3", "192.168.103.3", 
                        "192.168.104.3", "192.168.105.3", 
                        "192.168.106.3", "192.168.107.3", 
                        "192.168.108.3", "192.168.109.3"]
        self.server_thread = PingThread(ip_addresses, ping_count=4)
        self.server_thread.update_table.connect(self.update_server_table)
        self.server_thread.start()

    def stop_server_monitoring(self):
        if self.server_thread is not None:
            self.server_thread.terminate()
            self.server_thread.wait()
            self.server_thread = None

    def start_network_monitoring(self):
        ip_addresses = ["192.168.30.1", "192.168.54.1", 
                        "192.168.88.1", "192.168.89.1", 
                        "192.168.90.1", "192.168.91.1",
                        "192.168.92.1", "192.168.93.1", 
                        "192.168.94.1", "192.168.95.1",
                        "192.168.96.1", "192.168.98.1", 
                        "192.168.99.1", "192.168.101.1", 
                        "192.168.102.1", "192.168.103.1",
                        "192.168.104.1", "192.168.105.1", 
                        "192.168.106.1", "192.168.107.1",
                        "192.168.108.1", "192.168.109.1"]
        self.network_thread = PingThread(ip_addresses, ping_count=4)
        self.network_thread.update_table.connect(self.update_network_table)
        self.network_thread.start()

    def stop_network_monitoring(self):
        if self.network_thread is not None:
            self.network_thread.terminate()
            self.network_thread.wait()
            self.network_thread = None

    def update_server_table(self, ip, status, rtt):
        row_position = self.server_table.rowCount()
        self.server_table.insertRow(row_position)
        self.server_table.setItem(row_position, 0, QTableWidgetItem(ip))
        self.server_table.setItem(row_position, 1, QTableWidgetItem(status))
        self.server_table.setItem(row_position, 2, QTableWidgetItem(rtt))

    def update_network_table(self, ip, status, rtt):
        row_position = self.network_table.rowCount()
        self.network_table.insertRow(row_position)
        self.network_table.setItem(row_position, 0, QTableWidgetItem(ip))
        self.network_table.setItem(row_position, 1, QTableWidgetItem(status))
        self.network_table.setItem(row_position, 2, QTableWidgetItem(rtt))

    def save_server_to_excel(self):
        self.save_to_excel(self.server_table, 'Сервера')

    def save_network_to_excel(self):
        self.save_to_excel(self.network_table, 'Общая сеть')

    def save_to_excel(self, table, label):
        filename, _ = QFileDialog.getSaveFileName(self, f'Сохранить в Excel ({label})', '', 'Excel Files (*.xlsx)')
        if filename:
            data = [(table.item(row, 0).text(), table.item(row, 1).text(), table.item(row, 2).text())
                    for row in range(table.rowCount())]
            df = pd.DataFrame(data, columns=['IP Address', 'Status', 'RTT'])
            df.to_excel(filename, index=False, header=True)

def main():
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
