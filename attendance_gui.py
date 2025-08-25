import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGroupBox, QLabel, QLineEdit, QPushButton, QTableWidget, QTableWidgetItem,
    QTextEdit, QDateEdit, QMessageBox, QMenuBar, QAction
)
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtGui import QIntValidator, QIcon
from attendance_system import ZKTecoAttendance
import pandas as pd


class AttendanceWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("COC Attendance System")
        self.resize(1300, 650)

        # Set icon if available
        try:
            self.setWindowIcon(QIcon("icon.png"))
        except Exception:
            pass

        self.attendance_system = None

        # -------- Menu Bar (with icon space) --------
        menu_bar = self.menuBar()
        self.brand_action = QAction(QIcon("icon.png"), "", self)
        self.brand_action.setEnabled(False)
        menu_bar.addAction(self.brand_action)

        help_menu = menu_bar.addMenu("Help")
        about_action = QAction("About", self)
        about_action.triggered.connect(self.show_about_dialog)
        help_menu.addAction(about_action)

        # -------- Layout Setup --------
        central = QWidget()
        root_layout = QHBoxLayout(central)
        root_layout.setContentsMargins(16, 16, 16, 16)
        root_layout.setSpacing(16)

        # Left Column (Device, Date, Table)
        left_col = QVBoxLayout()
        left_col.setSpacing(12)

        # Device Connection Box
        conn_group = QGroupBox("Device Connection")
        conn_layout = QHBoxLayout(conn_group)
        conn_layout.addWidget(QLabel("IP Address:"))
        self.ip_edit = QLineEdit("192.168.1.201")
        self.ip_edit.setFixedWidth(160)
        conn_layout.addWidget(self.ip_edit)

        conn_layout.addWidget(QLabel("Port:"))
        self.port_edit = QLineEdit("4370")
        self.port_edit.setValidator(QIntValidator(1, 65535, self))
        self.port_edit.setFixedWidth(80)
        conn_layout.addWidget(self.port_edit)

        self.connect_btn = QPushButton("Connect")
        self.connect_btn.clicked.connect(self.connect_device)
        conn_layout.addWidget(self.connect_btn)
        conn_layout.addStretch(1)
        left_col.addWidget(conn_group)

        # Date Filter Box
        date_group = QGroupBox("Date Filter")
        date_layout = QHBoxLayout(date_group)
        date_layout.addWidget(QLabel("Start Date:"))
        self.start_date = QDateEdit(calendarPopup=True)
        self.start_date.setDisplayFormat("yyyy-MM-dd")
        self.start_date.setDate(QDate.currentDate())
        date_layout.addWidget(self.start_date)

        date_layout.addWidget(QLabel("End Date:"))
        self.end_date = QDateEdit(calendarPopup=True)
        self.end_date.setDisplayFormat("yyyy-MM-dd")
        self.end_date.setDate(QDate.currentDate())
        date_layout.addWidget(self.end_date)

        self.retrieve_btn = QPushButton("Retrieve Records")
        self.retrieve_btn.setEnabled(False)
        self.retrieve_btn.clicked.connect(self.retrieve_records)
        date_layout.addWidget(self.retrieve_btn)

        self.export_btn = QPushButton("Export to CSV")
        self.export_btn.setEnabled(False)
        self.export_btn.clicked.connect(self.export_records)
        date_layout.addWidget(self.export_btn)

        date_layout.addStretch(1)
        left_col.addWidget(date_group)

        # Attendance Records Table
        self.table = QTableWidget(0, 6)
        self.table.setAlternatingRowColors(True)
        self.table.setHorizontalHeaderLabels([
            "User ID", "Name", "Date", "Check In", "Check Out", "Duration (hours)"
        ])
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionBehavior(self.table.SelectRows)
        self.table.setEditTriggers(self.table.NoEditTriggers)
        self.table.horizontalHeader().setStretchLastSection(True)

        try:
            from PyQt5.QtWidgets import QHeaderView
            header = self.table.horizontalHeader()
            header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
            header.setSectionResizeMode(1, QHeaderView.Stretch)
            header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
            header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
            header.setSectionResizeMode(4, QHeaderView.ResizeToContents)
            header.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        except Exception:
            pass

        left_col.addWidget(self.table, 1)
        left_wrap = QWidget()
        left_wrap.setLayout(left_col)
        root_layout.addWidget(left_wrap, 3)

        # Right Column (Steps to Use)
        steps_group = QGroupBox("Steps to Use")
        steps_layout = QVBoxLayout(steps_group)
        self.steps_text = QTextEdit()
        self.steps_text.setReadOnly(True)
        self.steps_text.setPlainText(
            "A. Quick Setup\n"
            "  1) Plug in your ZKTeco MB460 device via ethernet (stable option)\n and ensure it's on the same network.\n"
            "  2) Confirm the IP address (e.g., 192.168.1.201).\n"

            "B. Connect\n"
            "  1) Enter IP and Port, click ‘Connect’.\n"
            "  2) Watch status bar for ‘Connected’ confirmation.\n\n"
            "C. Retrieve Logs\n"
            "  1) Select Start and End dates.\n"
            "  2) Click ‘Retrieve Records’ to populate the table.\n\n"
            "D. Export\n"
            "  1) Click ‘Export to CSV’ to save filtered records.\n"
            "  2) Use Excel or Google Sheets for analysis.\n\n"
            "E. Pro Tips\n"
            "  • Sync happens automatically only on Sunday, Monday, Wednesday, and Friday.\n"
            "  • Keep device clock accurate.\n"
            "  • Static IP must be set to prevent random disconnections.\n"
        )
        steps_layout.addWidget(self.steps_text)
        root_layout.addWidget(steps_group, 1)

        self.setCentralWidget(central)
        self.statusBar().showMessage("Not connected")

        # UI Styling
        self.setStyleSheet(
            "QMainWindow{background:#F7F9FC;}"
            "QMenuBar{background:#FFFFFF;border-bottom:1px solid #E5EAF2;padding:4px;}"
            "QMenuBar::item{padding:4px 10px;margin:0 2px;border-radius:6px;}"
            "QMenuBar::item:selected{background:#EEF2FA;}"
            "QGroupBox{font-weight:600;border:1px solid #D9DEE7;border-radius:10px;"
            "margin-top:8px;padding:10px;}"
            "QGroupBox::title{subcontrol-origin: margin; left:10px; padding:0 6px;}"
            "QLabel{font-size:11pt;color:#222;}"
            "QPushButton{padding:6px 12px;border:1px solid #C5CCD8;border-radius:8px;"
            "background:#FFFFFF;font-weight:600;}"
            "QPushButton:hover{background:#F6F8FF;}"
            "QPushButton:disabled{color:#888;background:#F0F2F6;}"
            "QTableWidget{background:#FFFFFF;border:1px solid #D9DEE7;border-radius:8px;"
            "gridline-color:#EDF1F7;}"
            "QHeaderView::section{background:#FBFCFE;border:none;border-bottom:1px solid #E6EBF3;"
            "padding:6px;font-weight:600;}"
            "QTableWidget::item:selected{background:#162761;}"
        )

    # -------- Logic --------
    def connect_device(self):
        try:
            ip = self.ip_edit.text().strip()
            port = int(self.port_edit.text().strip() or 4370)
            self.attendance_system = ZKTecoAttendance(ip, port=port)
            self.attendance_system.connect()

            if getattr(self.attendance_system, 'conn', None):
                self.statusBar().showMessage(f"Connected to {ip}")
                self.connect_btn.setEnabled(False)
                self.retrieve_btn.setEnabled(True)
            else:
                self.statusBar().showMessage("Connection failed")
        except Exception as e:
            QMessageBox.critical(self, "Connection Error", str(e))
            self.statusBar().showMessage("Connection failed")

    def retrieve_records(self):
        if not self.attendance_system or not getattr(self.attendance_system, 'conn', None):
            QMessageBox.critical(self, "Error", "Please connect to the device first")
            return
        try:
            self.table.setRowCount(0)
            start_date = self.start_date.date().toPyDate()
            end_date = self.end_date.date().toPyDate()
            records = self.attendance_system.get_attendance(start_date, end_date)

            if records is not None and not records.empty:
                self.populate_table(records)
                self.statusBar().showMessage(f"Retrieved {len(records)} records")
                self.export_btn.setEnabled(True)
            else:
                self.statusBar().showMessage("No records found")
                self.export_btn.setEnabled(False)
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))
            self.statusBar().showMessage("Error retrieving records")

    def populate_table(self, df: pd.DataFrame):
        expected_cols = ["user_id", "user_name", "date", "check_in", "check_out", "duration"]
        for col in expected_cols:
            if col not in df.columns:
                df[col] = pd.NA

        self.table.setRowCount(len(df))
        for r, (_, row) in enumerate(df.iterrows()):
            self.table.setItem(r, 0, QTableWidgetItem(str(row["user_id"])))
            self.table.setItem(r, 1, QTableWidgetItem(str(row["user_name"])))
            self.table.setItem(r, 2, QTableWidgetItem(str(row["date"])))
            self.table.setItem(r, 3, QTableWidgetItem(str(row["check_in"])))
            self.table.setItem(r, 4, QTableWidgetItem(str(row["check_out"])))
            duration = f"{row['duration']:.2f}" if pd.notnull(row['duration']) else "N/A"
            self.table.setItem(r, 5, QTableWidgetItem(duration))

    def export_records(self):
        if not self.attendance_system or not getattr(self.attendance_system, 'conn', None):
            QMessageBox.critical(self, "Error", "Please connect to the device first")
            return
        try:
            start_date = self.start_date.date().toPyDate()
            end_date = self.end_date.date().toPyDate()
            records = self.attendance_system.get_attendance(start_date, end_date)

            if records is not None and not records.empty:
                filename = f"attendance_records_{start_date}_{end_date}.csv"
                records.to_csv(filename, index=False)
                self.statusBar().showMessage(f"Exported to {filename}")
                QMessageBox.information(self, "Success", f"Records exported to {filename}")
            else:
                self.statusBar().showMessage("No records to export")
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))
            self.statusBar().showMessage("Error exporting records")

    def show_about_dialog(self):
        QMessageBox.information(
            self,
            "About",
            "COC Attendance System\nVersion 1.0\nSyncs automatically on Sunday, Monday, Wednesday, and Friday."
        )


def main():
    app = QApplication(sys.argv)
    win = AttendanceWindow()
    win.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
