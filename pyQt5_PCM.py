import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication,
    QMainWindow,
    QVBoxLayout,
    QWidget,
    QPushButton,
    QLabel,
    QFileDialog,
    QTableWidget,
    QTableWidgetItem,
    QInputDialog,
    QMessageBox,
)
from PyQt5.QtChart import (
    QChart,
    QChartView,
    QLineSeries,
    QBarSet,
    QBarSeries,
    QValueAxis,
)
from PyQt5.QtGui import QPainter


class ScatterPlotWindow(QWidget):
    def __init__(self, column_data):
        super().__init__()

        self.setWindowTitle("Scatter Plot")
        self.setGeometry(200, 200, 800, 600)

        layout = QVBoxLayout(self)

        chart = QChart()
        series = QLineSeries()

        for i, data in enumerate(column_data):
            series.append(i, float(data))

        chart.addSeries(series)
        chart.createDefaultAxes()
        chart.setTitle("Scatter Plot")

        chart_view = QChartView(chart)
        chart_view.setRenderHint(QPainter.Antialiasing)

        layout.addWidget(chart_view)

        data_table = QTableWidget(self)
        data_table.setColumnCount(1)
        data_table.setRowCount(len(column_data))
        data_table.setHorizontalHeaderLabels(["Data"])

        for i, data in enumerate(column_data):
            item = QTableWidgetItem(str(data))
            data_table.setItem(i, 0, item)

        layout.addWidget(data_table)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Excel 뷰어")
        self.setGeometry(100, 100, 800, 600)

        self.central_widget = QWidget(self)
        self.setCentralWidget(self.central_widget)

        self.layout = QVBoxLayout(self.central_widget)

        self.label = QLabel("선택된 파일 없음", self)
        self.layout.addWidget(self.label)

        self.upload_button = QPushButton("Excel 업로드", self)
        self.upload_button.clicked.connect(self.upload_excel)
        self.layout.addWidget(self.upload_button)

        self.table = QTableWidget(self)
        self.layout.addWidget(self.table)

        self.modify_button = QPushButton("데이터 수정", self)
        self.modify_button.clicked.connect(self.modify_data)
        self.layout.addWidget(self.modify_button)

        self.delete_button = QPushButton("데이터 삭제", self)
        self.delete_button.clicked.connect(self.delete_data)
        self.layout.addWidget(self.delete_button)

        self.filename = None
        self.df = None

    def upload_excel(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly

        file_dialog = QFileDialog(self)
        file_dialog.setNameFilter("Excel 파일 (*.xlsx *.xls)")
        file_dialog.setFileMode(QFileDialog.ExistingFile)
        file_dialog.setOptions(options)

        if file_dialog.exec_():
            self.filename = file_dialog.selectedFiles()[0]
            self.label.setText(f"선택된 파일: {self.filename}")

            # 새로운 파일 업로드 시 기존 데이터 초기화
            self.df = None

            try:
                self.df = pd.read_excel(self.filename)
            except Exception as e:
                print(f"Error reading Excel file: {e}")

            if self.df is not None:
                self.update_table()

    def update_table(self):
        self.table.clear()

        if self.df is not None:
            self.table.setRowCount(self.df.shape[0])
            self.table.setColumnCount(self.df.shape[1])

            for i in range(self.df.shape[0]):
                for j in range(self.df.shape[1]):
                    item = QTableWidgetItem(str(self.df.iloc[i, j]))
                    self.table.setItem(i, j, item)

            # 셀 선택 시 새로운 창 열기
            self.table.itemClicked.connect(self.show_scatter_plot)

    def modify_data(self):
        if self.df is not None:
            selected_item = self.table.currentItem()

            if selected_item is not None:
                selected_row = selected_item.row()
                selected_col = selected_item.column()

                new_value, ok_pressed = QInputDialog.getText(
                    self,
                    "데이터 수정",
                    "새로운 값을 입력하세요:",
                    text=self.df.iloc[selected_row, selected_col],
                )

                if ok_pressed:
                    self.df.iloc[selected_row, selected_col] = new_value
                    self.update_table()

    def delete_data(self):
        if self.df is not None:
            selected_item = self.table.currentItem()

            if selected_item is not None:
                selected_row = selected_item.row()
                selected_col = selected_item.column()

                reply = QMessageBox.question(
                    self,
                    "데이터 삭제",
                    "선택한 셀을 삭제하시겠습니까?\n(삭제 시 해당 값은 0으로 대체됩니다.)",
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.No,
                )

                if reply == QMessageBox.Yes:
                    self.df.iloc[selected_row, selected_col] = 0
                    self.update_table()

    def show_scatter_plot(self, item):
        selected_row = item.row()
        selected_col = item.column()

        column_data = self.df.iloc[:, selected_col].tolist()

        scatter_plot_window = ScatterPlotWindow(column_data)
        scatter_plot_window.show()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
