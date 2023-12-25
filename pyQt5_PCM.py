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
    QScatterSeries,
)
from PyQt5.QtGui import QPainter
from PyQt5.QtCore import Qt, QPointF


class ScatterPlotWindow(QWidget):
    def __init__(self, column_data):
        super().__init__()

        self.setWindowTitle("Scatter Plot")
        self.setGeometry(200, 200, 800, 600)

        layout = QVBoxLayout(self)

        self.chart = QChart()
        self.scatter_series = QScatterSeries()

        for i, data in enumerate(column_data):
            self.scatter_series.append(i, float(data))

        self.chart.addSeries(self.scatter_series)
        self.chart.createDefaultAxes()
        self.chart.setTitle("Scatter Plot")

        self.chart_view = QChartView(self.chart)
        self.chart_view.setRenderHint(QPainter.Antialiasing)

        layout.addWidget(self.chart_view)

        data_table = QTableWidget(self)
        data_table.setColumnCount(1)
        data_table.setRowCount(len(column_data))
        data_table.setHorizontalHeaderLabels(["Data"])

        for i, data in enumerate(column_data):
            item = QTableWidgetItem(str(data))
            data_table.setItem(i, 0, item)

        layout.addWidget(data_table)

        self.selected_point_index = None
        self.data_table = data_table  # 데이터 목록을 인스턴스 변수로 추가

        self.chart_view.mousePressEvent = self.on_mouse_click  # 마우스 클릭 이벤트 연결

    def find_nearest_point(self, pos):
        # 이전에 선택된 데이터의 하이라이트 제거
        if self.selected_point_index is not None:
            for i in range(self.data_table.rowCount()):
                for j in range(self.data_table.columnCount()):
                    item = self.data_table.item(i, j)
                    item.setBackground(Qt.white)

        # 화면 좌표를 차트 플롯 영역 좌표로 변환
        graph_point = self.chart.mapToValue(pos)

        # x좌표를 반올림해서 가장 가까운 자연수로 변환
        x_value = round(graph_point.x())

        # ScatterSeries에서 x가 가장 가까운 자연수를 찾기
        min_distance_x = float("inf")
        nearest_x_index = None

        for i, point in enumerate(self.scatter_series.points()):
            distance_x = abs(point.x() - x_value)
            if distance_x < min_distance_x:
                min_distance_x = distance_x
                nearest_x_index = i

        # 선택된 x값에 대응하는 데이터 찾기
        selected_data = None
        min_distance_y = float("inf")

        for i in range(self.data_table.rowCount()):
            item = self.data_table.item(i, 0)
            y_value = float(item.text())

            distance_y = abs(y_value - graph_point.y())
            if distance_y < min_distance_y:
                min_distance_y = distance_y
                selected_data = (i, y_value)

        if selected_data and nearest_x_index is not None:
            selected_index, selected_y_value = selected_data

            # 선택한 데이터 행을 초록색으로 하이라이트
            for j in range(self.data_table.columnCount()):
                item = self.data_table.item(selected_index, j)
                item.setBackground(Qt.green)

            # 좌표 표시
            QMessageBox.information(
                self,
                "Clicked Point",
                f"X: {x_value}, Y: {selected_y_value}",
            )
        else:
            QMessageBox.warning(self, "Point Not Found", "해당하는 점을 찾을 수 없습니다.")

    def on_mouse_click(self, event):
        pos = event.pos()

        # ScatterSeries에서 가장 가까운 점 찾기
        self.find_nearest_point(pos)


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

        self.scatter_plot_window = None

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

        # 인스턴스 변수로 선언한 scatter_plot_window를 사용
        self.scatter_plot_window = ScatterPlotWindow(column_data)
        self.scatter_plot_window.show()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
