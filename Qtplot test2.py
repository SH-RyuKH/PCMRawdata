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
    QListWidget,
    QListWidgetItem,
)
from PyQt5.QtGui import QPainter
from PyQt5.QtCore import Qt
import pyqtgraph as pg


class ScatterPlotWindow(QWidget):
    def __init__(self, column_data, raw_data_window):
        super().__init__()

        self.setWindowTitle("Scatter Plot")
        self.setGeometry(200, 200, 800, 600)

        self.raw_data_window = raw_data_window  # PCM Raw Data 창 참조

        layout = QVBoxLayout(self)

        # 그래프 위젯 생성
        self.plot_widget = pg.PlotWidget(self)
        layout.addWidget(self.plot_widget)

        # scatter 그래프 생성
        self.scatter_plot = self.plot_widget.plot(
            x=list(range(len(column_data))),
            y=column_data,
            pen=None,
            symbol="o",
            symbolSize=10,
        )

        # 그래프 클릭 이벤트 연결
        self.plot_widget.scene().sigMouseClicked.connect(self.onMouseClicked)

        # 선택된 좌표의 X 및 Y 값을 표시할 레이블 추가
        self.selected_coordinates_label = QLabel("선택한 데이터: X: ?  Y: ?", self)
        layout.addWidget(self.selected_coordinates_label)

        # 모든 데이터의 목록을 표시할 목록 위젯 추가
        self.data_list_widget = QListWidget(self)
        layout.addWidget(self.data_list_widget)

        # "데이터 삭제" 버튼 추가
        self.delete_data_button = QPushButton("데이터 삭제", self)
        self.delete_data_button.clicked.connect(self.delete_selected_data)
        layout.addWidget(self.delete_data_button)

        # "데이터 업데이트" 버튼 추가
        self.update_data_button = QPushButton("데이터 업데이트", self)
        self.update_data_button.clicked.connect(self.update_data)
        layout.addWidget(self.update_data_button)

        # 창 설정
        self.data = {"x": list(range(len(column_data))), "y": column_data}

        self.selected_point_index = None
        self.update_data_list()

    def handle_item_double_clicked(self, item):
        # 데이터 목록에서 더블 클릭한 항목의 인덱스 가져오기
        selected_index = self.data_list_widget.row(item)

        # 선택한 데이터의 좌표 표시
        x_value = self.data["x"][selected_index]
        y_value = self.data["y"][selected_index]
        self.selected_coordinates_label.setText(
            f"선택한 데이터: X: {x_value:.2f}  Y: {y_value:.2f}"
        )

    def onMouseClicked(self, event):
        pos = event.scenePos()
        pos_item = self.plot_widget.plotItem.vb.mapSceneToView(pos)

        # Scatter 그래프에서 가장 가까운 점 찾기
        distances = [
            ((pos_item.x() - x) ** 2 + (pos_item.y() - y) ** 2) ** 0.5
            for x, y in zip(self.data["x"], self.data["y"])
        ]
        min_distance = min(distances)
        min_index = distances.index(min_distance)

        # 가장 가까운 점의 좌표 표시
        x_value = self.data["x"][min_index]
        y_value = self.data["y"][min_index]
        self.selected_coordinates_label.setText(
            f"가장 가까운 점: X: {x_value:.2f}  Y: {y_value:.2f}"
        )

        # 현재 클릭한 점 하이라이트
        current_item = self.data_list_widget.item(min_index)

        # 이전에 클릭한 점의 스타일 초기화 (기본 스타일)
        for i in range(self.data_list_widget.count()):
            item = self.data_list_widget.item(i)
            item.setSelected(False)

        # 현재 클릭한 점 스타일 변경
        current_item.setSelected(True)

        # 선택한 데이터의 인덱스 저장
        self.selected_point_index = min_index

    def update_data_list(self):
        # 목록을 지우고 다시 추가
        self.data_list_widget.clear()
        for x, y in zip(self.data["x"], self.data["y"]):
            item = QListWidgetItem(f"X: {x:.2f}  Y: {y:.2f}")
            self.data_list_widget.addItem(item)

    def delete_selected_data(self):
        if self.selected_point_index is not None:
            # 선택한 데이터의 y 값을 0으로 만들기
            selected_y_value = self.data["y"][self.selected_point_index]
            self.data["y"][self.selected_point_index] = 0

            # 그래프 업데이트
            self.scatter_plot.setData(x=self.data["x"], y=self.data["y"])

            # 선택한 데이터의 인덱스 초기화
            self.selected_point_index = None

            # 데이터 목록 업데이트
            self.update_data_list()

            # 선택한 데이터의 좌표 표시 초기화
            self.selected_coordinates_label.setText("선택한 데이터: X: ?  Y: ?")

    def update_data(self):
        # Scatter Plot 창에서 PCM Raw Data 창의 데이터 업데이트 호출
        self.raw_data_window.update_data()

        # 업데이트된 데이터를 가져와서 그래프와 데이터 목록 업데이트
        column_data = self.raw_data_window.get_selected_column_data()
        self.scatter_plot.setData(x=list(range(len(column_data))), y=column_data)
        self.data = {"x": list(range(len(column_data))), "y": column_data}
        self.update_data_list()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("PCM Raw Data")
        self.setGeometry(100, 100, 1600, 1200)

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

        self.chart_button = QPushButton("차트", self)
        self.chart_button.clicked.connect(self.show_scatter_plot)
        self.layout.addWidget(self.chart_button)

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
            # "공정조건"이라는 column 추가
            self.df["공정조건"] = self.create_process_column(self.df.iloc[:, 0])

            # "공정조건" 칼럼을 두 번째 칼럼으로 이동
            cols = list(self.df.columns)
            cols = [cols[0], "공정조건"] + cols[1:-1] + [cols[-1]]
            self.df = self.df[cols]

            self.table.setRowCount(self.df.shape[0])
            self.table.setColumnCount(self.df.shape[1])

            for i in range(self.df.shape[0]):
                for j in range(self.df.shape[1]):
                    value = self.df.iloc[i, j]
                    item = QTableWidgetItem(self.get_formatted_value(value))
                    self.table.setItem(i, j, item)

    def get_formatted_value(self, value):
        if isinstance(value, int):
            return str(value)
        elif isinstance(value, float):
            return "{:.6f}".format(value)
        else:
            return str(value)

    def create_process_column(self, column_data):
        process_column = []
        current_group = None
        alphabet_index = 0

        for value in column_data:
            if current_group is None or value == 1:
                current_group = alphabet_index
                alphabet_index = (alphabet_index + 1) % 26
            process_column.append(chr(ord("A") + current_group))

        return process_column

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

    def show_scatter_plot(self):
        selected_item = self.table.currentItem()

        if selected_item is not None:
            selected_row = selected_item.row()
            selected_col = selected_item.column()

            column_data = self.df.iloc[:, selected_col].tolist()

            # 인스턴스 변수로 선언한 scatter_plot_window를 사용
            self.scatter_plot_window = ScatterPlotWindow(column_data, self)
            self.scatter_plot_window.show()

    def update_data(self):
        # 파일을 다시 불러와서 데이터 업데이트
        self.df = pd.read_excel(self.filename)

        # 업데이트된 데이터로 테이블과 차트 업데이트
        self.update_table()
        if self.scatter_plot_window:
            self.scatter_plot_window.update_data_list()

    def get_selected_column_data(self):
        selected_item = self.table.currentItem()
        if selected_item is not None:
            selected_col = selected_item.column()
            return self.df.iloc[:, selected_col].tolist()
        return []


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
