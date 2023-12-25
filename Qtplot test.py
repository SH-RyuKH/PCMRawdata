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
    def __init__(self, column_data):
        super().__init__()

        self.setWindowTitle("Scatter Plot")
        self.setGeometry(200, 200, 800, 600)

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
        self.selected_coordinates_label = QLabel("가장 가까운 점: X: ?  Y: ?", self)
        layout.addWidget(self.selected_coordinates_label)

        # 모든 데이터의 목록을 표시할 목록 위젯 추가
        self.data_list_widget = QListWidget(self)
        layout.addWidget(self.data_list_widget)

        # "데이터 삭제" 버튼 추가
        self.delete_data_button = QPushButton("데이터 삭제", self)
        self.delete_data_button.clicked.connect(self.delete_selected_data)
        layout.addWidget(self.delete_data_button)

        # 창 설정
        self.data = {"x": list(range(len(column_data))), "y": column_data}

        self.update_data_list()

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

            # 클릭한 데이터 보여주기
            self.show_data(
                [self.data["x"][self.selected_point_index], selected_y_value]
            )


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("PCM Raw Data")
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
            self.table.setRowCount(self.df.shape[0])
            self.table.setColumnCount(self.df.shape[1])

            for i in range(self.df.shape[0]):
                for j in range(self.df.shape[1]):
                    if j == 0:  # 첫 번째 열인 경우 int 형식으로 지정
                        item = QTableWidgetItem(str(int(self.df.iloc[i, j])))
                    else:
                        # 나머지 열은 소수점 6자리까지 표시하는 형식으로 지정
                        item = QTableWidgetItem("{:.6f}".format(self.df.iloc[i, j]))
                    self.table.setItem(i, j, item)

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
            self.scatter_plot_window = ScatterPlotWindow(column_data)
            self.scatter_plot_window.show()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
