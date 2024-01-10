import sys
import pandas as pd
import numpy as np
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
    QComboBox,
)
from PyQt5.QtGui import QPainter
from PyQt5.QtCore import pyqtSignal
from PyQt5.QtWidgets import (
    QSplitter,
    QSizePolicy,
    QVBoxLayout,
    QPushButton,
    QDockWidget,
    QGridLayout,
    QHBoxLayout,
)


import pyqtgraph as pg


class ScatterPlotWindow(QWidget):
    dataUpdated = pyqtSignal(str, list)

    def __init__(self, column_data, raw_data_window, selected_column_name):
        super().__init__()

        self.setWindowTitle("이상치 제거")
        self.setGeometry(1700, 100, 1200, 1200)
        self.selected_column_name = selected_column_name

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

        # 선택한 칼럼의 이름을 표시할 레이블 추가
        self.selected_column_label = QLabel(
            f"선택한 칼럼: {self.selected_column_name}", self
        )
        layout.addWidget(self.selected_column_label)

        # 선택된 좌표의 X 및 Y 값을 표시할 레이블 추가
        self.selected_coordinates_label = QLabel("선택한 데이터: X: ?  Y: ?", self)
        layout.addWidget(self.selected_coordinates_label)

        # 모든 데이터의 목록을 표시할 목록 위젯 추가
        self.data_list_widget = QListWidget(self)
        self.data_list_widget.itemDoubleClicked.connect(self.handle_item_double_clicked)
        layout.addWidget(self.data_list_widget)

        # "데이터 삭제" 버튼 추가
        self.delete_data_button = QPushButton("데이터 삭제", self)
        self.delete_data_button.clicked.connect(self.delete_selected_data)
        layout.addWidget(self.delete_data_button)

        # "데이터 적용" 버튼 추가
        self.update_data_button = QPushButton("데이터 적용", self)
        self.update_data_button.clicked.connect(
            lambda: self.update_data_scatterPlot(
                selected_column_name, column_data, raw_data_window
            )
        )
        layout.addWidget(self.update_data_button)

        # 창 설정
        self.data = {"x": list(range(len(column_data))), "y": column_data}

        self.selected_point_index = None
        self.update_data_list()

    def update_data_scatterPlot(
        self, selected_column_name, column_data, raw_data_window
    ):
        # 선택한 칼럼의 이름 및 현재 Y값 가져오기
        selected_column_name = selected_column_name
        updated_y_values = column_data

        self.dataUpdated.emit(selected_column_name, updated_y_values)

    def handle_item_double_clicked(self, item):
        # 데이터 목록에서 더블 클릭한 항목의 인덱스 가져오기
        selected_index = self.data_list_widget.row(item)

        # 선택한 데이터의 좌표 표시
        x_value = self.data["x"][selected_index]
        y_value = self.data["y"][selected_index]
        self.selected_coordinates_label.setText(
            f"선택한 데이터: X: {x_value:.6f}  Y: {y_value:.6f}"
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
            f"선택한 데이터: X: {x_value:.6f}  Y: {y_value:.6f}"
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
            item = QListWidgetItem(f"X: {x:.6f}  Y: {y:.6f}")
            self.data_list_widget.addItem(item)

    def delete_selected_data(self):
        if (
            self.selected_point_index is not None
            or self.data_list_widget.currentItem() is not None
        ):
            # 선택한 데이터의 y 값을 0으로 만들기
            if self.selected_point_index is not None:
                selected_y_value = self.data["y"][self.selected_point_index]
                self.data["y"][self.selected_point_index] = 0
            elif self.data_list_widget.currentItem() is not None:
                # 데이터 목록에서 선택한 항목의 인덱스 가져오기
                selected_index = self.data_list_widget.currentRow()
                selected_y_value = self.data["y"][selected_index]
                self.data["y"][selected_index] = 0

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

<<<<<<< HEAD
        # 좌측 영역
        self.central_widget = QWidget(self)
        self.setCentralWidget(self.central_widget)

        self.layout = QVBoxLayout(self.central_widget)
=======
        # 버튼 크기 설정
        button_width = 100
        button_height = 30
>>>>>>> 29caa351e28a63d0038263fcf6f55bb5e466b7e3

        # 버튼 및 테이블 등을 추가
        self.label = QLabel("선택된 파일 없음", self.central_widget)
        self.upload_button = QPushButton("Excel 업로드", self.central_widget)
        self.chart_button = QPushButton("이상치 제거", self.central_widget)

<<<<<<< HEAD
        # 버튼에 함수 연결
=======
        # 좌측 영역에 현재 내용 보여주는 위젯 추가
        self.left_widget = QWidget(self)
        self.left_layout = QVBoxLayout(self.left_widget)
        self.label = QLabel("선택된 파일 없음", self.left_widget)
        self.left_layout.addWidget(self.label)

        # 버튼 크기를 조정할 QSizePolicy 설정
        button_size_policy = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)

        # Central Widget을 Splitter로 변경
        self.central_splitter = QSplitter(Qt.Horizontal, self)
        self.setCentralWidget(self.central_splitter)

        # 좌측 영역에 현재 내용 보여주는 위젯 추가
        self.left_widget = QWidget(self)
        self.left_layout = QVBoxLayout(self.left_widget)
        self.label = QLabel("선택된 파일 없음", self.left_widget)
        self.left_layout.addWidget(self.label)

        # 버튼 크기를 조정할 QSizePolicy 설정
        button_size_policy = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)

        # 수평 박스 레이아웃 생성
        button_layout = QHBoxLayout()

        # Excel 업로드 버튼 추가
        self.upload_button = QPushButton("Excel 업로드", self.left_widget)
        self.upload_button.setSizePolicy(button_size_policy)
        self.upload_button.setFixedSize(button_width, button_height)
>>>>>>> 29caa351e28a63d0038263fcf6f55bb5e466b7e3
        self.upload_button.clicked.connect(self.upload_excel)
        self.chart_button.clicked.connect(self.show_scatter_plot)
        self.table = QTableWidget(self.central_widget)
        self.table_avg_std = QTableWidget(self.central_widget)

<<<<<<< HEAD
        self.layout.addWidget(self.label)
        self.layout.addWidget(self.upload_button)
        self.layout.addWidget(self.chart_button)
        self.layout.addWidget(self.table)
        self.layout.addWidget(self.table_avg_std)
=======
        # AVG / STD 버튼 추가
        self.avg_std_button = QPushButton("AVG / STD", self.left_widget)
        self.avg_std_button.setSizePolicy(button_size_policy)
        self.avg_std_button.setFixedSize(button_width, button_height)
        self.avg_std_button.clicked.connect(self.show_avg_std)
        button_layout.addWidget(self.avg_std_button)

        # 수평 박스 레이아웃을 왼쪽 위젯에 추가
        self.left_layout.addLayout(button_layout)

        # 위쪽에 테이블 추가
        self.table = QTableWidget(self.left_widget)
        self.left_layout.addWidget(self.table)

        # 아래쪽에 테이블 추가
        self.table_avg_std = QTableWidget(self.left_widget)
        self.left_layout.addWidget(self.table_avg_std)

        # 첫번째
        self.central_splitter_2 = QSplitter(Qt.Vertical, self.central_splitter)
        self.central_splitter.addWidget(self.left_widget)
        self.central_splitter.addWidget(self.central_splitter_2)

        self.middle_widget = QWidget(self)
        # 가운데 분할 화면에 추가할 QComboBox 초기화
        self.column_combobox_avg_std = QComboBox(self.central_splitter_2)
        self.column_combobox_avg_std.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.column_combobox_avg_std.setFixedSize(button_width, button_height)
        self.central_splitter_2.addWidget(self.column_combobox_avg_std)

        # 오른쪽 영역에 어디에 추가해야 할지를 보여주는 위젯 추가
        self.right_widget = QWidget(self)
        self.right_layout = QVBoxLayout(self.right_widget)
        # Scatter Plot 창을 QDockWidget으로 만들고 오른쪽에 추가
        self.scatter_plot_window = ScatterPlotWindow([], self)
        self.scatter_plot_dock = QDockWidget("Scatter Plot", self)
        self.scatter_plot_dock.setWidget(self.scatter_plot_window)
        self.addDockWidget(Qt.RightDockWidgetArea, self.scatter_plot_dock)

        self.central_splitter.addWidget(self.right_widget)
>>>>>>> 29caa351e28a63d0038263fcf6f55bb5e466b7e3

        self.filename = None
        self.df = None
        self.scatter_plot_window = None

<<<<<<< HEAD
=======
        self.set_splitter_sizes()

    def set_splitter_sizes(self):
        # Set initial sizes for the splitters
        self.central_splitter.setSizes(
            [self.width() // 3, self.width() // 3, self.width() // 3]
        )
        self.central_splitter_2.setSizes(
            [
                self.height() // 2,
                self.height() // 2,
            ]
        )

>>>>>>> 29caa351e28a63d0038263fcf6f55bb5e466b7e3
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
                # "공정조건"이라는 column 추가
                self.df["공정조건"] = self.create_process_column(self.df.iloc[:, 0])

                # "공정조건" 칼럼을 두 번째 칼럼으로 이동
                cols = list(self.df.columns)
                cols = [cols[0], "공정조건"] + cols[1:-1] + [cols[-1]]
                self.df = self.df[cols]

                # 마지막 칼럼 제거
                self.df = self.df.iloc[:, :-1]

                self.table.setRowCount(self.df.shape[0])
                self.table.setColumnCount(self.df.shape[1])
                self.update_table()

        self.show_avg_std()

    def update_table(self):
        self.table.clear()

        if self.df is not None:
            # 컬럼 헤더 추가
            for j in range(self.df.shape[1]):
                header_item = QTableWidgetItem(self.df.columns[j])
                self.table.setHorizontalHeaderItem(j, header_item)

            for i in range(self.df.shape[0]):
                for j in range(self.df.shape[1]):
                    value = self.df.iloc[i, j]
                    item = QTableWidgetItem(self.get_formatted_value(value))
                    self.table.setItem(i, j, item)

        self.table.resizeColumnsToContents()

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

    def show_scatter_plot(self):
        selected_item = self.table.currentItem()

        if selected_item is not None:
            selected_row = selected_item.row()
            selected_col = selected_item.column()
            selected_column_name = self.df.columns[selected_col]  # 선택한 칼럼의 이름 가져오기

            column_data = self.df.iloc[:, selected_col].tolist()

            if self.scatter_plot_window is not None:
                # ScatterPlotWindow가 이미 열려있으면 닫기
                self.scatter_plot_window.close()
                self.scatter_plot_window = None

            # ScatterPlotWindow 생성
            self.scatter_plot_window = ScatterPlotWindow(
                column_data, self, selected_column_name
            )
            self.scatter_plot_window.dataUpdated.connect(
                self.handle_scatter_data_updated
            )
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

    def show_avg_std(self):
        # 선택한 "공정 조건"에 대한 AVG와 STD 계산
        selected_col_index = 1  # "공정 조건"이 위치한 Column (두 번째 Column)

        # "공정 조건" Column의 고유한 값들
        process_conditions = self.df.iloc[:, selected_col_index].unique()

        avg_std_values = []

        # 3번째 Column부터 각 Column에 대한 AVG와 STD 계산
        for col_index in range(2, self.df.shape[1]):
            col_header = self.df.columns[col_index]

            # "item" 칼럼명인 경우 계산 넘어가기
            if col_header.lower() == "item":
                continue
            if pd.api.types.is_numeric_dtype(self.df.iloc[:, col_index]):
                for process_condition in process_conditions:
                    # "공정 조건"에 해당하는 데이터 필터링
                    filtered_df = self.df[
                        self.df.iloc[:, selected_col_index] == process_condition
                    ]
                    col_data_filtered = filtered_df.iloc[:, col_index]

                    # nan이 아닌 경우에만 계산
                    if not col_data_filtered.isna().all():
                        avg = col_data_filtered.mean()
                        std = col_data_filtered.std()

                        # 결과 리스트에 정보 추가
                        avg_std_values.append(
                            [col_header, process_condition, "AVG", avg]
                        )
                        avg_std_values.append(
                            [col_header, process_condition, "STD", std]
                        )

        # "table_avg_std" 테이블에 AVG 및 STD 표시
        self.update_avg_std_table(avg_std_values)

    def update_avg_std_table(self, avg_std_values):
        # "table_avg_std" 테이블 초기화
        self.table_avg_std.clear()

        # 고유한 "공정 조건" 및 Column 명 가져오기
        unique_process_conditions = sorted(set(item[1] for item in avg_std_values))
        unique_columns = sorted(set(item[0] for item in avg_std_values))

        # "table_avg_std" 테이블 열 수 설정
        num_columns = len(unique_columns)
        self.table_avg_std.setColumnCount(num_columns * 2)  # AVG와 STD이므로 2배

        # "table_avg_std" 테이블 헤더 설정
        header_labels = []
        for column in unique_columns:
            header_labels.extend([f"{column}_AVG", f"{column}_STD"])
        self.table_avg_std.setHorizontalHeaderLabels(header_labels)

        # "table_avg_std" 테이블 행 수 설정
        num_rows = len(unique_process_conditions)
        self.table_avg_std.setRowCount(num_rows)

        # 데이터 채우기
        for col, column in enumerate(unique_columns):
            for row, process_condition in enumerate(unique_process_conditions):
                avg_value = next(
                    (
                        item[3]
                        for item in avg_std_values
                        if item[0] == column
                        and item[1] == process_condition
                        and item[2] == "AVG"
                    ),
                    np.nan,
                )
                std_value = next(
                    (
                        item[3]
                        for item in avg_std_values
                        if item[0] == column
                        and item[1] == process_condition
                        and item[2] == "STD"
                    ),
                    np.nan,
                )

                self.table_avg_std.setItem(
                    row, col * 2, QTableWidgetItem("{:.2f}".format(avg_value))
                )
                self.table_avg_std.setItem(
                    row, col * 2 + 1, QTableWidgetItem("{:.2f}".format(std_value))
                )

                # "공정 조건" 표시
                process_condition_item = QTableWidgetItem(process_condition)
                self.table_avg_std.setVerticalHeaderItem(row, process_condition_item)

        # "table_avg_std" 테이블 크기 조정
        self.table_avg_std.resizeColumnsToContents()

    def handle_scatter_data_updated(self, selected_column_name, updated_y_values):
        # 선택한 칼럼의 이름 및 현재 Y값 가져오기
        selected_column_name = selected_column_name

        # 선택한 칼럼의 헤더 열 인덱스 가져오기
        selected_col_index = self.df.columns.get_loc(selected_column_name)

        # Y값 업데이트
        for row, y_value in enumerate(updated_y_values):
            item = QTableWidgetItem("{:.6f}".format(y_value))
            self.table.setItem(row, selected_col_index, item)

            # 테이블 크기 조정
            self.table.resizeColumnsToContents()

        self.show_avg_std()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
