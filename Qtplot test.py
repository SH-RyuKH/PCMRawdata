import sys
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
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
    QSpacerItem,
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

        # 데이터를 유지하기 위한 멤버 변수 추가
        self.column_data = column_data
        self.selected_column_name = selected_column_name

        # raw_data_window 멤버 변수 추가
        self.raw_data_window = raw_data_window
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
        self.delete_data_button.setFixedSize(150, 30)
        self.delete_data_button.clicked.connect(self.delete_selected_data)
        layout.addWidget(self.delete_data_button)

        # "데이터 적용" 버튼 추가
        self.update_data_button = QPushButton("데이터 적용", self)
        self.update_data_button.setFixedSize(150, 30)
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

        # 데이터 업데이트
        self.column_data = updated_y_values
        self.selected_column_name = selected_column_name

        # MainWindow에 데이터 업데이트 알리기
        self.dataUpdated.emit(selected_column_name, updated_y_values)

        self.close()

    def handle_item_double_clicked(self, item):
        # 데이터 목록에서 더블 클릭한 항목의 인덱스 가져오기
        selected_index = self.data_list_widget.row(item)

        # 선택한 데이터의 좌표 표시
        x_value = self.data["x"][selected_index]
        y_value = self.data["y"][selected_index]
        self.selected_coordinates_label.setText(
            f"선택한 데이터: X: {x_value:.15f}  Y: {y_value:.15f}"
        )

    def closeEvent(self, event):
        # 창이 닫힐 때 데이터를 저장
        self.save_data()

    def save_data(self):
        # 데이터를 ScatterPlotWindow의 멤버 변수에 저장
        self.raw_data_window.set_scatter_plot_data(
            self.selected_column_name, self.column_data
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
            f"선택한 데이터: X: {x_value:.15f}  Y: {y_value:.15f}"
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
            item = QListWidgetItem(f"X: {x:.15f}  Y: {y:.15f}")
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

    def get_saved_data(self):
        # ScatterPlotWindow의 데이터를 가져오는 메서드
        return self.selected_column_name, self.column_data


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("PCM Raw Data")
        self.setGeometry(100, 100, 1600, 1200)

        # 좌측 영역
        self.central_widget = QWidget(self)
        self.setCentralWidget(self.central_widget)

        # 전체 수직 레이아웃
        self.layout = QVBoxLayout(self.central_widget)

        Excel_layout = QHBoxLayout()
        button_layout = QHBoxLayout()

        # 버튼 및 테이블 등을 추가
        self.label = QLabel("선택된 파일 없음", self.central_widget)
        self.upload_button = QPushButton("Excel 업로드", self.central_widget)
        self.upload_button.setFixedSize(100, 30)

        self.chart_button = QPushButton("이상치 제거", self.central_widget)
        self.chart_button.setFixedSize(100, 30)
        # ComboBox 생성 및 초기화
        self.item_column_combobox = QComboBox(self.central_widget)
        self.item_column_combobox.addItem("선택된 열 없음")
        self.data_chart_button = QPushButton("차트", self.central_widget)
        self.data_chart_button.setFixedSize(100, 30)

        # 버튼에 함수 연결
        self.upload_button.clicked.connect(self.upload_excel)
        self.chart_button.clicked.connect(self.show_scatter_plot)
        self.data_chart_button.clicked.connect(self.avg_scatter_plot)
        self.table = QTableWidget(self.central_widget)
        self.table_avg_std = QTableWidget(self.central_widget)

        Excel_layout.addWidget(self.upload_button)
        Excel_layout.addWidget(self.label)

        button_layout.addWidget(self.chart_button)
        button_layout.addWidget(self.item_column_combobox)
        button_layout.addWidget(self.data_chart_button)

        # 수평 스페이서 추가
        spacer_item = QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum)
        button_layout.addItem(spacer_item)

        # 전체 레이아웃에 나머지 위젯들 추가
        self.layout.addLayout(Excel_layout)
        self.layout.addLayout(button_layout)
        self.layout.addWidget(self.table)
        self.layout.addWidget(self.table_avg_std)

        self.filename = None
        self.df = None
        self.scatter_plot_window = None

    def process_excel_data(self, excel_file_path):
        # Read Excel file
        df = pd.read_excel(excel_file_path, header=None)

        # 전처리 시작
        column_names = []
        data_values = []
        static_values = ["Avg", "Std", "Min", "Max"]
        wafer_id = None
        start_processing = False
        unique_second_values = set()
        wafer_no = 0

        # 행 별로 조회
        for index, row in df.iterrows():
            # 처음 Wafer ID가 나오면 처리 시작
            if "Wafer ID:" in str(row[0]):
                wafer_id = row[0]
                start_processing = True
                wafer_no += 1
            elif "Lot Summary" in str(row[0]):
                break
            elif start_processing and pd.isna(row[0]):
                # Wafer ID가 나온 이후에 다음 Wafer ID 전에 공백이 있다면 무시
                continue
            elif start_processing:
                # 처음에는 40자를 자르고, 이후부터는 11자씩 자르고 값 추출
                if len(str(row[0])) > 40:
                    values = [str(row[0])[:40].strip()] + [
                        value.strip()
                        for value in [
                            str(row[0])[i : i + 11]
                            for i in range(40, len(str(row[0])), 11)
                        ]
                        if value
                    ]
                else:
                    values = [str(row[0]).strip()]

                # 40자를 자른 부분이 모두 공백이면 중단
                if not values[0]:
                    break

                # 나머지 행은 데이터 값으로 사용
                data_values.append([wafer_id] + values)

        # column_names에 Wafer ID 추가
        for i in data_values:
            if i[1] not in unique_second_values:
                unique_second_values.add(i[1])
                column_names.append(i[1])

        # 데이터프레임 생성
        data_df = pd.DataFrame(data_values)
        data_df = data_df.transpose()

        # 데이터프레임을 잘라서 리스트에 저장
        num_columns = len(column_names)
        total_columns = wafer_no * num_columns
        result_dfs = []

        for i in range(0, total_columns, num_columns):
            subset_df = data_df.iloc[:, i : i + num_columns]
            subset_df.columns = column_names  # 각 데이터프레임의 열 이름 설정
            subset_df = subset_df.applymap(
                lambda x: float(x) if isinstance(x, (int, float)) else x
            )

            result_dfs.append(subset_df)

        # 리스트에 있는 데이터프레임들을 하나로 합치기
        result_df = pd.concat(result_dfs, ignore_index=True)

        # 행 별로 조회하면서 Column명과 동일한 행 제거
        result_df = result_df[
            ~result_df.apply(lambda row: any(row == column_names), axis=1)
        ]
        for i in range(0, 4):
            result_df = result_df[
                ~result_df.apply(lambda row: any(row == static_values[i]), axis=1)
            ]

        column_names.insert(0, "Wafer ID")
        # 행 별로 조회하면서 Slot 다음에 나오는 숫자를 다음 Slot이 나올 때까지 Wafer ID에 추가
        current_wafer_id = None
        for index, row in result_df.iterrows():
            if "Wafer ID:" in str(row[0]):
                current_wafer_id = str(row[0]).split()[-1]
                result_df = result_df.drop(index)

            elif current_wafer_id:
                result_df.at[index, "Wafer ID"] = current_wafer_id

        result_df.insert(0, "Wafer ID", result_df.pop("Wafer ID"))
        return result_df

    def upload_excel(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly

        def convert_to_float(value):
            try:
                return float(value)
            except ValueError:
                return value

        file_dialog = QFileDialog(self)
        file_dialog.setNameFilter("Excel 파일 (*.xlsx *.xls)")
        file_dialog.setFileMode(QFileDialog.ExistingFile)
        file_dialog.setOptions(options)

        if file_dialog.exec_():
            processed_filename = file_dialog.selectedFiles()[0]
            self.label.setText(f"선택된 파일: {self.filename}")

            try:
                processed_df = self.process_excel_data(processed_filename)
                self.df = processed_df

            except Exception as e:
                print(f"Error reading Excel file: {e}")

            if self.df is not None:
                self.table.setRowCount(self.df.shape[0])
                self.table.setColumnCount(self.df.shape[1])
                self.update_table()

                # ComboBox에 "item" 이후의 열 추가
                self.item_column_combobox.clear()
                self.item_column_combobox.addItem("선택된 열 없음")
                for col in self.df.columns[2:]:
                    if col == "item":
                        continue
                    self.item_column_combobox.addItem(col)

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
        # 테이블의 값을 새로운 데이터프레임으로 옮기기
        upload_df = pd.DataFrame(
            index=range(self.table.rowCount()), columns=self.df.columns
        )

        for i in range(self.table.rowCount()):
            for j in range(self.table.columnCount()):
                item = self.table.item(i, j)
                if item is not None:
                    upload_df.iloc[i, j] = self.get_original_value(item)

        self.df = upload_df
        self.show_avg_std()

    def get_original_value(self, item):
        # 이 함수는 QTableWidgetItem에서 원래 값으로 변환하는데 사용될 수 있습니다.
        # 예: 부동 소수점 값이나 다른 형태의 값을 가져오는 데 사용될 수 있습니다.
        # 실제로 사용하는 데이터 형식에 따라 구현이 달라질 수 있습니다.
        return item.text()

    def get_formatted_value(self, value):
        try:
            # 문자열을 부동 소수점으로 변환 시도
            float_value = float(value)
            return "{:.15f}".format(float_value)
        except ValueError:
            # 변환할 수 없는 경우는 그대로 반환
            return str(value)

    def show_scatter_plot(self):
        selected_item = self.table.currentItem()

        if selected_item is not None:
            selected_row = selected_item.row()
            selected_col = selected_item.column()
            selected_column_name = self.df.columns[selected_col]  # 선택한 칼럼의 이름 가져오기

            column_data = self.df.iloc[:, selected_col].tolist()

        if self.scatter_plot_window is not None:
            # ScatterPlotWindow가 이미 열려있으면 창을 닫고 데이터 로드
            self.scatter_plot_window.close()
            (
                selected_column_name,
                column_data,
            ) = self.scatter_plot_window.get_saved_data()

        # ScatterPlotWindow 생성
        self.scatter_plot_window = ScatterPlotWindow(
            column_data, self, selected_column_name
        )
        self.scatter_plot_window.dataUpdated.connect(self.handle_scatter_data_updated)
        self.scatter_plot_window.show()

    def set_scatter_plot_data(self, selected_column_name, column_data):
        # ScatterPlotWindow의 데이터를 저장하는 메서드
        self.scatter_plot_data = {
            "selected_column_name": selected_column_name,
            "column_data": column_data,
        }

    def get_scatter_plot_data(self):
        # ScatterPlotWindow의 데이터를 가져오는 메서드
        return (
            self.scatter_plot_data["selected_column_name"],
            self.scatter_plot_data["column_data"],
        )

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
        selected_col_index = 0  # "공정 조건"이 위치한 Column (두 번째 Column)

        # "공정 조건" Column의 고유한 값들
        process_conditions = self.df.iloc[:, selected_col_index].unique()

        avg_std_values = []

        # 3번째 Column부터 각 Column에 대한 AVG와 STD 계산
        for col_index in range(2, self.df.shape[1]):
            col_header = self.df.columns[col_index]

            # "item" 칼럼명인 경우 계산 넘어가기
            if col_header.lower() == "Item Name":
                continue
            if col_header.lower() == "Wafer ID":
                continue
            self.df.iloc[:, col_index] = pd.to_numeric(self.df.iloc[:, col_index])

            # 이미 숫자로 변환된 경우에만 계산 수행

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
                    avg_std_values.append([col_header, process_condition, "AVG", avg])
                    avg_std_values.append([col_header, process_condition, "STD", std])

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
                    row, col * 2, QTableWidgetItem("{:.15f}".format(avg_value))
                )
                self.table_avg_std.setItem(
                    row, col * 2 + 1, QTableWidgetItem("{:.15f}".format(std_value))
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
            item = QTableWidgetItem("{:.15f}".format(y_value))
            self.table.setItem(row, selected_col_index, item)

        # 데이터프레임(df) 업데이트
        for row, y_value in enumerate(updated_y_values):
            self.df.iloc[row, selected_col_index] = y_value

        # 테이블 크기 조정
        self.table.resizeColumnsToContents()

        # AVG, STD 업데이트
        self.show_avg_std()

    def handle_combobox_changed(self, index):
        # ComboBox 선택이 바뀌었을 때의 처리
        selected_column = self.item_column_combobox.currentText()
        if selected_column != "선택된 열 없음":
            # 선택한 열의 차트 표시
            self.show_scatter_plot()

    def avg_scatter_plot(self):
        selected_column = self.item_column_combobox.currentText()

        # 선택한 열이 "_AVG"로 끝나는 경우에만 처리
        avg_column_data = self.get_avg_column_data(selected_column)

        # AVG 데이터가 있는 경우에만 차트 표시
        if avg_column_data is not None:
            avg_data = {
                "Slot": self.df["Wafer ID"].unique().tolist(),  # "공정조건"의 값들을 가져옴
                "AVG": avg_column_data,
            }

            # 선택한 칼럼의 AVG를 matplotlib로 차트 그리기
            self.plot_avg_scatter_chart(selected_column, avg_data)

    def plot_avg_scatter_chart(self, selected_column, avg_data):
        x_values = avg_data["Slot"]

        # 그래프를 그릴 Figure 생성
        figure, ax = plt.subplots(figsize=(8, 6))

        # 알파벳 - AVG 차트 그리기
        ax.scatter(x_values, avg_data["AVG"], marker="o", s=100)

        ax.set_xticks(x_values)
        ax.set_xticklabels(avg_data["Slot"])
        ax.set_xlabel("공정 조건")
        ax.set_ylabel(selected_column + " AVG")
        ax.set_title(f"{selected_column} - AVG Chart")

        # 차트를 보여주기
        plt.show()

    def get_avg_column_data(self, selected_column):
        # 선택한 열의 "_AVG"에 해당하는 데이터를 가져옴
        avg_column_data = []

        # 수평 헤더 아이템 중에서 "_AVG"로 끝나는 아이템을 찾음
        for col in range(self.table_avg_std.columnCount()):
            header_item = self.table_avg_std.horizontalHeaderItem(col)
            if header_item.text().endswith("_AVG"):
                # 선택한 열과 일치하는 경우
                if header_item.text().startswith(selected_column):
                    # 해당 열의 데이터를 가져와서 리스트에 추가
                    for row in range(self.table_avg_std.rowCount()):
                        item = self.table_avg_std.item(row, col)
                        if item is not None:
                            avg_column_data.append(float(item.text()))

        return avg_column_data


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
