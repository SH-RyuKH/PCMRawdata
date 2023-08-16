import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


class ExcelChartApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Chart App")
        self.data = pd.DataFrame()  # DataFrame으로 데이터 저장
        self.selected_index = None  # 선택한 데이터 인덱스 저장
        self.create_widgets()

    def create_widgets(self):
        self.textarea = tk.Text(self.root, height=10, width=50)
        self.textarea.pack(padx=10, pady=10)

        self.load_button = tk.Button(
            self.root, text="Load Excel File", command=self.load_data
        )
        self.load_button.pack(pady=5)

        self.column_listbox = tk.Listbox(self.root, height=5, selectmode=tk.SINGLE)
        self.column_listbox.pack(pady=5)

        # 삭제 버튼 생성
        self.delete_button = tk.Button(
            self.root, text="Delete Selected Data", command=self.delete_data
        )
        self.delete_button.pack(pady=5)

        self.chart_label = tk.Label(self.root, text="Chart")
        self.chart_label.pack(pady=10)

        # 그래프를 표시할 Figure와 Axes 객체 생성
        self.figure, self.ax = plt.subplots(figsize=(8, 4))

        # FigureCanvasTkAgg 객체 생성하여 그래프를 Tkinter 창에 추가
        self.canvas = FigureCanvasTkAgg(self.figure, master=self.root)
        self.canvas.get_tk_widget().pack()

        # 그래프에서 점을 선택하는 이벤트 핸들러 등록
        self.selected_index = None

        # 선택한 데이터 값을 표시할 Label 위젯
        self.selected_value_label = tk.Label(self.root, text="")
        self.selected_value_label.pack(pady=5)

        # 선택한 데이터의 y값을 표시할 Textbox 위젯
        self.selected_y_value_textbox = tk.Entry(self.root, state="readonly")
        self.selected_y_value_textbox.pack(pady=5)

        # 선택한 데이터 목록을 보여줄 Listbox 위젯
        self.selected_data_listbox = tk.Listbox(self.root, height=5)
        self.selected_data_listbox.pack(pady=5)

        # 칼럼 선택 시 그래프 업데이트 이벤트 등록
        self.column_listbox.bind("<<ListboxSelect>>", self.plot_selected_column)

        # 데이터 다운로드 버튼 생성
        self.download_button = tk.Button(
            self.root, text="Download Data as Excel", command=self.download_data
        )
        self.download_button.pack(pady=5)

    def load_data(self):
        # 엑셀 파일 불러오기 대화 상자를 통해 파일 선택
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            try:
                # 엑셀 파일 데이터 읽기
                self.data = pd.read_excel(file_path)

                # Listbox에 칼럼 이름들을 추가
                self.column_listbox.delete(0, tk.END)
                for column in self.data.columns:
                    self.column_listbox.insert(tk.END, column)

                self.update_textarea()

            except Exception as e:
                messagebox.showerror("Error", f"Failed to load data: {e}")

    def update_textarea(self):
        # 리스트박스에서 선택한 항목 가져오기
        selected_indices = self.column_listbox.curselection()
        if not selected_indices:  # 만약 리스트박스에 선택한 항목이 없을 경우
            self.textarea.delete(1.0, tk.END)  # 텍스트 영역을 모두 지우고 함수 종료
            return

        selected_index = int(selected_indices[0])  # 첫 번째 선택한 인덱스 가져오기
        selected_column = self.column_listbox.get(selected_index)

        # 선택된 칼럼의 데이터를 텍스트 영역에 표시
        self.textarea.delete(1.0, tk.END)
        self.textarea.insert(tk.END, self.data[selected_column].to_string(index=False))

    def plot_selected_column(self, event=None):
        selected_indices = self.column_listbox.curselection()
        if not selected_indices:  # 리스트박스에서 선택된 항목이 없을 경우
            return

        selected_index = int(selected_indices[0])  # 첫 번째 선택된 인덱스 가져오기
        selected_column = self.column_listbox.get(selected_index)
        try:
            # 선택된 칼럼의 데이터를 읽어와서 리스트에 저장
            column_data = self.data[selected_column].tolist()

            # 차트를 그리기 위해 데이터 시각화
            self.ax.clear()
            self.ax.plot(column_data, "o")  # 'o' 스타일로 점만 표시
            self.ax.set_xlabel("Data Index")
            self.ax.set_ylabel("Data Value")
            self.ax.set_title(f"{selected_column} Chart")
            self.canvas.draw()

            # 그래프에서 점을 선택하는 이벤트 핸들러 등록
            self.canvas.mpl_connect("button_press_event", self.on_click)  # 여기서 변경

            # 선택한 칼럼의 데이터 목록을 Listbox에 표시
            self.selected_data_list = self.data[selected_column].tolist()
            self.selected_data_listbox.delete(0, tk.END)
            for data in self.selected_data_list:
                self.selected_data_listbox.insert(tk.END, data)

        except ValueError:
            messagebox.showerror("Error", "잘못된 데이터 형식입니다. 숫자 값을 입력하세요.")

    def on_hover(self, event):
        if event.xdata is not None and event.ydata is not None:
            # 그래프의 점 위에 마우스를 올린 경우
            x, y = event.xdata, event.ydata
            self.ax.plot(x, y, "ro", markersize=15, alpha=0.7)
            self.canvas.draw()

    def on_leave(self, event):
        # 마우스가 그래프 영역을 벗어난 경우
        self.ax.clear()
        self.plot_selected_column()  # 선택한 칼럼의 데이터 다시 그리기
        self.canvas.draw()

    def on_click(self, event):
        if event.xdata is not None and event.ydata is not None:
            # 그래프의 점을 클릭한 경우
            x, y = event.xdata, event.ydata
            selected_column = self.column_listbox.get(
                self.column_listbox.curselection()
            )

            if not selected_column:
                messagebox.showerror("Error", "Please select a column to plot.")
                return

            # 선택한 데이터 값을 가져오기 위해 라벨과 텍스트박스 업데이트
            selected_value = self.data[selected_column].iloc[int(x)]
            self.selected_value_label.config(text=f"Selected Value: {selected_value}")
            self.selected_y_value_textbox.config(state="normal")
            self.selected_y_value_textbox.delete(0, tk.END)
            self.selected_y_value_textbox.insert(0, str(selected_value))
            self.selected_y_value_textbox.config(state="readonly")

            # 선택한 데이터 인덱스 저장
            self.selected_index = int(x)

            # "Delete Selected Data" 버튼 기능 변경
            self.delete_button.config(
                text="Delete Selected Data", command=self.delete_data
            )

        else:
            # 그래프 영역을 클릭한 경우
            # 선택한 데이터 값을 초기화하고 "Delete Selected Data" 버튼 기능 변경
            self.selected_index = None
            self.selected_value_label.config(text="")
            self.selected_y_value_textbox.config(state="normal")
            self.selected_y_value_textbox.delete(0, tk.END)
            self.selected_y_value_textbox.config(state="readonly")
            self.delete_button.config(
                text="Plot Chart", command=self.plot_selected_column
            )

    def delete_data(self):
        selected_column_index = self.column_listbox.curselection()
        if selected_column_index:
            selected_column_index = selected_column_index[0]  # 첫 번째 선택된 항목 인덱스 가져오기
            selected_column = self.column_listbox.get(selected_column_index)
            if not selected_column:
                messagebox.showerror(
                    "Error", "Please select a column to delete data from."
                )
                return

            try:
                if self.selected_index is not None:
                    # 선택한 인덱스의 해당 칼럼 데이터만 삭제
                    self.data.loc[self.selected_index, selected_column] = None

                    # 변경된 데이터로 차트 업데이트
                    self.ax.clear()
                    self.ax.plot(
                        self.data[selected_column].tolist(), "o"
                    )  # 'o' 스타일로 점만 표시
                    self.ax.set_xlabel("Data Index")
                    self.ax.set_ylabel("Data Value")
                    self.ax.set_title(f"{selected_column} Chart")
                    self.canvas.draw()

                    # 선택한 데이터 값 라벨 초기화
                    self.selected_value_label.config(text="")
                    self.selected_y_value_textbox.config(state="normal")
                    self.selected_y_value_textbox.delete(0, tk.END)
                    self.selected_y_value_textbox.config(state="readonly")

                    # 선택한 데이터 목록 초기화
                    self.selected_data_list = []
                    self.selected_data_listbox.delete(0, tk.END)

                    self.selected_index = None

                    # "Delete Selected Data" 버튼을 "Plot Chart" 버튼으로 변경
                    self.delete_button.config(
                        text="Select Data", command=self.plot_selected_column
                    )

                else:
                    messagebox.showerror(
                        "Error", "No data selected. Please select data to delete."
                    )

            except tk.TclError:
                messagebox.showerror("Error", "An error occurred while deleting data.")

    def download_data(self):
        # 엑셀 파일로 데이터 다운로드
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")]
        )
        if file_path:
            try:
                self.data.to_excel(file_path, index=False)
                messagebox.showinfo("Success", "Data downloaded as Excel successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to download data: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelChartApp(root)
    root.mainloop()
