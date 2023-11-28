import tkinter as tk
from tkinter import filedialog
import pandas as pd
from pandastable import Table, TableModel
import statistics


class ExcelViewerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Viewer")

        self.file_path = None
        self.df = None

        # 업로드 버튼 생성
        self.upload_button = tk.Button(
            root, text="Excel 파일 업로드", command=self.upload_file
        )
        self.upload_button.pack(pady=20)

        # 데이터 표시 프레임 생성
        self.data_frame = tk.Frame(root)
        self.data_frame.pack()

        # 새로운 창을 열기 위한 버튼 생성
        self.process_button = tk.Button(
            root, text="데이터 처리 및 통계 보기", command=self.process_and_show_stats
        )
        self.process_button.pack(pady=10)

    def upload_file(self):
        self.file_path = filedialog.askopenfilename(
            filetypes=[("Excel 파일", "*.xls;*.xlsx")]
        )

        if self.file_path:
            # Excel 파일 읽기
            self.df = pd.read_excel(self.file_path)

            # 데이터 가공 및 표시
            self.process_and_display_data()

    def process_and_display_data(self):
        # 열 이름을 기준으로 그룹을 생성하고 알파벳 순으로 레이블 부여
        self.df["Group"] = (self.df.iloc[:, 0] == 1).cumsum()
        self.df["Group"] = self.df["Group"].apply(lambda x: chr(x + ord("A") - 1))

        # 그룹 열을 데이터프레임의 첫 번째 열로 이동
        cols = list(self.df.columns)
        cols = [cols[-1]] + cols[:-1]
        self.df = self.df[cols]

        # 데이터 표시 프레임 초기화
        for widget in self.data_frame.winfo_children():
            widget.destroy()

        # PandasTable을 사용하여 데이터 표시
        pt = Table(
            self.data_frame, dataframe=self.df, showtoolbar=False, showstatusbar=False
        )

        # 헤더 아이콘 제거
        pt.iconcolheader = None
        pt.iconcol = None
        pt.iconbigger = None
        pt.autoAddColumns = False

        pt.show()

    def process_and_show_stats(self):
        if self.df is None:
            return

        # 알파벳 별로 묶은 데이터 그룹들의 AVG(평균)와 STD(표준편차) 계산
        result_df = self.df.groupby("Group").agg(["mean", "std"])

        # 'A' 그룹을 앞으로 옮기기
        # 다른 그룹을 옮기려면 해당 그룹의 이름으로 바꿔주세요
        result_df = result_df.reorder_levels([1, 0], axis=1)

        # 새로운 창을 열어서 결과 표시
        stats_window = tk.Toplevel(self.root)
        stats_window.title("통계 결과")

        stats_frame = tk.Frame(stats_window)
        stats_frame.pack()

        # PandasTable을 사용하여 결과 표시
        stats_pt = Table(
            stats_frame, dataframe=result_df, showtoolbar=False, showstatusbar=False
        )

        # 헤더 아이콘 제거
        stats_pt.iconcolheader = None
        stats_pt.iconcol = None
        stats_pt.iconbigger = None
        stats_pt.autoAddColumns = False

        stats_pt.show()


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelViewerApp(root)
    root.mainloop()
