import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk


class ExcelUploaderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PCM Data 처리")
        self.root.geometry("800x600")

        self.upload_button = tk.Button(
            root, text="Excel Upload", command=self.load_excel
        )
        self.upload_button.grid(row=0, column=0, padx=10, pady=10)

    def load_excel(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx;*.xls")]
        )

        if file_path:
            try:
                df = pd.read_excel(file_path)
                self.add_alphabet_column(df)

                # 기존의 창을 먼저 닫고
                if hasattr(self, "data_window"):
                    self.data_window.destroy()

                # 새로운 창을 생성하여 데이터 표시
                self.display_data_in_new_window(df)
            except Exception as e:
                self.display_data_in_new_window(f"파일을 열 수 없습니다: {str(e)}")

    def display_data_in_new_window(self, data):
        # 기존의 창이 존재하면 닫기
        if hasattr(self, "data_window"):
            self.data_window.destroy()

        # 새로운 창 생성
        self.data_window = tk.Toplevel(self.root)
        self.data_window.title("데이터 프레임")

        # Treeview 생성
        tree_frame = tk.Frame(self.data_window)
        tree_frame.pack()

        self.tree = ttk.Treeview(
            tree_frame, columns=list(data.columns), show="headings", height=10
        )

        for col in data.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)

        tree_scrollbar_y = ttk.Scrollbar(
            tree_frame, orient="vertical", command=self.tree.yview
        )
        tree_scrollbar_x = ttk.Scrollbar(
            tree_frame, orient="horizontal", command=self.tree.xview
        )
        self.tree.configure(
            yscrollcommand=tree_scrollbar_y.set, xscrollcommand=tree_scrollbar_x.set
        )
        tree_scrollbar_y.pack(side="right", fill="y")
        tree_scrollbar_x.pack(side="bottom", fill="x")

        for index, row in data.iterrows():
            values = [row[col] for col in data.columns]
            self.tree.insert("", "end", text=row.name, values=values)

        # 트리뷰에서 행 클릭 이벤트 바인딩
        self.tree.bind("<ButtonRelease-1>", self.show_selected_row)

        self.tree.pack(fill="both", expand=True)

        # 열 선택 창 열기
        self.open_column_window(data)

    def add_alphabet_column(self, data):
        current_letter = "@"
        column_values = []

        for i in range(len(data)):
            if data.iloc[i, 0] == 1:
                current_letter = chr(ord(current_letter) + 1)
            column_values.append(current_letter)

        data["New_Column"] = column_values

    def open_column_window(self, data):
        # 열 선택 창 생성
        self.column_window = tk.Toplevel(self.root)
        self.column_window.title("열 선택")

        # 열 목록 생성
        columns_label = tk.Label(self.column_window, text="열 목록:")
        columns_label.pack(pady=10)

        columns_listbox = tk.Listbox(self.column_window, selectmode=tk.MULTIPLE)
        for col in data.columns:
            columns_listbox.insert(tk.END, col)

        columns_listbox.pack(pady=10)

        # 확인 버튼 클릭 시 선택한 열 표시
        confirm_button = tk.Button(
            self.column_window,
            text="확인",
            command=lambda: self.show_selected_columns(columns_listbox.get(0, tk.END)),
        )
        confirm_button.pack(pady=10)

    def show_selected_row(self, event):
        # 선택한 행의 데이터 가져오기
        selected_item = self.tree.selection()
        if selected_item:
            selected_row_data = self.tree.item(selected_item, "values")
            self.show_selected_data_in_column_window(selected_row_data)

    def show_selected_data_in_column_window(self, selected_row_data):
        # 선택한 데이터를 열 선택 창에 표시
        if hasattr(self, "column_window"):
            self.column_window.destroy()

        self.column_window = tk.Toplevel(self.root)
        self.column_window.title("열 선택")

        selected_data_label = tk.Label(self.column_window, text="선택한 데이터:")
        selected_data_label.pack(pady=10)

        for i, value in enumerate(selected_row_data):
            label = tk.Label(self.column_window, text=f"Column {i+1}: {value}")
            label.pack()


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelUploaderApp(root)
    root.mainloop()
