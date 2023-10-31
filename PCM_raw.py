import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import string


class ExcelUploaderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PCM Data 처리")
        self.root.geometry("800x600")

        self.upload_button = tk.Button(
            root, text="Excel Upload", command=self.load_excel
        )
        self.upload_button.grid(row=0, column=0, padx=10, pady=10)

        self.tree_frame = tk.Frame(root)
        self.tree_frame.grid(row=1, column=0, padx=10, pady=10)

    def load_excel(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx;*.xls")]
        )

        if file_path:
            try:
                df = pd.read_excel(file_path)
                self.add_alphabet_column(df)
                self.display_data_in_new_window(df)
            except Exception as e:
                self.display_data_in_new_window(f"파일을 열 수 없습니다: {str(e)}")

    def add_alphabet_column(self, data):
        # 알파벳 순으로 값 생성
        alphabet = string.ascii_uppercase  # A부터 Z까지의 알파벳
        column_values = []

        current_letter = "A"  # 초기 알파벳
        for i in range(len(data))
                if data.iloc[i, 0] == 1:
                    current_letter = chr(ord(current_letter) + 1)  # 다음 알파벳
                column_values.append(current_letter)

        # 새로운 칼럼 추가
        data["New_Column"] = column_values

    def display_data_in_new_window(self, data):
        new_window = tk.Toplevel(self.root)
        new_window.title("데이터 프레임")

        # Treeview의 크기 조절을 위한 프레임 생성
        tree_frame = tk.Frame(new_window)
        tree_frame.pack()

        # Treeview 생성
        tree = ttk.Treeview(
            tree_frame, columns=list(data.columns), show="headings", height=10
        )

        for col in data.columns:
            tree.heading(col, text=col)
            tree.column(col, width=100)

        # 수평 스크롤바 추가
        tree_scrollbar_y = ttk.Scrollbar(
            tree_frame, orient="vertical", command=tree.yview
        )
        tree_scrollbar_x = ttk.Scrollbar(
            tree_frame, orient="horizontal", command=tree.xview
        )
        tree.configure(
            yscrollcommand=tree_scrollbar_y.set, xscrollcommand=tree_scrollbar_x.set
        )
        tree_scrollbar_y.pack(side="right", fill="y")
        tree_scrollbar_x.pack(side="bottom", fill="x")

        for index, row in data.iterrows():
            values = [row[col] for col in data.columns]
            tree.insert("", "end", values=values)

        # Treeview 크기 조절
        tree.pack(fill="both", expand=True)


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelUploaderApp(root)
    root.mainloop()
