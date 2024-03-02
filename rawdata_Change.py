import pandas as pd

# 엑셀 파일 경로 설정
excel_file_path = "sda.xlsx"

# 엑셀 파일 읽기
df = pd.read_excel(excel_file_path, header=None)

# 전처리 시작
column_names = []
data_values = []
<<<<<<< HEAD
static_values = ["Avg", "Std", "Min", "Max"]
=======
static_values=["Avg","Std","Min","Max"]
>>>>>>> e50e507503bf68bfe9ab992773d0bb123de38a8a
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
    elif start_processing :
        # 처음에는 40자를 자르고, 이후부터는 11자씩 자르고 값 추출
        if len(str(row[0])) > 40:
            values = [str(row[0])[:40].strip()] + [
                value.strip()
                for value in [
                    str(row[0])[i : i + 11] for i in range(40, len(str(row[0])), 11)
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
    result_dfs.append(subset_df)

# 리스트에 있는 데이터프레임들을 하나로 합치기
result_df = pd.concat(result_dfs, ignore_index=True)

# 행 별로 조회하면서 Column명과 동일한 행 제거
result_df = result_df[~result_df.apply(lambda row: any(row == column_names), axis=1)]
<<<<<<< HEAD
for i in range(0, 4):
    result_df = result_df[
        ~result_df.apply(lambda row: any(row == static_values[i]), axis=1)
    ]
=======
for i in range(0,4):
    result_df = result_df[~result_df.apply(lambda row: any(row == static_values[i]), axis=1)]
>>>>>>> e50e507503bf68bfe9ab992773d0bb123de38a8a

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
<<<<<<< HEAD
result_df.to_excel("inventors.xlsx")
=======
result_df.to_excel('inventors.xlsx')
>>>>>>> e50e507503bf68bfe9ab992773d0bb123de38a8a
