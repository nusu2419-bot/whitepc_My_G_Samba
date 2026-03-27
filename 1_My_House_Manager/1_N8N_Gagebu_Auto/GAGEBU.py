import pandas as pd
import glob
import os

input_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Excel_Gagebu")
base_path = os.path.dirname(os.path.abspath(__file__))
output_parquet = os.path.join(base_path, "merged_gagebu.parquet")
output_duplicates = os.path.join(base_path, "duplicates_list.xlsx")

# 파일 읽기 및 병합
file_list = [f for f in glob.glob(os.path.join(input_dir, "*.xls*")) if not os.path.basename(f).startswith("~$")]
df_list = [pd.read_excel(file) for file in file_list]

if df_list:
    merged_df = pd.concat(df_list, ignore_index=True)
    merged_df.to_parquet(output_parquet, engine='pyarrow', index=False)
    
    # H열(인덱스 7) 기준 모든 중복 행 추출 및 엑셀 저장
    h_col = merged_df.columns[7]
    duplicates = merged_df[merged_df.duplicated(subset=[h_col], keep=False)]
    
    if not duplicates.empty:
        duplicates.to_excel(output_duplicates, index=False)
        print(f"중복 데이터 {len(duplicates)}건을 엑셀로 저장했습니다: {output_duplicates}")
    else:
        print("중복된 데이터가 없습니다.")