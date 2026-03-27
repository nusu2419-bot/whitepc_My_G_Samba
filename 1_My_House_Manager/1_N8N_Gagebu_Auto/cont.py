import pandas as pd
import os
from excel_value_only_writer import write_sheets_value_only

# 1. 경로 설정
_base = os.path.dirname(os.path.abspath(__file__))
_report_dir = "/mnt/photos/My_G_Samba/1_My_House_Manager/4_Report"
input_path = os.path.join(_base, "A_CONTRACT.xlsx")
output_path = os.path.join(_report_dir, "건물별_거주현황_명부.xlsx")

# 2. 각 건물의 전체 호실 리스트 (첨부하신 그림의 내용에 맞게 숫자를 수정하세요)
# 여기에 적힌 호실들이 엑셀의 왼쪽 두 열(건물명, 호실)에 기본으로 생성됩니다.
room_config = {
    "봉명동": [101, 102, 103, 104, 105 , 106,201, 202, 203, 204, 205,206,301, 302,303,304,305,306], # 예시 데이터
    "신부동": [101,201,202,203,204,205,206,301,302,303,304,305,306,401,402,403,404,405,406,501,502,503,504,505,506], # 예시 데이터
    "쌍용동": [101, 102, 103, 201, 202, 203],  # 예시 데이터
}

def create_management_sheet():
    try:
        # 데이터 불러오기
        if not os.path.exists(input_path):
            print(f"파일을 찾을 수 없습니다: {input_path}")
            return

        df = pd.read_excel(input_path)
        
        # '상태'가 '거주중'인 데이터만 필터링
        active_residents = df[df['상태'] == '거주중'].copy()
        
        # 최종적으로 추출할 열 목록
        target_cols = ['건물명', '호실', '임차인', 'Phone', '보증금', '월세', '관리비', '부가세', '입주일', '상태']
        
        sheet_data = {}
        for building, rooms in room_config.items():
            # 해당 건물의 전체 호실 틀 생성 (기본값 세팅)
            base_df = pd.DataFrame({
                '건물명': building,
                '호실': rooms
            })
            
            # '거주중' 데이터와 병합 (Left Join)
            # 건물명과 호실이 모두 일치하는 행만 가져옵니다.
            building_data = active_residents[active_residents['건물명'] == building]
            merged_df = pd.merge(base_df, building_data, on=['건물명', '호실'], how='left')
            # '입주일'을 날짜 문자열(YYYY-MM-DD)로 변환하여 시간 부분 제거
            if '입주일' in merged_df.columns:
                merged_df['입주일'] = pd.to_datetime(merged_df['입주일'], errors='coerce').dt.strftime('%Y-%m-%d')
            
            # 요청하신 열만 선택하고 호실 순서대로 정렬
            # 데이터가 없는 열은 자동으로 NaN(빈칸)으로 표시됩니다.
            final_df = merged_df[target_cols].sort_values(by='호실')
            sheet_data[building] = final_df
            print(f"[{building}] 시트 작성 완료 (호실 수: {len(rooms)})")

        # 기존 파일이 있으면 서식은 유지하고 값만 갱신
        write_sheets_value_only(output_path, sheet_data)
                
        print(f"\n파일 생성이 완료되었습니다: {output_path}")

    except PermissionError:
        print("오류: 원본 또는 결과 엑셀 파일이 열려 있습니다. 파일을 닫고 다시 실행해 주세요.")
    except Exception as e:
        print(f"오류 발생: {e}")

if __name__ == "__main__":
    create_management_sheet()