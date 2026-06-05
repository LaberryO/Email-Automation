import pandas as pd

input_path = "settings/input.csv"

# 1. 원본 CSV 파일 읽기
# 'input.csv' 자리에 본인의 파일 경로와 이름을 넣으세요.
# 만약 읽을 때 한글이 깨진다면 encoding='cp949'를 추가해보세요.
try:
    df = pd.read_csv(input_path, encoding='utf-8')
except UnicodeDecodeError:
    df = pd.read_csv(input_path, encoding='cp949')

# 2. 검증할 정규식 패턴 정의
# 02-123-4567, 02-1234-5678, 010-1234-5678 포맷만 통과
phone_pattern = r'^\d{2,3}-\d{3,4}-\d{4}$'

# 3. '전화번호' 컬럼에서 올바른 포맷만 필터링
# ※ '전화번호' 부분을 실제 CSV 파일의 연락처 컬럼명으로 정확히 수정해주세요!
# .astype(str)은 빈 칸(NaN)이나 숫자형으로 인식된 데이터를 문자열로 안전하게 변환합니다.
is_valid = df['연락처'].astype(str).str.match(phone_pattern, na=False)

# 올바른 데이터만 남기기
df_clean = df[is_valid]

# 4. 필터링된 데이터를 새로운 CSV로 저장
# index=False: 앞자리에 붙는 0, 1, 2... 순번 컬럼을 저장하지 않음
# encoding='utf-8-sig': 저장 후 MS 엑셀(Excel)에서 더블클릭해 열어도 한글이 깨지지 않음
df_clean.to_csv('phone_normalized_only.csv', index=False, encoding='utf-8-sig')

print(f"정규화 필터링 완료!")
print(f"원본 데이터 수: {len(df)}개 -> 올바른 포맷 데이터 수: {len(df_clean)}개")
print("결과가 'phone_normalized_only.csv' 파일로 저장되었습니다.")