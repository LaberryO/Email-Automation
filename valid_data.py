import pandas as pd

raw_df = pd.read_csv('settings/raw_data.csv')
fail_df = pd.read_csv('settings/backup_fail_data.csv')

# '이메일'을 기준으로 fail_df에 없는 데이터만 필터링
# raw_df['이메일']이 fail_df['이메일']에 포함되지(~) 않는 행만 남김
data_df = raw_df[~raw_df['업체명'].isin(fail_df['업체명'])]

# 결과 칼럼 선택 및 저장
target_columns = ['업체명', '이메일', '연락처']
result_df = data_df[target_columns]
result_df.to_csv('settings/data.csv', index=False)