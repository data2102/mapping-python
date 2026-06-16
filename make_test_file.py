import pandas as pd
import random

print("⏳ 1만 행 대용량 테스트 파일 생성 중...")

# 1. 1만 행짜리 거대한 Before 데이터 생성
total_rows = 10000
before_data = []

# 랜덤하게 조합할 품명과 규격 샘플들
items = ['방수 석고보드', '일반 석고보드', '인테리어 필름', 'LED 다운라이트', 'T5 간접조명', '멀티탭', '절연테이프', '스위치커버']
specs = ['9.5T * 3 * 6', '12.5T * 3 * 6', '우드 패턴', '3인치 주백색', '6인치 전구색', '1200mm 주광색', '4구 3m', '1구 화이트']
units = ['장', '롤', '개', '박스']

for i in range(1, total_rows + 1):
    # MAT-00001 ~ MAT-10000 구조의 고유 자재코드 생성
    code = f"MAT-{i:05d}"
    name = random.choice(items) + f"_{random.randint(1, 20)}"  # 구분을 위해 뒤에 숫자 결합
    spec = random.choice(specs)
    qty = random.randint(50, 500)  # 50개 ~ 500개 사이 랜덤 수량
    unit = random.choice(units)

    before_data.append({
        '자재코드': code,
        '품명': name,
        '규격': spec,
        '기존수량': qty,
        '단위': unit
    })

df_before = pd.DataFrame(before_data)
df_before.to_excel('before_huge.xlsx', index=False)
print("🏢 1. [before_huge.xlsx] (10,000행) 생성 완료!")

# 2. 변경 사항을 체감할 수 있는 콤팩트한 After 데이터 생성
# (Before의 일부 데이터를 가져와 수량을 바꾸고, 아예 새로운 신규 데이터도 추가)
after_data = []

# (A) 기존 1만 개 중 상위 50개 품목을 가져와서 '일부 항목의 수량'을 변경
for idx, row in df_before.head(50).iterrows():
    new_row = {
        '자재코드': row['자재코드'],
        '상품명': row['품명'],  # 일부러 타사 양식처럼 열 이름 변경
        '사이즈': row['규격'],  # 일부러 타사 양식처럼 열 이름 변경
        '수량': row['기존수량'],  # 기본은 기존 수량 유지
        '단위': row['단위']
    }
    # 그중 15개 항목은 수량을 강제로 변경 (+50 또는 -30)
    if idx % 3 == 0:
        new_row['수량'] = row['기존수량'] + random.choice([50, -30])

    after_data.append(new_row)

# (B) 마스터(Before)에는 아예 존재하지 않는 '완전 신규 자재' 5개 추가
for i in range(1, 6):
    after_data.append({
        '자재코드': f"NEW-{i:03d}",
        '상품명': f"🚀 외부업체 전용 특수 자재_{i}",
        '사이즈': "최신형 규격",
        '수량': random.randint(10, 100),
        '단위': "개"
    })

df_after = pd.DataFrame(after_data)
df_after.to_excel('after_sample.xlsx', index=False)
print("🚚 2. [after_sample.xlsx] (55행) 생성 완료!")
print("✨ 테스트 파일 준비 끝! 이제 Streamlit 웹을 켜서 두 파일을 넣어보세요.")