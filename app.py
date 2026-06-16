import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="자재 마스터 최신화 시스템", layout="wide")

st.title("📦 마스터 자재 목록 변동사항 기재기 (최종 완결판)")
st.write("실무 최적화된 컬럼 구조와 대용량 고속 엔진, 그리고 자유로운 항목 추가/삭제(X) 기능이 모두 포함된 완성형 시스템입니다.")

# 1. 파일 업로드
col1, col2 = st.columns(2)
with col1:
    file_a = st.file_uploader("🏢 1. 현재 관리 중인 자재 목록 (Before / 기준)", type=["xlsx"])
with col2:
    file_b = st.file_uploader("🚚 2. 변경/외부 업체 자재 목록 (After / 비교)", type=["xlsx"])

if file_a and file_b:
    try:
        df_before = pd.read_excel(file_a).fillna("-")
        df_after = pd.read_excel(file_b).fillna("-")

        cols_before = list(df_before.columns)
        cols_after = list(df_after.columns)

        st.markdown("---")
        st.subheader("⚙️ 짝맞춤 기준 및 검증 항목 설정")

        grid_col1, grid_col2 = st.columns(2)

        # ----------------------------------------------------------------
        # [기능 개선 1] 제품 식별 기준 설정 (개별 행 삭제 기능 추가)
        # ----------------------------------------------------------------
        with grid_col1:
            st.markdown("### 🔑 1. 제품 식별 기준 설정 (복합 일치 조건)")

            # 최초 실행 시 세션 스테이트에 리스트 형태로 관리구조 초기화
            if 'id_list' not in st.session_state:
                st.session_state.id_list = [0]  # 기본적으로 1개의 행 생성용 인덱스

            identity_mappings = []

            # 생성된 행들의 인덱스를 순회하며 UI를 그림
            for idx, i in enumerate(st.session_state.id_list):
                c_b, c_a, c_del = st.columns([4, 4, 1])  # 삭제 버튼을 위해 3칸으로 쪼갬

                with c_b:
                    def_b_idx = cols_before.index('품명') if idx == 0 and '품명' in cols_before else (
                        cols_before.index('규격') if idx == 1 and '규격' in cols_before else 0)
                    b_col = st.selectbox(f"Before 기준 열 #{idx + 1}", cols_before, index=def_b_idx,
                                         key=f"id_b_select_{i}")
                with c_a:
                    def_a_idx = cols_after.index('商品명') if idx == 0 and '상품명' in cols_after else (
                        cols_after.index('사이즈') if idx == 1 and '사이즈' in cols_after else 0)
                    a_col = st.selectbox(f"➔ After 비교 열 #{idx + 1}", cols_after, index=def_a_idx,
                                         key=f"id_a_select_{i}")
                with c_del:
                    st.markdown("<div style='padding-top: 24px;'></div>", unsafe_allow_html=True)  # 라벨 높이 맞추기용 빈칸
                    # 첫 번째 필수 기준 행을 제외하고는 우측에 삭제 버튼(X) 생성
                    if idx > 0:
                        if st.button("🗑️", key=f"id_del_btn_{i}", help="이 조건 삭제"):
                            st.session_state.id_list.remove(i)
                            st.rerun()

                identity_mappings.append((b_col, a_col))

            if st.button("➕ 식별 기준 항목 추가", key="add_id"):
                # 겹치지 않는 새로운 타임스탬프성 ID를 리스트에 추가
                new_id = max(st.session_state.id_list) + 1 if st.session_state.id_list else 0
                st.session_state.id_list.append(new_id)
                st.rerun()

        # ----------------------------------------------------------------
        # [기능 개선 2] 실제 변동값 검증 설정 (개별 행 삭제 기능 추가)
        # ----------------------------------------------------------------
        with grid_col2:
            st.markdown("### 🔢 2. 실제 변동값 검증 설정 (대조 대상)")

            if 'val_list' not in st.session_state:
                st.session_state.val_list = [0]

            value_mappings = []

            for idx, i in enumerate(st.session_state.val_list):
                c_b, c_a, c_del = st.columns([4, 4, 1])

                with c_b:
                    def_b_idx = cols_before.index('기존수량') if idx == 0 and '기존수량' in cols_before else 0
                    b_col = st.selectbox(f"Before 검증 대상 #{idx + 1}", cols_before, index=def_b_idx,
                                         key=f"val_b_select_{i}")
                with c_a:
                    def_a_idx = cols_after.index('수량') if idx == 0 and '수량' in cols_after else 0
                    a_col = st.selectbox(f"➔ After 대조 대상 #{idx + 1}", cols_after, index=def_a_idx,
                                         key=f"val_a_select_{i}")
                with c_del:
                    st.markdown("<div style='padding-top: 24px;'></div>", unsafe_allow_html=True)
                    if idx > 0:
                        if st.button("🗑️", key=f"val_del_btn_{i}", help="이 조건 삭제"):
                            st.session_state.val_list.remove(i)
                            st.rerun()

                value_mappings.append((b_col, a_col))

            if st.button("➕ 변동 검증 항목 추가", key="add_val"):
                new_id = max(st.session_state.val_list) + 1 if st.session_state.val_list else 0
                st.session_state.val_list.append(new_id)
                st.rerun()

        # 4. 실행 버튼
        st.markdown("---")
        if st.button("🚀 최종 가공 및 결과 엑셀 파일 생성", type="primary"):

            # 딕셔너리 해시 고속 맵핑 가동
            after_dict = {}
            for _, row in df_after.iterrows():
                j_key = " / ".join([str(row[a_col]).strip() for _, a_col in identity_mappings])
                after_dict[j_key] = row

            final_rows = []
            matched_after_keys = set()

            # 규칙 1 & 2: Before 데이터를 한 줄씩 고속 매칭
            for _, row_b in df_before.iterrows():
                j_key = " / ".join([str(row_b[b_col]).strip() for b_col, _ in identity_mappings])

                new_row = row_b.to_dict()
                new_row['변경수량'] = ""
                new_row['구분'] = '변경없음(유지)'
                new_row['변동내용'] = '변동 사항 없음'

                if j_key in after_dict:
                    matched_after_keys.add(j_key)
                    row_after = after_dict[j_key]

                    changes = []
                    for b_val_col, a_val_col in value_mappings:
                        v_before = row_b[b_val_col]
                        v_after = row_after[a_val_col]

                        if str(v_before).strip() != str(v_after).strip():
                            try:
                                if float(v_before) == int(float(v_before)): v_before = int(float(v_before))
                                if float(v_after) == int(float(v_after)): v_after = int(float(v_after))
                            except:
                                pass

                            new_row['변경수량'] = v_after
                            changes.append(f"기존 {v_before} ➔ 변경 {v_after}")

                    if changes:
                        new_row['구분'] = '수량변경'
                        new_row['변동내용'] = ", ".join(changes)

                final_rows.append(new_row)

            # 규칙 3: Before에 없던 신규 품목 아랫줄에 자연스럽게 결합
            for j_key, row_after in after_dict.items():
                if j_key not in matched_after_keys:
                    new_row = {col: "-" for col in cols_before}

                    if '자재코드' in cols_before and '자재코드' in df_after.columns:
                        new_row['자재코드'] = row_after['자재코드']

                    for b_col, a_col in identity_mappings:
                        new_row[b_col] = row_after[a_col]

                    for b_val_col, _ in value_mappings:
                        new_row[b_val_col] = "-"

                    v_new = row_after[value_mappings[0][1]]
                    try:
                        if float(v_new) == int(float(v_new)): v_new = int(float(v_new))
                    except:
                        pass

                    new_row['변경수량'] = v_new
                    new_row['구분'] = '신규추가'
                    new_row['변동내용'] = f"신규 자재 추가 (수량: {v_new})"

                    final_rows.append(new_row)

            df_result_final = pd.DataFrame(final_rows)

            # 기존수량 컬럼 우측에 변경수량 오도록 순서 재배치 안정화
            desired_order = []
            qty_b_col = value_mappings[0][0]
            idx_qty = cols_before.index(qty_b_col)

            for i, col in enumerate(cols_before):
                desired_order.append(col)
                if i == idx_qty:
                    desired_order.append('변경수량')

            desired_order.extend(['구분', '변동내용'])
            df_result_final = df_result_final[desired_order]

            # 최종 리포트 출력 및 엑셀 화일 변환
            st.markdown("---")
            st.subheader("📋 최종 완성된 리포트 미리보기")
            if len(df_result_final) > 500:
                st.warning(f"💡 데이터가 너무 많아 상위 500개만 화면에 미리 보여줍니다. 전체 데이터는 엑셀로 다운로드하세요.")
                st.dataframe(df_result_final.head(500), use_container_width=True)
            else:
                st.dataframe(df_result_final, use_container_width=True)

            out_master = io.BytesIO()
            with pd.ExcelWriter(out_master, engine='openpyxl') as writer:
                df_result_final.to_excel(writer, index=False)

            st.download_button(
                label="📥 캡처 양식과 똑같은 최종 결과 엑셀 다운로드",
                data=out_master.getvalue(),
                file_name="최종_자재_변동내역_리포트.xlsx",
                mime="application/vnd.ms-excel",
                type="primary"
            )
            st.success("🏁 대용량 매핑 검증 작업이 완벽히 끝났습니다!")

    except Exception as e:
        st.error(f"⚠️ 가공 중 예기치 못한 오류 발생. 선택하신 매핑 단어들을 다시 확인해 주세요. ({e})")