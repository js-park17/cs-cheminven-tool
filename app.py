# -*- coding: utf-8 -*-
"""
6호서식(유해화학물질 목록) 작성 툴
- MSDS(PDF) 여러 개를 업로드하면 CAS/물질명/비중/인화점/폭발한계/증기압/화관법 규제구분 등을
  자동 추출하여 별지 제6호서식과 화학물질인벤토리(내부관리용) 두 가지 결과물로 내려받을 수 있다.
"""
import streamlit as st
import pandas as pd

import msds_parser as mp
import xlsx_writers as xw
import kreach_lookup as kr

st.set_page_config(page_title="6호서식 작성 툴", layout="wide")
st.title("🧪 유해화학물질 목록(별지 제6호서식) 자동 작성 툴")
st.markdown("##### MSDS(PDF)를 업로드하면 CAS·비중·인화점·폭발한계·증기압·화관법 규제구분을 자동 추출합니다.")

with st.sidebar:
    st.header("ℹ️ 사용 안내")
    st.markdown(
        "1. 각 물질의 MSDS PDF 파일을 모두 선택해 업로드하세요.\n"
        "2. 자동 추출된 표를 검토하고, 노란색으로 표시된 항목은 직접 확인·수정하세요.\n"
        "3. 완료 후 아래에서 6호서식과 화학물질인벤토리를 각각 내려받으세요."
    )
    st.divider()
    st.caption(
        "※ 다성분 혼합물(제품) MSDS는 화학물질관리법상 성분별 규제번호가 표에서 겹쳐 표기되어 "
        "자동으로 신뢰성 있게 분리할 수 없는 경우가 많습니다. 이런 행은 '다성분혼합물'로 표시되며 "
        "화관법 규제 여부를 NCIS(화학물질정보시스템) 등에서 직접 확인해 주세요."
    )

uploaded_files = st.file_uploader(
    "MSDS PDF 파일 업로드 (여러 개 선택 가능)", type=["pdf"], accept_multiple_files=True
)

if uploaded_files:
    all_rows = []
    errors = []
    with st.spinner("MSDS 분석 중..."):
        for uf in uploaded_files:
            try:
                rows = mp.parse_msds(uf.read(), uf.name)
                all_rows.extend(rows)
            except Exception as e:
                errors.append((uf.name, str(e)))

    if errors:
        for fname, err in errors:
            st.error(f"'{fname}' 처리 중 오류: {err}")

    if all_rows:
        for i, row in enumerate(all_rows, start=1):
            row["연번"] = i
            hz = row.get("고유번호_유해화학물질") or ""
            ap = row.get("고유번호_사고대비물질") or ""
            parts = []
            if hz:
                parts.append(hz.split(",")[0])
            if ap:
                parts.append(f"사고대비 {ap.split(',')[0]}")
            row["고유번호"] = "\n".join(parts)
            note_parts = []
            if row.get("다성분혼합물"):
                note_parts.append(f"다성분 혼합물({row.get('product_name','')}) - 화관법 규제 수동확인 필요")
            if row.get("needs_review") and not row.get("다성분혼합물"):
                note_parts.append("자동추출값 확인 필요")
            row["비고"] = " / ".join(note_parts)

        df = pd.DataFrame(all_rows)

        display_cols = [
            "연번", "source_file", "물질명", "CAS번호", "고유번호", "물질상태", "함량(%)", "비중",
            "폭발하한", "폭발상한", "독성구분_항목", "독성구분_등급", "위험노출수준", "허용농도값",
            "증기압(mmHg)", "부식성", "비고", "needs_review",
        ]
        show_df = df[display_cols].rename(columns={"source_file": "출처파일", "needs_review": "확인필요"})

        st.write("---")
        c1, c2, c3 = st.columns(3)
        c1.metric("업로드 MSDS 수", f"{len(uploaded_files)}개")
        c2.metric("추출된 물질(행) 수", f"{len(df)}개")
        c3.metric("확인 필요 항목", f"{int(df['needs_review'].sum())}개")

        st.write("### 📋 추출 결과 검토 및 수정")
        st.caption("노란색 배경 없이도 셀을 직접 클릭해 값을 수정할 수 있습니다. '확인필요' 열이 True인 행은 특히 검토해 주세요.")

        edited_df = st.data_editor(
            show_df,
            column_config={
                "확인필요": st.column_config.CheckboxColumn(),
                "비중": st.column_config.NumberColumn(format="%.4f"),
                "폭발하한": st.column_config.NumberColumn(format="%.2f"),
                "폭발상한": st.column_config.NumberColumn(format="%.2f"),
                "증기압(mmHg)": st.column_config.NumberColumn(format="%.2f"),
            },
            disabled=["연번", "출처파일"],
            hide_index=True,
            use_container_width=True,
            num_rows="dynamic",
        )

        st.write("---")
        st.write("### 📥 결과물 다운로드")

        # edited_df -> 원본 rows 구조로 되돌려서 xlsx 작성 (누락 필드는 df에서 보완)
        merged = edited_df.rename(columns={"출처파일": "source_file", "확인필요": "needs_review"})
        export_rows = merged.to_dict("records")
        # xlsx_writers가 참조하는 나머지 필드(product_name, 다성분혼합물 등)는 원본 df에서 순번으로 매핑
        for i, r in enumerate(export_rows):
            if i < len(all_rows):
                r["product_name"] = all_rows[i].get("product_name", "")
                r["다성분혼합물"] = all_rows[i].get("다성분혼합물", False)

        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            ho6_bytes = xw.build_6ho_form(export_rows)
            st.download_button(
                label="[별지6호서식] 유해화학물질목록 다운로드",
                data=ho6_bytes.getvalue(),
                file_name="1-나-1_유해화학물질목록(별지6호서식).xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        with col_dl2:
            inv_bytes = xw.build_inventory(export_rows)
            st.download_button(
                label="[화학물질인벤토리-기본] 내부관리용 다운로드",
                data=inv_bytes.getvalue(),
                file_name="화학물질인벤토리.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        st.write("---")
        st.write("### 🌐 K-REACH(화학물질정보처리시스템) 연동 인벤토리")
        st.caption(
            "CAS번호로 K-REACH 화학물질통합검색(kreach.mcee.go.kr)을 자동 조회하여 "
            "KE번호·인체급성/인체만성/생태 고시번호·사고대비물질 번호·중점관리물질 등을 "
            "공식 데이터베이스 기준으로 채웁니다. 함량기준(%)은 K-REACH 목록화면에 없어 "
            "MSDS 원문에서 보조 추출한 값이며 상세 확인이 필요합니다."
        )
        run_kreach = st.button("K-REACH 조회 시작")

        if run_kreach:
            cas_list = [r.get("CAS번호", "") for r in export_rows if r.get("CAS번호")]
            progress = st.progress(0.0, text="K-REACH 조회 준비 중...")

            def _cb(i, total, cas):
                progress.progress(i / total, text=f"K-REACH 조회 중... ({i}/{total}) {cas}")

            kreach_results = kr.search_many(cas_list, progress_cb=_cb)
            progress.empty()

            kreach_rows = []
            for i, row in enumerate(export_rows, start=1):
                cas = row.get("CAS번호", "")
                kres = kreach_results.get(cas, {})
                content_pct = str(row.get("함량(%)", "") or "")
                nums = [float(n) for n in __import__("re").findall(r"\d+(?:\.\d+)?", content_pct)]
                min_pct = min(nums) if nums else ""
                max_pct = max(nums) if nums else ""

                # MSDS 텍스트에서 보조 함량기준(%) 추출 (예: "...10% 이상 함유한 혼합물")
                def _threshold_from_text(text):
                    m = __import__("re").search(r"(\d+(?:\.\d+)?)\s*%\s*이상", text or "")
                    return float(m.group(1)) if m else ""

                acute_pct = _threshold_from_text(row.get("고유번호_유해화학물질", "")) if kres.get("acute_no") else ""
                chronic_pct = _threshold_from_text(row.get("고유번호_유해화학물질", "")) if kres.get("chronic_no") and not acute_pct else ""
                accident_pct = _threshold_from_text(row.get("고유번호_사고대비물질", "")) if kres.get("accident_prep_no") else ""

                hazard_types = []
                if kres.get("acute_no"):
                    hazard_types.append("인체급성")
                if kres.get("chronic_no"):
                    hazard_types.append("인체만성")
                if kres.get("eco_no"):
                    hazard_types.append("생태")
                if kres.get("accident_prep_no"):
                    hazard_types.append("사고대비물질")

                kreach_rows.append({
                    "seq": i,
                    "product_name": row.get("product_name", ""),
                    "물질명": kres.get("name_kr") or row.get("물질명", ""),
                    "CAS번호": cas,
                    "ke_no": kres.get("ke_no", ""),
                    "함량(%)": row.get("함량(%)", ""),
                    "최소함량": min_pct,
                    "최고함량": max_pct,
                    "msds_reg_no": row.get("msds_reg_no", ""),
                    "ke_existing": kres.get("ke_no", ""),
                    "acute_no": kres.get("acute_no", ""),
                    "acute_flag": "O" if kres.get("acute_no") else "",
                    "acute_pct": acute_pct,
                    "chronic_no": kres.get("chronic_no", ""),
                    "chronic_flag": "O" if kres.get("chronic_no") else "",
                    "chronic_pct": chronic_pct,
                    "eco_no": kres.get("eco_no", ""),
                    "eco_flag": "O" if kres.get("eco_no") else "",
                    "eco_pct": "",
                    "accident_prep_no": kres.get("accident_prep_no", ""),
                    "accident_prep_flag": "O" if kres.get("accident_prep_no") else "",
                    "accident_prep_pct": accident_pct,
                    "restricted_flag": "O" if kres.get("restricted_raw") else "",
                    "restricted_no": kres.get("restricted_raw", ""),
                    "restricted_pct": "",
                    "prohibited_flag": "",
                    "prohibited_no": "",
                    "prohibited_pct": "",
                    "version": row.get("revision_no", ""),
                    "revision_date": row.get("revision_date", ""),
                    "enact_no": "",
                    "note": "",
                    "product_category": "",
                    "hazard_target": "유해화학물질 대상" if hazard_types else "-",
                    "hazard_type": ", ".join(hazard_types) if hazard_types else "-",
                    "kreach_status": kres.get("status", "error"),
                    "kreach_note": kres.get("detail_note", ""),
                })

            st.session_state["kreach_rows"] = kreach_rows

        if "kreach_rows" in st.session_state:
            kdf = pd.DataFrame(st.session_state["kreach_rows"])
            st.write(f"조회 완료: {len(kdf)}건 (실패/미등록 {int((kdf['kreach_status'] != 'ok').sum())}건)")
            st.dataframe(kdf, use_container_width=True, hide_index=True)
            kreach_bytes = xw.build_kreach_inventory(st.session_state["kreach_rows"])
            st.download_button(
                label="[화학물질인벤토리-KREACH연동] 다운로드",
                data=kreach_bytes.getvalue(),
                file_name="화학물질인벤토리_KREACH연동.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    else:
        st.warning("MSDS에서 물질 정보를 추출하지 못했습니다. 파일을 확인해 주세요.")
else:
    st.info("왼쪽 안내를 참고하여 MSDS PDF 파일을 업로드하면 자동 분석이 시작됩니다.")
