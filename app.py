import streamlit as st
import openpyxl
import io
from datetime import datetime

st.set_page_config(page_title="자동 견적서 생성 에이전트", layout="wide", page_icon="📄")

st.title("📄 통합 견적서 자동 생성 에이전트")
st.markdown("기본 정보와 본견적서 내용만 입력하면 **본견적서**, **가견적서(+5%)**, **타견적서(+10%)**가 만 원 단위 절삭되어 자동으로 완성됩니다.")

# ---------------------------------------------------------
# Helper Functions
# ---------------------------------------------------------
def num2kor(num):
    """숫자를 한글 금액 표현으로 변환 (예: 340000 -> 삼십사만)"""
    if num == 0:
        return "영"
    units = ['', '만', '억', '조']
    sub_units = ['', '십', '백', '천']
    digits = ['', '일', '이', '삼', '사', '오', '육', '칠', '팔', '구']
    
    result = []
    unit_idx = 0
    
    while num > 0:
        part = num % 10000
        if part > 0:
            part_str = ""
            for i in range(4):
                d = (part // (10 ** i)) % 10
                if d > 0:
                    digit_str = digits[d] if not (d == 1 and i > 0) else ""
                    part_str = digit_str + sub_units[i] + part_str
            result.append(part_str + units[unit_idx])
        num //= 10000
        unit_idx += 1
        
    return "".join(reversed(result))

def calculate_quote(items, rate, include_vat=True, cut_unit=10000):
    """
    품목 리스트, 부가세 계산 및 총액까지 
    모든 단계에서 cut_unit(기본 10,000원) 미만 단수를 버림(절삭) 처리
    """
    if not items:
        return [], 0, 0, 0
    
    adjusted_items = []
    for item in items:
        raw_adj = item["amount"] * (1 + rate / 100.0)
        if cut_unit > 1:
            adj_amt = int((raw_adj // cut_unit) * cut_unit)
        else:
            adj_amt = int(round(raw_adj))
        adjusted_items.append(adj_amt)
        
    supply_total = sum(adjusted_items)
    
    if include_vat:
        raw_vat = supply_total * 0.1
        # 부가세 자체도 만 원 단위 미만 절삭(버림)
        if cut_unit > 1:
            vat_total = int((raw_vat // cut_unit) * cut_unit)
        else:
            vat_total = int(round(raw_vat))
        grand_total = supply_total + vat_total
    else:
        vat_total = 0
        grand_total = supply_total
        
    return adjusted_items, supply_total, vat_total, grand_total

def safe_write_cell(ws, row, col, value):
    """병합 셀 에러 방지를 위한 안전한 값 입력 함수"""
    cell = ws.cell(row=row, column=col)
    if type(cell).__name__ == 'MergedCell':
        for rng in ws.merged_cells.ranges:
            if cell.coordinate in rng:
                ws.cell(row=rng.min_row, column=rng.min_col, value=value)
                return
    else:
        cell.value = value

# ---------------------------------------------------------
# Sidebar Settings
# ---------------------------------------------------------
st.sidebar.header("⚙️ 상세 설정")
ga_rate = st.sidebar.number_input("가견적서 인상률 (%)", value=5.0, step=1.0)
ta_rate = st.sidebar.number_input("타견적서 인상률 (%)", value=10.0, step=1.0)

cut_option = st.sidebar.selectbox(
    "품목 금액 및 부가세 절삭(버림) 단위",
    options=[10000, 1000, 100, 1],
    index=0, # 만 원 단위 기본 선택
    format_func=lambda x: "만 원 단위 절삭 (기본)" if x == 10000 else ("천 원 단위 절삭" if x == 1000 else ("백 원 단위 절삭" if x == 100 else "절삭 없음"))
)

# ---------------------------------------------------------
# Main Options & Form Inputs
# ---------------------------------------------------------
st.subheader("⚙️ VAT 적용 옵션 선택")
include_vat_option = st.radio(
    "VAT(부가가치세) 계산 방식을 선택하세요:",
    options=["VAT 포함 (공급가액 + 10% 부가세 자동 합산)", "VAT 별도/미포함 (부가세 계산 안 함)"],
    index=0,
    horizontal=True
)

include_vat = include_vat_option.startswith("VAT 포함")

st.divider()

col1, col2 = st.columns(2)

with col1:
    st.subheader("📌 기본 정보")
    client_name = st.text_input("발주처 (회사/기관명)", value="사단법인 커뮤니티와경제")
    project_name = st.text_input("용역명 (사업명)", value="2026년 사회적경제 활성화 지원 사업")
    issue_date = st.date_input("발행일자", datetime.today())
    manager_name = st.text_input("담당자 성명", value="강은경")

with col2:
    st.subheader("📂 템플릿 파일")
    uploaded_file = st.file_uploader("견적서 양식 엑셀 파일 (.xlsx)", type=["xlsx"])
    st.caption("※ 업로드하지 않으면 기본 양식을 사용합니다.")

st.divider()
st.subheader("💡 본견적서 항목 작성 (공급가액 기준)")

if "items" not in st.session_state:
    st.session_state["items"] = [
        {"category": "인건비", "detail": "연구원 인건비", "spec": "2명 * 3개월", "amount": 2000000, "note": ""},
        {"category": "경비", "detail": "회의비 및 임차료", "spec": "장소 대여 2회", "amount": 1000000, "note": ""}
    ]

def add_item():
    st.session_state["items"].append({"category": "", "detail": "", "spec": "", "amount": 0, "note": ""})

def remove_item(idx):
    if len(st.session_state["items"]) > 1:
        st.session_state["items"].pop(idx)

for idx, item in enumerate(st.session_state["items"]):
    cols = st.columns([2, 3, 3, 2, 2, 1])
    item["category"] = cols[0].text_input(f"지출구분 #{idx+1}", item["category"], key=f"cat_{idx}")
    item["detail"] = cols[1].text_input(f"내용 #{idx+1}", item["detail"], key=f"det_{idx}")
    item["spec"] = cols[2].text_input(f"산출내역 #{idx+1}", item["spec"], key=f"spec_{idx}")
    item["amount"] = cols[3].number_input(f"금액(원) #{idx+1}", value=int(item["amount"]), step=10000, key=f"amt_{idx}")
    item["note"] = cols[4].text_input(f"비고 #{idx+1}", item["note"], key=f"note_{idx}")
    
    if cols[5].button("❌", key=f"del_{idx}"):
        remove_item(idx)
        st.rerun()

st.button("➕ 항목 추가", on_click=add_item)

# ---------------------------------------------------------
# Calculations & Preview
# ---------------------------------------------------------
bon_amts, supply_bon, vat_bon, grand_bon = calculate_quote(st.session_state["items"], 0, include_vat=include_vat, cut_unit=cut_option)
ga_amts, supply_ga, vat_ga, grand_ga = calculate_quote(st.session_state["items"], ga_rate, include_vat=include_vat, cut_unit=cut_option)
ta_amts, supply_ta, vat_ta, grand_ta = calculate_quote(st.session_state["items"], ta_rate, include_vat=include_vat, cut_unit=cut_option)

vat_str = "(VAT 포함)" if include_vat else "(VAT 별도)"

st.markdown("---")
st.subheader(f"📊 견적 금액 미리보기 [{vat_str}] - 만 원 단위 절삭 적용")
p_col1, p_col2, p_col3 = st.columns(3)

p_col1.metric("본견적서 최종 금액", f"{grand_bon:,} 원", f"공급가액: {supply_bon:,}원 | 부가세: {vat_bon:,}원")
p_col2.metric(f"가견적서 (+{ga_rate}%) 최종 금액", f"{grand_ga:,} 원", f"공급가액: {supply_ga:,}원 | 부가세: {vat_ga:,}원")
p_col3.metric(f"타견적서 (+{ta_rate}%) 최종 금액", f"{grand_ta:,} 원", f"공급가액: {supply_ta:,}원 | 부가세: {vat_ta:,}원")

# ---------------------------------------------------------
# Excel Generation Logic
# ---------------------------------------------------------
def generate_excel():
    template_path = "본,가,타견적서 양식_본_가_타.xlsx" if uploaded_file is None else uploaded_file
    wb = openpyxl.load_workbook(template_path)
    
    # 1. 본견적서 작성 (시트 1)
    if "본견적서" in wb.sheetnames:
        ws = wb["본견적서"]
        safe_write_cell(ws, 7, 3, client_name)
        safe_write_cell(ws, 12, 4, client_name)
        safe_write_cell(ws, 14, 4, issue_date.strftime("%Y-%m-%d"))
        safe_write_cell(ws, 14, 9, manager_name)
        safe_write_cell(ws, 18, 4, project_name)
        safe_write_cell(ws, 21, 7, "(부가세 포함)" if include_vat else "(부가세 별도)")
        
        start_row = 26
        for idx, item in enumerate(st.session_state["items"]):
            r = start_row + idx
            safe_write_cell(ws, r, 3, item["category"])
            safe_write_cell(ws, r, 4, item["detail"])
            safe_write_cell(ws, r, 5, item["spec"])
            safe_write_cell(ws, r, 9, bon_amts[idx])
            safe_write_cell(ws, r, 10, item["note"])
            
        safe_write_cell(ws, 41, 9, vat_bon)
        safe_write_cell(ws, 42, 9, grand_bon)
        safe_write_cell(ws, 21, 4, num2kor(grand_bon))

    # 2. 가견적서 작성 (시트 2)
    if "가견적서(본견+5%)" in wb.sheetnames:
        ws = wb["가견적서(본견+5%)"]
        safe_write_cell(ws, 7, 3, client_name)
        safe_write_cell(ws, 12, 4, client_name)
        safe_write_cell(ws, 14, 4, issue_date.strftime("%Y-%m-%d"))
        safe_write_cell(ws, 14, 9, manager_name)
        safe_write_cell(ws, 18, 4, project_name)
        safe_write_cell(ws, 21, 7, "(부가세 포함)" if include_vat else "(부가세 별도)")
        
        start_row = 26
        for idx, item in enumerate(st.session_state["items"]):
            r = start_row + idx
            safe_write_cell(ws, r, 3, item["category"])
            safe_write_cell(ws, r, 4, item["detail"])
            safe_write_cell(ws, r, 5, item["spec"])
            safe_write_cell(ws, r, 9, ga_amts[idx])
            safe_write_cell(ws, r, 10, item["note"])
            
        safe_write_cell(ws, 41, 9, vat_ga)
        safe_write_cell(ws, 42, 9, grand_ga)
        safe_write_cell(ws, 21, 4, num2kor(grand_ga))

    # 3. 타견적서 작성 (시트 3)
    if "타견적서(본견+10%)" in wb.sheetnames:
        ws = wb["타견적서(본견+10%)"]
        safe_write_cell(ws, 3, 4, None)
        safe_write_cell(ws, 7, 3, client_name)
        safe_write_cell(ws, 10, 4, client_name)
        safe_write_cell(ws, 12, 4, issue_date.strftime("%Y-%m-%d"))
        safe_write_cell(ws, 12, 9, manager_name)
        safe_write_cell(ws, 18, 4, project_name)
        
        start_row = 23
        for idx, item in enumerate(st.session_state["items"]):
            r = start_row + idx
            supply_price = ta_amts[idx]
            vat_price = int((supply_price * 0.1 // 10000) * 10000) if include_vat else 0
            
            safe_write_cell(ws, r, 3, idx + 1)
            safe_write_cell(ws, r, 4, f"[{item['category']}] {item['detail']}")
            safe_write_cell(ws, r, 5, item["spec"])
            safe_write_cell(ws, r, 7, supply_price)
            safe_write_cell(ws, r, 8, vat_price)
            safe_write_cell(ws, r, 9, item["note"])
            
        safe_write_cell(ws, 38, 7, supply_ta)
        safe_write_cell(ws, 39, 7, vat_ta)
        safe_write_cell(ws, 40, 7, grand_ta)
        safe_write_cell(ws, 17, 4, num2kor(grand_ta))

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

st.divider()

if st.button("🚀 만 원 단위 절삭 견적서 3종 엑셀 파일 생성하기", type="primary", use_container_width=True):
    excel_data = generate_excel()
    file_name = f"통합견적서_{project_name}_{issue_date.strftime('%Y%m%d')}.xlsx"
    st.download_button(
        label="📥 생성된 엑셀 파일 다운로드",
        data=excel_data,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )