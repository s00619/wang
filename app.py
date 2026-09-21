import streamlit as st
import openpyxl
from openpyxl.styles import Alignment
import io
from datetime import datetime

st.set_page_config(page_title="자동 견적서 생성 에이전트", layout="wide", page_icon="📄")

st.title("📄 통합 견적서 자동 생성 에이전트")
st.markdown("기본 정보와 본견적서 내용만 입력하면 **본견적서**, **가견적서(+5%)**, **타견적서(+10%)**의 모든 금액이 **절삭(예: 27,300,000원)** 및 **천 단위 콤마(,)**, **한글 금액**, **자동 줄바꿈** 처리되어 완성됩니다.")

# ---------------------------------------------------------
# Helper Functions
# ---------------------------------------------------------
def num2kor(num):
    """숫자를 정확한 한글 금액 표현으로 변환 (예: 27300000 -> 이천칠백삼십만)"""
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

def calculate_quote(items, rate, include_vat=True, cut_unit=100000):
    """
    품목 리스트, 부가세 및 총액 계산 시 
    만 원 이하 자리를 버려서 27,300,000원 형태로 만듦 (100,000원 단위 절삭)
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
        if cut_unit > 1:
            vat_total = int((raw_vat // cut_unit) * cut_unit)
        else:
            vat_total = int(round(raw_vat))
        grand_total = supply_total + vat_total
    else:
        vat_total = 0
        grand_total = supply_total
        
    return adjusted_items, supply_total, vat_total, grand_total

def safe_write_cell(ws, row, col, value, wrap=False, is_number=False):
    """병합 셀 에러 방지, 천 단위 콤마 서식 및 자동 줄바꿈 지원 값 입력 함수"""
    cell = ws.cell(row=row, column=col)
    target_cell = cell
    if type(cell).__name__ == 'MergedCell':
        for rng in ws.merged_cells.ranges:
            if cell.coordinate in rng:
                target_cell = ws.cell(row=rng.min_row, column=rng.min_col)
                break
    target_cell.value = value
    
    if is_number and isinstance(value, (int, float)):
        target_cell.number_format = '#,##0'
        
    if wrap:
        target_cell.alignment = Alignment(wrap_text=True, vertical='center')

# ---------------------------------------------------------
# Sidebar Settings
# ---------------------------------------------------------
st.sidebar.header("⚙️ 상세 설정")
ga_rate = st.sidebar.number_input("가견적서 인상률 (%)", value=5.0, step=1.0)
ta_rate = st.sidebar.number_input("타견적서 인상률 (%)", value=10.0, step=1.0)

cut_option = st.sidebar.selectbox(
    "금액 절삭(버림) 단위 선택",
    options=[100000, 1000000, 10000, 1],
    index=0, # 만 원 이하 자리를 다 날려서 27,300,000원 형태로 만듦
    format_func=lambda x: "만 원 이하 절삭 (예: 27,300,000원 형태 - 기본)" if x == 100000 else ("백만 원 단위 절삭" if x == 1000000 else ("1천 원 이하 절삭" if x == 10000 else "절삭 없음"))
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
    manager_name = st.text_input("담당자 성명", value="이주영")

with col2:
    st.subheader("📂 템플릿 파일")
    uploaded_file = st.file_uploader("견적서 양식 엑셀 파일 (.xlsx)", type=["xlsx"])
    st.caption("※ 업로드하지 않으면 새로 수정된 확장 템플릿 양식을 기본으로 사용합니다.")

st.divider()
st.subheader("💡 본견적서 항목 작성 (공급가액 기준)")
st.caption("※ 산출내역이나 내용 입력 시 Shift+Enter로 줄바꿈을 입력할 수 있습니다.")

if "items" not in st.session_state:
    st.session_state["items"] = [
        {"category": "인건비", "detail": "연구원 인건비", "spec": "2명 * 3개월", "amount": 20000000, "note": ""},
        {"category": "경비", "detail": "회의비 및 임차료", "spec": "장소 대여 2회", "amount": 10000000, "note": ""}
    ]

def add_item():
    st.session_state["items"].append({"category": "", "detail": "", "spec": "", "amount": 0, "note": ""})

def remove_item(idx):
    if len(st.session_state["items"]) > 1:
        st.session_state["items"].pop(idx)

for idx, item in enumerate(st.session_state["items"]):
    cols = st.columns([2, 3, 3, 2, 2, 1])
    item["category"] = cols[0].text_input(f"지출구분 #{idx+1}", item["category"], key=f"cat_{idx}")
    item["detail"] = cols[1].text_area(f"내용 #{idx+1}", item["detail"], key=f"det_{idx}", height=68)
    item["spec"] = cols[2].text_area(f"산출내역 #{idx+1}", item["spec"], key=f"spec_{idx}", height=68)
    item["amount"] = cols[3].number_input(f"금액(원) #{idx+1}", value=int(item["amount"]), step=100000, key=f"amt_{idx}")
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
st.subheader(f"📊 견적 금액 미리보기 [{vat_str}]")
p_col1, p_col2, p_col3 = st.columns(3)

p_col1.metric("본견적서 최종 금액", f"{grand_bon:,} 원", f"한글: {num2kor(grand_bon)}원 | 부가세: {vat_bon:,}원")
p_col2.metric(f"가견적서 (+{ga_rate}%) 최종 금액", f"{grand_ga:,} 원", f"한글: {num2kor(grand_ga)}원 | 부가세: {vat_ga:,}원")
p_col3.metric(f"타견적서 (+{ta_rate}%) 최종 금액", f"{grand_ta:,} 원", f"한글: {num2kor(grand_ta)}원 | 부가세: {vat_ta:,}원")

# ---------------------------------------------------------
# Excel Generation Logic (확장 템플릿 지원)
# ---------------------------------------------------------
def generate_excel():
    template_path = "통합 견적서 자동 생성 에이전트 템플릿 파일.xlsx" if uploaded_file is None else uploaded_file
    wb = openpyxl.load_workbook(template_path)
    
    # 1. 본견적서 작성 (시트 1)
    if "본견적서" in wb.sheetnames:
        ws = wb["본견적서"]
        safe_write_cell(ws, 7, 3, client_name)
        safe_write_cell(ws, 12, 4, client_name)
        safe_write_cell(ws, 14, 4, issue_date.strftime("%Y-%m-%d"))
        safe_write_cell(ws, 14, 10, manager_name)
        safe_write_cell(ws, 18, 4, project_name)
        safe_write_cell(ws, 21, 7, "(부가세 포함)" if include_vat else "(부가세 별도)")
        
        start_row = 26
        # 기존 품목 영역 초기화 (Row 26 ~ 55)
        for r in range(start_row, 56):
            safe_write_cell(ws, r, 3, None)
            safe_write_cell(ws, r, 4, None)
            safe_write_cell(ws, r, 5, None)
            safe_write_cell(ws, r, 9, None)
            safe_write_cell(ws, r, 10, None)
            
        for idx, item in enumerate(st.session_state["items"]):
            r = start_row + idx
            if r < 56:
                safe_write_cell(ws, r, 3, item["category"])
                safe_write_cell(ws, r, 4, item["detail"], wrap=True)
                safe_write_cell(ws, r, 5, item["spec"], wrap=True)
                safe_write_cell(ws, r, 9, bon_amts[idx], is_number=True)
                safe_write_cell(ws, r, 10, item["note"], wrap=True)
            
        safe_write_cell(ws, 56, 9, vat_bon, is_number=True)
        safe_write_cell(ws, 57, 9, grand_bon, is_number=True)
        
        # 상단 한글 금액 표기(D21) 및 (\₩ ) 옆 숫자 표기(I21)
        safe_write_cell(ws, 21, 4, num2kor(grand_bon))
        safe_write_cell(ws, 21, 9, grand_bon, is_number=True)

    # 2. 가견적서 작성 (시트 2)
    if "가견적서(본견+5%)" in wb.sheetnames:
        ws = wb["가견적서(본견+5%)"]
        safe_write_cell(ws, 7, 3, client_name)
        safe_write_cell(ws, 12, 4, client_name)
        safe_write_cell(ws, 14, 4, issue_date.strftime("%Y-%m-%d"))
        safe_write_cell(ws, 14, 10, manager_name)
        safe_write_cell(ws, 18, 4, project_name)
        safe_write_cell(ws, 21, 7, "(부가세 포함)" if include_vat else "(부가세 별도)")
        
        start_row = 26
        for r in range(start_row, 55):
            safe_write_cell(ws, r, 3, None)
            safe_write_cell(ws, r, 4, None)
            safe_write_cell(ws, r, 5, None)
            safe_write_cell(ws, r, 9, None)
            safe_write_cell(ws, r, 10, None)

        for idx, item in enumerate(st.session_state["items"]):
            r = start_row + idx
            if r < 55:
                safe_write_cell(ws, r, 3, item["category"])
                safe_write_cell(ws, r, 4, item["detail"], wrap=True)
                safe_write_cell(ws, r, 5, item["spec"], wrap=True)
                safe_write_cell(ws, r, 9, ga_amts[idx], is_number=True)
                safe_write_cell(ws, r, 10, item["note"], wrap=True)
            
        safe_write_cell(ws, 55, 9, vat_ga, is_number=True)
        safe_write_cell(ws, 56, 9, grand_ga, is_number=True)
        
        # 상단 한글 금액 표기(D21) 및 (\₩ ) 옆 숫자 표기(I21)
        safe_write_cell(ws, 21, 4, num2kor(grand_ga))
        safe_write_cell(ws, 21, 9, grand_ga, is_number=True)

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
        for r in range(start_row, 52):
            safe_write_cell(ws, r, 4, None)
            safe_write_cell(ws, r, 5, None)
            safe_write_cell(ws, r, 7, None)
            safe_write_cell(ws, r, 8, None)
            safe_write_cell(ws, r, 9, None)

        for idx, item in enumerate(st.session_state["items"]):
            r = start_row + idx
            if r < 52:
                supply_price = ta_amts[idx]
                vat_price = int((supply_price * 0.1 // cut_option) * cut_option) if include_vat else 0
                
                safe_write_cell(ws, r, 3, idx + 1)
                safe_write_cell(ws, r, 4, f"[{item['category']}] {item['detail']}", wrap=True)
                safe_write_cell(ws, r, 5, item["spec"], wrap=True)
                safe_write_cell(ws, r, 7, supply_price, is_number=True)
                safe_write_cell(ws, r, 8, vat_price, is_number=True)
                safe_write_cell(ws, r, 9, item["note"], wrap=True)
            
        safe_write_cell(ws, 52, 7, supply_ta, is_number=True)
        safe_write_cell(ws, 53, 7, vat_ta, is_number=True)
        safe_write_cell(ws, 54, 7, grand_ta, is_number=True)
        
        # 상단 한글 금액 표기(D17) 및 (\₩ ) 옆 숫자 표기(H17)
        safe_write_cell(ws, 17, 4, num2kor(grand_ta))
        safe_write_cell(ws, 17, 8, grand_ta, is_number=True)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

st.divider()

if st.button("🚀 견적서 3종 엑셀 파일 생성하기", type="primary", use_container_width=True):
    excel_data = generate_excel()
    file_name = f"통합견적서_{project_name}_{issue_date.strftime('%Y%m%d')}.xlsx"
    st.download_button(
        label="📥 생성된 엑셀 파일 다운로드",
        data=excel_data,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )