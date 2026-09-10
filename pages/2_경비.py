import sys
import asyncio

# 🌟 윈도우 환경 asyncio 웹소켓 끊김 방지 코드
if sys.platform == 'win32':
    asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())

import streamlit as st
import pandas as pd
import json
import uuid
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, Border, Side, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from io import BytesIO
import gspread
import plotly.express as px

# ==========================================
# ⚙️ 페이지 설정 및 기본 설정
# ==========================================
st.set_page_config(page_title="부서 경비 정산 시스템", page_icon="💸", layout="wide")

TEAM_BUDGETS = {
    "시운전팀": 1000000, "생산팀": 500000, "판금생산팀": 0, "기술생산설계팀": 0,
    "영업팀": 500000, "영업2팀": 0, "영업3팀": 0, "전장팀": 800000,
    "법카2536": 0, "법카6035": 0, "법카7547": 0, "법카0624": 0
}
TEAMS_LIST = list(TEAM_BUDGETS.keys())

USER_CREDENTIALS = {
    "admin": {"password": "01234", "name": "관리자", "team": "관리자"},
    "시운전": {"password": "1234", "name": "시운전", "team": "시운전팀"},
    "생산": {"password": "1234", "name": "생산", "team": "생산팀"},
    "판금": {"password": "1234", "name": "판금생산", "team": "판금생산팀"},
    "기술생산설계": {"password": "1234", "name": "기술생산설계", "team": "기술생산설계팀"},
    "영업": {"password": "1234", "name": "영업", "team": "영업팀"},
    "영업2": {"password": "1234", "name": "영업2", "team": "영업2팀"},
    "영업3": {"password": "1234", "name": "영업3", "team": "영업3팀"},
    "전장": {"password": "1234", "name": "전장", "team": "전장팀"},
    "법카2536": {"password": "1234", "name": "법카2536", "team": "법카2536"},
    "법카6035": {"password": "1234", "name": "법카6035", "team": "법카6035"},
    "법카7547": {"password": "1234", "name": "법카7547", "team": "법카7547"},
    "법카0624": {"password": "1234", "name": "법카0624", "team": "법카0624"}
}

CATEGORIES = ["교통비", "주차비", "식비", "숙박비", "소모품비", "차량유지비", "운반비", "기타"]
HEADERS = ["trip_id", "order", "team", "date", "user", "place", "content", "items_desc", "total_amount", "details_json"]

# ==========================================
# 🌟 구글 스프레드시트 DB 연동 (보안 Secrets 방식)
# ==========================================
SHEET_URL = "https://docs.google.com/spreadsheets/d/1QBilxwvsIllnve90k438xaSsPSs8jwgQmq1Hpj8xGas/edit?gid=0#gid=0"

@st.cache_resource
def get_gsheet_client(url):
    try:
        creds_json_str = st.secrets["gcp_service_account"]
        creds_dict = json.loads(creds_json_str)
        gc = gspread.service_account_from_dict(creds_dict)
        doc = gc.open_by_url(url)
        return doc.sheet1
    except Exception as e:
        st.error(f"❌ 구글 시트 연결 실패! 스트림릿 Secrets 설정을 다시 확인해 주세요: {e}")
        return None

ws = get_gsheet_client(SHEET_URL)

@st.cache_data(ttl=30)
def get_all_trips():
    if not ws: return []
    try:
        records = ws.get_all_records()
        for r in records:
            r['order'] = int(r.get('order') if r.get('order') else 999)
            r['total_amount'] = int(r.get('total_amount') if r.get('total_amount') else 0)
            r['trip_id'] = str(r.get('trip_id', ''))
        return sorted(records, key=lambda x: x['order'])
    except Exception as e:
        st.error(f"데이터 로드 에러: {e}")
        return []

def clear_cache():
    st.cache_data.clear()

def is_match_team(db_team, target_team):
    if not db_team or not target_team: return False
    return db_team.replace('팀', '').strip() == target_team.replace('팀', '').strip()

# ==========================================
# 🛠️ 팝업 업데이트용 헬퍼 함수
# ==========================================
def _save_trip_updates(trip_id, edit_date, edit_user, edit_place, edit_content, details, team, order):
    total_amount = sum(d['amount'] for d in details)
    desc_parts = [f"{d['category']}: {d['amount']:,}원" for d in details]
    items_desc = " | ".join(desc_parts) if desc_parts else "등록된 영수증 없음"
    
    records = ws.get_all_records()
    for i, r in enumerate(records):
        if str(r.get('trip_id')) == trip_id:
            row_idx = i + 2
            updated_vals = [
                trip_id, order, team, str(edit_date),
                edit_user, edit_place, edit_content, items_desc, total_amount, json.dumps(details, ensure_ascii=False)
            ]
            try: ws.update(f"A{row_idx}:J{row_idx}", [updated_vals], value_input_option='USER_ENTERED')
            except: ws.update(f"A{row_idx}:J{row_idx}", [updated_vals])
            clear_cache()
            break
            
def _delete_trip_entirely(trip_id):
    records = ws.get_all_records()
    for i, r in enumerate(records):
        if str(r.get('trip_id')) == trip_id:
            ws.delete_rows(i + 2)
            clear_cache()
            break

# ==========================================
# 🎁 팝업창(Modal) UI (개별 X 버튼 삭제 구현)
# ==========================================
@st.dialog("✏️ 출장 정산 내역 및 영수증 관리", width="large")
def edit_trip_dialog(trip):
    st.info(f"💡 **{trip['user']}**님의 **{trip['content']}** 내역을 수정 중입니다.")
    
    e_col1, e_col2, e_col3 = st.columns(3)
    edit_date = e_col1.date_input("지출 일자", value=datetime.strptime(trip['date'], '%Y-%m-%d'))
    edit_user = e_col2.text_input("사용자 이름", value=trip['user'])
    edit_place = e_col3.text_input("출장지", value=trip['place'])
    edit_content = st.text_input("출장 목적 및 내용", value=trip['content'])
    
    try: old_details = json.loads(trip['details_json'])
    except: old_details = []
    
    st.markdown("#### 🧾 소속 영수증 관리")
    st.caption("항목 옆의 **❌ 버튼을 누르면 해당 영수증이 즉시 삭제**되며 화면에 반영됩니다.")
    
    updated_details = []
    
    # 🌟 개선포인트 2: X 버튼을 누르면 즉시 삭제되도록 컬럼 기반 UI 적용
    for i, d in enumerate(old_details):
        rc1, rc2, rc3 = st.columns([4, 4, 2])
        cat = rc1.selectbox("경비 구분", options=CATEGORIES, index=CATEGORIES.index(d['category']) if d['category'] in CATEGORIES else 0, key=f"cat_{trip['trip_id']}_{i}", label_visibility="collapsed")
        amt = rc2.number_input("금액(원)", min_value=0, value=int(d['amount']), key=f"amt_{trip['trip_id']}_{i}", label_visibility="collapsed")
        
        if rc3.button("❌ 삭제", key=f"del_{trip['trip_id']}_{i}", use_container_width=True):
            old_details.pop(i)
            _save_trip_updates(trip['trip_id'], edit_date, edit_user, edit_place, edit_content, old_details, trip['team'], trip['order'])
            st.rerun()
        else:
            updated_details.append({"id": d.get('id', f"r_{uuid.uuid4().hex[:6]}"), "category": cat, "amount": amt})
            
    st.markdown("---")
    st.markdown("##### ➕ 새 영수증 항목 추가")
    ac1, ac2, ac3 = st.columns([4, 4, 2])
    new_cat = ac1.selectbox("추가할 경비 구분", options=CATEGORIES, key=f"new_cat_{trip['trip_id']}", label_visibility="collapsed")
    new_amt = ac2.number_input("추가할 금액(원)", min_value=0, value=0, key=f"new_amt_{trip['trip_id']}", label_visibility="collapsed")
    
    if ac3.button("➕ 항목 추가", key=f"add_{trip['trip_id']}", use_container_width=True):
        if new_amt > 0:
            updated_details.append({"id": f"r_{uuid.uuid4().hex[:6]}", "category": new_cat, "amount": new_amt})
            _save_trip_updates(trip['trip_id'], edit_date, edit_user, edit_place, edit_content, updated_details, trip['team'], trip['order'])
            st.rerun()
        else:
            st.warning("금액을 1원 이상 입력해주세요.")

    st.markdown("---")
    c1, c2 = st.columns(2)
    if c1.button("💾 위 수정된 기본정보 모두 저장", type="primary", use_container_width=True):
        _save_trip_updates(trip['trip_id'], edit_date, edit_user, edit_place, edit_content, updated_details, trip['team'], trip['order'])
        st.success("저장되었습니다!")
        st.rerun()
        
    if c2.button("🗑️ 이 출장내역 전체 완전 삭제", use_container_width=True):
        _delete_trip_entirely(trip['trip_id'])
        st.success("전체 삭제되었습니다!")
        st.rerun()

# ==========================================
# 🔐 로그인 화면
# ==========================================
if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False
if 'receipt_rows' not in st.session_state:
    st.session_state['receipt_rows'] = [{"category": "", "amount": 0}]

if not st.session_state['logged_in']:
    st.markdown("<h2 style='text-align: center; color: #1e3a8a;'>🔒 경비 정산 시스템</h2>", unsafe_allow_html=True)
    with st.form("login_form"):
        username = st.text_input("아이디 (부서/계정명)", placeholder="예: 시운전, admin, 생산")
        password = st.text_input("비밀번호", type="password")
        submit = st.form_submit_button("로그인", use_container_width=True)
        
        if submit:
            if username in USER_CREDENTIALS and USER_CREDENTIALS[username]['password'] == password:
                st.session_state['logged_in'] = True
                st.session_state['username'] = USER_CREDENTIALS[username]['name']
                st.session_state['team'] = USER_CREDENTIALS[username]['team']
                st.rerun()
            else:
                st.error("아이디 또는 비밀번호가 올바르지 않습니다.")
    st.stop()

# ==========================================
# 📈 메인 대시보드 화면
# ==========================================
username = st.session_state['username']
team = st.session_state['team']

with st.sidebar:
    st.write(f"👑 **{username}**님 [권한: {team}]")
    if st.button("로그아웃", type="primary"):
        st.session_state['logged_in'] = False
        st.rerun()
        
    st.divider()
    selected_month = st.date_input("📅 마감 월 선택 조회", value=datetime.now().replace(day=1), format="YYYY-MM-DD")
    current_month_str = selected_month.strftime('%Y-%m')
    current_year_str = selected_month.strftime('%Y')

ALL_TRIPS = get_all_trips()

def custom_sort(t):
    t_team, t_date, t_user, t_order = str(t.get('team', '')).strip(), str(t.get('date', '')), str(t.get('user', '')), int(t.get('order', 999))
    if t_team == '시운전팀': return (0, t_user, t_date, t_order)
    else: return (1, t_date, t_order, t_user)

filtered_trips = [t for t in ALL_TRIPS if str(t.get('date', '')).startswith(current_month_str)]
filtered_trips.sort(key=custom_sort)

# 통계 연산
dashboard_stats = {'총합': 0}
for t_name in TEAMS_LIST: dashboard_stats[t_name] = 0

cat_stats = {cat: 0 for cat in CATEGORIES}
place_stats = {}
user_stats = {}

for t in filtered_trips:
    if team != "관리자" and not is_match_team(t.get('team'), team):
        continue
        
    amt = int(t.get('total_amount', 0))
    dashboard_stats['총합'] += amt
    
    t_team = str(t.get('team', '')).strip()
    if t_team not in TEAMS_LIST and t_team + '팀' in TEAMS_LIST: t_team += '팀'
    if t_team in dashboard_stats: dashboard_stats[t_team] += amt
    
    place = t.get('place', '')
    user = t.get('user', '')
    if place: place_stats[place] = place_stats.get(place, 0) + amt
    if user: user_stats[user] = user_stats.get(user, 0) + amt
    
    try: details = json.loads(t.get('details_json', '[]'))
    except: details = []
    for item in details:
        cat = item.get('category', '기타')
        cat_stats[cat] = cat_stats.get(cat, 0) + int(item.get('amount', 0))

st.subheader(f"📊 {current_month_str} 전사 부서/출장지 통합 대시보드" if team == "관리자" else f"📊 {current_month_str} {team} 현황 대시보드")

if team == "관리자":
    cols = st.columns(5)
    cols[0].metric("총 지출합계", f"{dashboard_stats['총합']:,}원")
    cols[1].metric("시운전팀", f"{dashboard_stats['시운전팀']:,}원")
    cols[2].metric("생산팀", f"{dashboard_stats['생산팀']:,}원")
    cols[3].metric("영업팀", f"{dashboard_stats['영업팀']:,}원")
    cols[4].metric("전장팀", f"{dashboard_stats['전장팀']:,}원")
else:
    st.metric(f"선택 월 지출 총합", f"{dashboard_stats['총합']:,}원")

def process_top_10(data_dict):
    sorted_items = sorted(data_dict.items(), key=lambda x: x[1], reverse=True)
    if len(sorted_items) <= 10: return sorted_items
    top10 = sorted_items[:10]
    others = sum(x[1] for x in sorted_items[10:])
    top10.append(("기타 합산", others))
    return top10

col_c1, col_c2 = st.columns(2)
with col_c1:
    top_places = process_top_10(place_stats)
    if top_places:
        df_places = pd.DataFrame(top_places, columns=["출장지", "금액"]).sort_values("금액", ascending=True)
        fig_place = px.bar(df_places, x="금액", y="출장지", orientation='h', title="📍 출장지 순위 (상위 10위)", text_auto='.2s', color_discrete_sequence=['#3b82f6'])
        st.plotly_chart(fig_place, use_container_width=True)
        
    if sum(cat_stats.values()) > 0:
        fig_pie = px.pie(names=list(cat_stats.keys()), values=list(cat_stats.values()), title="🧾 항목별 비용 비율", hole=0.3)
        st.plotly_chart(fig_pie, use_container_width=True)

with col_c2:
    top_users = process_top_10(user_stats)
    if top_users:
        df_users = pd.DataFrame(top_users, columns=["사용자", "금액"]).sort_values("금액", ascending=True)
        fig_user = px.bar(df_users, x="금액", y="사용자", orientation='h', title="👤 인원별 지출 순위 (상위 10위)", text_auto='.2s', color_discrete_sequence=['#10b981'])
        st.plotly_chart(fig_user, use_container_width=True)

    year_months = [f"{current_year_str}-{str(i).zfill(2)}" for i in range(1, 13)]
    yearly_trend = {m: 0 for m in year_months}
    
    for t in ALL_TRIPS:
        if team != "관리자" and not is_match_team(t.get('team'), team): continue
        date_str = str(t.get('date', ''))
        if date_str.startswith(current_year_str):
            ym = date_str[:7]
            if ym in yearly_trend:
                yearly_trend[ym] += int(t.get('total_amount', 0))
                
    df_trend = pd.DataFrame({
        "월": [f"{i}월" for i in range(1, 13)],
        "지출 금액": [yearly_trend[m] for m in year_months]
    })
    fig_trend = px.line(df_trend, x="월", y="지출 금액", title=f"📈 {current_year_str}년도 전체 추이 그래프", markers=True, color_discrete_sequence=['#8b5cf6'])
    fig_trend.update_traces(fill='tozeroy')
    st.plotly_chart(fig_trend, use_container_width=True)

st.divider()

# ==========================================
# 📝 경비 지출 신청 (등록 폼)
# ==========================================
st.subheader("📝 통합 출장 경비 신청 (영수증 일괄 등록)")
with st.form("add_expense_form"):
    col1, col2, col3, col4 = st.columns(4)
    target_team = team
    if team == "관리자":
        target_team = col1.selectbox("부서 선택 (대리 등록)", TEAMS_LIST)
    else:
        col1.text_input("부서", value=team, disabled=True)
        
    expense_date = col2.date_input("지출 일자", value=selected_month)
    user_name = col3.text_input("사용자 (이름)")
    place = col4.text_input("출장지")
    
    content = st.text_input("출장 목적 및 내용 (표지 제목)")
    
    st.markdown("**🧾 영수증 내역 입력**")
    receipt_df = pd.DataFrame(st.session_state['receipt_rows'])
    edited_receipts = st.data_editor(
        receipt_df,
        num_rows="dynamic",
        column_config={
            "category": st.column_config.SelectboxColumn("경비 구분", options=CATEGORIES, required=True),
            "amount": st.column_config.NumberColumn("금액(원)", min_value=0, required=True)
        },
        use_container_width=True
    )
    
    submit_btn = st.form_submit_button("💾 상기 모든 영수증 저장 및 제출하기", use_container_width=True)
    
    if submit_btn:
        details, total_amount, desc_parts = [], 0, []
        for _, row in edited_receipts.iterrows():
            if pd.notna(row['category']) and row['category'] != "" and row['amount'] > 0:
                details.append({"id": f"r_{uuid.uuid4().hex[:6]}", "category": row['category'], "amount": int(row['amount'])})
                total_amount += int(row['amount'])
                desc_parts.append(f"{row['category']}: {int(row['amount']):,}원")
                
        items_desc = " | ".join(desc_parts) if desc_parts else "등록된 영수증 없음"
        
        new_trip = [
            str(uuid.uuid4().hex[:8]), 999, target_team, str(expense_date), user_name, place, content, 
            items_desc, total_amount, json.dumps(details, ensure_ascii=False)
        ]
        
        try:
            ws.append_row(new_trip, value_input_option='USER_ENTERED')
        except:
            ws.append_row(new_trip)
            
        clear_cache()
        st.success("✅ 성공적으로 등록되었습니다!")
        st.rerun()

st.divider()

# ==========================================
# 📋 정산 목록 리스트 (체크박스 일괄 삭제 및 팝업 연동)
# ==========================================
st.subheader(f"📋 {current_month_str} 등록된 출장 정산 목록")
st.info("💡 **항목을 삭제하려면 왼쪽 체크박스(☑️)를 선택, 수정하려면 항목 1개를 선택 후 아래 버튼을 눌러주세요.**")
if team == "관리자": st.caption("전사 부서 전체 통합 노출 모드")

display_trips = [t for t in filtered_trips if team == "관리자" or is_match_team(t.get('team'), team)]

if display_trips:
    df_display = pd.DataFrame(display_trips)
    df_view = df_display[['team', 'date', 'user', 'place', 'content', 'items_desc', 'total_amount']].copy()
    df_view.columns = ['신청부서', '지출일자', '사용자', '출장지', '내용', '포함된 영수증 항목', '총 금액(원)']
    
    # 🌟 개선포인트 1: 목록 맨 앞에 '선택' 체크박스 열 추가
    df_view.insert(0, "선택", False)
    
    edited_list = st.data_editor(
        df_view,
        use_container_width=True,
        hide_index=True,
        column_config={"선택": st.column_config.CheckboxColumn("☑️ 선택", default=False)}
    )
    
    # 선택된 항목 추출
    selected_rows = edited_list[edited_list["선택"] == True]
    
    if len(selected_rows) > 0:
        st.markdown("---")
        c1, c2 = st.columns(2)
        
        if len(selected_rows) == 1:
            selected_idx = selected_rows.index[0]
            selected_trip_data = display_trips[selected_idx]
            with c1:
                # 1개 선택 시에만 팝업 버튼 활성화
                if st.button("✏️ 선택한 내역 수정/상세보기", use_container_width=True):
                    edit_trip_dialog(selected_trip_data)
        else:
            with c1:
                st.warning(f"ℹ️ {len(selected_rows)}개 항목이 선택되었습니다. (수정은 1개씩만 가능합니다)")
        
        with c2:
            # 일괄 삭제 버튼
            if st.button(f"🗑️ 선택한 {len(selected_rows)}개 내역 일괄 삭제", type="primary", use_container_width=True):
                trip_ids_to_delete = [display_trips[i]['trip_id'] for i in selected_rows.index]
                records = ws.get_all_records()
                
                rows_to_delete = []
                for i, r in enumerate(records):
                    if str(r.get('trip_id')) in trip_ids_to_delete:
                        rows_to_delete.append(i + 2)
                
                # 구글 시트 삭제 시 행 번호가 꼬이는 것을 막기 위해 아래(높은 번호)부터 지웁니다
                for r_idx in sorted(rows_to_delete, reverse=True):
                    ws.delete_rows(r_idx)
                    
                clear_cache()
                st.success(f"✅ {len(selected_rows)}개 항목이 정상적으로 일괄 삭제되었습니다!")
                st.rerun()
        
else:
    st.info("해당 월에 등록된 정산 내역이 없습니다.")

# ==========================================
# 📥 엑셀 다운로드
# ==========================================
st.divider()
st.subheader("📥 엑셀 다운로드")

def generate_excel(raw_data, target_team, target_month):
    wb = openpyxl.Workbook()
    font_title = Font(name='맑은 고딕', size=18, bold=True)
    font_header = Font(name='맑은 고딕', size=11, bold=True)
    font_main = Font(name='맑은 고딕', size=10)
    font_sum = Font(name='맑은 고딕', size=11, bold=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    fill_header = PatternFill(start_color='F3F4F6', end_color='F3F4F6', fill_type='solid')
    align_center = Alignment(horizontal='center', vertical='center')
    align_right = Alignment(horizontal='right', vertical='center')
    
    ws1 = wb.active
    ws1.title = f"{target_month[5:7]}월 정산서"
    ws1.merge_cells('A1:E2')
    ws1['A1'] = f"{target_month[5:7]}월 경비 사용내역서"
    ws1['A1'].font = font_title
    
    display_categories = ["교통비", "식비", "숙박비", "소모품비", "차량유지비", "기타"]
    headers1 = ["순번", "일자", "내 용", "출장지", "금액(합계)"] + display_categories + ["사용자"]
    
    for col_idx, h in enumerate(headers1, 1):
        cell = ws1.cell(row=5, column=col_idx, value=h)
        cell.font = font_header; cell.alignment = align_center; cell.border = thin_border; cell.fill = fill_header

    r_idx = 6
    for idx, trip in enumerate(raw_data, 1):
        ws1.cell(row=r_idx, column=1, value=idx).alignment = align_center
        ws1.cell(row=r_idx, column=2, value=str(trip.get('date', ''))).alignment = align_center
        ws1.cell(row=r_idx, column=3, value=trip.get('content', ''))
        ws1.cell(row=r_idx, column=4, value=trip.get('place', ''))
        ws1.cell(row=r_idx, column=5, value=int(trip.get('total_amount', 0))).number_format = '#,##0'
        
        cat_sums = {c: 0 for c in display_categories}
        try:
            details = json.loads(trip.get('details_json', '[]'))
            for item in details:
                c_name = item.get('category', '기타')
                amt = int(item.get('amount', 0))
                if c_name in ['교통비', '주차비']: cat_sums['교통비'] += amt
                elif c_name in ['운반비', '기타']: cat_sums['기타'] += amt
                elif c_name in cat_sums: cat_sums[c_name] += amt
                else: cat_sums['기타'] += amt
        except: pass
        
        for c_idx, cat_name in enumerate(display_categories, 6):
            cell = ws1.cell(row=r_idx, column=c_idx, value=cat_sums[cat_name] if cat_sums[cat_name] > 0 else "")
            cell.number_format = '#,##0'
            
        ws1.cell(row=r_idx, column=12, value=trip.get('user', '')).alignment = align_center
        
        for c in range(1, 13): ws1.cell(row=r_idx, column=c).border = thin_border
        r_idx += 1
        
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

if display_trips:
    col_dl1, col_dl2 = st.columns(2)
    if team == "관리자":
        with col_dl1:
            excel_data_all = generate_excel(filtered_trips, 'ALL', current_month_str)
            st.download_button(label="🌐 전사 통합 엑셀 다운로드", data=excel_data_all, file_name=f"전사통합_정산서_{current_month_str}.xlsx", type="primary")
        with col_dl2:
            target_dl_team = st.selectbox("부서별 다운로드", TEAMS_LIST)
            team_trips = [t for t in filtered_trips if is_match_team(t.get('team'), target_dl_team)]
            excel_data_team = generate_excel(team_trips, target_dl_team, current_month_str)
            st.download_button(label=f"📥 {target_dl_team} 엑셀 다운로드", data=excel_data_team, file_name=f"{target_dl_team}_정산서_{current_month_str}.xlsx")
    else:
        excel_data = generate_excel(display_trips, team, current_month_str)
        st.download_button(label="📥 소속 부서 엑셀 다운로드", data=excel_data, file_name=f"{team}_정산서_{current_month_str}.xlsx", type="primary")
