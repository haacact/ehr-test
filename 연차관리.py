import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta
import calendar
import re
import math
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
import io
import json
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- [구글 시트 연동 설정] ---
SPREADSHEET_NAME = "vacation_data"     

# --- [사내 아웃룩 연동] 메일 발송 설정 ---
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
SENDER_EMAIL = "haacact@gmail.com"          
SENDER_PASSWORD = "여기에_16자리_앱비밀번호_입력"

@st.cache_resource
def get_gspread_client():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds_dict = json.loads(st.secrets["gcp_service_account"])
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    return client

@st.cache_data(ttl=60)
def get_available_years():
    try:
        client = get_gspread_client()
        sheet = client.open(SPREADSHEET_NAME)
        worksheets = sheet.worksheets()
        years = set([2026])
        for ws in worksheets:
            if ws.title.startswith("Employees_"):
                try:
                    yr = int(ws.title.split("_")[1])
                    years.add(yr)
                except: pass
        return sorted(list(years))
    except:
        return [2026]

@st.cache_data(ttl=30)
def load_data(year):
    if "gcp_service_account" not in st.secrets:
        st.error("❌ Streamlit Secrets에 구글 API 키 정보가 없습니다.")
        st.stop()
        
    try:
        client = get_gspread_client()
        sheet = client.open(SPREADSHEET_NAME)
        
        emp_ws_name = f"Employees_{year}"
        try: ws_emp = sheet.worksheet(emp_ws_name)
        except:
            if year == 2026:
                try: ws_emp = sheet.worksheet("Employees")
                except: ws_emp = None
            else: ws_emp = None
            
        if ws_emp:
            emp_data = ws_emp.get_all_values()
            df_emp = pd.DataFrame(emp_data[1:], columns=emp_data[0]) if emp_data else pd.DataFrame()
        else:
            df_emp = pd.DataFrame()

        required_emp_cols = ['ID', 'PASSWORD', '이름', '팀', '파트', '직급', 'permission', 'EMAIL', '입사일', '연차기초', '사용', '연차계획', '연차잔액']
        if not df_emp.empty:
            for col in required_emp_cols:
                if col not in df_emp.columns: df_emp[col] = "" 
            df_emp['ID'] = df_emp['ID'].astype(str)
            df_emp['PASSWORD'] = df_emp['PASSWORD'].astype(str)
            for col in ['연차기초', '사용', '연차계획', '연차잔액']:
                df_emp[col] = pd.to_numeric(df_emp[col].replace('', 0), errors='coerce').astype(float)
        
        plan_ws_name = f"PLANS_{year}"
        try: ws_plans = sheet.worksheet(plan_ws_name)
        except:
            if year == 2026:
                try: ws_plans = sheet.worksheet("PLANS")
                except: ws_plans = None
            else: ws_plans = None
            
        if ws_plans:
            plan_data = ws_plans.get_all_values()
            df_plans = pd.DataFrame(plan_data[1:], columns=plan_data[0]) if plan_data else pd.DataFrame()
        else:
            df_plans = pd.DataFrame()
            
        if not df_plans.empty:
            required_plan_cols = ['ID', 'Emp_ID', 'Date', 'Status', 'Type', 'Reason', 'Manager_Sign', 'Part_Sign', 'Apply_Time', 'Approve_Time', 'Reminder_Sent']
            for col in required_plan_cols:
                if col not in df_plans.columns: df_plans[col] = "" 
            df_plans['Date'] = df_plans['Date'].astype(str)
            df_plans['Emp_ID'] = df_plans['Emp_ID'].astype(str)
            df_plans['Reason'] = df_plans['Reason'].fillna("").astype(str)
            df_plans['Manager_Sign'] = df_plans['Manager_Sign'].fillna("").astype(str)
            df_plans['Part_Sign'] = df_plans['Part_Sign'].fillna("").astype(str)
            df_plans['Apply_Time'] = df_plans['Apply_Time'].fillna("").astype(str)
            df_plans['Approve_Time'] = df_plans['Approve_Time'].fillna("").astype(str)
            df_plans['Reminder_Sent'] = df_plans['Reminder_Sent'].fillna("").astype(str)
                
        return df_emp, df_plans
    except Exception as e:
        st.error(f"❌ 구글 시트 데이터 로드 오류: {e}")
        st.stop()

@st.cache_data(ttl=30)
def load_notices():
    try:
        client = get_gspread_client()
        sheet = client.open(SPREADSHEET_NAME)
        ws_notices = sheet.worksheet("NOTICES")
        notice_data = ws_notices.get_all_values()
        df_notices = pd.DataFrame(notice_data[1:], columns=notice_data[0]) if notice_data else pd.DataFrame()
        
        for col in ["ID", "날짜", "제목", "내용"]:
            if col not in df_notices.columns: df_notices[col] = ""
        return df_notices
    except:
        return pd.DataFrame(columns=["ID", "날짜", "제목", "내용"])

def update_sheet(ws_name, df):
    try:
        client = get_gspread_client()
        sheet = client.open(SPREADSHEET_NAME)
        try: ws = sheet.worksheet(ws_name)
        except gspread.exceptions.WorksheetNotFound:
            ws = sheet.add_worksheet(title=ws_name, rows="100", cols="20")
            
        df_clean = df.fillna("").astype(str)
        data = [df_clean.columns.values.tolist()] + df_clean.values.tolist()
        ws.clear()
        try: ws.update(values=data, value_input_option='RAW')
        except TypeError: ws.update('A1', data, value_input_option='RAW')
        return True
    except Exception as e:
        st.error(f"❌ {ws_name} 저장 실패: {e}")
        return False

def get_ws_names(year):
    e_name = "Employees" if year == 2026 else f"Employees_{year}"
    p_name = "PLANS" if year == 2026 else f"PLANS_{year}"
    return e_name, p_name

def save_plans_only(df_plans, year):
    _, p_name = get_ws_names(year)
    if update_sheet(p_name, df_plans):
        load_data.clear()
        return True
    return False

def save_emp_and_plans(df_emp, df_plans, year):
    e_name, p_name = get_ws_names(year)
    s1 = update_sheet(e_name, df_emp)
    s2 = update_sheet(p_name, df_plans)
    if s1 and s2:
        load_data.clear()
        return True
    return False

def save_notices_only(df_notices):
    if update_sheet("NOTICES", df_notices):
        load_notices.clear()
        return True
    return False

def save_emp_only(df_emp, year):
    e_name, _ = get_ws_names(year)
    if update_sheet(e_name, df_emp):
        load_data.clear()
        return True
    return False

def calculate_vacation_accrual(join_date_str, target_year):
    try:
        j_dt = datetime.strptime(str(join_date_str).strip(), "%Y-%m-%d")
        total_months = (target_year - j_dt.year) * 12 + (1 - j_dt.month)
        if 1 < j_dt.day: total_months -= 1
        if total_months < 0: return 0.0 
        full_years = total_months // 12
        if full_years >= 1:
            calculated = 15 + math.floor((full_years - 1) / 2)
            return float(min(calculated, 25))
        else:
            prorated = 15 * (total_months / 12.0)
            return float(math.floor(prorated + 0.5))
    except:
        return 15.0 

def send_vacation_reminder_email(to_email, emp_name, date_str, v_type):
    if not to_email or "@" not in to_email:
        return False
    subject = f"[리마인드] {emp_name}님, {date_str} [{v_type}] 일주일 전 안내입니다."
    body = f"안녕하세요. {emp_name}님,\n\n신청하신 [{v_type}] 일정이 약 일주일 앞으로 다가와 안내해 드립니다.\n\n■ 일 자 : {date_str}\n■ 구 분 : {v_type}\n\n휴가 전 업무 인수인계를 잘 마무리하시고, 즐겁고 편안한 시간 보내시길 바랍니다!\n(※ '연차계획'으로 신청하신 경우, 시스템에 접속하여 실제 '연차'로 확정 변경을 부탁드립니다.)\n\n- 하이에어공조(주) 시스템 관리자 드림 -"
    try:
        msg = MIMEMultipart()
        msg['From'] = SENDER_EMAIL
        msg['To'] = to_email
        msg['Subject'] = Header(subject, 'utf-8')
        msg.attach(MIMEText(body, 'plain', 'utf-8'))
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.sendmail(SENDER_EMAIL, to_email, msg.as_string())
        server.quit()
        return True
    except:
        return False

def auto_convert_expired_plans(df_emp, df_plans, year):
    today_date = datetime.now().date()
    needs_save = False
    for idx, row in df_plans.iterrows():
        if row['Type'] == '연차계획' and row['Status'] == '승인':
            try:
                plan_date = datetime.strptime(str(row['Date']).strip(), "%Y-%m-%d").date()
                if (today_date - plan_date).days >= 3:
                    df_plans.at[idx, 'Type'] = '연차'
                    target_emp_id = str(row['Emp_ID'])
                    df_emp.loc[df_emp["ID"].astype(str) == target_emp_id, "사용"] += 1.0
                    df_emp.loc[df_emp["ID"].astype(str) == target_emp_id, "연차잔액"] -= 1.0
                    df_emp.loc[df_emp["ID"].astype(str) == target_emp_id, "연차계획"] -= 1.0
                    needs_save = True
            except: pass
    if needs_save:
        save_emp_and_plans(df_emp, df_plans, year)
    return df_emp, df_plans

# 🚀 관리자 전용 수동 알림 발송 처리 함수
def execute_manual_reminders(df_emp, df_plans, year):
    today_date = datetime.now().date()
    success_count = 0
    for idx, row in df_plans.iterrows():
        if row['Status'] != '반려' and str(row.get('Reminder_Sent', '')) != 'Y' and row['Type'].strip() != "":
            try:
                plan_date = datetime.strptime(str(row['Date']).strip(), "%Y-%m-%d").date()
                if 1 <= (plan_date - today_date).days <= 7:
                    emp_id = str(row['Emp_ID'])
                    emp_info = df_emp[df_emp['ID'].astype(str) == emp_id]
                    if not emp_info.empty:
                        emp_email = emp_info.iloc[0].get('EMAIL', '')
                        emp_name = emp_info.iloc[0]['이름']
                        if emp_email and "@" in str(emp_email):
                            if send_vacation_reminder_email(emp_email, emp_name, row['Date'], row['Type']):
                                df_plans.at[idx, 'Reminder_Sent'] = 'Y'
                                success_count += 1
                        else:
                            df_plans.at[idx, 'Reminder_Sent'] = 'E'
            except: pass
    if success_count > 0:
        save_plans_only(df_plans, year)
    return success_count

st.set_page_config(page_title="사내 연차 관리 시스템", layout="wide")

st.sidebar.page_link("https://hiairac-expense-sysem.onrender.com/", label="경비 시스템 가기", icon="💸")
st.sidebar.divider()

available_years = get_available_years()

if 'logged_in' not in st.session_state: st.session_state['logged_in'] = False
if 'user_info' not in st.session_state: st.session_state['user_info'] = None
if 'selected_year' not in st.session_state:
    now_y = datetime.now().year
    st.session_state['selected_year'] = now_y if now_y in available_years else available_years[-1]

if not st.session_state['logged_in']:
    st.title("🔐 사내 연차 관리 시스템")
    with st.form("login"):
        s_year = st.selectbox("📅 접속 연도 선택", available_years, index=available_years.index(st.session_state['selected_year']) if st.session_state['selected_year'] in available_years else len(available_years)-1)
        i_id, i_pw = st.text_input("ID(사번)"), st.text_input("비밀번호", type="password")
        if st.form_submit_button("로그인"):
            df_emp_check, _ = load_data(s_year)
            if df_emp_check.empty: st.error(f"❌ {s_year}년도 데이터가 아직 생성되지 않았습니다.")
            else:
                i_pw_stripped = i_pw.lstrip('0') if i_pw.lstrip('0') != '' else i_pw
                user = df_emp_check[(df_emp_check['ID'] == i_id) & ((df_emp_check['PASSWORD'] == i_pw) | (df_emp_check['PASSWORD'] == i_pw_stripped))]
                if not user.empty:
                    st.session_state.update({'logged_in': True, 'user_info': user.iloc[0], 'selected_year': s_year})
                    st.rerun()
                else: st.error("정보가 올바르지 않습니다.")
    st.stop()

sel_year = st.session_state['selected_year']
df_emp, df_plans = load_data(sel_year)

# 🚀 [가벼운 연동] 날짜 만료 체크만 수행 (메일 발송 부하 완전 제거)
df_emp, df_plans = auto_convert_expired_plans(df_emp, df_plans, sel_year)

user_info = df_emp[df_emp['ID'] == st.session_state['user_info']['ID']].iloc[0]
sub_title_str = f"{user_info['팀']}" + (f" / {user_info['파트']}" if user_info['파트'] else "") + (f" / {user_info['직급']}" if user_info['직급'] else "")

st.sidebar.title(f"👤 {user_info['이름']} ({user_info['permission']})")
st.sidebar.caption(sub_title_str)

new_year = st.sidebar.selectbox("📅 연도 전환", available_years, index=available_years.index(sel_year) if sel_year in available_years else len(available_years)-1)
if new_year != sel_year:
    st.session_state['selected_year'] = new_year
    st.rerun()
st.sidebar.divider()

menu = ["📢 공지사항(연차촉진)", "🏠 내 연차 신청/현황", "📑 신청서 출력", "📅 연차 현황 달력"]
if user_info['permission'] in ["파트장", "팀장", "총괄", "관리자"]: menu += ["✅ 팀원 결재 관리 (검토/승인)"]
if user_info['permission'] in ["파트장", "팀장", "관리자"]: menu += ["📊 부서/전사 모니터링"]
if user_info['permission'] == "관리자": menu += ["🌐 [관리자] 전사 통합 관리"]

choice = st.sidebar.radio("메뉴 이동", menu)
if st.sidebar.button("로그아웃"):
    st.session_state['logged_in'] = False
    st.rerun()

# --- 📢 공지사항 ---
if choice == "📢 공지사항(연차촉진)":
    st.header(f"📢 {sel_year}년 전사 공지사항 (연차촉진 안내)")
    df_notices = load_notices()
    if df_notices.empty or df_notices["제목"].str.strip().eq("").all(): st.info("현재 등록된 공지사항이 없습니다.")
    else:
        for idx, row in df_notices[df_notices["제목"].str.strip() != ""].iloc[::-1].iterrows():
            with st.expander(f"📌 [{row['날짜']}] {row['제목']}", expanded=True):
                st.write(row['내용'])
                st.caption("작성자: 시스템 관리자")

# --- 🏠 내 연차 신청/현황 ---
elif choice == "🏠 내 연차 신청/현황":
    st.header(f"📅 나의 {sel_year}년 연차 현황")
    if st.session_state.get('apply_success'):
        msg_txt = "총괄 승인 후 최종 반영됩니다." if user_info['permission'] == "팀장" else "파트장 검토 및 팀장 승인 후 최종 반영됩니다."
        st.components.v1.html(f"<script>alert('🎉 신청이 성공적으로 완료되었습니다! {msg_txt}');</script>", height=0, width=0)
        st.success(f"✅ 연차(휴가) 신청이 성공적으로 완료되었습니다! ({msg_txt})")
        st.session_state['apply_success'] = False
        
    if st.session_state.get('cancel_success'):
        st.components.v1.html("<script>alert('🗑️ 신청 내역이 취소되었습니다.');</script>", height=0, width=0)
        st.warning("🗑️ 신청 내역이 정상적으로 취소되었습니다.")
        st.session_state['cancel_success'] = False

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("기초 연차", f"{user_info['연차기초']}일")
    c2.metric("사용 완료", f"{user_info['사용']}일")
    c3.metric("연차 계획", f"{user_info['연차계획']}일")
    c4.metric("남은 잔액", f"{user_info['연차잔액']}일")
    st.divider()
    
    col_l, col_r = st.columns([1, 2])
    with col_l:
        st.subheader("📝 신규 신청")
        today = datetime.now().date()
        v_date = st.date_input("날짜 선택", value=(today, today))
        is_multi_day = isinstance(v_date, (tuple, list)) and len(v_date) == 2 and v_date[0] != v_date[1]
        type_options = ["연차", "연차계획", "휴가", "교육/훈련"] if is_multi_day else ["연차", "오전반차", "오후반차", "연차계획", "휴가", "교육/훈련"]
        v_type = st.selectbox("구분", type_options)
        v_reason = st.text_input("✍️ 신청 사유", placeholder="예: 개인 용무, 직무 교육 위탁 참여 등")
        
        if st.button("신청서 제출하기"):
            if isinstance(v_date, (tuple, list)):
                start_date, end_date = v_date[0], v_date[1] if len(v_date) == 2 else v_date[0]
            else: start_date = end_date = v_date
                
            if start_date and end_date:
                date_list = [(start_date + timedelta(days=i)).strftime("%Y-%m-%d") for i in range((end_date - start_date).days + 1) if (start_date + timedelta(days=i)).weekday() < 5]
                if not date_list: st.error("❌ 선택한 기간에 평일이 없습니다.")
                else:
                    dup_dates = [d_str for d_str in date_list if not df_plans[(df_plans['Emp_ID'] == user_info['ID']) & (df_plans['Date'] == d_str) & (df_plans['Status'] != '반려')].empty]
                    if dup_dates: st.error(f"❌ 이미 신청된 날짜가 포함되어 있습니다: {', '.join(dup_dates)}")
                    else: st.session_state.update({'confirm_apply': True, 'temp_dates': date_list, 'temp_type': v_type, 'temp_reason': v_reason if v_reason.strip() else "개인 용무"})
            else: st.warning("시작 날짜와 종료 날짜를 모두 선택해주세요.")

        if st.session_state.get('confirm_apply'):
            temp_dates = st.session_state['temp_dates']
            st.warning(f"⚠️ {temp_dates[0]} ~ {temp_dates[-1]} [{st.session_state['temp_type']}] 신청하시겠습니까?")
            if st.button("✅ 최종 확인"):
                new_id = int(pd.to_numeric(df_plans["ID"], errors='coerce').max() + 1) if not df_plans.empty else 1
                now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                new_rows = [{"ID": new_id + i, "Emp_ID": user_info['ID'], "Date": d_str, "Status": "대기", "Type": st.session_state['temp_type'], "Reason": st.session_state['temp_reason'], "Manager_Sign": "", "Part_Sign": "", "Apply_Time": now_str, "Approve_Time": "", "Reminder_Sent": ""} for i, d_str in enumerate(temp_dates)]
                df_plans = pd.concat([df_plans, pd.DataFrame(new_rows)], ignore_index=True)
                if save_plans_only(df_plans, sel_year):
                    st.session_state['apply_success'] = True
                    del st.session_state['confirm_apply']
                    st.rerun()

    with col_r:
        st.subheader(f"🔍 나의 {sel_year}년 신청 내역")
        my_h = df_plans[(df_plans['Emp_ID'] == user_info['ID']) & (df_plans['Type'] != "")].sort_values(by="Date", ascending=False)
        for idx, row in my_h.iterrows():
            cols = st.columns([3, 2, 3])
            cols[0].write(f"📅 {row['Date']} ({row['Type']})")
            cols[1].write(f"상태: {row['Status']}")
            with cols[2]:
                if row['Status'] in ['대기', '검토완료'] or (row['Type'] == '연차계획' and row['Status'] == '승인'):
                    if st.button("❌ 취소", key=f"cancel_{row['ID']}"):
                        if row['Type'] == '연차계획' and row['Status'] == '승인':
                            df_emp.loc[df_emp["ID"] == user_info['ID'], "연차계획"] -= 1.0
                            df_plans = df_plans[df_plans['ID'].astype(str) != str(row['ID'])]
                            if save_emp_and_plans(df_emp, df_plans, sel_year): st.session_state['cancel_success'] = True; st.rerun()
                        else:
                            df_plans = df_plans[df_plans['ID'].astype(str) != str(row['ID'])]
                            if save_plans_only(df_plans, sel_year): st.session_state['cancel_success'] = True; st.rerun()
                                
                if row['Type'] == "연차계획":
                    if st.button("✅ 연차로 확정(변경)", key=f"btn_{row['ID']}"):
                        df_plans.at[idx, "Type"] = "연차"
                        if row['Status'] == "승인": df_emp.loc[df_emp["ID"] == user_info['ID'], ["사용","연차잔액","연차계획"]] += [1.0, -1.0, -1.0]
                        if save_emp_and_plans(df_emp, df_plans, sel_year): st.rerun()

# --- 📑 신청서 출력 ---
elif choice == "📑 신청서 출력":
    st.header("🖨️ 연차 신청서 출력")
    my_valid_plans = df_plans[(df_plans['Emp_ID'] == user_info['ID']) & (df_plans['Type'] != "") & (df_plans['Status'] != '반려')].sort_values(by="Date", ascending=False)
    if my_valid_plans.empty: st.info("출력 가능한 연차 내역이 없습니다.")
    else:
        s_doc = st.selectbox("출력할 항목을 선택하세요", my_valid_plans.apply(lambda x: f"[{x['ID']}] {x['Date']} ({x['Type']}) - {x['Status']}", axis=1).tolist())
        doc = my_valid_plans[my_valid_plans['ID'].astype(str) == str(s_doc.split(']')[0].replace('[', ''))].iloc[0]
        sign_table_html = f"""<table style="border-collapse: collapse; border: 1px solid black; text-align: center; color: black;"><tr><th rowspan="2" style="border: 1px solid black; padding: 5px; width: 30px; background: #f2f2f2; font-size: 13px;">결<br>재</th>""" + (f"""<th style="border: 1px solid black; padding: 5px; width: 85px; background: #f2f2f2; font-size: 13px;">담당(팀장)</th><th style="border: 1px solid black; padding: 5px; width: 85px; background: #f2f2f2; font-size: 13px;">총괄승인</th>""" if user_info['permission'] == "팀장" else f"""<th style="border: 1px solid black; padding: 5px; width: 75px; background: #f2f2f2; font-size: 13px;">담당</th><th style="border: 1px solid black; padding: 5px; width: 75px; background: #f2f2f2; font-size: 13px;">파트검토</th><th style="border: 1px solid black; padding: 5px; width: 75px; background: #f2f2f2; font-size: 13px;">팀장승인</th>""") + f"""<th style="border: 1px solid black; padding: 5px; width: 75px; background: #f2f2f2; font-size: 13px;">대표승인</th></tr><tr><td style="border: 1px solid black; height: 55px; font-weight: bold; vertical-align: middle; font-size: 14px;">{user_info['이름']}</td>""" + (f"""<td style="border: 1px solid black; height: 55px; font-weight: bold; vertical-align: middle; color: blue; font-size: 14px;">{doc['Manager_Sign']}</td>""" if user_info['permission'] == "팀장" else f"""<td style="border: 1px solid black; height: 55px; font-weight: bold; vertical-align: middle; color: green; font-size: 14px;">{doc['Part_Sign']}</td><td style="border: 1px solid black; height: 55px; font-weight: bold; vertical-align: middle; color: blue; font-size: 14px;">{doc['Manager_Sign']}</td>""") + f"""<td style="border: 1px solid black; height: 55px; vertical-align: middle;"></td></tr></table>"""
        html_template = f"""<div style="border: 1px solid #000; padding: 40px; background-color: white; color: black; font-family: 'Malgun Gothic'; width: 700px; margin: 0 auto; position: relative;"><div style="display: flex; justify-content: flex-end;">{sign_table_html}</div><h1 style="text-align: center; margin-top: 15px; color: black; font-size: 28px; letter-spacing: 5px;">연 차 휴 가 신 청 서</h1><br><br><table style="width: 100%; border-collapse: collapse; border: 1px solid black; color: black; font-size: 14px;"><tr><th style="border: 1px solid black; padding: 12px; background: #f2f2f2; width: 20%; font-weight: bold;">성 명</th><td style="border: 1px solid black; padding: 12px; width: 30%;">{user_info['이름']}</td><th style="border: 1px solid black; padding: 12px; background: #f2f2f2; width: 20%; font-weight: bold;">사 번</th><td style="border: 1px solid black; padding: 12px; width: 30%;">{user_info['ID']}</td></tr><tr><th style="border: 1px solid black; padding: 12px; background: #f2f2f2; font-weight: bold;">부 서</th><td style="border: 1px solid black; padding: 12px;">{user_info['팀']}</td><th style="border: 1px solid black; padding: 12px; background: #f2f2f2; font-weight: bold;">파 트</th><td style="border: 1px solid black; padding: 12px;">{user_info['파트'] if user_info['파트'] else "-"}</td></tr><tr><th style="border: 1px solid black; padding: 12px; background: #f2f2f2; font-weight: bold;">직 급</th><td colspan="3" style="border: 1px solid black; padding: 12px;">{user_info['직급'] if user_info['직급'] else "-"}</td></tr><tr><th style="border: 1px solid black; padding: 12px; background: #f2f2f2; font-weight: bold;">휴가 일자</th><td colspan="3" style="border: 1px solid black; padding: 12px;">{doc['Date']}</td></tr><tr><th style="border: 1px solid black; padding: 12px; background: #f2f2f2; font-weight: bold;">휴가 구분</th><td colspan="3" style="border: 1px solid black; padding: 12px;">{doc['Type']}</td></tr><tr><th style="border: 1px solid black; padding: 12px; background: #f2f2f2; height: 100px; font-weight: bold;">신청 사유</th><td colspan="3" style="border: 1px solid black; padding: 12px; vertical-align: top;">{doc['Reason'] if doc['Reason'] else "개인 용무"}</td></tr></table><br><br><p style="text-align: center; margin-top: 40px; font-size: 16px; color: black;">위와 같이 연차 휴가를 신청하오니 승인하여 주시기 바랍니다.</p><p style="text-align: center; margin-top: 30px; font-size: 14px; color: black;">{datetime.now().strftime('%Y년 %m월 %d일')}</p><br><div style="text-align: right; margin-top: 20px; padding-right: 40px; font-size: 15px; color: black;">신청인 : <b style="font-size: 16px;">{user_info['이름']}</b> <span style="position: relative; display: inline-block; width: 30px; text-align: center;">(인)<div style="position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); width: 50px; height: 50px; border: 3px solid red; border-radius: 50%; color: red; font-weight: 900; font-size: 16px; line-height: 48px; text-align: center; z-index: 10; background: transparent;">확인</div></span></div><br><br><h2 style="text-align: center; margin-top: 20px; color: black; font-size: 22px; letter-spacing: 2px;">하이에어공조(주) 귀하</h2></div>"""
        st.components.v1.html(f"<script>function printPDF() {{ var w = window.open('', '_blank', 'width=800,height=900'); w.document.write('<html><head><title>신청서</title></head><body>' + {repr(html_template)} + '</body></html>'); w.document.close(); w.focus(); setTimeout(function() {{ w.print(); w.close(); }}, 250); }}</script><button onclick=\"printPDF()\" style=\"background-color: #FF4B4B; color: white; border: none; padding: 10px 20px; font-size: 15px; font-weight: bold; border-radius: 5px; cursor: pointer; margin-bottom: 20px; width: 100%;\">📥 연차신청서 PDF 다운로드 / 즉시 인쇄하기</button>", height=60)
        st.markdown(html_template, unsafe_allow_html=True)

# --- ✅ 팀원 결재 관리 ---
elif choice == "✅ 팀원 결재 관리 (검토/승인)":
    st.header(f"📥 팀원 결재 및 검토 관리 ({sel_year}년)")
    tab_pending, tab_history = st.tabs(["⏳ 결재/검토 대기", "📜 처리 완료 내역 (히스토리)"])
    all_merged = df_plans.merge(df_emp[['ID', '이름', '팀', '파트', '직급', 'permission']], left_on='Emp_ID', right_on='ID')
    
    with tab_pending:
        if user_info['permission'] == "파트장": display_df = all_merged[(all_merged['파트'] == user_info['파트']) & (all_merged['permission'] == "팀원") & (all_merged['Status'] == "대기")]
        elif user_info['permission'] == "팀장": display_df = all_merged[(all_merged['팀'] == user_info['팀']) & (all_merged['permission'] != "팀장") & (all_merged['Status'].isin(["대기", "검토완료"]))]
        elif user_info['permission'] == "총괄": display_df = all_merged[(all_merged['permission'] == "팀장") & (all_merged['Status'] == "대기")]
        else: display_df = all_merged[all_merged['Status'].isin(["대기", "검토완료"])]

        if display_df.empty: st.info("현재 결재 처리할 내역이 존재하지 않습니다.")
        else:
            display_df = display_df.copy(); display_df['선택'] = st.checkbox("전체 선택/해제")
            edited = st.data_editor(display_df.rename(columns={'Reason': '신청 사유', 'Status': '현재 상태', 'Apply_Time': '신청일시'})[['선택','ID_x','이름','직급','팀','파트','Date','Type','현재 상태','신청 사유','신청일시']], hide_index=True, use_container_width=True)
            s_ids = edited[edited['선택'] == True]['ID_x'].tolist()
            if s_ids:
                col_b1, col_b2 = st.columns(2)
                if col_b1.button("✅ 선택 항목 일괄 처리 진행", use_container_width=True):
                    for t_id in s_ids:
                        idx = df_plans[df_plans["ID"].astype(str) == str(t_id)].index[0]
                        if user_info['permission'] == "파트장": df_plans.at[idx, "Status"] = "검토완료"; df_plans.at[idx, "Part_Sign"] = str(user_info['이름'])
                        else:
                            df_plans.at[idx, "Status"] = "승인"; df_plans.at[idx, "Manager_Sign"] = str(user_info['이름']); df_plans.at[idx, "Approve_Time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            val = 0.5 if "반차" in str(df_plans.at[idx, "Type"]) else 1.0
                            if "연차계획" in str(df_plans.at[idx, "Type"]): df_emp.loc[df_emp["ID"].astype(str) == str(df_plans.at[idx, "Emp_ID"]), "연차계획"] += val
                            else: df_emp.loc[df_emp["ID"].astype(str) == str(df_plans.at[idx, "Emp_ID"]), ["사용","연차잔액"]] += [val, -val]
                    if save_emp_and_plans(df_emp, df_plans, sel_year): st.success("처리 완료!"); st.rerun()
                if col_b2.button("❌ 선택 항목 일괄 반려 처리", use_container_width=True):
                    for t_id in s_ids:
                        idx = df_plans[df_plans["ID"].astype(str) == str(t_id)].index[0]
                        df_plans.at[idx, "Status"] = "반려"
                        if user_info['permission'] == "파트장": df_plans.at[idx, "Part_Sign"] = f"반려됨({user_info['이름']})"
                        else: df_plans.at[idx, "Manager_Sign"] = f"반려됨({user_info['이름']})"; df_plans.at[idx, "Approve_Time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    if save_plans_only(df_plans, sel_year): st.warning("반려 완료"); st.rerun()

    with tab_history:
        history_df = all_merged[all_merged['파트'] == user_info['파트']] if user_info['permission'] == "파트장" else all_merged[all_merged['팀'] == user_info['팀']] if user_info['permission'] == "팀장" else all_merged
        st.dataframe(history_df.sort_values(by="Date", ascending=False)[['Date', 'Type', '이름', '직급', '팀', '파트', 'Status', 'Part_Sign', 'Manager_Sign', 'Apply_Time', 'Approve_Time', 'Reason']], hide_index=True, use_container_width=True)

# --- 📅 연차 현황 달력 ---
elif choice == "📅 연차 현황 달력":
    st.header(f"🗓️ {sel_year}년 연차 현황 달력")
    all_p = df_plans.merge(df_emp[['ID', '이름', '팀', '파트']], left_on='Emp_ID', right_on='ID')
    cal_p = all_p[all_p['Status'].isin(['승인', '대기', '검토완료'])]
    if user_info['permission'] == "파트장": cal_p = cal_p[cal_p['파트'] == user_info['파트']]
    elif user_info['permission'] in ["팀원", "팀장"]: cal_p = cal_p[cal_p['팀'] == user_info['팀']]
    
    y_col, m_col = st.columns(2); t = datetime.now()
    s_y = y_col.selectbox("연도", [t.year, t.year+1, t.year-1])
    s_m = m_col.selectbox("월", range(1, 13), index=t.month-1)
    
    c_heads = st.columns(7)
    for i, d_name in enumerate(["월","화","수","목","금","토","일"]): c_heads[i].write(f"**{d_name}**")
    for week in calendar.monthcalendar(s_y, s_m):
        c_days = st.columns(7)
        for i, day in enumerate(week):
            if day != 0:
                txt = f"**{day}**"; evs = cal_p[cal_p['Date'] == f"{s_y}-{s_m:02d}-{day:02d}"]
                for _, ev in evs.iterrows():
                    c = "#FFFDE7" if "반차" in ev['Type'] else "#F1F8E9" if "휴가" in ev['Type'] or "교육" in ev['Type'] else "#E3F2FD"
                    txt += f"\n<div style='font-size:0.7em; background:{c if ev['Status']=='승인' else '#ECEFF1'}; padding:2px; border-radius:3px; margin-top:2px; color:black;'>{ev['이름']}({ev['Type']})</div>"
                c_days[i].markdown(txt, unsafe_allow_html=True)
        st.divider()

# --- 📊 부서/전사 모니터링 ---
elif choice == "📊 부서/전사 모니터링":
    display_df = df_emp if user_info['permission'] == "관리자" else df_emp[df_emp['팀'] == user_info['팀']] if user_info['permission'] == "팀장" else df_emp[df_emp['파트'] == user_info['파트']]
    st.dataframe(display_df[['팀', '파트', '직급', 'ID', '이름', 'permission', '입사일', '연차기초', '사용', '연차계획', '연차잔액']], use_container_width=True, hide_index=True)

# --- 🌐 [관리자] 전사 통합 관리 ---
elif choice == "🌐 [관리자] 전사 통합 관리":
    buffer = io.BytesIO(); current_notices = load_notices()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_emp.to_excel(writer, sheet_name=f"Employees_{sel_year}", index=False)
        df_plans.to_excel(writer, sheet_name=f"PLANS_{sel_year}", index=False)
        current_notices.to_excel(writer, sheet_name="NOTICES", index=False)
    buffer.seek(0)
    st.download_button("📥 현재 구글시트 최신 데이터를 엑셀 백업본으로 다운로드", data=buffer, file_name=f"vacation_data_backup_{sel_year}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    st.divider()

    tab_list, tab_stat, tab_notice, tab_mail, tab_emp, tab_rollover = st.tabs(["📋 전사 로그 관리", "📈 월간 사용 통계", "📝 연차촉진 공지사항", "📧 리마인드 메일 발송", "👥 임직원 정보", "🗓️ 차기 연도 DB 생성(Rollover)"])
    
    with tab_list:
        all_logs = df_plans[df_plans['Emp_ID'] != ""].merge(df_emp[['ID', '이름', '팀']], left_on='Emp_ID', right_on='ID'); all_logs['선택'] = False
        ed_logs = st.data_editor(all_logs[['선택','ID_x','이름','팀','Date','Type','Status','Apply_Time','Approve_Time']], hide_index=True)
        del_ids = ed_logs[ed_logs['선택'] == True]['ID_x'].tolist()
        if del_ids and st.button("🗑️ 선택 항목 삭제 및 수치 복구"):
            for di in del_ids:
                row_idx = df_plans[df_plans['ID'].astype(str) == str(di)].index
                if not row_idx.empty:
                    row = df_plans.loc[row_idx[0]]
                    if row['Status'] == '승인' and "휴가" not in str(row['Type']) and "교육" not in str(row['Type']):
                        v = 0.5 if "반차" in row['Type'] else 1.0
                        if "연차계획" in row['Type']: df_emp.loc[df_emp['ID'].astype(str)==str(row['Emp_ID']), '연차계획'] -= v
                        else: df_emp.loc[df_emp['ID'].astype(str)==str(row['Emp_ID']), ['사용','연차잔액']] += [-v, v]
            df_plans = df_plans[~df_plans['ID'].astype(str).isin([str(x) for x in del_ids])]
            if save_emp_and_plans(df_emp, df_plans, sel_year): st.rerun()
            
    with tab_stat:
        for_s_date = st.date_input("기준 월 선택"); t_month = for_s_date.strftime("%Y-%m")
        m_plans = df_plans[(df_plans['Date'].str.startswith(t_month)) & (df_plans['Status'] == '승인')].copy()
        m_plans['val'] = m_plans['Type'].apply(lambda x: 0.0 if "휴가" in str(x) or "교육" in str(x) else 0.5 if "반차" in str(x) else 1.0)
        u_stat = m_plans.groupby('Emp_ID')['val'].sum().reset_index().merge(m_plans.groupby('Emp_ID').apply(lambda g: ", ".join([f"{int(r['Date'].split('-')[2])}일({r['Type']})" for _, r in g.sort_values(by="Date").iterrows()])).reset_index(name='상세'), on='Emp_ID', how='left')
        st.dataframe(df_emp[['팀','ID','이름']].merge(u_stat, left_on='ID', right_on='Emp_ID', how='left').fillna({'val':0, '상세':'-'}), use_container_width=True, hide_index=True)

    with tab_notice:
        st.subheader("📝 연차촉진 공지사항 관리")
        df_notices = load_notices(); t_add, t_edit = st.tabs(["등록", "수정/삭제"])
        with t_add:
            with st.form("공지등록", clear_on_submit=True):
                t_title, t_content = st.text_input("공지사항 제목"), st.text_area("공지 내용", height=200)
                if st.form_submit_button("📢 공지사항 등록하기") and t_title and t_content:
                    new_n_id = int(pd.to_numeric(df_notices["ID"], errors='coerce').max() + 1) if not df_notices.empty else 1
                    if save_notices_only(pd.concat([df_notices, pd.DataFrame([{"ID": new_n_id, "날짜": datetime.now().strftime("%Y-%m-%d"), "제목": t_title, "내용": t_content}])], ignore_index=True)): st.rerun()
        with t_edit:
            valid_notices = df_notices[df_notices["제목"].str.strip() != ""]
            if not valid_notices.empty:
                edit_target = st.selectbox("수정/삭제할 공지 선택", valid_notices["제목"].tolist())
                t_row = valid_notices[valid_notices["제목"] == edit_target].iloc[0]
                with st.form("공지수정"):
                    e_title, e_content = st.text_input("제목", value=t_row["제목"]), st.text_area("내용", value=t_row["내용"], height=150)
                    c_e1, c_e2 = st.columns(2)
                    if c_e1.form_submit_button("💾 수정 저장"):
                        df_notices.loc[df_notices["제목"] == edit_target, ["제목", "내용"]] = [e_title, e_content]
                        if save_notices_only(df_notices): st.rerun()
                    if c_e2.form_submit_button("🗑️ 영구 삭제", type="primary"):
                        if save_notices_only(df_notices[df_notices["제목"] != edit_target]): st.rerun()

    # 🚀 [업데이트] 무거운 자동 알림을 끄고, 관리자가 원클릭으로 보낼 수 있는 수동 제어판 탑재
    with tab_mail:
        st.subheader("📧 다가오는 휴가 일정 리마인드 발송 제어판 (D-7)")
        st.info("💡 아래 명단은 향후 1~7일 이내에 연차/반차/계획 등이 예정되어 있으나 아직 리마인드 메일을 받지 않은 임직원들입니다.")
        
        today_date = datetime.now().date()
        preview_rows = []
        for idx, row in df_plans.iterrows():
            if row['Status'] != '반려' and str(row.get('Reminder_Sent', '')) != 'Y' and row['Type'].strip() != "":
                try:
                    plan_date = datetime.strptime(str(row['Date']).strip(), "%Y-%m-%d").date()
                    if 1 <= (plan_date - today_date).days <= 7:
                        emp_info = df_emp[df_emp['ID'].astype(str) == str(row['Emp_ID'])]
                        if not emp_info.empty:
                            preview_rows.append({"사번": row['Emp_ID'], "이름": emp_info.iloc[0]['이름'], "부서": emp_info.iloc[0]['팀'], "예정일자": row['Date'], "구분": row['Type'], "이메일": emp_info.iloc[0].get('EMAIL', '')})
                except: pass
                
        if preview_rows:
            st.dataframe(pd.DataFrame(preview_rows), use_container_width=True, hide_index=True)
            if st.button("🚀 위 대상자 전원에게 리마인드 메일 즉시 일괄 발송", type="primary", use_container_width=True):
                with st.spinner("사내 메일 인프라 통신 중..."):
                    sent_cnt = execute_manual_reminders(df_emp, df_plans, sel_year)
                    st.success(f"🎉 총 {sent_cnt}명의 대상 직원에게 리마인드 안내 메일을 성공적으로 발송했습니다!")
                    st.rerun()
        else:
            st.success("✅ 향후 7일 이내에 리마인드 안내 메일을 발송할 대상자가 없습니다. (모두 발송 완료)")

    with tab_emp:
        st.subheader("👥 임직원 정보 관리")
        if st.session_state.get('emp_save_success'): st.success("✅ 업데이트 완료!"); st.session_state['emp_save_success'] = False
        edited_emp = st.data_editor(df_emp, num_rows="dynamic", use_container_width=True, height=400)
        if st.button("💾 임직원 정보 변경 사항 저장"):
            if save_emp_only(edited_emp, sel_year): st.session_state['emp_save_success'] = True; st.rerun()
                
    with tab_rollover:
        st.subheader("🗓️ 차기 연도 DB 자동 생성 및 연차 리셋")
        create_year = st.number_input("생성할 기준 연도", min_value=2026, value=sel_year + 1, step=1)
        if st.button(f"🚀 {create_year}년도 데이터베이스 완벽 자동 생성 (Rollover)", type="primary"):
            client = get_gspread_client(); sheet = client.open(SPREADSHEET_NAME)
            try: ws_base = sheet.worksheet(f"Employees_{create_year - 1}")
            except: ws_base = sheet.worksheet("Employees") if (create_year - 1) == 2026 else None
                
            if not ws_base: st.error("❌ 이전 연도 직원 데이터를 찾을 수 없습니다.")
            else:
                df_new = pd.DataFrame(ws_base.get_all_values()[1:], columns=ws_base.get_all_values()[0])
                df_new['연차기초'] = [calculate_vacation_accrual(r.get('입사일',''), create_year) for _, r in df_new.iterrows()]
                df_new['사용'] = 0.0; df_new['연차계획'] = 0.0; df_new['연차잔액'] = df_new['연차기초']
                
                for t_name, df_data in [(f"Employees_{create_year}", df_new), (f"PLANS_{create_year}", pd.DataFrame(columns=['ID', 'Emp_ID', 'Date', 'Status', 'Type', 'Reason', 'Manager_Sign', 'Part_Sign', 'Apply_Time', 'Approve_Time', 'Reminder_Sent']))]:
                    try: ws_t = sheet.worksheet(t_name)
                    except: ws_t = sheet.add_worksheet(title=t_name, rows="100", cols="20")
                    ws_t.clear(); df_c = df_data.fillna("").astype(str); ws_t.update(values=[df_c.columns.values.tolist()] + df_c.values.tolist(), value_input_option='RAW')
                get_available_years.clear(); st.success("🎉 rollover 생성 완료!")
