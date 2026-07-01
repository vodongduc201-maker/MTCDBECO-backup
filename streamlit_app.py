import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
from datetime import datetime, timedelta, timedelta
import pytz
import gspread
import time
from google.oauth2.service_account import Credentials

# --- 1. CẤU HÌNH ---
st.set_page_config(page_title="Báo Cáo MT - v4026", page_icon="🥤", layout="centered")

st.markdown("""
<style>
/* Thu padding 2 ben cho mobile */
.block-container {
    padding-left: 0.75rem !important;
    padding-right: 0.75rem !important;
    padding-top: 1rem !important;
    max-width: 480px !important;
}

/* Stat card */
.stat-row {
    display: flex;
    gap: 6px;
    margin: 6px 0;
}
.stat-card {
    flex: 1;
    background: #1e1e2e;
    border-radius: 10px;
    padding: 10px 4px 8px 4px;
    text-align: center;
    border: 1px solid #2e2e3e;
    min-width: 0;
}
.stat-number {
    font-size: 1.7rem;
    font-weight: 800;
    line-height: 1;
    margin-bottom: 3px;
}
.stat-label {
    font-size: 0.62rem;
    letter-spacing: 0.08em;
    color: #888;
    font-weight: 600;
    text-transform: uppercase;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}
.color-red    { color: #ff4d4d; }
.color-green  { color: #4dff91; }
.color-yellow { color: #ffd700; }
.color-gray   { color: #aaaaaa; }

/* Progress bar */
.prog-wrap { margin: 5px 0 1px 0; }
.prog-bar-bg {
    background: #2e2e3e;
    border-radius: 6px;
    height: 12px;
    overflow: hidden;
}
.prog-bar-fill {
    height: 100%;
    border-radius: 6px;
}
.prog-info {
    display: flex;
    justify-content: space-between;
    font-size: 0.72rem;
    color: #aaa;
    margin-top: 3px;
}

/* Ep 3 cot Facing/Thung/Le luon nam ngang */
[data-testid="stHorizontalBlock"] {
    flex-wrap: nowrap !important;
    gap: 6px !important;
}
[data-testid="stHorizontalBlock"] > [data-testid="stColumn"] {
    min-width: 0 !important;
    flex: 1 1 0 !important;
    padding: 0 !important;
}

/* Input number: thu gon toi da */
.stNumberInput label {
    font-size: 0.72rem !important;
    margin-bottom: 2px !important;
    white-space: nowrap;
}
.stNumberInput [data-testid="stNumberInputContainer"] {
    padding: 0 !important;
}
.stNumberInput input {
    font-size: 0.95rem !important;
    padding: 5px 4px !important;
    text-align: center !important;
    min-width: 0 !important;
}
.stNumberInput button {
    padding: 2px 5px !important;
    min-width: 24px !important;
    font-size: 0.9rem !important;
}

/* Product card */
.product-card {
    background: #1e1e2e;
    border-radius: 10px;
    padding: 12px;
    border: 1px solid #2e2e3e;
    margin-bottom: 10px;
}
.product-name {
    font-size: 0.9rem;
    font-weight: 600;
    margin-bottom: 8px;
}

/* Nut submit full width */
.stFormSubmitButton button,
.stButton button {
    width: 100% !important;
    height: 3rem !important;
    font-size: 1rem !important;
    border-radius: 10px !important;
}

/* Selectbox label */
.stSelectbox label {
    font-size: 0.85rem !important;
}

/* Confirm info box */
.confirm-box {
    background: #1e1e2e;
    border-radius: 10px;
    padding: 12px 14px;
    border: 1px solid #2e2e3e;
    margin-bottom: 10px;
    font-size: 0.88rem;
    line-height: 1.8;
}
.confirm-row {
    display: flex;
    justify-content: space-between;
    border-bottom: 1px solid #2e2e3e;
    padding: 3px 0;
}
.confirm-row:last-child { border-bottom: none; }
.confirm-key   { color: #888; font-size: 0.78rem; }
.confirm-value { font-weight: 600; text-align: right; }
</style>
""", unsafe_allow_html=True)


# --- 2. GHI DỮ LIỆU (WITH RETRY) ---
def safe_append_to_sheets(flat_row):
    """Ghi dữ liệu vào Sheets với retry 3 lần"""
    try:
        creds = Credentials.from_service_account_info(
            st.secrets["connections"]["gsheets"],
            scopes=[
                "https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive",
            ],
        )
        client = gspread.authorize(creds)
        sheet = client.open("Data_Bao_Cao_MT").worksheet("Data_Bao_Cao_MT")
        
        for attempt in range(3):
            try:
                sheet.append_row(flat_row)
                return True
            except Exception:
                if attempt == 2:
                    raise
                time.sleep(2 * (attempt + 1))
    except Exception as e:
        st.error(f"❌ Lỗi khi ghi dữ liệu: {e}")
        return False


# --- 3. ĐỌC DỮ LIỆU ---
@st.cache_data(ttl=120)
def load_data():
    import os
    possible_nv_files = ["data nhan vien.csv", "data_nhan_vien.csv"]
    possible_sp_files = ["danhmuc.csv", "danhmuc.txt"]
    
    nv_file = None
    for f in possible_nv_files:
        if os.path.exists(f):
            nv_file = f
            break
    
    if not nv_file:
        st.error("❌ Không tìm thấy file 'data nhan vien.xlsx'")
        return pd.DataFrame(), pd.DataFrame()
    
    sp_file = None
    for f in possible_sp_files:
        if os.path.exists(f):
            sp_file = f
            break
    
    if not sp_file:
        st.error("❌ Không tìm thấy file 'danhmuc.csv'")
        return pd.DataFrame(), pd.DataFrame()
    
    df_nv = pd.read_csv(nv_file, header=None).iloc[:, :4]
    df_nv.columns = ["NHAN VIEN", "HE THONG", "PHUONG", "SIEU THI"]
    df_sp = pd.read_csv(sp_file, header=None)
    df_sp.columns = ["HE THONG", "SAN PHAM"]
    return (
        df_nv.apply(lambda x: x.astype(str).str.strip()),
        df_sp.apply(lambda x: x.astype(str).str.strip()),
    )


# --- HELPER: Màu theo ngưỡng ---
def color_class(val, warn, ok):
    if val == 0:        return "color-gray"
    elif val < warn:    return "color-red"
    elif val < ok:      return "color-yellow"
    else:               return "color-green"


# --- HELPER: Render stat cards ---
def render_stat_cards(items):
    cards_html = "".join(
        f'<div class="stat-card">'
        f'<div class="stat-number {cls}">{value}</div>'
        f'<div class="stat-label">{label}</div>'
        f'</div>'
        for label, value, cls in items
    )
    st.markdown(f'<div class="stat-row">{cards_html}</div>', unsafe_allow_html=True)


# --- HELPER: Progress bar ---
def render_progress(done, total, label=""):
    pct = done / total if total > 0 else 0
    pct_int = int(pct * 100)
    color = "#4dff91" if pct_int >= 80 else ("#ffd700" if pct_int >= 50 else "#ff4d4d")
    st.markdown(f"""
    <div class="prog-wrap">
      <div class="prog-bar-bg">
        <div class="prog-bar-fill" style="width:{pct_int}%;background:{color};"></div>
      </div>
      <div class="prog-info"><span>{label}</span><span>{done}/{total} ({pct_int}%)</span></div>
    </div>""", unsafe_allow_html=True)


# --- HELPER: Confirm info row ---
def confirm_rows(pairs):
    rows = "".join(
        f'<div class="confirm-row">'
        f'<span class="confirm-key">{k}</span>'
        f'<span class="confirm-value">{v}</span>'
        f'</div>'
        for k, v in pairs
    )
    st.markdown(f'<div class="confirm-box">{rows}</div>', unsafe_allow_html=True)


# --- HAM KIEM TRA CHAN (HYBRID: SESSION + CACHE) ---
def check_block_status(sel_nv, today_str, now, df_history):
    """
    1. Session state uu tien (nhanh, 0 API)
    2. Fallback: df_history cache (300s, khong them API)
    3. Chong refresh: dung cache thay vi ttl=0
    """
    MIN_WAIT = 360
    tz = pytz.timezone("Asia/Ho_Chi_Minh")
    
    if hasattr(st.session_state, 'last_submit_time') and st.session_state.last_submit_time:
        last_time = st.session_state.last_submit_time
        if pd.notna(last_time):
            if hasattr(last_time, 'tzinfo') and last_time.tzinfo is None:
                last_time = tz.localize(last_time)
            diff = (now - last_time).total_seconds()
            if diff < MIN_WAIT:
                return True, int(MIN_WAIT - diff)
    
    my_history = df_history[(df_history["NHAN VIEN"] == sel_nv) & (df_history["NGAY"] == today_str)]
    if not my_history.empty:
        my_history_copy = my_history.copy()
        my_history_copy["DT"] = pd.to_datetime(
            my_history_copy["NGAY"] + " " + my_history_copy["GIO"], 
            format="%d/%m/%Y %H:%M:%S", errors="coerce"
        )
        last_dt = my_history_copy["DT"].max()
        if pd.notna(last_dt):
            if hasattr(last_dt, 'tzinfo') and last_dt.tzinfo is None:
                last_dt = tz.localize(last_dt)
            diff = (now - last_dt).total_seconds()
            if diff < MIN_WAIT:
                st.session_state.last_submit_time = last_dt
                return True, int(MIN_WAIT - diff)
    
    return False, 0


# ================================================================
# --- 4. HAM CHINH ---
# ================================================================
def main():

    tz = pytz.timezone("Asia/Ho_Chi_Minh")
    now = datetime.now(tz)
    today_str = now.strftime("%d/%m/%Y")
    UU_TIEN_LIST = ["CM", "SF", "CF", "MM", "GO!", "EMART", "CTY", "SM", "NHAN VAN"]
    GIO_CHAN = 1030

    st.title("🥤 Báo Cáo MT")

    df_master, df_sp = load_data()
    sel_nv = st.selectbox(
        "👤 Nhân viên",
        ["Chọn nhân viên..."] + sorted(df_master["NHAN VIEN"].unique().tolist()),
    )
    if sel_nv == "Chọn nhân viên...":
        return

    for key, val in {
        "confirm_pending": False,
        "pending_row": None,
        "form_inputs": {},
        "form_note": "",
        "is_submitting": False,
        "submitted_row_hash": None,
        "last_submit_time": None,
        "last_submit_date": None,
    }.items():
        if key not in st.session_state:
            st.session_state[key] = val
    
    if st.session_state.last_submit_date != today_str:
        st.session_state.last_submit_time = None
        st.session_state.last_submit_date = today_str

    conn = st.connection("gsheets", type=GSheetsConnection)
    try:
        df_history = conn.read(worksheet="Data_Bao_Cao_MT", ttl=60)  # 1 phút - fresh hơn để chặn hoạt động
        df_history.columns = df_history.columns.str.strip().str.upper()
        
        if df_history.empty:
            st.warning("⚠️ Sheet đang rỗng. Bắt đầu nhập dữ liệu mới...")
            df_history = pd.DataFrame(columns=["NGAY", "THU", "GIO", "NHAN VIEN", "HE THONG", "PHUONG", "SIEU THI"])
    except Exception as e:
        st.warning(f"⚠️ Không thể đọc sheet: {e}. Bắt đầu với dữ liệu trắng.")
        df_history = pd.DataFrame(columns=["NGAY", "THU", "GIO", "NHAN VIEN", "HE THONG", "PHUONG", "SIEU THI"])
    
    df_history["NGAY_DT"] = pd.to_datetime(df_history["NGAY"], format="%d/%m/%Y", errors="coerce") if not df_history.empty else pd.Series(dtype="datetime64[ns]")

    df_f1 = df_master[df_master["NHAN VIEN"] == sel_nv]
    sel_ht = st.selectbox("🏪 Hệ thống", sorted(df_f1["HE THONG"].unique().tolist()))
    ht_up = sel_ht.upper()
    
    # Tính số lần viếng & hiển thị cho hệ thống phụ
    st_list = sorted(df_f1[df_f1["HE THONG"] == sel_ht]["SIEU THI"].unique().tolist())
    st_display_list = []
    st_count_map = {}
    
    for st in st_list:
        so_lan = df_history[
            (df_history["NHAN VIEN"] == sel_nv) & (df_history["SIEU THI"] == st)
            & (df_history["NGAY_DT"].dt.month == now.month)
            & (df_history["NGAY_DT"].dt.year == now.year)
        ].shape[0]
        
        if ht_up not in UU_TIEN_LIST:  # Hệ thống phụ
            display = f"❌ {st} ({so_lan}/2 lần)" if so_lan >= 2 else f"✅ {st} ({so_lan}/2 lần)"
            st_display_list.append(display)
            st_count_map[display] = (st, so_lan)
        else:  # Hệ thống ưu tiên
            st_display_list.append(st)
            st_count_map[st] = (st, so_lan)
    
    sel_st_display = st.selectbox("🏬 Siêu thị", st_display_list)
    sel_st, so_lan_st = st_count_map[sel_st_display]

    df_visits = df_history[
        (df_history["NHAN VIEN"] == sel_nv)
        & (df_history["SIEU THI"] == sel_st)
    ].copy()
    if not df_visits.empty:
        df_visits["DT"] = pd.to_datetime(
            df_visits["NGAY"] + " " + df_visits["GIO"],
            format="%d/%m/%Y %H:%M:%S", errors="coerce"
        )
        last = df_visits.loc[df_visits["DT"].idxmax()]
        last_str = last["DT"].strftime("%d/%m/%Y %H:%M") if pd.notna(last["DT"]) else last["NGAY"]
        so_lan_thang = df_visits[
            (df_visits["DT"].dt.month == now.month)
            & (df_visits["DT"].dt.year == now.year)
        ].shape[0]
        st.info(f"🕐 Viếng thăm gần nhất: **{last_str}** · Tháng này: **{so_lan_thang} lần**")
    else:
        st.info("🆕 Chưa có lịch sử viếng thăm siêu thị này.")

    st.divider()
    df_today = df_history[(df_history["NHAN VIEN"] == sel_nv) & (df_history["NGAY"] == today_str)]
    tong_hn = len(df_today)
    ut_hn   = len(df_today[df_today["HE THONG"].str.upper().isin(UU_TIEN_LIST)])
    phu_hn  = tong_hn - ut_hn

    render_stat_cards([
        ("TỔNG HÔM NAY", tong_hn, color_class(tong_hn, 2, 4)),
        ("UT",           ut_hn,   color_class(ut_hn, 1, 3)),
        ("PHỤ",          phu_hn,  color_class(phu_hn, 1, 2)),
    ])

    if not df_today.empty:
        with st.expander(f"📋 Xem chi tiết hôm nay ({tong_hn} lượt)"):
            for _, row in df_today.sort_values("GIO", ascending=False).iterrows():
                ht_r = str(row.get("HE THONG", ""))
                badge = "🟢" if ht_r.upper() in UU_TIEN_LIST else "🟡"
                st.write(f"{badge} {row.get('SIEU THI','')}")
    else:
        st.caption("Chưa có lượt nào hôm nay.")

    st.divider()

    block_reasons = []

    if ht_up != "CTY" and (now.hour * 60 + now.minute) >= GIO_CHAN:
        block_reasons.append("🌙 Đã sau 17:10, không thể gửi.")

    if now.day > 21 and ht_up not in UU_TIEN_LIST:
        block_reasons.append("🚫 Sau ngày 21, chỉ hệ thống ưu tiên được gửi.")

    if ht_up not in UU_TIEN_LIST and so_lan_st >= 2:
        block_reasons.append(f"🚫 {sel_st} đã đủ 2 lần trong tháng.")

    # KIEM TRA CHAN 6 PHUT (HYBRID SMART CACHE)
    is_waiting, wait_time = check_block_status(sel_nv, today_str, now, df_history)
    
    # Debug: Hiển thị khi đang chặn
    if is_waiting and False:  # Set True để debug
        st.write(f"DEBUG: {sel_nv} bị chặn {wait_time}s (từ hàm check_block_status)")
    
    if is_waiting:
        block_reasons.append(f"⏳ Chờ thêm {wait_time} giây mới được gửi tiếp.")
        st.error(f"🚫 Bạn vừa mới gửi báo cáo. Vui lòng chờ {wait_time} giây để tiếp tục.")
        st.info("Báo cáo trước khi di chuyển sang điểm khác.")
        st.stop()

    is_blocked = bool(block_reasons)
    for r in block_reasons:
        st.error(r)

    san_pham_list = df_sp[df_sp["HE THONG"].str.upper() == ht_up]["SAN PHAM"].tolist()
    if not san_pham_list and ht_up != "CTY":
        st.warning("⚠️ Không tìm thấy sản phẩm cho hệ thống này.")
        return

    if st.session_state.confirm_pending:
        st.subheader("📋 Xác nhận trước khi gửi")

        confirm_rows([
            ("Nhân viên", sel_nv),
            ("Hệ thống",  sel_ht),
            ("Siêu thị",  sel_st),
            ("Thời gian", f"{today_str} {now.strftime('%H:%M')}"),
        ])

        sp_pairs = []
        for sp, vals in st.session_state.form_inputs.items():
            tong = (vals["t"] * (12 if "1.5L" in sp else 24)) + vals["l"]
            sp_pairs.append((sp, f"F={vals['f']}  T={tong}"))
        if sp_pairs:
            confirm_rows(sp_pairs)

        if st.session_state.form_note:
            st.caption(f"📝 Ghi chú: {st.session_state.form_note}")

        st.write("")
        row_hash = hash(tuple(st.session_state.pending_row))
        
        btn_label = "⏳ Đang gửi..." if st.session_state.is_submitting else "✅ Xác nhận & Gửi"
        if st.button(btn_label, type="primary", use_container_width=True,
                     disabled=st.session_state.is_submitting):
            if not st.session_state.is_submitting:
                st.session_state.is_submitting = True
                st.rerun()

        if st.session_state.is_submitting:
            if st.session_state.last_submit_time is None:
                st.session_state.last_submit_time = now
                st.session_state.last_submit_date = today_str
            
            if st.session_state.submitted_row_hash == row_hash:
                st.warning("⚠️ Dữ liệu này đã gửi rồi, vui lòng bấm 'Quay lại sửa'.")
            else:
                if safe_append_to_sheets(st.session_state.pending_row):
                    st.session_state.submitted_row_hash = row_hash
                    st.success("✅ Gửi thành công!")
                    st.session_state.confirm_pending = False
                    st.session_state.pending_row     = None
                    st.session_state.form_inputs     = {}
                    st.session_state.form_note       = ""
                    st.session_state.is_submitting   = False
                    for key in list(st.session_state.keys()):
                        if key.startswith("f_") or key.startswith("t_") or key.startswith("l_"):
                            del st.session_state[key]
                    st.rerun()
                else:
                    st.session_state.is_submitting = False

        if not st.session_state.is_submitting:
            if st.button("↩️ Quay lại sửa", use_container_width=True):
                st.session_state.confirm_pending = False
                st.rerun()
        return

    inputs = {}
    if ht_up != "CTY":
        for sp in san_pham_list:
            f_val = st.session_state.get(f"f_{sp}", 0)
            t_val = st.session_state.get(f"t_{sp}", 0)
            l_val = st.session_state.get(f"l_{sp}", 0)
            tong_preview = (t_val * (12 if "1.5L" in sp else 24)) + l_val
            has_data = f_val > 0 or tong_preview > 0
            
            st.markdown(f"""
            <div class="product-card">
                <div class="product-name">{'✅' if has_data else '➕'} {sp}</div>
            </div>
            """, unsafe_allow_html=True)
            
            c1, c2, c3 = st.columns([1, 1, 1])
            with c1:
                f = st.number_input("Facing", min_value=0, step=1, key=f"f_{sp}")
            with c2:
                t = st.number_input("Thùng",  min_value=0, step=1, key=f"t_{sp}")
            with c3:
                l = st.number_input("Lẻ",     min_value=0, step=1, key=f"l_{sp}")
            inputs[sp] = {"f": f, "t": t, "l": l}

    note = st.text_area("📝 Ghi chú", height=80, key="form_note_input")

    if not is_blocked:
        if st.button("🚀 GỬI BÁO CÁO", type="primary", use_container_width=True,
                     disabled=st.session_state.confirm_pending):
            p = df_f1[df_f1["SIEU THI"] == sel_st]["PHUONG"].values[0]
            thu_list = ["Thu Hai","Thu Ba","Thu Tu","Thu Nam","Thu Sau","Thu Bay","Chu Nhat"]
            flat_row = [today_str, thu_list[now.weekday()], now.strftime("%H:%M:%S"),
                        sel_nv, sel_ht, p, sel_st]
            if ht_up == "CTY":
                flat_row.extend([""] * 14)
            else:
                SP_CHUAN = ["Sa Xi Lon","Sa Xi Zero Lon","Xi Pet 390",
                            "Xi Pet 1.5L","Soda Kem Lon","Suoi 500mL","Soda Lon"]
                for sp in SP_CHUAN:
                    if sp in inputs:
                        tong = (inputs[sp]["t"] * (12 if "1.5L" in sp else 24)) + inputs[sp]["l"]
                        flat_row.extend([inputs[sp]["f"], tong])
                    else:
                        flat_row.extend([None, None])
            flat_row.extend([note, ""])

            st.session_state.pending_row     = flat_row
            st.session_state.form_inputs     = inputs
            st.session_state.form_note       = note
            st.session_state.confirm_pending = True
            st.rerun()
    else:
        st.button("🔒 Không thể gửi", disabled=True, use_container_width=True)

    st.divider()
    st.subheader(f"📊 Tháng {now.month}/{now.year}")

    all_ut  = df_master[(df_master["NHAN VIEN"] == sel_nv) &  df_master["HE THONG"].isin(UU_TIEN_LIST)]["SIEU THI"].unique()
    all_phu = df_master[(df_master["NHAN VIEN"] == sel_nv) & ~df_master["HE THONG"].isin(UU_TIEN_LIST)]["SIEU THI"].unique()

    base = (df_history["NHAN VIEN"] == sel_nv) & (df_history["NGAY_DT"].dt.month == now.month) & (df_history["NGAY_DT"].dt.year == now.year)
    done_ut_set  = set(df_history[base &  df_history["HE THONG"].isin(UU_TIEN_LIST)]["SIEU THI"].unique())
    done_phu_set = set(df_history[base & ~df_history["HE THONG"].isin(UU_TIEN_LIST)]["SIEU THI"].unique())

    total_ut = len(all_ut);   done_ut  = len(done_ut_set)
    total_phu= len(all_phu);  done_phu = len(done_phu_set)
    total_all= total_ut + total_phu; done_all = done_ut + done_phu

    render_stat_cards([
        ("TỔNG", done_all, color_class(done_all, total_all // 2 or 1, total_all)),
        ("UT",   done_ut,  color_class(done_ut,  total_ut  // 2 or 1, total_ut)),
        ("PHỤ",  done_phu, color_class(done_phu, total_phu // 2 or 1, total_phu)),
    ])

    st.write("")
    render_progress(done_ut,  total_ut,  "🟢 Ưu tiên")
    render_progress(done_phu, total_phu, "🟡 Phụ")
    render_progress(done_all, total_all, "📦 Tổng")

    chua_ut  = sorted(set(all_ut)  - done_ut_set)
    chua_phu = sorted(set(all_phu) - done_phu_set)
    if chua_ut or chua_phu:
        with st.expander(f"📋 Còn {len(chua_ut)+len(chua_phu)} siêu thị chưa đi"):
            if chua_ut:
                st.markdown("**🟢 Ưu tiên:**")
                for s in chua_ut:  st.write(f"• {s}")
            if chua_phu:
                st.markdown("**🟡 Phụ:**")
                for s in chua_phu: st.write(f"• {s}")
    else:
        st.success("🎉 Hoàn thành tất cả siêu thị tháng này!")


if __name__ == "__main__":
    main()
