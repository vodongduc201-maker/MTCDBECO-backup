import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
from datetime import datetime
import pytz, gspread, time
from google.oauth2.service_account import Credentials

APP_VERSION = "v4027"
SHEET_NAME = "Data_Bao_Cao_MT"
WORKSHEET_NAME = "Data_Bao_Cao_MT"
UU_TIEN_LIST = ["CM", "SF", "CF", "MM", "GO!", "EMART", "CTY", "GENSHAI", "SM", "NHAN VAN"]
GIO_CHAN_H, GIO_CHAN_M = 17, 10
NGAY_CHAN = 21
MAX_LAN_PHU = 2
MIN_WAIT_SECONDS = 360
THUNG_LON = 24
THUNG_NHO = 12
SP_CHUA_1_5L = "1.5L"
SP_CHUAN = ["Sa Xi Lon", "Sa Xi Zero Lon", "Xi Pet 390", "Xi Pet 1.5L", "Soda Kem Lon", "Suoi 500mL", "Soda Lon"]
NV_FILE_CANDIDATES = ["data nhan vien.csv", "data_nhan_vien.csv"]
SP_FILE_CANDIDATES = ["danhmuc.csv", "danhmuc.txt"]
TZ = pytz.timezone("Asia/Ho_Chi_Minh")
THU_LIST = ["Thu Hai", "Thu Ba", "Thu Tu", "Thu Nam", "Thu Sau", "Thu Bay", "Chu Nhat"]
DEFAULT_SESSION = {"confirm_pending": False, "pending_row": None, "form_inputs": {}, "form_note": "", "is_submitting": False, "submitted_row_hash": None, "last_submit_date": None, "df_history": None}
CSS = """<style>
.block-container { padding-left: 0.75rem !important; padding-right: 0.75rem !important; padding-top: 1rem !important; max-width: 480px !important; }
.stat-row { display: flex; gap: 6px; margin: 6px 0; }
.stat-card { flex: 1; background: #1e1e2e; border-radius: 10px; padding: 10px 4px 8px 4px; text-align: center; border: 1px solid #2e2e3e; min-width: 0; }
.stat-number { font-size: 1.7rem; font-weight: 800; line-height: 1; margin-bottom: 3px; }
.stat-label { font-size: 0.62rem; letter-spacing: 0.08em; color: #888; font-weight: 600; text-transform: uppercase; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.color-red { color: #ff4d4d; } .color-green { color: #4dff91; } .color-yellow { color: #ffd700; } .color-gray { color: #aaaaaa; }
.prog-wrap { margin: 5px 0 1px 0; } .prog-bar-bg { background: #2e2e3e; border-radius: 6px; height: 12px; overflow: hidden; }
.prog-bar-fill { height: 100%; border-radius: 6px; } .prog-info { display: flex; justify-content: space-between; font-size: 0.72rem; color: #aaa; margin-top: 3px; }
[data-testid="stHorizontalBlock"] { flex-wrap: nowrap !important; gap: 6px !important; }
[data-testid="stHorizontalBlock"] > [data-testid="stColumn"] { min-width: 0 !important; flex: 1 1 0 !important; padding: 0 !important; }
.stNumberInput label { font-size: 0.72rem !important; margin-bottom: 2px !important; white-space: nowrap; }
.stNumberInput [data-testid="stNumberInputContainer"] { padding: 0 !important; }
.stNumberInput input { font-size: 0.95rem !important; padding: 5px 4px !important; text-align: center !important; min-width: 0 !important; }
.stNumberInput button { padding: 2px 5px !important; min-width: 24px !important; font-size: 0.9rem !important; }
.product-card { background: #1e1e2e; border-radius: 10px; padding: 12px; border: 1px solid #2e2e3e; margin-bottom: 10px; }
.product-name { font-size: 0.9rem; font-weight: 600; margin-bottom: 8px; }
.stFormSubmitButton button, .stButton button { width: 100% !important; height: 3rem !important; font-size: 1rem !important; border-radius: 10px !important; }
.stSelectbox label { font-size: 0.85rem !important; }
.confirm-box { background: #1e1e2e; border-radius: 10px; padding: 12px 14px; border: 1px solid #2e2e3e; margin-bottom: 10px; font-size: 0.88rem; line-height: 1.8; }
.confirm-row { display: flex; justify-content: space-between; border-bottom: 1px solid #2e2e3e; padding: 3px 0; }
.confirm-row:last-child { border-bottom: none; } .confirm-key { color: #888; font-size: 0.78rem; } .confirm-value { font-weight: 600; text-align: right; }
</style>"""

def color_class(val, warn, ok):
    if val == 0: return "color-gray"
    elif val < warn: return "color-red"
    elif val < ok: return "color-yellow"
    else: return "color-green"

def render_stat_cards(items):
    cards = "".join(f'<div class="stat-card"><div class="stat-number {cls}">{value}</div><div class="stat-label">{label}</div></div>' for label, value, cls in items)
    st.markdown(f'<div class="stat-row">{cards}</div>', unsafe_allow_html=True)

def render_progress(done, total, label=""):
    if total <= 0:
        st.markdown(f'<div class="prog-wrap"><div class="prog-info"><span>{label}</span><span>0/0 (0%)</span></div></div>', unsafe_allow_html=True)
        return
    pct_int = int(done / total * 100)
    color = "#4dff91" if pct_int >= 80 else ("#ffd700" if pct_int >= 50 else "#ff4d4d")
    st.markdown(f'<div class="prog-wrap"><div class="prog-bar-bg"><div class="prog-bar-fill" style="width:{pct_int}%;background:{color};"></div></div><div class="prog-info"><span>{label}</span><span>{done}/{total} ({pct_int}%)</span></div></div>', unsafe_allow_html=True)

def confirm_rows(pairs):
    rows = "".join(f'<div class="confirm-row"><span class="confirm-key">{k}</span><span class="confirm-value">{v}</span></div>' for k, v in pairs)
    st.markdown(f'<div class="confirm-box">{rows}</div>', unsafe_allow_html=True)

@st.cache_data(ttl=120)
def load_data():
    import os
    nv_file = next((f for f in NV_FILE_CANDIDATES if os.path.exists(f)), None)
    if not nv_file:
        st.error(f"❌ Không tìm thấy file nhân viên. Đặt 1 trong: {NV_FILE_CANDIDATES}")
        return pd.DataFrame(), pd.DataFrame()
    sp_file = next((f for f in SP_FILE_CANDIDATES if os.path.exists(f)), None)
    if not sp_file:
        st.error(f"❌ Không tìm thấy file danh mục. Đặt 1 trong: {SP_FILE_CANDIDATES}")
        return pd.DataFrame(), pd.DataFrame()
    df_nv = pd.read_csv(nv_file, header=None).iloc[:, :4]
    df_nv.columns = ["NHAN VIEN", "HE THONG", "PHUONG", "SIEU THI"]
    df_sp = pd.read_csv(sp_file, header=None)
    df_sp.columns = ["HE THONG", "SAN PHAM"]
    return df_nv.apply(lambda x: x.astype(str).str.strip()), df_sp.apply(lambda x: x.astype(str).str.strip())

def safe_append_to_sheets(flat_row):
    last_err = None
    for attempt in range(3):
        try:
            creds = Credentials.from_service_account_info(st.secrets["connections"]["gsheets"], scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"])
            client = gspread.authorize(creds)
            sheet = client.open(SHEET_NAME).worksheet(WORKSHEET_NAME)
            sheet.append_row(flat_row)
            return True
        except Exception as e:
            last_err = e
            if attempt < 2:
                try: st.toast(f"⚠️ Thử lại lần {attempt + 2}…")
                except: pass
                time.sleep(2 * (attempt + 1))
    st.error(f"❌ Lỗi khi ghi dữ liệu: {last_err}")
    return False

def get_history_df(conn):
    try:
        df = conn.read(worksheet=WORKSHEET_NAME, ttl=60)
        df.columns = df.columns.str.strip().str.upper()
        if df.empty: return pd.DataFrame(columns=["NGAY", "THU", "GIO", "NHAN VIEN", "HE THONG", "PHUONG", "SIEU THI"])
        return df
    except Exception as e:
        st.warning(f"⚠️ Không thể đọc sheet: {e}. Bắt đầu với dữ liệu trắng.")
        return pd.DataFrame(columns=["NGAY", "THU", "GIO", "NHAN VIEN", "HE THONG", "PHUONG", "SIEU THI"])

def _per_nv_key(nv): return f"last_submit_time__{nv}"

def check_block_status(sel_nv, today_str, now, df_history):
    last_time = st.session_state.get(_per_nv_key(sel_nv))
    if last_time is not None and pd.notna(last_time):
        if last_time.tzinfo is None: last_time = TZ.localize(last_time)
        diff = (now - last_time).total_seconds()
        if diff < MIN_WAIT_SECONDS: return True, int(MIN_WAIT_SECONDS - diff)
    if df_history is not None and not df_history.empty:
        my_history = df_history[(df_history["NHAN VIEN"] == sel_nv) & (df_history["NGAY"] == today_str)].copy()
        if not my_history.empty:
            my_history["DT"] = pd.to_datetime(my_history["NGAY"] + " " + my_history["GIO"], format="%d/%m/%Y %H:%M:%S", errors="coerce")
            last_dt = my_history["DT"].max()
            if pd.notna(last_dt):
                if last_dt.tzinfo is None: last_dt = TZ.localize(last_dt)
                diff = (now - last_dt).total_seconds()
                if diff < MIN_WAIT_SECONDS:
                    st.session_state[_per_nv_key(sel_nv)] = last_dt
                    return True, int(MIN_WAIT_SECONDS - diff)
    return False, 0

def init_session_state():
    for key, val in DEFAULT_SESSION.items():
        if key not in st.session_state: st.session_state[key] = val

def render_top_controls(df_master, df_history, sel_nv, now):
    df_f1 = df_master[df_master["NHAN VIEN"] == sel_nv]
    sel_ht = st.selectbox("🏪 Hệ thống", sorted(df_f1["HE THONG"].unique().tolist()))
    sel_st = st.selectbox("🏬 Siêu thị", sorted(df_f1[df_f1["HE THONG"] == sel_ht]["SIEU THI"].unique().tolist()))
    df_visits = df_history[(df_history["NHAN VIEN"] == sel_nv) & (df_history["SIEU THI"] == sel_st)].copy()
    if not df_visits.empty:
        df_visits["DT"] = pd.to_datetime(df_visits["NGAY"] + " " + df_visits["GIO"], format="%d/%m/%Y %H:%M:%S", errors="coerce")
        last = df_visits.loc[df_visits["DT"].idxmax()]
        last_str = last["DT"].strftime("%d/%m/%Y %H:%M") if pd.notna(last["DT"]) else last["NGAY"]
        so_lan_thang = df_visits[(df_visits["DT"].dt.month == now.month) & (df_visits["DT"].dt.year == now.year)].shape[0]
        st.info(f"🕐 Viếng thăm gần nhất: **{last_str}** · Tháng này: **{so_lan_thang} lần**")
    else:
        st.info("🆕 Chưa có lịch sử viếng thăm siêu thị này.")
    return sel_ht, sel_st

def render_today_dashboard(df_history, sel_nv, today_str):
    df_today = df_history[(df_history["NHAN VIEN"] == sel_nv) & (df_history["NGAY"] == today_str)]
    tong_hn, ut_hn = len(df_today), len(df_today[df_today["HE THONG"].str.upper().isin(UU_TIEN_LIST)])
    render_stat_cards([("TỔNG HÔM NAY", tong_hn, color_class(tong_hn, 2, 4)), ("UT", ut_hn, color_class(ut_hn, 1, 3)), ("PHỤ", tong_hn - ut_hn, color_class(tong_hn - ut_hn, 1, 2))])
    if not df_today.empty:
        with st.expander(f"📋 Xem chi tiết hôm nay ({tong_hn} lượt)"):
            for _, row in df_today.sort_values("GIO", ascending=False).iterrows():
                badge = "🟢" if str(row.get("HE THONG", "")).upper() in UU_TIEN_LIST else "🟡"
                st.write(f"{badge} {row.get('SIEU THI','')}")
    else:
        st.caption("Chưa có lượt nào hôm nay.")

def render_form(san_pham_list):
    inputs = {}
    for sp in san_pham_list:
        f_val = st.session_state.get(f"f_{sp}", 0)
        t_val = st.session_state.get(f"t_{sp}", 0)
        l_val = st.session_state.get(f"l_{sp}", 0)
        thung_size = THUNG_NHO if SP_CHUA_1_5L in sp else THUNG_LON
        tong_preview = (t_val * thung_size) + l_val
        st.markdown(f'<div class="product-card"><div class="product-name">{"✅" if f_val > 0 or tong_preview > 0 else "➕"} {sp}</div></div>', unsafe_allow_html=True)
        c1, c2, c3 = st.columns([1, 1, 1])
        with c1: f = st.number_input("Facing", min_value=0, step=1, key=f"f_{sp}")
        with c2: t = st.number_input("Thùng", min_value=0, step=1, key=f"t_{sp}")
        with c3: l = st.number_input("Lẻ", min_value=0, step=1, key=f"l_{sp}")
        inputs[sp] = {"f": f, "t": t, "l": l}
    note = st.text_area("📝 Ghi chú", height=80, key="form_note_input")
    return inputs, note

def build_flat_row(today_str, now, sel_nv, sel_ht, sel_st, phuong, inputs, note):
    flat = [today_str, THU_LIST[now.weekday()], now.strftime("%H:%M:%S"), sel_nv, sel_ht, phuong, sel_st]
    if sel_ht.upper() == "CTY":
        flat.extend([""] * (len(SP_CHUAN) * 2))
    else:
        for sp in SP_CHUAN:
            if sp in inputs:
                thung_size = THUNG_NHO if SP_CHUA_1_5L in sp else THUNG_LON
                tong = (inputs[sp]["t"] * thung_size) + inputs[sp]["l"]
                flat.extend([inputs[sp]["f"], tong])
            else:
                flat.extend([None, None])
    flat.extend([note, ""])
    return flat

def check_recent_duplicate(df_history, sel_nv, sel_st, today_str, sp_part):
    if df_history is None or df_history.empty: return False, None
    same = df_history[(df_history["NHAN VIEN"] == sel_nv) & (df_history["SIEU THI"] == sel_st) & (df_history["NGAY"] == today_str)]
    if same.empty: return False, None
    sp_cols = list(same.columns[7:7 + len(sp_part)])
    for _, row in same.iterrows():
        existing = [row.get(c) for c in sp_cols]
        if [None if pd.isna(v) else v for v in existing] == sp_part: return True, "Đã gửi dữ liệu giống hệt cho ST này trong ngày."
    return False, None

def render_confirm_screen(sel_nv, sel_ht, sel_st, today_str, now, inputs, note, df_history):
    st.subheader("📋 Xác nhận trước khi gửi")
    confirm_rows([("Nhân viên", sel_nv), ("Hệ thống", sel_ht), ("Siêu thị", sel_st), ("Thời gian", f"{today_str} {now.strftime('%H:%M')}")])
    if inputs:
        sp_pairs = []
        for sp, vals in inputs.items():
            thung_size = THUNG_NHO if SP_CHUA_1_5L in sp else THUNG_LON
            tong = (vals["t"] * thung_size) + vals["l"]
            sp_pairs.append((sp, f"F={vals['f']}  T={tong}"))
        confirm_rows(sp_pairs)
    if note: st.caption(f"📝 Ghi chú: {note}")
    st.write("")
    row_hash = hash(tuple(st.session_state.pending_row))
    sp_part = st.session_state.pending_row[7:-2]
    btn_label = "⏳ Đang gửi..." if st.session_state.is_submitting else "✅ Xác nhận & Gửi"
    if st.button(btn_label, type="primary", use_container_width=True, disabled=st.session_state.is_submitting):
        if not st.session_state.is_submitting: st.session_state.is_submitting = True; st.rerun()
    if st.session_state.is_submitting:
        is_dup, dup_msg = check_recent_duplicate(df_history, sel_nv, sel_st, today_str, sp_part)
        if is_dup and st.session_state.submitted_row_hash != row_hash: st.warning(f"⚠️ {dup_msg}"); st.session_state.is_submitting = False; st.session_state.submitted_row_hash = row_hash; return
        if st.session_state.submitted_row_hash == row_hash:
            st.warning("⚠️ Dữ liệu này đã gửi rồi, vui lòng bấm 'Quay lại sửa'.")
        elif safe_append_to_sheets(st.session_state.pending_row):
            st.session_state[_per_nv_key(sel_nv)] = now; st.session_state.submitted_row_hash = row_hash; st.success("✅ Gửi thành công!")
            st.session_state.confirm_pending = False; st.session_state.pending_row = None; st.session_state.form_inputs = {}; st.session_state.form_note = ""; st.session_state.is_submitting = False
            for key in list(st.session_state.keys()):
                if key.startswith(("f_", "t_", "l_")): del st.session_state[key]
            st.rerun()
        else: st.session_state.is_submitting = False
    if not st.session_state.is_submitting:
        if st.button("↩️ Quay lại sửa", use_container_width=True): st.session_state.confirm_pending = False; st.rerun()

def render_monthly_summary(df_master, df_history, sel_nv, now):
    st.subheader(f"📊 Tháng {now.month}/{now.year}")
    all_ut = df_master[(df_master["NHAN VIEN"] == sel_nv) & df_master["HE THONG"].isin(UU_TIEN_LIST)]["SIEU THI"].unique()
    all_phu = df_master[(df_master["NHAN VIEN"] == sel_nv) & ~df_master["HE THONG"].isin(UU_TIEN_LIST)]["SIEU THI"].unique()
    base = (df_history["NHAN VIEN"] == sel_nv) & (df_history["NGAY_DT"].dt.month == now.month) & (df_history["NGAY_DT"].dt.year == now.year)
    done_ut_set = set(df_history[base & df_history["HE THONG"].isin(UU_TIEN_LIST)]["SIEU THI"].unique())
    done_phu_set = set(df_history[base & ~df_history["HE THONG"].isin(UU_TIEN_LIST)]["SIEU THI"].unique())
    total_ut, done_ut = len(all_ut), len(done_ut_set)
    total_phu, done_phu = len(all_phu), len(done_phu_set)
    total_all, done_all = total_ut + total_phu, done_ut + done_phu
    render_stat_cards([("TỔNG", done_all, color_class(done_all, total_all // 2 or 1, total_all)), ("UT", done_ut, color_class(done_ut, total_ut // 2 or 1, total_ut)), ("PHỤ", done_phu, color_class(done_phu, total_phu // 2 or 1, total_phu))])
    st.write("")
    render_progress(done_ut, total_ut, "🟢 Ưu tiên")
    render_progress(done_phu, total_phu, "🟡 Phụ")
    render_progress(done_all, total_all, "📦 Tổng")
    chua_ut, chua_phu = sorted(set(all_ut) - done_ut_set), sorted(set(all_phu) - done_phu_set)
    if chua_ut or chua_phu:
        with st.expander(f"📋 Còn {len(chua_ut)+len(chua_phu)} siêu thị chưa đi"):
            if chua_ut: st.markdown("**🟢 Ưu tiên:**"); [st.write(f"• {s}") for s in chua_ut]
            if chua_phu: st.markdown("**🟡 Phụ:**"); [st.write(f"• {s}") for s in chua_phu]
    else: st.success("🎉 Hoàn thành tất cả siêu thị tháng này!")

def main():
    st.set_page_config(page_title=f"Báo Cáo MT - {APP_VERSION}", page_icon="🥤", layout="centered")
    st.markdown(CSS, unsafe_allow_html=True)
    init_session_state()
    now = datetime.now(TZ)
    today_str = now.strftime("%d/%m/%Y")
    if st.session_state.last_submit_date != today_str:
        for k in [k for k in st.session_state.keys() if k.startswith("last_submit_time__")]: del st.session_state[k]
        st.session_state.last_submit_date = today_str
    st.title("🥤 Báo Cáo MT")
    df_master, df_sp = load_data()
    sel_nv = st.selectbox("👤 Nhân viên", ["Chọn nhân viên..."] + sorted(df_master["NHAN VIEN"].unique().tolist()))
    if sel_nv == "Chọn nhân viên...": return
    conn = st.connection("gsheets", type=GSheetsConnection)
    df_history = get_history_df(conn)
    if not df_history.empty and "NGAY" in df_history.columns:
        df_history["NGAY_DT"] = pd.to_datetime(df_history["NGAY"], format="%d/%m/%Y", errors="coerce")
    else:
        df_history["NGAY_DT"] = pd.Series(dtype="datetime64[ns]")
    st.session_state.df_history = df_history
    sel_ht, sel_st = render_top_controls(df_master, df_history, sel_nv, now)
    st.divider()
    render_today_dashboard(df_history, sel_nv, today_str)
    st.divider()
    block_reasons = []
    ht_up = sel_ht.upper()
    gio_chan_min = GIO_CHAN_H * 60 + GIO_CHAN_M
    if ht_up != "CTY" and (now.hour * 60 + now.minute) >= gio_chan_min: block_reasons.append(f"🌙 Đã sau {GIO_CHAN_H:02d}:{GIO_CHAN_M:02d}, không thể gửi.")
    if now.day > NGAY_CHAN and ht_up not in UU_TIEN_LIST: block_reasons.append(f"🚫 Sau ngày {NGAY_CHAN}, chỉ hệ thống ưu tiên được gửi.")
    so_lan = df_history[(df_history["NHAN VIEN"] == sel_nv) & (df_history["SIEU THI"] == sel_st) & (df_history["NGAY_DT"].dt.month == now.month) & (df_history["NGAY_DT"].dt.year == now.year)].shape[0]
    if ht_up not in UU_TIEN_LIST and so_lan >= MAX_LAN_PHU: block_reasons.append(f"🚫 {sel_st} đã đủ {MAX_LAN_PHU} lần trong tháng.")
    is_waiting, wait_time = check_block_status(sel_nv, today_str, now, df_history)
    if is_waiting: st.error(f"🚫 Bạn vừa mới gửi báo cáo. Vui lòng chờ {wait_time} giây để tiếp tục."); st.info("Báo cáo trước khi di chuyển sang điểm khác."); st.stop()
    is_blocked = bool(block_reasons)
    for r in block_reasons: st.error(r)
    if st.session_state.confirm_pending: render_confirm_screen(sel_nv, sel_ht, sel_st, today_str, now, st.session_state.form_inputs, st.session_state.form_note, df_history); return
    san_pham_list = df_sp[df_sp["HE THONG"].str.upper() == ht_up]["SAN PHAM"].tolist()
    if not san_pham_list and ht_up != "CTY": st.warning("⚠️ Không tìm thấy sản phẩm cho hệ thống này."); return
    inputs, note = ({}, "") if ht_up == "CTY" else render_form(san_pham_list)
    if not is_blocked:
        if st.button("🚀 GỬI BÁO CÁO", type="primary", use_container_width=True, disabled=st.session_state.confirm_pending):
            phuong_match = df_master[(df_master["NHAN VIEN"] == sel_nv) & (df_master["SIEU THI"] == sel_st)]["PHUONG"].values
            p = phuong_match[0] if len(phuong_match) > 0 else ""
            flat_row = build_flat_row(today_str, now, sel_nv, sel_ht, sel_st, p, inputs, note)
            st.session_state.pending_row = flat_row; st.session_state.form_inputs = inputs; st.session_state.form_note = note; st.session_state.confirm_pending = True; st.rerun()
    else: st.button("🔒 Không thể gửi", disabled=True, use_container_width=True)
    st.divider()
    render_monthly_summary(df_master, df_history, sel_nv, now)

if __name__ == "__main__": main()
