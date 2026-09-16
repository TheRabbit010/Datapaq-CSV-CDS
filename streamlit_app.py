import io
import openpyxl
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

# 1. ตั้งค่า Page Config
st.set_page_config(
    page_title="Datapaq NB1 CDS",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 2. บังคับ Dark Mode CSS + จัดการตารางให้อยู่ตรงกลางอย่างสมบูรณ์
st.markdown("""
    <style>
        /* ซ่อนแถบขาว Header ด้านบน */
        header[data-testid="stHeader"] {
            background-color: transparent !important;
            display: none !important;
        }
        [data-testid="stToolbar"] {
            display: none !important;
        }
        
        /* ตั้งค่าพื้นหลัง Dark Mode */
        html, body, .stApp, [data-testid="stAppViewContainer"] {
            background-color: #0e1117 !important;
            color: #ffffff !important;
        }
        [data-testid="stSidebar"] {
            background-color: #161b22 !important;
        }
        .stMarkdown, h1, h2, h3, p, span, label {
            color: #ffffff !important;
        }

        /* ปุ่มเคลียร์ข้อมูลใน Sidebar */
        [data-testid="stSidebar"] div.stButton > button {
            background-color: #21262d !important;
            color: #ffffff !important;
            border: 1px solid #F0B90B !important;
            font-weight: bold !important;
            width: 100% !important;
            padding: 8px 16px !important;
        }
        [data-testid="stSidebar"] div.stButton > button:hover {
            background-color: #F0B90B !important;
            color: #000000 !important;
        }

        /* กล่อง File Uploader */
        [data-testid="stFileUploader"] {
            background-color: #161b22 !important;
            border: 1.5px solid #F0B90B !important;
            border-radius: 8px !important;
            padding: 10px !important;
        }
        [data-testid="stFileUploader"] section {
            background-color: #1c2128 !important;
            border: 1px dashed #F0B90B !important;
            border-radius: 6px !important;
        }
        [data-testid="stFileUploader"] section div, 
        [data-testid="stFileUploader"] section span,
        [data-testid="stFileUploader"] section small {
            color: #e6edf3 !important;
        }

        /* การ์ดไฟล์ที่อัปโหลดแล้ว */
        [data-testid="stFileUploaderFileData"],
        [data-testid="stFileUploaderFileData"] > div,
        [data-testid="stFileUploaderFile"] {
            background-color: #21262d !important;
            border: 1px solid #F0B90B !important;
            border-radius: 6px !important;
        }
        [data-testid="stFileUploaderFileData"] *,
        [data-testid="stFileUploaderFile"] * {
            color: #ffffff !important;
            font-weight: bold !important;
        }

        /* สไตล์กล่องแสดง Header Metadata แบบ Raw Header (#key = value) */
        .raw-header-box {
            background-color: #161b22;
            border: 1px solid #30363d;
            border-left: 4px solid #F0B90B;
            border-radius: 6px;
            padding: 12px 18px;
            font-family: 'Courier New', Courier, monospace;
            font-size: 14px;
            color: #e6edf3;
            margin-bottom: 15px;
            line-height: 1.6;
        }
        .raw-header-key {
            color: #58a6ff;
            font-weight: bold;
        }
        .raw-header-val {
            color: #D29922;
            font-weight: bold;
        }

        /* ปรับแถบ Expander */
        [data-testid="stExpander"] {
            background-color: #161b22 !important;
            border: 1px solid #30363d !important;
            border-radius: 8px !important;
        }
        [data-testid="stExpander"] details summary {
            background-color: #21262d !important;
            color: #ffffff !important;
            border-radius: 8px !important;
        }
        [data-testid="stExpander"] details summary * {
            color: #ffffff !important;
        }

        /* ปรับแต่งตาราง Dataframe & บังคับจัดข้อความ/ตัวเลขทุกช่องให้อยู่ตรงกลางทั้งหมด */
        [data-testid="stDataFrame"], [data-testid="stTable"] {
            background-color: #161b22 !important;
            border: 1px solid #30363d !important;
            border-radius: 8px !important;
        }
        div[data-testid="stDataFrame"] div[role="grid"] {
            background-color: #161b22 !important;
            color: #ffffff !important;
        }
        div[data-testid="stDataFrame"] div[role="columnheader"] {
            background-color: #21262d !important;
            color: #ffffff !important;
            justify-content: center !important;
            text-align: center !important;
        }
        div[data-testid="stDataFrame"] div[role="columnheader"] * {
            justify-content: center !important;
            text-align: center !important;
        }
        div[data-testid="stDataFrame"] div[role="gridcell"] {
            justify-content: center !important;
            text-align: center !important;
            display: flex !important;
            align-items: center !important;
        }
        div[data-testid="stDataFrame"] div[role="gridcell"] * {
            text-align: center !important;
        }
        
        [data-testid="stTable"] th, [data-testid="stTable"] td {
            text-align: center !important;
            vertical-align: middle !important;
        }

        /* ปรับแต่งปุ่มดาวน์โหลด Excel */
        div.stDownloadButton > button {
            background-color: #21262d !important;
            border: 1.5px solid #F0B90B !important;
            border-radius: 6px !important;
            padding: 8px 16px !important;
            transition: all 0.2s ease-in-out;
        }
        div.stDownloadButton > button, 
        div.stDownloadButton > button *,
        div.stDownloadButton > button p,
        div.stDownloadButton > button span {
            color: #ffffff !important;
            font-weight: bold !important;
            font-size: 15px !important;
        }
        div.stDownloadButton > button:hover {
            background-color: #F0B90B !important;
            border-color: #F0B90B !important;
        }
        div.stDownloadButton > button:hover,
        div.stDownloadButton > button:hover *,
        div.stDownloadButton > button:hover p,
        div.stDownloadButton > button:hover span {
            color: #000000 !important;
        }
    </style>
""", unsafe_allow_html=True)

# 3. แสดงชื่อโปรแกรมหลัก
st.title("🏭 Datapaq NB1 CDS")

# 4. ฟังก์ชันแปลงวินาทีเป็นรูปแบบ mm:ss หรือ hh:mm:ss
def format_seconds_to_time(total_seconds):
    if pd.isna(total_seconds) or total_seconds == 0:
        return "00:00"
    
    total_sec = int(round(total_seconds))
    hours = total_sec // 3600
    minutes = (total_sec % 3600) // 60
    seconds = total_sec % 60
    
    if hours == 0:
        return f"{minutes:02d}:{seconds:02d}"
    else:
        return f"{hours:02d}:{minutes:02d}:{seconds:02d}"

# ฟังก์ชันแปลง Hex Color เป็น RGBA
def hex_to_rgba(hex_str, opacity=0.25):
    hex_str = hex_str.lstrip('#')
    r = int(hex_str[0:2], 16)
    g = int(hex_str[2:4], 16)
    b = int(hex_str[4:6], 16)
    return f"rgba({r}, {g}, {b}, {opacity})"

# 5. ฟังก์ชันอ่านไฟล์ CSV และดึงข้อมูล (รองรับค่า *OC* สายหลุด)
def parse_single_file(uploaded_file):
    uploaded_file.seek(0)
    raw_bytes = uploaded_file.read()
    
    text_content = None
    for enc in ['utf-8', 'cp932', 'shift_jis', 'tis-620', 'latin1']:
        try:
            text_content = raw_bytes.decode(enc)
            break
        except Exception:
            continue
            
    if text_content is None:
        text_content = raw_bytes.decode('utf-8', errors='ignore')

    lines = text_content.splitlines()
    
    probe_labels = {}
    data_rows = []
    
    metadata = {
        "paqfile start date": "-",
        "paqfile start time": "-",
        "title": "-",
        "operator": "-",
        "product": "CONDENSER 12SHP+ & KN6,9",
        "site": "VSTS / Power Chonburi",
        "note_1": "-",
        "raw_text": text_content
    }

    for line in lines:
        line_str = line.strip()
        if not line_str:
            continue
        
        if line_str.startswith("#"):
            line_clean = line_str.lstrip("#").strip()
            if "=" in line_clean:
                key, val = [p.strip() for p in line_clean.split("=", 1)]
                val = val.rstrip(",")
                
                if key.lower() == "title":
                    metadata["title"] = val
                elif key.lower() == "paqfile start date":
                    metadata["paqfile start date"] = val
                elif key.lower() == "paqfile start time":
                    metadata["paqfile start time"] = val
                elif key.lower() == "operator":
                    metadata["operator"] = val
                elif key.lower() == "product":
                    metadata["product"] = val if (val and val != "-") else "CONDENSER 12SHP+ & KN6,9"
                elif key.lower() == "site":
                    metadata["site"] = val if (val and val != "-") else "VSTS / Power Chonburi"
                elif "note" in key.lower():
                    metadata["note_1"] = val
                elif key.isdigit():
                    ch_num = int(key)
                    probe_labels[ch_num] = val
        else:
            parts = [p.strip() for p in line_str.split(",") if p.strip() != ""]
            if len(parts) >= 3:
                try:
                    time_str = parts[0].strip()
                    
                    if time_str.startswith("-"):
                        continue
                        
                    t_parts = time_str.split(":")
                    if len(t_parts) == 3:
                        elapsed_sec = int(t_parts[0]) * 3600 + int(t_parts[1]) * 60 + int(t_parts[2])
                    else:
                        elapsed_sec = len(data_rows)
                    
                    try:
                        dist_val = float(parts[1])
                    except ValueError:
                        dist_val = 0.0
                        
                    # แปลงค่าโพรบที่ไม่ใช่ตัวเลข (เช่น *OC*) เป็น np.nan
                    probe_vals = []
                    for p in parts[2:10]:
                        try:
                            probe_vals.append(float(p))
                        except ValueError:
                            probe_vals.append(np.nan)
                    
                    # เติม np.nan ให้ครบ 8 Probe หากแชนแนลเปิดไม่ครบ
                    while len(probe_vals) < 8:
                        probe_vals.append(np.nan)
                    
                    data_rows.append({
                        "elapsed_sec": elapsed_sec,
                        "time_str": time_str,
                        "dist_val": dist_val,
                        "probes": probe_vals
                    })
                except Exception:
                    continue

    if not data_rows:
        return pd.DataFrame(), metadata

    parsed_data = []
    for row in data_rows:
        row_dict = {
            "ElapsedSeconds": row["elapsed_sec"],
            "Time (HH:MM:SS)": row["time_str"],
            "Distance (m)": round(row["dist_val"], 2)
        }
        
        for i in range(1, 9):
            col_label = f"Probe #{i}"
            if i in probe_labels:
                lbl = probe_labels[i]
                col_label = f"Probe #{i}: {lbl[:15]}..." if len(lbl) > 15 else f"Probe #{i}: {lbl}"
            row_dict[col_label] = row["probes"][i-1]
            
        parsed_data.append(row_dict)

    return pd.DataFrame(parsed_data), metadata

def process_multiple_files(uploaded_files):
    combined_dfs = []
    first_metadata = None
    
    for file in uploaded_files:
        df_single, meta_single = parse_single_file(file)
        if not df_single.empty:
            combined_dfs.append(df_single)
            if first_metadata is None:
                first_metadata = meta_single
            
    if not combined_dfs:
        return pd.DataFrame(), {}

    full_df = pd.concat(combined_dfs, ignore_index=True)
    full_df = full_df.sort_values("ElapsedSeconds").reset_index(drop=True)
    return full_df, first_metadata

# ฟังก์ชันแปลง DataFrame + Summary Table + แนบรูปกราฟลงในไฟล์ Excel (.xlsx)
def to_excel_bytes(dataframe, summary_dataframe=None, fig_plotly=None):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if summary_dataframe is not None and not summary_dataframe.empty:
            summary_dataframe.to_excel(writer, sheet_name='Parameter Summary')
            
        df_export = dataframe.copy()
        df_export.to_excel(writer, index=False, sheet_name='Raw Log Data')

    if fig_plotly is not None:
        try:
            img_bytes = fig_plotly.to_image(format="png", width=1200, height=550)
            img_buf = io.BytesIO(img_bytes)
            
            wb = openpyxl.load_workbook(output)
            ws = wb['Parameter Summary'] if 'Parameter Summary' in wb.sheetnames else wb.active
            
            img = openpyxl.drawing.image.Image(img_buf)
            img.anchor = 'A12'
            ws.add_image(img)
            
            output = io.BytesIO()
            wb.save(output)
        except Exception:
            pass

    output.seek(0)
    return output.getvalue()

# 6. เมนู Sidebar
st.sidebar.header("📁 เมนูอัปโหลดข้อมูล")

if st.sidebar.button("🧹 เคลียร์ข้อมูลไฟล์เก่าทั้งหมด"):
    st.cache_data.clear()
    st.rerun()

uploaded_files = st.sidebar.file_uploader(
    "อัปโหลดไฟล์ CSV (.csv) ได้มากกว่า 1 ไฟล์", 
    type=["csv"],
    accept_multiple_files=True
)

# 7. แสดงผล Header Metadata + กราฟพร้อมโซนเวลา
if uploaded_files:
    df, metadata = process_multiple_files(uploaded_files)
    
    if df.empty:
        st.error("⚠️ ไม่สามารถอ่านข้อมูลจากไฟล์ที่อัปโหลดได้ กรุณาตรวจสอบว่าเป็นไฟล์ CSV จาก Datapaq หรือไม่")
    else:
        st.sidebar.success(f"รวมข้อมูลสำเร็จ {len(uploaded_files)} ไฟล์ ({len(df)} แถว)")

        st.sidebar.markdown("---")
        st.sidebar.header("🎛️ Dynamic Controls")
        
        # 📌 ระบบตรวจจับขอบเขตเวลาอัตโนมัติตาม Recipe Model ของไฟล์
        raw_text_meta = metadata.get("raw_text", "").upper()
        if "16XHP" in raw_text_meta:
            default_dryer_end = 270
            default_db_start = 330
            default_db_end = 840
            detected_model_name = "16XHP"
        else:
            default_dryer_end = 271
            default_db_start = 298
            default_db_end = 841
            detected_model_name = "27XHP / SU2"

        st.sidebar.info(f"🤖 ตรวจพบประเภทสูตรอัตโนมัติ: **{detected_model_name}**")

        # 🎛️ เพิ่ม Slider ปรับแต่งช่วงวินาทีเพื่อ Fine-Tune เพิ่มเติมได้ถ้าต้องการ
        with st.sidebar.expander("🛠️ ปรับขอบเขตวินาทีของโซน (Optional Zone Boundaries)"):
            dryer_max_sec = st.slider("Dryer End Sec (วินาทีที่จบ Dryer):", 200, 350, default_dryer_end)
            db_range_sec = st.slider("Debinder Zone Sec (ช่วงวินาที Debinder):", 250, 900, (default_db_start, default_db_end))

        color_shading_mode = st.sidebar.radio(
            "เลือกโหมดแสดงสี:",
            ["แสดงสีตามโซน (By Zone)", "แสดงสีตามกลุ่มงาน (By Process Group)"],
            index=0
        )

        position_naming = st.sidebar.radio(
            "เลือกรูปแบบชื่อตำแหน่งหัววัด (Location Name):",
            ["Bottom / Top", "Right / Left"],
            index=0
        )

        if color_shading_mode == "แสดงสีตามโซน (By Zone)":
            zones_data = [
                {"Start Time": "00:00:00", "End Time": "00:02:14", "Zone Name": "Dryer Z#1", "Color": "#F7DC6F"},
                {"Start Time": "00:02:15", "End Time": "00:04:28", "Zone Name": "Dryer Z#2", "Color": "#F39C12"},
                {"Start Time": "00:04:29", "End Time": "00:04:58", "Zone Name": "EXT Dryer", "Color": "#E67E22"},
                {"Start Time": "00:04:59", "End Time": "00:05:26", "Zone Name": "ENT DB", "Color": "#D35400"},
                {"Start Time": "00:05:27", "End Time": "00:07:41", "Zone Name": "DB Z#1", "Color": "#E74C3C"},
                {"Start Time": "00:07:42", "End Time": "00:09:32", "Zone Name": "DB Z#2", "Color": "#E63946"},
                {"Start Time": "00:09:33", "End Time": "00:11:23", "Zone Name": "DB Z#3", "Color": "#D90429"},
                {"Start Time": "00:11:24", "End Time": "00:13:38", "Zone Name": "DB Z#4", "Color": "#C1121F"},
                {"Start Time": "00:13:39", "End Time": "00:15:34", "Zone Name": "XFER#1", "Color": "#9B59B6"},
                {"Start Time": "00:15:35", "End Time": "00:17:48", "Zone Name": "Z#1", "Color": "#FF0033"},       
                {"Start Time": "00:17:49", "End Time": "00:19:43", "Zone Name": "Z#2", "Color": "#E6002E"},       
                {"Start Time": "00:19:44", "End Time": "00:21:40", "Zone Name": "Z#3", "Color": "#CC0029"},       
                {"Start Time": "00:21:41", "End Time": "00:23:07", "Zone Name": "Z#4", "Color": "#B30024"},       
                {"Start Time": "00:23:08", "End Time": "00:24:33", "Zone Name": "Z#5", "Color": "#CC0029"},       
                {"Start Time": "00:24:34", "End Time": "00:25:59", "Zone Name": "Z#6", "Color": "#E6002E"},       
                {"Start Time": "00:26:00", "End Time": "00:27:37", "Zone Name": "Z#7", "Color": "#FF0033"},       
                {"Start Time": "00:27:38", "End Time": "00:29:19", "Zone Name": "WatCool#1", "Color": "#00B4D8"},
                {"Start Time": "00:29:20", "End Time": "00:30:41", "Zone Name": "WatCool#2", "Color": "#0096C7"},
                {"Start Time": "00:30:42", "End Time": "00:32:02", "Zone Name": "Exit curtain box", "Color": "#0077B6"},
                {"Start Time": "00:32:03", "End Time": "00:32:28", "Zone Name": "XFER#2", "Color": "#023E8A"},
                {"Start Time": "00:32:29", "End Time": "00:33:21", "Zone Name": "AirCool#1", "Color": "#48CAE4"},
                {"Start Time": "00:33:22", "End Time": "00:34:15", "Zone Name": "AirCool#2", "Color": "#90E0EF"},
                {"Start Time": "00:34:16", "End Time": "00:35:35", "Zone Name": "Exit", "Color": "#CAF0F8"}
            ]
            angle_setting = -90
        else:
            zones_data = [
                {"Start Time": "00:00:00", "End Time": "00:04:58", "Zone Name": "Dryer", "Color": "#F39C12"},      
                {"Start Time": "00:04:59", "End Time": "00:15:34", "Zone Name": "Debinder", "Color": "#E74C3C"},   
                {"Start Time": "00:15:35", "End Time": "00:27:37", "Zone Name": "Brazing", "Color": "#FF0033"},    
                {"Start Time": "00:27:38", "End Time": "00:34:15", "Zone Name": "Cool", "Color": "#00B4D8"},       
                {"Start Time": "00:34:16", "End Time": "00:35:35", "Zone Name": "Exit", "Color": "#90E0EF"}        
            ]
            angle_setting = 0

        # 📋 แสดงผล Header Metadata
        col_h1, col_h2 = st.columns(2)
        with col_h1:
            st.markdown(f"""
                <div class="raw-header-box">
                    <div><span class="raw-header-key">#paqfile start date</span> = <span class="raw-header-val">{metadata.get('paqfile start date', '-')}</span></div>
                    <div><span class="raw-header-key">#paqfile start time</span> = <span class="raw-header-val">{metadata.get('paqfile start time', '-')}</span></div>
                    <div><span class="raw-header-key">#title</span> = <span class="raw-header-val">{metadata.get('title', '-')}</span></div>
                </div>
            """, unsafe_allow_html=True)
        with col_h2:
            st.markdown(f"""
                <div class="raw-header-box">
                    <div><span class="raw-header-key">#operator</span> = <span class="raw-header-val">{metadata.get('operator', '-')}</span></div>
                    <div><span class="raw-header-key">#product</span> = <span class="raw-header-val">{metadata.get('product', 'CONDENSER 12SHP+ & KN6,9')}</span></div>
                    <div><span class="raw-header-key">#site</span> = <span class="raw-header-val">{metadata.get('site', 'VSTS / Power Chonburi')}</span></div>
                </div>
            """, unsafe_allow_html=True)

        # สร้างกราฟ Plotly
        fig = make_subplots(specs=[[{"secondary_y": False}]])
        
        probe_colors = [
            "#FF0000", "#00FF00", "#0000FF", "#8B4513", 
            "#FF00FF", "#DAA520", "#800080", "#00FFFF"
        ]

        probe_cols = [c for c in df.columns if c.startswith("Probe #")]
        for idx, col in enumerate(probe_cols[:8]):
            fig.add_trace(
                go.Scatter(
                    x=df["Time (HH:MM:SS)"],
                    y=df[col],
                    name=col,
                    mode="lines",
                    line=dict(color=probe_colors[idx % len(probe_colors)], width=2)
                )
            )

        fig.add_trace(
            go.Scatter(
                x=df["Distance (m)"],
                y=[None] * len(df),
                xaxis="x2",
                showlegend=False,
                hoverinfo="skip"
            )
        )

        # แสดงแถบสีพื้นหลังตามโหมดที่ผู้ใช้เลือกใน Sidebar
        for idx, z_item in enumerate(zones_data):
            start_t = z_item["Start Time"]
            end_t = z_item["End Time"]
            z_name = z_item["Zone Name"]
            color_hex = z_item["Color"]

            fill_opacity = 0.28 if color_shading_mode == "แสดงสีตามกลุ่มงาน (By Process Group)" else 0.22
            fill_rgba = hex_to_rgba(color_hex, fill_opacity)
            
            fig.add_vrect(
                x0=start_t,
                x1=end_t,
                fillcolor=fill_rgba,
                layer="below",
                line_width=0
            )

            font_sz = 12 if color_shading_mode == "แสดงสีตามกลุ่มงาน (By Process Group)" else 10
            
            fig.add_annotation(
                x=start_t,
                y=610,
                text=f"<b>{z_name}</b>",
                showarrow=False,
                xanchor="left",
                yanchor="bottom",
                font=dict(color="#FFFFFF", size=font_sz, family="Arial Bold"),
                textangle=angle_setting
            )

        # คำนวณช่วง Tick สำหรับแกน Time ให้เหมาะสม
        step_tick = max(1, len(df) // 16)
        tick_indices = list(range(0, len(df), step_tick))
        if (len(df) - 1) not in tick_indices:
            tick_indices.append(len(df) - 1)
            
        # สร้างรายการ Tick สำหรับแกน Distance โดยเฉพาะ
        max_dist = df["Distance (m)"].max() if not df.empty else 50.0
        if max_dist <= 20:
            dist_dtick = 1.0
        elif max_dist <= 50:
            dist_dtick = 2.0
        else:
            dist_dtick = 4.0

        fig.update_layout(
            template="plotly_dark",
            plot_bgcolor="#161b22",
            paper_bgcolor="#0e1117",
            hovermode="x unified",
            showlegend=True,
            legend=dict(
                font=dict(color="#FFFFFF", size=11, family="Arial Bold"),
                bgcolor="rgba(27, 31, 36, 0.95)",
                bordercolor="#F0B90B",
                borderwidth=1.5,
                orientation="v",
                yanchor="top",
                y=0.88,
                xanchor="left",
                x=1.02
            ),
            yaxis=dict(
                title=dict(text="Temperature (°C)", font=dict(color="#FFFFFF", size=12)),
                tickfont=dict(color="#CCCCCC", size=10),
                showgrid=True,
                gridcolor="rgba(255,255,255,0.08)",
                zeroline=False,
                linecolor="#555555",
                domain=[0.22, 1.0],
                range=[0, 650]
            ),
            xaxis=dict(
                title=dict(text="Time (hh:mm:ss)", font=dict(color="#FFFFFF", size=11)),
                tickmode="array",
                tickvals=df.loc[tick_indices, "Time (HH:MM:SS)"].tolist(),
                tickfont=dict(color="#CCCCCC", size=10),
                showgrid=True,
                gridcolor="rgba(255,255,255,0.08)",
                showline=True,
                linewidth=1,
                linecolor="#888888",
                anchor="free",
                position=0.12,
                tickangle=0
            ),
            xaxis2=dict(
                title=dict(text="Distance (m)", font=dict(color="#F0B90B", size=11)),
                overlaying="x",
                anchor="free",
                position=0.00,
                
                tickmode="linear",
                tick0=0,
                dtick=dist_dtick,
                tickformat=".2f",
                
                range=[0, max_dist],
                
                tickfont=dict(color="#F0B90B", size=10),
                showgrid=False,
                showline=True,
                linewidth=1,
                linecolor="#F0B90B",
                
                minor=dict(
                    tickmode="linear",
                    tick0=0,
                    dtick=dist_dtick / 2,
                    ticklen=4,
                    tickcolor="#F0B90B",
                    showgrid=False
                ),
                tickangle=0,
                ticks="outside",
                ticklen=6,
            ),
            height=660,
            margin=dict(l=60, r=240, t=50, b=120)
        )

        st.plotly_chart(fig, use_container_width=True)

        # ---------------------------------------------------------
        # 📝 กล่องแสดงข้อความ #note #1 ด้านล่างรูปภาพกราฟ
        # ---------------------------------------------------------
        st.markdown(f"""
            <div class="raw-header-box" style="margin-top: -10px; margin-bottom: 25px;">
                <div><span class="raw-header-key">#note #1</span> = <span class="raw-header-val">{metadata.get('note_1', '-')}</span></div>
            </div>
        """, unsafe_allow_html=True)

        # ---------------------------------------------------------
        # 📊 ตารางสรุปค่า (อัปเดตใช้ช่วงเวลาตามสูตรที่ตรวจพบ)
        # ---------------------------------------------------------
        st.markdown("### 📊 ตารางสรุปผลการวิเคราะห์ (Data Table for Google Sheets Copy)")

        # 📌 การตัด Subset ตามสไลเดอร์/การตรวจจับสูตรอัตโนมัติ
        dryer_subset = df[(df["ElapsedSeconds"] >= 0) & (df["ElapsedSeconds"] <= dryer_max_sec)]
        debinder_subset = df[(df["ElapsedSeconds"] >= db_range_sec[0]) & (df["ElapsedSeconds"] <= db_range_sec[1])]
        
        # Brazing Zone Dwell Time: สะสมเวลาทั้งไฟล์
        brazing_ht_subset = df[(df["ElapsedSeconds"] >= 0)]
        
        # Brazing Zone Max Temp: ช่วงแช่อุณหภูมิสูงสุด
        brazing_max_subset = df[(df["ElapsedSeconds"] >= 900) & (df["ElapsedSeconds"] <= 1750)]

        # ลำดับ Probe ให้ตรงตามแม่แบบ: 1, 2, 3, 8, 4, 5, 6, 7
        probe_order = [1, 2, 3, 8, 4, 5, 6, 7]
        ordered_cols = []
        for p_num in probe_order:
            for c in probe_cols[:8]:
                if f"Probe #{p_num}" in c or f"Probe #{p_num}:" in c:
                    ordered_cols.append((p_num, c))
                    break

        summary_rows = []
        for p_num, col_name in ordered_cols:
            if position_naming == "Bottom / Top":
                location = "Bottom" if p_num in [1, 2, 3, 8] else "Top"
            else:
                location = "Right" if p_num in [1, 2, 3, 8] else "Left"

            short_pb_name = f"PB#{p_num}"
            
            # Max Temp (รองรับกรณีสายหลุดเป็น NaN)
            br_val = brazing_max_subset[col_name].max() if not brazing_max_subset.empty else np.nan
            br_max = f"{br_val:.1f}" if pd.notna(br_val) else "-"
            
            db_val = debinder_subset[col_name].max() if not debinder_subset.empty else np.nan
            db_max = f"{db_val:.1f}" if pd.notna(db_val) else "-"
            
            d_val = dryer_subset[col_name].max() if not dryer_subset.empty else np.nan
            d_max = f"{d_val:.1f}" if pd.notna(d_val) else "-"
            
            # Dwell Time
            br_dwell_600 = (brazing_ht_subset[col_name] >= 600.0).sum() if not brazing_ht_subset.empty else 0
            br_dwell_583 = (brazing_ht_subset[col_name] >= 583.0).sum() if not brazing_ht_subset.empty else 0
            br_dwell_577 = (brazing_ht_subset[col_name] >= 577.0).sum() if not brazing_ht_subset.empty else 0
            
            db_dwell_200 = (debinder_subset[col_name] >= 200.0).sum() if not debinder_subset.empty else 0
            d_dwell_175 = (dryer_subset[col_name] >= 175.0).sum() if not dryer_subset.empty else 0

            summary_rows.append([
                location,
                short_pb_name,
                br_max,
                db_max,
                d_max,
                format_seconds_to_time(br_dwell_600),
                format_seconds_to_time(br_dwell_583),
                format_seconds_to_time(br_dwell_577),
                format_seconds_to_time(db_dwell_200),
                format_seconds_to_time(d_dwell_175)
            ])

        multi_cols = pd.MultiIndex.from_tuples([
            ("", "Location"),
            ("", "Probe"),
            ("Max Temp (°C)", "Brazing"),
            ("Max Temp (°C)", "Debinder"),
            ("Max Temp (°C)", "Dryer"),
            ("Brazing Zone", "Dwell Time Above 600°C"),
            ("Brazing Zone", "Dwell Time Above 583°C"),
            ("Brazing Zone", "Dwell Time Above 577°C"),
            ("Debinder Zone", "Dwell Time Above 200°C"),
            ("Dryer Zone", "Dwell Time Above 175°C")
        ])

        display_summary_df = pd.DataFrame(summary_rows, columns=multi_cols)

        st.dataframe(display_summary_df, use_container_width=True, hide_index=True)

        # คำอธิบายเกณฑ์มาตรฐาน (Process Standards Legend)
        st.markdown("""
            <div style="background-color: #161b22; border: 1px solid #30363d; border-radius: 6px; padding: 12px 18px; font-size: 13px; color: #CCCCCC; margin-top: 10px;">
                <b style="color: #F0B90B;">📌 เกณฑ์มาตรฐานอ้างอิง (Process Standards):</b><br>
                • <b>Maximum Temperatures (°C):</b> Brazing (Corner Probes: <b>596 - 610 °C</b> | Center Probes #2, #5: <b>583 - 607 °C</b>) | Debinder: <b>200 - 375 °C</b> | Dryer: <b>175 - 260 °C</b><br>
                • <b>Brazing Dwell Time (คิดช่วงเวลา 00:00:00 to 00:35:35):</b> Dwell Time Above 600°C: <b>< 4:00 min (<240s)</b> | Dwell Time Above 583°C & 577°C: <b>2:30 - 6:00 min (150s - 360s)</b><br>
                • <b>Debinder Dwell Time:</b> Dwell Time Above 200°C: <b>> 2:00 min (>120s)</b><br>
                • <b>Dryer Dwell Time:</b> Dwell Time Above 175°C: <b>> 1:00 min (>60s)</b>
            </div>
        """, unsafe_allow_html=True)

        # ส่วนตรวจสอบและเลือกดาวน์โหลด Excel (.xlsx)
        with st.expander("📋 ตรวจสอบและเลือกดาวน์โหลดตารางข้อมูล Excel (.xlsx)"):
            st.dataframe(df)
            
            st.markdown("---")
            st.markdown("##### 📥 ตัวเลือกการดาวน์โหลดไฟล์ Excel")
            
            col_opt1, col_opt2 = st.columns([2, 1])
            with col_opt1:
                custom_filename = st.text_input(
                    "ตั้งชื่อไฟล์ดาวน์โหลด:", 
                    value="datapaq_nb1_cds_8probes_data.xlsx"
                )
                if not custom_filename.endswith('.xlsx'):
                    custom_filename += '.xlsx'
                    
            with col_opt2:
                st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                excel_bytes = to_excel_bytes(df, summary_dataframe=display_summary_df, fig_plotly=fig)
                st.download_button(
                    label="📊 ดาวน์โหลดไฟล์ Excel (พร้อมตารางและแนบรูปกราฟ)",
                    data=excel_bytes,
                    file_name=custom_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

else:
    st.info("👈 กรุณาเลือกอัปโหลดไฟล์ (.csv) ที่เมนูด้านซ้าย สามารถเลือกอัปโหลดได้มากกว่า 1 ไฟล์")
