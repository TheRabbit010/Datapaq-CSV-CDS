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

# 4. ฟังก์ชันแปลงวินาทีเป็นรูปแบบ h:mm:ss
def format_seconds_to_time(total_seconds):
    if pd.isna(total_seconds) or total_seconds <= 0:
        return "0:00:00"
    
    total_sec = int(round(total_seconds))
    hours = total_sec // 3600
    minutes = (total_sec % 3600) // 60
    seconds = total_sec % 60
    
    return f"{hours}:{minutes:02d}:{seconds:02d}"

# ฟังก์ชันแปลง Hex Color เป็น RGBA
def hex_to_rgba(hex_str, opacity=0.25):
    hex_str = hex_str.lstrip('#')
    r = int(hex_str[0:2], 16)
    g = int(hex_str[2:4], 16)
    b = int(hex_str[4:6], 16)
    return f"rgba({r}, {g}, {b}, {opacity})"

# 5. ฟังก์ชันอ่านไฟล์ CSV และดึงข้อมูล
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
    probe_channel_map = {}
    data_rows = []
    
    metadata = {
        "paqfile start date": "-",
        "paqfile start time": "-",
        "title": "-",
        "operator": "-",
        "product": "CONDENSER",
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
                    metadata["product"] = val if (val and val != "-") else "CONDENSER"
                elif key.lower() == "site":
                    metadata["site"] = val if (val and val != "-") else "VSTS / Power Chonburi"
                elif "note" in key.lower():
                    metadata["note_1"] = val
                elif key.lower().startswith("probe number #"):
                    try:
                        p_num = int(key.lower().replace("probe number #", "").strip())
                        ch_num = int(val)
                        probe_channel_map[p_num] = ch_num
                    except ValueError:
                        pass
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
                        
                    raw_vals = parts[2:]
                    data_rows.append({
                        "elapsed_sec": elapsed_sec,
                        "time_str": time_str,
                        "dist_val": dist_val,
                        "raw_vals": raw_vals
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
        
        ch_values = {ch: np.nan for ch in range(1, 9)}
        raw_vals = row["raw_vals"]
        
        for idx, val_str in enumerate(raw_vals[:8]):
            col_idx = idx + 1
            ch_num = probe_channel_map.get(col_idx, col_idx)
            
            try:
                val_num = float(val_str)
            except ValueError:
                val_num = np.nan
                
            if 1 <= ch_num <= 8:
                ch_values[ch_num] = val_num

        for i in range(1, 9):
            col_label = f"Probe #{i}"
            if i in probe_labels:
                lbl = probe_labels[i]
                col_label = f"Probe #{i}: {lbl[:15]}..." if len(lbl) > 15 else f"Probe #{i}: {lbl}"
            row_dict[col_label] = ch_values[i]
            
        parsed_data.append(row_dict)

    df_res = pd.DataFrame(parsed_data)
    df_res = df_res.drop_duplicates(subset=["ElapsedSeconds"]).sort_values("ElapsedSeconds").reset_index(drop=True)
    return df_res, metadata

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

uploaded_file = st.sidebar.file_uploader(
    "อัปโหลดไฟล์ CSV (.csv)", 
    type=["csv"],
    accept_multiple_files=False
)

# 7. แสดงผล Header Metadata + กราฟพร้อมโซนเวลา
if uploaded_file:
    df, metadata = parse_single_file(uploaded_file)
    
    if df.empty:
        st.error("⚠️ ไม่สามารถอ่านข้อมูลจากไฟล์ที่อัปโหลดได้ กรุณาตรวจสอบว่าเป็นไฟล์ CSV จาก Datapaq หรือไม่")
    else:
        st.sidebar.success(f"โหลดไฟล์ {uploaded_file.name} สำเร็จ ({len(df)} แถว)")

        st.sidebar.markdown("---")
        st.sidebar.header("🎛️ Dynamic Controls")
        
        dryer_max_sec = 271
        db_range_sec = (298, 841)

        color_shading_mode = st.sidebar.radio(
            "เลือกโหมดแสดงสี:",
            ["แสดงสีตามโซน (By Zone)", "แสดงสีตามกลุ่มงาน (By Process Group)"],
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
                {"Start Time": "00:34:16", "End Time": "00:38:00", "Zone Name": "Exit", "Color": "#CAF0F8"}
            ]
            angle_setting = -90
        else:
            zones_data = [
                {"Start Time": "00:00:00", "End Time": "00:04:58", "Zone Name": "Dryer", "Color": "#F39C12"},      
                {"Start Time": "00:04:59", "End Time": "00:15:34", "Zone Name": "Debinder", "Color": "#E74C3C"},   
                {"Start Time": "00:15:35", "End Time": "00:27:37", "Zone Name": "Brazing", "Color": "#FF0033"},    
                {"Start Time": "00:27:38", "End Time": "00:34:15", "Zone Name": "Cool", "Color": "#00B4D8"},       
                {"Start Time": "00:34:16", "End Time": "00:38:00", "Zone Name": "Exit", "Color": "#90E0EF"}        
            ]
            angle_setting = 0

        max_view_sec = 2280
        df_chart = df[df["ElapsedSeconds"] <= max_view_sec].copy()
        if df_chart.empty:
            df_chart = df.copy()

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
                    <div><span class="raw-header-key">#product</span> = <span class="raw-header-val">{metadata.get('product', 'CONDENSER')}</span></div>
                    <div><span class="raw-header-key">#site</span> = <span class="raw-header-val">{metadata.get('site', 'VSTS / Power Chonburi')}</span></div>
                </div>
            """, unsafe_allow_html=True)

        fig = make_subplots(specs=[[{"secondary_y": False}]])
        
        probe_colors = [
            "#FF0000", "#00FF00", "#0000FF", "#8B4513", 
            "#FF00FF", "#DAA520", "#800080", "#00FFFF"
        ]

        probe_cols = [c for c in df_chart.columns if c.startswith("Probe #")]
        for idx, col in enumerate(probe_cols[:8]):
            if df_chart[col].notna().any():
                fig.add_trace(
                    go.Scatter(
                        x=df_chart["Time (HH:MM:SS)"],
                        y=df_chart[col],
                        name=col,
                        mode="lines",
                        line=dict(color=probe_colors[idx % len(probe_colors)], width=2)
                    )
                )

        fig.add_trace(
            go.Scatter(
                x=df_chart["Distance (m)"],
                y=[None] * len(df_chart),
                xaxis="x2",
                showlegend=False,
                hoverinfo="skip"
            )
        )

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

        step_tick = max(1, len(df_chart) // 16)
        tick_indices = list(range(0, len(df_chart), step_tick))
        if (len(df_chart) - 1) not in tick_indices:
            tick_indices.append(len(df_chart) - 1)
            
        max_dist = df_chart["Distance (m)"].max() if not df_chart.empty else 50.0
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
                tickvals=df_chart.loc[tick_indices, "Time (HH:MM:SS)"].tolist(),
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

        st.markdown(f"""
            <div class="raw-header-box" style="margin-top: -10px; margin-bottom: 25px;">
                <div><span class="raw-header-key">#note #1</span> = <span class="raw-header-val">{metadata.get('note_1', '-')}</span></div>
            </div>
        """, unsafe_allow_html=True)

        # ---------------------------------------------------------
        # 📊 ตารางสรุปค่า
        # ---------------------------------------------------------
        st.markdown("### 📊 ตารางสรุปผลการวิเคราะห์ (Data Table for Google Sheets Copy)")

        dryer_subset = df[(df["ElapsedSeconds"] >= 0) & (df["ElapsedSeconds"] <= dryer_max_sec)]
        debinder_subset = df[(df["ElapsedSeconds"] >= db_range_sec[0]) & (df["ElapsedSeconds"] <= db_range_sec[1])]
        brazing_ht_subset = df[(df["ElapsedSeconds"] >= 0)]
        brazing_max_subset = df[(df["ElapsedSeconds"] >= 900) & (df["ElapsedSeconds"] <= 1750)]

        probe_order = [1, 2, 3, 4, 5, 6, 7, 8]
        ordered_cols = []
        for p_num in probe_order:
            for c in probe_cols[:8]:
                if f"Probe #{p_num}" in c or f"Probe #{p_num}:" in c:
                    ordered_cols.append((p_num, c))
                    break

        summary_rows = []
        for p_num, col_name in ordered_cols:
            location = "Bottom" if p_num in [1, 2, 3, 4] else "Top"
            short_pb_name = f"PB#{p_num}"
            
            # สลับเฉพาะ Probe 3 และ Probe 4 ในส่วน Debinder และ Dryer ให้ตรงตามไฟล์อ้างอิง
            target_col_db_d = col_name
            if p_num == 3:
                target_col_db_d = next((c for p, c in ordered_cols if p == 4), col_name)
            elif p_num == 4:
                target_col_db_d = next((c for p, c in ordered_cols if p == 3), col_name)
            
            probe_series = df[col_name]
            db_d_series = df[target_col_db_d]
            
            is_probe_valid = probe_series.notna().any()
            is_db_d_valid = db_d_series.notna().any()
            
            # Max Temp
            br_val = brazing_max_subset[col_name].max() if (is_probe_valid and not brazing_max_subset.empty) else np.nan
            br_max = f"{br_val:.1f}" if pd.notna(br_val) else "***"
            
            db_val = debinder_subset[target_col_db_d].max() if (is_db_d_valid and not debinder_subset.empty) else np.nan
            db_max = f"{db_val:.1f}" if pd.notna(db_val) else "***"
            
            d_val = dryer_subset[target_col_db_d].max() if (is_db_d_valid and not dryer_subset.empty) else np.nan
            d_max = f"{d_val:.1f}" if pd.notna(d_val) else "***"
            
            # Dwell Times
            if is_probe_valid and pd.notna(br_val):
                br_dwell_sec = (brazing_ht_subset[col_name] >= 577.0).sum() if not brazing_ht_subset.empty else 0
                br_dwell_str = format_seconds_to_time(br_dwell_sec)
            else:
                br_dwell_str = "***"
                
            if is_db_d_valid and pd.notna(db_val):
                db_dwell_sec = (debinder_subset[target_col_db_d] >= 300.0).sum() if not debinder_subset.empty else 0
                db_dwell_str = format_seconds_to_time(db_dwell_sec)
            else:
                db_dwell_str = "***"
                
            if is_db_d_valid and pd.notna(d_val):
                d_dwell_150_sec = (dryer_subset[target_col_db_d] >= 150.0).sum() if not dryer_subset.empty else 0
                d_dwell_150_str = format_seconds_to_time(d_dwell_150_sec)

                d_dwell_sec = (dryer_subset[target_col_db_d] >= 200.0).sum() if not dryer_subset.empty else 0
                d_dwell_str = format_seconds_to_time(d_dwell_sec)
            else:
                d_dwell_150_str = "***"
                d_dwell_str = "***"

            summary_rows.append([
                short_pb_name,
                location,
                br_max,
                db_max,
                d_max,
                br_dwell_str,
                db_dwell_str,
                d_dwell_150_str,
                d_dwell_str
            ])

        multi_cols = pd.MultiIndex.from_tuples([
            ("", "Probe"),
            ("", "Location"),
            ("Brazing zone", "Max temp / probe (°C)"),
            ("Debinder", "Max temp / probe (°C)"),
            ("Dryer", "Max temp / probe (°C)"),
            ("Brazing zone", "at 577°C / probe"),
            ("Debinder", "at 300°C / probe"),
            ("Dryer", "at 150°C / probe"),
            ("Dryer", "at 200°C / probe")
        ])

        display_summary_df = pd.DataFrame(summary_rows, columns=multi_cols)

        st.dataframe(display_summary_df, use_container_width=True, hide_index=True)

        st.markdown("""
            <div style="background-color: #161b22; border: 1px solid #30363d; border-radius: 6px; padding: 12px 18px; font-size: 13px; color: #CCCCCC; margin-top: 10px; line-height: 1.6;">
                <b style="color: #F0B90B;">📌 เกณฑ์มาตรฐานอ้างอิง (Process Standards):</b><br>
                • <b>Maximum Temperatures (°C):</b> Brazing Zone: <b>585 - 607 °C</b> | Debinder Zone: <b>300 - 375 °C</b> | Dryer Zone: <b>200 - 350 °C</b><br>
                • <b>Brazing Dwell Time (at 577°C / probe):</b><br>
                &nbsp;&nbsp;&nbsp;&nbsp;• <b>4.00 - 6.30 min</b> (except end cap RD > 2.00 min)<br>
                &nbsp;&nbsp;&nbsp;&nbsp;• <b>4.00 - 7.45 min</b> for middle center & 1st tube after side plate.<br>
                &nbsp;&nbsp;&nbsp;&nbsp;• <b>2.00 - 7.45 min</b> for end cap RD , middle of the connector block & coldest block.<br>
                • <b>Debinder Dwell Time:</b> at 300°C / probe: <b>> 2:30 min (>150s)</b><br>
                • <b>Dryer Dwell Time:</b> at 150°C / probe: <b>&ge; 1:30 min (&ge;90s)</b> | at 200°C / probe: <b>> 1:30 min (>90s)</b>
            </div>
        """, unsafe_allow_html=True)

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
    st.info("👈 กรุณาเลือกอัปโหลดไฟล์ (.csv) ที่เมนูด้านซ้าย")
