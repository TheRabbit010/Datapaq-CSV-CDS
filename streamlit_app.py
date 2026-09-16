# แทนที่ส่วนสรุปผล summary_rows ในโค้ดเดิมด้วย logic นี้:

summary_rows = []
for p_num, col_name in ordered_cols:
    location = "Bottom" if p_num in [1, 2, 3, 4] else "Top"
    short_pb_name = f"PB#{p_num}"
    
    # ดึง column ที่ถูกต้องสำหรับ Debinder และ Dryer (สลับเฉพาะ 3 กับ 4 ให้ตรงไฟล์อ้างอิง)
    target_col_db_d = col_name
    if p_num == 3:
        target_col_db_d = next((c for p, c in ordered_cols if p == 4), col_name)
    elif p_num == 4:
        target_col_db_d = next((c for p, c in ordered_cols if p == 3), col_name)
    
    probe_series = df[col_name]
    db_d_series = df[target_col_db_d]
    
    is_probe_valid = probe_series.notna().any()
    is_db_d_valid = db_d_series.notna().any()
    
    # 1. Max Temp
    br_val = brazing_max_subset[col_name].max() if (is_probe_valid and not brazing_max_subset.empty) else np.nan
    br_max = f"{br_val:.1f}" if pd.notna(br_val) else "***"
    
    db_val = debinder_subset[target_col_db_d].max() if (is_db_d_valid and not debinder_subset.empty) else np.nan
    db_max = f"{db_val:.1f}" if pd.notna(db_val) else "***"
    
    d_val = dryer_subset[target_col_db_d].max() if (is_db_d_valid and not dryer_subset.empty) else np.nan
    d_max = f"{d_val:.1f}" if pd.notna(d_val) else "***"
    
    # 2. Dwell Times
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
        d_dwell_sec = (dryer_subset[target_col_db_d] >= 200.0).sum() if not dryer_subset.empty else 0
        d_dwell_str = format_seconds_to_time(d_dwell_sec)
    else:
        d_dwell_str = "***"

    summary_rows.append([
        short_pb_name,
        location,
        br_max,
        db_max,
        d_max,
        br_dwell_str,
        db_dwell_str,
        d_dwell_str
    ])
