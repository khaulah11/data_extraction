def read_station_sheet(excel_file, sheet):
    raw = pd.read_excel(excel_file, sheet_name=sheet, header=None)
    header_row = find_header_row(raw)
    if header_row is None:
        return None, None
    
    info = get_station_info(raw, default_sheet_name=sheet)
    data = raw.iloc[header_row + 1:].copy().iloc[:, :4]
    data.columns = ["Year", "Month", "Day", "Rainfall"]
    
    # 1. Bersihkan kod TR / TRACE dan nilai bukan angka
    data["Rainfall"] = data["Rainfall"].astype(str).str.strip().str.upper()
    data["Rainfall"] = data["Rainfall"].replace({"TR": "0.1", "TRACE": "0.1", "NONE": "0.0", "NULL": "nan", "-": "nan"})
    
    data["Year"] = pd.to_numeric(data["Year"], errors="coerce")
    data["Month"] = pd.to_numeric(data["Month"], errors="coerce")
    data["Day"] = pd.to_numeric(data["Day"], errors="coerce")
    data["Rainfall"] = pd.to_numeric(data["Rainfall"], errors="coerce")
    
    # Tukar nilai negatif (ralat sensor AAWS seperti -33.3) kepada NaN
    data.loc[data["Rainfall"] < 0, "Rainfall"] = np.nan
    
    data = data.dropna(subset=["Year", "Month", "Day"])
    if len(data) == 0:
        return None, info
        
    data["Year"] = data["Year"].astype(int)
    data["Month"] = data["Month"].astype(int)
    data["Day"] = data["Day"].astype(int)
    return data, info


def generate_styled_excel(station_name, station_info, all_data_list):
    data = pd.concat(all_data_list, ignore_index=True)
    data = data.drop_duplicates(subset=["Year", "Month", "Day"], keep="first")
    data = data.sort_values(by=["Year", "Month", "Day"])
    
    output_stream = io.BytesIO()
    
    with pd.ExcelWriter(output_stream, engine="openpyxl") as writer:
        years = sorted(data["Year"].unique())
        for year in years:
            year = int(year)
            data_year = data[data["Year"] == year].copy()
            
            # 1. Bina Jadual Matriks 1-31 Hari x 1-12 Bulan
            table = data_year.pivot(index="Day", columns="Month", values="Rainfall")
            table = table.reindex(columns=range(1, 13)).reindex(range(1, 32))
            table.columns = months
            table.index.name = "DATE"
            
            # 2. Kira Rumusan Rasmi Borang Kosong
            monthly_totals = []
            monthly_rain_days = []
            monthly_highest_falls = []
            monthly_highest_dates = []
            
            for m in range(1, 13):
                m_df = data_year[data_year["Month"] == m]
                
                # TOTAL
                m_sum = m_df["Rainfall"].sum() if not m_df.empty else 0.0
                monthly_totals.append(round(float(m_sum), 1) if pd.notna(m_sum) else 0.0)
                
                # No.Of days (Hujan >= 0.1 mm)
                m_wet = (m_df["Rainfall"] >= 0.1).sum() if not m_df.empty else 0
                monthly_rain_days.append(int(m_wet))
                
                # Highest fall & Date
                if not m_df.empty and m_df["Rainfall"].notna().any():
                    max_v = m_df["Rainfall"].max()
                    if pd.notna(max_v) and max_v > 0:
                        monthly_highest_falls.append(round(float(max_v), 1))
                        top_days = m_df[m_df["Rainfall"] == max_v]["Day"].tolist()
                        monthly_highest_dates.append(",".join([str(d) for d in top_days]))
                    else:
                        monthly_highest_falls.append(0.0)
                        monthly_highest_dates.append("-")
                else:
                    monthly_highest_falls.append(0.0)
                    monthly_highest_dates.append("-")
            
            station, latitude, longitude, elevation = station_info
            clean_station = station.replace(":", "").strip()
            
            # Header Borang Rasmi
            info_df = pd.DataFrame({
                0: [
                    "JABATAN METEOROLOGI MALAYSIA",
                    "",
                    "DAILY RAINFALL RECORD IN MILLIMETRES",
                    "",
                    f"STATION  : {clean_station.upper()}                      YEAR : {year}"
                ]
            })
            
            info_df.to_excel(writer, sheet_name=str(year), index=False, header=False, startrow=0)
            table.to_excel(writer, sheet_name=str(year), startrow=5)
            
            ws = writer.sheets[str(year)]
            
            # 4 Baris Rumusan Borang Kosong
            ws.cell(row=38, column=1, value="TOTAL")
            ws.cell(row=39, column=1, value="No.Of days")
            ws.cell(row=40, column=1, value="Highest fall")
            ws.cell(row=41, column=1, value="Date")
            
            for m_idx in range(12):
                col_idx = m_idx + 2
                ws.cell(row=38, column=col_idx, value=monthly_totals[m_idx])
                ws.cell(row=39, column=col_idx, value=monthly_rain_days[m_idx])
                ws.cell(row=40, column=col_idx, value=monthly_highest_falls[m_idx])
                ws.cell(row=41, column=col_idx, value=monthly_highest_dates[m_idx])
                
            # Nota Kaki Rasmi
            ws.cell(row=43, column=1, value="P.K.0497(Litho)")
            ws.cell(row=43, column=10, value="TR: Amount less than 0.1mm")
                
    output_stream.seek(0)
    
    # 3. Format Font, Border, & Alignment Menggunakan Openpyxl
    wb = load_workbook(output_stream)
    bold_font = Font(name="Arial", size=9, bold=True)
    title_font = Font(name="Arial", size=10, bold=True)
    data_font = Font(name="Arial", size=9)
    italic_font = Font(name="Arial", size=8, italic=True)
    
    center = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")
    right_align = Alignment(horizontal="right", vertical="center")
    thin_border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
    
    for ws in wb.worksheets:
        ws["A1"].font = title_font
        ws["A3"].font = title_font
        ws["A5"].font = bold_font
        
        # Grid Header (DATE, JAN - DEC)
        for cell in ws[6]:
            if cell.column <= 13:
                cell.font = bold_font
                cell.alignment = center
                cell.border = thin_border
                
        # Grid Nilai Harian (Hari 1 hingga 31)
        for row in ws.iter_rows(min_row=7, max_row=37, min_col=1, max_col=13):
            for cell in row:
                cell.alignment = center
                cell.border = thin_border
                cell.font = data_font
                if cell.column == 1:
                    cell.font = bold_font
                elif isinstance(cell.value, (int, float)):
                    cell.number_format = "0.0"
                    
        # 4 Baris Rumusan
        for r_idx in range(38, 42):
            for c_idx in range(1, 14):
                c = ws.cell(row=r_idx, column=c_idx)
                c.font = bold_font
                c.alignment = center
                c.border = thin_border
                if c_idx >= 2 and r_idx in [38, 40] and isinstance(c.value, (int, float)):
                    c.number_format = "0.0"
                    
        # Footer
        ws["A43"].font = italic_font
        ws["J43"].font = italic_font
        ws["J43"].alignment = right_align
        
        # Lebar Lajur
        ws.column_dimensions["A"].width = 14
        for col in range(2, 14):
            ws.column_dimensions[get_column_letter(col)].width = 8
            
    final_output = io.BytesIO()
    wb.save(final_output)
    final_output.seek(0)
    return final_output, data