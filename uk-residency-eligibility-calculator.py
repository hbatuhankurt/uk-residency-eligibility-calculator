from datetime import datetime
import pandas as pd

# Data
data = {
    "Departure Date": ["00.00.0000", "00.00.0000"], # You must fill in the dates

    "Return Date": ["00.00.0000", "00.00.0000"] # You must fill in the dates
}

df = pd.DataFrame(data)
for col in ["Departure Date", "Return Date"]:
    df[col] = pd.to_datetime(df[col], dayfirst=True)

# Calculations
df["Days Outside UK"] = (df["Return Date"] - df["Departure Date"]).dt.days - 1
total_days_outside_uk = df["Days Outside UK"].sum()

# ILR 180-Day Rule Check
def check_180_day_rule(df):
    max_days = 0
    for start_date in df["Departure Date"]:
        end_date = start_date + pd.DateOffset(days=364)
        trips_in_window = df[(df["Departure Date"] <= end_date) & (df["Return Date"] >= start_date)]
        total = 0
        for _, row in trips_in_window.iterrows():
            overlap_start = max(row["Departure Date"], start_date)
            overlap_end = min(row["Return Date"], end_date)
            total += (overlap_end - overlap_start).days - 1
        if total > max_days:
            max_days = total
    return max_days

max_180_days = check_180_day_rule(df)
ilr_eligible = "Yes" if max_180_days <= 180 else "No"

# Citizenship 450 Days + Last 12 Months 90 Days Check
last_return = df["Return Date"].max()
last_12_months_start = last_return - pd.DateOffset(years=1)
last_12_months_trips = df[(df["Return Date"] > last_12_months_start)]
last_12_months_days = sum((last_12_months_trips["Return Date"] - last_12_months_trips["Departure Date"]).dt.days - 1)
citizenship_eligible = "Yes" if (total_days_outside_uk <= 450) and (last_12_months_days <= 90) else "No"

# Write to Excel
with pd.ExcelWriter("UK_Residency_Eligibility_Calculation.xlsx", engine="openpyxl") as writer:
    # Travel Details
    df.to_excel(writer, sheet_name="Travels", index=False)

    # Summary Sheet
    summary = pd.DataFrame({
        "Criteria": [
            "Total Days Outside UK (5 Years) - Toplam UK Dışı Gün (5 Yılda)",
            "Max Days Outside UK (12-Month Window) - En Fazla UK Dışı Gün (12 Aylık Pencere)",
            "ILR Eligibility (180-Day Rule) - ILR Uygunluğu (180 Gün Kuralı)",
            "Citizenship Total Days (450 Limit) - Vatandaşlık Toplam Gün (450 Limit)",
            "Days Outside UK in Last 12 Months (90 Limit) - Son 12 Ayda UK Dışı Gün (90 Limit)",
            "Citizenship Eligibility - Vatandaşlık Uygunluğu"
        ],
        "Value": [
            total_days_outside_uk, max_180_days, ilr_eligible, f"{total_days_outside_uk}/450",
            f"{last_12_months_days}/90", citizenship_eligible
        ]
    })
    summary.to_excel(writer, sheet_name="Summary", index=False)

    # Formatting
    workbook = writer.book
    for sheet in ["Travels", "Summary"]:
        ws = workbook[sheet]
        for row in ws.iter_rows(min_row=2, max_col=ws.max_column, max_row=ws.max_row):
            for cell in row:
                cell.number_format = "DD/MM/YYYY" if isinstance(cell.value, datetime) else "General"

    # Conditional Formatting (Summary)
    from openpyxl.styles import PatternFill

    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    ws = workbook["Summary"]
    for row in [3, 6]:  # ILR and Citizenship eligibility rows
        cell = ws[f"B{row}"]
        if cell.value == "Yes":
            cell.fill = green_fill
        elif cell.value == "No":
            cell.fill = red_fill

print("Excel file created: UK_Residency_Eligibility_Calculation.xlsx")
