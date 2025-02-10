import streamlit as st
import polars as pl
import xlwings as xw
import atexit
import time
import os

def automize_selection(wb, pu_code, rubber_code, fabric_code, liner_code, uv_code, liner_texture_code, sheet_name = None):
    pu_thickness = pu_df.filter(pl.col("SurrogateKey") == pu_code).select("THICKNESS_mm").item()
    rubber_thickness = rubber_df.filter(pl.col("SurrogateKey") == rubber_code).select("THICKNESS_mm").item()
    fabric_thickness = fabric_df.filter(pl.col("SurrogateKey") == fabric_code).select("THICKNESS_mm").item()
    liner_thickness = liner_df.filter(pl.col("SurrogateKey") == liner_code).select("THICKNESS_mm").item()
    uv_thickness = uv_print_df.filter(pl.col("SurrogateKey") == uv_code).select("THICKNESS_mm").item()

    total_thickness = pu_thickness + rubber_thickness + fabric_thickness + liner_thickness + uv_thickness

    pu_weight = pu_df.filter(pl.col("SurrogateKey") == pu_code).select("WEIGHT_gsm").item()
    rubber_weight = rubber_df.filter(pl.col("SurrogateKey") == rubber_code).select("WEIGHT_gsm").item()
    fabric_weight = fabric_df.filter(pl.col("SurrogateKey") == fabric_code).select("WEIGHT_gsm").item()
    liner_weight = liner_df.filter(pl.col("SurrogateKey") == liner_code).select("WEIGHT_gsm").item()
    uv_weight = uv_print_df.filter(pl.col("SurrogateKey") == uv_code).select("WEIGHT_gsm").item()

    total_weight = pu_weight + rubber_weight + fabric_weight + liner_weight + uv_weight

    pu_weight_percent = round(pu_weight / total_weight * 100)
    rubber_weight_percent = round(rubber_weight / total_weight * 100)
    fabric_weight_percent = round(fabric_weight / total_weight * 100)
    liner_weight_percent = round(liner_weight / total_weight * 100)
    uv_weight_percent = round(uv_weight / total_weight * 100)
    total_weight_percentage = pu_weight_percent + rubber_weight_percent + fabric_weight_percent + liner_weight_percent + uv_weight_percent

    if total_weight_percentage < 100:
        rubber_weight_percent = rubber_weight_percent + (100 - total_weight_percentage)
    elif total_weight_percentage > 100:
        rubber_weight_percent = rubber_weight_percent - (total_weight_percentage - 100)

    uv_pu_weight_percent = uv_weight_percent + pu_weight_percent

    biobased = 0

    if fabric_df.filter(pl.col("SurrogateKey") == fabric_code).select("unnatural").item() == 1 & liner_df.filter(pl.col("SurrogateKey") == liner_code).select("unnatural").item() == 1:
        biobased = round(pu_weight_percent*0.94, 2)
    elif fabric_df.filter(pl.col("SurrogateKey") == fabric_code).select("unnatural").item() == 0 & liner_df.filter(pl.col("SurrogateKey") == liner_code).select("unnatural").item() == 1:
        biobased = round(pu_weight_percent*0.94 + fabric_weight_percent, 2)
    elif fabric_df.filter(pl.col("SurrogateKey") == fabric_code).select("unnatural").item() == 1 & liner_df.filter(pl.col("SurrogateKey") == liner_code).select("unnatural").item() == 0:
        biobased = round(pu_weight_percent*0.94 + liner_weight_percent, 2)
    else :
        biobased = round(pu_weight_percent*0.94 + fabric_weight_percent + liner_weight_percent, 2)

    sheet_names = [sht.name for sht in wb.sheets]
    
    if sheet_name in sheet_names:
        wb.sheets[sheet_name].delete()  # Delete the existing sheet
    
    sht = wb.sheets.add(name=sheet_name)

    # Set the worksheet background color
    sht.range("A1:Z50").color = (215, 228, 202)  # Light green background

    liner_texture = liner_texture_df.filter(pl.col('SurrogateKey') == liner_texture_code).select(pl.col('Texture No.')).item()
    liner_number = liner_df.filter(pl.col('SurrogateKey') == liner_code).select(pl.col('Natuura_key')).item()


    #Layer
    rubber  = rubber_df.filter(pl.col('SurrogateKey') == rubber_code).select(pl.col('Short Name')).item()
    fabric  = fabric_df.filter(pl.col('SurrogateKey') == fabric_code).select(pl.col('Short Name')).item()
    liner  = liner_df.filter(pl.col('SurrogateKey') == liner_code).select(pl.col('Short Name')).item()
    uv = uv_print_df.filter(pl.col('SurrogateKey') == uv_code).select(pl.col('Short Name')).item()

    sht.range("A1:D1").value = ["Company\nTORY BURCH", "BRAND\nLUCID MAT™", "STYLE\nPatent", "DATE\n2025-01-20"]
    sht.range("A2:D2").value = ["FINISH UV\n printing/glossy", f"TEXTURE\n{liner_texture}", f"LINER NO.\n{liner_number}", f"WEIGHT\n{total_weight}gsm ± 20%"]

    sht.range("A1:D2").api.WrapText = True
    sht.range("A1:D2").row_height = None
    sht.range("A1:D2").column_width = None


    # Rows with 2 columns (spanning full width of 4 columns)
    sht.range("A3:C3").merge()
    sht.range("A3").value = [f"COLOR\n"]

    sht.range("D3").value = [f"THICKNESS\n{total_thickness}mm ± 0.2mm"]

    sht.range("A4:C9").merge()
    sht.range("A4").api.WrapText = True

    sht.range("D4:D9").merge()
    sht.range("D4").api.WrapText = True
    sht.range("A10:C10").merge()

    if uv != 'NA':
        sht.range("A4").value = [f"""COMPOSITION
{rubber}: 
{fabric}: 
{liner}:
DRY POLYURETHANE +TOP COATING & UV PRINT:"""]
        sht.range("A10").value = "ESTIMATED BIOBASED:"
        sht.range("D4").value = [f"""Weight %
{rubber_weight_percent}%
{fabric_weight_percent}%
{liner_weight_percent}%

{uv_pu_weight_percent}%"""]
        sht.range("D10").value = f"{biobased}%"
    else:
        sht.range("A4").value = [f"""COMPOSITION
{rubber}:
{fabric}:
{liner}:
    DRY POLYURETHANE:"""]
        sht.range("A10").value = "ESTIMATED BIOBASED:"
        sht.range("D4").value = [f"""Weight %
{rubber_weight_percent}%
{fabric_weight_percent}%
{liner_weight_percent}%

{pu_weight_percent}%"""]
        sht.range("D10").value = f"{biobased}%"

    sht.range("D10").api.HorizontalAlignment = -4131  

    sht.range("A11:D15").merge()
    sht.range("A11").api.WrapText = True
    sht.range("A11").row_height = None
    sht.range("A11").column_width = None
    sht.range("A11").value = """MOQ / MCQ, SIZE, LEAD TIME
MOQ – 300m ; MCQ – 300m
Bulk Lead Time – 2 Months (with In-Stock Liner)
Usable Width – 1.32m (Up to 30m Per Roll)
    """

    sht.range("A16:D23").merge()
    sht.range("A16").api.WrapText = True
    sht.range("A16").row_height = None
    sht.range("A16").column_width = None
    sht.range("A16").value = """OPTIONS
Overall Thickness – 1.0mm – 2.0mm
Pigmentation – Pantone, Custom
Variety of Continuous Textures
Liners – Biobased - 100% GOTS Organic Cotton, 100% Cotton, 100% TENCEL™ (Canvas or Interlock)
        Recycled - 100% GRS Post-Consumer Recycled Polyester
        (Knit or Suede)
    """

    sht.range("A1").column_width = 12.56
    sht.range("B1").column_width = 11.44
    sht.range("C1").column_width = 11.44
    sht.range("D1").column_width = 15.22

    # Forpu cells with borders
    last_row = 23  # Example last row; adjust as needed
    for row in range(1, last_row + 1):
        sht.range(f"A{row}:D{row}").api.Borders.Weight = 2  # Add borders

    # Left, Top, Bottom, Right 9,10,11,12

    sht.range("A10:D10").api.Insert()

    sht.range("D4:D11").api.Borders(7).LineStyle = -4142
    sht.range("A10:D11").api.Borders(8).LineStyle = -4142

    for i in [9, 10, 11, 12]:  # Left, Top, Bottom, Right
        sht.range("A10:D10").api.Borders(i).LineStyle = -4142

    # Force an Excel refresh to apply changes
    wb.app.calculate()

    # Reapply only the right border to D10
    border = sht.range("D10:D10").api.Borders(10)  # Right border
    border.LineStyle = 1  # xlContinuous (solid line)
    border.Weight = 2  # xlMedium (thicker border)


def automize_selection_revise(wb, pu_code, rubber_code, fabric_code, liner_code, uv_code, liner_texture_code, sheet_name = None):
    pu_thickness = pu_df.filter(pl.col("Short Name") == pu_code).select("THICKNESS_mm").item()
    rubber_thickness = rubber_df.filter(pl.col("Short Name") == rubber_code).select("THICKNESS_mm").item()
    fabric_thickness = fabric_df.filter(pl.col("Short Name") == fabric_code).select("THICKNESS_mm").item()
    liner_thickness = liner_df.filter(pl.col("Short Name") == liner_code).select("THICKNESS_mm").item()
    uv_thickness = uv_print_df.filter(pl.col("Short Name") == uv_code).select("THICKNESS_mm").item()

    total_thickness = pu_thickness + rubber_thickness + fabric_thickness + liner_thickness + uv_thickness

    pu_weight = pu_df.filter(pl.col("Short Name") == pu_code).select("WEIGHT_gsm").item()
    rubber_weight = rubber_df.filter(pl.col("Short Name") == rubber_code).select("WEIGHT_gsm").item()
    fabric_weight = fabric_df.filter(pl.col("Short Name") == fabric_code).select("WEIGHT_gsm").item()
    liner_weight = liner_df.filter(pl.col("Short Name") == liner_code).select("WEIGHT_gsm").item()
    uv_weight = uv_print_df.filter(pl.col("Short Name") == uv_code).select("WEIGHT_gsm").item()

    total_weight = pu_weight + rubber_weight + fabric_weight + liner_weight + uv_weight

    pu_weight_percent = round(pu_weight / total_weight * 100)
    rubber_weight_percent = round(rubber_weight / total_weight * 100)
    fabric_weight_percent = round(fabric_weight / total_weight * 100)
    liner_weight_percent = round(liner_weight / total_weight * 100)
    uv_weight_percent = round(uv_weight / total_weight * 100)
    total_weight_percentage = pu_weight_percent + rubber_weight_percent + fabric_weight_percent + liner_weight_percent + uv_weight_percent

    if total_weight_percentage < 100:
        rubber_weight_percent = rubber_weight_percent + (100 - total_weight_percentage)
    elif total_weight_percentage > 100:
        rubber_weight_percent = rubber_weight_percent - (total_weight_percentage - 100)

    uv_pu_weight_percent = uv_weight_percent + pu_weight_percent

    biobased = 0

    if fabric_df.filter(pl.col("Short Name") == fabric_code).select("unnatural").item() == 1 & liner_df.filter(pl.col("Short Name") == liner_code).select("unnatural").item() == 1:
        biobased = round(pu_weight_percent*0.94, 2)
    elif fabric_df.filter(pl.col("Short Name") == fabric_code).select("unnatural").item() == 0 & liner_df.filter(pl.col("Short Name") == liner_code).select("unnatural").item() == 1:
        biobased = round(pu_weight_percent*0.94 + fabric_weight_percent, 2)
    elif fabric_df.filter(pl.col("Short Name") == fabric_code).select("unnatural").item() == 1 & liner_df.filter(pl.col("Short Name") == liner_code).select("unnatural").item() == 0:
        biobased = round(pu_weight_percent*0.94 + liner_weight_percent, 2)
    else :
        biobased = round(pu_weight_percent*0.94 + fabric_weight_percent + liner_weight_percent, 2)

    sheet_names = [sht.name for sht in wb.sheets]
    
    if sheet_name in sheet_names:
        wb.sheets[sheet_name].delete()  # Delete the existing sheet
    
    sht = wb.sheets.add(name=sheet_name)

    # Set the worksheet background color
    sht.range("A1:Z50").color = (215, 228, 202)  # Light green background

    liner_texture = liner_texture_df.filter(pl.col('Surface Texture & Finish') == liner_texture_code).select(pl.col('Texture No.')).item()
    liner_number = liner_df.filter(pl.col('Short Name') == liner_code).select(pl.col('Natuura_key')).item()


    #Layer
    rubber  = rubber_df.filter(pl.col('Short Name') == rubber_code).select(pl.col('Short Name')).item()
    fabric  = fabric_df.filter(pl.col('Short Name') == fabric_code).select(pl.col('Short Name')).item()
    liner  = liner_df.filter(pl.col('Short Name') == liner_code).select(pl.col('Short Name')).item()
    uv = uv_print_df.filter(pl.col('Short Name') == uv_code).select(pl.col('Short Name')).item()

    sht.range("A1:D1").value = ["Company\nTORY BURCH", "BRAND\nLUCID MAT™", "STYLE\nPatent", "DATE\n2025-01-20"]
    sht.range("A2:D2").value = ["FINISH UV\n printing/glossy", f"TEXTURE\n{liner_texture}", f"LINER NO.\n{liner_number}", f"WEIGHT\n{total_weight}gsm ± 20%"]

    sht.range("A1:D2").api.WrapText = True
    sht.range("A1:D2").row_height = None
    sht.range("A1:D2").column_width = None


    # Rows with 2 columns (spanning full width of 4 columns)
    sht.range("A3:C3").merge()
    sht.range("A3").value = [f"COLOR\n"]

    sht.range("D3").value = [f"THICKNESS\n{total_thickness}mm ± 0.2mm"]

    sht.range("A4:C9").merge()
    sht.range("A4").api.WrapText = True

    sht.range("D4:D9").merge()
    sht.range("D4").api.WrapText = True
    sht.range("A10:C10").merge()

    if uv != 'NA':
        sht.range("A4").value = [f"""COMPOSITION
{rubber}: 
{fabric}: 
{liner}:
DRY POLYURETHANE +TOP COATING & UV PRINT:"""]
        sht.range("A10").value = "ESTIMATED BIOBASED:"
        sht.range("D4").value = [f"""Weight %
{rubber_weight_percent}%
{fabric_weight_percent}%
{liner_weight_percent}%

{uv_pu_weight_percent}%"""]
        sht.range("D10").value = f"{biobased}%"
    else:
        sht.range("A4").value = [f"""COMPOSITION
{rubber}:
{fabric}:
{liner}:
    DRY POLYURETHANE:"""]
        sht.range("A10").value = "ESTIMATED BIOBASED:"
        sht.range("D4").value = [f"""Weight %
{rubber_weight_percent}%
{fabric_weight_percent}%
{liner_weight_percent}%

{pu_weight_percent}%"""]
        sht.range("D10").value = f"{biobased}%"

    sht.range("D10").api.HorizontalAlignment = -4131  

    sht.range("A11:D15").merge()
    sht.range("A11").api.WrapText = True
    sht.range("A11").row_height = None
    sht.range("A11").column_width = None
    sht.range("A11").value = """MOQ / MCQ, SIZE, LEAD TIME
MOQ – 300m ; MCQ – 300m
Bulk Lead Time – 2 Months (with In-Stock Liner)
Usable Width – 1.32m (Up to 30m Per Roll)
    """

    sht.range("A16:D23").merge()
    sht.range("A16").api.WrapText = True
    sht.range("A16").row_height = None
    sht.range("A16").column_width = None
    sht.range("A16").value = """OPTIONS
Overall Thickness – 1.0mm – 2.0mm
Pigmentation – Pantone, Custom
Variety of Continuous Textures
Liners – Biobased - 100% GOTS Organic Cotton, 100% Cotton, 100% TENCEL™ (Canvas or Interlock)
        Recycled - 100% GRS Post-Consumer Recycled Polyester
        (Knit or Suede)
    """

    sht.range("A1").column_width = 12.56
    sht.range("B1").column_width = 11.44
    sht.range("C1").column_width = 11.44
    sht.range("D1").column_width = 15.22

    # Forpu cells with borders
    last_row = 23  # Example last row; adjust as needed
    for row in range(1, last_row + 1):
        sht.range(f"A{row}:D{row}").api.Borders.Weight = 2  # Add borders

    # Left, Top, Bottom, Right 9,10,11,12

    sht.range("A10:D10").api.Insert()

    sht.range("D4:D11").api.Borders(7).LineStyle = -4142
    sht.range("A10:D11").api.Borders(8).LineStyle = -4142

    for i in [9, 10, 11, 12]:  # Left, Top, Bottom, Right
        sht.range("A10:D10").api.Borders(i).LineStyle = -4142

    # Force an Excel refresh to apply changes
    wb.app.calculate()

    # Reapply only the right border to D10
    border = sht.range("D10:D10").api.Borders(10)  # Right border
    border.LineStyle = 1  # xlContinuous (solid line)
    border.Weight = 2  # xlMedium (thicker border)


def close_excel():
    try:
        wb.close()
        app.quit()
    except:
        pass  # Ignore errors if Excel is already closed


# Example file path (replace with your file path)
file_path = "data.xlsx"

# Load DataFrames
pu_df = pl.read_excel(file_path, sheet_name="PU")
rubber_df = pl.read_excel(file_path, sheet_name="Rubber")
fabric_df = pl.read_excel(file_path, sheet_name="Fabric")
liner_df = pl.read_excel(file_path, sheet_name="Liner")
liner_texture_df = pl.read_excel(file_path, sheet_name="liner_texture")
uv_print_df = pl.read_excel(file_path, sheet_name="UV_print")

fabric_df = fabric_df.with_columns(
    unnatural = (pl.col("Short Name").str.contains("PET", strict=False)).cast(pl.Int8())
)
liner_df = liner_df.with_columns(
    unnatural = (pl.col("Short Name").str.contains("PET", strict=False)).cast(pl.Int8())
)

# Define dropdown data
dropdown_dataframes = [pu_df, rubber_df, fabric_df, liner_df, uv_print_df, liner_texture_df]
dropdown_columns = ["Short Name", "Short Name", "Short Name", "Short Name", "Short Name", "Surface Texture & Finish"]
dropdown_titles = [
    "Select PU Option", 
    "Select Rubber Option", 
    "Select Fabric Option", 
    "Select Liner Option", 
    "Select UV Print Option", 
    "Select Liner Texture Option"
]

# ✅ **Initialize Session State**
if "user_selections" not in st.session_state:
    st.session_state.user_selections = [None] * 6
if "sheet_name" not in st.session_state:
    st.session_state.sheet_name = ""

app = xw.App(visible=False)
try:
    wb = xw.Book("value.xlsx")
except:
    wb = xw.Book()


atexit.register(close_excel)

# Function to map selection to SurrogateKey
def map_selection_to_code(selected_value, df, col_name):
    try:
        return df.filter(pl.col(col_name) == selected_value).select("SurrogateKey").item()
    except:
        return None  # Return None if not found

# Streamlit UI
st.title("Multi-Option Selection")

num_options = len(dropdown_dataframes)  # Automatically set size based on dropdowns
if "user_selections" not in st.session_state:
    st.session_state.user_selections = [None] * num_options

# 🎯 **Display all selections on the same page**
for i, df in enumerate(dropdown_dataframes):
    col_name = dropdown_columns[i]
    title = dropdown_titles[i]

    selected_value = st.selectbox(title, df[col_name].to_list(), key=f"selection_{i}")

    # Save mapped code to session state
    st.session_state.user_selections[i] = map_selection_to_code(selected_value, df, col_name)

# 📌 **Sheet Name Input**
sheet_name = st.text_input("Enter Sheet Name", value=st.session_state.get("sheet_name", ""))
st.session_state.sheet_name = sheet_name

# ✅ **Run Automation**
if st.button("Run Automation"):
    if None in st.session_state.user_selections or not st.session_state.sheet_name:
        st.error("⚠️ Please make all selections before running automation.")
    else:
        # Run automation function
        automize_selection(
            wb, 
            pu_code=st.session_state.user_selections[0], 
            rubber_code=st.session_state.user_selections[1], 
            fabric_code=st.session_state.user_selections[2], 
            liner_code=st.session_state.user_selections[3], 
            uv_code=st.session_state.user_selections[4], 
            liner_texture_code=st.session_state.user_selections[5], 
            sheet_name=st.session_state.sheet_name
        )
        save_path = "little_card.xlsx"
        # Save workbook
        try:
            wb.save(save_path)
            time.sleep(1)  # Small delay to allow system to release the file
            wb.close()
            app.quit()
        except Exception as e:
            st.error(f"❌ Failed to save file: {e}")
            st.warning("🔄 Trying to close Excel forcefully...")
            os.system("taskkill /f /im excel.exe")  # Kill any Excel processes
            time.sleep(2)  # Wait before retrying
            try:
                wb.save(save_path)  # Try saving again
                wb.close()
                app.quit()
                st.success(f"✅ File saved successfully after resolving lock: {save_path}")
            except Exception as e:
                st.error(f"❌ Could not save file after retry: {e}")

        st.success("🎉 Automation Completed! Data has been written to Excel.")

# 🔄 **Restart Button**
if st.button("Restart"):
    st.session_state.user_selections = [None] * num_options
    st.session_state.sheet_name = ""
    st.rerun()
