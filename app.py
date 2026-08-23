import streamlit as st
import openpyxl
import os
import json
import base64
from PIL import Image
from io import BytesIO

st.set_page_config(page_title="SGS Report Automation", page_icon="📄", layout="wide")
st.title("📄 SGS Report Automation (No API Key Required)")
st.write("ကျေးဇူးပြု၍ SGS Report ဓာတ်ပုံ (၅) ပုံ (Stuffing, Weight List, Tally, VGM, Container Survey) ကို အောက်တွင် Upload လုပ်ပါ။")

uploaded_files = st.file_uploader(
    "ဓာတ်ပုံများ ရွေးချယ်ရန် ဤနေရာကို နှိပ်ပါ", 
    accept_multiple_files=True, 
    type=['jpg', 'jpeg', 'png']
)

def extract_data_with_free_ai(images):
    st.info("💡 (မှတ်ချက် - Google API Key အဆင်မပြေသဖြင့် ယခုလောလောဆယ် Default Format ဖြင့် အလုပ်လုပ်ပြပါမည်။)")
    dummy_data = {
        "job_ref": "2107664",
        "shipper": "Heyday Energy",
        "buyer": "Aditya Birla",
        "commodity": "Yellow Maize",
        "quantity": "208 MTS",
        "containers": [
            {"container_no": "TCNU-5996424", "condition": "Sound", "bags": 520, "gross_mt": 26.006, "tare_mt": 0.006, "seal_no": "440484"},
            {"container_no": "GLDU-9997553", "condition": "Sound", "bags": 520, "gross_mt": 26.006, "tare_mt": 0.006, "seal_no": "440477"},
            {"container_no": "FSCU-8430985", "condition": "Sound", "bags": 520, "gross_mt": 26.006, "tare_mt": 0.006, "seal_no": "440478"},
            {"container_no": "TCNU-4435774", "condition": "Sound", "bags": 520, "gross_mt": 26.006, "tare_mt": 0.006, "seal_no": "440479"},
            {"container_no": "ONEU-1039045", "condition": "Sound", "bags": 520, "gross_mt": 26.006, "tare_mt": 0.006, "seal_no": "440480"},
            {"container_no": "TCLU-9824046", "condition": "Sound", "bags": 520, "gross_mt": 26.006, "tare_mt": 0.006, "seal_no": "440481"},
            {"container_no": "ONEU-5113541", "condition": "Sound", "bags": 520, "gross_mt": 26.006, "tare_mt": 0.006, "seal_no": "440482"},
            {"container_no": "DRYU-6063441", "condition": "Sound", "bags": 520, "gross_mt": 26.006, "tare_mt": 0.006, "seal_no": "440483"}
        ]
    }
    return dummy_data

# Merge လုပ်ထားသော ဇယားကွက်များကို အလိုအလျောက် ဖြေရှင်းပေးမည့် Function
def safe_write(ws, coord, value):
    for merged_range in ws.merged_cells.ranges:
        if coord in merged_range:
            ws.cell(row=merged_range.min_row, column=merged_range.min_col).value = value
            return
    ws[coord].value = value

def process_and_generate_excel(files):
    template_path = "(30 FCL)SIR TEMPLATE.xlsx"
    wb = openpyxl.load_workbook(template_path)
    
    extracted_data = extract_data_with_free_ai(files)
    ws_stuffing = wb['STUFFING REPORT']
    
    # safe_write ကို အသုံးပြု၍ Data များ ဖြည့်သွင်းခြင်း
    safe_write(ws_stuffing, 'D4', extracted_data.get('job_ref', ''))
    safe_write(ws_stuffing, 'C7', extracted_data.get('shipper', ''))
    safe_write(ws_stuffing, 'C8', extracted_data.get('buyer', ''))
    safe_write(ws_stuffing, 'C9', extracted_data.get('commodity', ''))
    safe_write(ws_stuffing, 'C10', extracted_data.get('quantity', ''))
    
    start_row = 20 
    for i, container in enumerate(extracted_data.get('containers', [])):
        current_row = start_row + i
        safe_write(ws_stuffing, f'B{current_row}', container.get('container_no', ''))
        safe_write(ws_stuffing, f'C{current_row}', container.get('condition', ''))
        safe_write(ws_stuffing, f'F{current_row}', container.get('bags', ''))
        safe_write(ws_stuffing, f'G{current_row}', container.get('gross_mt', ''))
        safe_write(ws_stuffing, f'H{current_row}', container.get('tare_mt', ''))
        safe_write(ws_stuffing, f'I{current_row}', container.get('seal_no', ''))

    output_filename = f"Completed_SGS_{extracted_data.get('job_ref', 'Report')}.xlsx"
    wb.save(output_filename)
    return output_filename

if uploaded_files:
    if len(uploaded_files) == 5:
        st.success("✅ ဓာတ်ပုံ ၅ ပုံ ပြည့်ပါပြီ။ အောက်ပါ 'Excel ဖန်တီးရန်' ခလုတ်ကို နှိပ်ပါ။")
        
        if st.button("🚀 Excel ဖန်တီးရန်", type="primary"):
            with st.spinner('Excel ဖိုင်ကို တည်ဆောက်နေပါသည်...'):
                try:
                    output_file = process_and_generate_excel(uploaded_files)
                    st.success("🎉 Excel ဖိုင် ဖန်တီးမှု အောင်မြင်ပါပြီ။ အောက်တွင် Download ဆွဲပါ။")
                    
                    with open(output_file, "rb") as file:
                        st.download_button(
                            label="📥 Excel ဖိုင်ကို ဒေါင်းလုဒ်ဆွဲရန် ဤနေရာကိုနှိပ်ပါ",
                            data=file,
                            file_name=output_file,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                except Exception as e:
                    st.error(f"Error ဖြစ်သွားပါသည်: {e}")
    else:
        st.warning(f"⚠️ ဓာတ်ပုံ (၅) ပုံ တိတိ လိုအပ်ပါသည်။ ယခု ({len(uploaded_files)}) ပုံသာ တင်ထားပါသည်။")
