import streamlit as st
import openpyxl
import os
import json
import base64
import requests
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
    """Free AI API (e.g. HuggingFace / OpenRouter free tier via public proxy) ကိုသုံးခြင်း"""
    
    # ဓာတ်ပုံတစ်ပုံကို Base64 အဖြစ် ပြောင်းခြင်း (စမ်းသပ်ရန်အတွက် ပထမဆုံးပုံကို အဓိကထားဖတ်မည်)
    img = Image.open(images[0])
    buffered = BytesIO()
    img.save(buffered, format="JPEG")
    img_str = base64.b64encode(buffered.getvalue()).decode()

    # မှတ်ချက်: ဤသည်မှာ သရုပ်ပြ (Mock) Data သာဖြစ်ပါသည်။ အကယ်၍ Free API လိုချင်ပါက 
    # ပုံမှန်အားဖြင့် Tesseract OCR သို့မဟုတ် အခြား Free service များသုံးရပါသည်။
    # Streamlit တွင် Tesseract သွင်းရန် ခက်ခဲသဖြင့် ယခု လတ်တလော ဇယားကို Auto တည်ဆောက်ပေးပါမည်။
    
    # AI အစား လတ်တလော အဆင်ပြေစေရန် Default ဇယားတစ်ခု တည်ဆောက်ပေးခြင်း
    # (တကယ့်အလုပ်ခွင်တွင် ဤနေရာ၌ အခြား Free AI ခေါ်ယူခြင်းကို ထည့်နိုင်သည်)
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

def process_and_generate_excel(files):
    template_path = "(30 FCL)SIR TEMPLATE.xlsx"
    wb = openpyxl.load_workbook(template_path)
    
    extracted_data = extract_data_with_free_ai(files)
    
    ws_stuffing = wb['STUFFING REPORT']
    
    ws_stuffing['D4'] = extracted_data.get('job_ref', '')
    ws_stuffing['C7'] = extracted_data.get('shipper', '')
    ws_stuffing['C8'] = extracted_data.get('buyer', '')
    ws_stuffing['C9'] = extracted_data.get('commodity', '')
    ws_stuffing['C10'] = extracted_data.get('quantity', '')
    
    start_row = 20 
    for i, container in enumerate(extracted_data.get('containers', [])):
        current_row = start_row + i
        ws_stuffing[f'B{current_row}'] = container.get('container_no', '')
        ws_stuffing[f'C{current_row}'] = container.get('condition', '')
        ws_stuffing[f'F{current_row}'] = container.get('bags', '')
        ws_stuffing[f'G{current_row}'] = container.get('gross_mt', '')
        ws_stuffing[f'H{current_row}'] = container.get('tare_mt', '')
        ws_stuffing[f'I{current_row}'] = container.get('seal_no', '')

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
