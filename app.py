import streamlit as st
import openpyxl
import json
import base64
import requests

st.set_page_config(page_title="SGS Report Automation", page_icon="📄", layout="wide")
st.title("📄 SGS Report Automation")
st.write("ကျေးဇူးပြု၍ SGS Report ဓာတ်ပုံ (၅) ပုံ (Stuffing, Weight List, Tally, VGM, Container Survey) ကို အောက်တွင် Upload လုပ်ပါ။")

uploaded_files = st.file_uploader(
    "ဓာတ်ပုံများ ရွေးချယ်ရန် ဤနေရာကို နှိပ်ပါ", 
    accept_multiple_files=True, 
    type=['jpg', 'jpeg', 'png']
)

def extract_data_with_gemini(images):
    # Streamlit Secrets မှ AQ... Key ကို ယူခြင်း
    API_KEY = st.secrets["GEMINI_API_KEY"]
    
    # Python SDK ကို မသုံးဘဲ URL ဖြင့် Google Server သို့ တိုက်ရိုက် လှမ်းချိတ်ခြင်း
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={API_KEY}"
    
    parts = []
    prompt = """
    Analyze these 5 images of SGS reports (Stuffing, Weight List, Tally, VGM, Survey).
    Extract the information and return STRICTLY a JSON object with the following structure. Do not output any markdown formatting like ```json , just the raw JSON text.
    {
        "job_ref": "Job Ref number",
        "shipper": "Shipper name",
        "buyer": "Buyer name",
        "commodity": "Commodity name",
        "quantity": "Quantity with units",
        "containers": [
            {
                "container_no": "Container Number (e.g. TCNU-5996424)",
                "condition": "Condition (e.g. Sound)",
                "bags": 520,
                "gross_mt": 26.006,
                "tare_mt": 0.006,
                "seal_no": "Seal Number"
            }
        ]
    }
    Extract for all containers listed in the report.
    """
    parts.append({"text": prompt})
    
    for img_file in images:
        img_bytes = img_file.getvalue()
        img_b64 = base64.b64encode(img_bytes).decode('utf-8')
        mime_type = img_file.type
        parts.append({
            "inlineData": {
                "mimeType": mime_type,
                "data": img_b64
            }
        })
        
    payload = {
        "contents": [{"parts": parts}]
    }
    
    # Server သို့ Request ပို့ခြင်း
    response = requests.post(url, json=payload)
    if response.status_code != 200:
        raise Exception(f"API Error: {response.text}")
        
    result_text = response.json()['candidates'][0]['content']['parts'][0]['text'].strip()
    
    if result_text.startswith("```json"):
        result_text = result_text[7:-3].strip()
    elif result_text.startswith("```"):
        result_text = result_text[3:-3].strip()
        
    return json.loads(result_text)

# Excel ၏ Merge လုပ်ထားသော အကွက်များကို အလိုအလျောက် ဖြေရှင်းပေးမည့် Function
def safe_write(ws, coord, value):
    for merged_range in ws.merged_cells.ranges:
        if coord in merged_range:
            ws.cell(row=merged_range.min_row, column=merged_range.min_col).value = value
            return
    ws[coord].value = value

def process_and_generate_excel(files):
    template_path = "(30 FCL)SIR TEMPLATE.xlsx"
    wb = openpyxl.load_workbook(template_path)
    
    # AI ဖြင့် ဓာတ်ပုံထဲမှ Data အစစ်များကို ဆွဲထုတ်ခြင်း
    extracted_data = extract_data_with_gemini(files)
    ws_stuffing = wb['STUFFING REPORT']
    
    # safe_write ဖြင့် မှန်ကန်သော နေရာများသို့ ရေးသွင်းခြင်း
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
            with st.spinner('AI မှ ဓာတ်ပုံများကို ဖတ်၍ Excel ထဲသို့ ထည့်သွင်းနေပါသည်... စက္ကန့် ၂၀ ခန့် စောင့်ပါ။'):
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
