import streamlit as st
import openpyxl
import os
import json
import google.generativeai as genai
from PIL import Image

# အစ်ကို၏ API Key
GOOGLE_API_KEY = "AQ.Ab8RN6Is09pGou__VTes-0VEBCGgh4hW_XS3tJfUMr9lFZ7WNQ" 
genai.configure(api_key=GOOGLE_API_KEY)
# နောက်ဆုံးပေါ် Vision Model ကို အသုံးပြုခြင်း
model = genai.GenerativeModel('gemini-1.5-flash')

st.set_page_config(page_title="SGS Report Automation", page_icon="📄", layout="wide")
st.title("📄 SGS Report Automation")
st.write("ကျေးဇူးပြု၍ SGS Report ဓာတ်ပုံ (၅) ပုံ (Stuffing, Weight List, Tally, VGM, Container Survey) ကို အောက်တွင် Upload လုပ်ပါ။")

uploaded_files = st.file_uploader(
    "ဓာတ်ပုံများ ရွေးချယ်ရန် ဤနေရာကို နှိပ်ပါ", 
    accept_multiple_files=True, 
    type=['jpg', 'jpeg', 'png']
)

def extract_data_with_ai(images):
    """ဓာတ်ပုံ ၅ ပုံကို ဖတ်၍ လိုအပ်သော Data များကို JSON အဖြစ် ဆွဲထုတ်မည့် Function"""
    
    # ဓာတ်ပုံများကို AI နားလည်သော ပုံစံသို့ ပြောင်းခြင်း
    pil_images = [Image.open(file) for file in images]
    
    # AI အား မည်သည့် Data များ ဆွဲထုတ်ရမည်ကို ညွှန်ကြားခြင်း (Prompt)
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
    
    # Prompt နှင့် ဓာတ်ပုံများကို AI ထံ ပို့၍ အဖြေတောင်းခြင်း
    response = model.generate_content([prompt] + pil_images)
    
    # AI ပြန်ပေးသော JSON စာသားကို ယူခြင်း
    result_text = response.text.strip()
    # Markdown (```json ... ```) များ ပါလာပါက ဖြတ်ထုတ်ခြင်း
    if result_text.startswith("```json"):
        result_text = result_text[7:-3].strip()
    elif result_text.startswith("```"):
        result_text = result_text[3:-3].strip()
        
    return json.loads(result_text)

def process_and_generate_excel(files):
    # Template ဖိုင်ကို ခေါ်ယူခြင်း
    template_path = "(30 FCL)SIR TEMPLATE.xlsx"
    wb = openpyxl.load_workbook(template_path)
    
    # ၁။ AI ဖြင့် Data များကို ဆွဲထုတ်ခြင်း
    extracted_data = extract_data_with_ai(files)
    
    # ၂။ 'STUFFING REPORT' Sheet ထဲသို့ Data များ ဖြည့်သွင်းခြင်း
    ws_stuffing = wb['STUFFING REPORT']
    
    # အထွေထွေ အချက်အလက်များ ဖြည့်ခြင်း (ဆဲလ် နေရာများကို လိုအပ်ပါက ပြင်နိုင်ပါသည်)
    ws_stuffing['D4'] = extracted_data.get('job_ref', '')
    ws_stuffing['C7'] = extracted_data.get('shipper', '')
    ws_stuffing['C8'] = extracted_data.get('buyer', '')
    ws_stuffing['C9'] = extracted_data.get('commodity', '')
    ws_stuffing['C10'] = extracted_data.get('quantity', '')
    
    # Container အချက်အလက်များကို ဇယားကွက်ထဲသို့ အစဉ်လိုက် ဖြည့်ခြင်း
    # Excel ၏ Container ဇယားသည် Row 20 (Excel index အရ) မှ စသည်ဟု ယူဆပါသည်
    start_row = 20 
    for i, container in enumerate(extracted_data.get('containers', [])):
        current_row = start_row + i
        ws_stuffing[f'B{current_row}'] = container.get('container_no', '')
        ws_stuffing[f'C{current_row}'] = container.get('condition', '')
        ws_stuffing[f'F{current_row}'] = container.get('bags', '')
        ws_stuffing[f'G{current_row}'] = container.get('gross_mt', '')
        ws_stuffing[f'H{current_row}'] = container.get('tare_mt', '')
        ws_stuffing[f'I{current_row}'] = container.get('seal_no', '')

    # ၃။ ပြီးစီးပါက ဖိုင်အသစ်အဖြစ် Save ခြင်း
    output_filename = f"Completed_SGS_{extracted_data.get('job_ref', 'Report')}.xlsx"
    wb.save(output_filename)
    return output_filename

# ဓာတ်ပုံအရေအတွက် စစ်ဆေးခြင်း နှင့် ခလုတ်နှိပ်ခြင်း
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
