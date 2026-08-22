import streamlit as st
import openpyxl
import os
import google.generativeai as genai
from PIL import Image

# အစ်ကို ပေးထားသော API Key ကို အသုံးပြုခြင်း
GOOGLE_API_KEY = "AQ.Ab8RN6Is09pGou__VTes-0VEBCGgh4hW_XS3tJfUMr9lFZ7WNQ" 
genai.configure(api_key=GOOGLE_API_KEY)
model = genai.GenerativeModel('gemini-1.5-flash')

# Website ၏ မျက်နှာပြင် ဒီဇိုင်း
st.set_page_config(page_title="SGS Report Automation", page_icon="📄")
st.title("📄 SGS Report Automation")
st.write("ကျေးဇူးပြု၍ SGS Report ဓာတ်ပုံ (၅) ပုံ (Stuffing, Weight List, Tally, VGM, Container Survey) ကို အောက်တွင် Upload လုပ်ပါ။")

# ဓာတ်ပုံများ Upload လုပ်ရန်
uploaded_files = st.file_uploader(
    "ဓာတ်ပုံများ ရွေးချယ်ရန် ဤနေရာကို နှိပ်ပါ", 
    accept_multiple_files=True, 
    type=['jpg', 'jpeg', 'png']
)

def extract_data_from_image(image, prompt):
    """ဓာတ်ပုံမှ Data ဖတ်ရန် AI ကို ခိုင်းစေသော Function"""
    response = model.generate_content([prompt, image])
    return response.text.strip()

def process_and_generate_excel(files):
    """Excel ထဲသို့ Data များ ဖြည့်သွင်းသော Function"""
    # Template ဖိုင်ကို ခေါ်ယူခြင်း
    template_path = "(30 FCL)SIR TEMPLATE.xlsx"
    wb = openpyxl.load_workbook(template_path)
    
    # ---------------------------------------------------------
    # AI ဖြင့် ဓာတ်ပုံများကို ဖတ်ပြီး Data ယူခြင်း (ဥပမာ အကြမ်းဖျင်း)
    # (တကယ့်အလုပ်လုပ်ချိန်တွင် ဓာတ်ပုံအလိုက် Prompt များ ခွဲရေးရပါမည်)
    # ---------------------------------------------------------
    
    # ပထမ ဓာတ်ပုံကို ယူ၍ စမ်းသပ်ဖတ်ခိုင်းခြင်း
    img = Image.open(files[0])
    
    # AI ကို ဘာ Data ယူရမလဲ ညွှန်ကြားခြင်း (ဥပမာ - Container Number ယူခိုင်းခြင်း)
    prompt = "Extract the Container Number from this image. Return ONLY the container number, nothing else."
    container_number = extract_data_from_image(img, prompt)
    
    # ---------------------------------------------------------
    # Excel Sheets များထဲသို့ နေရာချခြင်း
    # ---------------------------------------------------------
    
    # ၁။ STUFFING REPORT Sheet သို့ Data ဖြည့်ခြင်း
    ws_stuffing = wb['STUFFING REPORT']
    # ဥပမာ - Container Number ကို B19 တွင် ထည့်ခြင်း
    ws_stuffing['B19'] = container_number 
    
    # (မှတ်ချက် - ကျန်သော Weight List, Tally, VGM, Survey Sheet များအတွက်လည်း
    # ဤနေရာတွင် ထပ်မံဖြည့်သွင်းရပါမည်)
    
    # ဖိုင်ကို Save ခြင်း
    output_filename = "Completed_SGS_Report.xlsx"
    wb.save(output_filename)
    return output_filename


# ဓာတ်ပုံအရေအတွက် စစ်ဆေးခြင်း
if uploaded_files:
    if len(uploaded_files) == 5:
        st.success("✅ ဓာတ်ပုံ ၅ ပုံ ပြည့်ပါပြီ။ အောက်ပါ 'Excel ဖန်တီးရန်' ခလုတ်ကို နှိပ်ပါ။")
        
        if st.button("🚀 Excel ဖန်တီးရန်", type="primary"):
            with st.spinner('AI မှ ဓာတ်ပုံများကို ဖတ်၍ Excel ထဲသို့ ထည့်သွင်းနေပါသည်... စက္ကန့်အနည်းငယ် စောင့်ပါ။'):
                try:
                    output_file = process_and_generate_excel(uploaded_files)
                    st.success("🎉 Excel ဖိုင် ဖန်တီးမှု အောင်မြင်ပါပြီ။ အောက်တွင် Download ဆွဲပါ။")
                    
                    # Excel ဖိုင်ကို Download ဆွဲရန်
                    with open(output_file, "rb") as file:
                        st.download_button(
                            label="📥 Completed_SGS_Report.xlsx ကို ဒေါင်းလုဒ်ဆွဲရန်",
                            data=file,
                            file_name="Completed_SGS_Report.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                except Exception as e:
                    st.error(f"Error ဖြစ်သွားပါသည်: {e}. Template ဖိုင်နာမည် အမှန်တကယ် ရှိမရှိ စစ်ဆေးပါ။")
    else:
        st.warning(f"⚠️ ဓာတ်ပုံ (၅) ပုံ တိတိ လိုအပ်ပါသည်။ ယခု ({len(uploaded_files)}) ပုံသာ တင်ထားပါသည်။")
