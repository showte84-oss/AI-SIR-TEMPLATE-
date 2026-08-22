import streamlit as st
import openpyxl
import os

# Website ၏ ခေါင်းစဉ်
st.title("SGS Report Automation")
st.write("ကျေးဇူးပြု၍ SGS Report ဓာတ်ပုံ (၅) ပုံကို အောက်တွင် Upload လုပ်ပါ။")

# ဓာတ်ပုံများ Upload လုပ်ရန် နေရာ
uploaded_files = st.file_uploader(
    "ဓာတ်ပုံများ ရွေးချယ်ရန် ဤနေရာကို နှိပ်ပါ", 
    accept_multiple_files=True, 
    type=['jpg', 'jpeg', 'png']
)

def process_and_generate_excel(files):
    """
    ဤနေရာတွင် အရင်က အစ်ကိုရေးခဲ့သော OCR နှင့် Excel Data ဖြည့်သည့် Code များကို ထည့်ပါ။
    """
    # ဥပမာ - Template ကိုခေါ်ခြင်း
    # wb = openpyxl.load_workbook("30_FCL_Template.xlsx")
    # ws = wb.active
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "SGS Report"
    ws['A1'] = "SGS Report Completed Successfully!"
    
    # ဖိုင်ကို Save ခြင်း
    output_filename = "Completed_SGS_Report.xlsx"
    wb.save(output_filename)
    return output_filename

# ဓာတ်ပုံအရေအတွက် စစ်ဆေးခြင်း
if uploaded_files:
    if len(uploaded_files) == 5:
        st.success("ဓာတ်ပုံ ၅ ပုံ ပြည့်ပါပြီ။ 'Excel ဖန်တီးရန်' ခလုတ်ကို နှိပ်နိုင်ပါပြီ။")
        
        # Excel ဖန်တီးရန် ခလုတ်
        if st.button("Excel ဖန်တီးရန်"):
            with st.spinner('Data များကို Excel ထဲသို့ ထည့်သွင်းနေပါသည်...'):
                output_file = process_and_generate_excel(uploaded_files)
                
                # Excel ဖိုင်ကို Download ဆွဲရန် ခလုတ်
                with open(output_file, "rb") as file:
                    st.download_button(
                        label="📥 Excel ဖိုင်ကို ဒေါင်းလုဒ်ဆွဲရန် နှိပ်ပါ",
                        data=file,
                        file_name="Completed_SGS_Report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
    else:
        st.warning(f"ဓာတ်ပုံ (၅) ပုံ တိတိ လိုအပ်ပါသည်။ ယခု ({len(uploaded_files)}) ပုံသာ တင်ထားပါသည်။")
