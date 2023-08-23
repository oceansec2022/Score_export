import streamlit as st 
import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime
import os
from docx2pdf import convert
import shutil


# Upload file
st.title('XUẤT ĐIỂM')
st.markdown("Lấy file excel mẫu: [Get link](https://docs.google.com/spreadsheets/d/1JWlt470RF-CgiqEJeErCCKYIJcY82G3p/edit?usp=drive_link&ouid=114357774145411752777&rtpof=true&sd=true)")
st.markdown("**Lưu ý:** Để chương trình chạy được phải sử dụng định dạng (tên của từng cột phải giống như file mẫu) như file excel mẫu")
absolute_path = os.path.dirname(__file__)
full_path = os.path.join(absolute_path, "export_pdf")
full_path_zip = os.path.join(absolute_path, "export_zip")  
uploaded_file = st.file_uploader("Choose a file", type=['xlsx','xls'])
if uploaded_file is None:
    st.write("Vui lòng nhập file để xuất file điểm")
else:
    sheetInExcel = pd.ExcelFile(uploaded_file).sheet_names
    sheetInExcel.insert(0,'Chọn sheet')
    sheetName = st.selectbox('Lựa chọn sheet cần xuất file điểm', sheetInExcel)
    doc = DocxTemplate(absolute_path + "\\" + "Finaltest_mau.docx")
    df = pd.read_excel(uploaded_file, sheet_name=sheetName).fillna('')
    st.write(df)
    if st.button('Export PDF'):
        for index, file in df.iterrows():
            if file['Nghe'] != "":
                file['Nghe'] = int(file['Nghe'])
            if file['Nói'] != "":
                file['Nói'] = int(file['Nói'])
            if file['Đọc'] != "":
                file['Đọc'] = int(file['Đọc'])
            if file['Viết'] != "":
                file['Viết'] = int(file['Viết'])
            if file['Tổng'] != "":
                file['Tổng'] = int(file['Tổng'])
            context = {'sClass': file['Lớp'],
                    'datePoint': file['Ngày chấm'],
                    'GVHD': file['GVHD'],
                    'vName': file['Tên tiếng việt'],
                    'eName': file['Tên tiếng Anh'],
                    'gender': file['Giới tính'],
                    'dateReg': file['Ngày đăng kí'],
                    'dateEnd': file['Ngày kết thúc'],
                    'countLearn': file['Số buổi học'],
                    'lis': file['Nghe'],
                    'spe': file['Nói'],
                    'rea': file['Đọc'],
                    'wri': file['Viết'],
                    'total': file['Tổng'],
                    'evaInClass': file['Nhận xét trên lớp'],
                    'evaInTest': file['Nhận xét bài kiểm tra'],
                    'upGrade': file['Được lên lớp']}
            doc.render(context)
            doc.save(f"{full_path}\{index+1}_{file['Tên tiếng Anh']}.docx")
        convert(full_path)
        fileName = f"PDF_{df['Lớp'][0]}_{datetime.now().strftime('%d.%m.%Y')}"
        file_docx = [i for i in os.listdir(full_path) if i.endswith('.docx')]
        for i in range(0,len(file_docx)):
            os.remove(full_path +"/"+ file_docx[i])
        shutil.make_archive(f"{full_path_zip}\\{fileName}", 'zip', root_dir=full_path)
        with open(f"{full_path_zip}\\{fileName}.zip", "rb") as fp:
            st.download_button(
            label="Download zip file",   
            data=fp,
            file_name=f"{fileName}.zip",
            mime="application/zip")
        file_pdf = [i for i in os.listdir(full_path) if i.endswith('.pdf')]
        for i in range(0,len(file_pdf)):
            os.remove(full_path +"/"+ file_pdf[i])  
    else: 
        pass


