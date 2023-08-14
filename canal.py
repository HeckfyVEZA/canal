from pandas import DataFrame, ExcelWriter, set_option
import streamlit as st
from info_search import infos
from groupy import grouping
from check_file import check
st.set_page_config(layout='wide')
unprocessed = []
def to_excel(df, HEADER=False, START=1):
    output = __import__("io").BytesIO()
    writer = ExcelWriter(output, engine='xlsxwriter')
    DataFrame(df).to_excel(writer, index=False, header=HEADER,startrow=START, startcol=START, sheet_name='Sheet1')
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    format1 = workbook.add_format({'num_format': '0.00'})
    worksheet.set_column('A:A', None, format1)
    writer.close()
    return output.getvalue()
all_noms = []
st.session_state.uploaded_files = st.file_uploader('TECHNICAL SHEETS', accept_multiple_files=True)
p = 0
lp = len(st.session_state.uploaded_files)
if lp:
    pr_bar = st.progress(p, '')
exp = st.expander('Системы')
for file in st.session_state.uploaded_files:
    p+=1
    try:
        curinfo = infos(file)
        all_noms+=curinfo
        exp.markdown(f'<h2>Система {curinfo[0][0]}</h2>', unsafe_allow_html=True)
        for ci in curinfo:
            exp.markdown(f'<p><h4>{ci[1]}</h4>   <i>{ci[-1]} шт.</i></p>', unsafe_allow_html=True)
        exp.markdown('---')
    except:
        unprocessed.append(file.name)
    if lp:
        pr_bar.progress(p/lp, f"Обрабатывается {file.name}")
if lp:
    pr_bar.progress(p/lp, 'Готово')
if len(unprocessed):
    st.markdown('<h3>Необработанные файлы</h3>', unsafe_allow_html=True)
    for upcs in unprocessed:
        st.markdown(f"<h5>{upcs}</h5>", unsafe_allow_html=True)
all_nomenclature_names = list(set(list(map(lambda x: x[1], all_noms))))
gnoms = grouping(all_noms)
st.expander('Таблица').table(gnoms)
st.download_button(label='💾 Скачать файл для выгрузки в КП',data=to_excel(grouping(all_noms)), file_name= 'для кп.xls')
st.download_button(label='💾 Скачать файл для выгрузки в КП (По системам)',data=to_excel(list(map(lambda x: [x[1], x[2], x[0]], all_noms))), file_name= 'для кп.xls')
st.download_button(label='💾 Скачать проверочный файл',data=to_excel(DataFrame(check(all_noms)),START=0) ,file_name= 'проверка.xlsx')
