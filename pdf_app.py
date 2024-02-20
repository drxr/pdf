import streamlit as st
import pandas as pd
import PyPDF2
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from datetime import datetime
import zipfile


def seek(string: str, index: int) -> int:
    'search for words indexes that comes after spaces'
    if string[index-1] == ' ':
        return index
    return seek(string, index-1)


def wrap(string:str, length: int) -> str:
    'splits text into lines by length'
    if len(string) < length:
        return string
    pos = seek(string, length)
    return string[:pos:] + '\n' + wrap(string[pos::], length)


PdfWriter = PyPDF2.PdfWriter
PdfReader = PyPDF2.PdfReader

# getting current date
nowaday = datetime.today().date()
now_month = datetime.today().month
now_year = datetime.today().year
current_date = datetime(now_year, now_month, 1).date()

# list of pdf files to download
pdf_list = []

# intro text
st.title('**Интерфейс для создания благодарственных писем студентам**')
st.text('')
st.text('Для корректной работы скрипта необходимо указать название проекта,')
st.text('даты начала и окончания проекта, а также внести студентов в поле')
st.text('каждого с новой строки')
st.text('')


# making variables
project_name = st.text_input('Введите название проекта:', value=None)

if project_name is not None:
    st.write(f'**Название проекта:** {project_name}')

start_date, finish_date = st.date_input('Введите даты начала и окончания проекта:', (current_date, nowaday), format='DD.MM.YYYY')
start_date = start_date.strftime('%d.%m.%Y')
finish_date = finish_date.strftime('%d.%m.%Y')

if start_date is not None:
    st.write(f'Период проекта с {start_date} по {finish_date}')

students_raw = st.text_area('Введите имена и фамилии студентов с новой строки:', value=None)

if students_raw is not None:
    students = students_raw.split('\n')
    st.write(f'Добавлено {len(students)} студентов')
    
# text of letter
text_1 = f'Мастерская данных в лице Руководителя Мастерской данных Богданова Руслана Александровича благодарит вас за отличную работу над проектом «{project_name}» с {start_date} по {finish_date}.'
text_2 = f'Отдельное спасибо за вашу вовлеченность, творческий подход к поиску оптимального решения и проявленный профессионализм.'

# text split
st.text('')

# pdf scheme
if students_raw is not None:
    if len(project_name) > 71:
        text_pos = 260
    else:
        text_pos = 270
    for student in students:
    
        # adding file name to pdf list
        pdf_list.append(student + '.pdf')
    
        # setting up cyrillic and font 
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=letter)
        cyrenc = pdfmetrics.Encoding('utf8')
        pdfmetrics.registerFont(TTFont('YS Text Regular', 'YS Text Regular Regular.ttf'))
    
        # printing name
        can.setFont('YS Text Regular', 22)
        can.setFillColor('white')
        can.drawString(70, 400, student)
        can.setFont('YS Text Regular', 14)
        can.setFillColor('white')
        can.drawString(485, 64, nowaday.strftime('%d.%m.%Y'))
            
        # printing main text
        text = can.beginText()
        text.setFont('YS Text Regular', 16)
        #text.setLeading(15)
        text.setTextOrigin(70, 360)
        text.textLines(wrap(text_1, 62))
        text.setTextOrigin(70, text_pos)
        text.textLines(wrap(text_2, 62))
        can.drawText(text)
        
        # saving
        can.showPage()
        can.save()
        
        # moving to the beggining of the string
        packet.seek(0)
        
        # creating empty pdf
        new_pdf = PdfReader(packet)
        
        # reading target pdf
        existing_pdf = PdfReader('sample_pdf.pdf', "rb")
        output = PdfWriter()
    
        # merging pdfs
        page = existing_pdf.pages[0]
        page.merge_page(new_pdf.pages[0])
        output.add_page(page)
    
        # saving final pdf
        output_stream = open(student + ".pdf", "wb")
        output.write(output_stream)
        output_stream.close()

# creating zip archive and adding files from pdf list into it
with zipfile.ZipFile('archive.zip', 'w') as myzip:
    for file in pdf_list:
        myzip.write(file)

# button text
st.text('Скачать архив с грамотами')

# download archive button
st.download_button(
    help='Файл с архивом благодарственных писем для студентов',
    label="Скачать архив :sunglasses:",
    data=open('archive.zip', 'rb').read(),
    file_name='archive.zip',
    mime='application/zip'
)
