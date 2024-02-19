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
nowaday = datetime.today().strftime('%d.%m.%Y')

# list of pdf files to download
pdf_list = []

# intro text
st.title('Интерфейс для создания благодарственных писем студентам')
st.text('')
st.text('Из-за логики работы скрипта изначально необходимо загрузить файл со студентами')
st.text('и названием активности.')
st.text('Только после загрузки файла со студентами добавляйте файл с шаблоном, так как он')
st.text('требует переменные из файла со студентами.')
st.text('')

# students list upload button
file = st.file_uploader("Добавьте файл с названием активности и списком участников", type=["xls", "xlsx"])

if file is not None:
    # reading file
    df = pd.read_excel(file, header=0)
    # making variables
    students = df.student_name.tolist()
    project_name = df.project_name[0]
    start_date = df.start_date[0].strftime('%d.%m.%Y')
    finish_date = df.finish_date[0].strftime('%d.%m.%Y')

    # text of letter
    text_1 = f'Мастерская данных в лице Руководителя Мастерской данных Богданова Руслана Александровича благодарит вас за отличную работу над проектом «{project_name}» с {start_date} по {finish_date}.'
    text_2 = f'Отдельное спасибо за вашу вовлеченность, творческий подход к поиску оптимального решения и проявленный профессионализм.'

# text split
st.text('')

# letter scheme upload button
file_two = st.file_uploader("Добавьте файл с шаблоном благодарственного письма", type=["pdf"])

if file_two is not None:
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
        can.drawString(485, 64, nowaday)
        
        # printing main text
        text = can.beginText()
        text.setFont('YS Text Regular', 16)
        #text.setLeading(15)
        text.setTextOrigin(70, 360)
        text.textLines(wrap(text_1, 62))
        text.setTextOrigin(70, 250)
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
        existing_pdf = PdfReader(file_two, "rb")
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
