import os
import shutil
import tempfile
import PyPDF2
from docx2pdf import convert
from reportlab.pdfgen import canvas
from openpyxl import Workbook
from pptx import Presentation
from pptx.util import Inches
from pptx.enum.dml import MSO_FILL_TYPE
from PIL import Image
import comtypes.client as client



def read_text_file(filename):
    with open(filename, 'r', encoding='utf-8') as file:
        text = file.read()
    return text



def convert_ppt_to_pdf(input_path, output_dir):
    # Получаем полный путь к выходной директории
    output_path = os.path.join(output_dir, os.path.splitext(os.path.basename(input_path))[0] + ".pdf")

    # Создаем объект PowerPoint
    powerpoint = client.CreateObject("Powerpoint.Application")

    # Открываем файл презентации
    presentation = powerpoint.Presentations.Open(input_path)

    # Сохраняем презентацию в формате PDF
    presentation.ExportAsFixedFormat(output_path, 32)

    # Закрываем презентацию и выходим из PowerPoint
    presentation.Close()
    powerpoint.Quit()


def convert_docx_to_pdf(input_path, output_path):
    # Конвертируем docx-файл в PDF
    convert(input_path, output_path)



def convert_to_pdf(input_dir=os.getcwd(), output_dir=os.getcwd()):
    # Создаем выходную директорию, если она еще не создана
    os.makedirs(os.path.join(output_dir, "OUTPUT"), exist_ok=True)
    output_dir = os.path.join(output_dir, "OUTPUT")
    with open(os.path.join(output_dir,"report.txt"), "w") as report:
        report.write(f"Input directory - {input_dir}\n Output directory - {output_dir}\n\n\n")
        # Проходим по всем файлам внутри исходной директории
        for filename in os.listdir(input_dir):
            input_path = os.path.join(input_dir, filename)
            output_path = os.path.join(output_dir, os.path.splitext(filename)[0] + '.pdf')

            # Проверяем расширение файла, чтобы не конвертировать PDF-файлы
            if filename.endswith('.pdf'):
                # Если файл уже является PDF-файлом, то копируем его в выходную директорию
                shutil.copy(input_path, output_path)
                report.write(f"{filename} is already in pdf format\n")
                continue

            if filename.endswith('.txt'):
                text = read_text_file(input_path)
                with open(output_path, 'w', encoding='utf-8') as pdf_file:
                    pdf_file.write(text)
                report.write(f"{filename} successfully converted to PDF\n")

            elif filename.endswith(('.jpg', '.jpeg', '.png', '.bmp', '.gif')):
                # Конвертируем изображение в PDF
                img = canvas.Canvas(output_path)
                img.drawImage(input_path, 0, 0)
                img.save()
                report.write(f"{filename} successfully converted to PDF\n")

            elif filename.endswith(('.ppt', '.pptx')):
                convert_ppt_to_pdf(input_path, output_dir)
                report.write(f"{filename} successfully converted to PDF\n")


            elif filename.endswith(('.xls', '.xlsx')):
                # Конвертируем XLS или XLSX файл в PDF
                wb = Workbook()
                ws = wb.active
                with open(input_path, 'r') as xls_file:
                    for row in xls_file:
                        row = row.strip().split('\t')
                        ws.append(row)
                wb.save(output_path)
                report.write(f"{filename} successfully converted to PDF\n")

            elif filename.endswith('.docx'):
                convert_docx_to_pdf(input_path, output_path)
                report.write(f"{filename} successfully converted to PDF\n")

            else:
                # print(f'Cannot convert {filename} to PDF')
                report.write(f'Cannot convert {filename} to PDF\n')

        # print('Конвертация файлов завершена!')
        report.write('Конвертация файлов завершена!')

def clear_path(path):
    path = path.replace('"', '')  # Remove double quotes
    path = path.replace("'", "")  # Remove single quotes
    return path

if __name__ == "__main__":
    # input_dir = clear_path(input("Enter full path for the input folder or press enter to skip: "))
    # if input_dir == "":
    #     input_dir = os.getcwd()
    # output_dir = clear_path(input("Enter full path for the output folder or press enter to skip: "))
    # if output_dir == "":
    #     output_dir = os.getcwd()   
    input_dir= "D:\Desktop\Remote\Author24\pythonconvectortopdf\pdfconv\data"
    output_dir ="D:\Desktop\Remote\Author24\pythonconvectortopdf\pdfconv"
    convert_to_pdf(input_dir=input_dir, output_dir=output_dir)
