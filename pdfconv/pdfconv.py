
import shutil
import os
import chardet
from reportlab.pdfgen import canvas

def convert_to_pdf(input_dir=os.getcwd(), output_dir=os.getcwd()):
    # Создаем выходную директорию, если она еще не создана
    os.makedirs(os.path.join(output_dir, "OUTPUT"), exist_ok=True)
    output_dir = os.path.join(output_dir, "OUTPUT")
    # Проходим по всем файлам внутри исходной директории
    for filename in os.listdir(input_dir):
        input_path = os.path.join(input_dir, filename)
        output_path = os.path.join(output_dir, os.path.splitext(filename)[0] + '.pdf')

        # Проверяем расширение файла, чтобы не конвертировать PDF-файлы
        if filename.endswith('.pdf'):
            # Если файл уже является PDF-файлом, то копируем его в выходную директорию
            shutil.copy(input_path, output_path)
            continue

        # Пробуем определить кодировку файла
        encodings = ['utf-8', 'latin-1', 'windows-1252']
        for encoding in encodings:
            try:
                with open(os.path.join(input_dir, filename), 'r', encoding=encoding) as f:
                    text = f.read()
                    break
            except UnicodeDecodeError:
                continue

        if encoding is None:
            print(f"Could not determine encoding for {filename}")
            continue


        # Конвертируем файл в PDF
        if filename.endswith(('.jpg', '.jpeg', '.png', '.bmp', '.gif')):
            img = canvas.Canvas(output_path)
            img.drawImage(input_path, 0, 0)
            img.save()
        elif filename.endswith(('.txt', '.doc', '.docx', '.ppt', '.pptx', '.xls', '.xlsx')):
            with open(input_path, 'rb') as input_file, \
                 open(output_path, 'wb') as output_file:
                rawdata = input_file.read()
                result = chardet.detect(rawdata)
                if result['encoding'] is not None:
                    encoding = result['encoding']
                    text = rawdata.decode(encoding)
                    c = canvas.Canvas(output_file)
                    c.drawString(100, 750, text)
                    c.save()
                else:
                    print(f"Could not determine encoding for {filename}")

        else:
            print(f'Cannot convert {filename} to PDF')

    print('Конвертация файлов завершена!')

def clear_path(path):
    path = path.replace('"', '')  # Remove double quotes
    path = path.replace("'", "")  # Remove single quotes
    return path

if __name__ == "__main__":
    input_dir = clear_path(input("Enter full path for the input folder or press enter to skip: "))
    if input_dir == "":
        input_dir = os.getcwd()
    output_dir = clear_path(input("Enter full path for the output folder or press enter to skip: "))
    if output_dir == "":
        output_dir = os.getcwd()    
    convert_to_pdf(input_dir=input_dir, output_dir=output_dir)




  