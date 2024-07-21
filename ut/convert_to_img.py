from win32com import client
import pythoncom
import os
import fitz  # PyMuPDF
from PIL import Image


def pdf_to_image(pdf_path, output_path, page_num=0, bbox=(0, 0, 1000, 1000), image_format="PNG"):
    try:
        # Open the PDF file
        pdf_document = fitz.open(pdf_path)

        # Get the specified page
        page = pdf_document[page_num]

        # Convert PDF coordinates to image coordinates
        x0, y0, x1, y1 = bbox
        x0, y0, x1, y1 = x0, page.rect.height - y1, x1, page.rect.height - y0

        # Select the desired area from the page as a pixmap
        pix = page.get_pixmap(matrix=fitz.Matrix(300 / 72, 300 / 72), clip=(x0, y0, x1, y1))

        # Convert the pixmap to a Pillow Image object
        image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

        # Save the image to the specified output path
        image.save(output_path, format=image_format)

        # Close the PDF file
        pdf_document.close()
    except Exception as e:
        print(repr(e))


def xlsx_to_img(file):
    try:
        script_dir = os.path.dirname(__file__)
        script_dir_1 = '\\'.join(script_dir.split('\\')[:-1])
        script_dir = '\\'.join(script_dir.split('\\')[:-1] + ['file\\img\\pdf'])

        rel_path = f'{file}.xlsx'
        rel_path2 = f'{file}.pdf'
        abs_file_path = os.path.join(script_dir_1, rel_path)
        abs_file_path2 = os.path.join(script_dir, rel_path2)
        pythoncom.CoInitialize()
        excel = client.DispatchEx("Excel.Application")
        sheets = excel.Workbooks.Open(abs_file_path, ReadOnly=True)
        work_sheets = sheets.Worksheets[0]
        work_sheets.PageSetup.Orientation = 2
        work_sheets.ExportAsFixedFormat(0, abs_file_path2)
        sheets.Close(False)
        excel.Quit()
        del excel
        pdf_file_path = f'file/img/pdf/{file}.pdf'

        for i in range(2):
            output_image_path = f'file/img/{file}-{str(i)}.png'  # Output image file path
            pdf_to_image(pdf_file_path, output_image_path, i)

    except Exception as e:
        print(repr(e))