import PyPDF2
import pymupdf
import docx2pdf
from pymupdf.utils import getColor
import os
import ntpath

watermark_file_path = "static_content_files/watermark.pdf"
def add_first_and_last_page(doc, first_and_last_page_path):
    first_and_last_page = pymupdf.open(first_and_last_page_path)
    doc.insert_pdf(first_and_last_page, from_page=0, to_page=0, start_at=0)
    doc.insert_pdf(first_and_last_page, from_page=1, to_page=1, start_at=-1)

def add_watermark(file_path, first_and_last_page_added = 0):
    if not os.path.exists(watermark_file_path):
        print(f"Warning: Watermark file {watermark_file_path} not found. Skipping watermark.")
        return
        
    input_pdf = PyPDF2.PdfReader(open(file_path, 'rb'))
    watermark = PyPDF2.PdfReader(open(watermark_file_path, 'rb'))
    output = PyPDF2.PdfWriter()

    start = 0 + first_and_last_page_added
    stop = len(input_pdf.pages) - first_and_last_page_added
    for i in range(len(input_pdf.pages)):
        page = input_pdf.pages[i]
        
        if (i >= start and i < stop):
            page.merge_page(watermark.pages[0])
        output.add_page(page)

    with open(file_path, 'wb') as file:
        output.write(file)

def add_link_and_page_number(input_pdf, link = "https://www.sarrthiias.com/"):
    for idx, page in enumerate(input_pdf):
        width, height = page.rect.width, page.rect.height
        img_rect = pymupdf.Rect(0, height - 50, width, height)
        page_no = f"{idx + 1}"
        margin = 48 + (0 if len(page_no) == 1 else (1 + len(page_no)))

        page.insert_text((width - margin, height - 20), page_no, fontsize=12, color = getColor('black'))
        page.insert_link({"kind": pymupdf.LINK_URI, "from": img_rect, "uri": link})

def convert_all_pdfs(file_paths: list[str] = None, folder_name: str = None):
    first_and_last_page_path = os.path.join(folder_name, "first_and_last_page.pdf")
    link = input("Enter link to be embedded in footer?\n>").strip()
    for file_path in file_paths:
        if not os.path.exists(file_path):
            print(f"Warning: PDF file not found: {file_path}. Skipping.")
            continue
        input_pdf = pymupdf.open(file_path)
        add_link_and_page_number(input_pdf, link)
        add_first_and_last_page(input_pdf, first_and_last_page_path)

        file_name = ntpath.basename(file_path)
        print(file_name)

        input_pdf.save(os.path.join(folder_name, "output", file_name.removeprefix("tmp_")))
        add_watermark(os.path.join(folder_name, "output", file_name.removeprefix("tmp_")), first_and_last_page_added = 1)

if __name__ == '__main__':
    folder_name = input("Enter folder name that contains header.png, footer.png and first_and_last_page.pdf\n")
    filepaths = []
    for file in os.listdir(os.path.join(folder_name, "output-docx")):
        if (file.endswith(".docx")):
            docx2pdf.convert(os.path.join(folder_name, "output-docx", file), os.path.join(folder_name, "output", f"tmp_{file.removesuffix(".docx")}.pdf"))
            filepaths.append(os.path.join(folder_name, "output", f"tmp_{file.removesuffix(".docx")}.pdf"))
    
    convert_all_pdfs(filepaths, folder_name)
    for file_path in filepaths:
        os.remove(file_path)
