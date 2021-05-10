import io
import os
import pdfrw
import pandas as pd
from pathlib import Path
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

pdfmetrics.registerFont(TTFont('msjh', 'msjh.ttc'))


def run():
    Path("./output/").mkdir(parents=True, exist_ok=True)
    df = pd.read_excel('fill form.xlsx') # Edit path to link excel file
    pdfpath = '2019 Taiwan Tax Organizer - manual form2.pdf' # Edit path to link pdf file
    for index, row in df.iterrows():
        user_data = row.to_dict()
        print(f"{user_data['File Name']}.pdf\'s contant:")
        canvas_data = get_overlay_canvas(user_data, template_path=pdfpath)
        form = merge(canvas_data, template_path=pdfpath)
        save(form, filename=f"./output/{user_data['File Name']}.pdf")
        print('',end='\n\n')


def get_overlay_canvas(user_data: dict, template_path: str) -> io.BytesIO:
    template_pdf = pdfrw.PdfReader(template_path)
    data = io.BytesIO()
    pdf = canvas.Canvas(data)
    for index,page in enumerate(template_pdf.Root.Pages.Kids):
        print(f"  Page {index+1}")
        if page.Annots is None:
            continue
        for field in page.Annots:
            if not isinstance(field.T,str):
                print('    -------Field is missing-------')
                continue
            label = field.T[1:-1].split("_")[0].replace('\\','')
            sides_positions = field.Rect
            left = min(float(sides_positions[0]), float(sides_positions[2]))
            bottom = min(float(sides_positions[1]), float(sides_positions[3]))
            value = 'no'
            for i in user_data.keys():
                if label == i:
                    print(f"    {label}：\"{str(user_data[i])}\".")
                    value = str(user_data[i])
                    break
            if value in ['no','nan']:
                print(f"    {label}：Not found in dataset.")
                value = ''
            padding = 2
            line_height = 0
            pdf.setFont('msjh',10)
            pdf.drawString(x=left + padding, y=bottom + padding + line_height, text=value)
        pdf.showPage()
    pdf.save()
    data.seek(0)
    return data


def merge(overlay_canvas: io.BytesIO, template_path: str) -> io.BytesIO:
    template_pdf = pdfrw.PdfReader(template_path)
    overlay_pdf = pdfrw.PdfReader(overlay_canvas)
    for page, data in zip(template_pdf.pages, overlay_pdf.pages):
        overlay = pdfrw.PageMerge().add(data)[0]
        pdfrw.PageMerge(page).add(overlay).render()
    form = io.BytesIO()
    pdfrw.PdfWriter().write(form, template_pdf)
    form.seek(0)
    return form


def save(form: io.BytesIO, filename: str):
    with open(filename, 'wb') as f:
        f.write(form.read())

if __name__ == '__main__':
    run()