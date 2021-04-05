import io
import base64
import openpyxl
from IPython.display import HTML
from pandas import ExcelFile, ExcelWriter


def create_download_link_excel(df, title = "Скачайте пример с адресами Excel file", filename = "example.xlsx"):  
    output = io.BytesIO()
    writer = ExcelWriter(output, engine='openpyxl')
    df.to_excel(writer, sheet_name='sheet1', index=False)
    writer.save()
    
    excel_data = output.getvalue()
    b64 = base64.b64encode(excel_data)
    payload = b64.decode()
    
    html = '<a download="{filename}" href="data:text/xml;base64,{payload}" target="_blank">{title}</a>'
    html = html.format(payload=payload,title=title,filename=filename)
    
    return HTML(html)


def start():
    
    with open("addr.test", "rb") as f:
        text = f.read()

    excel_data = ExcelFile(io.BytesIO(text), engine='openpyxl')
    test_frame = excel_data.parse(excel_data.sheet_names[0])

    return create_download_link_excel(test_frame)