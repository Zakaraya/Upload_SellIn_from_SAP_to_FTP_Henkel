import tkinter as tk
from tkinter import filedialog
import pandas as pd
import openpyxl
from openpyxl.styles import numbers


# Функция прикрепления файла SellIn из SAP
def attach_file1():
    file_path = filedialog.askopenfilename()
    file1_label.config(text=file_path)


# Функция прикрепления файла, в который необходимо поместить данные из SAP
def attach_file2():
    file_path = filedialog.askopenfilename()
    file2_label.config(text=file_path)


# Функция преобразования файла
def process_files():
    file1_path = file1_label["text"]
    file2_path = file2_label["text"]
    # чтение файла из SAP
    ke24_SAP_df = pd.read_excel(file1_path, 'Sheet1', index_col=False)
    ke24_SAP_df.dropna(subset=['Product'], inplace=True)
    ke24_SAP_df['Product'] = ke24_SAP_df['Product'].astype('int')
    ke24_SAP_df['Product'] = ke24_SAP_df['Product'].astype(str)

    columns = [
        'Posting Date',
        'Ship To Country Customer',
        'Company Code Key',
        'Company Code',
        'Ship To Customer Number Global',
        'Ship To Customer Global',
        'Material Number',
        'Material',
        'Qty in CON',
        'CPV',
        'NES'
    ]
    # Создание результирующего датафрейма
    new_salesbi_file = pd.DataFrame(columns=columns, data={
        'Posting Date': ke24_SAP_df['Posting date'].dt.strftime("%d.%m.%Y"),
        'Ship To Country Customer': ke24_SAP_df['Country Text'],
        'Company Code Key': '7558',
        'Company Code': 'Henkel Rus LLC',
        'Ship To Customer Number Global': ke24_SAP_df['Ship-To Party'],
        'Ship To Customer Global': ke24_SAP_df['Ship-To Party Text'],
        'Material Number': ke24_SAP_df['Product'].astype('int'),
        'Material': ke24_SAP_df['Product Text'],
        'Qty in CON': ke24_SAP_df['Quantity in CON'],
        'CPV': round(ke24_SAP_df['CPV'].astype('float'), 2),
        'NES': round(ke24_SAP_df['NES'].astype('float'), 2),
    })
    # удаление пустых строк
    new_salesbi_file = new_salesbi_file.dropna(thresh=6)
    # удаляем продукт, который не относится к нашим продажам - Freight & Warehousing external Costs
    new_salesbi_file = new_salesbi_file[
        (new_salesbi_file['CPV'] != 0) & (new_salesbi_file['Material Number'] != '851810') & (
                    new_salesbi_file['Material'] != 'Freight & Warehousing external Costs')]
    # Оставляем только Россию и Беларусь. В выгрузке могут быть Казахстан, Армения, Азербайджан, Таджикистан
    new_salesbi_file = new_salesbi_file[
        new_salesbi_file['Ship To Country Customer'].isin(['Russian Feder.', 'Belarus'])]

    # Перезаписываем получившийся файл, сохраняя вид отчета
    book = openpyxl.load_workbook(file2_path)
    writer = pd.ExcelWriter(file2_path, engine='openpyxl')
    writer.book = book
    writer.sheets.update(dict((ws.title, ws) for ws in book.worksheets))

    sheet = writer.sheets[str(book.sheetnames[0])]

    new_salesbi_file.to_excel(writer, 'Grid', startrow=3, startcol=1, index=False, header=0)
    # writer.save()
    # for cell in sheet[f'F4:F5']:
    #     # print(cell)
    #     cell[0]= str(cell[0].value)
    writer.close()

    result_label.config(text="Files processed successfully!")


root = tk.Tk()
root.geometry("400x200")

file1_button = tk.Button(text="File from SAP", command=attach_file1)
file1_button.pack()

file1_label = tk.Label(text="")
file1_label.pack()

file2_button = tk.Button(text="File to FTP", command=attach_file2)
file2_button.pack()

file2_label = tk.Label(text="")
file2_label.pack()

process_button = tk.Button(text="Process files", command=process_files)
process_button.pack()

result_label = tk.Label(text="")
result_label.pack()

root.mainloop()
