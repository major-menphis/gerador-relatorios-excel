import pandas as pd
import os
import docx2txt
from tqdm import tqdm
from time import sleep

#diretory define
def diretory_define():
    path = "./"
    os.chdir(path)

# detect CTA in text
def read_CTA(file):
    tag = 'CHAMADA:'
    text_read = docx2txt.process(file)
    if tag in text_read:
        return 'SIM'
    else:
        return 'NÃO'
    
#read files in diretory
def read_files():
    prices = {
        '30W': 0.90,
        '50W': 1.50,
        '100W': 3.00, 
        '500W': 15.00, 
        '1000W': 30.00,
        '2000W': 60.00,
        'CHAMADA': 0.90
    }
    text_price_total = 0
    call_price_total = 0
    price_total = 0
    texts_monthly = list()
    for file in tqdm(os.listdir()):
        writed_text = dict()
        if file.endswith('.docx'):
            writed_text['Palavras'] = file.split(maxsplit=1)[0]
            writed_text['Título'] = file.split(maxsplit=1)[1].replace('.docx', '')
            writed_text['Chamada'] = read_CTA(file)
            if writed_text['Palavras'] in prices:
                writed_text['Valor do texto'] = prices[file.split(maxsplit=1)[0]]
            text_price_total += writed_text['Valor do texto']
            if writed_text['Chamada'] == 'SIM':
                writed_text['Valor da chamada'] = prices['CHAMADA']
            else:
                writed_text['Valor da chamada'] = 0
            call_price_total += writed_text['Valor da chamada']
            writed_text['Total geral'] = (writed_text['Valor do texto'] + writed_text['Valor da chamada'])
            price_total += writed_text['Total geral']
            texts_monthly.append(writed_text)

        sleep(0.1)
    # result presenting
    texts_monthly.append({
        'Palavras': '', 
        'Título': 'VALORES TOTAIS GERAIS GERADOS ==>> ', 
        'Chamada': ' ==>> ',
        'Valor do texto': (text_price_total),
        'Valor da chamada': (call_price_total),
        'Total geral': (price_total)
    })

    return texts_monthly

#saving report
def save_report():
    print('Processando...')
    direc = os.getcwd()
    mes = direc.split("\\")[-1]
    diretory_define()
    list_data = read_files()
    if list_data != []:
        df = pd.DataFrame(list_data)
        writer = pd.ExcelWriter(f'{mes}.xlsx', engine='xlsxwriter')
        df.to_excel(writer, sheet_name=f'{mes}', header=True, index=False)
        workbook = writer.book
        worksheet = writer.sheets[mes]
        #add custom formats
        fmt_centered = workbook.add_format({
            'valign': 'vcenter',
            'align': 'center'
        })
        fmt_text = workbook.add_format({
            'text_wrap': True
        })
        fmt_currency = workbook.add_format({
            'num_format': 'R$ #,##0.00',
            'valign': 'vcenter',
            'align': 'center'
        })
        fmt_header = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'vcenter',
            'align': 'center',
            'fg_color': '#5DADE2',
            'font_color': '#FFFFFF',
            'border': 1
        })
        #set zoom
        worksheet.set_zoom(100)
        #apply custom formats to the cells
        worksheet.set_column("A:A", 10, fmt_centered)
        worksheet.set_column("B:B", 100, fmt_text)
        worksheet.set_column("C:C", 10, fmt_centered)
        worksheet.set_column("D:F", 10, fmt_currency)
        #apply format header
        for col, value in enumerate(df.columns.values):
            worksheet.write(0, col, value, fmt_header)
        #save
        writer.close()
        print('Arquivo gerado com sucesso!')
        sleep(2)
    else:
        return

if __name__ == "__main__":  
    save_report()
