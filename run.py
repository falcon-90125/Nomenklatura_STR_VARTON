import pandas as pd
import datetime

#pdf2txt
import io
from pdfminer.converter import TextConverter
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfpage import PDFPage

file_directory_resalts = f'exchange/Номенклатура_СТР_{datetime.date.today()}.xlsx' # Директория сохранения файла
varton_price = pd.read_excel('exchange/varton_price.xlsx') #загружаем прайс
pdf_path = 'exchange/nomenkl_str.pdf' #Путь к файлу номенклатуры СТР

# print(varton_price.head(5))

#Функция конвертации pdf2txt
def extract_text_from_pdf(pdf_path):
    resource_manager = PDFResourceManager()
    fake_file_handle = io.StringIO()
    converter = TextConverter(resource_manager, fake_file_handle)
    page_interpreter = PDFPageInterpreter(resource_manager, converter)

    with open(pdf_path, 'rb') as fh:
        for page in PDFPage.get_pages(fh, 
                                      caching=True,
                                      check_extractable=True):
            page_interpreter.process_page(page)
 
        text = fake_file_handle.getvalue()
    # close open handles
    converter.close()
    fake_file_handle.close()
    if text:
        return text

# Конвернтируем pdf в txt
nom_load = extract_text_from_pdf(pdf_path)
# print(nom_load)
# # Функция очистки данных от мусора
def append_vol(text_snippet):
# На стыке страниц есть лишняя инфа, которая попадает в кол-во, проверяем - если она есть - отбрасываем.
    if '(Место выхода света)' in text_snippet:
        text_snippet_if = text_snippet[text_snippet.index('(Место выхода света)')+20:] #отбрасываем
        # Так же может быть, что к кол-ву прилипает "Изображениесветильников дается вфирменном каталоге."
        # Тоже проверяем это и удаляем.
        if 'каталоге.' in text_snippet_if:
            vol_list.append(int(text_snippet_if[text_snippet_if.index('каталоге.')+9:text_snippet_if.index('VARTON')]))
        else:
            vol_list.append(int(text_snippet_if[:text_snippet_if.index('VARTON')])) # и добавляем кол-во в список с количествами
    else:
        # Так же может быть, что к кол-ву прилипает "Изображениесветильников дается вфирменном каталоге."
        # Тоже проверяем это и удаляем.
        if 'каталоге.' in text_snippet:
            vol_list.append(int(text_snippet[text_snippet.index('каталоге.')+9:text_snippet.index('VARTON')]))
        else:
            vol_list.append(int(text_snippet[:text_snippet.index('VARTON')])) # Отбираем кол-во из фрагмента, записываем в список с количествами

nom_load =  nom_load[nom_load.index('Место выхода света')+19:] # Отбрасываем шапку
nom_load =  nom_load[:nom_load.index('Общий световой поток')] # Отбрасываем всё что после номенклатуры (если весь СТР загружен)

vol_list = [] # Список с количествами
vendor_code_list = [] # Список с артикулами

# Проходим по тексту
for i in range(nom_load.count('CRI')): # Повторяем проход столько раз, сколько позиций в СТР
    text_snippet = nom_load[:nom_load.index('CRI')+6] # Отбираем фрагмент текста от кол-ва по CRI
    if 'VARTON' in text_snippet: # Проверяем позиция 'VARTON' или нет
        append_vol(text_snippet) # добавляем кол-во в список с количествами
        if '+' in text_snippet: # если двойная позиция добавляем кол-во в список с количествами ещё раз
            append_vol(text_snippet)
        if 'bracket 2 pieces' in text_snippet: # если если кронштейны для школьных досок, то умножаем на 2
            vol_list[-1] *= 2

        art_next = text_snippet[text_snippet.index('VARTON - ')+9:] # Отбрасываем 'VARTON - ' перед артикулом
        art_next = art_next[:art_next.index(' ')] # Отбрасываем весь текст после артикула и записываем артикул в список артикулов
        if 'Место' in art_next: # Если в артикул записалось 'Место', записываем артикул без него в список артикулов
            vendor_code_list.append(art_next[:art_next.index('Место')])
        elif 'Свето' in art_next: # Если в артикул записалось 'Светодиодный светильник...', записываем артикул без него в список артикулов
            vendor_code_list.append(art_next[:art_next.index('Свето')])
        elif 'BLACKBOARD' in art_next: # Если в артикул записалось 'BLACKBOARD', записываем артикул без него в список артикулов
            vendor_code_list.append(art_next[:art_next.index('BLACKBOARD')])

        else: # Если одиночная позиция
            vendor_code_list.append(art_next) # записываем артикул в список артикулов
        if '+' in text_snippet: # если двойная позиция добавляем кол-во в список с количествами ещё раз
            text_snippet_ = text_snippet[text_snippet.index('+')+2:] # Берём весь текст после "+"
            text_snippet_ = text_snippet_[:26]
            if 'Emergen' in text_snippet_: # Если в артикул записалось 'Exit', записываем артикул без него в список артикулов
                vendor_code_list.append(text_snippet_[:text_snippet_.index('Emergen')])
            elif 'М' in text_snippet_: # Если в артикул записалось 'Exit', записываем артикул без него в список артикулов
                vendor_code_list.append(text_snippet_[:text_snippet_.index('М')])
            elif ' ' in text_snippet_:
                vendor_code_list.append(text_snippet_[:text_snippet_.index(' ')]) # Отбрасываем весь текст после артикула и записываем артикул в список артикулов
            else:
                vendor_code_list.append(text_snippet_) # Отбрасываем весь текст после артикула и записываем артикул в список артикулов
        nom_load = nom_load[len(text_snippet):] # Отбрасываем весь обработанный текст, переходим к следующему фрагменту
    else:
        nom_load = nom_load[len(text_snippet):] # Если не 'VARTON', пропускаем

df = pd.DataFrame({'Артикул': vendor_code_list, 'Кол-во': vol_list}) # собираем df из списков с артикулами и кол-вом
df_new = df.merge(varton_price, on='Артикул', how='left') # подтягиваем данные из прайса

df_new['Сумма, Вход'] = round(df_new['Кол-во'] * df_new['Вход ЭКС'], 2) # произведение Кол-во и Вход ЭКС
df_new['Сумма, МРЦ'] = round(df_new['Кол-во'] * df_new['МРЦ'], 2) # произведение Кол-во и МРЦ
df_new['№'] = df_new.index + 1 # нумерация строк
df_new['Вход ЭКС'] = round(df_new['Вход ЭКС'], 2) # приводим к читабельному формату
df_new['МРЦ'] = round(df_new['МРЦ'], 2) # приводим к читабельному формату

columns=['№', 'Номенклатура', 'Артикул', 'Ед. изм.', 'Кол-во', 'Вход ЭКС', 'Сумма, Вход', 'МРЦ', 'Сумма, МРЦ']
df_reorder = df_new.reindex(columns=columns) # переставляем колонки в нужном порядке
sum_input = round(df_reorder['Сумма, Вход'].sum(), 2)
sum_mrc = round(df_reorder['Сумма, МРЦ'].sum(), 2)

#Добавление строки с итоговыми суммами по колонкам "вход" и "МРЦ"
df_sums_line = [['', '', '', '', '', '', sum_input, '', sum_mrc]]
df_sums = pd.DataFrame(df_sums_line, columns = columns)
df_concat = pd.concat([df_reorder, df_sums])
print(df_concat)

writer = pd.ExcelWriter(file_directory_resalts, engine='xlsxwriter') # + file_name
df_concat.to_excel(writer, sheet_name='Sheet1', index=False) # Определяем сохранение xlsx методом ExcelWriter
workbook = writer.book #записываем объект 'xlsxwriter' в книгу, для последующих назначений форматов
format1 = workbook.add_format({'num_format': '#,##0.00'}) # Формат для колоноц с ценами
format2 = workbook.add_format({'num_format': '#,##0.00', 'bold': True}) # Формат для колонок с суммами

sheet_0 = writer.sheets['Sheet1'] # Определяем лист для форматирования

sheet_0.set_column(0, 0, 5) # №
sheet_0.set_column(1, 1, 70) #наименование
sheet_0.set_column(2, 2, 28) #артикул
sheet_0.set_column(3, 3, 7) #Ед. изм.
sheet_0.set_column(4, 4, 7) #кол-во
sheet_0.set_column(5, 5, 14, format1) #цены вход
sheet_0.set_column(6, 6, 15, format2) #сумма вход
sheet_0.set_column(7, 7, 14, format1) #цены МРЦ
sheet_0.set_column(8, 8, 15, format2) #сумма вход

# writer.save() #Для прямого запуска
writer._save() #Для контейнера