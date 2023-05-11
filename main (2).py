from openpyxl import load_workbook
from datetime import datetime
import PySimpleGUI as sg

layout = [[sg.Text('Поставщик'), sg.Push(), sg.Input(key='provider')],
          [sg.Text('Электронная почта'), sg.Push(), sg.Input(key='email')],
          [sg.Text('Модель'), sg.Push(), sg.Input(key='model')],
          [sg.Text('Производитель'), sg.Push(), sg.Input(key='manufacturer')],
          [sg.Text('Количество'), sg.Push(), sg.Input(key='quantity')],
          [sg.Text('Цена за единицу'), sg.Push(), sg.Input(key='price')],
          [sg.Button('Добавить'), sg.Button('Закрыть')]]
window = sg.Window('info', layout, element_justification='center')

while True:
    event, values = window.read()
    if event == sg.WINDOW_CLOSE_ATTEMPTED_EVENT or event == "Закрыть":
        break
    if event == 'Добавить':
        try:
            wb = load_workbook('data.xlsx')
            sheet = wb['Лист1']
            ID = len(sheet['ID'])
            time_stamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            data = [ID, values['provider'], values['email'], values['model'], values['manufacturer'], values['quantity'], values['price'], time_stamp]
            sheet.append(data)
            wb.save('data.xlsx')
            window['provider'].update(value='')
            window['email'].update(value='')
            window['model'].update(value='')
            window['manufacturer'].update(value='')
            window['quantity'].update(value='')
            window['price'].update(value='')
            window['provider'].set_focus()
            sg.popup('Данные сохранены')
        except PermissionError:
            sg.popup('File in use', 'File is being used by another User.\nPlease try again later.')
window.close()
