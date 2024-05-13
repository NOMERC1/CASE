from openpyxl import load_workbook
import PySimpleGUI as sg
from datetime import datetime

element_size = (25, 1)

layout = [
    [sg.Text('ФИО пациента', size=element_size), sg.Input(key='patient_name', size=element_size)],
    [sg.Text('Дата рождения', size=element_size), sg.Input(key='birth_date', size=element_size)],
    [sg.Text('Пол', size=element_size), sg.Radio('Мужской', group_id='gender', key='male'),
     sg.Radio('Женский', group_id='gender', key='female')],
    [sg.Text('Номер телефона', size=element_size), sg.Input(key='phone_number', size=element_size)],
    [sg.Text('Номер страхового полиса', size=element_size), sg.Input(key='insurance_number', size=element_size)],
    [sg.Text('Номер медицинской карты', size=element_size), sg.Input(key='medical_card_number', size=element_size)],
    [sg.Text('Группа крови', size=element_size), sg.Input(key='blood_type', size=element_size)],
    [sg.Text('Специализация врача', size=element_size), sg.Input(key='doctor_specialization', size=element_size)],
    [sg.Text('Квалификация врача', size=element_size), sg.Input(key='doctor_qualification', size=element_size)],
    [sg.Text('Номер кабинета', size=element_size), sg.Input(key='office_number', size=element_size)],
    [sg.Text('Дата приема', size=element_size), sg.Input(key='appointment_date', size=element_size)],
    [sg.Text('Время приема', size=element_size), sg.Input(key='appointment_time', size=element_size)],
    [sg.Button('Добавить'), sg.Button('Закрыть')]
]

window = sg.Window('Учет посетителей поликлиники', layout, element_justification='center')

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Закрыть':
        break
    if event == 'Добавить':
        try:
            wb = load_workbook('clinic_visits.xlsx')
            sheet = wb['Лист1']
            ID = len(sheet['ID'])
            time_stamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

            gender = 'Мужской' if values['male'] else 'Женский'

            data = [
                ID,
                values['patient_name'],
                values['birth_date'],
                gender,
                values['phone_number'],
                values['insurance_number'],
                values['medical_card_number'],
                values['blood_type'],
                values['doctor_specialization'],
                values['doctor_qualification'],
                values['office_number'],
                values['appointment_date'],
                values['appointment_time'],
                time_stamp
            ]
            sheet.append(data)
            wb.save('clinic_visits.xlsx')

            for key in values:
                window[key].update(value='')
            window['patient_name'].set_focus()
            sg.popup('Данные сохранены')
        except PermissionError:
            sg.popup('Ошибка доступа', 'Файл используется другим пользователем.\nПопробуйте позже.')

window.close()