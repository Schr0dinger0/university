import os
import PySimpleGUI as sg
from pdf2docx import Converter as PDFConverter
from docx2pdf import convert as docx_to_pdf_convert
from PIL import Image

def change_working_directory(new_directory):
    try:
        os.chdir(new_directory)
        return True
    except (FileNotFoundError, OSError) as e:
        sg.popup_error(f"Ошибка при смене рабочего каталога: {e}")
        return False

def confirm_directory_change(new_directory):
    layout = [
        [sg.Text(f"Новый рабочий каталог: {new_directory}")],
        [sg.Button("Подтвердить"), sg.Button("Отмена")]
    ]

    event, _ = sg.Window("Подтверждение смены каталога", layout, modal=True).read()

    return event == 'Подтвердить'

def convert_files(files, conversion_function, input_extension, output_extension):
    layout_conversion = [
        [sg.Text(f"Выберите {input_extension} файлы для конвертации:")],
        *[[sg.Checkbox(file, key=file)] for file in files],
        [sg.Button("Конвертировать"), sg.Button("Отмена")]
    ]

    window_conversion = sg.Window(f"Выбор файлов для конвертации {input_extension}", layout_conversion)

    while True:
        event_conversion, values_conversion = window_conversion.read()

        if event_conversion in (sg.WIN_CLOSED, "Отмена"):
            break

        elif event_conversion == "Конвертировать":
            selected_files = [file for file in files if values_conversion[file]]
            convert_files(selected_files, conversion_function, input_extension, output_extension)

    window_conversion.close()

def main():
    sg.theme('Default1')

    layout = [
        [sg.Button(generate_tree_structure(os.getcwd()), size=(30, 8), key='-TREE-', enable_events=True)],
        [sg.Button("Преобразовать PDF в Docx", size=(30, 1))],
        [sg.Button("Преобразовать Docx в PDF", size=(30, 1))],
        [sg.Button("Произвести сжатие изображений", size=(30, 1))],
        [sg.Button("Удалить группу файлов", size=(30, 1))],
        [sg.Exit(size=(30, 1))]
    ]

    window = sg.Window("Office Tweaks", layout, resizable=False, size=(310, 390), font=('Helvetica', 12), finalize=True)
    window.set_min_size((310, 390))

    while True:
        event, values = window.read()

        if event in (sg.WINDOW_CLOSED, 'Exit'):
            break

        if event == '-TREE-':
            new_directory = sg.popup_get_folder("Выберите новый рабочий каталог:")
            if new_directory and change_working_directory(new_directory) and confirm_directory_change(new_directory):
                window['-TREE-'].update(generate_tree_structure(new_directory))

        elif event == 'Преобразовать PDF в Docx':
            pdf_files = get_files_by_extension(".pdf")
            convert_files(pdf_files, pdf_to_docx, "PDF", "Docx")

        elif event == 'Преобразовать Docx в PDF':
            docx_files = get_files_by_extension(".docx")
            convert_files(docx_files, docx_to_pdf, "Docx", "PDF")

        elif event == 'Произвести сжатие изображений':
            image_files = get_files_by_extension((".png", ".jpg", ".jpeg"))
            layout_conversion = [
                [sg.Text("Выберите изображения для сжатия:")],
                *[[sg.Checkbox(file, key=file)] for file in image_files],
                [sg.Text("Процент сжатия:"), sg.Slider(range=(1, 85), orientation="h", default_value=50, size=(20, 15),
                                                      key="compression_slider")],
                [sg.Button("Сжать изображения"), sg.Button("Отмена")]
            ]
            window_conversion = sg.Window("Выбор изображений для сжатия", layout_conversion)
            while True:
                event_conversion, values_conversion = window_conversion.read()
                if event_conversion in (sg.WIN_CLOSED, "Отмена"):
                    break
                elif event_conversion == "Сжать изображения":
                    selected_files = [file for file in image_files if values_conversion[file]]
                    compression_percentage = int(values_conversion["compression_slider"])
                    for selected_file in selected_files:
                        compress_image(selected_file, compression_percentage)
            window_conversion.close()

        elif event == 'Удалить группу файлов':
            # Окно для выбора метода удаления и ввода подстроки
            layout_criteria = [
                [sg.Text("Выберите критерий удаления:")],
                [sg.Radio('1. Удалить все файлы начинающиеся на определенную подстроку', group_id='group1',
                          key='option1')],
                [sg.Radio('2. Удалить все файлы заканчивающиеся на определенную подстроку', group_id='group1',
                          key='option2')],
                [sg.Radio('3. Удалить все файлы содержащие определенную подстроку', group_id='group1', key='option3')],
                [sg.Radio('4. Удалить все файлы по расширению', group_id='group1', key='option4')],
                [sg.Text("Введите подстроку:"), sg.InputText(key='substring')],
                [sg.Button('OK'), sg.Button('Отмена')]
            ]
            # Создание окна выбора метода удаления и ввода подстроки
            window_criteria = sg.Window('Выбор метода удаления и ввода подстроки', layout_criteria)
            action = None
            substring = None
            # Основной цикл обработки событий для окна выбора метода удаления и ввода подстроки
            while True:
                event_criteria, values_criteria = window_criteria.read()
                # Обработка событий окна выбора метода удаления и ввода подстроки
                if event_criteria in (sg.WINDOW_CLOSED, 'Отмена'):
                    break
                elif event_criteria == 'OK':
                    action = int([k for k, v in values_criteria.items() if v][0][-1])  # Получение числа из ключа 'optionX'
                    substring = values_criteria['substring']
                    if not substring:
                        sg.popup_error("Неверный ввод. Пожалуйста, введите подстроку.")
                        continue
                    # Закрытие окна выбора метода удаления и ввода подстроки
                    window_criteria.close()
                    # Получение списка файлов для выбора
                    files = get_all_files()
                    # Фильтрация файлов по критериям
                    filtered_files = filter_files(files, action, substring)
                    # Окно для выбора файлов для удаления
                    layout_files = [
                        [sg.Text(f'Выберите файлы для удаления:')],
                        *[[sg.Checkbox(file, key=file)] for file in filtered_files],
                        [sg.Button('Удалить выбранные файлы'), sg.Button('Отмена')]
                    ]
                    # Создание окна выбора файлов для удаления
                    window_files = sg.Window('Выбор файлов для удаления', layout_files)
                    # Основной цикл обработки событий для окна выбора файлов для удаления
                    while True:
                        event_files, values_files = window_files.read()
                        # Обработка событий окна выбора файлов для удаления
                        if event_files in (sg.WINDOW_CLOSED, 'Отмена'):
                            break
                        elif event_files == 'Удалить выбранные файлы':
                            selected_files = [file for file in filtered_files if values_files[file]]
                            delete_files(selected_files)
                            window_files.close()
                    # Закрытие окна выбора файлов для удаления
            window_criteria.close()
    window.close()

if __name__ == "__main__":
    main()
