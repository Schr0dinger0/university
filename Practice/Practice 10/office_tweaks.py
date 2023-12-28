import os
import PySimpleGUI as sg
from pdf2docx import Converter as PDFConverter
from docx2pdf import convert as docx_to_pdf_convert
from PIL import Image

def change_working_directory(new_directory):
    try:
        os.chdir(new_directory)
        return True
    except FileNotFoundError:
        sg.popup_error("Раздел не найден.")
    except OSError as e:
        sg.popup_error(f"Произошла ошибка при смене рабочего каталога: {e}")
    return False

def confirm_directory_change(new_directory):
    layout = [
        [sg.Text(f"Новый рабочий каталог: {new_directory}")],
        [sg.Button("Подтвердить"), sg.Button("Отмена")]
    ]

    window = sg.Window("Подтверждение смены каталога", layout, modal=True)

    while True:
        event, values = window.read()

        if event == sg.WINDOW_CLOSED or event == 'Отмена':
            window.close()
            return False

        if event == 'Подтвердить':
            window.close()
            return True

def pdf_to_docx(pdf_file):
    docx_file = os.path.splitext(pdf_file)[0] + ".docx"

    pdf_converter = PDFConverter(pdf_file)
    pdf_converter.convert(docx_file)
    pdf_converter.close()

    sg.popup(f"Преобразование завершено. Результат: {docx_file}")

def docx_to_pdf(docx_file):
    pdf_file = os.path.splitext(docx_file)[0] + ".pdf"

    docx_to_pdf_convert(docx_file, pdf_file)

    sg.popup(f"Преобразование завершено. Результат: {pdf_file}")

def compress_image(image_file, compression_percentage):
    image = Image.open(image_file)
    compressed_image_file = f"compressed_{compression_percentage}_{os.path.basename(image_file)}"

    # Сжимаем изображение с оптимизацией для формата JPEG
    compressed_image_file_path = os.path.join(os.getcwd(), compressed_image_file)
    image.save(compressed_image_file_path, optimize=True, quality=compression_percentage)

    sg.popup(f"Сжатие завершено. Результат: {compressed_image_file_path}")

def delete_files(files_to_delete):
    current_directory = os.getcwd()  # Получаем текущую директорию

    # Окно для подтверждения удаления файлов
    confirm_layout = [
        [sg.Text(f"Будут удалены следующие файлы:")],
        *[[sg.Text(file)] for file in files_to_delete],
        [sg.Text("Вы уверены, что хотите удалить эти файлы?")],
        [sg.Button('Да'), sg.Button('Нет')]
    ]

    confirm_window = sg.Window('Подтверждение удаления файлов', confirm_layout)

    event, _ = confirm_window.read()

    # Обработка событий окна подтверждения удаления файлов
    if event == 'Да':
        for file_to_delete in files_to_delete:
            try:
                os.remove(os.path.join(current_directory, file_to_delete))
            except Exception as e:
                sg.popup_error(f"Ошибка при удалении файла {file_to_delete}: {e}")

        sg.popup("Удаление завершено.")

    confirm_window.close()

def convert_selected_files(files, conversion_function):
    for file in files:
        try:
            conversion_function(file)
        except Exception as e:
            sg.popup_error(f"Ошибка при конвертации файла {file}: {e}")


def generate_tree_structure(path):
    folders = os.path.normpath(path).split(os.sep)
    structure = ""
    for folder in folders[:-1]:
        structure += f"{folder} {os.sep}\n"
    structure += folders[-1]
    return structure

def get_selected_option(values):
    # Поиск выбранного варианта
    for option in ['option1', 'option2', 'option3', 'option4']:
        if values[option]:
            return int(option[-1])  # Получение числа из ключа 'optionX'

def main():
    sg.theme('DarkAmber')
    tree_button_style = {'size': (30, 8), 'key': '-TREE-', 'enable_events': True, 'font': ('Helvetica', 12)}
    action_button_style = {'size': (30, 1), 'font': ('Helvetica', 12)}
    text_style = {'font': ('Helvetica', 12)}

    layout = [
        [sg.Button(generate_tree_structure(os.getcwd()), **tree_button_style)],
        [sg.Button("Преобразовать PDF в Docx", **action_button_style)],
        [sg.Button("Преобразовать Docx в PDF", **action_button_style)],
        [sg.Button("Произвести сжатие изображений", **action_button_style)],
        [sg.Button("Удалить группу файлов", **action_button_style)],
        [sg.Exit(size=(30, 1), font=('Helvetica', 12))]
    ]

    window = sg.Window("Office Tweaks", layout, resizable=False, size=(310, 390), font=('Helvetica', 12), finalize=True)
    window.set_min_size((310, 390))

    while True:
        event, values = window.read()

        if event == sg.WINDOW_CLOSED or event == 'Exit':
            break

        if event == '-TREE-':
            new_directory = sg.popup_get_folder("Выберите новый рабочий каталог:")
            if new_directory and change_working_directory(new_directory):
                if confirm_directory_change(new_directory):
                    window['-TREE-'].update(generate_tree_structure(new_directory))




        elif event == 'Преобразовать PDF в Docx':

            pdf_files = [f for f in os.listdir() if f.lower().endswith(".pdf")]

            layout_conversion = [

                [sg.Text("Выберите файлы для конвертации:")],

                *[[sg.Checkbox(file, key=file)] for file in pdf_files],

                [sg.Button("Конвертировать"), sg.Button("Отмена")]

            ]

            window_conversion = sg.Window("Выбор файлов для конвертации", layout_conversion)

            while True:

                event_conversion, values_conversion = window_conversion.read()

                if event_conversion in (sg.WIN_CLOSED, "Отмена"):

                    break

                elif event_conversion == "Конвертировать":

                    selected_files = [file for file in pdf_files if values_conversion[file]]

                    convert_selected_files(selected_files, pdf_to_docx)

            window_conversion.close()


        elif event == 'Преобразовать Docx в PDF':

            docx_files = [f for f in os.listdir() if f.lower().endswith(".docx")]

            layout_conversion = [

                [sg.Text("Выберите файлы для конвертации:")],

                *[[sg.Checkbox(file, key=file)] for file in docx_files],

                [sg.Button("Конвертировать"), sg.Button("Отмена")]

            ]

            window_conversion = sg.Window("Выбор файлов для конвертации", layout_conversion)

            while True:

                event_conversion, values_conversion = window_conversion.read()

                if event_conversion in (sg.WIN_CLOSED, "Отмена"):

                    break

                elif event_conversion == "Конвертировать":

                    selected_files = [file for file in docx_files if values_conversion[file]]

                    convert_selected_files(selected_files, docx_to_pdf)

            window_conversion.close()



        elif event == 'Произвести сжатие изображений':

            image_files = [f for f in os.listdir() if f.lower().endswith((".png", ".jpg", ".jpeg"))]

            layout_conversion = [

                [sg.Text("Выберите изображения для сжатия:")],

                *[[sg.Checkbox(file, key=file)] for file in image_files],

                [sg.Text("Процент сжатия:"),
                 sg.Slider(range=(1, 85), orientation="h", default_value=50, size=(20, 15), key="compression_slider")],

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
                    action = int([k for k, v in values_criteria.items() if v][0][-1])
                    substring = values_criteria['substring']
                    if not substring:
                        sg.popup_error("Неверный ввод. Пожалуйста, введите подстроку.")
                        continue

                    # Закрытие окна выбора метода удаления и ввода подстроки
                    window_criteria.close()

                    # Получение списка файлов для выбора
                    files = [f for f in os.listdir() if os.path.isfile(f)]

                    # Фильтрация файлов по критериям
                    filtered_files = []
                    for file in files:
                        if action == 1 and file.startswith(substring):
                            filtered_files.append(file)
                        elif action == 2 and file.endswith(substring):
                            filtered_files.append(file)
                        elif action == 3 and substring in file:
                            filtered_files.append(file)
                        elif action == 4 and file.endswith(substring):
                            filtered_files.append(file)

                    # Окно для выбора файлов для удаления
                    layout_files = [
                        [sg.Text(
                            f'Выберите файлы для удаления:')],
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