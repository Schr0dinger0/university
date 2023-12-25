import os
import PySimpleGUI as sg
from pdf2docx import Converter
from docx2pdf import convert
from PIL import Image

def main_window():
    layout = [
        [sg.Text("Добро пожаловать в Office Tweaks!", font=("Helvetica", 16))],
        [sg.Button("Сменить рабочий каталог", size=(30, 2), font=("Helvetica", 12))],
        [sg.Button("Преобразовать PDF в Docx", size=(30, 2), font=("Helvetica", 12))],
        [sg.Button("Преобразовать Docx в PDF", size=(30, 2), font=("Helvetica", 12))],
        [sg.Button("Произвести сжатие изображений", size=(30, 2), font=("Helvetica", 12))],
        [sg.Button("Удалить группу файлов", size=(30, 2), font=("Helvetica", 12))],
        [sg.Button("Выход", size=(30, 2), font=("Helvetica", 12))]
    ]

    return sg.Window("Office Tweaks", layout, element_justification="center", margins=(20, 20))

def change_working_directory():
    layout = [
        [sg.Text("Введите путь к новому рабочему каталогу:")],
        [sg.InputText(key="directory_input", enable_events=True, default_text=os.getcwd()), sg.FolderBrowse()],
        [sg.Button("OK"), sg.Button("Отмена")]
    ]

    window = sg.Window("Сменить рабочий каталог", layout)

    while True:
        event, values = window.read()

        if event == sg.WIN_CLOSED or event == "Отмена":
            new_directory = None
            break
        elif event == "OK":
            new_directory = values["directory_input"]
            break

    window.close()
    return new_directory

def pdf_to_docx(pdf_path):
    pdf_path_str = pdf_path.name if isinstance(pdf_path, type(os.path)) else pdf_path
    docx_path = os.path.splitext(pdf_path_str)[0] + ".docx"

    cv = Converter(pdf_path_str)
    cv.convert(docx_path, start=0, end=None)
    cv.close()

    return docx_path

def docx_to_pdf(docx_path):
    pdf_path = os.path.splitext(docx_path)[0] + ".pdf"
    convert(docx_path, pdf_path)
    return pdf_path

def compress_images(directory, quality, selected_images):
    image_types = [".png", ".jpg", ".jpeg"]

    for root, dirs, files in os.walk(directory):
        for file in files:
            if any(file.lower().endswith(img_type) for img_type in image_types) and (not selected_images or file in selected_images):
                image_path = os.path.join(root, file)
                compressed_image_path = os.path.splitext(image_path)[0] + "_compressed.jpg"

                with Image.open(image_path) as img:
                    img.save(compressed_image_path, "JPEG", quality=quality)

    sg.popup(f"Сжатие завершено. Сжатые изображения сохранены в {directory}")


def delete_files(directory, filter_option=None, substring=None):
    deleted_files = []

    try:
        # Получаем все файлы в директории
        all_files = [os.path.join(root, file) for root, dirs, files in os.walk(directory) for file in files]


        # Используем словарь для соответствия методу удаления и соответствующей операции
        filter_options = {
            '1': lambda x: os.path.basename(x).startswith(substring),
            '2': lambda x: os.path.basename(x).endswith(substring),
            '3': lambda x: substring.lower() in os.path.basename(x).lower(),
            '4': lambda x: os.path.splitext(x)[1].lower() in {".png", ".jpg", ".jpeg"}
        }


        if all_files:

            # Фильтруем файлы согласно выбранным требованиям
            filtered_files = [file_path for file_path in all_files if filter_options.get(filter_option, lambda x: False)(file_path)]
            for file_path in all_files:
                comparison_result = filter_options.get(filter_option, lambda x: False)(file_path)
                print(
                    f"Проверка файла: {file_path}, Результат: {comparison_result}, Сравнение: {substring.lower()} in {os.path.basename(file_path).lower()} or {substring.lower()} in {os.path.splitext(os.path.basename(file_path))[0].lower()} or {substring.lower()} in {os.path.splitext(os.path.basename(file_path))[1].lower()}")

            if filtered_files:
                print(f"Выбранный метод: {filter_option}")
                print(f"Выбранная подстрока: {substring}")
                print(f"Файлы, удовлетворяющие критериям: {filtered_files}")

                for file_path in filtered_files:
                    try:
                        if os.path.isfile(file_path):
                            os.remove(file_path)
                        elif os.path.isdir(file_path):
                            os.rmdir(file_path)
                        deleted_files.append(file_path)
                    except Exception as e:
                        sg.popup_error(f"Ошибка при удалении файла {file_path}: {e}")

                sg.popup(f"Удаление завершено. Удалены следующие файлы:\n\n{', '.join(deleted_files)}")
            else:
                sg.popup("Нет файлов, удовлетворяющих выбранным критериям.")
        else:
            sg.popup("В выбранной директории нет файлов для удаления.")
    except Exception as e:
        sg.popup_error(f"Ошибка: {e}")


def main():
    window = main_window()
    current_directory = os.getcwd()  # Установим рабочий каталог по умолчанию

    while True:
        event, values = window.read()

        if event == sg.WIN_CLOSED or event == "Выход":
            break
        elif event == "Сменить рабочий каталог":
            new_directory = change_working_directory()

            if new_directory:
                current_directory = new_directory
                sg.popup(f"Рабочий каталог изменен на {current_directory}")

        elif event == "Преобразовать PDF в Docx":
            pdf_path = sg.popup_get_file("Выберите PDF-файл для преобразования в Docx:", no_window=True)

            if pdf_path:
                docx_path = pdf_to_docx(pdf_path)
                sg.popup(f"Преобразование завершено. Результат сохранен в {docx_path}")
            else:
                sg.popup("Выбор отменен.")

        elif event == "Преобразовать Docx в PDF":
            docx_path = sg.popup_get_file("Выберите Docx-файл для преобразования в PDF:", no_window=True)

            if docx_path:
                pdf_path = docx_to_pdf(docx_path)
                sg.popup(f"Преобразование завершено. Результат сохранен в {pdf_path}")
            else:
                sg.popup("Выбор отменен.")

        elif event == "Произвести сжатие изображений":
            if current_directory:
                layout_compression = [
                    [sg.Text("Выберите изображения для сжатия:")],
                    [sg.Listbox(values=[file for file in os.listdir(current_directory) if file.lower().endswith((".png", ".jpg", ".jpeg"))], key="selected_images", select_mode=sg.LISTBOX_SELECT_MODE_EXTENDED, size=(40, 10))],
                    [sg.Text("Выберите процент сжатия (1-85%):")],
                    [sg.Slider(range=(1, 85), orientation="h", size=(30, 15), default_value=85, key="compression_slider")],
                    [sg.Button("OK"), sg.Button("Отмена")]
                ]

                window_compression = sg.Window("Сжатие изображений", layout_compression)

                while True:
                    event_compression, values_compression = window_compression.read()

                    if event_compression == sg.WIN_CLOSED or event_compression == "Отмена":
                        quality = None
                        selected_images = None
                        break
                    elif event_compression == "OK":
                        quality = int(values_compression["compression_slider"])
                        selected_images = values_compression["selected_images"]
                        break

                window_compression.close()

                if quality is not None:
                    compress_images(current_directory, quality, selected_images)
            else:
                sg.popup("Выберите рабочий каталог перед сжатием изображений.")

        elif event == "Удалить группу файлов":
            if current_directory:
                filter_option = None

                layout_filter = [
                    [sg.Text("Выберите метод удаления:")],
                    [sg.Radio("Начинаются на подстроку", "deletion_method", default=True, key="deletion_method_startswith")],
                    [sg.Radio("Заканчиваются на подстроку", "deletion_method", key="deletion_method_endswith")],
                    [sg.Radio("Содержат подстроку", "deletion_method", key="deletion_method_contains")],
                    [sg.Radio("По расширению", "deletion_method", key="deletion_method_by_extension")],
                    [sg.Button("Далее"), sg.Button("Отмена")]
                ]

                window_filter = sg.Window("Выбор метода удаления", layout_filter)

                while True:
                    event_filter, values_filter = window_filter.read()

                    if event_filter == sg.WIN_CLOSED or event_filter == "Отмена":
                        break
                    elif event_filter == "Далее":
                        filter_option = next(key for key, value in values_filter.items() if value) if any(values_filter.values()) else None
                        break

                window_filter.close()

                if filter_option is not None:
                    layout_substring = [
                        [sg.Text("Введите подстроку для фильтра:")],
                        [sg.InputText(key="deletion_substring")],
                        [sg.Button("OK"), sg.Button("Отмена")]
                    ]

                    window_substring = sg.Window("Ввод подстроки для фильтра", layout_substring)

                    while True:
                        event_substring, values_substring = window_substring.read()

                        if event_substring == sg.WIN_CLOSED or event_substring == "Отмена":
                            break
                        elif event_substring == "OK":
                            substring = values_substring["deletion_substring"]
                            break

                    window_substring.close()

                    if substring is not None:
                        delete_files(current_directory, filter_option, substring)
            else:
                sg.popup("Выберите рабочий каталог перед удалением группы файлов.")

    window.close()

if __name__ == "__main__":
    main()
