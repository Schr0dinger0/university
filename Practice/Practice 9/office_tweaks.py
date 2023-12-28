import os
from pdf2docx import Converter as PDFConverter
from docx2pdf import convert as docx_to_pdf_convert
from PIL import Image


def display_menu(menu_title, options):
    print(f"\n{menu_title}")
    for i, option in enumerate(options, start=1):
        print(f"{i}. {option}")


def list_files_by_extension(extension):
    return [f for f in os.listdir() if f.lower().endswith(extension)]


def change_working_directory():
    new_directory = input("Введите путь к новому рабочему каталогу: ")
    try:
        os.chdir(new_directory)
        print(f"Текущий рабочий каталог: {os.getcwd()}")
    except FileNotFoundError:
        print("Раздел не найден.")
    except OSError as e:
        print(f"Произошла ошибка при смене рабочего каталога: {e}")


def pdf_to_docx():
    pdf_files = list_files_by_extension(".pdf")
    print("Выберите PDF-файл для преобразования в Docx:")
    for i, pdf_file in enumerate(pdf_files, start=1):
        print(f"{i}. {pdf_file}")

    try:
        choice = int(input())
        if 1 <= choice <= len(pdf_files):
            pdf_file = pdf_files[choice - 1]

            docx_file = os.path.splitext(pdf_file)[0] + ".docx"

            pdf_converter = PDFConverter(pdf_file)
            pdf_converter.convert(docx_file)
            pdf_converter.close()

            print(f"Преобразование завершено. Результат: {docx_file}")
        else:
            print("Неверный ввод. Пожалуйста, выберите существующий файл.")
    except ValueError:
        print("Неверный ввод. Пожалуйста, введите число.")


def docx_to_pdf():
    docx_files = list_files_by_extension(".docx")
    print("Выберите Docx-файл для преобразования в PDF:")
    for i, docx_file in enumerate(docx_files, start=1):
        print(f"{i}. {docx_file}")

    try:
        choice = int(input())
        if 1 <= choice <= len(docx_files):
            docx_file = docx_files[choice - 1]

            pdf_file = os.path.splitext(docx_file)[0] + ".pdf"

            docx_to_pdf_convert(docx_file, pdf_file)

            print(f"Преобразование завершено. Результат: {pdf_file}")
        else:
            print("Неверный ввод. Пожалуйста, выберите существующий файл.")
    except ValueError:
        print("Неверный ввод. Пожалуйста, введите число.")


def compress_images():
    image_files = list_files_by_extension((".png", ".jpg", ".jpeg"))
    print("Выберите изображение для сжатия:")
    for i, image_file in enumerate(image_files, start=1):
        print(f"{i}. {image_file}")

    try:
        choice = int(input())
        if 1 <= choice <= len(image_files):
            image_file = image_files[choice - 1]

            print("Введите процент сжатия (от 1 до 85):")
            compression_percentage = int(input())

            if not (1 <= compression_percentage <= 85):
                print("Неверный ввод. Процент сжатия должен быть от 1 до 85.")
                return

            image = Image.open(image_file)
            compressed_image_file = f"compressed_{compression_percentage}_{image_file}"

            # Сжимаем изображение с оптимизацией для формата JPEG
            image.save(compressed_image_file, optimize=True, quality=compression_percentage)

            print(f"Сжатие завершено. Результат: {compressed_image_file}")
        else:
            print("Неверный ввод. Пожалуйста, выберите существующий файл.")
    except ValueError:
        print("Неверный ввод. Пожалуйста, введите число.")


def delete_files():
    print("Выберите критерий удаления:")
    print("1. Удалить все файлы начинающиеся на определенную подстроку")
    print("2. Удалить все файлы заканчивающиеся на определенную подстроку")
    print("3. Удалить все файлы содержащие определенную подстроку")
    print("4. Удалить все файлы по расширению")

    try:
        choice = int(input())
        if 1 <= choice <= 4:
            substring = input("Введите подстроку: ")

            if choice == 1:
                files_to_delete = [f for f in os.listdir() if f.startswith(substring)]
            elif choice == 2:
                files_to_delete = [f for f in os.listdir() if f.endswith(substring)]
            elif choice == 3:
                files_to_delete = [f for f in os.listdir() if substring in f]
            elif choice == 4:
                files_to_delete = [f for f in os.listdir() if f.endswith(substring)]

            print("Будут удалены следующие файлы:")
            for file_name in files_to_delete:
                print(file_name)

            confirm_delete = input("Вы уверены, что хотите удалить эти файлы? (y/n): ")

            if confirm_delete.lower() == "y":
                for file_to_delete in files_to_delete:
                    os.remove(file_to_delete)
                print("Удаление завершено.")
            else:
                print("Удаление отменено.")
        else:
            print("Неверный ввод. Пожалуйста, выберите существующий критерий.")
    except ValueError:
        print("Неверный ввод. Пожалуйста, введите число.")


def main():
    while True:
        display_menu("Главное меню", ["Сменить рабочий каталог", "Преобразовать PDF в Docx",
                                      "Преобразовать Docx в PDF", "Произвести сжатие изображений",
                                      "Удалить группу файлов", "Выход"])
        try:
            choice = int(input())
        except ValueError:
            print("Неверный ввод. Пожалуйста, введите число.")
            continue

        if choice == 1:
            change_working_directory()
        elif choice == 2:
            pdf_to_docx()
        elif choice == 3:
            docx_to_pdf()
        elif choice == 4:
            compress_images()
        elif choice == 5:
            delete_files()
        elif choice == 6:
            print("Программа завершена.")
            break
        else:
            print("Неверный выбор. Пожалуйста, выберите действие снова.")


if __name__ == "__main__":
    main()
