from normalizer import normalize
from googletrans import Translator
import time

data = normalize('Text.txt')


def count_word_occurrences(word_list):
    word_count = {}
    for word in word_list:
        word_count[word] = word_count.get(word, 0) + 1
    sorted_word_count = sorted(word_count.items(), key=lambda item: item[1], reverse=True)
    return sorted_word_count


codata = count_word_occurrences(data)


def translate_words(sorted_list):
    translate_dictionary = []
    translator = Translator()
    for i, (word, count) in enumerate(sorted_list):
        try:
            translation = translator.translate(word, src='ru', dest='en').text
            translate_dictionary.append([word, translation, count])
            print(f'Переведено {i + 1}/{len(sorted_list)}...')
        except Exception as e:
            print(f'Произошла ошибка при переводе слова "{word}": {e}')

        # Добавим задержку после каждого 250 запроса
        if i % 250 == 0 and i != 0:
            print('Добавляю задержку...')
            time.sleep(5)  # Пауза в 5 секунд

    return translate_dictionary


translated_data = translate_words(codata)

# Сохранение результата в текстовый файл
output_file_path = 'translated_result.txt'
with open(output_file_path, 'w', encoding='utf-8') as output_file:
    output_file.write("Исходное слово | Перевод | Количество повторений\n")
    for item in translated_data:
        output_file.write(f"{item[0]} | {item[1]} | {item[2]}\n")

print(f"Результат сохранен в файл: {output_file_path}")
