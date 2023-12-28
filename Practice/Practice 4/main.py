import reader
import saver
data_name = input(f'Введите имя файла, который вы желаете отсортировать и сохранить ')
data = reader.read_file(data_name)
save_name = input(f'Введите имя файла, в который вы хотите сохранить данные ')
print(f'{saver.save_file(save_name, data)}')