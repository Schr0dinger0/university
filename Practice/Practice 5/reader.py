
def read_file(file_name):
    try:
        f = open (f'{file_name}', encoding="utf-8")
    except FileNotFoundError:
        print(f"Запрашиваемый файл {file_name} не найден")
    except OSError:
        print(f"Недопустимые символы в имени файла!")
    else:
        data = f.read().split()
        for i in range(len(data)):
            if data[i].isdigit() != True:
                data = 'Файл содержит не только числа!'
                return (print(data))
        for i in range(0,len(data)):
            data[i] = int(data[i])
        if int(data[0]) == 0:
            data = 'Число в первой строке == 0\n'
        elif int(data[0])+1 > len(data):
            data = 'Число в первой строке больше, чем чисел в файле!'
        elif int(data[0])+1 < len(data):
            print(f'Число в первой строке меньше, чем чисел в файле!\nВыведено только {data[0]} чисел(а)!')
            while len(data)-1 != data[0]:
                data.pop(-1)
        f.close()
        if str(data[0]).isdigit():
            data.pop(0)
        return(print(data))