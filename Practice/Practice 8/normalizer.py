import re
import pymorphy2
morph = pymorphy2.MorphAnalyzer() #сокращаю функцию, потому что могу

def normalize(inp):
    new_data = []
    with open(f'{inp}', mode='r', encoding='utf-8') as f:
        data = f.read().split() #получаю список состоящий из отдельных слов
        for i in range(0,len(data)):
            data[i] = morph.parse(re.sub("[^A-zА-я ]", "",data[i]))[0] #оставляю только буквы двух языков, после чего нормализую и записываю туда, откуда взял
            new_data.append(data[i].normal_form) #формирую новый список из нормализованных слов
        for i in range(len(new_data)-1,-1,-1):
            if new_data[i] =='': #удаляю пустые строки
                del(new_data[i])
    return(new_data)