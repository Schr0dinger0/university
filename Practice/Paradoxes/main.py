from MontyHall import *
from Birthdays import *

choose = input(f'Выберите парадокс\n1.Парадокс Монти Холла\n2.Парадокс Дней рождения\n')
choose = choose.lower()
match choose:
    case '1'|'монти холл'|'монти'|'холл':
        x = int(re.sub("[^0-9]", "", (input('Введите кол-во итераций: '))))
        print(montyhall(x))
    case '2' | 'дней рождения' | 'дней' | 'рождения':
        people = int(re.sub("[^0-9]", "", (input('Введите кол-во человек в группе: '))))
        repeats = int(re.sub("[^0-9]", "", (input('Введите кол-во итераций: '))))
        print(birthday(people,repeats))