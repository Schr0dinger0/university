import random
import re

def birthday(people, repeats):

    answer = 0 #кол-во групп, в которых было повторение
    for i in range(0,repeats):
        unique_days = set()
        for j in range(0,people):
            unique_days.add(random.randint(1,365)) #все дни в году уникальны. Например, за один год может быть только одно 2 февраля,
                                                         # которое всегда идет в строго определенном порядке =>
                                                         # мы можем принять каждый месяц+день как одно число,
                                                         # обозначающее порядок этого дня в году, таким образом, 2 февраля является 33 днём в году
        if len(unique_days) != people:
            answer += 1
    return(f'\tкол-во групп, в которых были как минимум два человека с одинаковыми днями рождения:'
           f'\n{answer}\n\tпроцент групп, в которых были как минимум два человека с одинаковыми днями рождения:\n{format((100/(repeats/answer)), ".2f")}%')

if __name__ == '__main__':
    people = int(re.sub("[^0-9]", "", (input('Введите кол-во человек в группе: '))))
    repeats = int(re.sub("[^0-9]", "", (input('Введите кол-во итераций: '))))
    print(birthday(people,repeats))