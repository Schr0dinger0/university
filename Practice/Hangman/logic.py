import dictionary
import re
lives = 0
score = 0
used = []
choose = (re.sub("[^0-9а-я]", "", (input(f'Выберите уровень сложности:'
              f'\n 1. Легко  - 7 жизней'
              f'\n 2. Средне - 5 жизней'
              f'\n 3. Сложно - 3 жизни\n')).lower()))
game = True
while game == True:
    end = False
    match choose:
        case '1' | 'легко':
            lives = 7
        case '2' | 'средне':
            lives = 5
        case '3' | 'сложно':
            lives = 3
    used.clear()
    word = (dictionary.word())
    if word == 'exception':
        print(f'Похоже, что вы отгадали все слова! \n0_0\nСтолько очков вам удалось набрать: {score}')
        break
    print(f'Загаданное слово состоит из {len(word)} букв')
    answer = '*'*len(word)
    while lives != 0 and answer != word:
        print('\n\n'+answer)
        guess = re.sub("[^А-я]", "", (input('Введите одну букву или слово целиком: ')))
        if guess in(used) and len(guess)==1:
            print('Вы уже вводили эту букву!')
        elif guess in(used) and len(guess)!=1:
            print('Вы уже вводили это слово!')
        else:
            used.append(guess)
            if guess not in word and len(guess)==1:
                lives -= 1
                print(f'К сожалению, такой буквы в слове нет!\n\nЕще осталось {lives} жизней')
            elif guess not in word and len(guess) != 1:
                lives -= 1
                print(f'К сожалению, загадано не это слово!\n\nЕще осталось {lives} жизней')
            elif guess == word:
                print(f'Вы угадали слово {word} целиком!\n У вас было в запасе еще столько жизней: {lives}')
                answer = guess
            elif guess in word and len(guess)==1:
                print(f'Вы угадали букву!\nЖизней в запасе: {lives}')
                for i in range(0,len(word)):
                    if word[i] == guess:
                        answer1 = answer[:i]+guess+answer[i+1:]
                        answer = answer1
        if answer == word:
            score += 1
            end = True
        elif lives == 0:
            end = True
        if end == True:
            new_game = re.sub("[^0-9А-я]", "", (input(f'\n\nБыло загадано слово {word}\nСтолько очков вы набрали: {score}\nХотите сыграть еще?'
                                                       f'\n1.Да'
                                                       f'\n2.Нет\n')).lower())
            match new_game:
                case '1'|'да':
                    game = True
                case _:
                    game = False
if game == False:
    with open(f'best_score.txt','r') as f:
        record = int(f.read().replace('\n',''))
    if record < score:
        with open(f'best_score.txt', 'w') as f:
            f.write(str(score))