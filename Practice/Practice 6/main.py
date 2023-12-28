with open ('data.txt', encoding="utf-8") as f:
    print('File has been read')
    strings = f.read().splitlines()
    for i in range(len(strings)-1,-1,-1):
        if strings[i][0] != 'Р':
            del(strings[i])
    for i in range(0,len(strings)):
        strings[i] = strings[i].split()

    with open('new_data.txt', 'w+', encoding="utf-8") as f:
        for i in range(len(strings)):
            f.write(f'[{strings[i][6]}] - Поезд № {strings[i][1]} {strings[i][3]} {strings[i][4]}\n')
    print('Data Saved')