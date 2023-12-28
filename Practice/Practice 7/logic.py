
with open ('books.csv', mode='r', encoding='utf-8') as f:
    spisok = f.read().splitlines()
    for i in range(0,len(spisok)):
        spisok[i]=spisok[i].split('|')

def reader():
    print('\n')
    for i in range(1, len(spisok)):
        print("%45s  %12s  %-24s | %-20s" % (spisok[i][1], '- by -', spisok[i][2], (spisok[i][0])))
    return''
def get_totals(spisok):
    lst = []
    for i in range(1, len(spisok)):
        if int(spisok[i][3])*float(spisok[i][4]) < 500:
            tpl = (spisok[i][0], int(spisok[i][3]) * float(spisok[i][4])+100)
        else:
            tpl = (spisok[i][0],int(spisok[i][3])*float(spisok[i][4]))
        lst.append(tpl)
    for i in range(0,len(lst)):
        print(lst[i])
    return''
def get_books(wish):
    print('\n')
    cnt = 0
    for i in range(1,int(len(spisok))-1):
        if str(wish).lower() in str(spisok[i]).lower():
            print("%-20s %-45s %-30s %-10s %-10s" % (spisok[i][0],spisok[i][1],spisok[i][2],spisok[i][3],spisok[i][4],))
            cnt += 1
    if cnt == 0:
        return 'Not Found!'
    else:
        return''