from logic import *

if input('Введите 1, если хотите увидеть список книг и авторов ') == '1':
    print(reader())
print(get_books(input('Введите автора/название книги/цифровой код книги ')))
input('Нажмите Enter для продолжения ')
print(get_totals(spisok))
