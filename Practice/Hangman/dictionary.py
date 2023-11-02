import random

with open(f'list.txt',encoding="utf-8") as f:
   dictionary = f.read().split()
def word():
   if len(dictionary) != 0:
      choose = random.choice(dictionary)
      dictionary.remove(choose)
      return(choose)
   else:
      return('exception')
