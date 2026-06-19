# Выборка номеров телефонов с портала и сохранение в Excel
import pandas as pd 
from pathlib import Path
file1 = Path("C:/") / "Users" / "KVG" / "Documents" / "Python" / "Parser phone" / "Страницы - Телефонный справочник МЭЦ.html"
file2 = Path("C:/") / "Users" / "KVG" / "Documents" / "Python" / "Parser phone" / "Телефонный справочник.txt"
a = 0
attempts = 0
position = ""
title=[]
pos=[]
dep=[]
phonecity=[]
phonevts=[]

data={'ФИО':title,
      'Должность': pos,
      'Подразделение': dep,
      'ВТС': phonevts,
      'Городской': phonecity
     }

with open(file1, 'r', encoding='utf-8') as f, \
     open(file2, 'w', encoding='utf-8') as f_out:
   while True:
      line = f.readline()
      if not line:
         break
      

      if line.count ("user-title") == 1:
        
         line = f.readline() #переход на следующую строку
         #выделение ФИО
         start = line.find('>') + 1
         end = line.find("<", start)
         fio = line[start:end] if start > 0 and end >0 else ''
         title.append(fio) #добавить ФИО в список
         # выделение должности "user-position"
         while line.count("user-position") != 1:
            line = f.readline()
            if line.count("user-position") ==1:
            # line = f.readline() #переход на следующую строку
               start = line.find('>') + 1
               end = line.find("<", start)
               position = line[start:end] if start > 0 and end > 0 else ''
               pos.append(position)
               break
         while line.count("user-department") != 1:
             line = f.readline()
             if line.count("user-department") ==1:
               start = line.find('>') + 1
               end = line.find("<", start)
               department = line[start:end] if start > 0 and end > 0 else ''
               dep.append(department)
               break
         while line.count("Телефоны") != 1:
             line = f.readline()          
             if line.count("Телефоны") == 1: 
               line = f.readline()
               saved_pos = f.tell() 
               while line.count("Городской") != 1 and attempts < 6:
                  line = f.readline()
                  attempts += 1
                  #phc = "    "
                  #phonecity.append(phc)
                  if line.count("Городской") ==1:
                     line = f.readline()
                     start = line.find(':') + 1
                     end = line.find(">", start) - 1
                     phc = line[start:end] if start > 0 and end > 0 else ''
                     phonecity.append(phc)
                     break 
                  if attempts == 6:    
                     f.seek(saved_pos)
                     attempts = 0
                     phonecity.append("нет")
                     break
               while line.count("Внутренний") != 1:
                  line = f.readline()
                  if line.count("Внутренний") ==1:
                     line = f.readline()
                     start = line.find(':') + 1
                     end = line.find(">", start) - 1
                     phv = line[start:end] if start > 0 and end > 0 else ''
                     phonevts.append(phv)
                     break   
               if attempts != 0:
                  attempts = 0
               break        
         #print(line, end='')
         print (f'{department} | {position:31} | {fio:30} | {phc:12} | {phv:12}')
         #print (position)

         # запись в файл
         f_out.write(department + '\n')
         f_out.write(f'{position:31} | {fio:30}\n')
         #f_out.write(fio + '\n')
         
         a+=1
print (f"Всего найдено записей: ", a) #счетчик записей
#print()
#print (title)
#print(pos)
#print(dep)
for key, value in data.items(): # Проверка количества записей в каждом
   print(f"Ключ '{key}': {len(value)} записей")

df = pd.DataFrame(data)        
df.to_excel(r'C:\Users\KVG\Documents\Python\Parser phone\output.xlsx', index=False)
print("Data exported to output.xlsx successfully.")
         
      