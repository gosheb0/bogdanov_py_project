import os

print(" Выберите, что запустить: ")
print("1 - Формирование отчета в виде exel файла ")
print("2 - Формирование графиков ")
result = input()
if int(result) == 1:
    os.system("forming_excel.py")
if int(result) == 2:
    os.system("forming_graph.py")
