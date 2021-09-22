import openpyxl
import random

workbook = openpyxl.load_workbook("a.xlsx")

worksheet = workbook["Sheet1"]
print(worksheet)

random_score = [random.randint(88, 92) for _ in range(33)]
print(random_score)

for i in range(len(random_score)):
    worksheet.cell(i + 2, 3, random_score[i])

workbook.save('德育学生打分表.xlsx')
