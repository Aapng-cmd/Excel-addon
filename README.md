# Excel-addon

import xlrd, xlwt
from random import randint as ri
import time

a = xlwt.Font()
a.name = 'Times New Roman'
a.colour_index = 2
a.bold = True
s0 = xlwt.XFStyle()
s0.font = a
wb = xlwt.Workbook()
ws = wb.add_sheet('A Test')

a = int(input("Впишите кол-во классов: "))
b = int(input("Впишите кол-во задач: "))
p = 0
o = 0
qu = 0
pop = []
qu_ = []
pop1 = []
u = 0
u_ = []
s = []

# ws.write(0, b + 8, "Итоговая сумма")
# print("Вводите баллы классов по порялку задач.")

for i in range(a + 1):
    for j in range(b + 1):
        if j == 0 and i != 0:
            ws.write(i, 0, "Класс {}".format(i + 7))
        elif i == 0 and j != 0:
            ws.write(0, j, "Задача {}".format(j))
        elif i != 0 and j != 0:
            # print("Класс № {}, Задание № {}".format(i, j))
            p = ri(0, 10)
            # p = int(input())

            '''try:
                p = int(input())
            except ValueError:
                print("Не число, введите ещё раз.")
                j -= 1
            '''
            ws.write(i, j, p)
            o += p

            pop.append(p)

"""    if i != 0:
        ws.write(i, j + 2, o)
        pop1.append(o)
    o = 0
ws.write(0, j + 2, "Сумма баллов")
for z in range(b):
    for q in range(len(pop)):
        if q % b == z:
            u += pop[q]
    ws.write(a + 2, z + 1, u)
    if u == 0:
        u = 1
    u_.append(u)
    u = 0
for y in range(a):
    for _ in range(len(u_)):
        qu += pop[_ + b * y] / u_[_]
    qu *= 10
    qu_.append(qu)
    qu = 0
for _1 in range(len(pop1)):
    s.append((pop1[_1] + qu_[_1]))
for plague in range(1, len(s) + 1):
    ws.write(plague, b + 8, s[plague - 1])
#ws.write(0, 0, "Test", s0)
"""
wb.save('example.xls')
name = input("Выберите файл, который надо обработать: ")
count = 0
f = open('example.xls')
while True:

    a = xlwt.Font()
    a.name = 'Times New Roman'
    a.colour_index = 2
    a.bold = True
    s0 = xlwt.XFStyle()
    s0.font = a
    wb = xlwt.Workbook()
    ws = wb.add_sheet('An answer')

    rb = xlrd.open_workbook(name + ".xls")
    sheet = rb.sheet_by_index(0)

    a = 0
    b = 0
    maximum_bal = []
    bal = []
    maximum_class = []
    just_class = []
    just = []
    # val = sheet.row_values(0)[1]
    val = sheet.row_len(0) - 1
    val1 = len(sheet.col_values(b)) - 1
    for i in range(1, val + 1):
        ws.write(0, i, sheet.row_values(0)[i])
    for i in range(1, val1 + 1):
        #ws.write(i, 0, sheet.row_values(i)[0])
        just_class.append(sheet.row_values(i)[0])

    b = val
    a = val1
    for_class_number = []
    p = 0
    o = 0
    lol = []
    qu = 0
    help_me_please = []
    o1 = -1
    v_main = 0
    pop = []
    qu_ = []
    pop1 = []
    u = 0
    u_ = []
    s = []
    if count == 0:
        ws.write(0, b + 5, "Итоговая сумма")
    # print("Вводите баллы за задачу по командам.")
    for i in range(a + 1):
        for j in range(b + 1):

            #if j == 0 and i != 0:
                #ws.write(i, 0, "Класс {}".format(i + 8))
            #elif i == 0 and j != 0:
                #ws.write(0, j, "Задача {}".format(j))

            if i != 0 and j != 0:
                p = sheet.row_values(i)[j]
                #ws.write(i, j, p)
                o += p
                pop.append(p)
        if i != 0:
            ws.write(i, j + 2, o)
            pop1.append(o)
        o = 0
    for _c_ in range(a):
        for _s_ in range(b):
            help_me_please.append(1)
        lol.append(help_me_please)
        help_me_please = []
    for _col_ in range(a):
        for _str_ in range(b):
            lol[_col_][_str_] = pop[_str_ + _col_ * b]
    ws.write(0, j + 2, "Сумма баллов")
    for z in range(b):
        for q in range(len(pop)):
            if q % b == z:
                u += pop[q]

        ws.write(a + 2, z + 1, u)
        if u == 0:
            u = 1
        u_.append(u)
        u = 0
    for y in range(a):
        for _ in range(len(u_)):
            qu += pop[_ + b * y] / u_[_]
        qu *= 10
        qu_.append(qu)
        bal.append(qu)
        qu = 0
    for _ in range(len(bal)):
        for v in range(len(bal)):
            if o1 != max(bal[v], o1):
                v_main = v
                o1 = max(bal[v], o1)
        maximum_class.append(o1)
        bal[v_main] = -10
        o1 = -1
    for ___ in range(len(maximum_class)):
        for _0 in range(len(qu_)):
            if qu_[_0] == maximum_class[___]:
                just.append(just_class[_0])
                just_class[_0] = -9
                for_class_number.append(_0)
                break

    for klop in range(1, len(just) + 1):
        ws.write(klop, 0, just[klop - 1])
    for _cucumber_ in range(len(for_class_number)):
        oo = lol[_cucumber_]
        lol[_cucumber_] = lol[for_class_number[_cucumber_]]
        lol[for_class_number[_cucumber_]] = oo
    for _i_ in range(1, a + 1):
        for _d_ in range(1, b + 1):
            ws.write(_i_, _d_, lol[_i_ - 1][_d_ - 1])
    for _1 in range(len(pop1)):
        s.append((pop1[_1] + qu_[_1]))
    for plague in range(1, len(s) + 1):
        ws.write(plague, b + 5, s[plague - 1])
    wb.save(name + '_обработанный.xls')
    count += 1
    time.sleep(60)
