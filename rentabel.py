
from types import NoneType
#import openpyxl
from openpyxl.styles.numbers import BUILTIN_FORMATS
from navantaj import navantaj


def findgname(sheet):
    groupname = sheet["A2"].value
    groupname = groupname.replace(' ', '')
    groupname = groupname[5:7]
    return groupname


def findgfname(sheet):
    groupname = sheet["A1"].value
    str_ = ''
    str_2 = ''
    groupname = groupname.strip()
    groupname = groupname.replace('-', ' ')
    #groupname = groupname.split(' ')
    for i in groupname:
        if i.isdigit():
            index = groupname.index(i)
            str_ = groupname[index:]
            break
    str_2 = str_
    for i in str_:
        if i == '(':
            str_2 = str_[:str_.index(i)]
            break
    str_2 = str_2.split(" ")
    while True:
        try:
            str_2.remove("")
        except ValueError:
            break

    return ' '.join(str_2)


def findgtype(sheet):
    groupname = sheet["A2"].value
    groupname = groupname[groupname.find('(')+1:]
    groupname = groupname[:groupname.find(' ')]
    return groupname


def findcourse(sheetname: str):
    sheetname = sheetname.replace(' ', '')
    if sheetname[0:1] == '5':
        return '1м', 1
    elif sheetname[0:1] == '6':
        return '2м', 1
    return sheetname[0:1], 0


def toFixed(numObj, digits=0):
    return f"{numObj:.{digits}f}"


def allmoney(I9, F9, G9, J9, H9, K9):

    # groupname = float(sheet["F"+str(row)].value) * \
    #     float(sheet["I"+str(row)].value)+float(sheet["G" +
    #                                                  str(row)].value)*float(sheet["J"+str(row)].value)+float(sheet["H"+str(row)].value)*float(sheet["K"+str(row)].value)
    money = toFixed(float(I9)*float(F9)+float(G9)
                    * float(J9)+float(H9)*float(K9))
    return money


# def navantaj(sheetfile1, sheetfile2, row1):
#     letter = "B"
#     number = 8
#     let_num = letter+str(number)
#     while number < 50:
#         if sheetfile1[let_num].value == None:
#             number += 1
#             let_num = letter+str(number)
#             continue
#         elif(str(sheetfile1[let_num].value).replace(' ', '') == "Разом"):
#             break
#         else:
#             number += 1
#             let_num = letter+str(number)

#     nav = (float(sheetfile1["O"+str(number)].value) *
#            int(sheetfile1["O7"].value[:3])+float(sheetfile1["R"+str(number)].value)) + \
#         (float(sheetfile1["S"+str(number)].value) *
#          int(sheetfile1["S7"].value[:3])+float(sheetfile1["V"+str(number)].value)) + \
#         int(sheetfile2['E'+str(row1)].value)
#     print("navantaj = ", nav)


def rentabel(sheetfile1, sheetfile2, savesheet, ser_nav_nav, ser_zar_plat, ESV, pev_vel, sheet, values):
    letter = "A"
    number = 8
    let_num = letter+str(number)
    while True:
        if(savesheet[let_num].value == None):
            break
        else:
            number += 1
            let_num = letter+str(number)
    # заполнение столбцов
    savesheet['A'+str(number)] = findgname(sheetfile1)
    savesheet["B"+str(number)] = findgfname(sheetfile1)
    savesheet["C"+str(number)] = findgtype(sheetfile1)
    course = findcourse(sheet)
    savesheet["D"+str(number)] = course[0]
    savesheet["F"+str(number)] = sheetfile2["F"+str(number)].value
    print(sheetfile2["F"+str(number)].value, " = F")
    savesheet["G"+str(number)] = sheetfile2["G"+str(number)].value
    print(sheetfile2["G"+str(number)].value, " = G")
    if sheetfile2["H"+str(number)].value != None:
        savesheet["H"+str(number)] = sheetfile2["H"+str(number)].value
    else:
        savesheet["H"+str(number)] = 0
    F = sheetfile2["F"+str(number)].value
    G = sheetfile2["G"+str(number)].value
    if sheetfile2["H"+str(number)].value != None:
        H = sheetfile2["H"+str(number)].value
    else:
        H = 0
    E = F+G+H
    savesheet["E"+str(number)] = E
    I = sheetfile2["I"+str(number)].value
    savesheet["I"+str(number)] = I
    J = sheetfile2["J"+str(number)].value
    savesheet["J"+str(number)] = J
    if sheetfile2["K"+str(number)].value != None:
        K = sheetfile2["K"+str(number)].value
    else:
        K = 0
    savesheet["K"+str(number)] = K
    # навантаження доделать 12 column
    L = navantaj(sheetfile1=sheetfile1, values=values,
                 kil_stud=E, sheet=sheet, course=course[1])
    print("navantaj = ", L, "kil_stud = ", E)
    # navantaj(sheetfile1, sheetfile2, number)
    savesheet["L"+str(number)] = L
    savesheet["L"+str(number+1)] = "=SUM(L9:"+"L"+str(number)+")"
    # ///
    M = float(toFixed(float(L)/float(ser_nav_nav), 2))
    savesheet["M"+str(number)] = M
    savesheet["M"+str(number+1)] = "=SUM(M9:"+"M"+str(number)+")"
    N = toFixed(float(F) * float(I) + float(G)*float(J)+float(H)*float(K))
    savesheet["N"+str(number)] = float(N)
    savesheet["N"+str(number+1)] = "=SUM(N9:"+"N"+str(number)+")"
    O = float(toFixed(float(M)*float(ser_zar_plat)
              * 12*float(ESV)+float(pev_vel)))
    savesheet["O"+str(number)] = O
    savesheet["O"+str(number+1)] = "=SUM(O9:"+"O"+str(number)+")"
    savesheet["P"+str(number)].number_format = BUILTIN_FORMATS[9]
    P = (float(O)/float(N))
    savesheet["P"+str(number)] = P
    # savesheet["P"+str(number+1)] = "=SUM(O9:"+"O"+str(number) + \
    #     ")/" + "(SUM(N9:"+"N"+str(number)+"))*100"
    savesheet["P"+str(number+1)].number_format = BUILTIN_FORMATS[9]
    savesheet["P"+str(number+1)] = "=(O"+str(number+1) + \
        "/"+"N"+str(number+1)+")"
    Q = float(toFixed(float(float(N)-float(O)), 0))
    savesheet["Q"+str(number)] = Q
    savesheet["Q"+str(number+1)] = "=SUM(Q9:"+"Q"+str(number)+")"


if __name__ == "__main__":
    print(1, 2)
