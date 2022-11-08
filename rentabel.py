from openpyxl.styles.numbers import BUILTIN_FORMATS
from navantaj import navantaj
import PySimpleGUI as sg


def findgname(sheet):
    groupname = sheet["A2"].value
    groupname = groupname.replace(' ', '')
    if groupname[10] == 'y' or groupname[10] == 'у':
        return groupname[5:7]+'у'
    return groupname[5:7]


def findgfname(sheet):
    groupname = sheet["A1"].value
    str_ = ''
    str_2 = ''
    groupname = groupname.strip()
    groupname = groupname.replace('-', ' ')
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
    money = toFixed(float(I9)*float(F9)+float(G9)
                    * float(J9)+float(H9)*float(K9))
    return money


def check_courses(course_in_file1, course_in_file2, sheetname, file2_number_of_course):
    print('course_in_file1: '+course_in_file1)
    print('course_in_file2: '+course_in_file2)
    sheetname = str(sheetname)
    file2_number_of_course = str(file2_number_of_course)
    if course_in_file1.upper() == course_in_file2.upper():
        if file2_number_of_course == '1м':
            file2_number_of_course = '5'
        elif file2_number_of_course == '2м':
            file2_number_of_course = '6'
        if file2_number_of_course in sheetname:
            return True
        
        print(file2_number_of_course in sheetname)
        print(sheetname)
        print(file2_number_of_course)
    return False

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
    if check_courses(findgname(sheetfile1), sheetfile2["A"+str(number)].value, sheet, sheetfile2["D"+str(number)].value):
        pass
    else:
        sg.popup(f'Курси або шифри у файлах не збігаються: {sheet} та {sheetfile2["D"+str(number)].value}, {findgname(sheetfile1)} та {sheetfile2["A"+str(number)].value}')
        return None
    savesheet['A'+str(number)] = findgname(sheetfile1)
    savesheet["B"+str(number)] = findgfname(sheetfile1)
    savesheet["C"+str(number)] = findgtype(sheetfile1)
    course = findcourse(sheet)
    savesheet["D"+str(number)] = course[0]
    savesheet["F"+str(number)] = sheetfile2["F"+str(number)].value
    savesheet["G"+str(number)] = sheetfile2["G"+str(number)].value
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
    E = int(F)+int(G)+int(H)
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
    try:
        minus_hours = float(sheetfile2["R"+str(number)].value)
    except:
        minus_hours = 0
    L = navantaj(sheetfile1=sheetfile1, values=values,
                 kil_stud=E, sheet=sheet, course=course[1], minus_hours=minus_hours)
    savesheet["L"+str(number)] = L
    savesheet["L"+str(number+1)] = "=SUM(L9:"+"L"+str(number)+")"
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
    savesheet["P"+str(number+1)].number_format = BUILTIN_FORMATS[9]
    savesheet["P"+str(number+1)] = "=(O"+str(number+1) + \
        "/"+"N"+str(number+1)+")"
    Q = float(toFixed(float(float(N)-float(O)), 0))
    savesheet["Q"+str(number)] = Q
    savesheet["Q"+str(number+1)] = "=SUM(Q9:"+"Q"+str(number)+")"

if __name__ == '__main__':
    a = findgname('Група                                         ПА-22у (денна форма навчання)')
    print(a)
