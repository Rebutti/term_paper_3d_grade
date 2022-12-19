import openpyxl
import decimal
import re


def toFixed(numObj, digits=0):
    return float(f"{numObj:.{digits}f}")


def findgtype(sheet):
    letter = 'A'
    number = 1
    groupname = None
    while number < 50:
        if sheet[letter+str(number)].value != None:
            if 'група' in sheet[letter+str(number)].value.lower():
                groupname = sheet[letter+str(number)].value
                groupname = groupname[groupname.find('(')+1:]
                groupname = groupname[:groupname.find(' ')]
                return groupname
            else:
                number += 1
        else:
            number += 1
    return "Не вдалося знайти комірку з даними про групу"


def findlpl(sheetfile1) -> int:
    letter = 'B'
    number = 8
    while True:
        if sheetfile1[letter+str(number)].value != None:
            if str(sheetfile1[letter+str(number)].value).strip() == "Разом":
                return (number, 0)
            if str(sheetfile1[letter+str(number)].value).strip() == "Разом за обов'язковими компонентами":
                if str(sheetfile1[letter+str(number+1)].value).strip().lower() == "разом":
                    return (number, 0)
                return (number, 1)
            else:
                number += 1
        else:
            number += 1


def find_vibir_disc(sheetfile1, obov_disc: int):
    letter = 'B'
    obov_disc = obov_disc+1
    kilkist_obov_disc = 0
    while True:
        print(sheetfile1[letter+str(obov_disc)].value)
        if sheetfile1[letter+str(obov_disc)].value == None:
            obov_disc += 1
            continue
        if sheetfile1[letter+str(obov_disc)].value.lower().find("вибірковими") != -1:
            break
        
        else:
            obov_disc += 1
            kilkist_obov_disc += 1
    return kilkist_obov_disc

def find_atec_ex(sheetfile1, obov_disc: int):
    letter = 'B'
    obov_disc = 1
    for row in range(1, 51):
        if sheetfile1[letter+str(obov_disc)].value == None:
            obov_disc+=1
            continue
        if str(sheetfile1[letter+str(obov_disc)].value).lower().find("атестаційний екзамен (ек)") != -1:
            return 1
        
        else:
            obov_disc += 1
    return 0

def findprakt(sheetfile1):
    letter = 'B'
    number = 8
    a = []
    while number < 100:

        if sheetfile1[letter+str(number)].value != None:
            if str(sheetfile1[letter+str(number)].value).strip() == "Вид практики":
                number += 1
                a.append((sheetfile1[letter+str(number)
                                     ].value.strip().lower(), number))
                number += 1
                continue
            elif str(sheetfile1[letter+str(number)].value).strip().lower().find("виробнича") != -1:
                a.append((sheetfile1[letter+str(number)
                                     ].value.strip().lower(), number))
                number += 1
                continue
            elif str(sheetfile1[letter+str(number)].value).strip().lower().find("навчальна") != -1:
                a.append((sheetfile1[letter+str(number)
                                     ].value.strip().lower(), number))
                number += 1
                continue

            else:
                number += 1
        else:
            number += 1
    if a != []:
        return a
    return None


def findatect(sheetfile1, kil_stud, values, sheet):
    letter = 'B'
    number = 8
    atestacia_adress = None
    result = 0
    sheet = sheet.strip()
    flag = False
    exz_or_icpit = 0
    while number < 100:
        if sheetfile1[letter+str(number)].value != None:
            if str(sheetfile1[letter+str(number)].value).strip().lower().find("атестація") != -1:
                flag = True
                number += 1
                break
            else:
                number += 1
        else:
            number += 1
    if flag == False:
        return 0
    counter = 0
    atestacia_adress = []
    while counter < 50:
        if sheetfile1["B"+str(number)].value == None:
            counter += 1
        else:
            atestacia_adress.append(
                [sheetfile1["B"+str(number)].value, number])
            counter = 0
        number += 1
    for atest in atestacia_adress:
        b = re.search('[(].+[)]', atest[0])
        if b == None:
            exz_or_icpit += float(values["атест_екз_консультації"])
            continue
        a = b.span()
        atest[0] = atest[0][a[0]:a[1]].replace('(', '').replace(')', '')
        if atest[0] == "ЕК":
            result += int(kil_stud) / \
                float(values['атест_ЕК']) * \
                int(sheetfile1["K"+str(atest[1])].value)
        elif atest[0] == "керівництво":
            if int(sheet[0]) <= 4:
                kval_rob = float(values["квал_роб_керівництво2_до_5к"])
            else:
                kval_rob = float(values["квал_роб_керівництво2_5_та_6к"])
            result += kil_stud * kval_rob
        elif atest[0] == "рецензування":
            if int(sheet[0]) <= 4:
                kval_rob = float(values["квал_роб_рецензування_до_5к"])
            else:
                kval_rob = float(values["квал_роб_рецензування_5_та_6к"])
            result += kil_stud * kval_rob
        else:
            kval_rob = float(values["квал_роб_керівництво1"])
            result += kil_stud * kval_rob
    result += exz_or_icpit
    return result


def find_kil_groups(sheet):
    letter = 'A'
    number = 1
    kil_groups = None
    while number < 20:
        if sheet[letter+str(number)].value != None:
            if 'кількість' in str(sheet[letter+str(number)].value).lower():
                kil_groups = sheet[letter+str(number)].value
                kilkist_groups = []
                for i in filter(lambda x: str(x).isdigit(), kil_groups):
                    kilkist_groups.append(i)
                kilkist_groups = int(kilkist_groups[0])
                return kilkist_groups
            else:
                number += 1
        else:
            number += 1
    return "Не вдалося знайти комірку з даними про кількіть груп"


def minuses_counter(sheet: openpyxl.Workbook, kil_gruops, lpl, amount_of_1weeks, amount_of_2weeks, spec_chislo):
    result = 0
    letter = 'D'
    number = 1
    for i in range(1, 15):
        if '4' in str(sheet[letter+str(number)].value):
            number += 1
            break
        else:
            number += 1
    for row in range(number, lpl):
        # считаю для колонки 4
        a = 0
        b = 0
        letter = 'D'
        if sheet[letter+str(row)].value != None:
            kol4 = find_number_pg(sheet, letter, row)
            if type(kol4) == tuple:
                kol4 = kol4[0]
            if kol4 != kil_gruops:
                kol16, kol20 = validation_for_hours_for_lessons(
                    sheet, 'P', 'T', row)
            if kol4 < kil_gruops:
                a = abs(kol16*amount_of_1weeks*kil_gruops -
                        kol16*amount_of_1weeks*kol4)*(-1)
                b = abs(kol20*amount_of_2weeks*kil_gruops -
                        kol20*amount_of_2weeks*kol4)*(-1)
            elif kol4 > kil_gruops:
                a = abs(kol16*amount_of_1weeks*(kol4-kil_gruops))
                b = abs(kol20*amount_of_2weeks*(kol4-kil_gruops))
            result = result + a + b
        # считаю для колонки 5
        letter = 'E'
        a = 0
        b = 0
        if sheet[letter+str(row)].value != None:
            kol5 = find_number_pg(sheet, letter, row)
            
            kol17, kol21 = validation_for_hours_for_lessons(
                sheet, 'Q', 'U', row)
            if type(kol5) == tuple:
                kol5 = kol5[0]
            if kol5 < kil_gruops*spec_chislo:
                a = abs(kol17*amount_of_1weeks*kil_gruops *
                        spec_chislo-kol17*amount_of_1weeks*kol5)*(-1)
                b = abs(kol21*amount_of_2weeks*kil_gruops *
                        spec_chislo-kol21*amount_of_2weeks*kol5)*(-1)
            elif kol5 > kil_gruops*spec_chislo:
                a = abs(kol17 * amount_of_1weeks *
                        (kol5-kil_gruops*spec_chislo))
                b = abs(kol21 * amount_of_2weeks *
                        (kol5-kil_gruops*spec_chislo))
            result = result + a + b
        # считаю для колонки 6
        letter = 'F'
        a = 0
        b = 0
        flag = False
        if sheet[letter+str(row)].value != None:
            kol6 = find_number_pg(sheet, letter, row)
            if type(kol6) == tuple:
                flag = kol6[1]
                kol6 = kol6[0]
            if kol6 != kil_gruops:
                kol15, kol19 = validation_for_hours_for_lessons(
                    sheet, 'O', 'S', row)
                if flag == True:
                    a = abs(kol15*amount_of_1weeks-kol15 *
                            amount_of_1weeks*1/kol6)*(-1)
                    b = abs(kol19*amount_of_2weeks-kol19 *
                            amount_of_2weeks*1/kol6)*(-1)
                else:
                    if kol6 < 1:
                        a = abs(kol15*amount_of_1weeks-kol15 *
                                amount_of_1weeks*kol6)*(-1)
                        b = abs(kol19*amount_of_2weeks-kol19 *
                                amount_of_2weeks*kol6)*(-1)
                    elif kol6 > 1:
                        a = abs(kol15*amount_of_1weeks*(kol6-1))
                        b = abs(kol19*amount_of_2weeks*(kol6-1))
            result = result + a + b
    return result


def find_number_pg(sheet, letter, row):
    if sheet[letter+str(row)].value == ' ':
        return 0
    try:
        kol = float(sheet[letter+str(row)].value)
        return kol
    except:
        kol = len(str(sheet[letter+str(row)].value).strip().split(' '))+1
        return kol, True


def validation_for_hours_for_lessons(sheet, letter_1_col, letter_2_col, row):
    if sheet[letter_1_col+str(row)].value != None:
        column1 = float(sheet[letter_1_col+str(row)].value)
    else:
        column1 = 0
    if sheet[letter_2_col+str(row)].value != None:
        column2 = float(sheet[letter_2_col+str(row)].value)
    else:
        column2 = 0
    return column1, column2


def navantaj(sheetfile1, values, kil_stud, sheet, course, minus_hours):
    number1 = findlpl(sheetfile1)
    number = number1[0]
    kil_tij_1_cem = int(str(sheetfile1["O7"].value).strip()[:2])
    kil_tij_2_cem = int(str(sheetfile1["S7"].value).strip()[:2])
    kilkist_groups = find_kil_groups(sheetfile1)
    if kilkist_groups == 'Не вдалося знайти комірку з даними про кількіть груп':
        return (kilkist_groups, 'Error')
    lekciya1 = toFixed(float(
        sheetfile1["O"+str(number)].value if sheetfile1["O"+str(number)].value != None else 0), 2)
    praktika1 = toFixed(float(
        sheetfile1["P"+str(number)].value if sheetfile1["P"+str(number)].value != None else 0), 2)
    labi1 = toFixed(float(
        sheetfile1["Q"+str(number)].value if sheetfile1["Q"+str(number)].value != None else 0), 2)
    lekciya2 = toFixed(float(
        sheetfile1["S"+str(number)].value if sheetfile1["S"+str(number)].value != None else 0), 2)
    praktika2 = toFixed(float(
        sheetfile1["T"+str(number)].value if sheetfile1["T"+str(number)].value != None else 0), 2)
    labi2 = toFixed(float(
        sheetfile1["U"+str(number)].value if sheetfile1["U"+str(number)].value != None else 0), 2)
    if kil_stud > float(values["студ_зал"]):
        spec_chislo = 2
    else:
        spec_chislo = 1
    sem1 = lekciya1*kil_tij_1_cem+praktika1 * \
        kil_tij_1_cem*kilkist_groups+labi1*kil_tij_1_cem * \
        kilkist_groups*spec_chislo
    sem2 = lekciya2*kil_tij_2_cem+praktika2 * \
        kil_tij_2_cem*kilkist_groups+labi2*kil_tij_2_cem * \
        kilkist_groups*spec_chislo

    exzamens = str(sheetfile1["G"+str(number)].value).strip().split(' ')
    for el in exzamens:
        if el == '':
            exzamens.remove(el)
    zaliki = str(sheetfile1["H"+str(number)].value).strip().split(' ')
    for el in zaliki:
        if el == '':
            zaliki.remove(el)
    all_hours = sheetfile1["M"+str(number)].value

    dia3 = kil_stud*float(values["екз"])
    if dia3 > 1:
        dia3 = int(str(decimal.Decimal(dia3).quantize(
            decimal.Decimal('0'), rounding=decimal.ROUND_HALF_UP)))
    else:
        dia3 = 1
    exzamens = [int(i) for i in exzamens]
    exzamens = sum(exzamens)
    zaliki = [int(i) for i in zaliki]
    zaliki = sum(zaliki)
    dia3 *= exzamens
    dia4 = exzamens*float(values["пров_екз"])*kilkist_groups

    dia5 = exzamens*float(values["конс_пред_екз"])*kilkist_groups
    if kil_stud > float(values["студ_зал"]):
        dia6 = zaliki*float(values["заліки"])*kilkist_groups
    else:
        dia6 = zaliki*float(values["заліки"])/2

    k_pot_kons = findgtype(sheetfile1)
    if k_pot_kons == "Не вдалося знайти комірку з даними про групу":
        return (k_pot_kons, "Error")
    if k_pot_kons == "денна":
        k_pot_kons = values['пот_конс_денна']
        k_indiv = values['індивід_денна/вечірня']
        if course == 0:
            k_vibirkovi_disc = values['вибір_дисц_бакалавр_денна']
        else:
            k_vibirkovi_disc = values['вибір_дисц_магістр_денна']
    elif k_pot_kons == "вечірня":
        k_pot_kons = values['пот_конс_вечірня']
        k_indiv = values['індивід_денна/вечірня']
        k_vibirkovi_disc = values['вибір_дисц_бакалавр_вечірня']
    elif k_pot_kons == "заочна":
        k_pot_kons = values['пот_конс_заочна']
        k_indiv = values['індивід_заочна']
        if course == 0:
            k_vibirkovi_disc = values['вибір_дисц_бакалавр_заочна']
        else:
            k_vibirkovi_disc = values['вибір_дисц_магістр_заочна']
    elif k_pot_kons == "дуальна":
        k_pot_kons = values['пот_конс_дуальна']
        k_indiv = values['індивід_дуальна']
    else:
        k_pot_kons = 1
    k_pot_kons = float(k_pot_kons)
    dia7 = all_hours*k_pot_kons*kil_stud/100*float(values["академ_груп"])
    kil_individ = 0
    Kr = 0
    Kp = 0
    row = 10
    letter1 = "K"
    nowords = ["кр", "кп"]
    while row < number:
        if sheetfile1[letter1+str(row)].value != None and str(sheetfile1[letter1+str(row)].value).lower() not in nowords:
            kil_individ += 1
        elif sheetfile1[letter1+str(row)].value != None:
            if str(sheetfile1[letter1+str(row)].value).lower() in nowords:
                kil_individ -= 1
            if str(sheetfile1[letter1+str(row)].value).lower() == "кр":
                Kr += 1
            else:
                Kp += 1
        row += 1
    dia8 = kil_stud*int(k_indiv)
    if dia8 < 1:
        dia8 = 1
    else:
        dia8 = int(str(decimal.Decimal(dia8).quantize(
            decimal.Decimal('0'), rounding=decimal.ROUND_HALF_UP)))*kil_individ
    dia9 = Kr*float(values["кр"])*kil_stud
    dia10 = Kp*float(values["кп"])*kil_stud
    dia11 = Kr*float(values["зах_кр"])*kil_stud
    dia12 = Kp*float(values["зах_кп"])*kil_stud
    dia13 = 0
    dia14 = 0
    dia14 = findatect(sheetfile1, kil_stud, values, sheet)
    praktika = findprakt(sheetfile1)
    if praktika != None:
        for name_, numbe in praktika:
            kil_tij_for_prakt = float(sheetfile1["E"+str(numbe)].value)
            if dia14 != 0:
                dia13 += float(values["вир_пр_переддипломна"]) * \
                    kil_stud*kil_tij_for_prakt
                continue
            if name_.find("навчальна") != -1:
                if kil_stud >= float(values["студ_нав_практ"]):
                    dia13 += float(values["нав_практ1"]) * \
                        kil_tij_for_prakt*kilkist_groups
                else:
                    dia13 += float(values["нав_практ2"]) * \
                        kil_tij_for_prakt*kil_stud
            else:
                if kil_stud >= float(values["студ_вир_практ"]):
                    dia13 += float(values["вир_практ1"]) * \
                        kil_tij_for_prakt*kilkist_groups
                else:
                    dia13 += float(values["вир_практ2"]) * \
                        kil_tij_for_prakt*kil_tij_for_prakt*kil_stud

    kil_vibir_disc = 0
    if number1[1] == 1:
        kil_vibir_disc = find_vibir_disc(sheetfile1, number)
    vibirkovi = kil_vibir_disc*float(k_vibirkovi_disc)*kil_stud
    nav = sem1+sem2 + dia3 + dia4 + \
        dia5 + dia6 + dia7 + dia8 + dia9+dia10+dia11+dia12+dia13 + \
        dia14 + vibirkovi + minus_hours + \
        minuses_counter(sheetfile1, kilkist_groups, number,
                        kil_tij_1_cem, kil_tij_2_cem, spec_chislo)
    # print(nav, ' = nav')
    # print(f"{sem1}+{sem2} + {dia3} + {dia4} +{dia5} + {dia6} + {dia7} + {dia8} + {dia9}+{dia10}+{dia11}+{dia12}+{dia13} + {dia14} + {vibirkovi} + {minus_hours}+{minuses_counter(sheetfile1, kilkist_groups, number, kil_tij_1_cem, kil_tij_2_cem, spec_chislo)}")
    nav = int(str(decimal.Decimal(nav).quantize(
        decimal.Decimal('0'), rounding=decimal.ROUND_HALF_UP)))
    return nav


if __name__ == "__main__":
    file1 = openpyxl.open("D:/Учеба/7_семестр/курсавая/tests/New folder/Копия ПАу 1-2к. 2022.xlsx",
                          read_only=True, data_only=True)
    sheet = "5-й курс "
    sheetfile1 = file1[sheet]
    E = 24
    course = 1
    values = {0: '580', 1: '17100', 2: '1.22', 3: '10590.41', 'вибіркові_дисципліни': '0',
              4: 'C:/Users/Touch.com.ua/Учеба/6 семестр/курсавая/курсова3py/main/2022 РП ІФ/xlsx/1-6k/Копия ІП 1-6 к.2022.xlsx',
              'Переглянути': 'C:/Users/Touch.com.ua/Учеба/6 семестр/курсавая/курсова3py/main/2022 РП ІФ/xlsx/1-6k/Копия ІП 1-6 к.2022.xlsx',
              5: 'C:/Users/Touch.com.ua/Учеба/6 семестр/курсавая/курсова3py/main/2022 РП ІФ/xlsx/1-6k/ІФ_РОЗРАХУНОК.xlsx',
              'Переглянути0': 'C:/Users/Touch.com.ua/Учеба/6 семестр/курсавая/курсова3py/main/2022 РП ІФ/xlsx/1-6k/ІФ_РОЗРАХУНОК.xlsx',
              'filetrue': False, 6: True, 7: '', 'dop_file': '', 'вибір_дисц_бакалавр_денна': '0.7',
              'вибір_дисц_магістр_денна': '0.67', 'вибір_дисц_бакалавр_заочна': '0.17', 'вибір_дисц_магістр_заочна': '0.16',
              'вибір_дисц_бакалавр_вечірня': '0.64', 'екз': '4', 'пров_екз': '2', 'конс_пред_екз': '2',
              'заліки': '2', 'студ_зал': '25', 'пот_конс_денна': '4', 'пот_конс_вечірня': '6', 'пот_конс_заочна': '8',
              'пот_конс_дуальна': '10', 'академ_груп': '25', 'індивід_денна/вечірня': '4', 'індивід_заочна': '2',
              'індивід_дуальна': '4', 'кр': '1', 'кп': '2', 'зах_кр': '1', 'зах_кп': '1', 'нав_практ1': '30',
                 'нав_практ2': '15', 'вир_практ1': '15', 'вир_практ2': '1', 'вир_пр_переддипломна': '1',
                 'студ_нав_практ': '15', 'студ_вир_практ': '15', 'атест_ЕК': '2', 'атест_екз_консультації': '8',
              'квал_роб_керівництво1': '0', 'квал_роб_керівництво2_до_5к': '8', 'квал_роб_керівництво2_5_та_6к': '31',
              'квал_роб_рецензування_до_5к': '2', 'квал_роб_рецензування_5_та_6к': '4', 8: 'Головна'}
    try:
        print(minuses_counter(sheetfile1, find_kil_groups(
            sheetfile1), findlpl(sheetfile1)[0], 16, 16, 1))
    except:
        print("Error")
    finally:
        file1.close()
    # minuses_counter(sheetfile1, find_kil_groups(
    #         sheetfile1), findlpl(sheetfile1)[0], 16, 16, 1)
