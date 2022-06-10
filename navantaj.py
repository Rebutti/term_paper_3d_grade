import openpyxl
import decimal
import re


def toFixed(numObj, digits=0):
    return float(f"{numObj:.{digits}f}")


def findgtype(sheet):
    groupname = sheet["A2"].value
    groupname = groupname[groupname.find('(')+1:]
    groupname = groupname[:groupname.find(' ')]
    return groupname


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
        if sheetfile1[letter+str(obov_disc)].value.lower().find("вибірковими") != -1:
            break
        else:
            obov_disc += 1
            kilkist_obov_disc += 1
    return kilkist_obov_disc


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
    while counter < 10:
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
            exz_or_icpit += int(values["атест_екз_консультації"])
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


def navantaj(sheetfile1, values, kil_stud, sheet, course):
    number1 = findlpl(sheetfile1)
    number = number1[0]

    kil_tij_1_cem = int(str(sheetfile1["O7"].value).strip()[:2])
    kil_tij_2_cem = int(str(sheetfile1["S7"].value).strip()[:2])
    kil_groups = sheetfile1["A3"].value.split(" ")
    kilkist_groups = []
    for i in filter(lambda x: str(x).isdigit(), kil_groups):
        kilkist_groups.append(i)
    kilkist_groups = int(kilkist_groups[0])
    lekciya1 = toFixed(float(sheetfile1["O"+str(number)].value), 2)
    praktika1 = toFixed(float(sheetfile1["P"+str(number)].value), 2)
    labi1 = toFixed(float(sheetfile1["Q"+str(number)].value), 2)
    lekciya2 = toFixed(float(sheetfile1["S"+str(number)].value), 2)
    praktika2 = toFixed(float(sheetfile1["T"+str(number)].value), 2)
    labi2 = toFixed(float(sheetfile1["U"+str(number)].value), 2)
    if kil_stud > int(values["студ_зал"]):
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
    zaliki = str(sheetfile1["H"+str(number)].value).strip().split(' ')
    all_hours = sheetfile1["M"+str(number)].value

    dia3 = kil_stud/int(values["екз"])
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
    dia4 = exzamens*int(values["пров_екз"])*kilkist_groups

    dia5 = exzamens*int(values["конс_пред_екз"])*kilkist_groups
    if kil_stud > int(values["студ_зал"]):
        dia6 = zaliki*int(values["заліки"])*kilkist_groups
    else:
        dia6 = zaliki*int(values["заліки"])/2

    k_pot_kons = findgtype(sheetfile1)
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
    k_pot_kons = int(k_pot_kons)
    dia7 = all_hours*k_pot_kons*kil_stud/100/int(values["академ_груп"])
    kil_individ = 0
    Kr = 0
    Kp = 0
    row = 10
    letter1 = "K"
    nowords = ["кр", "кп"]
    while row < number:
        if sheetfile1[letter1+str(row)].value != None and sheetfile1[letter1+str(row)].value.lower() not in nowords:
            kil_individ += 1
        elif sheetfile1[letter1+str(row)].value != None:
            if sheetfile1[letter1+str(row)].value.lower() in nowords:
                kil_individ -= 1
            if sheetfile1[letter1+str(row)].value.lower() == "кр":
                Kr += 1
            else:
                Kp += 1
        row += 1
    dia8 = kil_stud/int(k_indiv)
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
                if kil_stud >= int(values["студ_нав_практ"]):
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
        dia14 + vibirkovi
    nav = int(str(decimal.Decimal(nav).quantize(
        decimal.Decimal('0'), rounding=decimal.ROUND_HALF_UP)))
    return nav


if __name__ == "__main__":
    file1 = openpyxl.open("C:/Users/Touch.com.ua/Учеба/6 семестр/курсавая/курсова3py/main/excel files/ФПСО/Копия ДС 1-6к. 2022.xlsx",
                          read_only=True, data_only=True)
    sheet = "1-й курс "
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
        navantaj(sheetfile1, values, E, sheet, course)
    except:
        print("Error")
    finally:
        file1.close()
