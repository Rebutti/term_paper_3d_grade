coefficients = ["вибір_дисц_бакалавр_денна",
                "вибір_дисц_магістр_денна",
                "вибір_дисц_бакалавр_заочна",
                "вибір_дисц_магістр_заочна",
                "вибір_дисц_бакалавр_вечірня",
                "екз",
                "пров_екз",
                "конс_пред_екз",
                "заліки",
                "студ_зал",
                "пот_конс_денна",
                "пот_конс_вечірня",
                "пот_конс_заочна",
                "пот_конс_дуальна",
                "академ_груп",
                "індивід_денна/вечірня",
                "індивід_заочна",
                "індивід_дуальна",
                "кр",
                "кп",
                "зах_кр",
                "зах_кп",
                "нав_практ1",
                "нав_практ2",
                "вир_практ1",
                "вир_практ2",
                "вир_пр_переддипломна",
                "студ_нав_практ",
                "студ_вир_практ",
                "атест_ЕК",
                "атест_екз_консультації",
                "квал_роб_керівництво1",
                "квал_роб_керівництво2_до_5к",
                "квал_роб_керівництво2_5_та_6к",
                "квал_роб_рецензування_до_5к",
                "квал_роб_рецензування_5_та_6к",
                "поток",
                "D1coeff",
                'E1coeff',
                'НПП_витрати',
                'поточні_консультації',
                "бюджет",
                "контракт",
                "бюджет1",
                "бюджет2",
                "бюджет3",
                "бюджет4",
                "бюджет5",
                "бюджет6",
                "контракт1",
                "контракт2",
                "контракт3",
                "контракт4",
                "контракт5",
                "контракт6"]

main_values = [0, 1, 2, 3]


def save_coef(coeffs, coeff=True):
    if coeff == False:
        values = main_values
        file_name = 'main_values.txt'
    else:
        values = coefficients
        file_name = 'coefficients.txt'
    with open(file_name, 'w', encoding='utf-8') as file:
        for k in values:
            file.write(str(coeffs[k]))
            file.write('\n')


def check_coef():
    value = []
    file_name = ['main_values.txt', 'coefficients.txt']
    k_values = main_values + coefficients
    for f in file_name:
        with open(f, 'r', encoding="utf-8") as file:
            coefs = file.readlines()
            value += coefs
    coefs = ''.join(value).split('\n')
    # print(coefs)
    values = {}
    n = 0
    for k in k_values:
        values[k] = coefs[n]
        n += 1

    return values


if __name__ == "__main__":
    values = {0: '580', 1: '17100', 2: '1.22', 3: '10590.41', 4: 'D:/Учеба/7_семестр/курсавая/tests/Копия ПА 1-6 к. 2022.xlsx', 'Переглянути': 'D:/Учеба/7_семестр/курсавая/tests/Копия ПА 1-6 к. 2022.xlsx',
              5: 'D:/Учеба/7_семестр/курсавая/tests/ФПМ РОЗРАХУНОК (2).xlsx', 'Переглянути0': 'D:/Учеба/7_семестр/курсавая/tests/ФПМ РОЗРАХУНОК (2).xlsx',
              'filetrue': False, 6: True, 7: '', 'dop_file': '', 'вибір_дисц_бакалавр_денна': '0.7',
              'вибір_дисц_магістр_денна': '0.67', 'вибір_дисц_бакалавр_заочна': '0.17', 'вибір_дисц_магістр_заочна': '0.16', 'вибір_дисц_бакалавр_вечірня': '0.64', 'екз': '4', 'пров_екз': '2', 'конс_пред_екз': '2', 'заліки': '2', 'студ_зал': '25', 'пот_конс_денна': '4', 'пот_конс_вечірня': '6', 'пот_конс_заочна': '8', 'пот_конс_дуальна': '10', 'академ_груп': '25', 'індивід_денна/вечірня': '4', 'індивід_заочна': '2', 'індивід_дуальна': '4', 'кр': '1', 'кп': '2', 'зах_кр': '1', 'зах_кп': '1', 'нав_практ1': '30', 'нав_практ2': '15', 'вир_практ1': '15', 'вир_практ2': '1', 'вир_пр_переддипломна': '1', 'студ_нав_практ': '15', 'студ_вир_практ': '15', 'атест_ЕК': '2', 'атест_екз_консультації': '8', 'квал_роб_керівництво1': '0', 'квал_роб_керівництво2_до_5к': '8', 'квал_роб_керівництво2_5_та_6к': '31', 'квал_роб_рецензування_до_5к': '2', 'квал_роб_рецензування_5_та_6к': '4', 8: 'Головна'}
    file = 'main_values.txt'
    file2 = 'coefficients.txt'
    # save_main_values(values)
    # save_coef(values, coeff=False)
    print(check_coef())
