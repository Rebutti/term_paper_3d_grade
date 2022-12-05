import PySimpleGUI as sg
from excelreader import reader
from pathlib import Path
from update_coefficients import save_coef, check_coef
# from screeninfo import get_monitors

# def get_monitor_size():
#     for m in get_monitors():
#         if m.is_primary == False:
#             return m


def main():
    open_window = False
    values = check_coef()
    while True:
        layout_settings = [
            [sg.Text('К(вибір_дисц_бакалавр_денна):'),
             sg.InputText(values["вибір_дисц_бакалавр_денна"], size=(5, 1),
                          do_not_clear=True, key="вибір_дисц_бакалавр_денна")],
            [sg.Text('К(вибір_дисц_магістр_денна):'),
             sg.InputText(values["вибір_дисц_магістр_денна"], size=(5, 1),
                          do_not_clear=True, key="вибір_дисц_магістр_денна")],
            [sg.Text('К(вибір_дисц_бакалавр_заочна):'),
             sg.InputText(values["вибір_дисц_бакалавр_заочна"], size=(5, 1),
                          do_not_clear=True, key="вибір_дисц_бакалавр_заочна")],
            [sg.Text('К(вибір_дисц_магістр_заочна):'),
             sg.InputText(values["вибір_дисц_магістр_заочна"], size=(5, 1),
                          do_not_clear=True, key="вибір_дисц_магістр_заочна")],
            [sg.Text('К(вибір_дисц_бакалавр_вечірня):'),
             sg.InputText(values["вибір_дисц_бакалавр_вечірня"], size=(5, 1),
                          do_not_clear=True, key="вибір_дисц_бакалавр_вечірня")],
            [sg.Text('К(екз):'),
             sg.InputText(values["екз"], size=(5, 1),
                          do_not_clear=True, key="екз")],
            [sg.Text('К(пров_екз):'),
             sg.InputText(values["пров_екз"], size=(5, 1),
                          do_not_clear=True, key="пров_екз")],
            [sg.Text('К(конс_пред_екз):'),
             sg.InputText(
                values["конс_пред_екз"], size=(5, 1), do_not_clear=True, key="конс_пред_екз")],
            [sg.Text('К(заліки):'),
             sg.InputText(values["заліки"], size=(5, 1),
                          do_not_clear=True, key="заліки")],
            [sg.Text('К(студ_зал):'),
             sg.InputText(values["студ_зал"], size=(5, 1),
                          do_not_clear=True, key="студ_зал")],
            [sg.Text('К(пот_конс_денна):'),
             sg.InputText(values["пот_конс_денна"], size=(5, 1),
                          do_not_clear=True, key="пот_конс_денна")],
            [sg.Text('К(пот_конс_вечірня):'),
             sg.InputText(values["пот_конс_вечірня"], size=(5, 1),
                          do_not_clear=True, key="пот_конс_вечірня")],
            [sg.Text('К(пот_конс_заочна):'),
             sg.InputText(values["пот_конс_заочна"], size=(5, 1),
                          do_not_clear=True, key="пот_конс_заочна")],
            [sg.Text('К(пот_конс_дуальна):'),
             sg.InputText(values["пот_конс_дуальна"], size=(5, 1),
                          do_not_clear=True, key="пот_конс_дуальна")],
            [sg.Text('К(академ_груп):'),
             sg.InputText(values["академ_груп"], size=(5, 1),
                          do_not_clear=True, key="академ_груп")],
            [sg.Text('К(індивід_денна/вечірня):'),
             sg.InputText(values["індивід_денна/вечірня"], size=(5, 1),
                          do_not_clear=True, key="індивід_денна/вечірня")],
            [sg.Text('К(індивід_заочна):'),
             sg.InputText(values["індивід_заочна"], size=(5, 1),
                          do_not_clear=True, key="індивід_заочна")],
            [sg.Text('К(індивід_дуальна):'),
             sg.InputText(values["індивід_дуальна"], size=(5, 1),
                          do_not_clear=True, key="індивід_дуальна")],
            [sg.Text('К(кр):'),
             sg.InputText(values["кр"], size=(5, 1),
                          do_not_clear=True, key="кр")],
            [sg.Text('К(кп):'),
             sg.InputText(values["кп"], size=(5, 1),
                          do_not_clear=True, key="кп")],
            [sg.Text('К(зах_кр):'),
             sg.InputText(values["зах_кр"], size=(5, 1),
                          do_not_clear=True, key="зах_кр")],
            [sg.Text('К(зах_кп):'),
             sg.InputText(values["зах_кп"], size=(5, 1),
                          do_not_clear=True, key="зах_кп")],
            [sg.Text('К(нав_практ1):'),
             sg.InputText(values["нав_практ1"], size=(5, 1),
                          do_not_clear=True, key="нав_практ1")],
            [sg.Text('К(нав_практ2):'),
             sg.InputText(values["нав_практ2"], size=(5, 1),
                          do_not_clear=True, key="нав_практ2")],
            [sg.Text('К(вир_практ1):'),
             sg.InputText(values["вир_практ1"], size=(5, 1),
                          do_not_clear=True, key="вир_практ1")],
            [sg.Text('К(вир_практ2):'),
             sg.InputText(values["вир_практ2"], size=(5, 1),
                          do_not_clear=True, key="вир_практ2")],
            [sg.Text('К(вир_пр_переддипломна):'),
             sg.InputText(values["вир_пр_переддипломна"], size=(5, 1),
                          do_not_clear=True, key="вир_пр_переддипломна")],
            [sg.Text('К(студ_нав_практ):'),
             sg.InputText(values["студ_нав_практ"], size=(5, 1),
                          do_not_clear=True, key="студ_нав_практ")],
            [sg.Text('К(студ_вир_практ):'),
             sg.InputText(values["студ_вир_практ"], size=(5, 1),
                          do_not_clear=True, key="студ_вир_практ")],
            [sg.Text('К(атест_ЕК):'),
             sg.InputText(values["атест_ЕК"], size=(5, 1),
                          do_not_clear=True, key="атест_ЕК")],
            [sg.Text('К(атест_екз_консультації):'),
             sg.InputText(values["атест_екз_консультації"], size=(5, 1),
                          do_not_clear=True, key="атест_екз_консультації")],
            [sg.Text('К(квал_роб_керівництво1):'),
             sg.InputText(values["квал_роб_керівництво1"], size=(5, 1),
                          do_not_clear=True, key="квал_роб_керівництво1")],
            [sg.Text('К(квал_роб_керівництво2_до_5к):'),
             sg.InputText(values["квал_роб_керівництво2_до_5к"], size=(5, 1),
                          do_not_clear=True, key="квал_роб_керівництво2_до_5к")],
            [sg.Text('К(квал_роб_керівництво2_5_та_6к):'),
             sg.InputText(values["квал_роб_керівництво2_5_та_6к"], size=(5, 1),
                          do_not_clear=True, key="квал_роб_керівництво2_5_та_6к")],
            [sg.Text('К(квал_роб_рецензування_до_5к):'),
             sg.InputText(values["квал_роб_рецензування_до_5к"], size=(5, 1),
                          do_not_clear=True, key="квал_роб_рецензування_до_5к")],
            [sg.Text('К(квал_роб_рецензування_5_та_6к):'),
             sg.InputText(values["квал_роб_рецензування_5_та_6к"], size=(5, 1),
                          do_not_clear=True, key="квал_роб_рецензування_5_та_6к")],
            [sg.Text('К(D1coeff):'),
             sg.InputText(values["D1coeff"], size=(5, 1),
                          do_not_clear=True, key="D1coeff")],
            [sg.Text('К(E1coeff):'),
             sg.InputText(values["E1coeff"], size=(5, 1),
                          do_not_clear=True, key="E1coeff")],
            [sg.Text('К(НПП_витрати):'),
             sg.InputText(values['НПП_витрати'], size=(5, 1),
                          do_not_clear=True, key='НПП_витрати')],
            [sg.Submit("Зберегти", key='save_coef')],

        ]
        layout_main = [[sg.Text('Середнє навчальне навантаження.')],
                       [sg.InputText(values[0])],
                       [sg.Text('Середня заробітна плата')],
                       [sg.InputText(values[1])],
                       [sg.Text('Відрахування ЄСВ')],
                       [sg.InputText(values[2])],
                       [sg.Text('Певна величина')],
                       [sg.InputText(values[3])],
                       [sg.Text('Навчальний план')],
                       [sg.InputText(size=(31, 1)
                                     ), sg.FileBrowse(button_text="Переглянути", size=(10, 1), file_types=(("MIDI files", "*.xlsx"),))],
                       [sg.Text('Розрахунок')],
                       [sg.InputText(size=(31, 1)
                                     ), sg.FileBrowse(button_text="Переглянути", size=(10, 1), file_types=(("MIDI files", "*.xlsx"),))],
                       [sg.Text('Доповнити файл?')],
                       [sg.Radio('Так',
                                 "RADIO1", default=False, key="filetrue"),
                       sg.Radio('Ні',
                                "RADIO1", default=True)],
                       [sg.InputText(size=(31, 1)
                                     ), sg.FileBrowse(button_text="Переглянути", size=(10, 1), file_types=(("MIDI files", "*.xlsx"),), key="dop_file")],
                       [sg.Submit("Підтвердити"), sg.Cancel("Відмінити")],
                       ]
        layout_vart = [
                       [sg.Text('Навчальний план')],
                       [sg.InputText(size=(31, 1)
                                     ), sg.FileBrowse(button_text="Переглянути", size=(10, 1), file_types=(("MIDI files", "*.xlsx"),), key="plan_vart")],
                       [sg.Submit("Підтвердити", key='count_vartist'), sg.Cancel("Відмінити", key="cancel")],
                       ]
        layout_set = [
            [sg.Column(layout_settings, scrollable=True,  vertical_scroll_only=True, size_subsample_height=1.5, size_subsample_width=0.54)]]
        layout_main1 = [
            [sg.Column(layout_main, scrollable=True,  vertical_scroll_only=True, size_subsample_height=1, size_subsample_width=0.55)]]
        layout_vartist = [
            [sg.Column(layout_vart, scrollable=True,  vertical_scroll_only=True, size_subsample_height=1, size_subsample_width=0.55)]]

        tabgrp = [[
            sg.TabGroup([[sg.Tab("Головна", layout_main1),
                        sg.Tab("Налаштування", layout_set),
                        sg.Tab("Вартість", layout_vartist)]])
        ]]
        if open_window == False:
            window = sg.Window(
                'Рентабельність спеціальності/факультету', tabgrp, icon="DNU_gerb2.ico", size=(600, 400)).Finalize()
            open_window = True

        event, values = window.read()
        # print(values, event)
        if event == "Відмінити" or event == None or event == "cancel":
            flag = 0
            window.close()
            break
        elif event == 'save_coef':
            save_coef(values, False)
            save_coef(values)
            sg.popup('Нові дані збережені!')
            continue
        else:
            flag = 1
        if len(values) < 6:
            sg.popup('Ви ввели некоректні данні: ', v)
            break
        if flag == 1:
            flag = 0
            for k, v in values.items():
                if k == 4 or k == 5 or k == "Переглянути" or k == "Переглянути0" or k == 6 or k == 9 or k == 7 or k == 8 or k == "dop_file" or k == "plan_vart":
                    flag = 1
                    if k == 4 or k == 5 or k == "Переглянути":
                        a = Path(v)
                        if a.is_file():
                            continue
                        else:
                            sg.popup(
                                'Ви ввели не правильний шлях до файлу: ', v)
                            flag = 0
                            break
                else:
                    try:
                        v = (float(v))
                    except:
                        sg.popup('Ви ввели некоректні данні: ', v)
                        flag = 0
                        break
        if flag != 0:
            reader(values)
            sg.popup('Програма закінчила свою роботу!')


if __name__ == "__main__":
    main()
