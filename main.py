
import PySimpleGUI as sg
from excelreader import reader
from pathlib import Path


def main():
    open_window = False
    while True:
        layout_settings = [
            [sg.Text('К(вибір_дисц_бакалавр_денна):'),
             sg.InputText(0.7, size=(5, 1),
                          do_not_clear=True, key="вибір_дисц_бакалавр_денна")],
            [sg.Text('К(вибір_дисц_магістр_денна):'),
             sg.InputText(0.67, size=(5, 1),
                          do_not_clear=True, key="вибір_дисц_магістр_денна")],
            [sg.Text('К(вибір_дисц_бакалавр_заочна):'),
             sg.InputText(0.17, size=(5, 1),
                          do_not_clear=True, key="вибір_дисц_бакалавр_заочна")],
            [sg.Text('К(вибір_дисц_магістр_заочна):'),
             sg.InputText(0.16, size=(5, 1),
                          do_not_clear=True, key="вибір_дисц_магістр_заочна")],
            [sg.Text('К(вибір_дисц_бакалавр_вечірня):'),
             sg.InputText(0.64, size=(5, 1),
                          do_not_clear=True, key="вибір_дисц_бакалавр_вечірня")],
            [sg.Text('К(екз):'),
             sg.InputText(4, size=(5, 1),
                          do_not_clear=True, key="екз")],
            [sg.Text('К(пров_екз):'),
             sg.InputText(2, size=(5, 1),
                          do_not_clear=True, key="пров_екз")],
            [sg.Text('К(конс_пред_екз):'),
             sg.InputText(
                2, size=(5, 1), do_not_clear=True, key="конс_пред_екз")],
            [sg.Text('К(заліки):'),
             sg.InputText(2, size=(5, 1),
                          do_not_clear=True, key="заліки")],
            [sg.Text('К(студ_зал):'),
             sg.InputText(25, size=(5, 1),
                          do_not_clear=True, key="студ_зал")],
            [sg.Text('К(пот_конс_денна):'),
             sg.InputText(4, size=(5, 1),
                          do_not_clear=True, key="пот_конс_денна")],
            [sg.Text('К(пот_конс_вечірня):'),
             sg.InputText(6, size=(5, 1),
                          do_not_clear=True, key="пот_конс_вечірня")],
            [sg.Text('К(пот_конс_заочна):'),
             sg.InputText(8, size=(5, 1),
                          do_not_clear=True, key="пот_конс_заочна")],
            [sg.Text('К(пот_конс_дуальна):'),
             sg.InputText(10, size=(5, 1),
                          do_not_clear=True, key="пот_конс_дуальна")],
            [sg.Text('К(академ_груп):'),
             sg.InputText(25, size=(5, 1),
                          do_not_clear=True, key="академ_груп")],
            [sg.Text('К(індивід_денна/вечірня):'),
             sg.InputText(4, size=(5, 1),
                          do_not_clear=True, key="індивід_денна/вечірня")],
            [sg.Text('К(індивід_заочна):'),
             sg.InputText(2, size=(5, 1),
                          do_not_clear=True, key="індивід_заочна")],
            [sg.Text('К(індивід_дуальна):'),
             sg.InputText(4, size=(5, 1),
                          do_not_clear=True, key="індивід_дуальна")],
            [sg.Text('К(кр):'),
             sg.InputText(1, size=(5, 1),
                          do_not_clear=True, key="кр")],
            [sg.Text('К(кп):'),
             sg.InputText(2, size=(5, 1),
                          do_not_clear=True, key="кп")],
            [sg.Text('К(зах_кр):'),
             sg.InputText(1, size=(5, 1),
                          do_not_clear=True, key="зах_кр")],
            [sg.Text('К(зах_кп):'),
             sg.InputText(1, size=(5, 1),
                          do_not_clear=True, key="зах_кп")],
            [sg.Text('К(нав_практ1):'),
             sg.InputText(30, size=(5, 1),
                          do_not_clear=True, key="нав_практ1")],
            [sg.Text('К(нав_практ2):'),
             sg.InputText(15, size=(5, 1),
                          do_not_clear=True, key="нав_практ2")],
            [sg.Text('К(вир_практ1):'),
             sg.InputText(15, size=(5, 1),
                          do_not_clear=True, key="вир_практ1")],
            [sg.Text('К(вир_практ2):'),
             sg.InputText(1, size=(5, 1),
                          do_not_clear=True, key="вир_практ2")],
            [sg.Text('К(вир_пр_переддипломна):'),
             sg.InputText(1, size=(5, 1),
                          do_not_clear=True, key="вир_пр_переддипломна")],
            [sg.Text('К(студ_нав_практ):'),
             sg.InputText(15, size=(5, 1),
                          do_not_clear=True, key="студ_нав_практ")],
            [sg.Text('К(студ_вир_практ):'),
             sg.InputText(15, size=(5, 1),
                          do_not_clear=True, key="студ_вир_практ")],
            [sg.Text('К(атест_ЕК):'),
             sg.InputText(2, size=(5, 1),
                          do_not_clear=True, key="атест_ЕК")],
            [sg.Text('К(атест_екз_консультації):'),
             sg.InputText(8, size=(5, 1),
                          do_not_clear=True, key="атест_екз_консультації")],
            [sg.Text('К(квал_роб_керівництво1):'),
             sg.InputText(0, size=(5, 1),
                          do_not_clear=True, key="квал_роб_керівництво1")],
            [sg.Text('К(квал_роб_керівництво2_до_5к):'),
             sg.InputText(8, size=(5, 1),
                          do_not_clear=True, key="квал_роб_керівництво2_до_5к")],
            [sg.Text('К(квал_роб_керівництво2_5_та_6к):'),
             sg.InputText(31, size=(5, 1),
                          do_not_clear=True, key="квал_роб_керівництво2_5_та_6к")],
            [sg.Text('К(квал_роб_рецензування_до_5к):'),
             sg.InputText(2, size=(5, 1),
                          do_not_clear=True, key="квал_роб_рецензування_до_5к")],
            [sg.Text('К(квал_роб_рецензування_5_та_6к):'),
             sg.InputText(4, size=(5, 1),
                          do_not_clear=True, key="квал_роб_рецензування_5_та_6к")],

        ]
        layout_main = [[sg.Text('Середнє навчальне навантаження.')],
                       [sg.InputText(580)],
                       [sg.Text('Середня заробітна плата')],
                       [sg.InputText(17100)],
                       [sg.Text('Відрахування ЄСВ')],
                       [sg.InputText(1.22)],
                       [sg.Text('Певна величина')],
                       [sg.InputText(10590.41)],
                       [sg.Text('File1')],
                       [sg.InputText(
                       ), sg.FileBrowse(button_text="Переглянути", size=(10, 1), file_types=(("MIDI files", "*.xlsx"),))],
                       [sg.Text('File2')],
                       [sg.InputText(
                       ), sg.FileBrowse(button_text="Переглянути", size=(10, 1), file_types=(("MIDI files", "*.xlsx"),))],
                       [sg.Text('Доповнити файл?')],
                       [sg.Radio('Так',
                                 "RADIO1", default=False, key="filetrue"),
                       sg.Radio('Ні',
                                "RADIO1", default=True)],
                       [sg.InputText(
                       ), sg.FileBrowse(button_text="Переглянути", size=(10, 1), file_types=(("MIDI files", "*.xlsx"),), key="dop_file")],
                       [sg.Submit("Підтвердити"), sg.Cancel("Відмінити")],
                       ]
        layout = [
            [sg.Column(layout_settings, scrollable=True,  vertical_scroll_only=True, size_subsample_height=1.5, size_subsample_width=0.54)]]
        layout_main1 = [
            [sg.Column(layout_main, scrollable=True,  vertical_scroll_only=True, size_subsample_height=1, size_subsample_width=0.55)]]

        tabgrp = [[
            sg.TabGroup([[sg.Tab("Головна", layout_main1),
                        sg.Tab("Налаштування", layout)]])
        ]]
        if open_window == False:
            window = sg.Window(
                'Рентабельність спеціальності/факультету', tabgrp, icon="DNU_gerb2.ico", size=(500, 400))
            open_window = True

        event, values = window.read()
        if event == "Відмінити" or event == None:
            flag = 0
            window.close()
            break
        else:
            flag = 1
        if len(values) < 6:
            sg.popup('Ви ввели некоректні данні: ', v)
            break
        if flag == 1:
            flag = 0
            for k, v in values.items():
                if k == 4 or k == 5 or k == "Переглянути" or k == "Переглянути0" or k == 6 or k == 9 or k == 7 or k == 8 or k == "dop_file":
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
