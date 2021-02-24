import PySimpleGUI as sg  # интерфейс
import docx  # работа с вордом
import re  # регулярные выражения


# функции, использованные в программе

def numbers(string_with_nums):
    """
    Функция принимает на входе строку
    и выдает список цифр, содержащихся в ней 
    """
    return list(map(float, re.findall(r'(?<!\d)-?\d*[.,]?\d+', string_with_nums.replace(',', '.'))))


def more(need_list, have_list):
    """
    Проверяет условие need_num_list[0] <= have_num_list[0]

    Принимает на входе числа, полученные функцией numbers
    """

    if need_list[0] <= have_list[0]:
        return f'OK, строка: {wordDoc.tables[1].rows[string].cells[0].text}'

    return f'Число больше чем нужно, строка: {wordDoc.tables[1].rows[string].cells[0].text}'


def low(need_list, have_list):
    """
    Проверяет условие need_num_list[0] >= have_num_list[0]

    Принимает на входе числа, полученные функцией numbers
    """

    if need_list[0] >= have_list[0]:
        return f'OK, строка: {wordDoc.tables[1].rows[string].cells[0].text}'

    return f'Число меньше чем нужно, строка: {wordDoc.tables[1].rows[string].cells[0].text}'


def in_interval_equally(need_list, have_list):
    """
    Проверяет условие min(need_num_list) <= have_num_list[0] <= max(need_num_list)
    а так же наличие ошибок в указанных условиях

    Принимает на входе числа, полученные функцией numbers
    """

    if len(need_list) == 2:

        if min(need_list) == need_list[0] and max(need_list) == need_list[1]:
            if min(need_list) <= have_list[0] <= max(need_list):
                return f'OK, строка: {wordDoc.tables[1].rows[string].cells[0].text}'

            return f'Число не подходит под условия "не менее чем и не более чем", ' \
                   f'строка: {wordDoc.tables[1].rows[string].cells[0].text} '

        else:
            if need_list[1] <= have_list[0] <= need_list[0]:
                return f'OK, но возможна ошибка в условии, лучше проверить, ' \
                       f'строка: {wordDoc.tables[1].rows[string].cells[0].text}'

            return f'Число не подходит под условия "не менее чем и не более чем", ' \
                   f'строка: {wordDoc.tables[1].rows[string].cells[0].text}'

    elif len(need_list) == len(have_list) == 1:

        if need_list[0] == have_list[0]:
            return f'OK, но возможна ошибка в условии, лучше проверить, ' \
                   f'строка: {wordDoc.tables[1].rows[string].cells[0].text}'

        return f'Число не подходит под условия "не менее чем и не более чем", ' \
               f'строка: {wordDoc.tables[1].rows[string].cells[0].text}'

    return f'Число не подходит под условия "не менее чем и не более чем", ' \
           f'строка: {wordDoc.tables[1].rows[string].cells[0].text}'


def in_interval_not_equally(need_list, have_list):
    """
    Проверяет условие min(need_num_list) < have_num_list[0] < max(need_num_list)
    а так же наличие ошибок в указанных условиях

    Принимает на входе числа, полученные функцией numbers
    """

    if min(need_list) == need_list[0] and max(need_list) == need_list[1]:
        if min(need_list) < have_list[0] < max(need_list):
            return f'OK, строка: {wordDoc.tables[1].rows[string].cells[0].text}'

        return f'Число не подходит под условия "не менее чем и не более чем", ' \
               f'строка: {wordDoc.tables[1].rows[string].cells[0].text} '

    else:
        if need_list[1] < have_list[0] < need_list[0]:
            return f'OK, но возможна ошибка в условии, лучше проверить, ' \
                   f'строка: {wordDoc.tables[1].rows[string].cells[0].text}'

        return f'Число не подходит под условия "не менее чем и не более чем", ' \
               f'строка: {wordDoc.tables[1].rows[string].cells[0].text}'


def in_interval_biggest(need_list, have_list):
    """
    Проверяет условие min(need_num_list) < have_num_list[0] <= max(need_num_list)
    а так же наличие ошибок в указанных условиях

    Принимает на входе числа, полученные функцией numbers
    """

    if min(need_list) == need_list[0] and max(need_list) == need_list[1]:
        if min(need_list) < have_list[0] <= max(need_list):
            return f'OK, строка: {wordDoc.tables[1].rows[string].cells[0].text}'

        return f'Число не подходит под условия "< & <=", строка: {wordDoc.tables[1].rows[string].cells[0].text}'

    else:
        if need_list[1] <= have_list[0] <= need_list[0]:
            return f'OK, но возможна ошибка в условии, лучше проверить, ' \
                   f'строка: {wordDoc.tables[1].rows[string].cells[0].text}'

        return f'Число не подходит под условия "< & <=", строка: {wordDoc.tables[1].rows[string].cells[0].text}'


def in_interval_lowest(need_list, have_list):
    """
    Проверяет условие min(need_num_list) <= have_num_list[0] < max(need_num_list)
    а так же наличие ошибок в указанных условиях

    Принимает на входе числа, полученные функцией numbers
    """

    if min(need_list) == need_list[0] and max(need_list) == need_list[1]:
        if min(need_list) <= have_list[0] < max(need_list):
            return f'OK, строка: {wordDoc.tables[1].rows[string].cells[0].text}'

        return f'Число не подходит под условия "<= & <", строка: {wordDoc.tables[1].rows[string].cells[0].text}'

    else:
        if need_list[1] <= have_list[0] <= need_list[0]:
            return f'OK, но возможна ошибка в условии, лучше проверить, ' \
                   f'строка: {wordDoc.tables[1].rows[string].cells[0].text}'

        return f'Число не подходит под условия "<= & <", строка: {wordDoc.tables[1].rows[string].cells[0].text}'


def lowest(need_list, have_list):
    """
    Проверяет условие need_num_list[0] > have_num_list[0]

    Принимает на входе числа, полученные функцией numbers
    """
    if need_list[0] > have_list[0]:
        return f'OK, строка: {wordDoc.tables[1].rows[string].cells[0].text}'

    return f'Число больше чем нужно, строка: {wordDoc.tables[1].rows[string].cells[0].text}'


def biggest(need_list, have_list):
    """
    Проверяет условие need_num_list[0] < have_num_list[0]

    Принимает на входе числа, полученные функцией numbers
    """

    if need_list[0] < have_list[0]:
        return f'OK, строка: {wordDoc.tables[1].rows[string].cells[0].text}'

    return f'Число меньше чем нужно, строка: {wordDoc.tables[1].rows[string].cells[0].text}'


def only_numbers(need_list, have_list):
    """
    Функция проверяет находятся ли в условии только числа (или интервал значений)

    Принимает на входе числа, полученные функцией numbers
    """
    if len(have_list) == 1 and len(need_list) == 2:
        if need_list[0] <= have_list[0] <= need_list[1]:
            return f'OK, строка: {wordDoc.tables[1].rows[string].cells[0].text}'

        return f'Надо проверить, строка: {wordDoc.tables[1].rows[string].cells[0].text}'

    return f'Надо проверить, строка: {wordDoc.tables[1].rows[string].cells[0].text}'


def string_conditions(need_text, have_text):
    """
    Проверяет текстовое условие

    Принимает на входе числа, полученные функцией numbers
    """
    need_conditions_list = need_text.lower().replace(',', '').split()
    have_conditions_list = have_text.lower().replace(',', '').split()
    if 'или' in need_conditions_list and 'и' not in need_conditions_list:
        result = []
        for word1 in need_conditions_list:
            for word2 in have_conditions_list:
                if word1 == word2:
                    result.append(word1)
                    break

        if set(result) == set(have_conditions_list):
            return f'OK, строка: {wordDoc.tables[1].rows[string].cells[0].text}'

        elif len((set(result)) - set(have_conditions_list)) >= 1:
            return f'Возможна опечатка в текстовом условии, строка: {wordDoc.tables[1].rows[string].cells[0].text}'

        return f'Текстовое условие не выполнено, строка: {wordDoc.tables[1].rows[string].cells[0].text}'

    elif 'и' in need_conditions_list:
        return f'Надо проверить, текстовое условие "И" строка: {wordDoc.tables[1].rows[string].cells[0].text}'

    return f'Надо проверить, строка: {wordDoc.tables[1].rows[string].cells[0].text}'


# интерфейс программы

layout = [
    [sg.Text('Файл заявки'), sg.InputText(), sg.FileBrowse('Указать путь к файлу')],
    [sg.Output(size=(88, 20))],
    [sg.Submit('Старт'), sg.Cancel('Отмена')]
]
window = sg.Window('СУПЕР ПРОВЕРЯЛЬЩИК ЗАЯВОК 3000', layout, background_color='green')
while True:  # The Event Loop
    event, values = window.read()
    # print(event, values) #debug
    if event in (None, 'Exit', 'Cancel', 'Отмена'):
        break
    if event == 'Submit' or event == 'Старт':
        # сама программа

        wordDoc = docx.Document(values[0])

        table_size = len(wordDoc.tables[1].rows)  # узнаем сколько во второй таблице в документе строк

        print('Начинаю проверку заявки. Строк к проверке:', table_size)

        for string in range(1, table_size):
            # заносим в переменные текст из ячейки с условием (need) и тем что у нас есть (have)
            need = wordDoc.tables[1].rows[string].cells[2].text
            have = wordDoc.tables[1].rows[string].cells[3].text

            # выделяем числа из выбранных ячеек в отдельные списки при помощи функции numbers
            need_num_list = numbers(need)
            have_num_list = numbers(have)

            if (any(word in wordDoc.tables[1].rows[string].cells[1].text.lower()
                    for word in ('диап', 'диоп', 'интер'))):
                # проверяем, есть ли какие-то диапазоны в описании, если есть - явно на них указываем
                print(
                    f'Указан какой-то интервал, необходимо согласовать с заказчиком, '
                    f'строка: {wordDoc.tables[1].rows[string].cells[0].text}')

            if need.lower().replace(' ', '') == have.lower().replace(' ', ''):
                # если ячейки равны друг другу, то продолжаем дальше, т.к. нас это устраивает
                print(f'OK, строка: {wordDoc.tables[1].rows[string].cells[0].text}')

            elif need.lower().replace(' ', '') == '' or have.lower().replace(' ', '') == '':
                # если одна из ячеек пустая - сообщаем об этом
                print(f'Пустая ячейка, строка: {wordDoc.tables[1].rows[string].cells[0].text}')

            elif (any(word in need.lower().replace(' ', '') for word in ('неменее', '≥', 'неменьше'))
                  and all(word not in need.lower().replace(' ', '') for word in ('неболее', '≤', 'небольше', '<'))):

                # тут проверяем условие "не менее чем" (need <= have)
                print(more(need_num_list, have_num_list))

            elif (any(word in need.lower().replace(' ', '') for word in ('неболее', '≤', 'небольше'))
                  and all(word not in need.lower().replace(' ', '') for word in ('неменее', '≥', 'неменьше', '>'))):

                # тут проверяем условие "не более чем" (need >= have)
                print(low(need_num_list, have_num_list))

            elif (any(word in need.lower().replace(' ', '') for word in ('неболее', '≤', 'небольше'))
                  and any(word in need.lower().replace(' ', '') for word in ('неменее', '≥', 'неменьше'))):

                # тут проверяем условие "не менее чем и не более чем" (need1 <= have <= need2)
                print(in_interval_equally(need_num_list, have_num_list))

            elif (any(word in need.lower().replace(' ', '') for word in ('более', '>', 'больше'))
                  and any(word in need.lower().replace(' ', '') for word in ('неболее', '≤', 'небольше'))):

                # тут проверяем условие "< & <=" (need1 < have <= need2)
                print(in_interval_biggest(need_num_list, have_num_list))

            elif (any(word in need.lower().replace(' ', '') for word in ('менее', '<', 'меньше'))
                  and any(word in need.lower().replace(' ', '') for word in ('неменее', '≥', 'неменьше'))):

                # тут проверяем условие "<= & <" (need1 < have <= need2)
                print(in_interval_lowest(need_num_list, have_num_list))

            elif (any(word in need.lower().replace(' ', '') for word in ('менее', '<', 'меньше'))
                  and all(word not in need.lower().replace(' ', '') for word in ('более', '>', 'больше', '≥'))):

                # тут проверяем условие "менее чем" (need > have)
                print(lowest(need_num_list, have_num_list))

            elif (any(word in need.lower().replace(' ', '') for word in ('более', '>', 'больше'))
                  and all(word not in need.lower().replace(' ', '') for word in ('менее', '<', 'меньше', '≤'))):

                # тут проверяем условие "более чем" (need < have)
                print(biggest(need_num_list, have_num_list))

            elif (any(word in need.lower().replace(' ', '') for word in ('более', '>', 'больше'))
                  and any(word in need.lower().replace(' ', '') for word in ('менее', '<', 'меньше'))):

                # тут проверяем условие "более чем и менее чем" (need1 < have < need2)
                print(in_interval_not_equally(need_num_list, have_num_list))

            elif 0 < len(need_num_list) <= 2:
                # проверяем заданы ли в условии только числа (или их интервал)
                print(only_numbers(need_num_list, have_num_list))

            elif len(set(list(need.lower())) - set(list(have.lower()))) >= 1:
                # проверяем чисто текстовые условия, а так же возможные ошибки или опечатки
                if ' и ' in need.lower() or ' или ' in need.lower():
                    print(string_conditions(need, have))
                else:
                    print(f'Возможна ошибка или опечатка, строка: {wordDoc.tables[1].rows[string].cells[0].text}')

            elif len(need.lower()) > len(have.lower()):
                # проверяем чисто текстовые условия с "И" и/или "ИЛИ"
                print(string_conditions(need, have))

            else:
                print(f'Я не знаю что делать с этой строкой: {wordDoc.tables[1].rows[string].cells[0].text}')

        print('Проверка файла завершена. Количество проверенных строк:', table_size)

    else:
        print('Что-то пошло не так')
