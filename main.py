import os
import time
import re
# import platform
import json
import concurrent.futures as cf
from multiprocessing import cpu_count as cc

import undetected_chromedriver as uc
import pandas as pd
import openpyxl
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
from fake_useragent import UserAgent as UA



data_dir = os.path.join(os.getcwd(), 'result_data')
xlsx_dir = os.path.join(os.getcwd(), 'xlsx_data')
json_dir = os.path.join(os.getcwd(), 'json_data')

if not os.path.exists(data_dir):
    os.mkdir(data_dir)

params_sites_search = [
    'https://www.regard.ru/catalog?search=',
    'https://www.onlinetrade.ru/sitesearch.html?query=',
    'https://www.novo-market.ru/search/?q=',
    'https://www.dns-shop.ru/search/?q='
]

ua = UA()

alphabeth = ['A', 'B']

class DriverInitialize:

    def __new__(cls, headless: bool = True):
        options = uc.ChromeOptions()
        # options.add_argument('--disable-blink-features=AutomationControlled')
        
        if headless:
            options.add_argument('--headless')

        driver = uc.Chrome(options=options)
        return driver


def get_data_from_xlsx(filename: str):
    data = pd.read_excel(f'{filename}')
    sku_list = data['sku'].tolist()
    part_list = data['part'].to_list()
    vendor_list = data['Вендор'].to_list()
    name_list = data['Наименование'].to_list()

    return sku_list, part_list, name_list, vendor_list


def get_char(name_char, value_char, site, monitors: bool, mice: bool, ddr=bool, cartridges=bool):
    each_one_char = dict()
    if monitors:
        if name_char == 'Разрешение экрана' or name_char == 'Максимальное разрешение':
            if site == 'https://www.regard.ru/catalog?search=' or site == 'https://www.novo-market.ru/search/?q=':
                try:
                    each_one_char.update(
                        
                        {
                            'Соотношение сторон': {
                                'value': value_char.replace('пикс.', '').split(' ')[1].replace('(', '').replace(')', '').strip(),
                                'char_name': 'Соотношение сторон',
                                'vid_name': 'Обычный',
                                'unit_name': '(без наименования)'
                            }
                        }
                    )
                    each_one_char.update(
                        {   
                            'Максимальное разрешение': {
                                'value': value_char.replace('пикс.', '').split(' ')[0].strip(),
                                'char_name': 'Максимальное разрешение',
                                'vid_name': 'Обычный',
                                'unit_name': '(без наименования)'
                            }
                        }
                    )
                except:
                    pass
            elif site == 'https://www.dns-shop.ru/search/?q=' or site == 'https://www.onlinetrade.ru/sitesearch.html?query=':
                try:
                    each_one_char.update(
                        
                        {
                            'Максимальное разрешение': {
                                'value': value_char.replace('пикс.', '').split(' ')[0],
                                'char_name': 'Максимальное разрешение',
                                'vid_name': 'Обычный',
                                'unit_name': '(без наименования)'
                            }
                        }
                    )
                except:
                    pass
        elif name_char == 'Диагональ' or name_char == 'Диагональ экрана' or name_char == 'Диагональ экрана (дюйм)' or name_char == 'Размер экрана':
            try:
                each_one_char.update(
                    {
                        'Диагональ': {
                            'value': value_char.strip().replace(' ', ''),
                            'char_name': 'Диагональ',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Поддержка HDR':
            if value_char == 'да' or value_char == 'Да':
                each_one_char.update(
                    {   
                        'Поддержка HDR': {
                            'value': 'есть',
                            'char_name': 'Поддержка HDR',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            else:
                each_one_char.update(
                    {   
                        'Поддержка HDR': {
                            'value': value_char,
                            'char_name': 'Поддержка HDR',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
        elif name_char == 'Размеры (ШхВхГ)' or name_char == 'Размеры (Ш x Г x В), мм' or name_char == 'Размеры':
            try:
                each_one_char.update(
                    {
                        'Ширина': {
                            'value': value_char.split('x')[0].strip().replace('мм', '').strip(),
                            'char_name': 'Ширина',
                            'vid_name': 'Обычный',
                            'unit_name': 'миллиметр'
                        }
                    }
                )
                each_one_char.update(
                    {
                        'Высота': {
                            'value': value_char.split('x')[1].strip().replace('мм', '').strip(),
                            'char_name': 'Высота',
                            'vid_name': 'Обычный',
                            'unit_name': 'миллиметр'
                        }
                    }
                )
                each_one_char.update(
                    {
                        'Глубина': {
                            'value': value_char.split('x')[2].replace('мм', '').strip(),
                            'char_name': 'Глубина',
                            'vid_name': 'Обычный',
                            'unit_name': 'миллиметр'
                        }
                    }
                )
            except IndexError:
                each_one_char.update(
                    {
                        'Ширина': {
                            'value': value_char.split('х')[0].strip().replace('мм', '').strip(),
                            'char_name': 'Ширина',
                            'vid_name': 'Обычный',
                            'unit_name': 'миллиметр'
                        }
                    }
                )
                each_one_char.update(
                    {
                        'Высота': {
                            'value': value_char.split('х')[1].strip().replace('мм', '').strip(),
                            'char_name': 'Высота',
                            'vid_name': 'Обычный',
                            'unit_name': 'миллиметр'
                        }
                    }
                )
                each_one_char.update(
                    {
                        'Глубина': {
                            'value': value_char.split('х')[2].replace('мм', '').strip(),
                            'char_name': 'Глубина',
                            'vid_name': 'Обычный',
                            'unit_name': 'миллиметр'
                        }
                    }
                )
        elif name_char == 'Вес':
            try:
                each_one_char.update(
                    {
                        'Вес': {
                            'value': value_char.strip().replace('кг&nbsp;', '').replace('кг ', '').replace('кг', '').strip(),
                            'char_name': 'Вес',
                            'vid_name': 'Обычный',
                            'unit_name': 'килограмм'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Широкоформатный монитор':
            try:
                each_one_char.update(
                    {
                        'Широкоформатный': {
                            'value': value_char,
                            'char_name': 'Широкоформатный',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Встроенные динамики' or name_char == 'Мощность динамиков':
            try:
                if len(value_char) <= 20:
                    try:
                        count_char_dyn = value_char.strip().replace('Вт', '').replace(' ', '').split('(')[-1].replace(')', '').split('x')[0]
                        power_char_dyn = value_char.strip().replace('Вт', '').replace(' ', '').split('(')[-1].replace(')', '').split('x')[1]       
                        each_one_char.update(
                            {
                                'Мощность динамиков (на канал)': {
                                    'value': power_char_dyn,
                                    'char_name': 'Мощность динамиков (на канал)',
                                    'vid_name': 'Обычный',
                                    'unit_name': 'Вт'
                                }
                            }
                        )
                        each_one_char.update(
                            {
                                'Количество встроенных динамиков': {
                                    'value': count_char_dyn,
                                    'char_name': 'Количество встроенных динамиков',
                                    'vid_name': 'Обычный',
                                    'unit_name': 'Вт'
                                }
                            }
                        )
                    except IndexError:
                        try:
                            count_char_dyn = value_char.strip().replace('Вт', '').replace(' ', '').split('(')[-1].replace(')', '').split('х')[0]
                            power_char_dyn = value_char.strip().replace('Вт', '').replace(' ', '').split('(')[-1].replace(')', '').split('х')[1]  
                            each_one_char.update(
                                {
                                    'Мощность динамиков (на канал)': {
                                        'value': power_char_dyn,
                                        'char_name': 'Мощность динамиков (на канал)',
                                        'vid_name': 'Обычный',
                                        'unit_name': 'Вт'
                                    }
                                }
                            )
                            each_one_char.update(
                                {
                                    'Количество встроенных динамиков': {
                                        'value': count_char_dyn,
                                        'char_name': 'Количество встроенных динамиков',
                                        'vid_name': 'Обычный',
                                        'unit_name': 'Вт'
                                    }
                                }
                            )
                        except IndexError:
                            count_char_dyn = value_char.strip().replace('Вт', '').replace(' ', '').split('(')[-1].replace(')', '').split('+')[0].split('х')[0]
                            power_char_dyn = value_char.strip().replace('Вт', '').replace(' ', '').split('(')[-1].replace(')', '').split('+')[0].split('х')[1]
                            each_one_char.update(
                                {
                                    'Мощность динамиков (на канал)': {
                                        'value': power_char_dyn,
                                        'char_name': 'Мощность динамиков (на канал)',
                                        'vid_name': 'Обычный',
                                        'unit_name': 'Вт'
                                    }
                                }
                            )
                            each_one_char.update(
                                {
                                    'Количество встроенных динамиков': {
                                        'value': count_char_dyn,
                                        'char_name': 'Количество встроенных динамиков',
                                        'vid_name': 'Обычный',
                                        'unit_name': 'Вт'
                                    }
                                }
                            )
            except:
                pass
        elif name_char == 'Разъёмы' or name_char == 'Видео разъемы' or name_char == 'Порты и разъемы' or name_char == 'Интерфейсы' or name_char == 'Другие разъемы':
            try:
                each_one_char.update(
                    {
                        'Разъёмы': {
                            'value': value_char.split(', '),
                            'char_name': 'Разъёмы',
                            'vid_name': 'несколько из',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Угол обзора по горизонтали':
            try:
                each_one_char.update(
                    {
                        'Гориз. область обзора, градусов': {
                            'value': value_char.replace('°', '').strip(),
                            'char_name': 'Гориз. область обзора, градусов',
                            'vid_name': 'Обычный',
                            'unit_name': 'градус'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Угол обзора по вертикали':
            try:
                each_one_char.update(
                    {
                        'Верт. область обзора, градусов': {
                            'value': value_char.replace('°', '').strip(),
                            'char_name': 'Верт. область обзора, градусов',
                            'vid_name': 'Обычный',
                            'unit_name': 'градус'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Переменная частота обновления' or name_char == 'Технология динамического обновления экрана' or name_char == 'Синхронизации кадров':
            try:
                each_one_char.update(
                    {
                        'Переменная частота обновления': {
                            'value': value_char.strip(),
                            'char_name': 'Переменная частота обновления',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'USB-хаб' or name_char == 'USB-концентратор':
            try:
                each_one_char.update(
                    {
                        'USB-концентратор': {
                            'value': value_char.split(',')[0].strip(),
                            'char_name': 'USB-концентратор',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
                each_one_char.update(
                    {   'Интерфейсы': {
                            'value': value_char.split(',')[1].strip(),
                            'char_name': 'Интерфейсы',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Время отклика':
            try:
                each_one_char.update(
                    {
                        'Время отклика': {
                            'value': value_char.split(' ')[0].replace('&nbsp;', '').strip(),
                            'char_name': 'Время отклика',
                            'vid_name': 'Обычный',
                            'unit_name': 'мс'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Контрастность':
            try:
                each_one_char.update(
                    {
                        'Контрастность': {
                            'value': value_char.strip().replace(' ', ''),
                            'char_name': 'Контрастность',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Динамическая контрастность':
            try:
                each_one_char.update(
                    {
                        'Динамическая контрастность': {
                            'value': value_char.strip().replace(' ', '').replace('М', '000000').strip(),
                            'char_name': 'Динамическая контрастность',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Яркость' or name_char == 'Яркость экрана' or name_char == 'Яркость, кд/м2':
            try:
                each_one_char.update(
                    {
                        'Яркость, кд/м2': {
                            'value': value_char.replace('кд/м2', '').replace('Кд/м²', '').replace('кд/кв.м', '').strip(),
                            'char_name': 'Яркость, кд/м2',
                            'vid_name': 'Обычный',
                            'unit_name': 'кд/м2'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Поверхность экрана' or name_char == 'Покрытие экрана':
            try:
                if value_char == 'матовая' or value_char == 'Матовая':
                    each_one_char.update(
                        {
                            'Поверхность экрана': {
                                'value': 'матовое',
                                'char_name': 'Поверхность экрана',
                                'vid_name': 'Обычный',
                                'unit_name': '(без наименования)'
                            }
                        }
                    )
                else:
                    each_one_char.update(
                        {
                            'Поверхность экрана': {
                                'value': value_char.strip(),
                                'char_name': 'Поверхность экрана',
                                'vid_name': 'Обычный',
                                'unit_name': '(без наименования)'
                            }
                        }
                    )
            except:
                pass
        elif name_char == 'Подсветка матрицы' or name_char == 'Подсветка':
            try:
                each_one_char.update(
                    {
                        'Подсветка': {
                            'value': value_char.strip(),
                            'char_name': 'Подсветка',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Тип матрицы' or name_char == 'Технология изготовления матрицы':
            try:
                each_one_char.update(
                    {
                        'Тип матрицы экрана': {
                            'value': value_char.strip(),
                            'char_name': 'Тип матрицы экрана',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Сенсорный экран':
            try:
                if value_char == 'да':
                    each_one_char.update(
                        {
                            'Сенсорный экран': {
                                'value': 'есть',
                                'char_name': 'Сенсорный экран',
                                'vid_name': 'Обычный',
                                'unit_name': '(без наименования)'
                            }
                        }
                    )
                else:
                    each_one_char.update(
                        {
                            'Сенсорный экран': {
                                'value': value_char.strip(),
                                'char_name': 'Сенсорный экран',
                                'vid_name': 'Обычный',
                                'unit_name': '(без наименования)'
                            }
                        }
                    )

            except:
                pass
        elif name_char == 'Потребляемая мощность при работе':
            try:
                each_one_char.update(
                    {
                        'Потребляемая мощность при работе': {
                            'value': value_char.replace('Вт').strip(),
                            'char_name': 'Потребляемая мощность при работе',
                            'vid_name': 'Обычный',
                            'unit_name': 'Вт'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Потребляемая мощность в режиме ожидания':
            try:
                each_one_char.update(
                    {
                        'Потребляемая мощность в режиме ожидания': {
                            'value': value_char.replace('Вт').strip(),
                            'char_name': 'Потребляемая мощность в режиме ожидания',
                            'vid_name': 'Обычный',
                            'unit_name': 'Вт'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Потребляемая мощность в спящем режиме':
            try:
                each_one_char.update(
                    {
                        'Потребляемая мощность в спящем режиме': {
                            'value': value_char.replace('Вт').strip(),
                            'char_name': 'Потребляемая мощность в спящем режиме',
                            'vid_name': 'Обычный',
                            'unit_name': 'Вт'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Блок питания' or name_char == 'Расположение блока питания':
            try:
                each_one_char.update(
                    {
                        'Блок питания': {
                            'value': value_char.strip(),
                            'char_name': 'Блок питания',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Поворот на 90 градусов' or name_char == 'Поворот на 90° (портретный режим)':
            try:
                if value_char == 'да' or value_char == 'Да':
                    each_one_char.update(
                        {
                            'Поворот на 90 градусов': {
                                'value': 'есть',
                                'char_name': 'Поворот на 90 градусов',
                                'vid_name': 'Обычный',
                                'unit_name': '(без наименования)'    
                            }
                        }
                    )
                else:
                    each_one_char.update(
                        {
                            'Поворот на 90 градусов': {
                                'value': value_char,
                                'char_name': 'Поворот на 90 градусов',
                                'vid_name': 'Обычный',
                                'unit_name': '(без наименования)'    
                            }
                        }
                    ) 
            except:
                pass
        elif name_char == 'Частота обновления кадров' or name_char == 'Частота обновления' or name_char == 'Максимальная частота обновления экрана':
            try:
                each_one_char.update(
                    {
                        'Макс. частота обновления кадров': {
                            'value': value_char.strip(),
                            'char_name': 'Макс. частота обновления кадров',
                            'vid_name': 'Обычный',
                            'unit_name': 'Гц'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Крепление на стену (VESA)' or name_char == 'Размер VESA' or name_char == 'Наличие крепления VESA' or name_char == 'Размер крепления VESA':
            try:
                if value_char == 'есть':
                    each_one_char.update(
                        {
                            'Крепление на стену (VESA)': {
                                'value': value_char.strip(),
                                'char_name': 'Крепление на стену (VESA)',
                                'vid_name': 'Обычный',
                                'unit_name': '(без наименования)'
                            }
                        }
                    )
                elif value_char == 'отсутствует' or value_char == 'н.д.' or value_char == 'нет':
                    pass
                else:
                    each_one_char.update(
                        {
                            'Крепление на стену (VESA)': {
                                'value': f'есть, {value_char.replace("мм", "").replace(" ", "").strip()}',
                                'char_name': 'Крепление на стену (VESA)',
                                'vid_name': 'Обычный',
                                'unit_name': '(без наименования)'
                            }
                        }
                    )
            except:
                pass
        elif name_char == 'Регулировка по высоте':
            try:
                each_one_char.update(
                    {
                        'Регулировка по высоте': {
                            'value': 'есть',
                            'char_name': 'Регулировка по высоте',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Максимальное количество цветов':
            if len(value_char.split(',')[0]) == 1 or len(value_char.split(',')[0]) == 2:
                try:
                    each_one_char.update(
                        {
                            'Максимальное количество цветов': {
                                'value': value_char.strip().replace(',', '.'),
                                'char_name': 'Максимальное количество цветов',
                                'vid_name': 'один из',
                                'unit_name': '(без наименования)'
                            }
                        }
                    )
                except:
                    pass
            else:
                try:
                    each_one_char.update(
                        {
                            'Максимальное количество цветов': {
                                'value': value_char.split(',')[0].replace('/', '.'),
                                'char_name': 'Максимальное количество цветов',
                                'vid_name': 'один из',
                                'unit_name': '(без наименования)'
                            }
                        }
                    )
                except:
                    pass
        elif name_char == 'Веб-камера':
            try:
                try:
                    each_one_char.update(
                        {
                            'Встроенная веб-камера': {
                                'value': value_char.split(','),
                                'char_name': 'Встроенная веб-камера',
                                'vid_name': 'один из',
                                'unit_name': '(без наименования)'
                            }
                        }
                    )
                except:
                    each_one_char.update(
                        {
                            'Встроенная веб-камера': {
                                'value': value_char.split.strip(),
                                'char_name': 'Встроенная веб-камера',
                                'vid_name': 'один из',
                                'unit_name': '(без наименования)'
                            }
                        }
                    )                  
            except:
                pass
        elif name_char == 'Углы обзора экрана (гор/верт)' or name_char == 'Угол обзора гор/верт':
            try:
                try:
                    each_one_char.update(
                        {
                            'Гориз. область обзора, градусов': {
                                'value': value_char.split('/')[0].replace('°', '').strip(),
                                'char_name': 'Гориз. область обзора, градусов',
                                'vid_name': 'один из',
                                'unit_name': 'градус'
                            }
                        }
                    )
                    each_one_char.update(
                        {
                            'Верт. область обзора, градусов': {
                                'value': value_char.split('/')[1].replace('°', '').strip(),
                                'char_name': 'Верт. область обзора, градусов',
                                'vid_name': 'один из',
                                'unit_name': 'градус'
                            }
                        }
                    )
                except:
                    each_one_char.update(
                        {
                            'Гориз. область обзора, градусов': {
                                'value': value_char.split('/')[0].strip(),
                                'char_name': 'Гориз. область обзора, градусов',
                                'vid_name': 'один из',
                                'unit_name': 'градус'
                            }
                        }
                    )
                    each_one_char.update(
                        {
                            'Верт. область обзора, градусов': {
                                'value': value_char.split('/')[1].replace('°', '').strip(),
                                'char_name': 'Верт. область обзора, градусов',
                                'vid_name': 'один из',
                                'unit_name': 'градус'
                            }
                        }
                    ) 
            except:
                pass
        elif name_char == 'Покрытие':
            try:
                each_one_char.update(
                    {
                        'Покрытие': {
                            'value': value_char.strip(),
                            'char_name': 'Покрытие',
                            'vid_name': 'один из',
                            'unit_name': '(без наименования)'
                        }
                    }
                ) 
            except:
                pass
        elif name_char == 'Плотность пикселей (PPI)' or name_char == 'Плотность пикселей':
            try:
                each_one_char.update(
                    {
                        'Плотность пикселей (PPI)': {
                            'value': value_char.strip(),
                            'char_name': 'Плотность пикселей (PPI)',
                            'vid_name': 'один из',
                            'unit_name': 'градус'
                        }
                    }
                ) 
            except:
                pass
        elif name_char == 'Размер пикселя':
            try:
                each_one_char.update(
                    {
                        'Размер пикселя': {
                            'value': value_char.strip().replace('мкм', '').strip(),
                            'char_name': 'Размер пикселя',
                            'vid_name': 'один из',
                            'unit_name': 'мкм'
                        }
                    }
                ) 
            except:
                pass
        elif name_char == 'Цвет корпуса' or name_char == 'Цвет' or name_char == 'Основной цвет':
            try:
                each_one_char.update(
                    {
                        'Цвет': {
                            'value': value_char.strip(),
                            'char_name': 'Цвет',
                            'vid_name': 'один из',
                            'unit_name': '(без наименования)'
                        }
                    }
                ) 
            except:
                pass
        elif name_char == 'Радиус изогнутости':
            try:
                each_one_char.update(
                    {
                        'Изогнутый экран': {
                            'value': 'да',
                            'char_name': 'Изогнутый экран',
                            'vid_name': 'один из',
                            'unit_name': '(без наименования)'
                        }
                    }
                ) 
            except:
                pass  
        elif name_char == 'Тип подсветки матрицы':
            try:
                each_one_char.update(
                    {
                        'Подсветка': {
                            'value': value_char.strip(),
                            'char_name': 'Подсветка',
                            'vid_name': 'один из',
                            'unit_name': '(без наименования)'
                        }
                    }
                ) 
            except:
                pass 
        elif name_char == 'Время отклика пикселя (MPRT)':
            try:
                each_one_char.update(
                    {
                        'Время отклика MPRT': {
                            'value': value_char.strip(),
                            'char_name': 'Время отклика MPRT',
                            'vid_name': 'один из',
                            'unit_name': '(без наименования)'
                        }
                    }
                ) 
            except:
                pass 
        elif name_char == 'Глубина цвета':
            try:
                each_one_char.update(
                    {
                        'Глубина цвета': {
                            'value': value_char.strip(),
                            'char_name': 'Глубина цвета',
                            'vid_name': 'один из',
                            'unit_name': '(без наименования)'
                        }
                    }
                ) 
            except:
                pass 
        elif name_char == 'Количество USB' or name_char == 'Количество портов USB':
            try:
                each_one_char.update(
                    {
                        'Интерфейсы USB': {
                            'value': f'USB x{value_char.strip().replace("шт", "").strip()}',
                            'char_name': 'Интерфейсы',
                            'vid_name': 'один из',
                            'unit_name': '(без наименования)'
                        }
                    }
                ) 
            except:
                pass
        elif name_char == 'Выход на наушники':
            try:
                each_one_char.update(
                    {
                        'Выходы': {
                            'value': 'на наушники',
                            'char_name': 'Выход',
                            'vid_name': 'один из',
                            'unit_name': '(без наименования)'
                        }
                    }
                ) 
            except:
                pass
        elif name_char == 'Разъем HDMI' or name_char == 'Разъёмов HDMI':
            try:
                each_one_char.update(
                    {
                        'Разъем HDMI': {
                            'value': value_char.strip(),
                            'char_name': 'Разъем HDMI',
                            'vid_name': 'один из',
                            'unit_name': '(без наименования)'
                        }
                    }
                ) 
            except:
                pass
        elif name_char == 'Разъем DisplayPort' or name_char == 'Разъёмов Display Port':
            try:
                each_one_char.update(
                    {
                        'Вход Mini DisplayPort': {
                            'value': value_char.strip(),
                            'char_name': 'Вход Mini DisplayPort',
                            'vid_name': 'один из',
                            'unit_name': '(без наименования)'
                        }
                    }
                ) 
            except:
                pass
        elif name_char == 'Разъем DVI' or name_char == 'Разъёмов DVI':
            try:
                each_one_char.update(
                    {
                        'Разъем DVI': {
                            'value': value_char.strip(),
                            'char_name': 'Разъем DVI',
                            'vid_name': 'один из',
                            'unit_name': '(без наименования)'
                        }
                    }
                ) 
            except:
                pass
        elif name_char == 'Разъем VGA' or name_char == 'Разъёмов VGA (D-SUB)':
            try:
                each_one_char.update(
                    {
                        'Разъем VGA': {
                            'value': value_char.strip(),
                            'char_name': 'Разъем VGA',
                            'vid_name': 'один из',
                            'unit_name': '(без наименования)'
                        }
                    }
                ) 
            except:
                pass
        elif name_char == 'Цветовой охват sRGB':
            try:
                each_one_char.update(
                    {
                        'Цветовой охват sRGB': {
                            'value': value_char.strip(),
                            'char_name': 'Цветовой охват sRGB',
                            'vid_name': 'один из',
                            'unit_name': '(без наименования)'
                        }
                    }
                ) 
            except:
                pass
        elif name_char == 'Комплектация':
            try:
                each_one_char.update(
                    {
                        'Комплектация': {
                            'value': value_char.split(', '),
                            'char_name': 'Комплектация',
                            'vid_name': 'один из',
                            'unit_name': '(без наименования)'
                        }
                    }
                ) 
            except:
                pass
        elif name_char == 'Ширина без подставки':
            try:
                each_one_char.update(
                    {
                        'Ширина': {
                            'value': value_char.strip().replace('мм', '').strip(),
                            'char_name': 'Ширина',
                            'vid_name': 'один из',
                            'unit_name': '(без наименования)'
                        }
                    }
                ) 
            except:
                pass
        elif name_char == 'Высота без подставки':
            try:
                each_one_char.update(
                    {
                        'Высота': {
                            'value': value_char.split('x')[1].strip().replace('мм', '').strip(),
                            'char_name': 'Высота',
                            'vid_name': 'один из',
                            'unit_name': '(без наименования)'
                        }
                    }
                ) 
            except:
                pass
        elif name_char == 'Толщина без подставки':
            try:
                each_one_char.update(
                    {
                        'Глубина': {
                            'value': value_char.strip().replace('мм', '').strip(),
                            'char_name': 'Глубина',
                            'vid_name': 'один из',
                            'unit_name': '(без наименования)'
                        }
                    }
                ) 
            except:
                pass
        elif name_char == 'Вес без подставки':
            try:
                each_one_char.update(
                    {
                        'Вес': {
                            'value': value_char.strip().replace('кг', '').replace('&nbsp;', '').replace('кг&nbsp;', '').replace('кг ', '').strip(),
                            'char_name': 'Вес',
                            'vid_name': 'один из',
                            'unit_name': '(без наименования)'
                        }
                    }
                ) 
            except:
                pass
        elif name_char == 'Вес с подставкой':
            try:
                each_one_char.update(
                    {
                        'Вес нетто': {
                            'value': value_char.strip().replace('кг', '').replace('&nbsp;', '').replace('кг&nbsp;', '').replace('кг ', '').strip(),
                            'char_name': 'Вес нетто',
                            'vid_name': 'один из',
                            'unit_name': '(без наименования)'
                        }
                    }
                ) 
            except:
                pass
        elif name_char == 'Комплектация':
            try:
                each_one_char.update(
                    {
                        'Комплектация': {
                            'value': value_char.strip(),
                            'char_name': 'Комплектация',
                            'vid_name': 'один из',
                            'unit_name': '(без наименования)'
                        }
                    }
                ) 
            except:
                pass
        elif name_char == 'Особенности':
            try:
                each_one_char.update(
                    {
                        'Особенности': {
                            'value': value_char.strip(),
                            'char_name': 'Особенности',
                            'vid_name': 'один из',
                            'unit_name': '(без наименования)'
                        }
                    }
                ) 
            except:
                pass
        elif name_char == 'Встроенная WEB-камера':
            try:
                each_one_char.update(
                    {
                        'Встроенная веб-камера': {
                            'value': value_char.strip(),
                            'char_name': 'Встроенная веб-камера',
                            'vid_name': 'один из',
                            'unit_name': '(без наименования)'
                        }
                    }
                ) 
            except:
                pass
        elif name_char == 'Соотношение сторон':
            try:
                each_one_char.update(
                    {
                        'Соотношение сторон': {
                            'value': value_char.strip(),
                            'char_name': 'Соотношение сторон',
                            'vid_name': 'один из',
                            'unit_name': '(без наименования)'
                        }
                    }
                ) 
            except:
                pass        
        elif name_char == 'Стандарты':
            try:
                each_one_char.update(
                    {
                        'Стандарты': {
                            'value': value_char.strip().split(','),
                            'char_name': 'Стандарты',
                            'vid_name': 'один из',
                            'unit_name': '(без наименования)'
                        }
                    }
                ) 
            except:
                pass
        elif name_char == 'Дополнительная информация':
            try:
                each_one_char.update(
                    {
                        'Дополнительная информация': {
                            'value': value_char.strip(),
                            'char_name': 'Дополнительная информация',
                            'vid_name': 'один из',
                            'unit_name': '(без наименования)'
                        }
                    }
                ) 
            except:
                pass
        elif name_char == 'Функциональность':
            try:
                each_one_char.update(
                    {
                        'Функциональность': {
                            'value': value_char.strip(),
                            'char_name': 'Функциональность',
                            'vid_name': 'один из',
                            'unit_name': '(без наименования)'
                        }
                    }
                ) 
            except:
                pass
        elif name_char == 'Игровой монитор':
            try:
                each_one_char.update(
                    {
                        'Игровой': {
                            'value': value_char.strip(),
                            'char_name': 'Игровой',
                            'vid_name': 'один из',
                            'unit_name': '(без наименования)'
                        }
                    }
                ) 
            except:
                pass
        elif name_char == 'Встроенный микрофон':
            try:
                each_one_char.update(
                    {
                        'Встроенный микрофон': {
                            'value': value_char.strip(),
                            'char_name': 'Встроенный микрофон',
                            'vid_name': 'один из',
                            'unit_name': '(без наименования)'
                        }
                    }
                ) 
            except:
                pass
        elif name_char == 'Соотношение сторон':
            try:
                each_one_char.update(
                    
                    {
                        'Максимальное разрешение': {
                            'value': value_char.strip(),
                            'char_name': 'Максимальное разрешение',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Страна-производитель':
            try:
                each_one_char.update(
                    
                    {
                        'Страна производства': {
                            'value': value_char.strip(),
                            'char_name': 'Страна производства',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
    elif mice:
        if name_char == 'Материал изготовления':
            try:
                each_one_char.update(
                    {
                        'Материал обоймы': {
                            'value': value_char.strip(),
                            'char_name': 'Материал обоймы',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Время автономной работы мыши':
                    try:
                        each_one_char.update(
                            {
                                'Время работы': {
                                    'value': value_char.strip(),
                                    'char_name': 'Время работы',
                                    'vid_name': 'Обычный',
                                    'unit_name': '(без наименования)'
                                }
                            }
                        )
                    except:
                        pass
        elif name_char == 'Дополнительная информация' or name_char == 'Комплектация':
            try:
                each_one_char.update(
                    {
                        'Комплектация': {
                            'value': value_char.strip(),
                            'char_name': 'Комплектация',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Вес':
            try:
                if 'кг' in value_char:
                    each_one_char.update(
                        {
                            'Вес': {
                                'value': float(value_char.replace('кг', '').strip()) * 1000,
                                'char_name': 'Вес',
                                'vid_name': 'Обычный',
                                'unit_name': '(без наименования)'
                            }
                        }
                    ) 
                else:
                    each_one_char.update(
                        {
                            'Вес': {
                                'value': value_char.replace('г', '').strip(),
                                'char_name': 'Вес',
                                'vid_name': 'Обычный',
                                'unit_name': '(без наименования)'
                            }
                        }
                    )
            except:
                pass
        elif name_char == 'Программируемые кнопки':
            try:
                each_one_char.update(
                    {
                        'Количество программируемых клавиш': {
                            'value': value_char.strip(),
                            'char_name': 'Количество программируемых клавиш',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Сканер отпечатка пальца':
            try:
                each_one_char.update(
                    {
                        'Сканер отпечатка пальца': {
                            'value': value_char.strip(),
                            'char_name': 'Сканер отпечатка пальца',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Система регулировки веса':
            try:
                each_one_char.update(
                    {
                        'Система регулировки веса': {
                            'value': value_char.strip(),
                            'char_name': 'Система регулировки веса',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Бесшумные кнопки':
            try:
                each_one_char.update(
                    {
                        'Бесшумное нажатие клавиш мыши': {
                            'value': value_char.strip(),
                            'char_name': 'Бесшумное нажатие клавиш мыши',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Категория':
            try:
                each_one_char.update(
                    {
                        'Назначение': {
                            'value': value_char.strip(),
                            'char_name': 'Назначение',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Страна-производитель':
            try:
                each_one_char.update(
                    {
                        'Страна производства': {
                            'value': value_char.strip(),
                            'char_name': 'Страна производства',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Цвет' or name_char == 'Основной цвет':
            try:
                each_one_char.update(
                    {
                        'Цвет': {
                            'value': value_char.strip(),
                            'char_name': 'Цвет',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Общее количество кнопок' or name_char == 'Количество клавиш':
            try:
                each_one_char.update(
                    {
                        'Количество клавиш': {
                            'value': value_char.strip(),
                            'char_name': 'Количество клавиш',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Горизонтальная прокрутка':
            try:
                each_one_char.update(
                    {
                        'Горизонтальная прокрутка': {
                            'value': value_char.strip(),
                            'char_name': 'Горизонтальная прокрутка',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Тип беспроводной связи':
            try:
                each_one_char.update(
                    {
                        'Тип беспроводной связи': {
                            'value': value_char.strip(),
                            'char_name': 'Тип беспроводной связи',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Разрешение сенсора' or name_char == 'Максимальное разрешение датчика':
            try:
                each_one_char.update(
                    {
                        'Разрешение оптического сенсора': {
                            'value': value_char.strip(),
                            'char_name': 'Разрешение оптического сенсора',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Тип':
            try:
                each_one_char.update(
                    {
                        'Тип': {
                            'value': value_char.strip(),
                            'char_name': 'Тип',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Источник питания мыши' or name_char == 'Питание' or name_char == 'Тип источника питания':
            try:
                each_one_char.update(
                    {
                        'Источник питания мыши': {
                            'value': value_char.strip(),
                            'char_name': 'Источник питания мыши',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Хват' or name_char == 'Дизайн мыши':
            try:
                each_one_char.update(
                    {
                        'Дизайн': {
                            'value': value_char.strip(),
                            'char_name': 'Дизайн',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Особенности, дополнительно':
            try:
                each_one_char.update(
                    {
                        'Особенности': {
                            'value': value_char.strip(),
                            'char_name': 'Особенности',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Радиус действия беспроводной связи':
            try:
                each_one_char.update(
                    {
                        'Радиус действия беспроводной связи': {
                            'value': value_char.strip(),
                            'char_name': 'Радиус действия беспроводной связи',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Интерфейс подключения' or name_char == 'Интерфейс':
            try:
                each_one_char.update(
                    {
                        'Интерфейс подключения': {
                            'value': value_char.strip(),
                            'char_name': 'Интерфейс подключения',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Сенсорная прокрутка':
            try:
                each_one_char.update(
                    {
                        'Сенсорная прокрутка': {
                            'value': value_char.strip(),
                            'char_name': 'Сенсорная прокрутка',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Беспроводная связь':
            try:
                if value_char == 'нет':
                    each_one_char.update(
                        {
                            'Беспроводная связь': {
                                'value': value_char.strip(),
                                'char_name': 'Беспроводная связь',
                                'vid_name': 'Обычный',
                                'unit_name': '(без наименования)'
                            }
                        }
                    )
                else:
                    each_one_char.update(
                        {
                            'Беспроводная связь': {
                                'value': 'есть',
                                'char_name': 'Беспроводная связь',
                                'vid_name': 'Обычный',
                                'unit_name': '(без наименования)'
                            }
                        }
                    )
            except:
                pass
        elif name_char == 'Тип подключения' or name_char == 'Тип соединения':
            try:
                each_one_char.update(
                    {
                        'Тип подключения': {
                            'value': value_char.strip(),
                            'char_name': 'Тип подключения',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Размеры (ШхВхГ)' or name_char == 'Размеры (Ш x Г x В), мм' or name_char == 'Габариты':
            try:
                each_one_char.update(
                    {
                        'Ширина': {
                            'value': value_char.split('x')[0].strip().replace('мм', '').strip(),
                            'char_name': 'Ширина',
                            'vid_name': 'Обычный',
                            'unit_name': 'миллиметр'
                        }
                    }
                )
                each_one_char.update(
                    {
                        'Высота': {
                            'value': value_char.split('x')[1].strip().replace('мм', '').strip(),
                            'char_name': 'Высота',
                            'vid_name': 'Обычный',
                            'unit_name': 'миллиметр'
                        }
                    }
                )
                each_one_char.update(
                    {
                        'Глубина': {
                            'value': value_char.split('x')[2].replace('мм', '').strip(),
                            'char_name': 'Глубина',
                            'vid_name': 'Обычный',
                            'unit_name': 'миллиметр'
                        }
                    }
                )
            except IndexError:
                each_one_char.update(
                    {
                        'Ширина': {
                            'value': value_char.split('х')[0].strip().replace('мм', '').strip(),
                            'char_name': 'Ширина',
                            'vid_name': 'Обычный',
                            'unit_name': 'миллиметр'
                        }
                    }
                )
                each_one_char.update(
                    {
                        'Высота': {
                            'value': value_char.split('х')[1].strip().replace('мм', '').strip(),
                            'char_name': 'Высота',
                            'vid_name': 'Обычный',
                            'unit_name': 'миллиметр'
                        }
                    }
                )
                each_one_char.update(
                    {
                        'Глубина': {
                            'value': value_char.split('х')[2].replace('мм', '').strip(),
                            'char_name': 'Глубина',
                            'vid_name': 'Обычный',
                            'unit_name': 'миллиметр'
                        }
                    }
                )
        elif name_char == 'Ширина':
            try:
                each_one_char.update(
                    {
                        'Ширина': {
                            'value': value_char.strip().replace('мм', '').strip(),
                            'char_name': 'Ширина',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Высота':
            try:
                each_one_char.update(
                    {
                        'Высота': {
                            'value': value_char.strip().replace('мм', '').strip(),
                            'char_name': 'Высота',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Длина':
            try:
                each_one_char.update(
                    {
                        'Длина': {
                            'value': value_char.strip().replace('мм', '').strip(),
                            'char_name': 'Длина',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Материал покрытия':
            try:
                each_one_char.update(
                    {
                        'Материал покрытия': {
                            'value': value_char.strip(),
                            'char_name': 'Материал покрытия',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Тип сенсора мыши' or name_char == 'Тип сенсора':
            try:
                if 'светодиодная' in value_char or 'светодиодный' in value_char:
                    each_one_char.update(
                        {
                            'Принцип работы': {
                                'value': 'оптическая светодиодная',
                                'char_name': 'Принцип работы',
                                'vid_name': 'Обычный',
                                'unit_name': '(без наименования)'
                            }
                        }
                    )
                elif 'лазерный' in value_char or 'лазерная' in value_char:
                    each_one_char.update(
                        {
                            'Принцип работы': {
                                'value': 'оптическая лазерная',
                                'char_name': 'Принцип работы',
                                'vid_name': 'Обычный',
                                'unit_name': '(без наименования)'
                            }
                        }
                    )
            except:
                pass
        elif name_char == 'Частота опроса' or name_char == 'Частота':
            try:
                each_one_char.update(
                    {
                        'Частота опроса': {
                            'value': value_char.strip(),
                            'char_name': 'Частота опроса',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Подсветка':
            try:
                each_one_char.update(
                    {
                        'Подсветка': {
                            'value': value_char.strip(),
                            'char_name': 'Подсветка',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Скорость (IPS)' or name_char == 'Скорость':
            try:
                each_one_char.update(
                    {
                        'Скорость (IPS)': {
                            'value': value_char.strip(),
                            'char_name': 'Скорость (IPS)',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Длина кабеля':
            try:
                each_one_char.update(
                    {
                        'Длина провода': {
                            'value': value_char.strip().replace('м', '').strip(),
                            'char_name': 'Длина провода',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Радиус действия беспроводной связи':
            try:
                each_one_char.update(
                    {
                        'Радиус действия беспроводной связи': {
                            'value': value_char.strip(),
                            'char_name': 'Радиус действия беспроводной связи',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Размеры в упаковке (Ш x Г x В), см':
            try:
                each_one_char.update(
                    {
                        'Ширина': {
                            'value': value_char.split('x')[0].strip(),
                            'char_name': 'Ширина',
                            'vid_name': 'Обычный',
                            'unit_name': 'миллиметр'
                        }
                    }
                )
                each_one_char.update(
                    {
                        'Высота': {
                            'value': value_char.split('x')[1].strip(),
                            'char_name': 'Высота',
                            'vid_name': 'Обычный',
                            'unit_name': 'миллиметр'
                        }
                    }
                )
                each_one_char.update(
                    {
                        'Глубина': {
                            'value': value_char.split('x')[2].replace('мм', '').strip(),
                            'char_name': 'Глубина',
                            'vid_name': 'Обычный',
                            'unit_name': 'миллиметр'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Вес в упаковке':
            try:
                each_one_char.update(
                    {
                        'Вес в упаковке': {
                            'value': value_char.strip().replace('г', '').strip(),
                            'char_name': 'Вес в упаковке',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
    elif ddr:
        if name_char == 'Ранговость':
            try:
                if value_char.strip() == 'одноранговая':
                    each_one_char.update(
                        {
                            'Количество рангов': {
                                'value': '1',
                                'char_name': 'Количество рангов',
                                'vid_name': 'Обычный',
                                'unit_name': '(без наименования)'
                            }
                        }
                    )
                elif value_char.strip() == 'двухранговая':
                    each_one_char.update(
                        {
                            'Количество рангов': {
                                'value': '2',
                                'char_name': 'Количество рангов',
                                'vid_name': 'Обычный',
                                'unit_name': '(без наименования)'
                            }
                        }
                    )
                elif value_char.strip() == 'четырехранговая':
                    each_one_char.update(
                        {
                            'Количество рангов': {
                                'value': '4',
                                'char_name': 'Количество рангов',
                                'vid_name': 'Обычный',
                                'unit_name': '(без наименования)'
                            }
                        }
                    )
                elif value_char.strip() == 'восьмиранговая':
                    each_one_char.update(
                        {
                            'Количество рангов': {
                                'value': '8',
                                'char_name': 'Количество рангов',
                                'vid_name': 'Обычный',
                                'unit_name': '(без наименования)'
                            }
                        }
                    )
            except:
                pass
        elif name_char == 'Вес':
            try:
                each_one_char.update(
                    {
                        'Вес нетто': {
                            'value': value_char.strip().replace('кг', '').strip(),
                            'char_name': 'Вес нетто',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Вес в упаковке':
            try:
                each_one_char.update(
                    {
                        'Вес в упаковке': {
                            'value': value_char.strip(),
                            'char_name': 'Вес в упаковке',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Поддержка XMP' or name_char == 'Профили Intel XMP' or name_char == 'XMP':
            try:
                if value_char != 'нет':
                    each_one_char.update(
                        {
                            'Поддержка XMP': {
                                'value': 'есть',
                                'char_name': 'Поддержка XMP',
                                'vid_name': 'Обычный',
                                'unit_name': '(без наименования)'
                            }
                        }
                    )
                else:
                    each_one_char.update(
                        {
                            'Поддержка XMP': {
                                'value': 'нет',
                                'char_name': 'Поддержка XMP',
                                'vid_name': 'Обычный',
                                'unit_name': '(без наименования)'
                            }
                        }
                    )
            except:
                pass
        elif name_char == 'Радиатор' or name_char == 'Наличие радиатора':
            try:
                each_one_char.update(
                    {
                        'Радиатор': {
                            'value': value_char.strip(),
                            'char_name': 'Радиатор',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Форм-фактор памяти' or name_char == 'Форм-фактор':
            clear_char = value_char.strip().split(' ')[0].replace('-', '')
            try:
                if clear_char == 'DIMM':
                    each_one_char.update(
                        {
                            'Форм-фактор модуля памяти': {
                                'value': 'DIMM',
                                'char_name': 'Форм-фактор модуля памяти',
                                'vid_name': 'Обычный',
                                'unit_name': '(без наименования)'
                            }
                        }
                    )
                elif clear_char == 'LRDIMM':
                    each_one_char.update(
                        {
                            'Форм-фактор модуля памяти': {
                                'value': 'LRDIMM',
                                'char_name': 'Форм-фактор модуля памяти',
                                'vid_name': 'Обычный',
                                'unit_name': '(без наименования)'
                            }
                        }
                    )
                elif clear_char == 'MicroDIMM':
                    each_one_char.update(
                        {
                            'Форм-фактор модуля памяти': {
                                'value': 'MicroDIMM',
                                'char_name': 'Форм-фактор модуля памяти',
                                'vid_name': 'Обычный',
                                'unit_name': '(без наименования)'
                            }
                        }
                    )
                elif clear_char == 'RDIMM':
                    each_one_char.update(
                        {
                            'Форм-фактор модуля памяти': {
                                'value': 'RDIMM',
                                'char_name': 'Форм-фактор модуля памяти',
                                'vid_name': 'Обычный',
                                'unit_name': '(без наименования)'
                            }
                        }
                    )
                elif clear_char == 'SODIMM':
                    each_one_char.update(
                        {
                            'Форм-фактор модуля памяти': {
                                'value': 'SODIMM',
                                'char_name': 'Форм-фактор модуля памяти',
                                'vid_name': 'Обычный',
                                'unit_name': '(без наименования)'
                            }
                        }
                    )
                elif clear_char == 'FBDIMM':
                    each_one_char.update(
                        {
                            'Форм-фактор модуля памяти': {
                                'value': 'FB-DIMM',
                                'char_name': 'Форм-фактор модуля памяти',
                                'vid_name': 'Обычный',
                                'unit_name': '(без наименования)'
                            }
                        }
                    )
            except:
                pass
        elif name_char == 'Упаковка чипов':
            try:
                each_one_char.update(
                    {
                        'Упаковка чипов': {
                            'value': value_char.strip(),
                            'char_name': 'Упаковка чипов',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Высота':
            try:
                each_one_char.update(
                    {
                        'Высота': {
                            'value': value_char.strip().replace('мм', '').strip(),
                            'char_name': 'Высота',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Row Precharge Delay (tRP)':
            try:
                each_one_char.update(
                    {
                        'Row Precharge Delay (tRP)': {
                            'value': value_char.strip(),
                            'char_name': 'Row Precharge Delay (tRP)',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Напряжение питания':
            try:
                if value_char != 'н.д. В':
                    each_one_char.update(
                        {
                            'Напряжение питания': {
                                'value': value_char.strip(),
                                'char_name': 'Напряжение питания',
                                'vid_name': 'Обычный',
                                'unit_name': '(без наименования)'
                            }
                        }
                    )
                else:
                    pass
            except:
                pass
        elif name_char == 'Буферизованная (Registered)' or name_char == 'Буферизованная (RDIMM)':
            try:
                each_one_char.update(
                    {
                        'Буферизованная (Registered)': {
                            'value': value_char.strip(),
                            'char_name': 'Буферизованная (Registered)',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'CAS Latency (CL)':
            try:
                each_one_char.update(
                    {
                        'CAS Latency (CL)': {
                            'value': value_char.strip(),
                            'char_name': 'CAS Latency (CL)',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Игровая':
            try:
                each_one_char.update(
                    {
                        'Игровая': {
                            'value': value_char.strip(),
                            'char_name': 'Игровая',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Низкопрофильная (Low Profile)':
            try:
                if value_char == 'есть':
                    each_one_char.update(
                    {
                        'Низкопрофильная (Low Profile)': {
                            'value': 'да',
                            'char_name': 'Низкопрофильная (Low Profile)',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
                else:
                    each_one_char.update(
                        {
                            'Низкопрофильная (Low Profile)': {
                                'value': 'нет',
                                'char_name': 'Низкопрофильная (Low Profile)',
                                'vid_name': 'Обычный',
                                'unit_name': '(без наименования)'
                            }
                        }
                    )
            except:
                pass
        elif name_char == 'Пропускная способность':
            try:
                each_one_char.update(
                    {
                        'Пропускная способность': {
                            'value': value_char.strip().replace('Мб/с', 'МБ/с'),
                            'char_name': 'Пропускная способность',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Страна-производитель':
            try:
                each_one_char.update(
                    {
                        'Страна производства': {
                            'value': value_char.strip(),
                            'char_name': 'Страна производства',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Activate to Precharge Delay (tRAS':
            try:
                each_one_char.update(
                    {
                        'Activate to Precharge Delay (tRAS': {
                            'value': value_char.strip(),
                            'char_name': 'Activate to Precharge Delay (tRAS',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Количество модулей в комплекте' or name_char == 'Кол-во модулей в упаковке':
            try:
                each_one_char.update(
                    {
                        'Количество модулей в комплекте': {
                            'value': value_char.strip().replace('шт.', '').strip(),
                            'char_name': 'Количество модулей в комплекте',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Количество контактов':
            try:
                each_one_char.update(
                    {
                        'Количество контактов': {
                            'value': value_char.strip(),
                            'char_name': 'Количество контактов',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Совместимость':
            try:
                each_one_char.update(
                    {
                        'Совместимость': {
                            'value': value_char.strip(),
                            'char_name': 'Совместимость',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Объем одного модуля':
            try:
                each_one_char.update(
                    {
                        'Объем одного модуля': {
                            'value': value_char.strip().replace('Гб', 'ГБ'),
                            'char_name': 'Объем одного модуля',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Тип памяти':
            try:
                each_one_char.update(
                    {
                        'Тип памяти': {
                            'value': value_char.strip().replace('DIMM', '').strip(),
                            'char_name': 'Тип памяти',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'RAS to CAS Delay (tRCD)':
            try:
                each_one_char.update(
                    {
                        'RAS to CAS Delay (tRCD)': {
                            'value': value_char.strip(),
                            'char_name': 'RAS to CAS Delay (tRCD)',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Поддержка ECC':
            try:
                each_one_char.update(
                    {
                        'Поддержка ECC': {
                            'value': value_char.strip(),
                            'char_name': 'Поддержка ECC',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Тактовая частота':
            try:
                each_one_char.update(
                    {
                        'Тактовая частота': {
                            'value': value_char.strip(),
                            'char_name': 'Тактовая частота',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Низкопрофильная (Low Profile)':
            try:
                each_one_char.update(
                    {
                        'Низкопрофильная (Low Profile)': {
                            'value': value_char.strip(),
                            'char_name': 'Низкопрофильная (Low Profile)',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Частота':
            try:
                each_one_char.update(
                    {
                        'Частота': {
                            'value': value_char.strip(),
                            'char_name': 'Частота',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Объем одного модуля памяти':
            try:
                each_one_char.update(
                    {
                        'Объем одного модуля памяти': {
                            'value': value_char.strip(),
                            'char_name': 'Объем одного модуля памяти',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Суммарный объем памяти всего комплекта':
            try:
                each_one_char.update(
                    {
                        'Суммарный объем памяти всего комплекта': {
                            'value': value_char.strip(),
                            'char_name': 'Суммарный объем памяти всего комплекта',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Подсветка элементов платы':
            try:
                each_one_char.update(
                    {
                        'Подсветка элементов платы': {
                            'value': value_char.strip(),
                            'char_name': 'Подсветка элементов платы',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Количество чипов модуля':
            try:
                each_one_char.update(
                    {
                        'Количество чипов каждого модуля': {
                            'value': value_char.strip(),
                            'char_name': 'Количество чипов каждого модуля',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Двухсторонняя установка чипов':
            try:
                each_one_char.update(
                    {
                        'Подсветка элементов платы': {
                            'value': value_char.strip(),
                            'char_name': 'Подсветка элементов платы',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Тип':
            try:
                each_one_char.update(
                    {
                        'Тип': {
                            'value': value_char.strip(),
                            'char_name': 'Тип',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
    elif cartridges:
        if name_char == 'Совместимые бренды':
            try:
                each_one_char.update(
                    {
                        'Совместимые бренды': {
                            'value': value_char.strip(),
                            'char_name': 'Совместимые бренды',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Поддерживаемые модели принтеров':
            try:
                each_one_char.update(
                    {
                        'Поддерживаемые модели принтеров': {
                            'value': value_char.strip(),
                            'char_name': 'Поддерживаемые модели принтеров',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Назначение' or name_char == 'Тип картриджа':
            try:
                each_one_char.update(
                    {
                        'Назначение': {
                            'value': value_char.strip(),
                            'char_name': 'Назначение',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Вес':
            try:
                if 'кг' in value_char:
                    each_one_char.update(
                        {
                            'Вес': {
                                'value': float(value_char.replace('кг', '').strip()) * 1000,
                                'char_name': 'Вес',
                                'vid_name': 'Обычный',
                                'unit_name': '(без наименования)'
                            }
                        }
                    ) 
                else:
                    each_one_char.update(
                        {
                            'Вес': {
                                'value': value_char.replace('г', '').strip(),
                                'char_name': 'Вес',
                                'vid_name': 'Обычный',
                                'unit_name': '(без наименования)'
                            }
                        }
                    )
            except:
                pass
        elif name_char == 'Ресурс':
            try:
                each_one_char.update(
                    {
                        'Ресурс': {
                            'value': value_char.strip(),
                            'char_name': 'Ресурс',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Количество в упаковке':
            try:
                each_one_char.update(
                    {
                        'Количество картриджей': {
                            'value': value_char.strip(),
                            'char_name': 'Количество картриджей',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Количество листов бумаги':
            try:
                each_one_char.update(
                    {
                        'Количество листов бумаги': {
                            'value': value_char.strip(),
                            'char_name': 'Количество листов бумаги',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Тип':
            try:
                each_one_char.update(
                    {
                        'Тип': {
                            'value': value_char.strip(),
                            'char_name': 'Тип',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Страна-производитель':
            try:
                each_one_char.update(
                    {
                        'Страна производства': {
                            'value': value_char.strip(),
                            'char_name': 'Страна производства',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Оригинальный':
            try:
                each_one_char.update(
                    {
                        'Оригинальный': {
                            'value': value_char.strip(),
                            'char_name': 'Оригинальный',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Цвет печати':
            try:
                each_one_char.update(
                    {
                        'Цвет': {
                            'value': value_char.strip(),
                            'char_name': 'Цвет',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Европодвес':
            try:
                each_one_char.update(
                    {
                        'Европодвес': {
                            'value': value_char.strip(),
                            'char_name': 'Европодвес',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Ресурс ленты, млн. знаков':
            try:
                each_one_char.update(
                    {
                        'Ресурс ленты, млн. знаков': {
                            'value': value_char.strip(),
                            'char_name': 'Ресурс ленты, млн. знаков',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Размеры (ШхВхГ)' or name_char == 'Размеры (Ш x Г x В), мм' or name_char == 'Габариты':
            try:
                each_one_char.update(
                    {
                        'Ширина': {
                            'value': value_char.split('x')[0].strip(),
                            'char_name': 'Ширина',
                            'vid_name': 'Обычный',
                            'unit_name': 'миллиметр'
                        }
                    }
                )
                each_one_char.update(
                    {
                        'Высота': {
                            'value': value_char.split('x')[1].strip(),
                            'char_name': 'Высота',
                            'vid_name': 'Обычный',
                            'unit_name': 'миллиметр'
                        }
                    }
                )
                each_one_char.update(
                    {
                        'Глубина': {
                            'value': value_char.split('x')[2].replace('мм', '').strip(),
                            'char_name': 'Глубина',
                            'vid_name': 'Обычный',
                            'unit_name': 'миллиметр'
                        }
                    }
                )
            except IndexError:
                each_one_char.update(
                    {
                        'Ширина': {
                            'value': value_char.split('х')[0].strip(),
                            'char_name': 'Ширина',
                            'vid_name': 'Обычный',
                            'unit_name': 'миллиметр'
                        }
                    }
                )
                each_one_char.update(
                    {
                        'Высота': {
                            'value': value_char.split('х')[1].strip(),
                            'char_name': 'Высота',
                            'vid_name': 'Обычный',
                            'unit_name': 'миллиметр'
                        }
                    }
                )
                each_one_char.update(
                    {
                        'Глубина': {
                            'value': value_char.split('х')[2].replace('мм', '').strip(),
                            'char_name': 'Глубина',
                            'vid_name': 'Обычный',
                            'unit_name': 'миллиметр'
                        }
                    }
                )
        elif name_char == 'Ширина':
            try:
                each_one_char.update(
                    {
                        'Ширина': {
                            'value': value_char.strip().replace('мм'),
                            'char_name': 'Ширина',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Высота':
            try:
                each_one_char.update(
                    {
                        'Высота': {
                            'value': value_char.strip().replace('мм'),
                            'char_name': 'Высота',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Длина':
            try:
                each_one_char.update(
                    {
                        'Длина': {
                            'value': value_char.strip().replace('мм'),
                            'char_name': 'Длина',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Наличие чипа':
            try:
                each_one_char.update(
                    {
                        'Наличие чипа': {
                            'value': value_char.strip(),
                            'char_name': 'Наличие чипа',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Емкость':
            try:
                each_one_char.update(
                    {
                        'Емкость': {
                            'value': value_char.strip(),
                            'char_name': 'Емкость',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Вид':
            try:
                each_one_char.update(
                    {
                        'Вид': {
                            'value': value_char.strip(),
                            'char_name': 'Вид',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
    return each_one_char


def search_monitors(name_xlsx: str):
    path_to_xlsx = os.path.join(xlsx_dir, name_xlsx)
    result_xlsx = get_data_from_xlsx(path_to_xlsx)

    nice_width_list = list()
    nice_part_list = list()
    nice_vendor_list = list()

    for name in result_xlsx[2]:
        nice_width_pattern = re.search(
            r'(\d+\.\d\"|\d+\b\"|\d+\,\d\"|[1]\,\d|\d+\.\d\'|\d+\w?\'|\d+\'|\d+\.\d\”|\d+\,\d)', name
        )
        try:
            nice_width_list.append(str(nice_width_pattern[0]).replace("'", '"').replace('”', '"').replace(' ', '"'))
        except:
            nice_width_list.append('')
    
    for part in result_xlsx[1]:
        better_id = part.strip().split(' ')[0].split('/')[0].split('(')
        if len(better_id) >=2:
            nice_partial = re.sub(r'[^a-zA-Z]', '',better_id[1])
            if nice_partial != '':
                better_id = better_id[0] + '(' + nice_partial + ')'
            else:
                better_id = better_id[0] + nice_partial
            
            nice_part_list.append(better_id)
        else:
            better_id[0].split('#')
            nice_part_list.append(better_id[0].split('#')[0])

    for vendor in result_xlsx[3]:
        if vendor.strip() == 'Hewlett-Packard':
            nice_vendor_list.append('HP')
        elif vendor.strip() == 'Elo Touch Solutions':
            nice_vendor_list.append('ELO')
        elif vendor.strip() == 'АБР ТЕХНОЛОДЖИ':
            nice_vendor_list.append('ABR')
        else:
            nice_vendor_list.append(vendor)

    search_req_mon = list(map(lambda a, x, y, z: str(a) + "/" + x + " " + y + " " + z, result_xlsx[0], nice_vendor_list, nice_width_list, nice_part_list))

    result_dict_mon = grab_data(req=search_req_mon, monitors=True, mice=False, ddr=False, cartridges=False)
    
    with open(os.path.join(data_dir, 'result_monitors.json'), 'a', encoding='utf-8') as file:
        json.dump(result_dict_mon, file, indent=4, ensure_ascii=False)


def search_mice(name_xlsx: str):
    path_to_xlsx = os.path.join(xlsx_dir, name_xlsx)
    result_xlsx = get_data_from_xlsx(path_to_xlsx)

    nice_part_list = list()
    nice_vendor_list = list()


    for part in result_xlsx[1]:
        better_id = str(part).split('/')[0]
        if 'STONE BLACK' in better_id:
            better_id.replace('STONE BLACK', 'черный')
            nice_part_list.append(better_id)
        elif 'WHITE' in better_id:
            better_id.replace('WHITE', 'белый')
            nice_part_list.append(better_id)
        else:
            nice_part_list.append(better_id)
    
    for vendor in result_xlsx[3]:
        nice_vendor_list.append(vendor)

    search_req_mice = list(map(lambda a, x, y: str(a) + "/" + str(x) + " " + y, result_xlsx[0], nice_vendor_list, nice_part_list))

    result_dict_mice = grab_data(req=search_req_mice, monitors=False, mice=True, ddr=False, cartridges=False)

    with open(os.path.join(data_dir, 'result_mice.json'), 'a', encoding='utf-8') as file:
        json.dump(result_dict_mice, file, indent=4, ensure_ascii=False)


def search_ddr(name_xlsx:str):
    path_to_xlsx = os.path.join(xlsx_dir, name_xlsx)
    result_xlsx = get_data_from_xlsx(path_to_xlsx)

    nice_part_list = list()
    nice_vendor_list = list()

    for part in result_xlsx[1]:
        better_id = str(part).split(',')[0]
        nice_part_list.append(better_id)

    for vendor in result_xlsx[3]:
        if vendor.strip() == 'Hewlett-Packard':
            nice_vendor_list.append('HP')
        else:
            nice_vendor_list.append(vendor)

    search_req_ddr = list(map(lambda a, x, y: str(a) + "/" + str(x) + " " + y, result_xlsx[0], nice_vendor_list, nice_part_list))

    result_dict_ddr = grab_data(req=search_req_ddr, monitors=False, mice=False, ddr=True, cartridges=False)

    with open(os.path.join(data_dir, 'result_ddr.json'), 'a', encoding='utf-8') as file:
        json.dump(result_dict_ddr, file, indent=4, ensure_ascii=False)


def search_cartridges(name_xlsx:str):
    path_to_xlsx = os.path.join(xlsx_dir, name_xlsx)
    result_xlsx = get_data_from_xlsx(path_to_xlsx)

    nice_part_list = list()
    nice_vendor_list = list()

    for part in result_xlsx[1]:
        better_id = str(part).split('/')[0]
        nice_part_list.append(better_id)

    for vendor in result_xlsx[3]:
        if vendor.strip() == 'Hewlett-Packard':
            nice_vendor_list.append('HP')
        else:
            nice_vendor_list.append(vendor)

    search_req_cartridges = list(map(lambda a, x, y: str(a) + "/" + str(x) + " " + y, result_xlsx[0], nice_vendor_list, nice_part_list))

    result_dict_cartridges = grab_data(req=search_req_cartridges, monitors=False, mice=False, ddr=False, cartridges=True)

    with open(os.path.join(data_dir, 'result_cartridges.json'), 'a', encoding='utf-8') as file:
        json.dump(result_dict_cartridges, file, indent=4, ensure_ascii=False)


def grab_data(req, monitors=False, mice=False, ddr=False, cartridges=False):
    driver = DriverInitialize(headless=True)

    result_dict = list()

    count_cur_req = 0
    count_all_req = len(req)
    with driver:
        for full_req in req:
            count_cur_req += 1
            print(f'\n{full_req}\t{count_cur_req}\{count_all_req}\n')

            result_char_dict = dict()
            char_list = list()

            for site in params_sites_search:
                req = full_req.split('/')[1].strip()
                sku = full_req.split('/')[0].strip()
                search_url = f'{site}{req}'
                try:
                    driver.get(search_url)
                    
                    print(f'[!] URL: {search_url}')

                    if site == 'https://www.regard.ru/catalog?search=':

                        nginx_soup = BeautifulSoup(driver.page_source, 'html.parser')
                        try:
                            too_many = nginx_soup.find('h1').text.strip()
                        except:
                            continue
                    
                        if too_many == '429 Too Many Requests':
                            time.sleep(15)
                            driver.get(search_url)
                            time.sleep(3)
                            print('\t[-] Wait for response url (429)')
                        
                        try:
                            WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, '//*[@id="__next"]/div/div/main')))
                        except:
                            continue

                        item_soup = BeautifulSoup(driver.page_source, 'html.parser')
                        try:
                            print('\t[+] Responsed - OK')
                            true_our_item = item_soup.find('div', class_='rendererWrapper').find('div', class_='ListingRenderer_row__0VJXB').find_all('div', class_='Card_row__6_JG5')

                            if len(true_our_item) == 1:
                                item = item_soup.find('div', class_='rendererWrapper').find('div', class_='ListingRenderer_row__0VJXB').find('div', class_='Card_row__6_JG5').find('a').get('href').strip()
                                item_url = f'https://www.regard.ru/{item}'

                                driver.get(item_url)
                                time.sleep(1)

                                char_item_soup = BeautifulSoup(driver.page_source, 'html.parser')

                                try:
                                    image_item = char_item_soup.find('div', class_='product-slider-container').find('div', class_='swiper-wrapper').find('div', class_='swiper-zoom-container').find('img').get('src')
                                    img = {
                                        'img': {
                                            'img': f'https://www.regard.ru{image_item}'
                                        }
                                    }

                                    result_char_dict.update(img)
                                except:
                                    print('\t[-] Image not found')

                                try:

                                    char_item_wrap = char_item_soup.find('div', class_='ProductCharacteristics_wrap__3RjsG').find('div', class_='ProductCharacteristics_masonry__Ut6Zp').find_all('section', class_='CharacteristicsSection_section__ZctKC')

                                except Exception as ex:
                                    print(ex)   

                                try:
                                    for char_item in char_item_wrap:
                                        char_item_content = char_item.find('div', class_='CharacteristicsSection_content__5BpzM').find_all('div', class_='CharacteristicsItem_item__QnlK2')
                                        for char_content in char_item_content:
                                            name_char = char_content.find('div', class_='CharacteristicsItem_left__ux_qb').find('div', class_='CharacteristicsItem_name__Q7B8V').find('span').text.strip()
                                            value_char = char_content.find('div', class_='CharacteristicsItem_value__fgPkc').text.strip()
                                            if monitors:
                                                char = get_char(name_char=name_char, value_char=value_char, monitors=True, site=site, mice=False, ddr=False, cartridges=False)
                                            elif mice:
                                                char = get_char(name_char=name_char, value_char=value_char, monitors=False, site=site, mice=True, ddr=False, cartridges=False)
                                            elif ddr:
                                                char = get_char(name_char=name_char, value_char=value_char, monitors=False, site=site, mice=False, ddr=True, cartridges=False)
                                            elif cartridges:
                                                char = get_char(name_char=name_char, value_char=value_char, monitors=False, site=site, mice=False, ddr=False, cartridges=True)
                                            result_char_dict.update(char)

                                    print('\t[+] Characteristics grabed')
                                            
                                except Exception as ex:
                                    print(ex)
                                    print('\t[-] No characteristics')
                            else:
                                print('\t[-] Item don\'t found') 
                        except:
                            print('\t[-] Item don\'t found')
                    elif site == 'https://www.onlinetrade.ru/sitesearch.html?query=':
                        try:
                            WebDriverWait(driver, 6).until(EC.element_to_be_clickable((By.ID, 'otv3_submit'))).click()
                            time.sleep(1)
                            print('\t[+] Captcha solved')
                            WebDriverWait(driver, 6).until(EC.presence_of_element_located((By.XPATH, '//*[@id="wrap"]')))
                        except Exception as ex:
                            print('\t[+] No captcha')
                        
                        # try:
                        #     WebDriverWait(driver, 6).until(EC.presence_of_element_located((By.XPATH, '//*[@id="wrap"]')))
                        # except:
                        #     continue

                        item_soup = BeautifulSoup(driver.page_source, 'html.parser')

                        try:
                            true_our_choice = item_soup.find_all('div', class_='indexGoods__item')
                            if len(true_our_choice) == 1:
                                item = item_soup.find('div', class_='goods__items').find('div', class_='indexGoods__item').find('div', class_='indexGoods__item__flexCover').find('a').get('href').strip()
                                item_url = f'https://www.onlinetrade.ru/{item}'

                                driver.get(item_url)
                                time.sleep(3)

                                char_item_soup = BeautifulSoup(driver.page_source, 'html.parser')

                                try:
                                    image_item = char_item_soup.find('div', class_='productPage__displayedItem').find('div', class_='productPage__displayedItem__images').find('div', class_='productPage__displayedItem__images__big').find('a').get('href')
                                    img = {
                                        'img': {
                                            'img': image_item
                                        }
                                    }

                                    result_char_dict.update(img)
                                except:
                                    print('\t[-] Image not found')

                                try:
                                    char_item_wrap = char_item_soup.find('div', attrs={'id': 'tabs_description'}).find('ul', class_='featureList').find_all('li', class_='featureList__item')

                                except Exception as ex:
                                    print(ex)

                                try:
                                    for char_item in char_item_wrap:
                                        name_char = char_item.find('span').text.strip().replace(':', '')
                                        # char_item.find('span').decompose()
                                        value_char = char_item.contents[1].replace('\xa0', '')

                                        if monitors:
                                            char = get_char(name_char=name_char, value_char=value_char, monitors=True, site=site, mice=False, ddr=False, cartridges=False)
                                        elif mice:
                                            char = get_char(name_char=name_char, value_char=value_char, monitors=False, site=site, mice=True, ddr=False, cartridges=False)
                                        elif ddr:
                                            char = get_char(name_char=name_char, value_char=value_char, monitors=False, site=site, mice=False, ddr=True, cartridges=False)
                                        elif cartridges:
                                            char = get_char(name_char=name_char, value_char=value_char, monitors=False, site=site, mice=False, ddr=False, cartridges=True)
                                        result_char_dict.update(char)
                                    print('\t[+] Characteristics grabed')
                                except Exception as ex:
                                    print(ex)
                                    print('\t[-] No characteristics')
                            else:
                                print('\t[-] Item don\'t found') 
                        except:
                            print('\t[-] Item don\'t found')   
                    elif site ==  'https://www.dns-shop.ru/search/?q=':
                        try:
                            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CLASS_NAME, 'product-card-description-specs')))
                            time.sleep(1)
                            # driver.find_element(By.CLASS_NAME, 'product-characteristics__expand_in').click()
                            cur_url = driver.current_url

                            driver.get(f'{cur_url}characteristics/')

                            time.sleep(1)
                            
                            char_item_soup = BeautifulSoup(driver.page_source, 'html.parser')

                            try:
                                image_item = char_item_soup.find('div', class_='product-card-top__images').find('div', class_='product-images-slider').find('picture', class_='product-images-slider__main').find('source').get('srcset')
                                img = {
                                    'img': {
                                        'img': image_item
                                    }
                                }

                                result_char_dict.update(img)
                            except:
                                print('\t[-] Image not found')

                            try:
                                char_item_wrapper = char_item_soup.find('div', class_='product-card-description').find('div', class_='product-card-description-specs').find('div', class_='product-characteristics').find('div', class_='product-characteristics-content')
                            except Exception as ex:
                                print(ex)

                            try:
                                # for char_item_group in char_item_wrapper:
                                char_item = char_item_wrapper.find_all('div', class_='product-characteristics__spec')
                                for char in char_item:
                                    name_char = char.find('div', class_='product-characteristics__spec-title').text.strip()
                                    value_char = char.find('div', class_='product-characteristics__spec-value').text.strip()

                                    if monitors:
                                        char = get_char(name_char=name_char, value_char=value_char, monitors=True, site=site, mice=False, ddr=False, cartridges=False)
                                    elif mice:
                                        char = get_char(name_char=name_char, value_char=value_char, monitors=False, site=site, mice=True, ddr=False, cartridges=False)
                                    elif ddr:
                                        char = get_char(name_char=name_char, value_char=value_char, monitors=False, site=site, mice=False, ddr=True, cartridges=False)
                                    elif cartridges:
                                        char = get_char(name_char=name_char, value_char=value_char, monitors=False, site=site, mice=False, ddr=False, cartridges=True)
                                    result_char_dict.update(char)
                                print('\t[+] Characteristics grabed')
                            except Exception as ex:
                                print(ex)
                                print('\t[-] No characteristics')

                        except Exception as ex:
                            print('\t[-] Item don\'t found')  
                    elif site == 'https://www.novo-market.ru/search/?q=':
                        try:
                            WebDriverWait(driver, 6).until(EC.presence_of_element_located((By.CLASS_NAME, 'products')))
                            
                            wrap_soup = BeautifulSoup(driver.page_source, 'html_parser')
                            title_item = wrap_soup.find_all('div', class_='xml_article')
                            if len(title_item) == 1:
                                try:
                                    item_url = wrap_soup.find('a', class_='js-detail_page_url').get('href').strip()

                                    driver.get(f'https://www.novo-market.ru{item_url}')
                                except Exception as ex:
                                    print(ex)

                                try:
                                    image_item = char_item_soup.find('div', class_='product-gallery').find('div', class_='product-detail-carousel__images').find('div', class_='preview-wrap').find('a').find('imag').get('src')
                                    img = {
                                        'img': {
                                            'img': f'https://www.novo-market.ru{image_item}'
                                        }
                                    }

                                    result_char_dict.update(img)
                                except:
                                    print('\t[-] Image not found')

                                try:
                                    char_item_soup = BeautifulSoup(driver.page_source, 'html.parser')

                                    char_tab = char_item_soup.find('div', attrs={'id': 'properties'}).find_all('div', clas_='tech-info-block')

                                    try:
                                        for char_block in char_tab:
                                            char_miniblock = char_block.find('dl', class_='expand-content').find_all('div')
                                            for char in char_miniblock:
                                                name_char = char.find('dt').text().strip()
                                                value_char = char.find('dd').text().strip()

                                                if monitors:
                                                    char = get_char(name_char=name_char, value_char=value_char, monitors=True, site=site, mice=False, ddr=False, cartridges=False)
                                                elif mice:
                                                    char = get_char(name_char=name_char, value_char=value_char, monitors=False, site=site, mice=True, ddr=False, cartridges=False)
                                                elif ddr:
                                                    char = get_char(name_char=name_char, value_char=value_char, monitors=False, site=site, mice=False, ddr=True, cartridges=False)
                                                elif cartridges:
                                                    char = get_char(name_char=name_char, value_char=value_char, monitors=False, site=site, mice=False, ddr=False, cartridges=True)
                                                result_char_dict.update(char)
                                        print('\t[+] Characteristics grabed')
                                                
                                    except Exception as ex:
                                        print(ex)

                                except Exception as ex:
                                    print(ex)
                            else:
                                print('\t[-] Item don\'t found') 
                        except:
                            print('\t[-] Item don\'t found')  
                except Exception as ex:
                    print(ex)
                    continue

            for v in result_char_dict.values():
                char_list.append(v)
            
            result_dict.append(
                {
                    'sku': sku,
                    'characteristics': char_list
                }
            )
    return result_dict


def main():
        search_monitors(name_xlsx='product_templates_products_monitors.xlsx')
        # search_mice(name_xlsx='product_templates_products_mice.xlsx')
        # search_ddr(name_xlsx='product_templates_products_ddr.xlsx')
        # search_cartridges(name_xlsx='product_templates_products_cartridges.xlsx')


if __name__ == '__main__':
    main()