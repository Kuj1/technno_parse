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

def get_char(name_char, value_char, monitors: bool, site):
    each_one_char = dict()
    if monitors:
        if name_char == 'Разрешение экрана' or name_char == 'Максимальное разрешение':
            if site == 'https://www.regard.ru/catalog?search=' or site == 'https://www.novo-market.ru/search/?q=':
                try:
                    value_char.split('(')
                    try:
                        each_one_char.update(
                            
                            {
                                'Соотношение сторон': {
                                    'value': value_char[1].replace(')', '').strip(),
                                    'char_name': 'Соотношение сторон',
                                    'vid_name': 'Обычный',
                                    'unit_name': '(без наименования)'
                                }
                            }
                        )
                        each_one_char.update(
                            {   
                                'Максимальное разрешение': {
                                    'value': value_char[0].strip(),
                                    'char_name': 'Максимальное разрешение',
                                    'vid_name': 'Обычный',
                                    'unit_name': '(без наименования)'
                                }
                            }
                        )
                    except:
                        pass
                except:
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
            elif site == 'https://www.dns-shop.ru/search/?q=' or site == 'https://www.onlinetrade.ru/sitesearch.html?query=':
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
        elif name_char == 'Диагональ' or name_char == 'Диагональ экрана' or name_char == 'Диагональ экрана (дюйм)' or name_char == 'Размер экрана':
            try:
                each_one_char.update(
                    {
                        'Диагональ': {
                            'value': value_char,
                            'char_name': 'Диагональ',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Поддержка HDR':
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
        elif name_char == 'Вес':
            try:
                each_one_char.update(
                    {
                        'Вес': {
                            'value': value_char.strip().replace('кг', '').strip(),
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
                try:
                    each_one_char.update(
                        {
                            'Количество и мощность встроенных динамиков': {
                                'value': value_char.split(',')[1].strip().replace('Вт', ''),
                                'char_name': 'Количество и мощность встроенных динамиков',
                                'vid_name': 'Обычный',
                                'unit_name': 'Вт'
                            }
                        }
                    )
                except:
                    each_one_char.update(
                        {
                            'Количество и мощность встроенных динамиков': {
                                'value': value_char.split.strip().replace('Вт', ''),
                                'char_name': 'Количество и мощность встроенных динамиков',
                                'vid_name': 'Обычный',
                                'unit_name': 'Вт'
                            }
                        }
                    )
            except:
                pass
        elif name_char == 'Разъёмы' or name_char == 'Видео разъемы' or name_char == 'Порты и разъемы':
            try:
                each_one_char.update(
                    {
                        'Разъёмы': {
                            'value': value_char.split(','),
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
                        'Гориз. область обзора': {
                            'value': value_char.replace('°', '').strip(),
                            'char_name': 'Гориз. область обзора',
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
                        'Верт. область обзора': {
                            'value': value_char.replace('°', '').strip(),
                            'char_name': 'Верт. область обзора',
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
                        'USB-хаб': {
                            'value': value_char.split(',')[0].strip(),
                            'char_name': 'USB-хаб',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
                each_one_char.update(
                    {
                        'value': value_char.split(',')[1].strip(),
                        'char_name': 'Кол-во и тип USB-портов',
                        'vid_name': 'Обычный',
                        'unit_name': '(без наименования)'
                    }
                )
            except:
                pass
        elif name_char == 'Время отклика':
            try:
                each_one_char.update(
                    {
                        'Время отклика': {
                            'value': value_char.split(',')[0].replace('мс', '').strip(),
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
                            'value': value_char.strip(),
                            'char_name': 'Контрастность',
                            'vid_name': 'Обычный',
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
                        'value': value_char.strip().split(','),
                        'char_name': 'Функциональность',
                        'vid_name': 'Обычный',
                        'unit_name': 'несколько из'
                    }
                )
            except:
                pass
        elif name_char == 'Динамическая контрастность':
            try:
                each_one_char.update(
                    {
                        'Динамическая контрастность': {
                            'value': value_char.strip(),
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
                        'Яркость': {
                            'value': value_char.replace('кд/м2', ''),
                            'char_name': 'Яркость',
                            'vid_name': 'Обычный',
                            'unit_name': 'кд/м2'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Поверхность экрана' or name_char == 'Покрытие экрана':
            try:
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
        elif name_char == 'Изогнутый экран':
            try:
                each_one_char.update(
                    {
                        'Изогнутый экран': {
                            'value': value_char.split(',')[0],
                            'char_name': 'Изогнутый экран',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
                each_one_char.update(
                    {
                        'Радиус изогнутости': {
                            'value': value_char.split(',')[1],
                            'char_name': 'Радиус изогнутости',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'LED подсветка' or name_char == 'Светодиодная подсветка (LED)':
            try:
                each_one_char.update(
                    {
                        'LED подсветка': {
                            'value': value_char.strip(),
                            'char_name': 'LED подсветка',
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
                        'Тип матрицы': {
                            'value': value_char.strip(),
                            'char_name': 'Тип матрицы',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Сенсорный экран':
            try:
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
                        'Частота обновления кадров': {
                            'value': value_char.replace('Гц', '').strip(),
                            'char_name': 'Частота обновления кадров',
                            'vid_name': 'Обычный',
                            'unit_name': 'Гц'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Крепление на стену (VESA)' or name_char == 'Размер VESA' or name_char == 'Наличие крепления VESA' or name_char == 'Размер крепления VESA':
            try:
                each_one_char.update(
                    {
                        'Крепление на стену (VESA)': {
                            'value': value_char,
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
                            'value': value_char,
                            'char_name': 'Регулировка по высоте',
                            'vid_name': 'Обычный',
                            'unit_name': '(без наименования)'
                        }
                    }
                )
            except:
                pass
        elif name_char == 'Максимальное количество цветов':
            try:
                each_one_char.update(
                    {
                        'Максимальное количество цветов': {
                            'value': value_char.split(','),
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
                            'Веб-камера': {
                                'value': value_char.split(','),
                                'char_name': 'Веб-камера',
                                'vid_name': 'один из',
                                'unit_name': '(без наименования)'
                            }
                        }
                    )
                except:
                    each_one_char.update(
                        {
                            'Веб-камера': {
                                'value': value_char.split.strip(),
                                'char_name': 'Веб-камера',
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
                            'Гориз. область обзора': {
                                'value': value_char.split('/')[0].replace('°', ''),
                                'char_name': 'Гориз. область обзора',
                                'vid_name': 'один из',
                                'unit_name': 'градус'
                            }
                        }
                    )
                    each_one_char.update(
                        {
                            'Верт. область обзора': {
                                'value': value_char.split('/')[1].replace('°', ''),
                                'char_name': 'Верт. область обзора',
                                'vid_name': 'один из',
                                'unit_name': 'градус'
                            }
                        }
                    )
                except:
                    each_one_char.update(
                        {
                            'Гориз. область обзора': {
                                'value': value_char.split('/')[0].strip(),
                                'char_name': 'Гориз. область обзора',
                                'vid_name': 'один из',
                                'unit_name': 'градус'
                            }
                        }
                    )
                    each_one_char.update(
                        {
                            'Верт. область обзора': {
                                'value': value_char.split('/')[1].replace('°', ''),
                                'char_name': 'Верт. область обзора',
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
                            'unit_name': 'градус'
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
                            'value': value_char.strip(),
                            'char_name': 'Размер пикселя',
                            'vid_name': 'один из',
                            'unit_name': 'градус'
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
                        'Радиус изогнутости': {
                            'value': value_char.strip(),
                            'char_name': 'Радиус изогнутости',
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
                        'Тип подсветки матрицы': {
                            'value': value_char.strip(),
                            'char_name': 'Тип подсветки матрицы',
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
                        'Время отклика пикселя (MPRT)': {
                            'value': value_char.strip(),
                            'char_name': 'Время отклика пикселя (MPRT)',
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
        elif name_char == 'Другие разъемы':
            try:
                each_one_char.update(
                    {
                        'Другие разъемы': {
                            'value': value_char.strip(),
                            'char_name': 'Другие разъемы',
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
                        'Количество USB': {
                            'value': value_char.strip(),
                            'char_name': 'Количество USB',
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
                        'Выход на наушники': {
                            'value': value_char.strip(),
                            'char_name': 'Выход на наушники',
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
                        'Разъем DisplayPort': {
                            'value': value_char.strip(),
                            'char_name': 'Разъем DisplayPort',
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
                            'value': value_char.split(','),
                            'char_name': 'Комплектация',
                            'vid_name': 'один из',
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
                            'value': value_char.strip(),
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
                            'value': value_char.strip(),
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
                            'value': value_char.strip(),
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
                            'value': value_char.strip(),
                            'char_name': 'Вес',
                            'vid_name': 'один из',
                            'unit_name': '(без наименования)'
                        }
                    }
                ) 
            except:
                pass
        elif name_char == 'Ширина с подставкой':
            try:
                each_one_char.update(
                    {
                        'Ширина с подставкой': {
                            'value': value_char.strip(),
                            'char_name': 'Ширина с подставкой',
                            'vid_name': 'один из',
                            'unit_name': '(без наименования)'
                        }
                    }
                ) 
            except:
                pass
        elif name_char == 'Минимальная высота с подставкой':
            try:
                each_one_char.update(
                    {
                        'Минимальная высота с подставкой': {
                            'value': value_char.strip(),
                            'char_name': 'Минимальная высота с подставкой',
                            'vid_name': 'один из',
                            'unit_name': '(без наименования)'
                        }
                    }
                ) 
            except:
                pass
        elif name_char == 'Толщина с подставкой':
            try:
                each_one_char.update(
                    {
                        'Толщина с подставкой': {
                            'value': value_char.strip(),
                            'char_name': 'Толщина с подставкой',
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
                        'Вес с подставкой': {
                            'value': value_char.strip(),
                            'char_name': 'Вес с подставкой',
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
                        'Встроенная WEB-камера': {
                            'value': value_char.strip(),
                            'char_name': 'Встроенная WEB-камера',
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
        elif name_char == 'Интерфейсы':
            try:
                each_one_char.update(
                    {
                        'Интерфейсы': {
                            'value': value_char.strip(),
                            'char_name': 'Интерфейсы',
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
                        'Игровой монитор': {
                            'value': value_char.strip(),
                            'char_name': 'Игровой монитор',
                            'vid_name': 'один из',
                            'unit_name': '(без наименования)'
                        }
                    }
                ) 
            except:
                pass
        elif name_char == 'Встроенные колонки':
            try:
                each_one_char.update(
                    {
                        'Встроенные колонки': {
                            'value': value_char.strip(),
                            'char_name': 'Встроенные колонки',
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
    
    return each_one_char




def search_data(url,):
    pass
    #     driver = DriverInitialize(headless=False)
    #     delay = 30

    #     with driver:
    #         name_site = url.split('.')[1]
    #         if name_site in url and name_site == 'regard':
    #             try:
    #                 driver.get(url)
    #                 WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.ID, '__next')))
    #                 print(f'Hi, {name_site} this is first "if"')
    #                 driver.save_screenshot('sh_1.png')
    #             except:
    #                 print('page is not ready')
    #         elif name_site in url and name_site == 'onlinetrade':
    #             try:
    #                 try:
    #                     driver.get(url)
    #                     WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.ID, 'wrap')))
    #                     print(f'Hi, {name_site} this is second "if"')
    #                     driver.save_screenshot('sh_2.png')
    #                 except:
    #                     WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.ID, 'otv3_submit'))).click()
    #                     time.sleep(1)
    #                     driver.save_screenshot('sh_busted_2.png')
    #                     print('busted')
    #             except:
    #                 print('page is not ready')
    #         elif name_site in url and name_site == 'novo-market':
    #             try:
    #                 driver.get(url)
    #                 WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.ID, 'panel')))
    #                 print(f'Hi, {name_site} this is third "if"')
    #                 driver.save_screenshot('sh_3.png')
    #             except:
    #                 print('page is not ready')
    #         elif name_site in url and name_site == 'dns-shop':
    #             try:
    #                 driver.get(url)
    #                 WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.XPATH, 'container category-child')))
    #                 print(f'Hi, {name_site} this is fourth "if"')
    #                 driver.save_screenshot('sh_4.png')
    #             except:
    #                 print('page is not ready')


def search_monitors(name_xlsx: str):
    path_to_xlsx = os.path.join(xlsx_dir, name_xlsx)
    resurlt_xlsx = get_data_from_xlsx(path_to_xlsx)

    nice_width_list = list()
    nice_part_list = list()
    nice_vendor_list = list()

    result_monitors = list()

    for name in resurlt_xlsx[2]:
        nice_width_pattern = re.search(
            r'(\d+\.\d\"|\d+\b\"|\d+\,\d\"|[1]\,\d|\d+\.\d\'|\d+\w?\'|\d+\'|\d+\.\d\”|\d+\,\d)', name
        )
        try:
            nice_width_list.append(str(nice_width_pattern[0]).replace("'", '"').replace('”', '"').replace(' ', '"'))
        except:
            nice_width_list.append('')
    
    for part in resurlt_xlsx[1]:
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

    for vendor in resurlt_xlsx[3]:
        if vendor.strip() == 'Hewlett-Packard':
            nice_vendor_list.append('HP')
        elif vendor.strip() == 'Elo Touch Solutions':
            nice_vendor_list.append('ELO')
        elif vendor.strip() == 'АБР ТЕХНОЛОДЖИ':
            nice_vendor_list.append('ABR')
        else:
            nice_vendor_list.append(vendor)

    search_req_mon = list(map(lambda a, x, y, z: str(a) + "/" + x + " " + y + " " + z, resurlt_xlsx[0], nice_vendor_list, nice_width_list, nice_part_list))
    driver = DriverInitialize(headless=False)

    count_cur_req = 0
    count_all_req = len(search_req_mon)
    with driver:
        for full_req in search_req_mon:
            count_cur_req += 1
            print(f'\n{full_req}\t{count_cur_req}\{count_all_req}\n')

            result_char_dict = dict()
            char_list = list()

            for site in params_sites_search:
                # site = 'https://www.dns-shop.ru/search/?q='
                req = full_req.split('/')[1].strip()
                sku = full_req.split('/')[0].strip()
                search_url = f'{site}{req}'
                driver.get(search_url)
                print(f'[!] URL: {search_url}')
                # time.sleep(1)

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
                        WebDriverWait(driver, 6).until(EC.presence_of_element_located((By.XPATH, '//*[@id="__next"]/div/div/main')))
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

                                char_item_wrap = char_item_soup.find('div', class_='ProductCharacteristics_wrap__3RjsG').find('div', class_='ProductCharacteristics_masonry__Ut6Zp').find_all('section', class_='CharacteristicsSection_section__ZctKC')

                            except Exception as ex:
                                print(ex)   

                            try:
                                for char_item in char_item_wrap:
                                    char_item_content = char_item.find('div', class_='CharacteristicsSection_content__5BpzM').find_all('div', class_='CharacteristicsItem_item__QnlK2')
                                    for char_content in char_item_content:
                                        name_char = char_content.find('div', class_='CharacteristicsItem_left__ux_qb').find('div', class_='CharacteristicsItem_name__Q7B8V').find('span').text.strip()
                                        value_char = char_content.find('div', class_='CharacteristicsItem_value__fgPkc').text.strip()

                                        char = get_char(name_char=name_char, value_char=value_char, monitors=True, site=site)
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
                    except Exception as ex:
                        print('\t[+] No captcha')
                    
                    try:
                        WebDriverWait(driver, 12).until(EC.presence_of_element_located((By.XPATH, '//*[@id="wrap"]')))
                    except:
                        continue

                    item_soup = BeautifulSoup(driver.page_source, 'html.parser')

                    try:
                        true_our_choice = item_soup.find_all('div', class_='indexGoods__item')
                        if len(true_our_choice) == 1:
                            item = item_soup.find('div', class_='goods__items').find('div', class_='indexGoods__item').find('div', class_='indexGoods__item__flexCover').find('a').get('href').strip()
                            item_url = f'https://www.onlinetrade.ru/{item}'

                            driver.get(item_url)
                            time.sleep(1)

                            # WebDriverWait(driver, 12).until(EC.element_to_be_clickable((By.ID, 'ui-id-2'))).click()

                            char_item_soup = BeautifulSoup(driver.page_source, 'html.parser')

                            try:

                                char_item_wrap = char_item_soup.find('div', attrs={'id': 'tabs_description'}).find('ul', class_='featureList').find_all('li', class_='featureList__item')

                            except Exception as ex:
                                print(ex)

                            try:
                                for char_item in char_item_wrap:
                                    name_char = char_item.find('span').text.strip().replace(':', '')
                                    # char_item.find('span').decompose()
                                    value_char = char_item.contents[1].replace('\xa0', '')

                                    char = get_char(name_char=name_char, value_char=value_char, monitors=True, site=site)
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
                            char_item_wrapper = char_item_soup.find('div', class_='product-card-description').find('div', class_='product-card-description-specs').find('div', class_='product-characteristics').find('div', class_='product-characteristics-content').find_all('div', class_='product-characteristics__group')
                        except Exception as ex:
                            print(ex)

                        try:
                            for char_item_group in char_item_wrapper:
                                char_item = char_item_group.find_all('div', class_='product-characteristics__spec')
                                for char in char_item:
                                    name_char = char.find('div', class_='product-characteristics__spec-title').text.strip()
                                    value_char = char.find('div', class_='product-characteristics__spec-value').text.strip()

                                    char = get_char(name_char=name_char, value_char=value_char, monitors=True, site=site)
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
                                char_item_soup = BeautifulSoup(driver.page_source, 'html.parser')

                                char_tab = char_item_soup.find('div', attrs={'id': 'properties'}).find_all('div', clas_='tech-info-block')

                                try:
                                    for char_block in char_tab:
                                        char_miniblock = char_block.find('dl', class_='expand-content').find_all('div')
                                        for char in char_miniblock:
                                            name_char = char.find('dt').text().strip()
                                            value_char = char.find('dd').text().strip()

                                            char = get_char(name_char=name_char, value_char=value_char, monitors=True, site=site)
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

            for v in result_char_dict.values():
                char_list.append(v)
            
            result_monitors.append(
                {
                    'sku': sku,
                    'characteristics': char_list
                }
            )

        with open(os.path.join(data_dir, 'result_monitors.json'), 'a', encoding='utf-8') as file:
            json.dump(result_monitors, file, indent=4, ensure_ascii=False)

def main():
    search_monitors(name_xlsx='product_templates_products_monitors.xlsx')


if __name__ == '__main__':
    # while True:
    main()