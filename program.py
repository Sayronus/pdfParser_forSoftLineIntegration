"""
Программа для извлечения таблиц спецификаций из PDF-файлов проектной документации,
классификации оборудования по направлениям и определения типа монтажа.
Результат сохраняется в Excel по заданному шаблону.
"""

import os
import re
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import camelot
import pdfplumber
import logging

# Настройка логирования
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


# ----- Конфигурация -----
INPUT_DIR = "pdf_input"          # папка с исходными PDF-файлами
OUTPUT_EXCEL = "result.xlsx"     # выходной Excel-файл
TEMPLATE_COLUMNS = [
    "Номер ПП", "Категория товара", "Помещение", "Артикул", "Производитель",
    "Наименование позиции", "Модель Софтлайн", "Кол-во", "Ед.измер.",
    "Закупочная цена", "Сумма закупочной цены", "НДС закупки",
    "Комментарии", "Ссылка", "Поставщик", "Страна", "Срок поставки",
    "ТХ", "Ответственный", "Монтаж", "Цена продажи", "Регистрация"
]

# Словарь для классификации категорий по подразделам
SECTION_CATEGORY_MAP = {
    "входная зона": "Мебель/Интерьер",
    "гардероб": "Мебель/Интерьер",
    "актовый зал": "Оборудование сцены и звука",
    "столовая": "Кухонное оборудование",
    "пищеблок": "Кухонное оборудование",
    "учительская": "Мебель/ИТ",
    "кабинет психолога": "Мебель/ИТ",
    "кабинет логопеда": "Мебель/ИТ",
    "медицинский": "Медицинское оборудование",
    "коридоры": "Мебель/Интерьер",
    "административный": "Мебель/ИТ",
    "спортивный": "Спортивное оборудование",
    "спортзал": "Спортивное оборудование",
    "библиотека": "Мебель/ИТ",
    "информатика": "ИТ-оборудование",
    "физика": "Лабораторное оборудование",
    "химия": "Лабораторное оборудование",
    "биология": "Лабораторное оборудование",
    "мастерская": "Мастерское оборудование",
    "тир": "Спортивное оборудование",
    "серверная": "ИТ-оборудование",
    "охрана": "Системы безопасности",
}

# Ключевые слова для определения монтажа (монтируемое)
MOUNT_KEYWORDS = [
    "настенн", "навесн", "встраива", "креплени", "монтаж", "установк",
    "подвесн", "стационарн", "пристенн", "потолочн", "фиксиру", "анкер",
    "подключа", "присоедин", "штанга", "турникет", "стойк", "ферм",
    "щит баскетбольн", "кольцо", "сетка", "канат", "шведск", "брусь",
    "перекладин", "плинтус"
]
# Ключевые слова для немонтируемого (переносное)
NON_MOUNT_KEYWORDS = [
    "переносн", "мобильн", "на колес", "складн", "передвижн",
    "настольн", "ручн", "портативн", "транспортир", "сумк",
    "чехол", "кейс", "ящик", "подставк", "тележк"
]


def get_pdf_files(directory):
    """Возвращает список PDF-файлов в указанной директории."""
    return [os.path.join(directory, f) for f in os.listdir(directory) if f.lower().endswith('.pdf')]


def extract_tables_camelot(pdf_path, pages='all', flavor='lattice'):
    """
    Извлечение таблиц с помощью Camelot.
    Возвращает список DataFrame.
    """
    try:

        tables = camelot.read_pdf(pdf_path, pages=pages, flavor=flavor)
        print('!')
        dfs = [table.df for table in tables]
        logging.info(f"Camelot ({flavor}) извлёк {len(dfs)} таблиц из {pdf_path}")
        return dfs
    except Exception as e:
        logging.warning(f"Camelot не удался для {pdf_path}: {e}")
        return []


def extract_tables_pdfplumber(pdf_path):
    """
    Извлечение таблиц с помощью pdfplumber (для случаев, когда Camelot не справляется).
    Возвращает список DataFrame.
    """
    dfs = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                # Пытаемся извлечь таблицы
                tables = page.extract_tables()
                for table in tables:
                    if table:
                        # Преобразуем в DataFrame
                        df = pd.DataFrame(table)
                        # Очистка от пустых строк
                        df = df.dropna(how='all')
                        if not df.empty:
                            dfs.append(df)
        logging.info(f"pdfplumber извлёк {len(dfs)} таблиц из {pdf_path}")
    except Exception as e:
        logging.warning(f"pdfplumber не удался для {pdf_path}: {e}")
    return dfs


def extract_tables(pdf_path):
    """
    Основная функция извлечения таблиц.
    Сначала пытается Camelot (lattice), затем stream, затем pdfplumber.
    Возвращает список DataFrame.
    """
    dfs = []
    for flavor in ['lattice', 'stream']:

        dfs = extract_tables_camelot(pdf_path, flavor=flavor)

        if dfs:
            break
    if not dfs:
        dfs = extract_tables_pdfplumber(pdf_path)
    return dfs


def clean_table(df):
    """
    Очистка DataFrame: удаление пустых строк/столбцов, объединение многострочных ячеек.
    """
    # Удаляем полностью пустые строки
    df = df.dropna(how='all')
    # Удаляем строки, где все значения - пустые строки
    df = df[~df.apply(lambda row: row.astype(str).str.strip().eq('').all(), axis=1)]
    # Заменяем NaN на пустую строку
    df = df.fillna('')
    # Объединяем многострочный текст в ячейках (замена \n на пробел)
    df = df.applymap(lambda x: ' '.join(str(x).split()) if isinstance(x, str) else x)
    return df


def detect_header(df):
    """
    Пытается определить строку заголовка по ключевым словам.
    Возвращает индекс строки-заголовка или -1, если не найден.
    """
    header_keywords = ['наименование', 'количество', 'ед.изм', 'артикул', 'модель', 'примечание', 'кол-во', '№']
    for i, row in df.iterrows():
        # Объединяем все ячейки строки в одну строку (приводим к нижнему регистру)
        combined = ' '.join(row.astype(str).str.lower())
        # Проверяем, есть ли хотя бы два ключевых слова
        matches = sum(1 for kw in header_keywords if kw in combined)
        if matches >= 2:
            return i
    return -1


def map_columns(df, header_row_idx):
    """
    На основе строки заголовка определяет соответствие столбцов требуемым полям.
    Возвращает словарь: {'name': col_idx, 'qty': col_idx, ...}
    """
    if header_row_idx == -1:
        # Если заголовок не найден, используем порядок: 0 - номер, 1 - наименование, 2 - кол-во, 3 - ед.изм., последний - примечание
        col_map = {
            'num': 0,
            'name': 1,
            'qty': 2 if len(df.columns) > 2 else None,
            'unit': 3 if len(df.columns) > 3 else None,
            'note': len(df.columns) - 1 if len(df.columns) > 1 else None,
            'article': None,
            'manufacturer': None,
            'model': None,
        }
        return col_map

    header_row = df.iloc[header_row_idx].astype(str).str.lower()
    col_map = {}
    # Определяем ключевые слова для каждого поля
    patterns = {
        'num': ['№', 'п/п', 'поз', 'позиция'],
        'name': ['наименование', 'название', 'описание'],
        'qty': ['кол-во', 'количество', 'кол'],
        'unit': ['ед.изм', 'единица', 'ед'],
        'note': ['примечание', 'комментарий'],
        'article': ['артикул', 'код', 'арт'],
        'manufacturer': ['производитель', 'завод', 'изготовитель'],
        'model': ['модель', 'тип', 'марка'],
    }
    for field, keywords in patterns.items():
        for idx, cell in enumerate(header_row):
            for kw in keywords:
                if kw in cell:
                    col_map[field] = idx
                    break
            if field in col_map:
                break
    # Если не нашли наименование, предполагаем второй столбец (индекс 1)
    if 'name' not in col_map:
        col_map['name'] = 1 if len(df.columns) > 1 else 0
    # Если не нашли количество, предполагаем третий столбец (индекс 2)
    if 'qty' not in col_map and len(df.columns) > 2:
        col_map['qty'] = 2
    # Если не нашли единицу измерения, предполагаем четвертый столбец (индекс 3)
    if 'unit' not in col_map and len(df.columns) > 3:
        col_map['unit'] = 3
    # Примечание - последний столбец, если не найден
    if 'note' not in col_map and len(df.columns) > 1:
        col_map['note'] = len(df.columns) - 1
    return col_map


def extract_context_from_page(pdf_path, page_num):
    """
    Извлекает текст страницы и пытается определить название подраздела / помещения.
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if page_num < len(pdf.pages):
                page = pdf.pages[page_num]
                text = page.extract_text()
                if text:
                    lines = text.split('\n')
                    # Ищем ключевые слова "Подраздел", "Блок", "Пом."
                    context = ''
                    for line in lines:
                        if 'подраздел' in line.lower() or 'блок' in line.lower() or 'пом.' in line.lower():
                            context = line.strip()
                            break
                    return context
    except:
        pass
    return ''


def classify_category(context, item_name):
    """
    Классифицирует категорию товара на основе контекста и наименования.
    """
    context_lower = context.lower()
    # Проверяем по контексту
    for key, category in SECTION_CATEGORY_MAP.items():
        if key in context_lower:
            return category
    # Если контекст не помог, проверяем по наименованию
    name_lower = item_name.lower()
    if 'стол' in name_lower or 'стул' in name_lower or 'шкаф' in name_lower or 'кресло' in name_lower or 'диван' in name_lower:
        return 'Мебель'
    if 'компьютер' in name_lower or 'ноутбук' in name_lower or 'монитор' in name_lower or 'мфу' in name_lower or 'сервер' in name_lower:
        return 'ИТ-оборудование'
    if 'плита' in name_lower or 'холодильник' in name_lower or 'морозильник' in name_lower or 'стол производственный' in name_lower or 'ванна моечная' in name_lower:
        return 'Кухонное оборудование'
    if 'спорт' in name_lower or 'гимнастический' in name_lower or 'баскетбольный' in name_lower or 'волейбольный' in name_lower:
        return 'Спортивное оборудование'
    if 'медицинский' in name_lower or 'кушетка' in name_lower or 'ширма' in name_lower or 'весы медицинские' in name_lower:
        return 'Медицинское оборудование'
    return 'Прочее'


def determine_mount(item_name, note):
    """
    Определяет, монтируемое ли оборудование (Да/Нет) на основе текста.
    """
    text = (item_name + ' ' + note).lower()
    # Проверка на монтируемое
    for kw in MOUNT_KEYWORDS:
        if kw in text:
            return 'Да'
    # Проверка на немонтируемое (если есть явное указание)
    for kw in NON_MOUNT_KEYWORDS:
        if kw in text:
            return 'Нет'
    # Если нет явных признаков, по умолчанию считаем немонтируемым
    return 'Нет'


def process_table(df, context, start_index):
    """
    Обрабатывает один DataFrame (таблицу), извлекает строки спецификации и возвращает список словарей для шаблона.
    """
    # Очистка таблицы
    df_clean = clean_table(df)
    if df_clean.empty:
        return []

    # Определяем строку заголовка
    header_idx = detect_header(df_clean)
    # Если заголовок найден, удаляем его из данных
    if header_idx != -1:
        df_data = df_clean.iloc[header_idx+1:].reset_index(drop=True)
        header_row = df_clean.iloc[header_idx]
    else:
        df_data = df_clean.copy()
        header_row = None

    # Маппинг столбцов
    col_map = map_columns(df_data, header_idx)
    # Если не удалось определить колонки, используем порядок
    if 'name' not in col_map or col_map['name'] is None:
        col_map['name'] = 0
    if 'qty' not in col_map or col_map['qty'] is None:
        col_map['qty'] = 1 if len(df_data.columns) > 1 else None
    if 'unit' not in col_map or col_map['unit'] is None:
        col_map['unit'] = 2 if len(df_data.columns) > 2 else None
    if 'note' not in col_map or col_map['note'] is None:
        col_map['note'] = len(df_data.columns) - 1 if len(df_data.columns) > 1 else None

    # Извлекаем данные
    records = []
    for idx, row in df_data.iterrows():
        # Получаем значения ячеек
        name = row[col_map['name']] if col_map['name'] is not None else ''
        qty = row[col_map['qty']] if col_map['qty'] is not None else ''
        unit = row[col_map['unit']] if col_map['unit'] is not None else ''
        note = row[col_map['note']] if col_map['note'] is not None else ''
        # Артикул, производитель, модель могут быть в других столбцах
        article = row[col_map.get('article')] if col_map.get('article') is not None else ''
        manufacturer = row[col_map.get('manufacturer')] if col_map.get('manufacturer') is not None else ''
        model = row[col_map.get('model')] if col_map.get('model') is not None else ''

        # Пропускаем пустые строки (без наименования)
        if not name or name.strip() == '':
            continue

        # Классификация
        category = classify_category(context, name)
        mount = determine_mount(name, note)

        # Формируем запись
        record = {
            'Номер ПП': start_index + idx + 1,
            'Категория товара': category,
            'Помещение': context,
            'Артикул': article,
            'Производитель': manufacturer,
            'Наименование позиции': name,
            'Модель Софтлайн': model,
            'Кол-во': qty,
            'Ед.измер.': unit,
            'Комментарии': note,
            'Монтаж': mount,
            # Остальные поля оставляем пустыми (или можно заполнить по умолчанию)
            'Закупочная цена': '',
            'Сумма закупочной цены': '',
            'НДС закупки': '',
            'Ссылка': '',
            'Поставщик': '',
            'Страна': '',
            'Срок поставки': '',
            'ТХ': '',
            'Ответственный': '',
            'Цена продажи': '',
            'Регистрация': '',
        }
        records.append(record)

    return records


def process_pdf(pdf_path):
    """
    Обрабатывает один PDF-файл: извлекает таблицы, обрабатывает каждую, возвращает список записей.
    """
    logging.info(f"Обработка PDF: {pdf_path}")
    all_records = []
    # Извлекаем таблицы

    dfs = extract_tables(pdf_path)
    
    if not dfs:
        logging.warning(f"В {pdf_path} не найдено таблиц.")
        return all_records

    start_index = 0
    # Для каждой таблицы пытаемся получить контекст (из названия подраздела)
    with pdfplumber.open(pdf_path) as pdf:
        for table_idx, df in enumerate(dfs):

            # Определяем страницу таблицы (примерно, но у Camelot есть .page)
            # В Camelot таблицы содержат номер страницы, но мы его не сохранили
            # Поэтому используем приблизительный контекст: ищем в первых строках таблицы
            # или просто берем контекст из всего документа (упрощаем)
            context = ''
            # Ищем в первых нескольких строках таблицы ключевые слова
            for i in range(min(3, len(df))):
                row_text = ' '.join(df.iloc[i].astype(str))
                if 'подраздел' in row_text.lower() or 'блок' in row_text.lower():
                    context = row_text.strip()
                    break
            if not context:
                # Если не нашли, можно извлечь текст страницы (но у нас нет номера страницы)
                # В качестве упрощения используем имя файла (без расширения)
                context = os.path.splitext(os.path.basename(pdf_path))[0]
            # Обрабатываем таблицу
            records = process_table(df, context, start_index)
            all_records.extend(records)
            start_index += len(records)

    return all_records


def save_to_excel(records, output_path):
    """
    Сохраняет список записей в Excel по шаблону.
    """
    df_result = pd.DataFrame(records, columns=TEMPLATE_COLUMNS)
    # Сохраняем в Excel
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df_result.to_excel(writer, index=False, sheet_name='Спецификации')
    logging.info(f"Результат сохранён в {output_path}")


def main():
    """Главная функция."""
    # Проверяем наличие папки с PDF
    if not os.path.exists(INPUT_DIR):
        logging.error(f"Папка {INPUT_DIR} не найдена.")
        return

    pdf_files = get_pdf_files(INPUT_DIR)
    if not pdf_files:
        logging.warning("Нет PDF-файлов для обработки.")
        return

    all_records = []
    for pdf in pdf_files:
        records = process_pdf(pdf)
        all_records.extend(records)

    if not all_records:
        logging.warning("Не удалось извлечь ни одной позиции.")
        return

    # Сортировка по номеру ПП
    all_records.sort(key=lambda x: x['Номер ПП'])
    # Сброс номеров (чтобы были последовательными)
    for i, rec in enumerate(all_records):
        rec['Номер ПП'] = i + 1

    save_to_excel(all_records, OUTPUT_EXCEL)
    logging.info(f"Всего обработано позиций: {len(all_records)}")


if __name__ == "__main__":
    main()