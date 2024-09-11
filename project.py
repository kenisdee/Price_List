import os
import csv
import re
import openpyxl


def sanitize_filename(filename):
    '''
    Очищает название файла от недопустимых символов.
    '''
    return re.sub(r'[\\/*?:"<>|]', "", filename)


class PriceMachine:

    def __init__(self):
        self.data = []
        self.name_length = 0

    def load_prices(self, directory='prices'):
        '''
        Сканирует указанный каталог. Ищет файлы со словом price в названии.
        В файле ищет столбцы с названием товара, ценой и весом.
        '''
        if not os.path.isdir(directory):
            return "Папка 'prices' не найдена в текущей директории"

        for root, dirs, files in os.walk(directory):
            for file in files:
                if 'price' in file.lower():
                    file_path = os.path.join(root, file)
                    print(f"Загрузка данных из файла: {file_path}")
                    with open(file_path, 'r', encoding='utf-8') as f:
                        reader = csv.reader(f, delimiter=',')  # Используем запятую как разделитель
                        headers = next(reader)
                        print(f"Заголовки файла {file_path}: {headers}")
                        product_col, price_col, weight_col = self._search_product_price_weight(headers)
                        if product_col is None or price_col is None or weight_col is None:
                            print(f"В файле {file_path} не найдены необходимые столбцы.")
                            continue
                        for row in reader:
                            if len(row) > max(product_col, price_col, weight_col):
                                product_name = row[product_col]
                                price = float(row[price_col])
                                weight = float(row[weight_col])
                                price_per_kg = price / weight
                                self.data.append({
                                    'product_name': product_name,
                                    'price': price,
                                    'weight': weight,
                                    'file': file,
                                    'price_per_kg': price_per_kg
                                })
                                self.name_length = max(self.name_length, len(product_name))
        if self.data:
            print("Данные загружены")
        else:
            print("Данные не загружены")
        return "Загрузка завершена"

    def _search_product_price_weight(self, headers):
        '''
        Возвращает номера столбцов для названия товара, цены и веса.
        '''
        product_names = {'товар', 'название', 'наименование', 'продукт'}
        price_names = {'розница', 'цена'}
        weight_names = {'вес', 'масса', 'фасовка'}

        product_col = price_col = weight_col = None

        for i, header in enumerate(headers):
            if header.lower() in product_names:
                product_col = i
            elif header.lower() in price_names:
                price_col = i
            elif header.lower() in weight_names:
                weight_col = i

        return product_col, price_col, weight_col

    def export_to_html(self, fname='output.html'):
        '''
        Выгружает все данные в html файл.
        '''
        sorted_data = sorted(self.data, key=lambda x: x['price_per_kg'])
        result = '''
        <!DOCTYPE html>
        <html>
        <head>
            <title>Позиции продуктов</title>
        </head>
        <body>
            <table border="1">
                <tr>
                    <th>№ п/п</th>
                    <th>Название</th>
                    <th>Цена</th>
                    <th>Вес</th>
                    <th>Файл</th>
                    <th>Цена за кг.</th>
                </tr>
        '''
        for i, item in enumerate(sorted_data, 1):
            result += f'''
                <tr>
                    <td>{i}</td>
                    <td>{item['product_name']}</td>
                    <td>{item['price']}</td>
                    <td>{item['weight']}</td>
                    <td>{item['file']}</td>
                    <td>{item['price_per_kg']:.2f}</td>
                </tr>
            '''
        result += '''
            </table>
        </body>
        </html>
        '''
        with open(fname, 'w', encoding='utf-8') as f:
            f.write(result)
        return "HTML файл создан"

    def find_text(self, text):
        '''
        Получает текст и возвращает список позиций, содержащий этот текст в названии продукта.
        '''
        found_items = [item for item in self.data if text.lower() in item['product_name'].lower()]
        return sorted(found_items, key=lambda x: x['price_per_kg'])

    def export_search_results_to_html(self, results, query):
        '''
        Выгружает результаты поиска в html файл.
        '''
        sanitized_query = sanitize_filename(query)
        output_dir = 'search_results_html'
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        fname = os.path.join(output_dir, f'{sanitized_query}.html')
        result = '''
        <!DOCTYPE html>
        <html>
        <head>
            <title>Результаты поиска</title>
        </head>
        <body>
            <table border="1">
                <tr>
                    <th>№ п/п</th>
                    <th>Название</th>
                    <th>Цена</th>
                    <th>Вес</th>
                    <th>Файл</th>
                    <th>Цена за кг.</th>
                </tr>
        '''
        for i, item in enumerate(results, 1):
            result += f'''
                <tr>
                    <td>{i}</td>
                    <td>{item['product_name']}</td>
                    <td>{item['price']}</td>
                    <td>{item['weight']}</td>
                    <td>{item['file']}</td>
                    <td>{item['price_per_kg']:.2f}</td>
                </tr>
            '''
        result += '''
            </table>
        </body>
        </html>
        '''
        with open(fname, 'w', encoding='utf-8') as f:
            f.write(result)
        return f"HTML файл с результатами поиска '{fname}' создан"

    def export_search_results_to_excel(self, results, query):
        '''
        Выгружает результаты поиска в excel файл.
        '''
        sanitized_query = sanitize_filename(query)
        output_dir = 'search_results_xslx'
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        fname = os.path.join(output_dir, f'{sanitized_query}.xlsx')
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Результаты поиска"
        headers = ['№ п/п', 'Название', 'Цена', 'Вес', 'Файл', 'Цена за кг.']
        ws.append(headers)
        for i, item in enumerate(results, 1):
            ws.append([i, item['product_name'], item['price'], item['weight'], item['file'], item['price_per_kg']])
        wb.save(fname)
        return f"Excel файл с результатами поиска '{fname}' создан"

    def run_console_interface(self):
        '''
        Циклически получает информацию от пользователя и выводит результаты поиска.
        '''
        while True:
            query = input("Введите текст для поиска или 'exit' для завершения: ").strip()
            if query.lower() == 'exit':
                print("Работа завершена.")
                break
            results = self.find_text(query)
            if results:
                print(
                    f"{'№':<3} {'Наименование':<{self.name_length}} {'цена':<6} {'вес':<5} {'файл':<15} {'цена за кг.'}")
                for i, item in enumerate(results, 1):
                    print(
                        f"{i:<3} {item['product_name']:<{self.name_length}} {item['price']:<6.2f} {item['weight']:<5.2f} {item['file']:<15} {item['price_per_kg']:.2f}")
                print(self.export_search_results_to_html(results, query))
                print(self.export_search_results_to_excel(results, query))
            else:
                print("Ничего не найдено.")


# Пример использования
pm = PriceMachine()
print(pm.load_prices())
print(pm.export_to_html())
pm.run_console_interface()