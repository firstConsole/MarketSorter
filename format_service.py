import os
import pandas as pd
import yandex as ya
import wildberries as wb
import ozon


def file_format(file):
    try:
        if not os.path.exists(file):
            print(f"Файл {file} не найден.")
            return None
        if file.endswith(('.xls', '.xlsx')):
            df = pd.read_excel(file)
            if not df.empty and df.shape[1] > 0:
                print("Файл прочитан успешно")
                if df.columns[0] == "Информация о заказе":
                    print("Файл Яндекс Маркет")
                    return ya.format_yandex_file(file)
                else:
                    return wb.format_wb_file(file)
            else:
                print("Пустой или неправильный файл")
                return None
        elif file.endswith('.csv'):
            return ozon.format_ozon_file(file)
        else:
            print(f"Неверный формат файла: {file}")
            return None
    except Exception as e:
        print(f"Ошибка при обработке файла: {e}")
        return None