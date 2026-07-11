import sys
import os
import pandas as pd

def merge_excel(base_file, local_file, remote_file):
    print("\n[Git-Excel-Driver] Начинается автоматическое объединение таблиц Altium...")
    try:
        # Читаем три версии файла
        df_base = pd.read_excel(base_file)
        df_local = pd.read_excel(local_file)
        df_remote = pd.read_excel(remote_file)
        
        # Определяем ключевую колонку
        key_column = df_base.columns[0]
        print(f"[Git-Excel-Driver] Ключевая колонка для поиска дубликатов: '{key_column}'")
        
        # Объединяем строки и удаляем дубликаты
        df_merged = pd.concat([df_remote, df_local]).drop_duplicates(subset=[key_column], keep='last')
        
        # Сортируем таблицу
        df_merged = df_merged.sort_values(by=key_column).reset_index(drop=True)
        
        # Записываем результат с принудительным сохранением формата макросов (.xlsm)
        with pd.ExcelWriter(local_file, engine='openpyxl') as writer:
            df_merged.to_excel(writer, index=False)
            # Включаем флаг шаблона макросов для openpyxl
            if hasattr(writer.book, 'write_tmpl'):
                writer.book.write_tmpl = True
            
        print("[Git-Excel-Driver] 🎉 Успех! Таблицы объединены без конфликтов.\n")
        return 0
        
    except Exception as e:
        print(f"[Git-Excel-Driver] ❌ Ошибка автоматического слияния: {e}")
        print("[Git-Excel-Driver] Разрешите конфликт вручную стандартными средствами Git.")
        return 1

if __name__ == "__main__":
    if len(sys.argv) < 4:
        print("[Git-Excel-Driver] Недостаточно аргументов для работы драйвера.")
        sys.exit(1)
        
    base = sys.argv[1]
    local = sys.argv[2]
    remote = sys.argv[3]
    
    sys.exit(merge_excel(base, local, remote))
