import sys
import os
import pandas as pd

def merge_excel(base_file, local_file, remote_file):
    print("\n[Git-Excel-Driver] Начинается автоматическое объединение таблиц Altium...")
    try:
        # 1. Считываем данные из всех трех версий файла
        df_base = pd.read_excel(base_file)
        df_local = pd.read_excel(local_file)
        df_remote = pd.read_excel(remote_file)
        
        # Определяем имя первой ключевой колонки (например, 'Part Number')
        key_column = df_base.columns[0]
        print(f"[Git-Excel-Driver] Ключевая колонка для поиска дубликатов: '{key_column}'")
        
        # 2. Объединяем строки из локальной и удаленной веток, отсекая дубликаты
        df_merged = pd.concat([df_remote, df_local]).drop_duplicates(subset=[key_column], keep='last')
        
        # Сортируем таблицу по ключевой колонке
        df_merged = df_merged.sort_values(by=key_column).reset_index(drop=True)
        
        # 3. Перезаписываем данные, строго сохраняя бинарник VBA-макросов
        # Параметр keep_vba=True внутри engine_kwargs критически важен для сохранения .xlsm структуры
        with pd.ExcelWriter(
            local_file, 
            engine='openpyxl', 
            mode='a', 
            engine_kwargs={'keep_vba': True},
            if_sheet_exists='replace'
        ) as writer:
            df_merged.to_excel(writer, sheet_name='Components', index=False)
            
        print("[Git-Excel-Driver] 🎉 Успех! Таблицы объединены, макросы сохранены.\n")
        return 0
        
    except Exception as e:
        print(f"[Git-Excel-Driver] ❌ Ошибка автоматического слияния: {e}")
        print("[Git-Excel-Driver] Разрешите конфликт вручную стандартными средствами Git.")
        return 1

if __name__ == "__main__":
    if len(sys.argv) < 4:
        print("[Git-Excel-Driver] Недостаточно аргументов.")
        sys.exit(1)
        
    base = sys.argv[1]
    local = sys.argv[2]
    remote = sys.argv[3]
    
    sys.exit(merge_excel(base, local, remote))
