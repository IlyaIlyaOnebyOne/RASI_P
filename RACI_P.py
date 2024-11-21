import win32com.client
import pandas as pd
from openpyxl import load_workbook

# Инициализация MS Project
ms_project = win32com.client.Dispatch("MSProject.Application")
ms_project.Visible = False  # Не отображать GUI

# Открытие файла .mpp
file_path = r"Project_MS.mpp"  # Путь к файлу
ms_project.FileOpen(file_path)

# Извлечение данных
project = ms_project.ActiveProject
tasks_data = []

# Извлекаем информацию о задачах
for task in project.Tasks:
    if task is not None:
        tasks_data.append({
            "Название задачи/Стейкхолдер": task.Name,
            "R": task.Text5 or "",  # Исполняет
            "A": task.Text2 or "",  # Несет ответственность
            "C": task.Text3 or "",  # Консультирует до исполнения
            "I": task.Text4 or "",  # Оповещает после исполнения
        })

# Создаем DataFrame
df = pd.DataFrame(tasks_data)

# Разворачиваем стейкхолдеров в отдельные столбцы
def build_raci_matrix(df):
    raci_matrix = pd.DataFrame()

    for _, row in df.iterrows():
        task_name = row["Название задачи/Стейкхолдер"]

        # Обрабатываем каждую роль (R, A, C, I)
        for role in ["R", "A", "C", "I"]:
            if pd.notna(row[role]):  # Проверяем, что значение не NaN
                stakeholders = str(row[role]).split(", ")
                for stakeholder in stakeholders:
                    stakeholder = stakeholder.strip()  # Убираем лишние пробелы
                    if stakeholder:  # Исключаем пустые строки
                        if stakeholder not in raci_matrix.columns:
                            raci_matrix[stakeholder] = ""  # Добавляем нового стейкхолдера в столбцы
                        
                        # Объединяем роли, если они уже существуют
                        current_roles = raci_matrix.loc[task_name, stakeholder] if task_name in raci_matrix.index else ""
                        new_roles = f"{current_roles}/{role}" if current_roles else role
                        raci_matrix.loc[task_name, stakeholder] = new_roles.strip("/")

    # Очистка NaN и пустых строк
    raci_matrix = raci_matrix.fillna("")  # Убираем NaN (тип float)
    raci_matrix = raci_matrix.replace("nan", "", regex=True)  # Убираем строковые "nan"
    raci_matrix = raci_matrix.applymap(lambda x: "/".join(filter(None, x.split("/"))) if isinstance(x, str) else x)
    
    # Сброс индекса для удобного сохранения
    raci_matrix.index.name = "Название задачи/Стейкхолдер"
    raci_matrix.reset_index(inplace=True)
    return raci_matrix


# Генерация RACI-матрицы
raci_matrix = build_raci_matrix(df)

# Сохранение RACI-матрицы в Excel
raci_file = r"RACI_matrix.xlsx"
raci_matrix.to_excel(raci_file, index=False, engine='openpyxl')

# Автоподгон ширины столбцов
wb = load_workbook(raci_file)
ws = wb.active

for column in ws.columns:
    max_length = 0
    col_letter = column[0].column_letter  # Получаем букву столбца
    for cell in column:
        try:  # Для обработки пустых ячеек
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        except:
            pass
    adjusted_width = max_length + 2  # Небольшой запас для читаемости
    ws.column_dimensions[col_letter].width = adjusted_width

wb.save(raci_file)
ms_project.Quit()

print(f"RACI-матрица сохранена как {raci_file}")
