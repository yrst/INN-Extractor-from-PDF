import pdfplumber
import pandas as pd
from openpyxl import Workbook
import re
import tkinter as tk
from tkinter import filedialog, ttk


def extract_tables_from_pdf(file_path):
    """Извлечение таблиц из PDF-файла."""
    with pdfplumber.open(file_path) as pdf:
        tables = []
        for page in pdf.pages:
            tables.extend(page.extract_tables())
        return tables


def save_tables_to_xlsx(tables, output_file_path):
    """Сохранение таблиц в формате xlsx."""
    wb = Workbook()
    ws = wb.active
    for table in tables:
        for row in table:
            if row[0]:  # Проверка наличия текста в первой ячейке
                ws.append(row)
    wb.save(output_file_path)


def load_tables_from_xlsx(file_path):
    """Загрузка таблиц из файла xlsx."""
    df = pd.read_excel(file_path, header=None)
    return df


def extract_inn(df):
    """Извлечение ИНН из таблицы."""
    inn_pattern = r'\b\d{10,12}\b'  # Шаблон для ИНН (10 или 12 цифр)
    inn_column = df.iloc[:, 0].apply(lambda x: str(x).replace('\n', '').replace(' ', '')).apply(lambda x: re.search(inn_pattern, x).group() if re.search(inn_pattern, x) else x)
    return inn_column.tolist()


def compare_clients(old_inn, new_inn):
    """Сравнение ИНН клиентов."""
    added_clients = list(set(new_inn) - set(old_inn))
    removed_clients = list(set(old_inn) - set(new_inn))
    return added_clients, removed_clients


class GUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Сравнение ИНН клиентов")

        self.old_file_label = tk.Label(self.root, text="Старый файл:")
        self.old_file_label.pack()
        self.old_file_entry = tk.Entry(self.root, width=50)
        self.old_file_entry.pack()
        self.old_file_button = tk.Button(self.root, text="Выбрать файл", command=self.select_old_file)
        self.old_file_button.pack()

        self.new_file_label = tk.Label(self.root, text="Новый файл:")
        self.new_file_label.pack()
        self.new_file_entry = tk.Entry(self.root, width=50)
        self.new_file_entry.pack()
        self.new_file_button = tk.Button(self.root, text="Выбрать файл", command=self.select_new_file)
        self.new_file_button.pack()

        self.compare_button = tk.Button(self.root, text="Сравнить", command=self.compare_files)
        self.compare_button.pack()

        self.result_frame = tk.Frame(self.root)
        self.result_frame.pack()

    def select_old_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        self.old_file_entry.delete(0, tk.END)
        self.old_file_entry.insert(0, file_path)

    def select_new_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        self.new_file_entry.delete(0, tk.END)
        self.new_file_entry.insert(0, file_path)

    def compare_files(self):
        old_file_path = self.old_file_entry.get()
        new_file_path = self.new_file_entry.get()

        old_tables = extract_tables_from_pdf(old_file_path)
        new_tables = extract_tables_from_pdf(new_file_path)

        old_xlsx_file_path = 'old_file.xlsx'
        new_xlsx_file_path = 'new_file.xlsx'

        save_tables_to_xlsx(old_tables, old_xlsx_file_path)
        save_tables_to_xlsx(new_tables, new_xlsx_file_path)

        old_df = load_tables_from_xlsx(old_xlsx_file_path)
        new_df = load_tables_from_xlsx(new_xlsx_file_path)

        old_inn = extract_inn(old_df)
        new_inn = extract_inn(new_df)

        added_clients, removed_clients = compare_clients([x for x in old_inn if x.isdigit()], [x for x in new_inn if x.isdigit()])

        # Создание таблиц для вывода результатов
        self.result_frame.destroy()
        self.result_frame = tk.Frame(self.root)
        self.result_frame.pack()

        left_frame = tk.Frame(self.result_frame)
        left_frame.pack(side=tk.LEFT, padx=10)
        right_frame = tk.Frame(self.result_frame)
        right_frame.pack(side=tk.RIGHT, padx=10)

        tk.Label(left_frame, text="Добавленные клиенты:").pack()
        added_tree = ttk.Treeview(left_frame, columns=('INN',), show='headings')
        added_tree.heading('INN', text='ИНН')
        added_tree.pack()
        for inn in added_clients:
            added_tree.insert('', 'end', values=(inn,))

        tk.Label(right_frame, text="Удаленные клиенты:").pack()
        removed_tree = ttk.Treeview(right_frame, columns=('INN',), show='headings')
        removed_tree.heading('INN', text='ИНН')
        removed_tree.pack()
        for inn in removed_clients:
            removed_tree.insert('', 'end', values=(inn,))

        # Добавление кнопок для копирования ИНН
        def copy_added_inn():
            self.root.clipboard_clear()
            self.root.clipboard_append('\n'.join(added_clients))

        def copy_removed_inn():
            self.root.clipboard_clear()
            self.root.clipboard_append('\n'.join(removed_clients))

        tk.Button(left_frame, text="Копировать", command=copy_added_inn).pack()
        tk.Button(right_frame, text="Копировать", command=copy_removed_inn).pack()

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    gui = GUI()
    gui.run()