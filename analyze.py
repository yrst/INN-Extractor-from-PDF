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

        self.left_frame = tk.Frame(self.result_frame)
        self.left_frame.pack(side='left')

        self.old_clients_count_label = tk.Label(self.left_frame, text="")
        self.old_clients_count_label.pack()

        self.removed_clients_tree = ttk.Treeview(self.left_frame, columns=("INN",), show='headings')
        self.removed_clients_tree.heading(0, text='Удаленные ИНН')
        self.removed_clients_tree.pack()

        self.copy_removed_button = tk.Button(self.left_frame, text="Копировать", command=self.copy_removed)
        self.copy_removed_button.pack()

        self.right_frame = tk.Frame(self.result_frame)
        self.right_frame.pack(side='left')

        self.new_clients_count_label = tk.Label(self.right_frame, text="")
        self.new_clients_count_label.pack()

        self.added_clients_tree = ttk.Treeview(self.right_frame, columns=("INN",), show='headings')
        self.added_clients_tree.heading(0, text='Добавленные ИНН')
        self.added_clients_tree.pack()

        self.copy_added_button = tk.Button(self.right_frame, text="Копировать", command=self.copy_added)
        self.copy_added_button.pack()

        self.result_label = tk.Label(self.root, text="")
        self.result_label.pack()

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

        self.removed_clients_tree.delete(*self.removed_clients_tree.get_children())
        self.added_clients_tree.delete(*self.added_clients_tree.get_children())

        for client in removed_clients:
            self.removed_clients_tree.insert('', 'end', values=(client,))

        for client in added_clients:
            self.added_clients_tree.insert('', 'end', values=(client,))

        self.old_clients_count_label.config(text=f"Количество в старом файле: {len([x for x in old_inn if x.isdigit()])}")
        self.new_clients_count_label.config(text=f"Количество в новом файле: {len([x for x in new_inn if x.isdigit()])}")

        result_text = f"Добавленные клиенты: {len(added_clients)}\n"
        result_text += f"Удаленные клиенты: {len(removed_clients)}"

        self.result_label.config(text=result_text)

    def copy_removed(self):
        self.root.clipboard_clear()
        items = self.removed_clients_tree.get_children()
        text = ""
        for item in items:
            text += self.removed_clients_tree.item(item, 'values')[0] + "\n"
        self.root.clipboard_append(text.strip())
        self.root.update()  # Обновление окна для применения изменений

    def copy_added(self):
        self.root.clipboard_clear()
        items = self.added_clients_tree.get_children()
        text = ""
        for item in items:
            text += self.added_clients_tree.item(item, 'values')[0] + "\n"
        self.root.clipboard_append(text.strip())
        self.root.update()  # Обновление окна для применения изменений

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    gui = GUI()
    gui.run()
