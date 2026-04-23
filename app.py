import customtkinter as ctk
from CTkTreeview import CTkTreeview
import os
import datetime
from docx import Document
from docx.shared import Cm
import openpyxl
from openpyxl.styles import Font
import threading

# Настройка внешнего вида
ctk.set_appearance_mode("Dark")  # Тёмная тема
ctk.set_default_color_theme("blue")  # Цветовая схема

class FileExplorerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Modern File Explorer — Просмотр файлов с метаданными")
        self.geometry("1210x800")
        self.resizable(True, True)

        # Создаём основной фрейм
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        self.create_widgets()

    def create_widgets(self):
        # Заголовок
        title_label = ctk.CTkLabel(
            self.main_frame,
            text="📁 Modern File Explorer",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title_label.pack(pady=(0, 20))

        # Метка пути
        self.path_label = ctk.CTkLabel(
            self.main_frame,
            text="Выбранная папка: не выбрана",
            font=ctk.CTkFont(size=14),
            wraplength=1100
        )
        self.path_label.pack(fill="x", pady=10)

        # Кнопки управления
        button_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        button_frame.pack(fill="x", pady=10)

        self.select_button = ctk.CTkButton(
            button_frame,
            text="📂 Выбрать папку",
            command=self.select_folder,
            font=ctk.CTkFont(size=14),
            height=40
        )
        self.select_button.pack(side="left", padx=(0, 10))

        export_frame = ctk.CTkFrame(button_frame, fg_color="transparent")
        export_frame.pack(side="right")

        self.docx_button = ctk.CTkButton(
            export_frame,
            text="💾 Экспорт в DOCX",
            command=self.export_to_docx,
            font=ctk.CTkFont(size=12),
            height=35,
            fg_color="#1e88e5",
            hover_color="#1565c0"
        )
        self.docx_button.pack(side="left", padx=5)

        self.xlsx_button = ctk.CTkButton(
            export_frame,
            text="💾 Экспорт в XLSX",
            command=self.export_to_xlsx,
            font=ctk.CTkFont(size=12),
            height=35,
            fg_color="#43a047",
            hover_color="#2e7d32"
        )
        self.xlsx_button.pack(side="left", padx=5)

        # Индикатор прогресса
        self.progress = ctk.CTkProgressBar(self.main_frame)
        self.progress.pack(fill="x", pady=10, padx=10)
        self.progress.set(0)
        self.progress.pack_forget()  # Скрываем изначально

        # Treeview с скроллбаром
        tree_frame = ctk.CTkFrame(self.main_frame)
        tree_frame.pack(fill="both", expand=True, pady=10)
        columns = ('name', 'type', 'size', 'creation', 'modification', 'access')
        self.tree = CTkTreeview(
            tree_frame,
            columns=columns,
            show='headings',
            height=25
        )
        # Заголовки колонок
        self.tree.heading('name', text='Имя файла/папки')
        self.tree.heading('type', text='Тип')
        self.tree.heading('size', text='Размер')
        self.tree.heading('creation', text='Создан')
        self.tree.heading('modification', text='Изменён')
        self.tree.heading('access', text='Открыт')
        # Ширина колонок
        self.tree.column('name', width=500)
        self.tree.column('type', width=80)
        self.tree.column('size', width=100)
        self.tree.column('creation', width=150)
        self.tree.column('modification', width=150)
        self.tree.column('access', width=150)
        # Скроллбар
        scrollbar = ctk.CTkScrollbar(tree_frame, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        # Строка статуса
        self.status_label = ctk.CTkLabel(
            self.main_frame,
            text="Выберите папку для отображения файлов",
            font=ctk.CTkFont(size=12)
        )
        self.status_label.pack(fill="x", pady=5)
    def select_folder(self):
        """Открывает диалоговое окно для выбора папки и обновляет список файлов."""
        folder_path = ctk.filedialog.askdirectory(title="Выберите папку")
        if folder_path:
            self.path_label.configure(text=f"Выбранная папка: {folder_path}")
            # Запускаем обновление списка в отдельном потоке
            threading.Thread(target=self.update_file_list, args=(folder_path,), daemon=True).start()
    def get_file_info(self, file_path):
        """Получает информацию о файле: размер, даты создания, изменения, открытия."""
        try:
            stat = os.stat(file_path)
            # Размер файла в байтах, преобразуем в КБ или МБ
            size_bytes = stat.st_size
            if size_bytes < 1024:
                size = f"{size_bytes} Б"
            elif size_bytes < 1024**2:
                size = f"{size_bytes / 1024:.2f} КБ"
            else:
                size = f"{size_bytes / (1024**2):.2f} МБ"
            # Даты
            creation_time = datetime.datetime.fromtimestamp(stat.st_ctime)
            modification_time = datetime.datetime.fromtimestamp(stat.st_mtime)
            access_time = datetime.datetime.fromtimestamp(stat.st_atime)
            return {
                'size': size,
                'creation': creation_time.strftime('%Y-%m-%d %H:%M:%S'),
                'modification': modification_time.strftime('%Y-%m-%d %H:%M:%S'),
                'access': access_time.strftime('%Y-%m-%d %H:%M:%S')
            }
        except Exception:
            return {'size': 'N/A', 'creation': 'N/A', 'modification': 'N/A', 'access': 'N/A'}
    def collect_all_files(self, folder_path, progress_callback=None):
        """Рекурсивно собирает все файлы и папки с их метаданными."""
        all_items = []
        total_items = 0
        # Первый проход: подсчёт общего количества элементов
        for root, dirs, files in os.walk(folder_path):
            total_items += len(dirs) + len(files)
        processed_items = 0
        # Второй проход: сбор информации с обновлением прогресса
        for root, dirs, files in os.walk(folder_path):
            # Обработка папок
            for dir_name in dirs:
                dir_path = os.path.join(root, dir_name)
                relative_path = os.path.relpath(dir_path, folder_path)
                all_items.append({
                    'name': relative_path,
                    'type': 'Папка',
                    'size': '-',
                    'creation': '-',
                    'modification': '-',
                    'access': '-'
                })
                processed_items += 1
                if progress_callback:
                    progress_callback(processed_items, total_items)

            # Обработка файлов
            for file_name in files:
                file_path = os.path.join(root, file_name)
                relative_path = os.path.relpath(file_path, folder_path)
                info = self.get_file_info(file_path)
                all_items.append({
                    'name': relative_path,
                    'type': 'Файл',
                    'size': info['size'],
                    'creation': info['creation'],
                    'modification': info['modification'],
                    'access': info['access']
                })
                processed_items += 1
                if progress_callback:
                    progress_callback(processed_items, total_items)

        return all_items

    def update_file_list(self, folder_path):
        """Обновляет список файлов в Treeview с метаданными (рекурсивно)."""
        def update_progress(current, total):
            """Обновляет индикатор прогресса."""
            if total > 0:
                self.progress.set(current / total)
            self.status_label.configure(
                text=f"Обработано: {current}/{total} элементов"
            )
            self.update_idletasks()

        # Очищаем текущий список
        for item in self.tree.get_children():
            self.tree.delete(item)

        try:
            # Показываем прогрессбар
            self.progress.pack(fill="x", padx=10, pady=10)
            self.progress.set(0)

            all_items = self.collect_all_files(folder_path, update_progress)
            file_count = sum(1 for item in all_items if item['type'] == 'Файл')
            folder_count = sum(1 for item in all_items if item['type'] == 'Папка')


            # Сортируем по имени для удобства
            all_items.sort(key=lambda x: x['name'])

            for item in all_items:
                self.tree.insert('', 'end', values=(
                    item['name'],
                    item['type'],
                    item['size'],
                    item['creation'],
                    item['modification'],
                    item['access']
                ))

            self.status_label.configure(
                text=f"Найдено: {len(all_items)} элементов "
                f"({file_count} файлов, {folder_count} папок)"
            )

            # Скрываем прогрессбар после завершения
            self.progress.pack_forget()

        except PermissionError:
            self.status_label.configure(text="Ошибка: нет доступа к папке")
            self.progress.pack_forget()
        except Exception as e:
            self.status_label.configure(text=f"Ошибка: {str(e)}")
            self.progress.pack_forget()

    def export_to_docx(self):
        """Экспортирует список файлов в формат DOCX."""
        if not self.tree.get_children():
            self.status_label.configure(text="Нет данных для экспорта")
            return

        file_path = ctk.filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word документы", "*.docx"), ("Все файлы", "*.*")]
        )

        if file_path:
            doc = Document()

            section = doc.sections[0]
            section.left_margin = Cm(1)
            section.right_margin = Cm(1)
            section.top_margin = Cm(1.5)
            section.bottom_margin = Cm(1.5)

            doc.add_heading('Список файлов', 0)

            # Добавляем путь к папке
            #folder_text = self.path_label.cget("text")
            #doc.add_paragraph(f"Папка: {folder_text.replace('Выбранная папка: ', '')}")
            doc.add_paragraph()

            # Создаём таблицу
            headers = ['Имя', 'Тип', 'Размер', 'Создан', 'Изменён', 'Открыт']
            table = doc.add_table(rows=1, cols=len(headers))

            table.style = 'Table Grid'
            table.autofit = False
            table.allow_autofit = False
            #table.width = Cm(18)
            table.columns[0].width = Cm(4)


            # Заголовки таблицы
            hdr_cells = table.rows[0].cells
            for i, header in enumerate(headers):
                hdr_cells[i].text = header

            # Данные таблицы
            for row in self.tree.get_children():
                values = self.tree.item(row)['values']
                row_cells = table.add_row().cells
                for i, value in enumerate(values):
                    row_cells[i].text = str(value)
            doc.save(file_path)
            self.status_label.configure(text=f"Данные экспортированы в DOCX: {file_path}")

    def export_to_xlsx(self):
        """Экспортирует список файлов в формат XLSX."""
        if not self.tree.get_children():
            self.status_label.configure(text="Нет данных для экспорта")
            return

        file_path = ctk.filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel файлы", "*.xlsx"), ("Все файлы", "*.*")]
        )

        if file_path:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Список файлов"
            ws.column_dimensions['A'].width = 80
            ws.column_dimensions['C'].width = 15
            ws.column_dimensions['D'].width = 20
            ws.column_dimensions['E'].width = 20
            ws.column_dimensions['F'].width = 20

            # Заголовки
            headers = ['Имя', 'Тип', 'Размер', 'Создан', 'Изменён', 'Открыт']
            for col, header in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col, value=header)
                cell.font = Font(bold=True)

            # Данные
            row_idx = 2
            for tree_row in self.tree.get_children():
                values = self.tree.item(tree_row)['values']
                for col_idx, value in enumerate(values, 1):
                    ws.cell(row=row_idx, column=col_idx, value=str(value))
                row_idx += 1

            wb.save(file_path)
            self.status_label.configure(text=f"Данные экспортированы в XLSX: {file_path}")


# Запуск приложения
if __name__ == "__main__":
    app = FileExplorerApp()
    app.mainloop()