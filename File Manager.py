import os
import shutil
import sys
import psutil
import json
import openpyxl
from PyQt5.QtCore import Qt, QDir, QFileSystemWatcher, QTimer
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QFileSystemModel, QTreeView, QVBoxLayout, QWidget,
    QMenu, QAction, QInputDialog, QTabWidget, QLineEdit, QHBoxLayout, QStyle, QListWidget, QListWidgetItem, QSplitter, QPushButton, QFileDialog
)


class FileExplorer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Advanced File Explorer")
        self.setGeometry(500, 300, 1000, 500)

        # Устанавливаем иконку окна
        icon_path = os.path.join(os.path.dirname(__file__), "Img", "folder.png")  # Путь до иконки
        self.setWindowIcon(QIcon(icon_path))

        # Создаем вкладки с крестиком
        self.tabs = QTabWidget()
        self.tabs.setTabsClosable(True)
        self.tabs.tabCloseRequested.connect(self.close_tab_action)

        # Создаем боковую панель
        self.sidebar = QListWidget()
        self.sidebar.setFixedWidth(200)
        self.populate_sidebar()

        # Создаем основной вид
        self.main_widget = QSplitter(Qt.Horizontal)
        self.main_widget.addWidget(self.sidebar)
        self.main_widget.addWidget(self.tabs)
        self.setCentralWidget(self.main_widget)

        # Добавляем вкладку с последним использованным путем
        last_path = self.load_last_path()
        print(f"Загруженный последний путь: {last_path}")  # Отладка

        if last_path and os.path.exists(last_path):
            self.add_new_tab(last_path)
        else:
            self.add_new_tab("C:\\")

        # Меню
        self.create_menu()

        # Таймер для обновления списка съемных дисков
        self.disk_update_timer = QTimer(self)
        self.disk_update_timer.timeout.connect(self.update_disks)
        self.disk_update_timer.start(5000)  # обновляем список каждые 5 секунд

    def create_menu(self):
        menu_bar = self.menuBar()
        file_menu = menu_bar.addMenu("File")
        
        new_tab_action = QAction("New Tab", self)
        new_tab_action.triggered.connect(self.new_tab_action)
        file_menu.addAction(new_tab_action)

        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

    def populate_sidebar(self):
        """Добавление ярлыков на боковую панель."""
        icons = self.style()
        items = [
            ("Рабочий стол", QDir.homePath() + "/Desktop", icons.standardIcon(QStyle.SP_DesktopIcon)),
            ("Документы", QDir.homePath() + "/Documents", icons.standardIcon(QStyle.SP_FileIcon)),
            ("Изображения", QDir.homePath() + "/Pictures", icons.standardIcon(QStyle.SP_DirIcon)),
            ("Музыка", QDir.homePath() + "/Music", icons.standardIcon(QStyle.SP_FileIcon)),
            ("Загрузки", QDir.homePath() + "/Downloads", icons.standardIcon(QStyle.SP_DirClosedIcon)),
            ("Диск C:", "C:\\", icons.standardIcon(QStyle.SP_DriveHDIcon)),
        ]

        # Добавляем съемные устройства
        self.add_removable_disks(items, icons)

        for name, path, icon in items:
            item = QListWidgetItem(icon, name)
            item.setData(Qt.UserRole, path)
            self.sidebar.addItem(item)

        self.sidebar.itemClicked.connect(self.sidebar_item_clicked)

    def add_removable_disks(self, items, icons):
        """Добавление съемных дисков в список."""
        for disk in psutil.disk_partitions():
            if 'removable' in disk.opts or disk.fstype == '':
                items.append((f"Съемный диск ({disk.device.strip(':\\')})", disk.device, icons.standardIcon(QStyle.SP_DriveHDIcon)))

    def update_disks(self):
        """Обновление списка съемных дисков."""
        icons = self.style()
        items = [
            ("Рабочий стол", QDir.homePath() + "/Desktop", icons.standardIcon(QStyle.SP_DesktopIcon)),
            ("Документы", QDir.homePath() + "/Documents", icons.standardIcon(QStyle.SP_FileIcon)),
            ("Изображения", QDir.homePath() + "/Pictures", icons.standardIcon(QStyle.SP_DirIcon)),
            ("Музыка", QDir.homePath() + "/Music", icons.standardIcon(QStyle.SP_FileIcon)),
            ("Загрузки", QDir.homePath() + "/Downloads", icons.standardIcon(QStyle.SP_DirClosedIcon)),
            ("Диск C:", "C:\\", icons.standardIcon(QStyle.SP_DriveHDIcon)),
        ]

        # Получаем список актуальных съемных дисков
        current_items = self.get_sidebar_items()

        # Очищаем текущий список и добавляем актуальные элементы
        self.sidebar.clear()

        # Добавляем съемные устройства заново
        self.add_removable_disks(items, icons)

        for name, path, icon in items:
            if (name, path, icon) not in current_items:
                item = QListWidgetItem(icon, name)
                item.setData(Qt.UserRole, path)
                self.sidebar.addItem(item)

        self.sidebar.itemClicked.connect(self.sidebar_item_clicked)

    def get_sidebar_items(self):
        """Получить текущие элементы на боковой панели для отслеживания изменений."""
        return [(self.sidebar.item(i).text(), self.sidebar.item(i).data(Qt.UserRole), self.sidebar.item(i).icon()) for i in range(self.sidebar.count())]

    def sidebar_item_clicked(self, item):
        """Переход по выбранному ярлыку в текущей вкладке."""
        path = item.data(Qt.UserRole)
        current_tab = self.tabs.currentWidget()
        if isinstance(current_tab, FileExplorerTab):
            current_tab.update_path(path)
            self.save_last_path(path)  # Сохраняем последний путь

    def new_tab_action(self):
        """Добавление новой вкладки."""
        self.add_new_tab(QDir.homePath())

    def close_tab_action(self, index):
        """Закрытие вкладки."""
        if self.tabs.count() > 1:  # Оставляем хотя бы одну вкладку
            self.tabs.removeTab(index)

    def add_new_tab(self, path):
        """Добавить новую вкладку с файловой системой."""
        new_tab = FileExplorerTab(path)
        self.tabs.addTab(new_tab, os.path.basename(path) or path)
        self.tabs.setCurrentWidget(new_tab)

    def save_last_path(self, path):
        """Сохранение последнего пути в файл."""
        try:
            with open("last_path.json", "w") as file:
                json.dump({"last_path": path}, file)
            print(f"Путь сохранен: {path}")  # Отладка
        except Exception as e:
            print(f"Error saving last path: {e}")

    def load_last_path(self):
        """Загрузка последнего пути из файла."""
        try:
            if os.path.exists("last_path.json"):
                with open("last_path.json", "r") as file:
                    data = json.load(file)
                    return data.get("last_path")
        except Exception as e:
            print(f"Error loading last path: {e}")
        return None


class FileExplorerTab(QWidget):
    def __init__(self, path):
        super().__init__()
        self.layout = QVBoxLayout(self)

        # Поле ввода пути
        self.path_bar = QLineEdit()
        self.path_bar.setText(path)
        self.path_bar.returnPressed.connect(self.navigate_to_path)

        # Дерево файловой системы
        self.model = QFileSystemModel()
        self.model.setRootPath(path)
        self.tree = QTreeView()
        self.tree.setModel(self.model)
        self.tree.setRootIndex(self.model.index(path))
        self.tree.setSortingEnabled(True)
        self.tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tree.customContextMenuRequested.connect(self.open_context_menu)
        self.tree.doubleClicked.connect(self.open_file_or_folder)

        # Отслеживание изменений в файловой системе
        self.watcher = QFileSystemWatcher()
        self.watcher.addPath(path)
        self.watcher.directoryChanged.connect(self.refresh)

        # Верхняя панель
        top_layout = QHBoxLayout()
        top_layout.addWidget(self.path_bar)

        # Кнопка для экспорта в Excel
        self.export_button = QPushButton("Экспорт в Excel", self)
        self.export_button.clicked.connect(self.export_to_excel)
        top_layout.addWidget(self.export_button)

        self.layout.addLayout(top_layout)
        self.layout.addWidget(self.tree)

    def export_to_excel(self):
        """Экспорт списка файлов в Excel."""
        path = self.path_bar.text()
        if not os.path.exists(path):
            print("Path does not exist!")
            return
        
        # Получаем все файлы и папки в текущем каталоге
        files = self.get_files(path)

        # Открываем диалог для выбора места сохранения
        save_path, _ = QFileDialog.getSaveFileName(self, "Save Excel File", "", "Excel Files (*.xlsx);;All Files (*)")
        if save_path:
            # Создаем рабочую книгу Excel
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "File List"

            # Заполняем таблицу
            ws.append(["File Name", "File Path", "File Size (bytes)"])

            for file in files:
                ws.append([file["name"], file["path"], file["size"]])

            # Сохраняем файл
            try:
                wb.save(save_path)
                print(f"File list saved to {save_path}")
            except Exception as e:
                print(f"Error saving Excel file: {e}")

    def get_files(self, path):
        """Получить список файлов в каталоге."""
        files = []
        for root, dirs, filenames in os.walk(path):
            for filename in filenames:
                file_path = os.path.join(root, filename)
                file_size = os.path.getsize(file_path)
                files.append({"name": filename, "path": file_path, "size": file_size})
        return files

    def update_path(self, path):
        """Обновить отображаемую папку."""
        if os.path.exists(path):
            self.path_bar.setText(path)
            self.tree.setRootIndex(self.model.index(path))
            self.watcher.removePaths(self.watcher.directories())
            self.watcher.addPath(path)

    def navigate_to_path(self):
        """Переход по указанному пути."""
        path = self.path_bar.text()
        if os.path.exists(path):
            self.tree.setRootIndex(self.model.index(path))
            self.path_bar.setText(path)
            self.watcher.removePaths(self.watcher.directories())
            self.watcher.addPath(path)
        else:
            self.path_bar.setText("Invalid path!")

    def refresh(self):
        """Обновление содержимого."""
        current_path = self.path_bar.text()
        self.model.setRootPath(current_path)
        self.tree.setRootIndex(self.model.index(current_path))

    def open_context_menu(self, position):
        """Контекстное меню для операций с файлами."""
        index = self.tree.indexAt(position)
        if not index.isValid():
            return

        menu = QMenu(self)
        copy_action = QAction("Copy", self)
        paste_action = QAction("Paste", self)
        delete_action = QAction("Delete", self)
        rename_action = QAction("Rename", self)

        copy_action.triggered.connect(lambda: self.copy_item(index))
        paste_action.triggered.connect(self.paste_item)
        delete_action.triggered.connect(lambda: self.delete_item(index))
        rename_action.triggered.connect(lambda: self.rename_item(index))

        menu.addAction(copy_action)
        menu.addAction(paste_action)
        menu.addAction(delete_action)
        menu.addAction(rename_action)
        menu.exec_(self.tree.viewport().mapToGlobal(position))

    def copy_item(self, index):
        """Копирование файла/папки."""
        self.copied_item = self.model.filePath(index)

    def paste_item(self):
        """Вставка скопированного элемента."""
        current_path = self.model.filePath(self.tree.rootIndex())
        if hasattr(self, 'copied_item') and os.path.exists(self.copied_item):
            destination = os.path.join(current_path, os.path.basename(self.copied_item))
            try:
                if os.path.isdir(self.copied_item):
                    shutil.copytree(self.copied_item, destination)
                else:
                    shutil.copy2(self.copied_item, destination)
                self.refresh()
            except Exception as e:
                print(f"Error pasting item: {e}")

    def delete_item(self, index):
        """Удаление файла/папки."""
        item_path = self.model.filePath(index)
        try:
            if os.path.isdir(item_path):
                shutil.rmtree(item_path)
            else:
                os.remove(item_path)
            self.refresh()
        except Exception as e:
            print(f"Error deleting item: {e}")

    def rename_item(self, index):
        """Переименование файла/папки."""
        item_path = self.model.filePath(index)
        new_name, _ = QInputDialog.getText(self, "Rename Item", "Enter new name:")
        if new_name:
            new_path = os.path.join(os.path.dirname(item_path), new_name)
            try:
                os.rename(item_path, new_path)
                self.refresh()
            except Exception as e:
                print(f"Error renaming item: {e}")

    def open_file_or_folder(self, index):
        """Открытие файла или папки по двойному клику."""
        item_path = self.model.filePath(index)
        if os.path.isdir(item_path):
            self.update_path(item_path)
        else:
            os.startfile(item_path)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    file_explorer = FileExplorer()
    file_explorer.show()
    sys.exit(app.exec_())
