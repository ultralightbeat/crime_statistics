import pandas
import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, simpledialog
import matplotlib.pyplot as plt
import matplotlib
from matplotlib.backends.backend_tkagg import (
    FigureCanvasTkAgg, NavigationToolbar2Tk)


def open_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        df = pd.read_excel(file_path)
        display_data(df)


def display_data(df):
    def filter_data():
        filter_value = filter_combobox.get()
        if filter_value:
            filtered_df = df[df['object_name'] == filter_value]
            display_filtered_data(filtered_df)
        else:
            display_filtered_data(df)

    def display_filtered_data(filtered_df):
        # Очищаем Treeview
        for item in tree.get_children():
            tree.delete(item)
        # Вставляем отфильтрованные данные
        for index, row in filtered_df.iterrows():
            values = [row[col] for col in columns_to_display]
            tree.insert("", tk.END, values=values)

    display_frame = tk.Toplevel(root)
    display_frame.title("Данные о преступности")
    # display_frame.
    # Создание Treeview
    tree = ttk.Treeview(display_frame)
    tree["show"] = "headings"
    # Определение колонок
    columns_to_display = ['object_name', 'object_level', 'year2011', 'year2012', 'year2013', 'year2014',
                          'year2015', 'year2016', 'year2017', 'year2018', 'year2019', 'year2020', 'year2021',
                          'year2022', 'comment']
    columns_headers = {'object_name': 'Объект', 'object_level': 'Уровень объекта', 'comment': 'Комментарий'}

    tree["columns"] = columns_to_display

    # Форматирование колонок
    for col in columns_to_display:
        if col in columns_headers:
            tree.heading(col, text=columns_headers[col])
        else:
            year = int(col.replace("year", ""))
            tree.heading(col, text=str(year))
        tree.column(col, width=100, anchor=tk.CENTER, stretch=True)

    # Добавление полос прокрутки
    vsb = ttk.Scrollbar(display_frame, orient="vertical", command=tree.yview)
    vsb.pack(side=tk.RIGHT, fill=tk.Y)
    hsb = ttk.Scrollbar(display_frame, orient="horizontal", command=tree.xview)
    hsb.pack(side=tk.BOTTOM, fill=tk.X)

    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

    df = df.loc[df["indicator_name"].str.startswith("Всего зарегистрировано преступлений") |
                df["indicator_name"].str.startswith("Количество")]

    # Вставка данных
    for index, row in df.iterrows():
        values = [row[col] for col in columns_to_display]
        tree.insert("", tk.END, values=values)

    # Фильтр
    filter_frame = tk.Frame(display_frame)
    filter_frame.pack(pady=10)

    analyze_button = tk.Button(display_frame, text="Анализ трендов", command=lambda: analyze_crime_trend(df))

    analyze_button.pack(pady=10)

    filter_label = tk.Label(filter_frame, text="Фильтр по объекту:")
    filter_label.pack(side=tk.LEFT)

    unique_objects = sorted(set(df['object_name']))
    filter_combobox = ttk.Combobox(filter_frame, values=unique_objects)
    filter_combobox.pack(side=tk.LEFT, padx=10)

    filter_button = tk.Button(filter_frame, text="Фильтр", command=filter_data)
    filter_button.pack(side=tk.LEFT)

    graphics_button = tk.Button(filter_frame, text="Отобразить графики",
                                command=lambda: plot_crime_trend(display_frame, df, filter_combobox.get()))
    graphics_button.pack(side=tk.LEFT)
    # Упаковка Treeview
    tree.pack(expand=tk.YES, fill=tk.BOTH, padx=10)


class Graphic(tk.Frame):
    def __init__(self, x, y, title, master=None, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.figure = matplotlib.figure.Figure(figsize=(5, 5), dpi=100, tight_layout=True)
        ax = self.figure.add_subplot(111)
        ax.plot(x[:12], y[:12], marker='o', mfc="r", mec="r", markersize=5, linestyle='-', linewidth=2,
                color='b', label=r"$\ данные о преступлениях $")  # Задаем стиль линии и точек
        if len(x) > 12:
            ax.plot(x[11:], y[11:], marker='o', mfc="m", mec="c", markersize=5, linestyle='--', linewidth=2,
                    color='c', label=r"$\ предсказание $")  # Задаем стиль линии и точек
        ax.set_xlabel("Года", fontsize=12)
        ax.set_ylabel("Количество преступлений", fontsize=12)
        ax.grid(color="r", linewidth=1.0)
        ax.legend()

        self.canvas = FigureCanvasTkAgg(self.figure, master=self)
        self.canvas.draw()
        tk.Label(self, text=title, width=100, wraplength=260).pack()

        self.canvas.get_tk_widget().pack(anchor="center", expand=1, fill=tk.X)
        NavigationToolbar2Tk(self.canvas, self)
        self.canvas.get_tk_widget().pack(anchor="center", expand=1, fill=tk.BOTH)


def get_crime_prediction(x, y: list, prediction_years_count):
    new_x = [i for i in range(2023, 2023 + prediction_years_count)]
    predicted_y = []
    new_y = y.copy()
    for i in range(prediction_years_count):
        new_y.append(sum(new_y[-12:]) // 12)
        predicted_y.append(sum(new_y[-12:]) // 12)
    return new_x, predicted_y


def show_prediction_window():
    years = simpledialog.askinteger("Предсказание", "Введите количество лет для предсказания:")
    if years is None:
        return -1
    if not years:
        years = 0
    try:
        years = int(years)
        if years < 0:
            tk.messagebox.showerror("Ошибка", "Введите неотрицательное число!")
        else:
            return years
    except ValueError:
        tk.messagebox.showerror("Ошибка", "Нужно ввести число!")
    return years


def plot_crime_trend(root, df: pandas.DataFrame, filter_value):
    current_graphic_id = 0
    current_df = df

    def get_current_graphic_id(action):
        nonlocal current_graphic_id
        current_graphic_id = current_graphic_id
        graphics[current_graphic_id].pack_forget()
        if action == "<":
            if current_graphic_id - 1 > 0:
                current_graphic_id -= 1
            else:
                current_graphic_id = len(graphics) - 1
        elif action == ">":
            if current_graphic_id + 1 < len(graphics):
                current_graphic_id += 1
            else:
                current_graphic_id = 0

    def show_graphics(graphics, action):
        get_current_graphic_id(action)
        graphics[current_graphic_id].pack(
            side=tk.TOP, fill=tk.X, expand=True)
        window.eval('tk::PlaceWindow . center')

    if filter_value:
        current_df = df[df['object_name'] == filter_value]
        if current_df.empty:
            tk.messagebox.showinfo(title="Информация",
                                   message="Такого региона нет в базе данных")
            return
    else:
        tk.messagebox.showinfo(title="Информация",
                               message="Выберите конкретный регион, графики по которому вы \nхотите увидеть")
        return

    prediction_years_count = show_prediction_window()
    if prediction_years_count == -1:
        return

    window = tk.Tk()
    window.eval('tk::PlaceWindow . center')
    window.attributes("-topmost", True)
    window.title(f"Crime trends: {filter_value}")
    graphics = []

    left_button = tk.Button(window, text="<", width=10,
                            command=lambda: show_graphics(graphics, "<"))
    left_button.pack(side=tk.LEFT, fill=tk.Y)

    right_button = tk.Button(window, text=">", width=10,
                             command=lambda: show_graphics(graphics, ">"))
    right_button.pack(side=tk.RIGHT, fill=tk.Y)

    # Построение графика зависимости от года
    graphic_count = 1
    for row_id, row in current_df.iterrows():
        x = []
        y = []
        for name, value in row.items():
            if "year" in name:
                x.append(int(name[4:]))
                y.append(value)
        predicted_x, predicted_y = get_crime_prediction(x, y, prediction_years_count)
        x = x + predicted_x
        y = y + predicted_y
        old_title = row["indicator_name"]
        new_title_with_line_breaks = ""
        for i in range(len(old_title)):
            if i + 1 % 66 == 0:
                new_title_with_line_breaks += "\n"
            new_title_with_line_breaks += old_title[i]
        graphics.append(
            Graphic(x, y, title=f"{graphic_count}\{len(current_df)} {new_title_with_line_breaks}", master=window))
        graphic_count += 1

    show_graphics(graphics, "")


def analyze_crime_trend(df):
    # Добавление столбца 'total_crimes', который представляет сумму всех преступлений за все года
    df["total_crimes"] = df.filter(like="year").sum(axis=1)

    # Группировка данных по типу преступлений и региону
    grouped_data = df.groupby(['comment', 'object_name']).sum()
    max_crimes_per_type = grouped_data.groupby('comment').agg({'total_crimes': 'max'})
    regions_with_max_crimes = {}
    for comment, max_crimes in max_crimes_per_type.iterrows():
        max_crime_row = grouped_data.loc[(grouped_data.index.get_level_values('comment') == comment) &
                                         (grouped_data['total_crimes'] == max_crimes['total_crimes'])]
        regions_with_max_crimes[comment] = max_crime_row.index.get_level_values('object_name').tolist()

    def show_selected_crime(event):
        selected_crime = crime_combobox.get()
        max_crimes = max_crimes_per_type.loc[selected_crime, 'total_crimes']
        regions = ", ".join(regions_with_max_crimes[selected_crime])
        crime_info_label.config(text=f"Максимальное количество преступлений: {max_crimes}\n"
                                     f"Регионы с максимальным количеством преступлений: {regions}")

    result_window = tk.Toplevel(root)
    result_window.title("Результаты анализа")


    # Установка положения и размеров окна по центру экрана
    result_window.geometry("600x200+450+300")

    crime_combobox = ttk.Combobox(result_window, values=max_crimes_per_type.index.tolist(), width=150)

    crime_combobox.bind("<<ComboboxSelected>>", show_selected_crime)
    crime_combobox.pack(pady=10)

    crime_info_label = tk.Label(result_window, text="")
    crime_info_label.pack(pady=10)





root = tk.Tk()
root.title("Crime Analysis App")
root.eval('tk::PlaceWindow . center')
open_button = tk.Button(root, text="Open Crime Data", command=open_file)
open_button.pack(pady=20)

root.mainloop()
