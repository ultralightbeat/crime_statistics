import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import matplotlib.pyplot as plt

def open_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        df = pd.read_excel(file_path)
        display_data(df)
        plot_crime_trend(df)
        analyze_crime_trend(df)

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

    # Создание Treeview
    tree = ttk.Treeview(display_frame)

    # Определение колонок
    columns_to_display = ['object_name', 'object_level', 'year2011', 'year2012', 'year2013', 'year2014',
                          'year2015', 'year2016', 'year2017', 'year2018', 'year2019', 'year2020', 'year2021', 'year2022', 'comment']
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

    # Вставка данных
    for index, row in df.iterrows():
        values = [row[col] for col in columns_to_display]
        tree.insert("", tk.END, values=values)

    # Фильтр
    filter_frame = tk.Frame(display_frame)
    filter_frame.pack(pady=10)

    filter_label = tk.Label(filter_frame, text="Фильтр по объекту:")
    filter_label.pack(side=tk.LEFT)

    unique_objects = sorted(set(df['object_name']))
    filter_combobox = ttk.Combobox(filter_frame, values=unique_objects)
    filter_combobox.pack(side=tk.LEFT, padx=10)

    filter_button = tk.Button(filter_frame, text="Фильтр", command=filter_data)
    filter_button.pack(side=tk.LEFT)

    # Упаковка Treeview
    tree.pack(expand=tk.YES, fill=tk.BOTH)


def plot_crime_trend(df):
    # Построение графика зависимости от года
    plt.figure(figsize=(10, 6))
    for col in df.columns:
        if col.startswith("year"):
            plt.plot(df["section_name"], df[col], label=col)
    plt.xlabel("Crime Type")
    plt.ylabel("Number of Crimes")
    plt.xticks(rotation=90)
    plt.legend()
    plt.title("Crime Trend Over the Years")
    plt.show()

def analyze_crime_trend(df):
    # Анализ изменения уровня преступности за 10 лет
    df["total_crimes"] = df.filter(like="year").sum(axis=1)
    df["change"] = (df["year2022"] - df["year2011"]) / df["year2011"] * 100
    max_decrease = df.loc[df["change"].idxmin()]
    max_increase = df.loc[df["change"].idxmax()]
    print("Type of crime with the largest decrease over 10 years:", max_decrease["section_name"])
    print("Type of crime with the largest increase over 10 years:", max_increase["section_name"])

def extrapolate_crime_trend(df, n_years):
    # Статистическое прогнозирование методом экстраполяции по скользящей средней
    smoothed_data = df.filter(like="year").rolling(window=3, axis=1).mean()
    extrapolated_data = smoothed_data.mean(axis=0) * (1 + df.shape[0] / n_years)
    plt.figure(figsize=(8, 5))
    plt.plot(df.columns[8:], df.mean()[8:], label="Actual Data")
    plt.plot(extrapolated_data.index, extrapolated_data, label="Extrapolated Data", linestyle='--')
    plt.xlabel("Year")
    plt.ylabel("Number of Crimes")
    plt.title("Crime Trend Extrapolation for Next {} Years".format(n_years))
    plt.legend()
    plt.show()

root = tk.Tk()
root.title("Crime Analysis App")

open_button = tk.Button(root, text="Open Crime Data", command=open_file)
open_button.pack(pady=20)

root.mainloop()
