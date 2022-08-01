import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from pandas import DataFrame


def read_data(path: str) -> DataFrame:
    """
    Функция чтения файла из данной папки
    :param path: путь к файлу
    :return: Файл графика ремонтов подается на функцию make_max_date
    """
    df_grafic = pd.read_excel(path, header=1)
    df_grafic = df_grafic.dropna(subset=['Оборудование', 'Контролируемое сечение'])
    df_grafic.loc[:, 'Время ремонта. Количество дней'] = None
    df_grafic.loc[:, 'Время ремонта. Начало'] = pd.to_datetime(df_grafic['Время ремонта. Начало'], format='%d.%m.%y')
    df_grafic.loc[:, 'Время ремонта. Конец'] = pd.to_datetime(df_grafic['Время ремонта. Конец'], format='%d.%m.%y')
    df_grafic.loc[:, 'Время ремонта. Количество дней'] = pd.to_timedelta(df_grafic['Время ремонта. Количество дней'])
    return df_grafic.sort_values(by='Время ремонта. Начало', ascending=True)


def make_max_date(path: str) -> DataFrame:
    """
    Функция определения длительности ремонта
    :param path: путь к файлу
    :return: График с учетом длительности ремонтов
    """
    grafic = read_data(path)
    start = grafic['Время ремонта. Начало']
    finish = grafic['Время ремонта. Конец']
    grafic.loc[:, 'Время ремонта. Начало'] = start
    grafic.loc[:, 'Время ремонта. Конец'] = finish
    grafic.loc[:, 'Время ремонта. Количество дней'] = grafic['Время ремонта. Конец'] \
                                                      - grafic['Время ремонта. Начало']
    return grafic.rename(columns={"Контролируемое сечение": 'Контролируемое_сечение'})


def make_df_for_plotting(sechen: str, path: str) -> DataFrame:
    """
    Функция подготовки графика к построению диаграммы Ганта.
    :param path:
    :param sechen: Название сечения
    :return: График ремонтов для выбранного сечения с указанием ремонтов
    в смежных сечения для выбранного оборудования
    """
    df = make_max_date(path)
    df_outer = df.copy()
    df_sechen = df.query(f'Контролируемое_сечение ==  "{sechen}"')
    list_equipment = df_sechen['Оборудование'].unique()
    df_outer = df_outer.query('Оборудование.isin(@list_equipment)', engine='python')
    return pd.concat((df_sechen, df_outer), axis=0)


def graf_plotty(plot_graf: DataFrame) -> None:
    """
    Функция построения графика
    :param plot_graf: дата фрейм графика ремонтов для заданного сечения
    :return: Изображение графика ремонтов
    """
    from datetime import timedelta
    # дата начала ремонта
    proj_start = plot_graf['Время ремонта. Начало'].min()
    # количество дней с начала графика ренмонта по текущему сечению до старта ремонта ремонта текущего оборудования
    plot_graf['start_num'] = (plot_graf['Время ремонта. Начало'] - proj_start).dt.days
    # количество дней от старта ремонта по теукущему сечению до конца ремонта заанного оборудования
    plot_graf['end_num'] = (plot_graf['Время ремонта. Конец'] - proj_start).dt.days
    # длительность ремонта
    plot_graf['days_start_to_end'] = plot_graf.end_num - plot_graf.start_num

    fig, ax = plt.subplots(figsize=(40, 8), dpi=120)
    ax.barh(plot_graf["Оборудование"], plot_graf.days_start_to_end + 1, left=plot_graf.start_num, height=0.5)

    xticks = np.arange(0, plot_graf.end_num.max() + 2, 1)
    xticks_labels = pd.date_range(proj_start,
                                  end=(plot_graf['Время ремонта. Конец'] + timedelta(days=1)).max()).strftime(
        "%d/%m")
    xticks_minor = np.arange(0, plot_graf.end_num.max(), 1)
    ax.set_xticks(xticks)
    ax.set_xticks(xticks_minor, minor=True)
    plt.xticks(rotation=90, fontsize=14)
    ax.set_xticklabels(xticks_labels[::1])
    plt.yticks(fontsize=14)

    ax.grid(which='major',
            color='k',
            linewidth=0.8)

    ax.grid(which='minor',
            color='k',
            linestyle=':',
            linewidth=0.6)

    print(plot_graf['Контролируемое_сечение'].unique())
    ax.legend(plot_graf['Контролируемое_сечение'].unique(), loc='upper right', fontsize=14)
    plt.savefig('graf.jpg')


def make_sechen_list(grafic: DataFrame) -> list:
    """
    Функция получения оборудования входящего в сечение
    :param grafic: график ремонтов
    :return: список сечений
    """
    return grafic['Контролируемое_сечение'].unique()
