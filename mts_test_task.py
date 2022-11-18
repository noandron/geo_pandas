import os
import time

import geopandas as gpd
import pandas as pd
import xlwings as xw
from geopandas.tools import overlay

# делаем декоратор

def timeit(verbose):
    def outer(func):
        def wrapper(*args, **kwargs):
            start = time.time()
            result = func(*args, **kwargs)
            end = time.time()
            duration = round(end - start, 2)
            if verbose is True:
                print(
                    f'start_time = {time.strftime("%H:%M:%S", time.localtime(start))}, end_time = {time.strftime("%H:%M:%S", time.localtime(end))}, duration = {duration} seconds')
            elif verbose is False:
                print(f'duration = {duration} seconds')
            return result
        return wrapper
    return outer

# перереходим к task_1

def saving_results(dataframe, plot):  # функция для сохранения результатов для task1 в .xlsx
    os.chdir('C://PythonProjects//geo_pandas//results/task1/')
    with xw.App(visible=False) as app:
        wb = xw.Book()
        name = 'task1_results'
        sht = wb.sheets[0]
        sht.name = name
        sht['A1'].value = dataframe  # записываем df
        fig = plot.get_figure()
        sht.pictures.add(
            fig,
            left=sht.range('G1').left,
            height=300, width=300
        )
        wb.save(f'{name}.xlsx')
        wb.close()


@timeit(verbose=True)
def pivot_and_plot(file):  # функция для расчетов по заданию
    os.chdir('C://PythonProjects//geo_pandas//test_polygons')
    gdf = gpd.read_file(file)
    pivot = gdf[['ID', 'CATEGORY']].drop_duplicates().groupby(
        'CATEGORY').count()
    pivot['PERCENT'] = round(pivot['ID'] / pivot['ID'].sum() * 100, 2)
    pivot = pivot.rename(columns={'ID': 'POLYGON_COUNT'})
    plot = pivot.plot(kind='pie', y='PERCENT', autopct='%1.0f%%', legend=False)
    saving_results(pivot, plot)  # сохраняем результаты в файл

# исполнение task_1
pivot_and_plot('test.shp')

# переходим к task_2

os.chdir('C://PythonProjects//geo_pandas//test_polygons')
@timeit(verbose=False)
def overlay_proccesing(file):
    # читаем shp file и переходим в 3857 чтобы получить метры вместо градусов
    gdf = gpd.read_file(file).to_crs(3857)
    # сразу посчитаем площадь каждого полигона
    gdf['S'] = gdf.area
    # загоняем в список нумерацию строк
    poly_list = list(range(len(gdf)))
    # сделаем копию gdf чтобы считать перекрытия
    gdf_copy = gdf.copy()
    # и сделаем копию листа со списком строк, чтобы потом удалять из него уже пройденные в цикле строки
    check_list = poly_list.copy()
    cnt = 1  # для нумерации новых полигонов
    gdf_new_list = []  # тут будем собирать все перекрывающиеся полигоны

    for i in poly_list:
        # тут сразу удаляем первую строку, чтобы не сравнивать самого с собой
        check_list.remove(i)
        for k in check_list:
            # сравниваем строки каждую с каждой в двух gdf, фильтруем строки по индексам, ищем пересечения
            # сдлеано именно по строкам потому, что есть полигон с одним именем но двумя разными значениями (POLY_2)
            overlay_df = overlay(gdf.filter(items=[i], axis=0), gdf_copy.filter(
                items=[k], axis=0), how='intersection')
            if len(overlay_df) > 0:  # если пересение есть, то добавляем вычисления
                overlay_df['ID_NEW'] = f'NEW_POLY_{cnt}'
                overlay_df['S_NEW'] = overlay_df['geometry'].area
                overlay_df['VALUE_NEW'] = (overlay_df['S_NEW']/overlay_df['S_1'] * overlay_df['VALUE_1']) + (
                    overlay_df['S_NEW'] / overlay_df['S_2'] * overlay_df['VALUE_2'])
                # добавляем полученный массив в общий список
                gdf_new_list.append(overlay_df)
                cnt += 1  # чтобы next ID обновился

    over_gdf = pd.concat(gdf_new_list)  # объединям все df в один
    over_gdf = over_gdf[['ID_NEW', 'VALUE_NEW', 'geometry']].rename(
        columns={'ID_NEW': 'ID', 'VALUE_NEW': 'VALUE'}).reset_index(drop=True)
    os.chdir('C://PythonProjects//geo_pandas//results/task2/')
    over_gdf.to_file('new_shape.shp')   # сохраняем в файл

# исполнение task_2
overlay_proccesing('test.shp')
