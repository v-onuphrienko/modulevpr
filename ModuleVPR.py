import streamlit as st
import pandas as pd
import openpyxl

import io

def load_excel(file, header_row):
    wb = openpyxl.load_workbook(filename=file, data_only=True)
    sheet = wb.active

    data = list(sheet.values)
    cols = data[header_row]  # Заголовки начинаются с этой строки
    data = data[header_row + 1:]  # Данные начинаются после строки с заголовками

    # Удаление столбцов, где название 'None' или 'Unnamed'
    cols = [col if col is not None and not str(col).startswith("Unnamed") else None for col in cols]

    # Создание DataFrame из данных
    df = pd.DataFrame(data, columns=cols)

    # Удаление столбцов, где все значения None
    df.dropna(axis=1, how='all', inplace=True)

    return df

def get_dublicate_columns(df):
    duplicated_columns = df.columns[df.columns.duplicated()]
    return duplicated_columns

def merge_excel_files(df1, df2, key_columns, target_columns):
    missing_columns = [col for col in target_columns if col not in df1.columns]

    if missing_columns:
        merged_df = pd.merge(df1, df2[missing_columns + key_columns], on=key_columns, how='left')
    else:
        merged_df = df1

    return merged_df

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    processed_data = output.getvalue()
    return processed_data



st.title("ФУНКЦИЯ ВПР")

file1 = st.file_uploader("Загрузите 1 файл, в который планируете добавить столбцы", type=['xlsx'])
file2 = st.file_uploader("Загрузите 2 файл из которого планируете добавить столбцы", type=['xlsx'])

if file1 and file2:
    original_file_name = file1.name
    header_row_file1 = st.number_input("С какой строчки Названия столбцов в 1м файле", min_value=0, value=0, step=1)
    header_row_file2 = st.number_input("С какой строчки Названия столбцов во 2м файле", min_value=0, value=0, step=1)

    df1 = load_excel(file1, header_row_file1)
    df2 = load_excel(file2, header_row_file2)

    key_column_left = st.selectbox("Выберите ключ из первого файла", options=df1.columns)
    key_column_right = st.selectbox("Выберите ключ из второго файла", options=df2.columns)

    target_columns = st.multiselect("Какие столбцы добавить", options=df2.columns.difference(df1.columns))

    if st.button("Использовать ВПР"):
        # Выполнение слияния
        result_df = pd.merge(df1, df2[target_columns + [key_column_right]], left_on=key_column_left,
                             right_on=key_column_right, how='left')

        # Удаление дублирующегося ключевого столбца из второго файла, если он присутствует
        if key_column_right in result_df.columns and key_column_right != key_column_left:
            result_df.drop(key_column_right, axis=1, inplace=True)

        st.write(result_df)

        # Кнопка для скачивания результата слияния
        merge_result = to_excel(result_df)
        st.download_button(label="Скачать результат слияния", data=merge_result, file_name="merged_result.xlsx",
                           mime="application/vnd.ms-excel")
