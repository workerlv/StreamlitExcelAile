import streamlit as st
import pandas as pd
from collections import Counter

df = pd.DataFrame()

uploaded_file = st.file_uploader("Izvēlēties failu", type=["csv", "xlsx"])

if uploaded_file is not None:
    if uploaded_file.type == "text/csv":
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file, header=None)

if uploaded_file is not None:
    st.header("Esošais excelis")

    df.drop(index=[0, 1, 2, 3], axis=0, inplace=True)
    df.drop(columns=[0, 2, 7, 8, 9, 10, 11, 12, 13], axis=1, inplace=True)

    df_concat = df[1] + "?" + df[4].astype(str) + "?" + df[5].astype(str) + "?" + df[6].astype(str)
    df_concat.drop(index=[4], axis=0, inplace=True)
    df_concat.dropna(inplace=True)
    df_set = sorted(set(df_concat.tolist()))
    df_set = pd.DataFrame(df_set)

    splited = df_set[0].str.split("?", n=3, expand=True)

    count_values = Counter(splited[0])
    count_values = [key for key, val in count_values.items() if val > 1]

    error_df = splited[splited[0].isin(count_values)]
    error_df.sort_values(by=[0])

    st.write(df)
    st.write(error_df)
