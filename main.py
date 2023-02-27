import streamlit as st
import pandas as pd
from io import BytesIO


def to_excel(data_frame):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    data_frame.to_excel(writer, index=False, sheet_name='Sheet1')
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    format1 = workbook.add_format({'num_format': '0.00'})
    worksheet.set_column('A:A', None, format1)
    writer.save()
    processed_data = output.getvalue()
    return processed_data


with st.expander("IS exceļa pārbaude"):
    uploaded_file_excels_checkup = st.file_uploader("Ievietot IS excel failu", type=["csv", "xlsx"])

    df_delivery_spec = pd.DataFrame()

    if uploaded_file_excels_checkup is not None:
        if uploaded_file_excels_checkup.type == "text/csv":
            df_delivery_spec = pd.read_excel(uploaded_file_excels_checkup, sheet_name="3. DELIVERY SPEC.", dtype=str)
        else:
            df_delivery_spec = pd.read_excel(uploaded_file_excels_checkup, sheet_name="3. DELIVERY SPEC.", dtype=str)

    if uploaded_file_excels_checkup is not None:

        weight = st.checkbox(f"Rādīt svaru")
        detail_type = st.checkbox(f"Rādīt apakšelementu")
        # position = st.checkbox(f"Rādīt pozīciju")

        df_delivery_spec.drop([f"Unnamed: {v}" for v in range(0, 2)], axis=1, inplace=True)
        # df_delivery_spec.drop([f"Unnamed: {v}" for v in range(3, 5)], axis=1, inplace=True)

        if not detail_type:
            df_delivery_spec.drop(["Unnamed: 5"], axis=1, inplace=True)

        # if not position:
        df_delivery_spec.drop(["Unnamed: 4"], axis=1, inplace=True)

        df_delivery_spec.drop(["Unnamed: 3", "Unnamed: 7", "Unnamed: 10", "Unnamed: 11", "Unnamed: 12"], axis=1, inplace=True)
        df_delivery_spec.drop([f"Unnamed: {v}" for v in range(13, 28)], axis=1, inplace=True)
        df_delivery_spec.drop(["Unnamed: 31", "Unnamed: 32", "Unnamed: 33"], axis=1, inplace=True)
        df_delivery_spec.drop([f"Unnamed: {v}" for v in range(35, 46)], axis=1, inplace=True)
        df_delivery_spec.drop([0, 1], axis=0, inplace=True)

        df_delivery_spec.fillna(" ", inplace=True)

        df_delivery_spec['Unnamed: 2'] = df_delivery_spec['Unnamed: 2'].map({"Jā": 1, "Nē": 0})
        df_delivery_spec = df_delivery_spec[df_delivery_spec["Unnamed: 2"] > 0]

        df_delivery_spec["Unnamed: 9"] = df_delivery_spec["Unnamed: 9"].str.replace(' ', '')
        df_delivery_spec["Unnamed: 28"] = df_delivery_spec["Unnamed: 28"].str.replace(' ', '')
        df_delivery_spec["Unnamed: 29"] = df_delivery_spec["Unnamed: 29"].str.replace(' ', '')
        df_delivery_spec["Unnamed: 30"] = df_delivery_spec["Unnamed: 30"].str.replace(' ', '')

        if not weight:
            df_delivery_spec.drop(["Unnamed: 34"], axis=1, inplace=True)

        if weight:
            df_delivery_spec["Unnamed: 34"] = df_delivery_spec["Unnamed: 34"].str.replace(' ', '')

        # st.dataframe(df_delivery_spec)
        df_delivery_spec.drop_duplicates(inplace=True)
        df_delivery_spec['Column1_Value_Counts'] = df_delivery_spec['Unnamed: 6'].map(df_delivery_spec['Unnamed: 6'].value_counts())
        df_final = df_delivery_spec[df_delivery_spec["Column1_Value_Counts"] > 1].sort_values(by="Unnamed: 6")
        df_final.drop(["Column1_Value_Counts", "Unnamed: 2"], axis=1, inplace=True)

        df_final = df_final.rename(columns={"Unnamed: 6": "Marķējums",
                                            "Unnamed: 8": "Nosaukums",
                                            "Unnamed: 9": "Krāsa",
                                            "Unnamed: 28": "Platums",
                                            "Unnamed: 29": "Augstums",
                                            "Unnamed: 30": "Biezums",
                                            "Unnamed: 34": "Svars"})

        if weight:
            df_final = df_final.rename(columns={"Unnamed: 34": "Svars"})

        if detail_type:
            df_final = df_final.rename(columns={"Unnamed: 5": "Apakšelements"})

        # if position:
        #     df_final = df_final.rename(columns={"Unnamed: 4": "Pozīcija"})

        st.dataframe(df_final)


with st.expander("Meklēt vērtības"):
    df_main_values = pd.DataFrame()
    df_lookup_values = pd.DataFrame()

    uploaded_file = st.file_uploader("Ievietot failu", type=["csv", "xlsx"])

    if uploaded_file is None:
        with open("lookup_values_template.xlsx", "rb") as template_excel_lookup:
            st.download_button(
                label="Lejuplādēt template",
                data=template_excel_lookup,
                file_name="lookup_values_template.xlsx"
            )

    if uploaded_file is not None:
        if uploaded_file.type == "text/csv":
            df_main_values = pd.read_excel(uploaded_file, sheet_name="main_values")
            df_lookup_values = pd.read_excel(uploaded_file, sheet_name="look_up_values", index_col=0,
                                             dtype=str)
        else:
            df_main_values = pd.read_excel(uploaded_file, sheet_name="main_values")
            df_lookup_values = pd.read_excel(uploaded_file, sheet_name="look_up_values", index_col=0,
                                             dtype=str)

    if uploaded_file is not None:

        value_dictionary = df_lookup_values.to_dict('index')
        main_values_list = df_main_values["Main values"].tolist()

        new_df_dict = {
            "Main values": [],
            "value 1": [],
            "value 2": [],
            "value 3": [],
            "value 4": [],
            "value 5": []
        }

        new_df = pd.DataFrame(new_df_dict)

        for main_value in main_values_list:
            if main_value in value_dictionary.keys():
                new_df.loc[len(new_df.index)] = [main_value, value_dictionary[main_value]["value 1"],
                                                 value_dictionary[main_value]["value 2"],
                                                 value_dictionary[main_value]["value 3"],
                                                 value_dictionary[main_value]["value 4"],
                                                 value_dictionary[main_value]["value 5"]]
            else:
                new_df.loc[len(new_df.index)] = [main_value, "", "", "", "", ""]

        new_df.fillna("", inplace=True)

        st.write("Rezultāts")
        st.dataframe(new_df)

        st.download_button(
                    label="Lejuplādēt rezultātu",
                    data=to_excel(new_df),
                    file_name='result.xlsx'
                )
