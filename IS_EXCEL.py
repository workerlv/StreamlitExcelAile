import pandas as pd


class IsExcel:
    def __init__(self, is_excel_file):
        self.is_excel_file = is_excel_file

        if self.is_excel_file.type == "text/csv":
            delivery_spec_sheet_df = pd.read_excel(self.is_excel_file, sheet_name="3. DELIVERY SPEC.", dtype=str)
        else:
            delivery_spec_sheet_df = pd.read_excel(self.is_excel_file, sheet_name="3. DELIVERY SPEC.", dtype=str)

        delivery_spec_sheet_df.drop([0, 1], axis=0, inplace=True)
        delivery_spec_sheet_df.fillna(" ", inplace=True)

        data_set = pd.DataFrame()
        data_set["Aktīvs"] = delivery_spec_sheet_df["Unnamed: 2"].map({"Jā": 1, "Nē": 0})
        data_set["Pozīcija"] = delivery_spec_sheet_df["Unnamed: 4"]
        data_set["Apakšelements"] = delivery_spec_sheet_df["Unnamed: 5"]
        data_set["Marķējums"] = delivery_spec_sheet_df["Unnamed: 6"]
        data_set["Nosaukums"] = delivery_spec_sheet_df["Unnamed: 8"]
        data_set["Krāsa"] = delivery_spec_sheet_df["Unnamed: 9"].str.replace(' ', '')
        data_set["Platums"] = delivery_spec_sheet_df["Unnamed: 28"].str.replace(' ', '')
        data_set["Augstums"] = delivery_spec_sheet_df["Unnamed: 29"].str.replace(' ', '')
        data_set["Biezums"] = delivery_spec_sheet_df["Unnamed: 30"].str.replace(' ', '')
        data_set["Svars"] = delivery_spec_sheet_df["Unnamed: 34"].str.replace(' ', '')
        data_set.fillna(" ", inplace=True)

        self.full_data_set = data_set[data_set["Aktīvs"] > 0]

    def check_excel_by_marking(self, weight, detail_type):
        marking_check = self.full_data_set.copy()

        marking_check.drop(["Aktīvs", "Pozīcija"], axis=1, inplace=True)

        if not detail_type:
            marking_check.drop(["Apakšelements"], axis=1, inplace=True)

        if not weight:
            marking_check.drop(["Svars"], axis=1, inplace=True)

        marking_check.drop_duplicates(inplace=True)
        marking_check['Column_Value_Counts'] = marking_check['Marķējums'].map(marking_check['Marķējums'].value_counts())
        df_final = marking_check[marking_check["Column_Value_Counts"] > 1].sort_values(by="Marķējums")
        df_final.drop(["Column_Value_Counts"], axis=1, inplace=True)

        return df_final

    def check_excel_by_position(self):
        position_check = pd.DataFrame()
        position_check["Elementa numurs / marķējums"] = self.full_data_set["Pozīcija"] + " / " + self.full_data_set["Marķējums"]
        position_check['Cik reizes atkārtojas'] = position_check['Elementa numurs / marķējums'].map(position_check['Elementa numurs / marķējums'].value_counts())

        position_check = position_check[position_check["Cik reizes atkārtojas"] > 1]
        position_check.drop_duplicates(inplace=True)
        return position_check

    def test_columns(self):
        return self.full_data_set
