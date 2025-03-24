from datetime import datetime
from dateutil.relativedelta import relativedelta
import pandas as pd

import re


def save_df_to_excel(df, output):
    now = datetime.now()
    now_str = now.strftime("%Y%m%d_%H%M%S")
    writer = pd.ExcelWriter(f'{output}_{now_str}.xlsx', engine='xlsxwriter',
                            engine_kwargs={'options': {'strings_to_urls': False}})
    df.to_excel(writer, index=False)
    writer.close()


def extract_price_range(segment):
    match = re.findall(r'(\d+)', segment)
    if '<' in segment:
        return (-1, int(match[0]))
    elif '>' in segment:
        return (int(match[0]), float('inf'))
    elif len(match) == 2:
        return (int(match[0]), int(match[1]))
    else:
        return (float('inf'), float('inf'))


def sort_price_segment(df_sorted, column_name):
    df_sorted["_sort_key"] = df_sorted[column_name].apply(extract_price_range)
    df_sorted = df_sorted.sort_values(by="_sort_key").drop(columns=["_sort_key"])
    return df_sorted


def get_start_end_date_of_py_pp(start_date: str, end_date: str):
    start_dt = datetime.strptime(start_date, "%Y%m")
    end_dt = datetime.strptime(end_date, "%Y%m")
    start_py_dt = start_dt - relativedelta(years=1)
    end_py_dt = end_dt - relativedelta(years=1)
    diff_months = (end_dt.year - start_dt.year) * 12 + (end_dt.month - start_dt.month)
    end_pp_dt = start_dt - relativedelta(months=1)
    start_pp_dt = end_pp_dt - relativedelta(months=diff_months)
    start_py = start_py_dt.strftime("%Y%m")
    end_py = end_py_dt.strftime("%Y%m")
    start_pp = start_pp_dt.strftime("%Y%m")
    end_pp = end_pp_dt.strftime("%Y%m")
    return int(start_py), int(end_py), int(start_pp), int(end_pp)


def expand_benefit(df, value_col):
    benefit_cols = [col for col in df.columns if col.startswith("benefit_")]
    records = []
    for _, row in df.iterrows():
        for col in benefit_cols:
            if pd.notna(row[col]) and row[col] != "":
                record = {k: row[k] for k in value_col}
                record["benefit_type"] = col
                record["benefit"] = row[col]
                records.append(record)
    df_long = pd.DataFrame(records)
    return df_long


class ExtractedData:
    """
    0: revenue
    1: sale
    """

    def __init__(self, data: pd.DataFrame, start_date: str, end_date: str):
        self.data = data
        start = int(start_date)
        end = int(end_date)
        all_columns = self.data.columns
        self.revenue_columns = [
            col for col in all_columns
            if col.startswith('revenue') and
               start <= int(col.split('_')[-1]) <= end
        ]
        # print(f"revenue_columns: {self.revenue_columns}")
        self.sale_columns = [
            col for col in all_columns
            if col.startswith('sale_') and
               start <= int(col.split('_')[-1]) <= end
        ]
        # print(f"sale_columns: {self.sale_columns}")
        start_py, end_py, start_pp, end_pp = get_start_end_date_of_py_pp(start_date, end_date)
        if f"sale_{start_py}" in all_columns and f"sale_{end_py}" in all_columns:
            self.sale_columns_py = [
                col for col in all_columns
                if col.startswith('sale') and
                   start_py <= int(col.split('_')[-1]) <= end_py
            ]
            # print(f"sale_columns_py: {self.sale_columns_py}")
        if f"sale_{start_pp}" in all_columns and f"sale_{end_pp}" in all_columns:
            self.sale_columns_pp = [
                col for col in all_columns
                if col.startswith('sale') and
                   start_pp <= int(col.split('_')[-1]) <= end_pp
            ]
            # print(f"sale_columns_pp: {self.sale_columns_pp}")
        if f"revenue_{start_py}" in all_columns and f"revenue_{end_py}" in all_columns:
            self.revenue_columns_py = [
                col for col in all_columns
                if col.startswith('revenue') and
                   start_py <= int(col.split('_')[-1]) <= end_py
            ]
            # print(f"revenue_columns_py: {self.revenue_columns_py}")
        if f"revenue_{start_pp}" in all_columns and f"revenue_{end_pp}" in all_columns:
            self.revenue_columns_pp = [
                col for col in all_columns
                if col.startswith('revenue') and
                   start_pp <= int(col.split('_')[-1]) <= end_pp
            ]
            # print(f"revenue_columns_pp: {self.revenue_columns_pp}")
        self.total_revenue = self.data[self.revenue_columns].sum(axis=1).sum()
        self.total_unit = self.data[self.sale_columns].sum(axis=1).sum()

    def caculate_overview(self, lst_filter_columns=None):
        if lst_filter_columns:
            df_filter = self.data[self.data['cate'].isin(lst_filter_columns)]
        else:
            df_filter = self.data.copy()
        sales_value = df_filter[self.revenue_columns].sum(axis=1).sum()
        unit_sale = df_filter[self.sale_columns].sum(axis=1).sum()
        sales_value_py = df_filter[self.revenue_columns_py].sum(axis=1).sum()
        unit_sale_py = df_filter[self.sale_columns_py].sum(axis=1).sum()
        growth_sales_value = (sales_value - sales_value_py) / sales_value_py if sales_value_py != 0 else 0
        growth_unit_sale = (unit_sale - unit_sale_py) / unit_sale_py if unit_sale_py != 0 else 0
        df_result = pd.DataFrame({
            'Sales Value': [sales_value],
            'Unit sales': [unit_sale],
            'PY (%)\nBy: Sales Value': [growth_sales_value],
            'PY (%)\nBy: Unit Sales': [growth_unit_sale]
        })
        return df_result


    def caculate_overview_by_group_column(self, group_column, dict_include=None, dict_exclude=None, top_n=None):
        df_filtered = self.data.copy()
        if dict_include:
            for key, value in dict_include.items():
                df_filtered = df_filtered[df_filtered[key].isin(value)]
        if dict_exclude:
            for key, value in dict_exclude.items():
                df_filtered = df_filtered[~df_filtered[key].isin(value)]
        df_grouped = df_filtered.groupby(group_column)[self.revenue_columns + self.revenue_columns_py + self.sale_columns + self.sale_columns_py ].sum()

        df_grouped["Sales Value"] = df_grouped[self.revenue_columns].sum(axis=1)
        df_grouped["Unit Sales"] = df_grouped[self.sale_columns].sum(axis=1)
        if not top_n and group_column == 'partner_brand':
            top_n = 10
        if top_n and len(df_grouped) > top_n:
            df_grouped = df_grouped.sort_values(by="Sales Value", ascending=False).head(top_n)
        df_grouped["Sales Value PY"] = df_grouped[self.revenue_columns_py].sum(axis=1)
        df_grouped["PY (%)\nBy: Sales Value"] = df_grouped.apply(
            lambda row: (row["Sales Value"] - row["Sales Value PY"]) / row["Sales Value PY"]
            if row["Sales Value PY"] != 0 else 0,
            axis=1
        )
        df_grouped["Share (%)\nBy: Sales Value"] = df_grouped[
                                              "Sales Value"] / self.total_revenue if self.total_revenue != 0 else 0

        df_grouped["Unit Sales PY"] = df_grouped[self.sale_columns_py].sum(axis=1)
        df_grouped["PY (%)\nBy: Unit Sales"] = df_grouped.apply(
            lambda row: (row["Unit Sales"] - row["Unit Sales PY"]) / row["Unit Sales PY"]
            if row["Unit Sales PY"] != 0 else 0,
            axis=1
        )
        df_grouped["Share (%)\nBy: Unit Sales"] = df_grouped[
                                              "Unit Sales"] / self.total_unit if self.total_unit != 0 else 0

        df_grouped = df_grouped.reset_index()
        df_result = df_grouped[[group_column] + ["Sales Value", "Unit Sales","PY (%)\nBy: Sales Value","PY (%)\nBy: Unit Sales","Share (%)\nBy: Sales Value" ,"Share (%)\nBy: Unit Sales"]]
        if 'price_range' in df_result.columns:
            df_result = sort_price_segment(df_result, 'price_range')
        return df_result


    def caculate_tp_marketsize(self,dict_include=None, dict_exclude=None):
        df_filtered = self.data.copy()
        if dict_include:
            for key, value in dict_include.items():
                df_filtered = df_filtered[df_filtered[key].isin(value)]
        if dict_exclude:
            for key, value in dict_exclude.items():
                df_filtered = df_filtered[~df_filtered[key].isin(value)]
        sales_value = df_filtered[self.revenue_columns].sum(axis=1).sum()
        sales_value_py = df_filtered[self.revenue_columns_py].sum(axis=1).sum()
        growth_sales_value = (sales_value - sales_value_py) / sales_value_py if sales_value_py != 0 else 0
        df_result = pd.DataFrame({
            'Market size': [sales_value],
            'PY (%)': [growth_sales_value],
        })
        return df_result

    def caculate_salevalue_by_group_column(self, group_column, dict_include=None, dict_exclude=None, top_n=None):
        df_filtered = self.data.copy()
        if dict_include:
            for key, value in dict_include.items():
                df_filtered = df_filtered[df_filtered[key].isin(value)]
        if dict_exclude:
            for key, value in dict_exclude.items():
                df_filtered = df_filtered[~df_filtered[key].isin(value)]
        df_grouped = df_filtered.groupby(group_column)[self.revenue_columns + self.revenue_columns_py].sum()

        df_grouped["Sales Value"] = df_grouped[self.revenue_columns].sum(axis=1)
        total_revenue_filtered = df_grouped["Sales Value"].sum()
        if group_column == 'partner_brand':
            df_grouped = df_grouped[df_grouped.index != "Chưa biết"]
        if not top_n and group_column == 'partner_brand':
            top_n = 10
        if top_n and len(df_grouped) > top_n:
            df_grouped = df_grouped.sort_values(by="Sales Value", ascending=False).head(top_n)
        df_grouped["Sales Value PY"] = df_grouped[self.revenue_columns_py].sum(axis=1)
        df_grouped["PY (%)"] = df_grouped.apply(
            lambda row: (row["Sales Value"] - row["Sales Value PY"]) / row["Sales Value PY"]
            if row["Sales Value PY"] != 0 else 0,
            axis=1
        )
        df_grouped["Share (%)"] = df_grouped[
                                              "Sales Value"] / total_revenue_filtered if total_revenue_filtered != 0 else 0
        df_grouped = df_grouped.reset_index()
        df_result = df_grouped[[group_column] + ["Sales Value","PY (%)", "Share (%)"]]
        if 'price_range' in df_result.columns:
            df_result = sort_price_segment(df_result, 'price_range')
        return df_result

    def caculate_client(self, dict_include=None, dict_exclude=None):
        df_filtered = self.data.copy()
        if dict_include:
            for key, value in dict_include.items():
                df_filtered = df_filtered[df_filtered[key].isin(value)]
        if dict_exclude:
            for key, value in dict_exclude.items():
                df_filtered = df_filtered[~df_filtered[key].isin(value)]
        df_grouped = df_filtered.groupby('partner_brand')[self.revenue_columns + self.revenue_columns_py].sum()

        df_grouped["Sales Value"] = df_grouped[self.revenue_columns].sum(axis=1)
        total_revenue_filtered = df_grouped["Sales Value"].sum()
        df_grouped = df_grouped.reset_index()
        df_grouped = df_grouped[df_grouped['partner_brand'] == 'Cetaphil']
        df_grouped["Sales Value PY"] = df_grouped[self.revenue_columns_py].sum(axis=1)
        df_grouped["PY (%)"] = df_grouped.apply(
            lambda row: (row["Sales Value"] - row["Sales Value PY"]) / row["Sales Value PY"]
            if row["Sales Value PY"] != 0 else 0,
            axis=1
        )
        df_grouped["Share (%)"] = df_grouped[
                                      "Sales Value"] / total_revenue_filtered if total_revenue_filtered != 0 else 0
        df_result = df_grouped[["Sales Value", "PY (%)", "Share (%)"]]
        print()
        return df_result

    def caculate_tp_us(self, brand,dict_include=None, dict_exclude=None):
        df_filtered = self.data[self.data["partner_brand"] == brand]
        if dict_include:
            for key, value in dict_include.items():
                df_filtered = df_filtered[df_filtered[key].isin(value)]
        if dict_exclude:
            for key, value in dict_exclude.items():
                df_filtered = df_filtered[~df_filtered[key].isin(value)]

        sales_value = df_filtered[self.revenue_columns].sum().sum()
        sales_value_py = df_filtered[self.revenue_columns_py].sum().sum()
        unit_sales = df_filtered[self.sale_columns].sum().sum()
        unit_sales_py = df_filtered[self.sale_columns_py].sum().sum()
        py_percentage = (sales_value - sales_value_py) / sales_value_py if sales_value_py != 0 else 0
        py_unit_percentage = (unit_sales - unit_sales_py) / unit_sales_py if unit_sales_py != 0 else 0
        df_result = pd.DataFrame([
            ["Unit sales", unit_sales, "Doanh số", sales_value],
            ["PY (%)", py_unit_percentage, "Tăng trưởng (%)",py_percentage]
        ])
        return df_result
    def caculate_lst_product_model(self,brand):
        df_filtered = self.data[self.data['partner_brand'] == brand]
        df_grouped = df_filtered.groupby('product_model')[self.revenue_columns].sum()
        df_sorted = df_grouped.sum(axis=1).reset_index()
        df_sorted = df_sorted.sort_values(by=0, ascending=False)
        return df_sorted
    def caculate_model_share(self,brand,model):
        df_filtered = self.data[self.data['partner_brand'] == brand]
        total_sum = df_filtered[self.revenue_columns].sum(axis=1).sum()
        df_filtered_model = df_filtered[df_filtered['product_model'] == model]
        model_sum = df_filtered_model[self.revenue_columns].sum(axis=1).sum()
        share = model_sum / total_sum if total_sum != 0 else model_sum
        df_result = pd.DataFrame([
            [f"Model {model.lower().title()}", model_sum],
            ["Share (%)", share]
        ])
        return df_result
    def calculate_monthly_revenue_by_platform(self,dict_include=None, dict_exclude=None):
        df_filtered = self.data.copy()
        if dict_include:
            for key, value in dict_include.items():
                df_filtered = df_filtered[df_filtered[key].isin(value)]
        if dict_exclude:
            for key, value in dict_exclude.items():
                df_filtered = df_filtered[~df_filtered[key].isin(value)]

        revenue_mapping = {col: f"T{col[-2:]}" for col in self.revenue_columns}
        df_grouped = df_filtered.groupby("platform")[self.revenue_columns].sum().reset_index()
        df_grouped.rename(columns=revenue_mapping, inplace=True)
        return df_grouped

