from sqlalchemy import create_engine, text
import pandas as pd
import pandas.io.sql as psql
from render_galderma.main.constants import *
from render_galderma.helper.data_table_helper import save_df_to_excel
from render_galderma.helper.data_table_helper import ExtractedData

# query data từ db và lưu lại nếu có thay đổi data ở db
# engine = create_engine("postgresql+psycopg2://dateam:d4T3am_m3tr1cVn@13.214.246.7:54325/datademo")
# sql = f"""
# SELECT * FROM {serum_table}
# """
# df_table = psql.read_sql(sql, engine)
# df_table = df_table.fillna('')
# with pd.ExcelWriter(f'../input_template/excel_of_table/{serum_table}.xlsx', engine='xlsxwriter',
#                             engine_kwargs={'options': {'strings_to_urls': False}}) as writer:
#     df_table.to_excel(writer, index=False)
# nếu data không thay đổi thì có thể comment lại phần trên và sử dụng data đã lưu lại từ lần trước bằng calamine engine để tăng tốc
# df_table = pd.read_excel(f'../input_template/excel_of_table/{serum_table}.xlsx', engine='calamine')
# thay data raw
df_table = pd.read_excel(f'/Users/anhkieu/workspace/Metric/da-team-common/gaderma 2025/data_all_pkg_gop_brand_20250214_171531.xlsx', engine='calamine')
df_table = df_table.fillna('')

# # Các bước tiền xử lý data

df_table = df_table[(df_table['cate'] != 'x') & (df_table['is_fake_sales'] != 1)]
# & (df_table['price_range'] != '<400K')
# col_to_drop = [col for col in df_table.columns if 'firsthalf' in col or 'revenue_2024.1' in col]
# df_table = df_table.drop(columns=col_to_drop)

#
# df_table_l400 = pd.read_excel(f'../input_template/excel_of_table/{serum_table}.xlsx', engine='calamine')
# tinh toan
# A
df_total_serum = ExtractedData(df_table, start_date, end_date)
# df_serum_l400 = ExtractedData(df_table_l400, start_date, end_date)
df_overview_total_market = df_total_serum.caculate_overview()
df_overview_channel = df_total_serum.caculate_overview_by_group_column('platform')
df_overview_pricerange = df_total_serum.caculate_overview_by_group_column('price_range')
# df_l400_pricerange = df_serum_l400.caculate_overview_by_group_column('price_range')
# df_overview_pricerange = pd.concat([df_overview_pricerange, df_l400_pricerange])
df_overview_function = df_total_serum.caculate_overview_by_group_column('partner_function')

#B
df_tp_b_1_1 = df_total_serum.caculate_tp_marketsize()
df_brand_b_1_1 = df_total_serum.caculate_salevalue_by_group_column('partner_brand')
df_client_b_1_1 = df_total_serum.caculate_client()
df_tp_b_1_2 = df_total_serum.caculate_tp_marketsize(dict_include={'platform':['Shopee']})
df_brand_b_1_2 = df_total_serum.caculate_salevalue_by_group_column('partner_brand', dict_include={'platform':['Shopee']})
df_client_b_1_2 = df_total_serum.caculate_client(dict_include={'platform':['Shopee']})

df_tp_b_1_3 = df_total_serum.caculate_tp_marketsize(dict_include={'platform':['Lazada']})
df_brand_b_1_3 = df_total_serum.caculate_salevalue_by_group_column('partner_brand', dict_include={'platform':['Lazada']})
df_client_b_1_3 = df_total_serum.caculate_client(dict_include={'platform':['Lazada']})
df_tp_b_1_4 = df_total_serum.caculate_tp_marketsize(dict_include={'platform':['Tiktok Shop']})
df_brand_b_1_4 = df_total_serum.caculate_salevalue_by_group_column('partner_brand', dict_include={'platform':['Tiktok Shop']})
df_client_b_1_4 = df_total_serum.caculate_client(dict_include={'platform':['Tiktok Shop']})
df_tp_b_1_5 = df_total_serum.caculate_tp_marketsize(dict_include={'platform':['Tiki']})
df_brand_b_1_5 = df_total_serum.caculate_salevalue_by_group_column('partner_brand', dict_include={'platform':['Tiki']})
df_client_b_1_5 = df_total_serum.caculate_client(dict_include={'platform':['Tiki']})
#C
df_tp_c_1_2 = df_total_serum.caculate_tp_marketsize(dict_include={'partner_function':['Trắng da']})
df_brand_c_1_2 = df_total_serum.caculate_salevalue_by_group_column('partner_brand', dict_include={'partner_function':['Trắng da']})
df_client_c_1_2 = df_total_serum.caculate_client(dict_include={'partner_function':['Trắng da']})
df_tp_c_1_3 = df_total_serum.caculate_tp_marketsize(dict_include={'partner_function':['Cấp ẩm']})
df_brand_c_1_3 = df_total_serum.caculate_salevalue_by_group_column('partner_brand', dict_include={'partner_function':['Cấp ẩm']})
df_client_c_1_3 = df_total_serum.caculate_client(dict_include={'partner_function':['Cấp ẩm']})
df_tp_c_1_4 = df_total_serum.caculate_tp_marketsize(dict_include={'partner_function':['Chống lão hoá']})
df_brand_c_1_4 = df_total_serum.caculate_salevalue_by_group_column('partner_brand', dict_include={'partner_function':['Chống lão hoá']})
df_client_c_1_4 = df_total_serum.caculate_client(dict_include={'partner_function':['Chống lão hoá']})
df_tp_c_1_5 = df_total_serum.caculate_tp_marketsize(dict_include={'partner_function':['Giảm mụn']})
df_brand_c_1_5 = df_total_serum.caculate_salevalue_by_group_column('partner_brand', dict_include={'partner_function':['Giảm mụn']})
df_client_c_1_5 = df_total_serum.caculate_client(dict_include={'partner_function':['Giảm mụn']})
df_tp_c_1_6 = df_total_serum.caculate_tp_marketsize(dict_include={'partner_function':['Phục hồi']})
df_brand_c_1_6 = df_total_serum.caculate_salevalue_by_group_column('partner_brand', dict_include={'partner_function':['Phục hồi']})
df_client_c_1_6 = df_total_serum.caculate_client(dict_include={'partner_function':['Phục hồi']})
#D
dict_brand_player = {
 "L'Oréal": 'loreal',
 'La Roche-Posay': 'laroche',
 'Garnier': 'garnier'
}
dict_df_results = {}
for brand, brand_label in dict_brand_player.items():
     dict_df_results[f'tp_us_{brand_label}'] = df_total_serum.caculate_tp_us(brand)
     dict_df_results[f'{brand_label}_platform'] = df_total_serum.caculate_overview_by_group_column('platform', dict_include={'partner_brand': [brand]})
     df_model_share = df_total_serum.caculate_lst_product_model(brand)
     lst_productmodel = df_model_share['product_model'].to_list()
     dict_df_results[f'infor_autocf_{brand_label}'] = len(lst_productmodel)
     for idx, model in enumerate(lst_productmodel, start=1):
         dict_df_results[f'autocf_share_{brand_label}_model_{idx}'] = df_total_serum.caculate_model_share(brand, model)
         dict_df_results[f'autocf_platformbytime_{brand_label}_model_{idx}'] = df_total_serum.calculate_monthly_revenue_by_platform(dict_include={'partner_brand':[brand], 'product_model': [model]})





config_to_table_serum = {'Serum': {'overview_total_market': df_overview_total_market, 'overview_channel': df_overview_channel,
                  'overview_pricerange': df_overview_pricerange, 'overview_function': df_overview_function,
                 'tp_b_1_1': df_tp_b_1_1, 'tp_b_1_2': df_tp_b_1_2,
                  'tp_b_1_3': df_tp_b_1_3, 'tp_b_1_4': df_tp_b_1_4, 'tp_b_1_5': df_tp_b_1_5,
                  'brand_b_1_1': df_brand_b_1_1, 'brand_b_1_2': df_brand_b_1_2, 'brand_b_1_3': df_brand_b_1_3,
                  'brand_b_1_4': df_brand_b_1_4, 'brand_b_1_5': df_brand_b_1_5, 'client_b_1_1': df_client_b_1_1,
                  'client_b_1_2': df_client_b_1_2, 'client_b_1_3': df_client_b_1_3,
                  'client_b_1_4': df_client_b_1_4, 'client_b_1_5': df_client_b_1_5, 'tp_c_1_2': df_tp_c_1_2,
                  'tp_c_1_3': df_tp_c_1_3, 'tp_c_1_4': df_tp_c_1_4, 'tp_c_1_5': df_tp_c_1_5,
                  'tp_c_1_6': df_tp_c_1_6, 'brand_c_1_2': df_brand_c_1_2,
                  'brand_c_1_3': df_brand_c_1_3, 'brand_c_1_4': df_brand_c_1_4, 'brand_c_1_5': df_brand_c_1_5,
                  'brand_c_1_6': df_brand_c_1_6, 'client_c_1_2': df_client_c_1_2,
                  'client_c_1_3': df_client_c_1_3, 'client_c_1_4': df_client_c_1_4,
                  'client_c_1_5': df_client_c_1_5, 'client_c_1_6': df_client_c_1_6,'tp_c_1_1': df_tp_b_1_1,'brand_c_1_1': df_brand_b_1_1, 'client_c_1_1': df_client_b_1_1,
                  }}

config_to_table_serum['Serum'].update(dict_df_results)