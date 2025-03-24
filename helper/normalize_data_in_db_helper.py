from render_galderma.main.constants import *
import psycopg2

for sub_subcate, table_name in dict_sub_subcate.items():
    with psycopg2.connect(
            dbname=config_datademo.get('dbname'), user=config_datademo.get('user'),
            password=config_datademo.get('password'), host=config_datademo.get('host'), port=config_datademo.get('port')
    ) as conn:
        with conn.cursor() as cur:
            try:
                print(f'----start: {table_name}')
                add_column_query = f"""
                                                        ALTER TABLE {table_name} ADD COLUMN IF NOT EXISTS price_range TEXT;
                                                        """
                cur.execute(add_column_query)
                update_query = f"""
                                                            UPDATE {table_name}
                                                            SET price_range = CASE
                                                                WHEN price < 400000 THEN '<400K'
                                                                WHEN price >= 400000 AND price < 700000 THEN '400K-700K'
                                                                WHEN price >= 700000 AND price < 1000000 THEN '700K-1000K'
                                                                WHEN price >= 1000000 AND price < 1500000 THEN '1000K-1500K'
                                                                WHEN price >= 1500000 AND price < 3000000 THEN '1500K-3000K'
                                                                ELSE '>3000K'
                                                            END;
                                                            """
                cur.execute(update_query)
                print("Cột 'price_range' đã được tạo và cập nhật.")

                lst_initcap = ['partner_brand', 'platform', 'partner_function']
                for col in lst_initcap:
                    sql_init = f"""
                                                        UPDATE {table_name}
                                                        SET {col} = initcap({col})
                                            """
                    cur.execute(sql_init)
                conn.commit()
                print(f'----done: {table_name}')
            except:
                conn.rollback()
                print(f'----error: {table_name}')
                conn.close()
