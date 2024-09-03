import os
import pandas as pd
import snowflake.connector

def get_snowflake_connection():
    conn = snowflake.connector.connect(
        user=os.getenv('SNOWFLAKE_USER'),
        password=os.getenv('SNOWFLAKE_PASSWORD'),
        account=os.getenv('SNOWFLAKE_ACCOUNT'),
        warehouse=os.getenv('SNOWFLAKE_WAREHOUSE'),
        database=os.getenv('SNOWFLAKE_DATABASE'),
        schema=os.getenv('SNOWFLAKE_SCHEMA')
    )
    return conn

def query_snowflake(query):
    conn = get_snowflake_connection()
    try:
        cur = conn.cursor()
        cur.execute(query)
        df = cur.fetch_pandas_all()
        return df
    except Exception as e:
        print(f"An error occurred: {e}")
        return pd.DataFrame() 
    finally:
        conn.close()

def get_data(view):
    query_string = "SELECT * FROM " + view
    df = query_snowflake(query_string)
    return df

items_view = "DEPT_FINANCE.PUBLIC.BENS_MASTER_LOOKUP_ITEMS"
discounts_view = "DEPT_FINANCE.PUBLIC.BENS_MASTER_LOOKUP_DISCOUNTS"
employees_view = "DEPT_FINANCE.PUBLIC.BENS_MASTER_LOOKUP_EMPLOYEES"
consignors_view = "DEPT_FINANCE.PUBLIC.BENS_MASTER_LOOKUP_CONSIGNORS"
partners_view = "DEPT_FINANCE.PUBLIC.BENS_MASTER_LOOKUP_PARTNERS"
departments_view = "DEPT_FINANCE.PUBLIC.BENS_MASTER_LOOKUP_DEPARTMENTS"
customers_view = "DEPT_FINANCE.PUBLIC.BENS_MASTER_LOOKUP_CUSTOMERS"

items_df = get_data(items_view)
discounts_df = get_data(discounts_view)
employees_df = get_data(employees_view)
consignors_df = get_data(consignors_view)
partners_df = get_data(partners_view)
departments_df = get_data(departments_view)
departments_df.drop(columns=['Department Full Name'], inplace=True)
customers_df = get_data(customers_view)


with pd.ExcelWriter('Master Lookup Test.xlsx') as writer:
    customers_df.to_excel(writer, sheet_name='Customers', engine='xlsxwriter', index=False)
    employees_df.to_excel(writer, sheet_name='Employees', engine='xlsxwriter', index=False)
    partners_df.to_excel(writer, sheet_name='Partners', engine='xlsxwriter', index=False)
    departments_df.to_excel(writer, sheet_name='Departments', engine='xlsxwriter', index=False)
    items_df.to_excel(writer, sheet_name='Items', engine='xlsxwriter', index=False)
    consignors_df.to_excel(writer, sheet_name='Consignors', engine='xlsxwriter', index=False)
    discounts_df.to_excel(writer, sheet_name='Discounts', engine='xlsxwriter', index=False)
    
    

    
    
    
    




