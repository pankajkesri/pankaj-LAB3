import os
import sys
import datetime
import pandas as pd
import re

def main():
    sales_csv = get_sales_csv()
    orders_dir = create_orders_dir(sales_csv)
    process_sales_data(sales_csv, orders_dir)

# Get path of sales data CSV file from the command line
def get_sales_csv():
    # Check whether command line parameter provided
    # Check whether provide parameter is valid path of file
    if len(sys.argv) != 2:
        print("ERROR: provide the path to the sales data csv file.")
        sys.exit(1)
    
    sales_data_csv = sys.argv[1]
    if not os.path.isfile(sales_data_csv):
        print(f"ERROR: This file '{sales_data_csv}' does not exit.")
        sys.exit(1)

    return sales_data_csv

# Create the directory to hold the individual order Excel sheets
def create_orders_dir(sales_csv):
    # Get directory in which sales data CSV file resides
    sales_dir = os.path.dirname(sales_csv)
    # Determine the name and path of the directory to hold the order data files
    today_str = datetime.date.today().isoformat()
    orders_dir = os.path.join(sales_dir, f"Orders_{today_str}")
    # Create the order directory if it does not already exist
    if not os.path.exists(orders_dir):
        os.makedirs(orders_dir)
    
    return orders_dir 

# Split the sales data into individual orders and save to Excel sheets
def process_sales_data(sales_csv, orders_dir):
    # Import the sales data from the CSV file into a DataFrame
    df = pd.read_csv(sales_csv)
    # Insert a new "TOTAL PRICE" column into the DataFrame
    df['TOTAL PRICE'] = df['ITEM QUANTITY'] * df['ITEM PRICE']
    # Remove columns from the DataFrame that are not needed
    columns_to_keep = ['ORDER ID', 'ITEM NUMBER', 'PRODUCT CODE', 'PRODUCT LINE', 'ITEM QUANTITY', 'ITEM PRICE', 'TOTAL PRICE']
    df = df[columns_to_keep]   
    # Group the rows in the DataFrame by order ID
    grouped = df.groupby('ORDER ID')
    # For each order ID:
    for order_id, order_df in grouped:
        # Remove the "ORDER ID" column
        order_df = order_df.drop(columns=['ORDER ID'])
        # Sort the items by item number
        order_df = order_df.sort_values(by='ITEM NUMBER')
        # Append a "GRAND TOTAL" row
        grand_total = order_df['TOTAL PRICE'].sum()
        grand_total_row = pd.DataFrame({'ITEM NUMBER': [''], 'PRODUCT CODE': [''], 'PRODUCT LINE': [''], 'ITEM QUANTITY': [''], 'ITEM PRICE': [''], 'TOTAL PRICE': [grand_total]})
        order_df = pd.concat([order_df, grand_total_row], ignore_index=True)
        # Determine the file name and full path of the Excel sheet
        order_file = os.path.join(orders_dir, f"Order_{order_id}.xlsx")
        # Export the data to an Excel sheet
        with pd.ExcelWriter(order_file, engine='xlsxwriter') as writer:
            order_df.to_excel(writer, index=False, sheet_name=f'Order {order_id}')
        # Format the Excel sheet
        workbook = writer.book
        worksheet = writer.sheets[f'Order {order_id}']
        # Define format for the money columns
        money_format = workbook.add_format({'num_format': '$#,##0.00'})
        for col_num in range(5, 7):
            worksheet.set_column(col_num, col_num, None, money_format)
        # Format each colunm
        worksheet.set_column('A:A', 12)
        worksheet.set_column('B:B', 15)
        worksheet.set_column('C:C', 25)
        worksheet.set_column('D:D', 12)
        worksheet.set_column('E:E', 12)
        worksheet.set_column('F:F', 12)
        worksheet.set_column('G:G', 15)
        # close the sheet

if __name__ == '__main__':
    main()