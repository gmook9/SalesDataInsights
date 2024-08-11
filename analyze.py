import os
import pandas as pd
from glob import glob
import calendar

def load_and_combine_data(input_folder):
    all_files = glob(os.path.join(input_folder, "*.csv"))
    df_list = []
    
    for file in all_files:
        df = pd.read_csv(file, parse_dates=['Date of sale', 'Date of listing'])
        df['Year'] = df['Date of sale'].dt.year
        df['Month'] = df['Date of sale'].dt.month
        df['Total'] = pd.to_numeric(df['Total'].replace('[\$,]', '', regex=True), errors='coerce')
        df['Depop fee'] = pd.to_numeric(df['Depop fee'].replace('[\$,]', '', regex=True), errors='coerce')

        df_list.append(df)
    
    combined_df = pd.concat(df_list, ignore_index=True)
    return combined_df

def analyze_yearly_data(df):
    yearly_data = {}

    for year in df['Year'].unique():
        yearly_df = df[df['Year'] == year]
        monthly_summary = yearly_df.groupby('Month').agg({
            'Total': 'sum',
            'Depop fee': 'sum',
            'Size': lambda x: x.value_counts().to_dict(),
            'State': lambda x: (x.str.lower().str.contains('ca|california')).sum()
        }).reset_index()
        
        monthly_summary.rename(columns={
            'Total': 'Total Sales',
            'Depop fee': 'Total Fees Paid',
            'State': 'Shipments inside CA'
        }, inplace=True)

        monthly_summary['Net Sales'] = monthly_summary['Total Sales'] - monthly_summary['Total Fees Paid']

        sizes = ['XS', 'S', 'M', 'L']
        for size in sizes:
            monthly_summary[size] = monthly_summary['Size'].apply(lambda s: s.get(size, 0))
        
        monthly_summary.drop(columns=['Size'], inplace=True)
        monthly_summary['Month'] = monthly_summary['Month'].apply(lambda x: calendar.month_name[x])
        yearly_data[year] = monthly_summary
    
    return yearly_data

def summarize_customers(df):
    customer_summary = df.groupby('Buyer').agg({
        'Total': 'sum',
        'Date of sale': 'max',
        'Buyer': 'count'
    }).rename(columns={'Buyer': '# of orders', 'Total': 'Amount Spent', 'Date of sale': 'Last Purchase Date'})
    
    return customer_summary.reset_index()

def save_output(yearly_data, customer_summary, output_folder="output"):
    csv_folder = os.path.join(output_folder, "csv")
    spreadsheet_folder = os.path.join(output_folder, "spreadsheets")
    os.makedirs(csv_folder, exist_ok=True)
    os.makedirs(spreadsheet_folder, exist_ok=True)

    for year, data in yearly_data.items():
        monthly_csv_file = os.path.join(csv_folder, f"Yearly_Summary_{year}.csv")
        data.to_csv(monthly_csv_file, index=False)
        
        yearly_totals = data.sum(numeric_only=True)
        yearly_totals.to_csv(monthly_csv_file, mode='a', header=True)

        monthly_excel_file = os.path.join(spreadsheet_folder, f"Yearly_Summary_{year}.xlsx")
        with pd.ExcelWriter(monthly_excel_file, engine='xlsxwriter') as writer:
            data.to_excel(writer, sheet_name='Monthly Summary', index=False)
            yearly_totals.to_excel(writer, sheet_name='Yearly Totals', index=False)

    customer_csv_file = os.path.join(csv_folder, "Customer_Summary.csv")
    customer_summary.to_csv(customer_csv_file, index=False)

    customer_excel_file = os.path.join(spreadsheet_folder, "Customer_Summary.xlsx")
    customer_summary.to_excel(customer_excel_file, index=False)

def main():
    input_folder = "input"
    output_folder = "output"
    
    combined_df = load_and_combine_data(input_folder)
    yearly_data = analyze_yearly_data(combined_df)
    customer_summary = summarize_customers(combined_df)
    
    save_output(yearly_data, customer_summary, output_folder)

if __name__ == "__main__":
    main()
