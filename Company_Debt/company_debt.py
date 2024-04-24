#import libraries
import pandas as pd
import datetime
import wget
import os

# Download file form the url
def download_file(url, output_path):
    try:
        wget.download(url, output_path)
        print(f"\nFile downloaded successfully: {output_path}\n")
    except:
        print(f"\nFailed to download the file\n")
        raise

# Convert csv to excel
def convert_csv_to_excel(csv_path, excel_path, code_list):
    try:
        df = pd.read_csv(csv_path, skiprows=3, encoding='big5')
        df = df[df['代號'].isin(code_list)]
        buy = '今日最佳買入殖利率(%)/百元價'
        sell = '今日最佳賣出殖利率(%)/百元價'
        df['AVG'] = df[[buy, sell]].astype(float).mean(axis = 1)
        df.to_excel(excel_path, index=False)
        print(f"\nFile converted to Excel successfully: {excel_path}\n")
    except:
        print(f"\nFailed to convert CSV to Excel\n")
        raise

# Remove original csv file
def remove_file(path):
    try:
        os.remove(path)
        print(f"\nFile removed successfully: {path}\n")
    except:
        print(f"\nFailed to remove file\n")

# Main function
if __name__ == '__main__':
    today = datetime.datetime.today().strftime('%Y%m%d')
    
    current_dir = os.path.dirname(os.path.abspath(__file__))
    url = 'https://www.tpex.org.tw/storage/bond_xml/ch/str_bond/fix/FixBondQS.csv'
    path = f'{current_dir}/FixBondQS_{today}'
    code_list = ['B99503', 'B99603', 'B99604', 'B99702']

    download_file(url, f'{path}.csv')
    convert_csv_to_excel(f'{path}.csv', f'{path}.xlsx', code_list)
    remove_file(f'{path}.csv')