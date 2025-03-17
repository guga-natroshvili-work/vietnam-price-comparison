import pandas as pd

def enter_data(ro_file: str, samo_file: str, ro_sheet: str, samo_sheet: str):
    ro_df = pd.read_excel(ro_file, sheet_name=ro_sheet)
    summary = pd.read_excel(samo_file, sheet_name=samo_sheet)
    return ro_df, summary


if __name__ == '__main__':
    ro_df, summary = enter_data('Test.xlsx', 'Test.xlsx', 'RO', 'SAMO')
    print(ro_df.head())
    print(summary.head())