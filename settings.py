import pandas as pd

def read_settings(file_path, sheet_name, header_idx):
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_idx)
    # データフレームから空白の値を含む行を削除する
    df_formatted = df.dropna(how='any', axis=0)
    return df_formatted



