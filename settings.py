import pandas as pd

def read_settings(file_path, sheet_name, header_idx, cols):
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_idx, usecols=cols)
    # データフレームから全ての列が空白となっている行を削除する
    df_formatted = df.dropna(how='all', axis=0)
    return df_formatted



