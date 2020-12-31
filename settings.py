import pandas as pd

def read_settings(file_path, sheet_name, header_idx, cols):
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_idx, usecols=cols)
    # データフレームから空白の値を含む行を削除する
    print(df)
    df_formatted = df.dropna(how='any', axis=0)
    print(df_formatted)
    return df_formatted



