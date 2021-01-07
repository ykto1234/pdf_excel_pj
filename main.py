import os
import traceback
import settings
import pdf
import excel

if __name__ == '__main__':

    try:
        print('処理を開始します')
        print('設定ファイル（setting.xlsx）を読み込みます')
        # 共通設定の読み込み
        common_setting_list = settings.read_settings('setting.xlsx', '共通設定', 1, 'A:I')
        PDF_INPUT_PATH = common_setting_list.at[0, 'PDF_INPUT_PATH']
        PDF_OUTPUT_PATH = common_setting_list.at[0, 'PDF_OUTPUT_PATH']
        INPUT_TARGET_X = int(common_setting_list.at[0, 'INPUT_TARGET_X'])
        INPUT_TARGET_Y = int(common_setting_list.at[0, 'INPUT_TARGET_Y'])
        PDF_TMP_FILENAME1 = common_setting_list.at[0, 'PDF_TMP_FILENAME1']
        PDF_TMP_FILENAME2 = common_setting_list.at[0, 'PDF_TMP_FILENAME2']
        TMP_DIR_PATH = common_setting_list.at[0, 'TMP_DIR_PATH']
        EXCEL_TMP_FILENAME = common_setting_list.at[0, 'EXCEL_TMP_FILENAME']
        EXCEL_INPUT_PATH = common_setting_list.at[0, 'EXCEL_INPUT_PATH']

        print('----------------------------------------')

        # エリア設定の読み込み
        area_setting_list = settings.read_settings('setting.xlsx', 'エリア設定', 2, 'A:H')

        for i in range(0, len(area_setting_list)):
            print(str(i+1) + 'つ目のファイルの処理を開始します')

            INPUT_TEXT = str(area_setting_list.at[i, 'INPUT_TEXT'])
            PDF_INPUT_FILENAME = area_setting_list.at[i, 'PDF_INPUT_FILENAME']
            EXCEL_INPUT_FILENAME1 = area_setting_list.at[i, 'EXCEL_INPUT_FILENAME1']
            EXCEL_INPUT_SHEETNAME1 = area_setting_list.at[i, 'EXCEL_INPUT_SHEETNAME1']
            # 列番号を指定して地点番号を読み込めるように修正
            COL_NUM = int(area_setting_list.at[i, 'COL_NUM'])
            #START_ROW = int(area_setting_list.at[i, 'START_ROW'])
            #END_ROW = int(area_setting_list.at[i, 'END_ROW'])
            EXCEL_INPUT_FILENAME2 = area_setting_list.at[i, 'EXCEL_INPUT_FILENAME2']
            EXCEL_INPUT_SHEETNAME2 = area_setting_list.at[i, 'EXCEL_INPUT_SHEETNAME2']
            PDF_MERGE_FILENAME = area_setting_list.at[i, 'PDF_MERGE_FILENAME']

            # PDFにテキストを追加する
            pdf.insert_text_pdf(
                PDF_INPUT_PATH + '/' + PDF_INPUT_FILENAME,
                TMP_DIR_PATH + '/' + PDF_TMP_FILENAME1,
                INPUT_TEXT,
                INPUT_TARGET_X,
                INPUT_TARGET_Y)
            print('PDFに掲載日時の追加が完了しました')

            # 地点番号のExcelを開いて地点番号をコピー
            numbers = []
            numbers = excel.copy_location_number(
                EXCEL_INPUT_PATH + '/' + EXCEL_INPUT_FILENAME1,
                EXCEL_INPUT_SHEETNAME1,
                COL_NUM,
                1  # 固定で１行から地点番号をコピー
            )
            print('地点番号のコピーが完了しました')

            # 地点番号から需要者名を呼び出すExcelを開いて貼り付け
            excel.paste_location_number(
                numbers,
                EXCEL_INPUT_PATH + '/' + EXCEL_INPUT_FILENAME2,
                EXCEL_INPUT_SHEETNAME2,
                TMP_DIR_PATH + '/' + EXCEL_TMP_FILENAME
            )
            print('地点番号の貼り付けが完了しました')

            # ExcelからPDFに変換して保存
            excel.excel_to_pdf(
                TMP_DIR_PATH + '/' + EXCEL_TMP_FILENAME,
                EXCEL_INPUT_SHEETNAME2,
                TMP_DIR_PATH + '/' + PDF_TMP_FILENAME2
            )
            print('ExcelからPDFへの変換が完了しました')

            # 作成した２つのPDFを結合
            pdf.merge_pdf(
                TMP_DIR_PATH + '/',
                PDF_OUTPUT_PATH + '/' + PDF_MERGE_FILENAME,
                PDF_TMP_FILENAME1,
                PDF_TMP_FILENAME2
            )
            print('PDFの結合が完了しました')
            print('出力先：' + os.getcwd() + '\\' + PDF_OUTPUT_PATH + '\\' + PDF_MERGE_FILENAME)
            print(str(i+1) + 'つ目のファイルの処理が完了しました')
            print('----------------------------------------')

        #『続行するには何かキーを押してください . . .』と表示させる
        print('処理が完了しました')
        os.system('PAUSE')

    except Exception as err:
        print('処理が失敗しました')
        print(err)
        print(traceback.format_exc())
        #『続行するには何かキーを押してください . . .』と表示させる
        os.system('PAUSE')

