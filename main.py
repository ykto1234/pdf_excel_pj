import os
import settings
import pdf
import excel

if __name__ == '__main__':

    try:
        # 共通設定の読み込み
        common_setting_list = settings.read_settings('setting.xlsx', '共通設定', 1)
        PDF_INPUT_PATH = common_setting_list.at[0, 'PDF_INPUT_PATH']
        PDF_OUTPUT_PATH = common_setting_list.at[0, 'PDF_OUTPUT_PATH']
        INPUT_TARGET_X = int(common_setting_list.at[0, 'INPUT_TARGET_X'])
        INPUT_TARGET_Y = int(common_setting_list.at[0, 'INPUT_TARGET_Y'])
        PDF_TMP_FILENAME1 = common_setting_list.at[0, 'PDF_TMP_FILENAME1']
        PDF_TMP_FILENAME2 = common_setting_list.at[0, 'PDF_TMP_FILENAME2']
        TMP_DIR_PATH = common_setting_list.at[0, 'TMP_DIR_PATH']
        EXCEL_TMP_FILENAME = common_setting_list.at[0, 'EXCEL_TMP_FILENAME']
        EXCEL_INPUT_PATH = common_setting_list.at[0, 'EXCEL_INPUT_PATH']

        # エリア設定の読み込み
        area_setting_list = settings.read_settings('setting.xlsx', 'エリア設定', 2)

        for i in range(0, len(area_setting_list)):
            INPUT_TEXT = str(area_setting_list.at[i, 'INPUT_TEXT'])
            PDF_INPUT_FILENAME = area_setting_list.at[i, 'PDF_INPUT_FILENAME']
            EXCEL_INPUT_FILENAME1 = area_setting_list.at[i, 'EXCEL_INPUT_FILENAME1']
            EXCEL_INPUT_SHEETNAME1 = area_setting_list.at[i, 'EXCEL_INPUT_SHEETNAME1']
            COL_NUM = int(area_setting_list.at[i, 'COL_NUM'])
            START_ROW = int(area_setting_list.at[i, 'START_ROW'])
            END_ROW = int(area_setting_list.at[i, 'END_ROW'])
            EXCEL_INPUT_FILENAME2 = area_setting_list.at[i, 'EXCEL_INPUT_FILENAME2']
            EXCEL_INPUT_SHEETNAME2 = area_setting_list.at[i, 'EXCEL_INPUT_SHEETNAME2']
            PDF_MERGE_FILENAME = area_setting_list.at[i, 'PDF_MERGE_FILENAME']

            # PDFにテキストを追加する
            print('処理を開始します')
            pdf.insert_text_pdf(
                PDF_INPUT_PATH + PDF_INPUT_FILENAME,
                TMP_DIR_PATH + PDF_TMP_FILENAME1,
                INPUT_TEXT,
                INPUT_TARGET_X,
                INPUT_TARGET_Y)
            print('PDFに掲載日時の追加が完了しました')

            # 地点番号のExcelを開いて地点番号をコピー
            numbers = []
            numbers = excel.copy_location_number(
                EXCEL_INPUT_PATH + EXCEL_INPUT_FILENAME1,
                EXCEL_INPUT_SHEETNAME1,
                COL_NUM,
                START_ROW,
                END_ROW
            )
            print('地点番号のコピーが完了しました')

            # 地点番号から需要者名を呼び出すExcelを開いて貼り付け
            excel.paste_location_number(
                numbers,
                EXCEL_INPUT_PATH + EXCEL_INPUT_FILENAME2,
                EXCEL_INPUT_SHEETNAME2,
                TMP_DIR_PATH + EXCEL_TMP_FILENAME
            )
            print('地点番号の貼り付けが完了しました')

            # ExcelからPDFに変換して保存
            excel.excel_to_pdf(
                TMP_DIR_PATH + EXCEL_TMP_FILENAME,
                EXCEL_INPUT_SHEETNAME2,
                TMP_DIR_PATH + PDF_TMP_FILENAME2
            )
            print('ExcelからPDFへの変換が完了しました')

            # 作成した２つのPDFを結合
            pdf.merge_pdf(
                TMP_DIR_PATH,
                PDF_OUTPUT_PATH + PDF_MERGE_FILENAME
            )
            print('PDFの結合が完了しました')
            print('出力先：' + os.getcwd() + '\\' + PDF_OUTPUT_PATH + PDF_MERGE_FILENAME)

        #『続行するには何かキーを押してください . . .』と表示させる
        os.system('PAUSE')

    except Exception as err:
        print('処理が失敗しました')
        print(err)
        #『続行するには何かキーを押してください . . .』と表示させる
        os.system('PAUSE')

