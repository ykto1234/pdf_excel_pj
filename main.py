import settings
import pdf
import excel
import os

if __name__ == '__main__':

    try:
        # PDFにテキストを追加する
        pdf.insert_text_pdf(
            settings.PDF_INPUT_PATH + settings.PDF_INPUT_FILENAME,
            settings.WORK_DIR_PATH + settings.PDF_OUTPUT_FILENAME1,
            settings.INPUT_TEXT,
            settings.INPUT_TARGET_X,
            settings.INPUT_TARGET_Y)

        # 地点番号のExcelを開いて地点番号をコピー
        numbers = []
        numbers = excel.copy_location_number(
            settings.EXCEL_INPUT_PATH + settings.EXCEL_INPUT_FILENAME1,
            settings.EXCEL_INPUT_SHEETNAME1
        )
        print('----------------------地点番号コピー完了----------------------')

        # 地点番号から需要者名を呼び出すExcelを開いて貼り付け
        excel.paste_location_number(
            numbers,
            settings.EXCEL_INPUT_PATH + settings.EXCEL_INPUT_FILENAME2,
            settings.EXCEL_INPUT_SHEETNAME2,
            settings.WORK_DIR_PATH + settings.WORK_FILENAME
        )
        print('----------------------地点番号貼り付け完了----------------------')

        # ExcelからPDFに変換して保存
        excel.excel_to_pdf(
            settings.WORK_DIR_PATH + settings.WORK_FILENAME,
            settings.EXCEL_INPUT_SHEETNAME2,
            settings.WORK_DIR_PATH + settings.PDF_OUTPUT_FILENAME2
        )
        print('----------------------ExcelからPDFに変換完了----------------------')

        # 作成した２つのPDFを結合
        pdf.merge_pdf(
            settings.WORK_DIR_PATH,
            settings.PDF_OUTPUT_PATH + settings.PDF_MERGE_FILENAME
        )
        print('----------------------PDF結合完了----------------------')
        print('出力先：' + os.getcwd() + '\\' + settings.PDF_OUTPUT_PATH + settings.PDF_MERGE_FILENAME)
        #『続行するには何かキーを押してください . . .』と表示させる
        os.system('PAUSE')

    except Exception as err:
        print('----------------------処理が失敗しました----------------------')
        print(err)
        #『続行するには何かキーを押してください . . .』と表示させる
        os.system('PAUSE')

