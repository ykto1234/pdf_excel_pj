import os
import pathlib
from pathlib import Path
import PyPDF2
import fitz
from PyPDF2 import PdfFileWriter, PdfFileReader
from io import BytesIO
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.lib.units import mm

def insert_text_pdf(in_pdf_file, out_pdf_file, insert_text, target_x, target_y):
    """
    既存のPDFファイルに文字を挿入し、別名で出力します
    :param in_pdf_file:         挿入対象のPDFファイルのパス
    :param out_pdf_file:        挿入後のPDFファイルのパス
    :param insert_text:         挿入するテキスト
    :param target_x:            挿入するテキストのX位置
    :param target_y:            挿入するテキストのY位置
    :return:
    """
    buffer = BytesIO()
    # PDF新規作成
    p = canvas.Canvas(buffer, pagesize=A4)

    # フォントサイズ定義
    font_size = 14
    pdfmetrics.registerFont(UnicodeCIDFont('HeiseiKakuGo-W5')) # フォント
    p.setFont('HeiseiKakuGo-W5', font_size)

    # 挿入位置(mm指定)
    p.drawString(int(target_x)*mm, int(target_y)*mm, insert_text)
    p.showPage()
    p.save()

    # move to the beginning of the StringIO buffer
    buffer.seek(0)
    new_pdf = PdfFileReader(buffer)
    # read your existing PDF
    existing_pdf = PdfFileReader(open(in_pdf_file, 'rb'), strict=False)
    output = PdfFileWriter()
    # 既存PDFの1ページ目を読み取り
    page = existing_pdf.getPage(0)
    # 新規PDFにマージ
    page.mergePage(new_pdf.getPage(0))

    output.addPage(page)

    # フォルダが存在しない場合作成
    dirpath = os.path.split(os.path.abspath(out_pdf_file))
    if  not os.path.exists(dirpath[0]):
        os.mkdir(dirpath[0])

    output_stream = open(out_pdf_file, 'wb')
    output.write(output_stream)
    output_stream.close()

def merge_pdf(work_path, output_pdf_filepath):
    """
    指定したフォルダに存在するPDFファイルを結合する
    :param work_path:            PDFファイルが存在するフォルダのパス
    :param output_pdf_filepath:  結合後のPDFファイルの出力パス
    :return:
    """
    # 結合したい(結合元)PDFを取得
    curdir = os.getcwd() + '/' + work_path
    files = list(pathlib.Path(curdir).glob('*.pdf'))

    # 結合元PDFを並び替え
    sfiles = sorted(files)

    # 結合先のPDFを新規作成
    doc = fitz.open()

    # 結合元PDFを開く
    for file in sfiles:
        infile = fitz.open(file)

        # 結合先PDFと結合元PDFのページ番号を指定
        doc_lastPage = len(doc)
        infile_lastPage = len(infile)
        doc.insertPDF(infile, from_page=0, to_page=infile_lastPage, start_at=doc_lastPage, rotate=0)

        # 結合元PDFを閉じる
        infile.close()

    # フォルダが存在しない場合作成
    dirpath = os.path.split(os.path.abspath(output_pdf_filepath))
    if  not os.path.exists(dirpath[0]):
        os.mkdir(dirpath[0])

    # 結合先PDFを保存する
    filename = output_pdf_filepath
    doc.save(filename)
