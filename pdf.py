import settings
from pathlib import Path
import PyPDF2
from PyPDF2 import PdfFileMerger
import os
import pathlib
import fitz

def insert_text_pdf(in_pdf_file, out_pdf_file, insert_text, target_x, target_y):
    """
    既存のPDFファイルに文字を挿入し、別名で出力します
    :param pdf_file_path:       既存のPDFファイルパス
    :param insert_text:         挿入するテキスト
    :return:
    """
    from PyPDF2 import PdfFileWriter, PdfFileReader
    from io import BytesIO
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm

    buffer = BytesIO()
    # PDF新規作成
    p = canvas.Canvas(buffer, pagesize=A4)
    # 挿入位置(mm指定)
    #target_x, target_y = 200*mm, 400*mm
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
    # 出力名
    #output_name = "./tmp/pdf1.pdf"
    output_stream = open(out_pdf_file, 'wb')
    output.write(output_stream)
    output_stream.close()

def merge1():
    # フォルダ内のPDFファイル一覧
    pdf_dir = Path("./tmp")
    pdf_files = sorted(pdf_dir.glob("*.pdf"))

    # １つのPDFファイルにまとめる
    pdf_writer = PyPDF2.PdfFileWriter()
    for pdf_file in pdf_files:
        pdf_reader = PyPDF2.PdfFileReader(str(pdf_file))
        for i in range(pdf_reader.getNumPages()):
            pdf_writer.addPage(pdf_reader.getPage(i))

    # 保存ファイル名（先頭と末尾のファイル名で作成）
    merged_file = pdf_files[0].stem + "-" + pdf_files[-1].stem + ".pdf"

    # 保存
    with open(merged_file, "wb") as f:
        pdf_writer.write(f)

def merge2():
    pdf_file_merger = PdfFileMerger()

    with open('./tmp/pdf1.pdf','rb') as f1:

	    pdf_file_merger.append(f1)

    with open('./tmp/pdf2.pdf','rb') as f2:

        pdf_file_merger.append(f2)

    with open('./tmp/merge2.pdf','wb') as f3:

        pdf_file_merger.write(f3)

    pdf_file_merger.close()

def merge_pdf(work_path, output_pdf_filepath):
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

    # 結合先PDFを保存する
    filename = output_pdf_filepath
    doc.save(filename)



