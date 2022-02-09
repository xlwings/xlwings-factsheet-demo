import os
import sys
import tempfile
import datetime as dt
from pathlib import Path

import docx
import segno
import pandas as pd
import xlwings as xw
from xlwings.pro.reports import Markdown, MarkdownStyle, Image


# Global variables/constants
THIS_DIR = Path(__file__).resolve().parent
XLWINGS_GREEN = '#15a43a'
S3_BUCKET = 'xlwings'
S3_FOLDER = 'funds'


def main():
    """Main entry point and called from the 'Run' button in Excel"""

    try:
        run_sheet = xw.Book.caller().sheets['Run']
        # Create 'reports' directories if they don't exist yet
        for subdir in ['pdf', 'xlsx']:
            os.makedirs(THIS_DIR / 'reports' / subdir, exist_ok=True)
        # Read in the settings as dictionary
        settings = run_sheet['settings'].expand().options(dict).value
        # Translate 'ALL' into '*' so that glob below understands it
        fund_selection = '*' if settings['Fund Selection'] == 'ALL' else settings['Fund Selection']
        with xw.App(visible=False) as app:
            for directory in (THIS_DIR / 'data' / 'funds').glob(fund_selection):
                # glob with an exact match changes the capitalization on Windows,
                # so don't use the directory.name in the single fund case
                fundname = directory.name if fund_selection == '*' else fund_selection
                # The three steps: 1) pre-processing 2) report generation 3) post-processing
                data = preprocess(fundname, run_sheet)
                pdf_path = create_xlsx_and_pdf_reports(data, fundname, settings, run_sheet, app)
                postprocess(pdf_path, fundname, settings, run_sheet)
    finally:
        # Exceptions are already handled by xlwings and are shown as Excel dialogues
        run_sheet['status'].value = None


def preprocess(fundname, run_sheet):
    """Data acquisistion and manipulation"""

    run_sheet['status'].value = f'Preparing Data: {fundname}'

    # Read CSV files with pandas
    csv_dir = THIS_DIR / 'data' / 'funds' / fundname
    holdings = pd.read_csv(csv_dir / 'holdings.csv')
    history = pd.read_csv(csv_dir / 'history.csv', index_col='Date', parse_dates=True)

    # Read a Word document with python-docx
    doc = docx.Document(THIS_DIR / 'data' / 'common' / 'intro.docx')
    intro = ''
    for paragraph in doc.paragraphs:
        intro += paragraph.text + '\n'

    # Read a text file
    with open(THIS_DIR / 'data' / 'common' / 'disclaimer.md', 'r', encoding='utf-8') as fh:
        disclaimer = fh.read()

    # Calculate historical performance
    history = history.sort_index()
    total_returns = history.iloc[-1, :] / history.iloc[0, :] - 1
    fund_return = round(total_returns['Fund'], 4)
    # Insert Jan 1 so the chart won't start in the middle of the year
    history.loc[dt.datetime(history.index[0].year, 1, 1), :] = None
    history = history.sort_index()

    # Calculate sector weights
    sectors = holdings.groupby('Industry').sum()

    # Produce QR code
    qr = segno.make(f'https://www.xlwings.org/funds/{fundname.replace(" ", "-")}')
    if sys.platform.startswith("darwin"):
        extension = "pdf"
    else:
        extension = "svg"
    qrcode_path = tempfile.gettempdir() + f'/qr.{extension}'
    qr.save(qrcode_path, scale=5, border=0, finder_dark=XLWINGS_GREEN)

    # Define Markdown styling
    style = MarkdownStyle()
    style.h1.font.size = 11
    style.h1.font.color = XLWINGS_GREEN

    # Note that the index of pandas DataFrames are not passed over to Excel, so
    # make sure to call reset_index() if you need the index data in Excel.
    return dict(
        fundname=fundname,
        intro=Markdown(intro, style),
        disclaimer=Markdown(disclaimer, style),
        fund_return=fund_return,
        asofdate=dt.datetime.now(),  # local time
        holdings=holdings,
        sectors=sectors.reset_index(),
        qrcode=Image(qrcode_path),
        history=history.reset_index(),
    )


def create_xlsx_and_pdf_reports(data, fundname, settings, run_sheet, app):
    """Creates the Excel and PDF reports."""

    # Excel report
    run_sheet['status'].value = f'Creating Excel Report: {fundname}'
    report_book = app.create_report(
        THIS_DIR / 'template' / 'template.xlsx',
        THIS_DIR / 'reports' / 'xlsx' / f"{fundname}.xlsx",
        **data,
    )

    # PDF report
    run_sheet['status'].value = f'Creating PDF Report: {fundname}'
    pdf_path = THIS_DIR / 'reports' / 'pdf' / f"{fundname}.pdf"
    report_book.to_pdf(
        path=pdf_path,
        layout=THIS_DIR / 'template' / 'layout.pdf',
        show=True if settings['Open PDFs'] else False,
    )

    return pdf_path


def postprocess(file_path, fundname, settings, run_sheet):
    """Upload a file to S3"""

    import boto3  # importing here since this step is optional

    if settings['Upload PDFs']:
        run_sheet['status'].value = f'Uploading to S3: {fundname}'
        s3_client = boto3.client('s3')
        s3_client.upload_file(str(file_path), S3_BUCKET, f'{S3_FOLDER}/{file_path.name}')


if __name__ == '__main__':
    # This part is only used when you run the script from Python instead of Excel
    xw.Book('demo.xlsm').set_mock_caller()
    main()
