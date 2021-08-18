import os
import sys
import tempfile
import datetime as dt
from pathlib import Path

import docx
import segno
import pandas as pd
import PySimpleGUI as sg
import xlwings as xw
from xlwings.pro.reports import Markdown, MarkdownStyle, Image


# Global variables/constants
if not hasattr(sys, 'frozen'):
    THIS_DIR = Path(__file__).resolve().parent
else:
    THIS_DIR = Path(sys.executable).parent
XLWINGS_GREEN = '#15a43a'
S3_BUCKET = 'xlwings'
S3_FOLDER = 'funds'


def main():
    """Main entry point and called from the 'Run' button in Excel"""

    # Create 'reports' directories if they don't exist yet
    for subdir in ['pdf', 'xlsx']:
        os.makedirs(THIS_DIR / 'reports' / subdir, exist_ok=True)

    # GUI
    layout = [
        [sg.Text('xlwings Reports: Factsheet Demo', font=('Arial 14'))],
        [sg.Text('Fund Selection', key='-FUND_SELECTION-', size=(12, 1)), sg.DropDown(('ALL', 'Fund A', 'Fund B', 'Fund C'), 'Fund A', size=(12, 1))],
        [sg.Text('Open PDFs', key='-OPEN_PDFS-', size=(12, 1)), sg.DropDown((True, False), 'False', size=(12, 1))],
        [sg.Text('Upload PDFs', key='-UPLOAD_PDFS-', size=(12, 1)), sg.DropDown((True, False), 'False', size=(12, 1))],
        [sg.Text()],
        [sg.Text('Status:')],
        [sg.Output(key='-OUTPUT-')],
        [sg.Button('Run'), sg.Button('Cancel')]
    ]

    window = sg.Window('Factsheet Demo', layout, icon='xlwings.ico')

    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Cancel':
            break
        if event == 'Run':
            window.set_cursor('wait')
            # Read in the settings as dictionary
            value_map = {0: 'Fund Selection',  1: 'Open PDFs', 2: 'Upload PDFs'}
            settings = dict((value_map[key], value) for (key, value) in values.items())

            print("Starting...")

            # Translate 'ALL' into '*' so that glob below understands it
            fund_selection = '*' if settings['Fund Selection'] == 'ALL' else settings['Fund Selection']
            with xw.App(visible=False) as app:
                for directory in (THIS_DIR / 'data' / 'funds').glob(fund_selection):
                    # glob with an exact match changes the capitalization on Windows,
                    # so don't use the directory.name in the single fund case
                    fundname = directory.name if fund_selection == '*' else fund_selection
                    print(f'Processing {fundname}')
                    # The three steps: 1) pre-processing 2) report generation 3) post-processing
                    data = preprocess(fundname)
                    pdf_path = create_xlsx_and_pdf_reports(data, fundname, settings, app)
                    postprocess(pdf_path, fundname, settings)

        window['-OUTPUT-'].update('')
        window.set_cursor('arrow')
    window.close()


def preprocess(fundname):
    """Data acquisistion and manipulation"""

    print(f'Preparing Data: {fundname}')

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
    qrcode_path = tempfile.gettempdir() + '/qr.svg'
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


def create_xlsx_and_pdf_reports(data, fundname, settings, app):
    """Creates the Excel and PDF reports."""

    # Excel report
    print(f'Creating Excel Report: {fundname}')
    report_book = app.create_report(
        THIS_DIR / 'template' / 'template.xlsx',
        THIS_DIR / 'reports' / 'xlsx' / f"{fundname}.xlsx",
        **data,
    )

    # PDF report
    print(f'Creating PDF Report: {fundname}')
    pdf_path = THIS_DIR / 'reports' / 'pdf' / f"{fundname}.pdf"
    report_book.to_pdf(
        path=pdf_path,
        layout=THIS_DIR / 'template' / 'layout.pdf',
        show=True if settings['Open PDFs'] else False,
    )

    return pdf_path


def postprocess(file_path, fundname, settings):
    """Upload a file to S3"""

    import boto3  # importing here since this step is optional

    if settings['Upload PDFs']:
        print(f'Uploading to S3: {fundname}')
        s3_client = boto3.client('s3')
        s3_client.upload_file(str(file_path), S3_BUCKET, f'{S3_FOLDER}/{file_path.name}')

if __name__ == '__main__':
    main()
