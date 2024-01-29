import os
import requests
import pandas as pd

# Constants for file paths
dirname = os.path.dirname(__file__)
CSV_INPUT_PATH = os.path.join(dirname, 'files/csvInputs/')
DOCX_OUTPUT_PATH = os.path.join(dirname, 'files/docxOutputs/')

# https://documents.ats.aq/SATCM12/wp/SATCM12_wp001_e.pdf
# https://documents.ats.aq/SATCM1/wp/SATCM1_wp001_e.pdf
# https://documents.ats.aq/ATCM45/wp/ATCM45_wp001_e.docx
# https://documents.ats.aq/ATCM31/wp/ATCM31_wp001_e.doc

def xlsx_to_csv(xlsx_file, csv_file):
    """Convert an XLSX file to a CSV file.

    :param xlsx_file: The path to the input XLSX file.
    :param csv_file: The path to the output CSV file.
    :return: True if the conversion was successful, otherwise False.
    """
    try:
        # Read the XLSX file into a DataFrame
        df = pd.read_excel(xlsx_file)

        # Save the DataFrame as a CSV file
        df.to_csv(csv_file, index=False, encoding='utf-8')

        print(f"Conversion successful: {xlsx_file} -> {csv_file}")
        return True
    except Exception as e:
        print(f"Conversion failed: {e}")
        return False


def download_article_file(url: str, sub_folder: str) -> bool:
    """Download an article from a URL to a local directory.

    :param url: The URL of the article file to be downloaded.
    :param sub_folder: The sub-folder where the article will be saved.
    :return: True if the article file was successfully downloaded, otherwise False.
    """
    response = requests.get(url, stream=True)
    docx_file_name = os.path.basename(url)

    if response.status_code == 200:
        filepath = os.path.join(os.getcwd(), DOCX_OUTPUT_PATH, sub_folder, docx_file_name)
        os.makedirs(os.path.dirname(filepath), exist_ok=True)

        with open(filepath, 'wb') as docx_object:
            docx_object.write(response.content)
            print(f'{docx_file_name} was successfully saved!')
            return True
    else:
        print(f'Could not download {docx_file_name},')
        print(f'HTTP response status code: {response.status_code}')
        return False


def from_roman_numeral(numeral: str) -> int:
    """Convert a Roman numeral to an integer.

    :param numeral: The Roman numeral string to convert.
    :return: The integer equivalent of the Roman numeral.
    """
    value_map = {"I": 1, "V": 5, "X": 10, "L": 50, "C": 100, "D": 500, "M": 1000}
    value = 0
    last_digit_value = 0

    for roman_digit in reversed(numeral):
        digit_value = value_map[roman_digit]

        if digit_value >= last_digit_value:
            value += digit_value
            last_digit_value = digit_value
        else:
            value -= digit_value

    return value


def csv2links(file: str) -> list:
    """Convert data from a CSV file into a list of download links.

    :param file: The path to the CSV file.
    :return: A list of download links based on the CSV data.
    """
    links = []
    df = pd.read_csv(file)

    for _, row in df.iterrows():
        meeting = row['Meeting'].split('-')[0].strip()
        roman = str(from_roman_numeral(meeting.split(' ')[1]))
        s_final_1 = meeting.split(' ')[0] + roman

        no = row['No.'].split('-')[0].strip()
        article_type = no.split(' ')[0][0:2].lower()

        if len(no) > 5:
            s_final_2 = no.split(' ')[0].lower() + '_rev' + no.split(' ')[2]
        else:
            s_final_2 = no.split(' ')[0].lower()

        # links.append(f'https://documents.ats.aq/{s_final_1}/{article_type}/{s_final_1}_{s_final_2}_e.docx')
        # links.append(f'https://documents.ats.aq/{s_final_1}/{article_type}/{s_final_1}_{s_final_2}_e.doc')
        links.append(f'https://documents.ats.aq/{s_final_1}/{article_type}/{s_final_1}_{s_final_2}_e.pdf')

    return links


xlsx_to_csv(os.path.join(CSV_INPUT_PATH, 'listofpapersSIP.xlsx'), os.path.join(CSV_INPUT_PATH, 'listofpapersSIP.csv'))
xlsx_to_csv(os.path.join(CSV_INPUT_PATH, 'listofpapersSWP.xlsx'), os.path.join(CSV_INPUT_PATH, 'listofpapersSWP.csv'))
# Convert CSV data into lists of download links
IPArticles = csv2links(os.path.join(CSV_INPUT_PATH, 'listofpapersSIP.csv'))
WPArticles = csv2links(os.path.join(CSV_INPUT_PATH, 'listofpapersSWP.csv'))

# Download articles and save them to specified sub-folders
for A in IPArticles:
    download_article_file(A, 'SIPN/')
for A in WPArticles:
    download_article_file(A, 'SWPN/')
