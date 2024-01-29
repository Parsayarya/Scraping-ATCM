import os
import re
import pandas as pd
import docx
import string


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


def csv2name(file: str) -> None:
    """Convert data from a CSV file into a list of download links and append to the CSV file.

    :param file: The path to the CSV file.
    """
    df = pd.read_csv(file)

    links = []
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

        links.append(f'{s_final_1}_{s_final_2}_e.docx')

    df['Generated_Name'] = links
    df.to_csv(file, index=False, encoding='utf-8')


def xlsx_to_csv(xlsx_file: str, csv_file: str) -> bool:
    """
    Convert an XLSX file to a CSV file.

    :param xlsx_file: The path to the input XLSX file.
    :param csv_file: The path to the output CSV file.
    :return: True if the conversion was successful, otherwise False.
    """
    try:
        df = pd.read_excel(xlsx_file)
        df.to_csv(csv_file, index=False, encoding='utf-8')
        print(f"Conversion successful: {xlsx_file} -> {csv_file}")
        return True
    except Exception as e:
        print(f"Conversion failed: {e}")
        return False


# xlsx_to_csv(r'D:\python\ATS\DownloadCorpusBuilder\files\csvInputs\XLSX for Corpus\listofpapersAIP.xlsx',
#             r'D:\python\ATS\DownloadCorpusBuilder\files\csvInputs\XLSX for Corpus\listofpapersAIP.csv')
# xlsx_to_csv(r'D:\python\ATS\DownloadCorpusBuilder\files\csvInputs\XLSX for Corpus\listofpapersAWP.xlsx',
#             r'D:\python\ATS\DownloadCorpusBuilder\files\csvInputs\XLSX for Corpus\listofpapersAWP.csv')
# xlsx_to_csv(r'D:\python\ATS\DownloadCorpusBuilder\files\csvInputs\XLSX for Corpus\listofpapersSIP.xlsx',
#             r'D:\python\ATS\DownloadCorpusBuilder\files\csvInputs\XLSX for Corpus\listofpapersSIP.csv')
# xlsx_to_csv(r'D:\python\ATS\DownloadCorpusBuilder\files\csvInputs\XLSX for Corpus\listofpapersSWP.xlsx',
#             r'D:\python\ATS\DownloadCorpusBuilder\files\csvInputs\XLSX for Corpus\listofpapersSWP.csv')
# csv2name(r'D:\python\ATS\DownloadCorpusBuilder\files\csvInputs\XLSX for Corpus\listofpapersAIP.csv')
# csv2name(r'D:\python\ATS\DownloadCorpusBuilder\files\csvInputs\XLSX for Corpus\listofpapersAWP.csv')
# csv2name(r'D:\python\ATS\DownloadCorpusBuilder\files\csvInputs\XLSX for Corpus\listofpapersSIP.csv')
# csv2name(r'D:\python\ATS\DownloadCorpusBuilder\files\csvInputs\XLSX for Corpus\listofpapersSWP.csv')


def get_text(file_path: str) -> str:
    """
    Extract text from a DOCX file.

    :param file_path: The path to the DOCX file.
    :return: The extracted text as a single string.
    """
    doc = docx.Document(file_path)
    full_text = [para.text for para in doc.paragraphs]
    return '\n'.join(full_text)


def clean_text(text: str) -> str:
    """
    Clean and preprocess text.

    :param text: The input text to be cleaned.
    :return: The cleaned text.
    """
    text = text.lower()
    text = re.sub(r'\[.*?\]', ' ', text)
    text = re.sub(r'[%s]' % re.escape(string.punctuation), ' ', text)
    text = re.sub(r'\w*\d\w*', ' ', text)
    text = text.replace('�', ' ')
    return text


def build_text_df(text_folder: str) -> pd.DataFrame:
    """
    Build a DataFrame of filenames and cleaned texts from a folder of DOCX files.

    :param text_folder: The path to the folder containing DOCX files.
    :return: A DataFrame with two columns: file name and text data.
    """
    data_list = []
    for filename in os.listdir(text_folder):
        file_path = os.path.join(text_folder, filename)
        if os.path.splitext(filename)[1].lower() == '.docx':
            try:
                cleaned_text = re.sub(' {2,}', ' ', clean_text(get_text(file_path).replace('\n', '').strip()))
                data_list.append({'Generated_Name': filename, 'text': cleaned_text})
            except Exception as e:
                print(f"Error processing {filename}: {e}")
                data_list.append({'Generated_Name': filename, 'text': pd.NA})

    return pd.DataFrame(data_list)


# df = build_text_df(r'D:\python\ATS\DownloadCorpusBuilder\files\docxOutputs\SWPN')
# main_df = pd.read_csv(r'D:\python\ATS\DownloadCorpusBuilder\files\csvInputs\XLSX for Corpus\listofpapersSWP.csv')
# merged_df = main_df.merge(df, on=['Generated_Name'], how='left')
# merged_df.to_csv(r'D:\python\ATS\DownloadCorpusBuilder\files\Data frames\SWPFinalCorpus.csv', index=False)

import pandas as pd

# List of file paths
file_paths = [
    r'D:\python\ATS\DownloadCorpusBuilder\files\Data frames\IPFinalCorpus.csv',
    r'D:\python\ATS\DownloadCorpusBuilder\files\Data frames\SIPFinalCorpus.csv',
    r'D:\python\ATS\DownloadCorpusBuilder\files\Data frames\SWPFinalCorpus.csv',
    r'D:\python\ATS\DownloadCorpusBuilder\files\Data frames\WPFinalCorpus.csv',
    r'D:\python\ATS\DownloadCorpusBuilder\files\Data frames\ScopusCorpus.csv'
]

# Load all the CSV files into DataFrames
dataframes = [pd.read_csv(file_path, low_memory=False) for file_path in file_paths]

# Combine all DataFrames into a single DataFrame
combined_df = pd.concat(dataframes, ignore_index=True)

# Drop rows where the "text" column is empty
combined_df.dropna(subset=['text'], inplace=True)

# Keep only the "Generated_Name" and "text" columns
final_df = combined_df[['Generated_Name', 'text']]

# Save the final DataFrame to a CSV file
final_csv_path = r'D:\python\ATS\DownloadCorpusBuilder\files\Data frames\FinalCorpus2.csv'
final_df.to_csv(final_csv_path, index=False, encoding='utf-8')
