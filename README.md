# Scraping-ATCM
Title: Antarctic Treaty System Document Processing and Corpus Building Scripts

**Description**:
These Python scripts are part of a suite designed to automate the process of downloading, converting, and cleaning official documents from the Antarctic Treaty System (ATS) for corpus building. The scripts handle tasks ranging from file format conversion, downloading articles, generating standardized names, and cleaning text data for analysis.

**Functionality**:


Converts XLSX files to CSV for easier handling of document metadata.
Downloads article files (PDFs, DOCXs) from the ATS website based on the information in the CSV files.
Generates standardized file names for downloaded documents based on their metadata.

Extracts text from DOCX files.
Cleans and preprocesses the text data by removing punctuation, non-standard characters, and excess whitespace.
Constructs a DataFrame mapping file names to their corresponding cleaned text.

**Corpus Building**:

Merges metadata from CSV files with corresponding text data.
Consolidates multiple DataFrames into a single comprehensive DataFrame.
Filters and formats the final DataFrame to include essential information, ready for analysis.
Usage:


**Dataset descriptions**:
This dataset contains a collection of information and working papers from the Antarctic Treaty Consultative Meeting (ATCM) held since 1961. Each record in the dataset offers a comprehensive view of individual documents presented during the meeting, including:

* **ID**: A unique identifier for each document, typically including the ATCM session number, document number, and file extension (e.g., ATCM45_ip001_e.docx).
* **DocumentID**: A numerical identifier assigned to each document for easy reference and categorization.
* **Type**: The classification of the document, such as 'ip' (Information Paper) or 'wp' (Working Paper), indicating the nature of the content.
* **Year**: The year of the ATCM session, signifying the temporal context of the document (e.g., 2023).
* **Title**: The official title of the document, providing a concise summary of its subject or focus.
* **Submitted By**: The country or organization that submitted the document, reflecting the diverse international participation in the ATCM.
* **Category**: The thematic category assigned to the document, such as 'Cooperation with other organizations' or 'Operation of the Antarctic Treaty System', helps in understanding the focus areas of the meeting.

**Acknowledgement:**

Research funded by Australian Research Council SRIEAS Grant SR200100005 Securing Antarctica’s Environmental Future.

**Note**:

Run the scripts in order: first for downloading and format conversion, then for text extraction and cleaning.
Modify file paths and URLs as per your directory structure and document URLs.


Please adhere to ethical guidelines and legal restrictions when downloading and using documents from online sources.
The scripts are structured for clarity and modularity, allowing easy modification and extension for specific research needs.
