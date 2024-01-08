Python Notebooks:

1. These are the Python notebooks that were used during our scoping review. Both "findfiles" and "datatransfer" 
are helper notebooks. Whilst they will not run directly, they must be kept in the same folder as the other 
notebooks to ensure correct operation.

- The "findfiles" class as defualt is commented out. It opens a filedialog box to help users idenfity and 
copy the filepaths to Excel workbooks.

- The "datatransfer" class is used to help import and export data from Excel to Pandas and vice-versa.

2. The Python notebooks import scoping review data. As default, the scoping review data (Excel workbooks) should be kept 
on the desktop in a folder called "ScopingData". This may be changed but the code in the Python notebooks 
must be edited accordingly.

3. The required dependencies are listed in the "fiddes_enviornment.ymal".


To run the Python notebooks:

1. Open a blank Excel workbook and run "1_make_search_terms_worksheet.ipynb" in a Python notebook.
The number of search terms columns can be changed to between 2 and 5.

2. Open the Excel workbook "3a. Search Terms Worksheet (28-Feb-2023).xlsx" and run "2_make_search_strings.ipynb" in a Python notebook.

3. Open the Excel workbook "4a. Endnote Export Raw.xlsx" and run "3_endnote_csv_formatting.ipynb" in a Python notebook.

4. Open a blank Excel workbook and run "4_results_deduplication.ipynb" in a Python notebook. Requires access 
to Excel workbooks ("5. Cinhal.xlsx", "5. Embase.xlsx", "5. Medline.xlsx", "5. Scopus.xlsx", "5. Web of Science.xlsx").

5. Open a blank Excel workbook and run "5_results_filtering.ipynb" in a Python notebook.
Requires access to Excel workbooks ("6. Deduplication and Filtering (06-Mar-2023).xlsx").

6. Open a blank Excel workbook and run "6_data_analysis.ipynb" in a Python notebook.
Requires access to Excel workbooks ("7. Data Extraction (29-May-2023).xlsx").
