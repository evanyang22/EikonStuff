## Important Programs

### SearchScrapper.ipynb

Purpose: To create an output csv file with important information from the folders where the PDF files are stored. The output csv file contains data from the Advanced Research Search excel file that is automatically generated along with information on whether or not the file is downloaded. If it is downloaded, it also contains the path to the PDF. This data can be analyzed to determine download trends and download volumes, and I've attached some example programs in this module that I've used to create graphs concerning various metrics. I've included an example of the output file here (Aggregation2022.zip).

To use: Drag the ipynb file into the folder (Ex: BetaTest) containing the pdf reports organized by company first, then year second. Run the ipynb file using Jupyter Notebook and a csv file will be generated within the folder after it. You can also convert this program to a Python file and run it using the terminal ("python SearchScrapper.py" after navigating to the directory "cd xxx")

### 11_8_2022_Lists_Comparison.ipynb

Purpose: This programs reads in csv data from an old Thomson Dataset (1984-2011) and the newly downloaded report data (2012-2022) and creates graphs such as number of total downloaded reports per year and reports from shared contributors per year.
