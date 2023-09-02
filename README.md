# Web Scraping using Python and exporting all Microsoft Excel formulas in an Excel file

This repository contains Python code in the form of a Jupyter Notebook that demonstrates how to scrape data from a web page and export it to an Excel file. It utilizes libraries such as Pandas, BeautifulSoup, and XlsxWriter to perform these tasks.

## Table of Content:
- [Overview]
- [Prerequisites]
- [Usage]
- [License]

---

## Overview:

This Jupyter Notebook provides a step-by-step guide on how to:

- Send an HTTP GET request to a web page.
- Parse the HTML content using BeautifulSoup.
- Extract tabular data from the web page.
- Process and clean the data, including splitting a combined "Type and Description" column.
- Export the data to an Excel file with formatting.

## Prerequisites:

Before using this code, ensure you have the following:

- Python installed on your system (Python 3 recommended).
- Jupyter Notebook installed (you can install it using `pip`).
- Required Python libraries installed (install them using `pip`):
- For Example: pip install pandas
  - pandas
  - requests
  - xlsxwriter
  - bs4 (Beautiful Soup)

## Usage:

1. Clone this repository to your local machine or download the Jupyter Notebook file (`web_scraping_and_excel_export.ipynb`).

2. Open the Jupyter Notebook in your Jupyter environment.

3. Customize the following variables in the code according to your needs:

   - `URL`: The URL of the web page you want to scrape.
   - `output_directory`: The directory where you want to save the Excel file.

4. Execute the code cell by cell in your Jupyter Notebook environment.

5. The scraped data will be saved in an Excel file named "your_file_name.xlsx" in the specified output directory.

6. Review and modify the code as needed for your specific web scraping tasks.

---

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
