# Purchase Tracking and Spreadsheet Creation

This Python script tracks purchases of selected items and creates a spreadsheet with the collected data. The script allows users to select a file containing purchase information and performs data processing and manipulation to generate summary reports in an Excel file.

## Features

- Select a CSV file containing purchase data using a file dialog.
- Process the selected file to extract relevant columns and perform data cleaning operations.
- Generate pivot tables summarizing total quantity and cost per unit for vendors and restaurants.
- Save the processed data to an Excel file with separate sheets for vendors, restaurants, and detailed information.

## Prerequisites

To run the script, ensure you have the following dependencies installed:

- Python 3
- pandas
- numpy
- tkinter

## Usage

1. Clone or download the script to your local machine.
2. Open a terminal or command prompt and navigate to the directory where the script is located.
3. Install the necessary dependencies by running the following command:

pip install pandas numpy

4. Run the script using the following command:

python script.py

5. The script will prompt you to select a CSV file containing purchase data. Choose the appropriate file.
6. The script will process the selected file and generate an Excel file with the processed data.
7. The Excel file will include separate sheets for vendors, restaurants, and detailed information.

## License

This project is licensed under the [MIT License](LICENSE).

Feel free to use and modify the code according to your needs.
