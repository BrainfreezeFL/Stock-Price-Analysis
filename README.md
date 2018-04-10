# Stock-Price-Analysis
Old convoluted code initially meant to analyze over 50 years of stock data, later used to find the expected return of as stock. I will slowly be updating this code to be more efficient, readable and usable.

# To Run the Program
I use version 2 of python, not version 3
Add your python27 directory to the system path

Run in python27 directory:
>get-pip.py
</br>
>pip install xlrd
</br>
</br>
This program uses XLRD to read excel files. To obtain the data you wish to manipulate:
  - Go to Yahoo finance
  - Search for your stock
  - Select Historical prices
  - Select the max period for dates
  - Download the data into the same file as this script
</br>
Once the data is obtained, change the name in the script to open the name of the file you have
</br>
Run:
python financialexcelsheetdataparser.py
</br>
Side Note: It is worth noting that I use version 2 of python, not version 3
