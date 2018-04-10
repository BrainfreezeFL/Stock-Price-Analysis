# Stock-Price-Analysis
Old convoluted code initially meant to analyze over 50 years of stock data, later used to find the expected return of as stock. I will slowly be updating this code to be more efficient, readable and usable.

# To Run the Program

Add your python27 directory to the system path
</br>
Run in python27 directory:
>get-pip.py
</br>
>pip install xlrd
</br>
</br>
This program uses XLRD to read excel files. To obtain the data you wish to manipulate:
  - Go to Yahoo finance</br>
  - Search for your stock</br>
  - Select Historical prices</br>
  - Select the max period for dates</br>
  - Download the data into the same file as this script</br>
</br>
Once the data is obtained, change the name in the script to open the name of the file you have
<br>
</br>
Run:
python financialexcelsheetdataparser.py
</br>
</br>
Side Note: It is worth noting that I use version 2 of python, not version 3
