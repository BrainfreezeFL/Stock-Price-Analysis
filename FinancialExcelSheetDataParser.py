# Samuel Lewis
# 11/5/2016
# Parse an excel Sheet
import xlrd
def menuChoice2():
    print("Enter the beta: ")
    beta = input()
    print("Enter the risk free return: ")
    riskFree = input()
    print("Enter expected return of market: ")
    marketExpectedReturn = input()
    print("Enter the last divdend price: ")
    dividend = input()
    print("Enter the price when divdend was issued: ")
    priceWhenDivdendIssued = input()
    growth = (riskFree+(beta*(marketExpectedReturn-riskFree)))-(dividend/priceWhenDivdendIssued)
    print("The growth is: %f") %growth
    return
#print ("1")
print("loading...")
#open up the excel sheet
workbook = xlrd.open_workbook("S&P500dataTo1950.xlsx")
#print ("Please enter your input:")

#open up the sheet on the excel file
worksheet = workbook.sheet_by_index(0)
#print ("3")
x=0
while (x is not 1):
    
    print("1. Market Expected return")
    print("2. Expected Growth")
    print("3. Exit")
    #get the users input about what the stock price is
    menuChoice = input()
    if(menuChoice is 1):
        #print ("4")
        print("Please enter how high above the moving averages(%) you would like to test:")
        stockPrice = input()
        #declare a dynamic array(for each moving average)
        #to catch each instance of when it has been at that
        #level above the moving average
        twentyDay = [0 for i in range(16000)]
        fiftyDay = [0 for i in range(16000)]
        twoHundredDay = [0 for i in range(16000)]
        #print ("5")

        #Make some counters to see how many spaces are in the arrays
        twentyCounter = 0;
        fiftyCounter = 0;
        twoHundredCounter = 0;
        #print ("6")
        #parse through the file now(20 day) and catch instances
        #within .5 percent
        #make the for loop

        row = 0
        print ("loading...")
        for row in range(300,16000) :
            #print ("8")
            #the rows are incrementing but the col is 15
            tempVal = worksheet.cell(row, 16).value
            tempVal4 = worksheet.cell(row, 15).value
            tempVal5 = worksheet.cell(row,14).value
            #print("%d") %tempVal
            #print("9")
            #check if the number in the column is within .5 percent
            if stockPrice <5:
                if stockPrice-.5<tempVal and stockPrice+.5>tempVal:
                
                    twoHundredDay[twoHundredCounter] = row
                    twoHundredCounter = twoHundredCounter+1
                elif stockPrice-.5<tempVal4 and stockPrice+.5>tempVal4:
                
                    fiftyDay[fiftyCounter] = row
                    fiftyCounter = fiftyCounter+1
                elif stockPrice-.5<tempVal5 and stockPrice+.5>tempVal5:
                
                    twentyDay[twentyCounter] = row
                    twentyCounter = twentyCounter+1
                    
            else:
                if stockPrice-2<tempVal and stockPrice+2>tempVal:
                
                    twoHundredDay[twoHundredCounter] = row
                    twoHundredCounter = twoHundredCounter+1
                elif stockPrice-2<tempVal4 and stockPrice+2>tempVal4:
                
                    fiftyDay[fiftyCounter] = row
                    fiftyCounter = fiftyCounter+1
                elif stockPrice-2<tempVal5 and stockPrice+2>tempVal5:
                
                    twentyDay[twentyCounter] = row
                    twentyCounter = twentyCounter+1
            
        i=0
        sumOfTwentyCounter = 0.0
        sumOfTwentyCounter1 = 0.0
        sumOfTwentyCounter2 = 0.0
        sumOfFiftyCounter = 0.0
        sumOfFiftyCounter1 = 0.0
        sumOfFiftyCounter2 = 0.0
        sumOfTwoHundredCounter = 0.0
        sumOfTwoHundredCounter1 = 0.0
        sumOfTwoHundredCounter2 = 0.0
        average = 0.0
        average1 = 0.0
        average2 = 0.0
        for i in range(twentyCounter):
            #print("They are close on the row: %d\n") %twentyDay[i]
            tempthree = twentyDay[i]
            tempTwo = worksheet.cell(tempthree, 17).value
            sumOfTwentyCounter = tempTwo + sumOfTwentyCounter
            tempTen = worksheet.cell(tempthree,19).value
            sumOfTwentyCounter1 = tempTen + sumOfTwentyCounter1
            tempten = worksheet.cell(tempthree,20).value
            sumOfTwentyCounter2 = tempTen + sumOfTwentyCounter2
          
            
        if(twentyCounter is not 0):
            average = sumOfTwentyCounter/twentyCounter
            average1 = sumOfTwentyCounter1/twentyCounter
            average2 = sumOfTwentyCounter2/twentyCounter
            
            print("20 Day ROI for 1 year: %f") %average
            print("20 Day ROI for 1 month: %f") %average1
            print("20 Day ROI for 1 week: %f") %average2
        else:
            print("Sorry, no matches in 20 Day MA.")
        for i in range(fiftyCounter):
            #print("They are close on the row: %d\n") %fiftyDay[i]
            tempthree = fiftyDay[i]
            tempTwo = worksheet.cell(tempthree, 17).value
            #print(tempTwo)
            sumOfFiftyCounter = tempTwo + sumOfFiftyCounter
            tempTen = worksheet.cell(tempthree,19).value
            sumOfTFiftyCounter1 = tempTen + sumOfFiftyCounter1
            tempEleven = worksheet.cell(tempthree,20).value
            sumOfFiftyCounter2 = tempEleven + sumOfFiftyCounter2
            
        if(fiftyCounter is not 0):
            average = sumOfFiftyCounter/fiftyCounter
            average1 = sumOfFiftyCounter1/fiftyCounter
            average2 = sumOfFiftyCounter2/fiftyCounter
            
            print("50 Day ROI for 1 year: %f") %average
            print("50 Day ROI for 1 month: %f") %average1
            print("50 Day ROI for 1 week: %f") %average2
        else:
            print("Sorry, not matches in 50 Day MA.")
        for i in range(twoHundredCounter):
            #print("They are close on the row: %d\n") %twentyDay[i]
            tempthree = twoHundredDay[i]
            tempTwo = worksheet.cell(tempthree, 17).value
            sumOfTwoHundredCounter = (float)(tempTwo + sumOfTwoHundredCounter)
            tempTen = worksheet.cell(tempthree,19).value
            sumOfTwoHundredCounter1 =(float) (tempTen + sumOfTwoHundredCounter1)
            tempEleven = worksheet.cell(tempthree,20).value
            sumOfTwoHundredCounter2 = (float)(tempEleven + sumOfTwoHundredCounter2)
            
        if(twoHundredCounter is not 0):
            average = sumOfTwentyCounter/twoHundredCounter
            average1 = sumOfTwentyCounter1/twoHundredCounter
            average2 = sumOfTwentyCounter2/twoHundredCounter
            
            
            print("200 Day ROI for 1 year: %f") %average
            print("200 Day ROI for 1 month: %f") %average1
            print("200 Day ROI for 1 week: %f") %average2
        else:
            print("Sorry, no matches in 200 Day MA.")
            
        #print("The average rate of return in this market state is: %d") %average

        #print("The value of 16,2 is: %d") %worksheet.cell(2,16).value
            
        ##print("Would you like to run a scan again?")
        #print("-----------------   NEW SCAN   -----------------")
        #print("Y=0    N= 1")
        #x=input()
    if(menuChoice is 3):
        x=1
    elif(menuChoice is 2):
        menuChoice2()
    
    print("------------------------------------------------")
        
def menuChoice2():
    print("Enter the beta: ")
    beta = input()
    print("Enter the risk free return: ")
    riskFree = input()
    print("Enter expected return of market: ")
    marketExpectedReturn = input()
    print("Enter the last divdend price: ")
    dividend = input()
    print("Enter the price when divdend was issued: ")
    priceWhenDivdendIssued = input()
    growth = (riskFree+(beta*(marketExpectedReturn-riskFree)))-(dividend/priceWhenDivdendIssued)
    print("The growth is: %f") %growth
    return
    
