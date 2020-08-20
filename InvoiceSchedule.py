from datetime import datetime
from datetime import timedelta

Order_Term_Start_Date = datetime(Quote.EffectiveDate).date()
Order_Term_End_Date_String = Quote.GetCustomField('Order Term End Date').Content
Order_Term_End_Date = datetime.strptime(Order_Term_End_Date_String, '%d/%m/%y').date()

No_Of_Years = (Order_Term_End_Date.year - Order_Term_Start_Date.year)
No_Of_Months = Order_Term_End_Date.month - Order_Term_Start_Date.month

if No_Of_Months < 0:
    globals()['No_Of_Years'] -= 1
    globals()['No_Of_Months'] = 12 - abs(No_Of_Months)

No_Of_Days = abs(Order_Term_End_Date.day - Order_Term_Start_Date.day)
Total_No_Of_Months = No_Of_Years * 12 + No_Of_Months

Invoice_Date = Order_Term_Start_Date

Quote_Items = Quote.Items

Total_Amounts = {"Subscriptions": 0, "Events": 0, "Education Services": 0}

for Item in Quote_Items:
    Total_Amounts[Item.CategoryName] = Total_Amounts[Item.CategoryName] + Item.ExtendedAmount

quoteTable = Quote.QuoteTables['Invoice_Details']
quoteTable.Rows.Clear()

if Total_Amounts['Events'] > 0:
    invRow = quoteTable.AddNewRow()
    invRow['Invoice_Schedule'] = 'Knowledge & Other'
    invRow['Invoice_Date'] = "Upon Signature"
    invRow['Amount'] = Total_Amounts['Events']
    invRow['Estimated_Tax'] = 0
    invRow['Grand_Total'] = Total_Amounts['Events']

if Total_Amounts['Education Services'] > 0:
    invRow = quoteTable.AddNewRow()
    invRow['Invoice_Schedule'] = 'Training Fees'
    invRow['Invoice_Date'] = "Upon Signature"
    invRow['Amount'] = Total_Amounts['Education Services']
    invRow['Estimated_Tax'] = 0
    invRow['Grand_Total'] = Total_Amounts['Education Services']

Invoice_Schedule = ''
if No_Of_Months > 0 and No_Of_Days > 0:
    globals()['Invoice_Schedule'] = str(No_Of_Months) + ' Months and ' + str(No_Of_Days) + " days Subscription Fee"
elif No_Of_Months > 0:
    globals()['Invoice_Schedule'] = str(No_Of_Months) + " Months Subscription Fee"
else:
    globals()['Invoice_Schedule'] = str(No_Of_Days) + " Days Subscription Fee"

if Total_Amounts['Subscriptions'] > 0:
    if Invoice_Schedule != '':
        invRow = quoteTable.AddNewRow()
        invRow['Invoice_Schedule'] = Invoice_Schedule
        invRow['Invoice_Date'] = "Upon Signature"
        invRow['Amount'] = (Total_Amounts['Subscriptions'] / Total_No_Of_Months) * No_Of_Months
        invRow['Estimated_Tax'] = 0
        invRow['Grand_Total'] = (Total_Amounts['Subscriptions'] / Total_No_Of_Months) * No_Of_Months

    for count in range(No_Of_Years):
        invRow = quoteTable.AddNewRow()
        invRow['Invoice_Schedule'] = 'Annual Subscription Fee'
        if count == 0:
            invRow['Invoice_Date'] = "Upon Signature"
        else:
            Invoice_Date = Invoice_Date + timedelta(days=365)
            invRow['Invoice_Date'] = Invoice_Date.strftime("%B %d,%Y").ToString()
        invRow['Amount'] = (Total_Amounts['Subscriptions'] / Total_No_Of_Months) * 12
        invRow['Estimated_Tax'] = 0
        invRow['Grand_Total'] = (Total_Amounts['Subscriptions'] / Total_No_Of_Months) * 12

quoteTable.Save()
