from datetime import datetime
from datetime import timedelta
from System import DateTime

Order_Term_Start_Date = datetime(Quote.EffectiveDate).date()
Order_Term_End_Date_String = Quote.GetCustomField('Order Term End Date').Content
Order_Term_End_Date = datetime.strptime(Order_Term_End_Date_String, '%d/%m/%y').date()

No_Of_Years = (Order_Term_End_Date.year - Order_Term_Start_Date.year)
No_Of_Months = abs(Order_Term_End_Date.month - Order_Term_Start_Date.month)
No_Of_Days = abs(Order_Term_End_Date.day - Order_Term_Start_Date.day)

Invoice_Date = Order_Term_Start_Date

quoteTable = Quote.QuoteTables['Invoice_Details']
quoteTable.Rows.Clear()

Invoice_Schedule = ''
if No_Of_Months > 0 and No_Of_Days > 0:
	globals()['Invoice_Schedule'] = str(No_Of_Months) + ' Months and ' + str(No_Of_Days) + " days Subscription Fee"
elif No_Of_Months > 0:
	globals()['Invoice_Schedule'] = str(No_Of_Months) + " Months Subscription Fee"
else:
	globals()['Invoice_Schedule'] = str(No_Of_Days) + " Days Subscription Fee"

if Invoice_Schedule != '':
	invRow = quoteTable.AddNewRow()
	invRow['Invoice_Schedule'] = Invoice_Schedule
	invRow['Invoice_Date'] = "Upon Signature"
	invRow['Amount'] = Quote.Total.TotalAmount/No_Of_Years
	invRow['Estimated_Tax'] = 2.45
	invRow['Grand_Total'] = Quote.Total.TotalAmount/No_Of_Years

for count in range(No_Of_Years):
	invRow = quoteTable.AddNewRow()
	invRow['Invoice_Schedule'] = 'Annual Subscription Fee'
	if count == 0:
		invRow['Invoice_Date'] = "Upon Signature"
	else:
		Invoice_Date = Invoice_Date + timedelta(days=365)
		invRow['Invoice_Date'] = Invoice_Date.strftime("%B %d,%Y").ToString()
	invRow['Amount'] = Quote.Total.TotalAmount/No_Of_Years
	invRow['Estimated_Tax'] = 2.45
	invRow['Grand_Total'] = Quote.Total.TotalAmount/No_Of_Years

quoteTable.Save()
