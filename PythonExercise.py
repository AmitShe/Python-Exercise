# run: python PythomExercise.py [yourGmail@gmail.com] [Gmail password]

import xlrd
import datetime
import sys
import imaplib
import email

FROM_EMAIL  = sys.argv[1]
FROM_PWD    = sys.argv[2]
SMTP_SERVER = "imap.gmail.com"
SMTP_PORT   = 993

mail = imaplib.IMAP4_SSL(SMTP_SERVER)
mail.login(FROM_EMAIL,FROM_PWD) # user + password
mail.select('inbox')

type, data = mail.search(None, '(FROM "testing@outlook.com" SUBJECT "fileTesting")')

mail_ids = data[0]
id_list = mail_ids.split()
first_email_id = int(id_list[0])
latest_email_id = int(id_list[-1])

typ, data = mail.fetch(id_list[-1], '(RFC822)' )
for response_part in data:
    if isinstance(response_part, tuple):
        msg = email.message_from_string(response_part[1])
        email_subject = msg['subject']
        email_from = msg['from']
	for part in msg.walk():
		#find the attachment part
		if part.get_content_maintype() == 'multipart': continue
		if part.get('Content-Disposition') is None: continue

		#save the attachment in the program directory
		filename = part.get_filename()
		fp = open(filename, 'wb')
		fp.write(part.get_payload(decode=True))
		fp.close()
		print '%s saved!' % filename


for i in range(latest_email_id,first_email_id, -1):
            typ, data = mail.fetch(i, '(RFC822)' )

            for response_part in data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_string(response_part[1])
                    email_subject = msg['subject']
                    email_from = msg['from']
                    print 'From : ' + email_from + '\n'
                    print 'Subject : ' + email_subject + '\n'





book = xlrd.open_workbook(filename)  
sheetToRead = book.sheet_by_index(0)

for row in range(sheetToRead.nrows):
	valOfFirstCell = sheetToRead.cell_value(rowx=row, colx=0)
	if valOfFirstCell=="Date":
		ValueNamesRow = row
	elif valOfFirstCell=="Totals":
		TotalRow = row

num_cols = sheetToRead.ncols
for col_idx in range(0, num_cols):
	cell_obj = sheetToRead.cell_value(ValueNamesRow, col_idx)
	if cell_obj=="Date":
		dateCol = col_idx
	elif cell_obj=="App":
		appCol = col_idx
	elif cell_obj=="Campaign":
		OSCol = col_idx
	elif cell_obj=="Cost":
		CostCol = col_idx
		
#The Total Installs and Total Cost

Installs = sheetToRead.cell(TotalRow, sheetToRead.ncols-1).value
Cost = sheetToRead.cell(TotalRow, sheetToRead.ncols-2).value

print 'Total Installs:', '{:,.0f}'.format(Installs)
print 'Total Cost:', '${:,.2f}'.format(Cost), '\n'

#The Total Installs and Total Cost for each Date using the next format: "%Y-%m-%d"
#The Total Installs and Total Cost for each App and Platform (iOS/Android)

dateSum = {}
OSSum = {}
i = 4
while i<sheetToRead.nrows-1:
	date = sheetToRead.cell_value(rowx=i, colx=dateCol)
	app = sheetToRead.cell_value(rowx=i, colx=appCol)
	OSList = sheetToRead.cell_value(rowx=i, colx=OSCol).split() 
	OS = OSList[0]
	appAndOs = app + " running on " + OS
	cost = sheetToRead.cell_value(rowx=i, colx=CostCol)
	date_as_datetime = datetime.datetime(*xlrd.xldate_as_tuple(date, book.datemode)).strftime('%Y-%m-%d')
	if date_as_datetime in dateSum:
		dateSum[date_as_datetime]=dateSum[date_as_datetime]+cost
	else:
		dateSum[date_as_datetime]=cost
	
	if appAndOs in OSSum:
		OSSum[appAndOs]=OSSum[appAndOs]+cost
	else:
		OSSum[appAndOs]=cost
		
	i+=1
	
for x,y in dateSum.items():
	print 'For Date', x, 'the total cost was:',  '${:,.2f}'.format(y)
print ''
for x,y in OSSum.items():
	print 'For', x, 'the total cost was:',  '${:,.2f}'.format(y)