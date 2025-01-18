import smtplib
import time
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import ntpath
import types
import openpyxl
from openpyxl import worksheet
from openpyxl.utils import range_boundaries
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
import os
from loading import load_data, load_datas, create_file_names, month_hours, month_state_holidays_hours, month_names, report_input_file, report_attendance_file, month_days, state_holidays, day_names_short

def send_timesheet(sender, recipient, suplement_files, supplement_names):

	# Create the container (outer) email message.
	msg = MIMEMultipart()
	msg['Subject'] = f'SUBJECT'
	# me == the sender's email address
	# family = the list of all recipients' email addresses
	msg['From'] = sender
	msg['To'] = recipient
	msg.preamble = 'This is a multi-part message in MIME format.\n'

	with open('input/text.txt', 'r', encoding='utf-8') as content_file:
		msgtext = content_file.read()
	with open('input/html.html', 'r', encoding='utf-8') as content_file:
		htmlmsgtext = content_file.read()
		
	body = MIMEMultipart('alternative')
	body.attach(MIMEText(msgtext))
	body.attach(MIMEText(htmlmsgtext, 'html'))
	msg.attach(body)

	# Assume we know that the image files are all in PNG format
	maintype = 'application'
	subtype = 'vnd.openxmlformats-officedocument.spreadsheetml.sheet'
	for file_name, name in zip(suplement_files, supplement_names):
		fp = open(file_name, 'rb')
		attachment = MIMEBase(maintype, subtype)
		attachment.set_payload(fp.read())
		fp.close()
		# Encode the payload using Base64
		encoders.encode_base64(attachment)
		attachment.add_header('Content-Disposition', 'attachment', filename=name)
		msg.attach(attachment)

	# Send the email via our own SMTP server.
	s = smtplib.SMTP('SMTP')
	s.sendmail(sender, recipient, msg.as_string())
	s.quit()



persons = load_datas()
	

emailTofiles = {}
for person in persons:
	if person.key_activity is None:
		person.key_activity = "KA"
	
	filename = '{}/{}/{}/{}'.format('output', person.key_activity, person.spp, person.file_name)
	if person.email is not None:
		if not person.email in emailTofiles:
			emailTofiles[person.email] = [filename]
		else:
			emailTofiles[person.email].append(filename)

print("Person count: ", len(persons))
print("Unique emails:", len(emailTofiles))

count_total = 0
count_sent = 0

x = sorted(emailTofiles.keys());
for index, email in enumerate(sorted(emailTofiles.keys())):
	
	files = emailTofiles[email]
	
	file_paths = []
	file_names = []
	for f in files:
		file_paths.append(f)
		file_names.append(ntpath.basename(f))
		
	email_address = str(email).strip()
	
	print("{:5} {:40}{:2} ".format(index, email_address, len(file_paths)), end='', flush=True)
	count_total += 1
	
	send_timesheet('SENDER', email_address, file_paths, file_names)
	count_sent += 1
	print("SENT")
	# break
	time.sleep(13)
	
