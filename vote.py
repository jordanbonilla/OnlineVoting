#!/usr/bin/env python2

#--------------------------------------------------------------------------
# Project: 
# Create a secure, transparent, anonymized, convinient online voting system
# for using in Avery House. This system interfaces Google surveys with 
# Python for a clean front end and back end. The voting system uses the 
# IRV (instnt runoff voting) strategy. 
#
# When the script finds that the voting period has ended, an email will be
# sent to all voters in the survey, informing them of the results and of 
# all the other voters who participated in the survey]
#
# See the README.md for more information
#
# Prerequisites: 
# Manual survey creation on the Google Forms website 
# required. Several gmail configuration steps and library installations are 
# required to get this working on a fresh project. See README.md
#
# Author: Jordan Bonilla
# Date  : March 2016
# License: All rights Reserved. See LICENSE.txt
#--------------------------------------------------------------------------

# Allows Google API calls
import gspread
from oauth2client.service_account import ServiceAccountCredentials
# Common headers
import sys
import random
import time
# Enables basic email functionality
import imaplib
import smtplib
import textwrap
# Headers needed for packing votes into xlsx and attaching to email
import xlsxwriter
from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
# Allows commandline text to be entered without echoing
import getpass
# Allows pinging to check internet connection
import subprocess
# Allows us to check what OS this script is running on
import os

# Constant used in program.
LARGE_POSITIVE_INT = 1e15
# The number of digits in each voter ID
VOTER_ID_LENGTH = 128
# The number of digits in the survey ID
SURVEY_ID_LENGTH = 4
# Global list of unique numbers used in vote validation
all_voter_ids = []
# Votes to ignore (invalid ID or repeated ID)
blacklist = []
# Global 2D array of worksheet holding all voter responses
all_data = []
# Global string holding the survey URL
survey_url = ''
# Global sting holding the title of the worksheet holding raw voter data
WORKSHEET_TITLE = ''
# Global list of all emails invited to this survey
all_email_addresses = []
# Cooresponding names belonging to those emails 
all_first_names = []
all_full_names = []
# Unique id for all voters that identify this particular excution of this script 
all_survey_ids = []
# Array of votes detected in the survey results. Used to make sure no votes are deleted.
# Encodes votes by concatenating all the data in that row into a string.
votes_seen_so_far = []
# Holds all information about this vote and the calculation of results
all_output = ''
# The account that to send all emails. Not tested with non-gmail accounts.
HOST_GMAIL_ACCOUNT = "averyexcomm@gmail.com"
# Param that allows APIs to read from Google docs
SCOPES = [ "https://docs.google.com/feeds/ https://spreadsheets.google.com/feeds/"]
# Key to Google spreadsheet that hosts the survey results. 
# Must explicitly link to your local machine.
LINKED_SPREADSHEET_KEY = "1Pfzdngzcxt94iFSpPxf88TyMehsUcLS-zf5TovR0Ks8"
# Json file in current directory with oauth2 credentials. Downloaded from Google API dashboard
SECRETS = "OnlineVoting-e363607f6925.json"
# Holds the user-specified subject line for emails
SUBJECT = ''
# Minimum number of votes reach quorum - 50% of the number of undergrads living in Avery
QUORUM =  69
# Give voters time to reach quorum (seconds). Gets renewed if quorum is not reached.
# Default 24 hrs
TIME_LIMIT_QUORUM = 86400
# The number of columns in the spreadsheet holding survey results.
# Make global to reduce Google API calls
NUM_COLS = -1
# Time to wait in between vote manipulation checks (seconds).
# Too low could cause Google API error
CHECKS_INTERVAL = 5

# Ensure the local machine is connected to the internet. Exit if no internet.
def verify_internet_access():
	success = True
	host = "8.8.8.8" # Google
	#Windows OS
	if(os.name == "nt"):
		output = subprocess.Popen(["ping.exe",host],stdout = subprocess.PIPE).communicate()[0]
		if("0% loss" in output):
			print "--- Internet status: OK ---\n"
		else:
			success = False
	# UNIX
	else:
		output = subprocess.Popen(['ping', '-c 1 -W 1 ', host], stdout=subprocess.PIPE).communicate()[0]
		if("0% packet loss" in output):
			print "--- Internet status: OK ---\n"
		else:
			success = False
	if(success == False):
		print output
		sys.exit(-1)
		
# Perform a normal print call but also write output to global string "all_output"
# which will be emailed out at the end of the survey
def print_write(in_string):
	print in_string
	global all_output
	all_output += in_string + '\n'
		
# Generate time-seeded random number with n digits. Used to generate voter IDs and pins
def random_with_N_digits(n):
	# Time-seed random values
	random.seed
	range_start = 10**(n-1)
	range_end = (10**n)-1
	return random.randint(range_start, range_end)

# Use credentials to return an authenticated worksheet object with voter data
def renewed_worksheet():
	if(WORKSHEET_TITLE == ''):
		write_print("FATAL: Worksheet title does not exit")
		sys.exit(-1)
	times_attempted = 0
	max_num_attempts = 5
	while(times_attempted < max_num_attempts):
		try:
			credentials = ServiceAccountCredentials.from_json_keyfile_name(SECRETS, scopes=SCOPES)
			gc = gspread.authorize(credentials)
			sh = gc.open_by_key(LINKED_SPREADSHEET_KEY)
			authenticated_worksheet = sh.worksheet(WORKSHEET_TITLE)
			time.sleep(2) # Ensure we don't call Google APIs too rapidly
			break
		except: #Maybe we are using APIs too much. Try again after waiting
			print "Unable to fetch worksheet. Trying again..."
			verify_internet_access()
			time.sleep(10)
			times_attempted = times_attempted + 1
	if(times_attempted == max_num_attempts):
		print "Unable to recover"
	else:
		return authenticated_worksheet

# Try to read a column from a recently authenticated voter data worksheet
def grab_col_safe(authenticated_worksheet, col_num):
	if(WORKSHEET_TITLE == ''):
		write_print("FATAL: Worksheet title does not exit")
		sys.exit(-1)
	times_attempted = 0
	max_num_attempts = 5
	while(times_attempted < max_num_attempts):
		try:
			requested_col = authenticated_worksheet.col_values(col_num)
			time.sleep(2) # Ensure we don't call Google APIs too rapidly
			break
		except: #Maybe we are using APIs too much. Try again after waiting
			print "Unable to grab column. Trying again..."
			verify_internet_access()
			time.sleep(10)
			authenticated_worksheet = renewed_worksheet()
			times_attempted = times_attempted + 1
	if(times_attempted == max_num_attempts):
		print "Unable to recover"
	else:
		return requested_col
	
# Try to read a row from a recently authenticated voter data worksheet
def grab_row_safe(authenticated_worksheet, row_num):
	if(WORKSHEET_TITLE == ''):
		write_print("FATAL: Worksheet title does not exit")
		sys.exit(-1)
	times_attempted = 0
	max_num_attempts = 5
	while(times_attempted < max_num_attempts):
		try:
			requested_row = authenticated_worksheet.row_values(row_num)
			time.sleep(2) # Ensure we don't call Google APIs too rapidly
			break
		except: #Maybe we are using APIs too much. Try again after waiting
			print "Unable to grab row.  Trying again..."
			verify_internet_access()
			time.sleep(10)
			authenticated_worksheet = renewed_worksheet()
			times_attempted = times_attempted + 1
	if(times_attempted == max_num_attempts):
		print "Unable to recover"
	else:
		return requested_row
	
# Try to read all data from a recently authenticated worksheet with voter data
def grab_all_data_safe(authenticated_worksheet):
	if(WORKSHEET_TITLE == ''):
		write_print("FATAL: Worksheet title does not exit")
		sys.exit(-1)
	times_attempted = 0
	max_num_attempts = 5
	while(times_attempted < max_num_attempts):
		try:
			requested_data = authenticated_worksheet.get_all_values()
			time.sleep(2) # Ensure we don't call Google APIs too rapidly
			break
		except: #Maybe we are using APIs too much. Try again after waiting
			print "Unable to grab all data. Try again later"
			verify_internet_access()
			time.sleep(10)
			authenticated_worksheet = renewed_worksheet()
			times_attempted = times_attempted + 1
	if(times_attempted == max_num_attempts):
		print "Unable to recover"
	else:
		return requested_data

# Try to update an entry in a worksheet hosting the voter data
def update_worksheet_cell_safe(worksheet, row, col, new_val):
	if(WORKSHEET_TITLE == ''):
		write_print("FATAL: Worksheet title does not exit")
		sys.exit(-1)
	times_attempted = 0
	max_num_attempts = 5
	while(times_attempted < max_num_attempts):
		try:
			worksheet.update_cell(row, col, new_val)
			time.sleep(2) # Ensure we don't call Google APIs too rapidly
			break
		except: #Maybe we are using APIs too much. Try again in 1 minute
			print "Unable to update cell. Try again later"
			verify_internet_access()
			time.sleep(10)
			worksheet = renewed_worksheet()
			times_attempted = times_attempted + 1
	if(times_attempted == max_num_attempts):
		print "Unable to recover"
		
# Return the first row of the specified worksheet, this is the header info.
# Need to reomve blank entries since Google spreadsheets pad blanks.
def get_first_row(worksheet):
	first_row = grab_row_safe(worksheet, 1)
	index_of_blank = -1
	for i in range(len(first_row)):
		if(first_row[i] == ''):
			index_of_blank = i
			break
	if(index_of_blank == -1):
		return first_row
	else:
		return first_row[0:index_of_blank]

# Grab the first row from the global variable all_data.
# Reduces number of Google API calls
def get_first_row_from_all_data():
	global all_data
	if(all_data == []):
		print "FATAL: all data not loaded"
		sys.exit(-1)
	first_row = all_data[0]
	# Account for padding
	index_of_blank = -1
	for i in range(len(first_row)):
		if(first_row[i] == ''):
			index_of_blank = i
			break
	if(index_of_blank == -1):
		return first_row
	else:
		return first_row[0:index_of_blank]
		
# Ascertain the number of columns in the survey results spreadsheet.
# We load this value in global memory to mitigate API calls later on.
def get_num_columns():
	global NUM_COLS
	worksheet = renewed_worksheet()
	first_row = get_first_row(worksheet)
	NUM_COLS = len(first_row)
	
# Returns the number of responses in the survey so far
def get_num_responses():
	if(WORKSHEET_TITLE == ''):
		print_write("FATAL: worksheet title not specified")
		sys.exit(-1)
		
	worksheet = renewed_worksheet()
	first_col = grab_col_safe(worksheet, 1)
	i = 0
	while(first_col[i] != ''):
		i = i + 1
	# Don't count the column header as a response
	return i - 1

# get number of responses from a voting data worksheet directly. 
# Must make sure the worksheet was recently renewed or this won't work.
# Prevents an extra API call
def get_num_responses_on_recently_renewed_worksheet(renewed_worksheet):
	if(WORKSHEET_TITLE == ''):
		print_write("FATAL: worksheet title not specified")
		sys.exit(-1)
	try:
		first_col = grab_col_safe(renewed_worksheet, 1)
		i = 0
		while(first_col[i] != ''):
			i = i + 1
		# Don't count the column header as a response
		return i - 1
	except:
		# Perhaps it was was not recently renewed so call the version of
		# this function that explicitly renews a worksheet
		return get_num_responses


# Returns the number of valid responses in the survey so far
def get_num_valid_responses():
	if(WORKSHEET_TITLE == ''):
		print_write("FATAL: worksheet title not specified")
		sys.exit(-1)
	elif(NUM_COLS == -1):
		print_write("FATAL: number of spreadsheet columns not specified")
		sys.exit(-1)
	# Grab worksheet linked to the current survey
	worksheet = renewed_worksheet()
	num_responses = get_num_responses_on_recently_renewed_worksheet(worksheet)

	encountered_IDs = {}
	last_col = grab_col_safe(worksheet, NUM_COLS)[1:num_responses + 1]
	num_invalid = 0
	for i in range(num_responses):
		if last_col[i] == '':
			num_invalid += 1
		elif last_col[i] not in all_voter_ids:
			num_invalid += 1
		elif last_col[i] in encountered_IDs:
			num_invalid += 1
		else:
			encountered_IDs[last_col[i]] = True
			
	return num_responses - num_invalid

# Called if an email could not be sent. Re-try the email 
def email_recovery(server, FROM, TO, message, recipient):
	# Perhaps it was a network error? Wait and try again
	times_attempted = 1
	max_attempts = 5
	while(times_attempted < max_attempts):
		print "Previous email failed. Retying email to " + recipient
		try:
			time.sleep(10)
			server.sendmail(FROM, TO, message)
			break
		except:
			times_attempted = times_attempted + 1
	if(times_attempted == max_attempts):
			print_write("FATAL: unable to send email to " + recipient)
			server.quit()
			sys.exit(-1)
			
# If a voter ID is invalid, overwrite it with an error message and blacklist the vote
def identify_invalid_votes(num_responses):
	global all_data
	if(WORKSHEET_TITLE == ''):
		print_write("FATAL: worksheet title not specified")
		sys.exit(-1)
	elif(NUM_COLS == -1):
		print_write("FATAL: number of spreadsheet columns not specified")
		sys.exit(-1)
	elif(all_data == []):
		print_write("FATAL: voter data not loaded into global variable al_data")
		sys.exit(-1)
		
	# Grab worksheet linked to the current survey
	worksheet = renewed_worksheet()
	encountered_IDs = {}
	last_col = grab_col_safe(worksheet, NUM_COLS)[1:num_responses + 1]
	for i in range(num_responses):
		if last_col[i] == '':
			blacklist.append(i)
			# +2 offset since the API calls are 1-indexed and we must account for header data
			update_worksheet_cell_safe(worksheet, i + 2, NUM_COLS, "INVALID VOTE! EMPTY ID: " + last_col[i])
			# Update local data to have this markup
			all_data[i+1][NUM_COLS - 1] = "INVALID VOTE! EMPTY ID: " + last_col[i]
		elif last_col[i] not in all_voter_ids:
			blacklist.append(i)
			update_worksheet_cell_safe(worksheet, i + 2, NUM_COLS, "INVALID VOTE! EMPTY ID: " + last_col[i])
			all_data[i+1][NUM_COLS - 1] = "INVALID VOTE! UNAUTHORIZED ID: " + last_col[i]
		elif last_col[i] in encountered_IDs:
			blacklist.append(i)
			update_worksheet_cell_safe(worksheet, i + 2, NUM_COLS, "INVALID VOTE! EMPTY ID: " + last_col[i])
			all_data[i+1][NUM_COLS - 1] = "INVALID VOTE! REPEATED ID: " + last_col[i]
		else:
			encountered_IDs[last_col[i]] = True

# Read the private Google spreadsheet holding info for all eligible Avery voters
def get_all_elgible_email_address():
	global all_email_addresses
	global all_first_names
	all_email_addresses.append("jordan@caltech.edu")
	all_email_addresses.append("sunbonilla@yahoo.com")
	all_first_names.append("Jordan")
	all_first_names.append("sun")
	return

	credentials = ServiceAccountCredentials.from_json_keyfile_name(SECRETS, scopes=SCOPES)
	gc = gspread.authorize(credentials)
	sh = gc.open_by_key('1Kodv_Fzz9Oki6q9w14jGddP49XFWD8VnXfFlxyViMVY');
	email_worksheet = sh.get_worksheet(0)
	first_col = grab_col_safe(email_worksheet, 1)
	all_data = grab_all_data_safe(email_worksheet)
	emails = []

	# Skip over header entry
	for i in range(1, len(first_col)):
		if(first_col[i] != ''):
			row_index = i
			# 1st column is first name, 2nd column is nickname, 3rd column is last
			full_name = all_data[row_index][0]
			if(all_data[row_index][1] != ''):
				first_name = all_data[row_index][1]
				full_name += ' "' + all_data[row_index][1] + '"'
			else:
				first_name = all_data[row_index][0]
			full_name += ' ' + all_data[row_index][2]
			
			all_first_names.append(first_name)
			all_full_names.append(full_name)
			# The 4th column has the email data
			all_email_addresses.append(all_data[row_index][3])

# delete sent folder to ensure voter anonymity. This prevents association of
# voter ID with email address and prevents pins from being recovered
def delete_sent_folder(sender, password):
	m = imaplib.IMAP4_SSL("imap.gmail.com")  # server to connect to
	m.login(sender, password)

	# Move sent folder to trash
	print_write("Moving sent folder to trash... ")
	print_write(str(m.select('[Gmail]/Sent Mail')))
	m.store("1:*",'+X-GM-LABELS', '\\Trash')
	m.expunge()
	
	#This block empties trash
	print_write("Emptying Trash...")
	print_write(str(m.select('[Gmail]/Trash')))  # select all trash
	m.store("1:*", '+FLAGS', '\\Deleted')  #Flag all Trash as Deleted
	m.expunge()  # not need if auto-expunge enabled

	m.close()
	m.logout()
	print_write("Sent folder successfully deleted.\n")	
	
# Send the following data to all eligible voters:
# 1) list of all email addresses invited to the survey (not visible to public)
# 2) text file of instant runoff results containing all program output
# 3) xlsx file of the raw vote counts pulled directly from the Google spreadsheet
def email_results(gmail_password):
	num_averites = len(all_email_addresses)
	if(num_averites == 0):
		print_write('FATAL: No email addresses')
		sys.exit(-1)
	elif(SUBJECT == ''):
		print_write('FATAL: No subject line')
		sys.exit(-1)
	elif(all_output == ''):
		print_write('FATAL: No output')
		sys.exit(-1)
	elif(all_data == []):
		print_write('FATAL: no spreadsheet data')
		sys.exit(-1)
	elif(NUM_COLS == -1):
		print_write("FATAL: number of spreadsheet columns not specified")
		sys.exit(-1)
	elif(all_survey_ids == []):
		print_write("FATAL: Survey IDs not generated")
		sys.exit(-1)
		
	# Array of all filenames to send
	all_files = []
	
	print_write("Creating xlsx file with vote data...")
	# Create an new Excel file and add a worksheet.
	this_file_name = 'raw_votes.xlsx'
	workbook = xlsxwriter.Workbook(this_file_name)
	output_xlsx_file = workbook.add_worksheet()
	# Write data with row/column notation.
	num_rows = get_num_responses() + 1
	for i in range(num_rows):
		for j in range(NUM_COLS):
			output_xlsx_file.write(i, j, all_data[i][j])
	# Widen columns appropriately
	for i in range(NUM_COLS):
		if(i == 0):
			output_xlsx_file.set_column(i,i, 20)
		else:
			output_xlsx_file.set_column(i,i, len(all_data[0][i]))
	workbook.close()
	print_write("SUCCESS!")
	all_files.append(this_file_name)
	
	print_write("Creating xlsx file with eligible voter data...")
	# Create an new Excel file and add a worksheet.
	this_file_name = 'eligible_voters.xlsx'
	workbook = xlsxwriter.Workbook(this_file_name)
	output_xlsx_file = workbook.add_worksheet()
	# Write data with row/column notation.
	num_rows = len(all_first_names)
	cols = 2
	for i in range(num_rows):
		output_xlsx_file.write(i, 0, all_full_names[i])
		output_xlsx_file.write(i, 1, all_email_addresses[i])
	# Widen columns appropriately
	for i in range(cols):
		output_xlsx_file.set_column(i,i, 30)
	workbook.close()
	print_write("SUCCESS!")
	all_files.append(this_file_name)
	
	# Create text file with runoff results
	this_file_name = "runoff_results.txt"
	print_write("Sending results emails.")
	file = open(this_file_name, "w")
	file.write(all_output)
	file.close()
	all_files.append(this_file_name)
	
	sender = HOST_GMAIL_ACCOUNT
	server = smtplib.SMTP("smtp.gmail.com", 587)
	server.ehlo()
	server.starttls()
	server.login(sender, gmail_password)
	FROM = sender
	
	# Send name-tailored email to every eligible voter with their custom pin
	for i in range(num_averites):
		try:
			TO = all_email_addresses[i]
			TEXT = "Hi again " + all_first_names[i] + ', ' \
			+ "\n\nThe survey has closed and the votes have been counted." \
			+ "\nAll email addresses that were sent a link are in eligible_voters.xlsx" \
			+ "\nRaw vote data is in raw_votes.xlsx" \
			+ "\nRunoff results are in runoff_results.txt" \
			+ "\n\nThank you for keeping Avery great," \
			+ "\n\n<3 your ExComm" \
			+ "\nSurvey ID (should match the first email): " + str(all_survey_ids[i]) \
			+ "\nGithub repo: https://github.com/jordanbonilla/OnlineVoting"

			# Prepare actual message
			message = MIMEMultipart()
			message['From'] = FROM
			message['To'] = TO
			message['Subject'] = "*RESULTS* " + SUBJECT
			message.attach(MIMEText(TEXT))

			# Attach all files
			for j in range(len(all_files)):
				file_name = all_files[j]
				with open(file_name, "rb") as fil:
					message.attach(MIMEApplication(
						fil.read(),
						Content_Disposition='attachment; filename="%s"' % basename(file_name),
						Name=basename(file_name)
						))
			# Send
			server.sendmail(FROM, TO, message.as_string())
		except:
			email_recovery(server, FROM, TO, message.as_string(), TO)
		# Don't risk sending too many emails in too short a time span
		time.sleep(2)
		# Write progress to local terminal
		sys.stdout.write('\r')
		sys.stdout.write("[%-20s] %d%%" % ('='*(20 * (i+1)/num_averites), 100 * (i+1)/num_averites))
		sys.stdout.flush()
	
	print_write("\nSuccess. Total number of emails sent: " + str(num_averites) + " / " + str(num_averites))
	delete_sent_folder(sender, gmail_password)
	server.quit()

	from time import sleep
import sys
	
# Send links to the survey, along with the unique voter IDs (embedded in URL), and survey pin
def email_the_links(gmail_password):
	print_write("Sending emails to all eligible voters...")
	if(all_email_addresses == []):
		write_print("FATAL: no email addresses loaded")
		sys.exit(-1)
	num_averites = len(all_email_addresses)
	if(num_averites == 0):
		print_write('FATAL: No email addresses')
		sys.exit(-1)
	elif(SUBJECT == ''):
		print_write('FATAL: No subject line')
		sys.exit(-1)
		
	sender = HOST_GMAIL_ACCOUNT
	server = smtplib.SMTP("smtp.gmail.com", 587)
	server.ehlo()
	server.starttls()

	server.login(sender, gmail_password)
	FROM = sender

	global all_voter_ids
	global all_survey_ids
	unique_urls = []
	# Generate voter IDs and pins
	used_voter_ids = {}
	for i in range(num_averites):
		this_voter_id = random_with_N_digits(VOTER_ID_LENGTH)
		while(this_voter_id in used_voter_ids):
			this_voter_id = random_with_N_digits(VOTER_ID_LENGTH)
		used_voter_ids[this_voter_id] = True
		all_voter_ids.append(str(this_voter_id))
		all_survey_ids.append(str(random_with_N_digits(SURVEY_ID_LENGTH)))
		unique_urls.append(survey_url + all_voter_ids[i])
		
	BODY = \
			textwrap.fill("This is a unique link assigned to you. For this reason, " \
			+ "do not share this link or forward this email to anyone else. If you are not" \
			+ " a current Averite or know of a current Averite who did not receive a voting " \
			+ "link, please contact a member of ExComm so we can correct the eligible voter" \
			+ " mailing list. Lastly, please save this email message so that you have your" \
			+ " unique pin and url on file should a vote's legitimacy fall into question.") \
			
	for i in range(num_averites):
		try:
			TO = all_email_addresses[i]
			TEXT = "Hi " + all_first_names[i] + ', ' \
			+ "\n\nHere is your link to vote: \n" + unique_urls[i] \
			+ "\n\n" + BODY \
			+ "\n\nThank you for keeping Avery great," \
			+ "\n\n<3 your ExComm" \
			+ "\nSurvey ID: " + str(all_survey_ids[i]) \
			+ "\nGithub repo: https://github.com/jordanbonilla/OnlineVoting"
			# Prepare actual message
			message = """\From: %s\nTo: %s\nSubject: %s\n\n%s
			""" % (FROM, TO, SUBJECT, TEXT)
			server.sendmail(FROM, TO, message)
		except:
			email_recovery(server, FROM, TO, message, TO)
		# Don't risk sending too many emails in too short a time span
		time.sleep(2)
		# Write progress to local terminal
		sys.stdout.write('\r')
		sys.stdout.write("[%-20s] %d%%" % ('='*(20 * (i+1)/num_averites), 100 * (i+1)/num_averites))
		sys.stdout.flush()
		
	print_write("\nAll unique links sent. Total number of emails sent: " + str(num_averites) + " / " + str(num_averites))
	
	delete_sent_folder(sender, gmail_password)
	server.quit()

# Read in results from spreadsheet holding voter data and calculate winners
def get_results():
	global all_data
	if(WORKSHEET_TITLE == ''):
		print_write("FATAL: no worksheet title specified")
		sys.exit(-1)
		
	worksheet = renewed_worksheet()
	num_responses = get_num_responses_on_recently_renewed_worksheet(worksheet)
	
	all_data = grab_all_data_safe(worksheet)
	position_delimiters = []
	position_encountered = {}
	first_row = get_first_row_from_all_data()
	
	position_names = []
	candidates_adjoined = []
	# Grab names of candidates cooresponding to each position
	# Skip first col (timestamp), last col (unique ID)
	for i in range(1, len(first_row) - 1):
		parsed = first_row[i].split('[')
		this_position = parsed[0]
		this_candidate = parsed[1][:-1]
		if(this_position not in position_names):
			position_delimiters.append(i)
			position_encountered[this_position] = True
			position_names.append(this_position)
			candidates_adjoined.append(this_candidate)
			
		else:
			candidates_adjoined[-1] += ('\t' + this_candidate)
	
	print_write("Number of votes cast in this survey: " + str(num_responses))
	# Identify votes that are are invalid
	identify_invalid_votes(num_responses)

	num_invalid = len(blacklist)
	print_write("Number of invalid votes: " + str(num_invalid))
	print_write("Number of valid votes: " + str(num_responses - num_invalid))
	print_write("Refer to raw vote data for more information")
	num_positions = len(candidates_adjoined)
	print_write("\nNumber of positions to assign in this election: " + str(num_positions))
	
	# Split the adjoined candidate lists
	candidates_split = []
	for i in range(num_positions):
		print_write("\nPosition Title: " + position_names[i])
		all_candidates_for_this_position = candidates_adjoined[i].split('\t')
		candidates_split.append(all_candidates_for_this_position)
		for j in range(len(all_candidates_for_this_position)):
			print_write("    Candidate #" + str(j + 1) + ": " + all_candidates_for_this_position[j])

	# Find the winners
	print_write('\nBegin Runoff\n_____________________________________________\n')
	for i in range(num_positions):
		print_write(position_names[i])
		num_candidates = len(candidates_split[i])
		remaining_candidates = range(num_candidates)
		start_index = position_delimiters[i] # starting column

		# Proceed with runoff
		round_num = 1
		while(len(remaining_candidates) > 1):
			print_write("Round: " + str(round_num))
			count_spread = [0] * num_candidates
			run_off(count_spread, remaining_candidates, start_index, num_candidates, \
			num_positions, num_responses)
			for j in range(num_candidates):
				if(j in remaining_candidates or count_spread[j] is not 0):
					print_write("    " + candidates_split[i][j] + " - " + str(count_spread[j]))
			for j in remaining_candidates:
				if(len(remaining_candidates) == 1):
					print_write("    CONGRATULATIONS WINNER: " + candidates_split[i][j])
				else:
					print_write("    Advance: " + candidates_split[i][j])
			round_num = round_num + 1
			print_write('\n')

		print_write('_____________________________________________\n')

# Perform one iteration/round of instant run off voting
def run_off(count_spread, remaining_candidates, start_index, num_candidates, \
num_positions, num_responses):
	# Make sure data was read before this function was called
	if(len(all_data) is 0):
		print_write('FATAL: Worksheet not populated')
		sys.exit(-1)
		
	# Get the votes from all voters, taking into account elminated candidates and invalid votes
	for i in range(num_responses):
		if(i in blacklist):
			continue
			
		relevant_votes = all_data[i + 1][start_index : start_index + num_candidates]
		final_vote = -1; # index in "relevant_votes" designating the desired candidate
		eliminated = [] # eliminated candidates or 
		while (final_vote == -1 and len(eliminated) < len(relevant_votes)):
			best_rank = LARGE_POSITIVE_INT
			index_of_best_rank = -1
			# Eliminate NULL votes
			for j in range(len(relevant_votes)):
				if(relevant_votes[j] == ''):
					eliminated.append(j)
			# Scan for highest eligble vote
			for j in range(len(relevant_votes)):
				if(j not in eliminated):
					if int(relevant_votes[j]) < best_rank:
						best_rank = int(relevant_votes[j])
						index_of_best_rank = j
			if(index_of_best_rank in remaining_candidates):
				final_vote = index_of_best_rank
			else:
				eliminated.append(index_of_best_rank) 
		# Record this vote
		if(final_vote is not -1):
			count_spread[index_of_best_rank] += 1
			
	# Coordesponding minimum number of votes for the candidate[s] to evict
	min_for_this_round = LARGE_POSITIVE_INT;
	for i in remaining_candidates:
		if(count_spread[i] < min_for_this_round):
			min_for_this_round = count_spread[i]
	
	# Check for strict majority win 
	max = -1
	index_of_max = -1
	for i in remaining_candidates:
		if(count_spread[i] > max):
			max = count_spread[i]
			index_of_max = i
	if(max * 2 > sum(count_spread) ):
		print_write('    Strict majority victory for candidate #' + str(index_of_max + 1))
		# Remove all other candidates from "remaining_candidates" array
		del remaining_candidates[:]
		remaining_candidates.append(index_of_max)
		return
	
	# No strict majority. Begin the elimination process.
	min_indices = []
	for i in remaining_candidates:
		if(count_spread[i] == min_for_this_round):
			min_indices.append(i)
			
	# There was a tie. Print voting data for manual sorting and print error message.
	if(len(min_indices) > 1):
		print("    Draw for lowest vote count. ")
		for i in min_indices:
			print_write("    candidate#" + str(i + 1) + " vote count: " + str(count_spread[i]))
		print_write("    Refer to constitution for tie-breaking procedure.") 
		print_write("    All votes and their associated frequencies:")
		encountered_vote_patterns = {}
		for i in range(num_responses):			
			relevant_votes = all_data[i + 1][start_index : start_index + num_candidates]
			this_vote_pattern = ''
			for j in range(num_candidates):
				if(relevant_votes[j] == ''):
					this_vote_pattern+= "X"
				else:
					this_vote_pattern+= str(relevant_votes[j])
			if this_vote_pattern in encountered_vote_patterns:
				encountered_vote_patterns[this_vote_pattern] += 1
			else:
				encountered_vote_patterns[this_vote_pattern] = 1
		print_write("    " + str(encountered_vote_patterns))
		del remaining_candidates[:]
	# There was no tie. Exactly one candidate to eliminate
	else:	
		remaining_candidates.remove(min_indices[0])

# Comapre a list of verified votes with the current set of votes to make sure the current
# set of votes is a superset of the verified set of votes
def ensure_no_votes_manipulated():
	num_averites = len(all_email_addresses)
	if(num_averites == 0):
		print_write('FATAL: No email addresses')
		sys.exit(-1)
	elif(SUBJECT == ''):
		print_write('FATAL: No subject line')
		sys.exit(-1)
	elif(WORKSHEET_TITLE == ''):
		print_write("FATAL: worksheet title not specified")
		sys.exit(-1)
	elif(all_email_addresses == []):
		print_write("FATAL: no emails specified")
		sys.exit(-1)
	elif(NUM_COLS == -1):
		print_write("FATAL: number of spreadsheet columns not specified")
		sys.exit(-1)
		
	worksheet = renewed_worksheet()
	num_responses = get_num_responses_on_recently_renewed_worksheet(worksheet)
	latest_data = grab_all_data_safe(worksheet)
	latest_votes = {}
	# Encode vote info into a string
	for i in range(num_responses):
		relevant_votes = latest_data[i + 1] # +1 to account for header row
		encoded_vote = ''
		for j in range(len(relevant_votes)):
			encoded_vote += str(relevant_votes[j])
		latest_votes[encoded_vote] = True
		
	# Array of all encoded votes seen so far
	global votes_seen_so_far
	num_votes_seen_so_far = len(votes_seen_so_far)
	# Check that all existing votes exist in the latest batch of pulled votes
	for i in range(num_votes_seen_so_far):
		if votes_seen_so_far[i] not in latest_votes:
			email_tamper_notification()
	# Add any new votes to existing votes
	for encoded_vote in latest_votes.keys():
		if encoded_vote not in votes_seen_so_far:
			votes_seen_so_far.append(encoded_vote)
					
# Vote tamper detected. Email all eligible voters and exit.
def email_tamper_notification():
	sender = HOST_GMAIL_ACCOUNT
	server = smtplib.SMTP("smtp.gmail.com", 587)
	server.ehlo()
	server.starttls()
	server.login(sender, gmail_password)
	FROM = sender
	
	num_averites = len(all_first_names)
	for i in range(num_averites):
		try:
			TO = all_email_addresses[i]
			TEXT = "Hi " + all_first_names[i] + ', ' \
			+ "\n\nThe vote has been terminated due to detection of vote manipulation\n" \
			+ "\n\nSurvey ID (should match the first email): " + str(all_survey_ids[i]) \
			+ "\nGithub repo: https://github.com/jordanbonilla/OnlineVoting"
			# Prepare actual message
			message = """\From: %s\nTo: %s\nSubject: %s\n\n%s
			""" % (FROM, TO, "*CANCELED* " + SUBJECT, TEXT)
			server.sendmail(FROM, TO, message)
		except:
			email_recovery(server, FROM, TO, message, TO)
		# Don't risk sending too many emails in too short a time span
		time.sleep(2)
	print_write("Vote terminated due to a vote being manipulated")
	server.quit()
	sys.exit(-1)
		
# Verify survey URL meets specifications
def verify_survey(survey_url):
	if(survey_url[-1] != '='):
		print_write('FATAL: Invalid spreadsheet URL, did you remember to change the URL to pre-fill ID?')
		sys.exit(-1)
	elif('https://docs.google.com/forms' not in survey_url):
		print_write("FATAL: Invalid spreadsheet URL. Please check for typos.")
		sys.exit(-1)
		
# Make sure that the worksheet exists on the spreadsheet specified by LINKED_SPREADSHEET_KEY
def verify_voter_data_worksheet():
	if(WORKSHEET_TITLE == ''):
		write_print("FATAL: Global WORKSHEET_TITLE not populated")
		sys.exit(-1)
	try:
		worksheet = renewed_worksheet()
	except:
		print_write("Worksheet does not exist in spreadsheet. Exiting")
		sys.exit(-1)

	# Read back the election info so it can be confirmed
	first_row = get_first_row(worksheet)
	encountered_positions = []
	candidates_per_position = []
	candidate_names = []
	for i in range(1, len(first_row) - 1):	# Skip timestamp entry and voter ID entry
		parsed = first_row[i].split('[')
		this_position = parsed[0]
		this_candidate = (parsed[1])[:-1]
		if this_position not in encountered_positions:
			encountered_positions.append(this_position)
			candidates_per_position.append(1)
			candidate_names.append(this_candidate)
		else:
			candidates_per_position[-1] += 1
			candidate_names[-1] += ", " + this_candidate
	for i in range(len(encountered_positions)):
		print "\nPosition #" + str(i) + ": " + encountered_positions[i]
		print str(candidates_per_position[i]) + " candidates:"
		print str(candidate_names[i])
		
	print "\nVoter ID column name:\n" + first_row[-1]
	user_input = raw_input("\nConfirm the above information to continue [y/n] ")
	if(user_input.lower() != 'y'):
		sys.exit()
		
	# Read through the first column for any existing data. 
	first_col = grab_col_safe(worksheet, 1)
	for i in range(1, len(first_col)):	#Skip header info
		if(first_col[i] != ''):
			print_write("FATAL: specified worksheet is not blank. Do NOT reuse worksheets")
			sys.exit(-1)
			
# Make sure that the gmail password associated with the host is correct
def verify_gmail_pass(gmail_password):
	try:
		sender = HOST_GMAIL_ACCOUNT
		server = smtplib.SMTP("smtp.gmail.com", 587)
		server.ehlo()
		server.starttls()
		server.login(sender, gmail_password)
		server.quit()
	except:
		write_print("Wrong gmail password name. Exiting")
		sys.exit(-1)
		
# Entry point
if __name__ == "__main__":
	verify_internet_access()
	# URL retrived from manually-created Google survey 
	survey_url = raw_input('Enter survey URL:') 
	verify_survey(survey_url)
	# Title of worksheet holding voter data.
	# Established when creating the worksheet via Google form creation.
	WORKSHEET_TITLE = raw_input('Enter worksheet title:') 
	verify_voter_data_worksheet()
	# Password known by all members of the ExComm
	gmail_password = getpass.getpass('[ECHO DISABLED] Enter averyexcomm password:')
	verify_gmail_pass(gmail_password)
	# Subject as it will appear in emails
	SUBJECT = raw_input('Enter email subject:')
	
	# Load all email addresses to be used in this survey into global variable "all_email_addresses"
	get_all_elgible_email_address()
	# Load the number of columns in the spreadsheet into global variable to minimize API calls
	get_num_columns()
	# All input params are good. Email links to the survey
	email_the_links(gmail_password)
	
	# Time-seed random values
	random.seed
	while(1):
		print_write('Waiting 24 hours for quorum to be reached...')
		elapsed_seconds = 0.0
		while(elapsed_seconds < TIME_LIMIT_QUORUM):
			start_time = time.time()
			random_time_slice = CHECKS_INTERVAL + random.randint(1, 5) # Add variability for security
			time.sleep(random_time_slice)
			ensure_no_votes_manipulated()
			elapsed_time = time.time() - start_time
			elapsed_seconds += elapsed_time
		num_valid_votes = get_num_valid_responses()
		print_write("Number of valid votes so far: " + str(num_valid_votes) + ", Quorum: " + str(QUORUM))
		if(num_valid_votes >= QUORUM):
			print_write('Quorum Reached!')
			break
	# Read in results
	get_results()
	# Email results
	email_results(gmail_password)
