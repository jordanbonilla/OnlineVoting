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

# Constant used in program.
LARGE_POSITIVE_INT = 1e15
# Global list of unique numbers used in vote validation
unique_nums = []
# Votes to ignore (invalid ID or repeated ID)
blacklist = []
# Global 2D array of worksheet holding all voter responses
all_data = []
# Global string holding the survey URL
survey_url = ''
# Global sting holding the title of the worksheet holding raw voter data
worksheet_title = ''
# Global worksheet object holding raw voting data
worksheet = None
# Global list of all emails invited to this survey
all_email_addresses = []
# Cooresponding names belonging to those emails 
all_names = []
# Holds all information about this vote and the calculation of results
all_output = ''
# Holds the user-specified subject line for emails
SUBJECT = ''

# Perform a normal print call but also write output to global string "all_output"
# which will be emailed out at the end of the survey
def print_write(in_string):
	print in_string
	global all_output
	all_output += in_string + '\n'
	
# Returns the number of responses in the survey so far
def get_num_responses():
	global worksheet
	if(worksheet == None):
		print_write("FATAL: no spreadsheet loaded")
		sys.exit(-1)
	first_col = worksheet.col_values(1)
	i = 0
	while(first_col[i] != ''):
		i = i + 1
	# Don't count the column header as a response
	return i - 1

# Returns the number of valid responses in the survey so far
def get_num_valid_responses():
	scopes = [ "https://docs.google.com/feeds/ https://spreadsheets.google.com/feeds/"]
	credentials = ServiceAccountCredentials.from_json_keyfile_name(
	    'OnlineVoting-e363607f6925.json', scopes=scopes)
	gc = gspread.authorize(credentials)
	sh = gc.open_by_key('1Pfzdngzcxt94iFSpPxf88TyMehsUcLS-zf5TovR0Ks8')
	num_responses = get_num_responses()
	# Grab worksheet linked to the current survey
	worksheet = sh.worksheet(worksheet_title)
	all_data = worksheet.get_all_values()
	first_row = all_data[0]
	num_cols = len(first_row)
	encountered_IDs = {}
	last_col = worksheet.col_values(num_cols)[1:num_responses + 1]
	num_invalid = 0
	for i in range(num_responses):
		if last_col[i] == '':
			num_invalid += 1
		elif last_col[i] not in unique_nums:
			num_invalid += 1
		elif last_col[i] in encountered_IDs:
			num_invalid += 1
		else:
			encountered_IDs[last_col[i]] = True
			
	return num_responses - num_invalid
	
	
# Read the private Google spreadsheet holding info for all eligible Avery voters
def get_all_elgible_email_address():
	scopes = [ "https://docs.google.com/feeds/ https://spreadsheets.google.com/feeds/"]
	credentials = ServiceAccountCredentials.from_json_keyfile_name(
	    'OnlineVoting-e363607f6925.json', scopes=scopes)
	gc = gspread.authorize(credentials)
	sh = gc.open_by_key('1Kodv_Fzz9Oki6q9w14jGddP49XFWD8VnXfFlxyViMVY');
	email_worksheet = sh.get_worksheet(0)
	first_col = email_worksheet.col_values(1)
	all_data = email_worksheet.get_all_values()
	emails = []

	# Skip over header entry
	for i in range(1, len(first_col)):
		if(first_col[i] != ''):
			row_index = i
			# 1st column is first name, 2nd column is nickname, 3rd column is last
			this_name = all_data[row_index][0]
			if(all_data[row_index][1] != ''):
				this_name += (' "' + all_data[row_index][1] + '"')
			this_name += (' ' + all_data[row_index][2])
			all_names.append(this_name)
			# The 4th column has the email data
			all_email_addresses.append(all_data[row_index][3])

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
	
	# Array of all filenames to send
	all_files = []
	
	print_write("Creating xlsx file with vote data...")
	# Create an new Excel file and add a worksheet.
	this_file_name = 'raw_votes.xlsx'
	workbook = xlsxwriter.Workbook(this_file_name)
	output_xlsx_file = workbook.add_worksheet()
	# Write data with row/column notation.
	num_rows = get_num_responses() + 1
	num_cols = len(all_data[0])
	for i in range(num_rows):
		for j in range(num_cols):
			output_xlsx_file.write(i, j, all_data[i][j])
	# Widen columns appropriately
	for i in range(num_cols):
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
	num_rows = len(all_names)
	num_cols = 2
	for i in range(num_rows):
		output_xlsx_file.write(i, 0, all_names[i])
		output_xlsx_file.write(i, 1, all_email_addresses[i])
	# Widen columns appropriately
	for i in range(num_cols):
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
	
	sender = "averyexcomm@gmail.com"
	server = smtplib.SMTP("smtp.gmail.com", 587)
	server.ehlo()
	server.starttls()
	server.login(sender, gmail_password)
	FROM = sender
	
	# Send name-tailored email to every eligible voter
	for i in range(num_averites):
		try:
			TO = all_email_addresses[i]
			first_name = all_names[i].split()[0]
			TEXT = "Hi again " + first_name + ', ' \
			+ "\n\nThe survey has closed and the votes have been counted." \
			+ "\nAll email addresses that were sent a link are in eligible_voters.xlsx:" \
			+ "\nRaw vote data is in raw_votes.xlsx:" \
			+ "\nRunoff results are in runoff_results.txt" \
			+ "\n\nThank you for keeping Avery great," \
			+ "\n\n<3 your ExComm" \
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
			print_write("Error: unable to send email to " + all_email_addresses[i])
			server.quit()
			e = sys.exc_info()[0]
			print e
			sys.exit(-1)
		# Don't risk sending too many emails in too short a time span
		time.sleep(2)

	print_write("Success. Total number of emails sent: " + str(num_averites) + " / " + str(num_averites))
	delete_sent_folder(sender, gmail_password)
	server.quit()

	
	
# If a voter ID is invalid, overwrite it with an error message and blacklist the vote
def identify_invalid_votes(num_cols, num_responses):
	global worksheet
	if(worksheet == None):
		print_write("FATAL: spreadsheet not read")
		sys.exit(-1)
	encountered_IDs = {}
	last_col = worksheet.col_values(num_cols)[1:num_responses + 1]
	for i in range(num_responses):
		if last_col[i] == '':
			blacklist.append(i)
			# +2 offset since the API calls are 1-indexed and we must account for header data
			worksheet.update_cell(i + 2, num_cols, "INVALID VOTE! EMPTY ID: " + last_col[i])
			# Make sure Google APIs calls aren't too close together - could cause error
			time.sleep(5)
		elif last_col[i] not in unique_nums:
			blacklist.append(i)
			worksheet.update_cell(i + 2, num_cols, "INVALID VOTE! UNAUTHORIZED ID: " + last_col[i])
			time.sleep(5)
		elif last_col[i] in encountered_IDs:
			blacklist.append(i)
			worksheet.update_cell(i + 2, num_cols, "INVALID VOTE! REPEATED ID: " + last_col[i])
			time.sleep(5)
		else:
			encountered_IDs[last_col[i]] = True
	#print num_responses, num_cols, last_col
	
# Generate time-seeded random number with n digits. Used to generate voter IDs.
def random_with_N_digits(n):
	# Time-seed random values
	random.seed
	range_start = 10**(n-1)
	range_end = (10**n)-1
	return random.randint(range_start, range_end)

# Send links to the survey, along with the unique voter IDs (embedded in URL)
def email_the_links(gmail_password):
	print_write("Sending emails to all eligible voters...")
	# Load all emails addresses in global variable "all_email_addresses"
	get_all_elgible_email_address()
	num_averites = len(all_email_addresses)
	if(num_averites == 0):
		print_write('FATAL: No email addresses')
		sys.exit(-1)
	elif(SUBJECT == ''):
		print_write('FATAL: No subject line')
		sys.exit(-1)
		
	sender = "averyexcomm@gmail.com"
	server = smtplib.SMTP("smtp.gmail.com", 587)
	server.ehlo()
	server.starttls()

	server.login(sender, gmail_password)
	FROM = sender

	unique_urls = [None] * num_averites
	for i in range(num_averites):
		unique_nums.append(str(random_with_N_digits(128)))
		unique_urls[i] = survey_url + unique_nums[i]

	BODY = \
			textwrap.fill("This is a unique link assigned to you. For this reason, " \
			+ "do not share this link or forward this email to anyone else. If you are not" \
			+ " a current Averite or know of a current Averite who did not receive a voting " \
			+ "link, please contact a member of ExComm so we can correct the eligible voter" \
			+ " mailing list. Lastly, please save this email message so that you have your" \
			+ " unique link on file should a vote's legitimacy fall into question.") \
			
	for i in range(num_averites):
		try:
			first_name = all_names[i].split()[0]
			TO = all_email_addresses[i]
			TEXT = "Hi " + first_name + ', ' \
			+ "\n\nHere is your link to vote: \n" + unique_urls[i] \
			+ "\n\n" + BODY \
			+ "\n\nThank you for keeping Avery great," \
			+ "\n\n<3 your ExComm"
			# Prepare actual message
			message = """\From: %s\nTo: %s\nSubject: %s\n\n%s
			""" % (FROM, TO, SUBJECT, TEXT)
			server.sendmail(FROM, TO, message)
		except:
			print_write("Error: unable to send email to " + all_email_addresses[i])
			server.quit()
			e = sys.exc_info()[0]
			print e
			sys.exit(-1)
		# Don't risk sending too many emails in too short a time span
		time.sleep(2)

	print_write("All unique links sent. Total number of emails sent: " + str(num_averites) + " / " + str(num_averites))
	
	delete_sent_folder(sender, gmail_password)
	server.quit()

# delete sent folder to ensure voter anonymity. This prevents association of
# voter ID with email address.
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

# Read in results from spreadsheet holding voter data and calculate winners
def get_results():
	global all_data
	global worksheet
	global worksheet_title
	if(worksheet_title == ''):
		print_write("FATAL: no worksheet title specified")
		sys.exit(-1)
		
	scopes = [ "https://docs.google.com/feeds/ https://spreadsheets.google.com/feeds/"]
	credentials = ServiceAccountCredentials.from_json_keyfile_name(
	    'OnlineVoting-e363607f6925.json', scopes=scopes)

	gc = gspread.authorize(credentials)
	sh = gc.open_by_key('1Pfzdngzcxt94iFSpPxf88TyMehsUcLS-zf5TovR0Ks8')
	# Grab worksheet linked to the current survey
	worksheet = sh.worksheet(worksheet_title)

	all_data = worksheet.get_all_values()
	position_delimiters = []
	position_encountered = {}
	first_row = all_data[0]
	
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
	
	# Print survey statistics
	num_cols = len(first_row)
	num_responses  = get_num_responses()
	print_write("Number of votes cast in this survey: " + str(num_responses))
	# Identify votes that are are invalid
	identify_invalid_votes(num_cols, num_responses)

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
		
# Entry point
if __name__ == "__main__":
	# URL retrived from manually-created Google survey 
	survey_url = raw_input('Enter survey URL:') #https://docs.google.com/forms/d/1Ql535WTs_w4UMFfdrwqx1-dxxPhSXO22jd80Hg6VpzQ/viewform?entry.1388757669='
	if(survey_url[-1] != '='):
		print_write('FATAL: Invalid spreadsheet URL, did you remember to change the URL to pre-fill ID?')
		sys.exit(-1)
	elif('https://docs.google.com/forms' not in survey_url):
		print_write("FATAL: Invalid spreadsheet URL. Please check for typos.")
		sys.exit(-1)
	# Title of worksheet. Established when creating the worksheet via Google form creation
	worksheet_title = raw_input('Enter worksheet title:') #"Voting Results"
	try:
		scopes = [ "https://docs.google.com/feeds/ https://spreadsheets.google.com/feeds/"]
		credentials = ServiceAccountCredentials.from_json_keyfile_name(
			'OnlineVoting-e363607f6925.json', scopes=scopes)
		gc = gspread.authorize(credentials)
		sh = gc.open_by_key('1Pfzdngzcxt94iFSpPxf88TyMehsUcLS-zf5TovR0Ks8')
		worksheet = sh.worksheet(worksheet_title)
		first_col = worksheet.col_values(1)
		# Read through the first column for any existing data. Skip header info
		for i in range(1, len(first_col)):	
			if(first_col[i] != ''):
				print_write("FATAL: specified worksheet is not blank. Do NOT reuse worksheets")
				sys.exit(-1)
	except:
		print_write("Invalid worksheet name. Exiting")
		sys.exit(-1)
	# Password known by all members of the ExComm
	gmail_password = getpass.getpass('[ECHO DISABLED] Enter averyexcomm password:') #"makeaverygreatagain"
	try:
		sender = "averyexcomm@gmail.com"
		server = smtplib.SMTP("smtp.gmail.com", 587)
		server.ehlo()
		server.starttls()
		server.login(sender, gmail_password)
		server.quit()
	except:
		write_print("Wrong gmail password name. Exiting")
		sys.exit(-1)
	# Subject as it will appear in emails
	SUBJECT = raw_input('Enter email subject:') #"Please Vote for Your Favorite Color and Animal"

	# 50% of the number of undergrads living in Avery
	QUORUM =  69
	# All input params are good. Email links to the survey
	email_the_links(gmail_password)
	
	# Give voters time to reach quorum in batches of 24 hours
	SECONDS_PER_DAY = 86400
	while(1):
		print_write('Waiting 24 hours for quorum to be reached...')
		time.sleep(SECONDS_PER_DAY)
		num_valid_votes = get_num_valid_responses()
		print_write("Number of valid votes so far: " + str(num_valid_votes) + ", Quorum: " + str(QUORUM))
		if(num_valid_votes >= QUORUM):
			print_write('Quorum Reached!')
			break
	# Read in results
	get_results()
	# Email results
	email_results(gmail_password)
