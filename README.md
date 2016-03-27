## Caltech Avery House Online Voting Project 

# Goals:
1. Create a free, secure, transparent, anonymized, convinient online voting system.
2. Remove the need for "middleman" survey services.
3. Avery House glory.

# Summary:
This framework interfaces Google surveys with Python for a clean front end and back end. The voting system uses the IRV (instnt runoff voting) strategy. The workflow can be summarized as:

1. Manually create a Google survey and link it to a Google spreadsheet. Enter the link to the survey into the Python script.

2. Python script generates unique links to the survey and emails out links. One link per email addresses. These links can be hardcoded or read from a Google doc (default)

3. Results automatically go into the linked Google Spreadsheet

4. When the voting period has ended, the Python script reads in the results from this spreadsheet, parses it, and calculates the winners based on instant runoff. The results are automatically emailed out.


# Security and anonymity:

1. The python script generates a unique 128-digit-time-seeded random int for each email address that is eligible to vote. This number is embedded in the URL of the Google survey link and is automatically added to the “ID” field in the survey. When votes are read from the Python script, it will check the “ID” field of submitted votes and confirm that they corresponding to originally-generated ID.This makes it so duplicate and unauthorized votes can be invalidated.


2. The 128-digit IDs are not stored anywhere except in the python script’s internal state. This is possible because the script does not finish executing until the final email results are sent out. This prevents manipulation of IDs, ensures that vote results can not be modified, and maintains confidentiality of voters. Furthermore, the “sent” folder of the gmail account hosting the survey has its sent box deleted after the emails get sent - this prevents people with access to the email account from seeing which IDs were associated with specific email addresses. 


3. Votes that are manually added to the spreadsheet will not have a valid 128-digit entry and will be invalidated in the final count. Votes that are deleted from the spreadsheet will raise attention since the email with the election results will contain a list of all data from the vote, including the IDs associated with each vote. If someone believes their vote has been deleted, they can compare their voter ID (in the email) with the voter IDs found in the raw vote data which is automatically emailed at the end of a vote.

4. Adding non-eligible emails to the eligible voter list is easily spotted since the email with the results also contains a list of all emails (xlsx format) that participated in the survey. All eligible voters are emailed the raw vote data (xlsx format) in the results email. 




# Prerequisitites for running locally:

1. OAuth2 credentials in json format.
      - http://gspread.readthedocs.org/en/latest/oauth2.html
      

2. Python libraries:
      - gspread
      - oauth2client
      - xlsxwriter
  
  
3. A manually-created Google Survey
      - The survey must be a multiple choice grid 
      - For every candidate, make a new column with a "rank"
      - Ranks go from 1 to num_candidates with 1 being the best
      - Last question must be a short answer question that holds voter ID
      - Get the URL from the pre-fill link such that values appended to the URL automatically fill in "voter ID"
      - This URL looks something like: https://docs.google.com/forms/d/xxx...xxx/viewform?entry.1299711861=
      - Required survey option: 1 response per column
      - Suggested survey options: shuffle row order, disable all confirmation page links
      - Example of a valid Google survey: 
      ![alt tag](https://raw.githubusercontent.com/jordanbonilla/OnlineVoting/master/example%20correct%20survey%20format.png)

  
4. A linked Google spreadsheet that gives explicit edit access to the local machine running this code
      - Select "share" on the spreadsheet and manually add your authorized console email.
      - Should look something like this: local-xxxxxxxxxxxx.iam.gserviceaccount.com


5. A new, blank worksheet in that spreadsheet with a unique name. 
      - Trying to re-use worksheets will result in failure
      - Google does not allow a complete wipe of existing worksheets
      
