## Caltech Avery House Online Voting Project 

# Goals:
1. Create a free, secure, transparent, anonymized, convinient online voting system.
2. Remove the need for "middleman" survey services.
3. Avery House glory.

# Summary:
This framework interfaces Google surveys with Python for a clean front end and back end. The voting system uses the IRV (instnt runoff voting) strategy. The workflow can be summarized as:

1. Manually create a Google survey and link it to a Google spreadsheet. Enter the link to the survey into the Python script.

2. Python script generates unique links to the survey and emails out links. One link per email addresses. These links can be hardcoded or read from a Google doc (default)

3. Results automatically go into the linked Google Spreadsheet. As the script waits for the time limit to be reached, it actively checks the spreadsheet for tampering. 

4. When the voting period has ended, the Python script reads in the results from this spreadsheet, parses it, and calculates the winners based on instant runoff. The results are automatically emailed out.


# Security and anonymity:

1. The python script generates a unique 128-digit-time-seeded random int for each email address that is eligible to vote. This number is embedded in the URL of the Google survey link and is automatically added to the “voter ID” field in the survey. When votes are read from the Python script, it will check the “voter ID” field of submitted votes and confirm that they corresponding to originally-generated ID.This makes it so duplicate/unauthorized votes can be invalidated.

2. Every email also contains a randomly-generated 4 digit pin which represents the survey ID. This survey ID is sent in the initial email as well as the results email. Every voter should see the same survey ID but different voters see differnt survey IDs. The purpose of the survey ID is prevent a scenario where someone termiantes the script and sends out fake results emails.

3. The voter/survey IDs are not stored anywhere except in the python script’s internal state. This is possible because the script does not finish executing until the final email results are sent out. This prevents manipulation of IDs, ensures that vote results can not be modified, and maintains confidentiality of voters. Furthermore, the “sent” folder of the gmail account hosting the survey has its sent box deleted after the emails get sent - this prevents people with access to the email account from seeing which IDs were associated with specific email addresses. 

4. If someone manually adds a vote to the spreadsheet, it will certainly not have a valid 128-digit entry and will be invalidated in the final count.

5. While the script waits for quorum to be reached, it periodically compares a local copy of all voter data with the most recent batch of voter data. If the recent data is missing any of the old data, the script emails all voters that the spreadsheet has been tamepred with and terminates. If, by luck, the manipulation occurs before a refresh of the local data, a second layer of security is used - the election results will contain a list of all data from the vote, including the IDs associated with each vote. If someone believes their vote has been deleted, they can compare their voter ID (in the email) with the voter IDs found in the raw vote data which is automatically emailed at the end of a vote.A thrid layer of security is in the Google spreadsheet hosting the survey results which will contain all edit history of the spreadsheet. The edit history can be referenced if a votes legitimacy falls into question

6. Adding non-eligible emails to the eligible voter list is easily spotted since the email with the results also contains a list of all emails (xlsx format) that participated in the survey. All eligible voters are emailed the raw vote data (xlsx format) in the results email. 




# Prerequisitites for running locally:

1. Python 2.x and the following libraries:
      - gspread
      - oauth2client
      - xlsxwriter


2. OAuth2 credentials in json format.
      - http://gspread.readthedocs.org/en/latest/oauth2.html
      - This step only needs to be completed once


3. A linked Google spreadsheet that gives explicit edit access to the local machine running this code
      - Select "share" on the spreadsheet and manually add your authorized console email.
      - Should look something like this: local-xxxxxxxxxxxx.iam.gserviceaccount.com
      - This step only needs to be completed once


4. A manually-created Google Survey
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

  
5. A new, blank worksheet in your linked spreadsheet (step 3) with a unique name. 
      - This worksheet must be created manually and specified as the data destination when creating the survey (step 4)
      - Trying to re-use worksheets will result in failure
      - This step must be completed for every new survey
      ![alt tag](https://raw.githubusercontent.com/jordanbonilla/OnlineVoting/master/worksheet%20guide.png)

# Notes:
When choosing a local machine to run this script on, keep in mind that the script's execution can not be interrupted or the vote will end. Additionally, the script needs to internet access throughout the duration of its execution time so that it can actively check for vote manipulation. 

The script should should be portable and has been tested on Windows 10 and Ubuntu 14
