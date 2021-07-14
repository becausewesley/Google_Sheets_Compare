## Google Sheets Excel Compare
>This script was originally designed to check a list of contacts, against
> a list of "opt-outs" (people who no longer wanted to be on the list). 

The script works by looping over two separate spreadsheets:
- A Google sheet
- A local Excel sheet

First it loops over the Google sheet, and for each row (contact), it loops
through the local file to check if it is contained there. If it is, then 
that contact has 'opted-out' and so we need to remove them from the list. 

_*NB: This script does not make any deletions on the GSheet, it only
flags contacts which match the criteria logic for manual/semi-manual 
deletion._


## Adding Google Service Account:
To use this script you will need to create a Google Service account from
the [Google Dev Console](console.developers.google.com/).<br />
You will also have to generate a JSON file containing your credentials
in order for the script to have access.<br /><br />

You can read more about how to get this set up [HERE](https://stackoverflow.com/questions/46287267/how-can-i-get-the-file-service-account-json-for-google-translate-api) 
and [HERE](https://help.talend.com/r/E3i03eb7IpvsigwC58fxQg/ol2OwTHmFbDiMjQl3ES5QA).
<br /><br />
**NB***: **You need to give the service account access to your Google Doc/Sheet.
You do this by getting the Service Account email address contained in the JSON
file, and then adding it as a Editor from the Google Sheet itself.**

## Third Party Libraries
- This script uses the [tqdm](https://tqdm.github.io/) module. This is a progress bar module
and is non-essential to the functioning of the script. I added this because the 
script takes a while to execute because of all the looping - this will improve
when I replace the local excel with a SQLite DB.<br />
  `pip install tqdm` <br /><br />
  
- This script also uses the [GSpread](https://docs.gspread.org/en/latest/)
module for API access to the Google Sheet. <br />
  `pip install gspread` <br /><br />
  
### # TODO: 
There are a couple of TODOs in this project, and I will update the 'Issues'
tab to reflect these, there is however one major upgrade which I will be handling soon
and that is:
- Replacing the local excel file, with a SQLite DB. This will make the process
of querying 'opt-outs' much quicker, without having to loop through the
  entire list each time.

