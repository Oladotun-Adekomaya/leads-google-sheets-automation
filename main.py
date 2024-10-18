import threading
import gspread
import pickle
from google.oauth2.service_account import Credentials
import os
from googleapiclient.http import MediaFileUpload
from Google import Create_Service
from dotenv import load_dotenv # type: ignore


# Google drive
CLIENT_SECRET_FILE = 'drivecredentials.json'
API_NAME = 'drive'
API_VERSION = 'v3'
SCOPES = ['https://www.googleapis.com/auth/drive']

# Google sheets
scopes = ["https://www.googleapis.com/auth/spreadsheets"]
creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
client = gspread.authorize(creds)

# 
service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION,SCOPES)




load_dotenv()


uploadPath = os.getenv("uploadPath")
csvfiles = os.listdir(uploadPath)
driveFolderId = os.getenv("driveFolderId")
counter = 1
sheetlist = []


def export_csv_file(filepath: str, parents: list=None):
   # print(f'File ID: {file.get("id")}')
   if not os.path.exists(filepath):
      print(f"{filepath} not found.")
      return
   try:
      filemetadata = {
         'name': os.path.basename(filepath).replace('.csv',''),
         'mimeType': 'application/vnd.google-apps.spreadsheet',
         'parents' : parents
      }
      
      media = MediaFileUpload(filename=filepath, mimetype='text/csv')
      
      response = service.files().create(
         media_body=media,
         body=filemetadata
      ).execute()
      fileid = response.get("id")
      filename = response.get("name")
      filedict = {
         "name": filename,
         "id":fileid
      }

      # print(f'File ID: {fileid}')
      sheetlist.append(filedict)
      # print(response)
      # return response
      return
   except Exception as e:
      print(e)
      return

def write_to_file(variable, filename):
   with open(filename, "a") as file:
      file.write(variable + "\n")


# Upload files
print("Uploading files...")
for csvfile in csvfiles:
    # export_csv_file(os.path.join(path, csvfile))
   export_csv_file(os.path.join(uploadPath, csvfile), parents=[driveFolderId])
   print(f"Uploading file {counter}")
   counter += 1

# Save id list in file, so can be used for downloading later
print("Saving id list in sheet_id_list file...\n")
with open('sheet_list', 'wb') as fp:
   pickle.dump(sheetlist, fp)




# SPREADSHEET

#get a hold of the spreadsheet
for sheet in sheetlist:
   sheet_id = sheet["id"]
   sheet_name = sheet["name"]
   sheet = client.open_by_key(sheet_id)
   details = f"Name: {sheet_name}\nid:{sheet_id}\n\n"

   #get worksheet
   print("Getting worksheet...")
   worksheet = sheet.sheet1

   #Determine the number of rows and column
   print("Determining number of rows in worksheet...")
   rows_no = int(len(worksheet.col_values(1)))
   print(rows_no)
   print("Determining number of rows in worksheet...")
   col_no = int(len(worksheet.row_values(1)))
   print(col_no)

   #If spreasheet is faulty i.e has less than 7 columns
   if col_no < 7:
      print("Terminating cleaning and adding file details to bad list...\n")
      # Write the name and id of the sheet to a text file and continue
      thread = threading.Thread(target=write_to_file, args=(details, "badlist.txt"))
      thread.start()
      continue

   #get saved contact
   print("Getting saved contact")
   savedcontact = worksheet.findall("TRUE")

   #get row number of first saved contact and last saved contact
   if savedcontact:
      startrow = savedcontact[0].row
      endrow = savedcontact[-1].row
      #delete rows of saved contact
      print("Deleting rows of saved contact...")
      worksheet.delete_rows(startrow,endrow)
   else:
      print("No saved contact")

   #fetch group name from worksheet 1
   print("Fetching group name from worksheet...")
   groupname = worksheet.acell('G2').value

   #Clear first column
   print("Clearing first column...")
   worksheet.batch_clear([f"A1:A{rows_no}"])


   #Delete useless columns
   print("Deleting useless rows...")
   worksheet.delete_columns(2,3)
   worksheet.delete_columns(3,5)

   #create list to store contact names
   namelist = [['Name']]

   #create contact name using group name
   print("Generating contact name using group name...")
   for n in range(1,rows_no):
      namelist.append([f"{n} {groupname}"])


   #update worksheet with contact names
   print("Updating generated contact names...")
   worksheet.update(namelist,f"A1:A{rows_no}")
   print("Contact name updated successfully.")

   #update header
   print("Updating worksheet header")
   sheetheader = ['Name','Phone']
   worksheet.update([sheetheader], 'A1:B1')
   print(f"{groupname} spreadsheet successfully cleaned.\n")