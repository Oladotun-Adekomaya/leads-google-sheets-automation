import os
import io
import pickle
from googleapiclient.http import MediaIoBaseDownload
from Google import Create_Service
from dotenv import load_dotenv # type: ignore

CLIENT_SECRET_FILE = 'drivecredentials.json'
API_NAME = 'drive'
API_VERSION = 'v3'
SCOPES = ['https://www.googleapis.com/auth/drive']

service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION,SCOPES)

load_dotenv()

with open ('sheet_list', 'rb') as fp:
    filelist = pickle.load(fp)

downloadPath = os.getenv("downloadPath")


def download_file(file_id, file_name):
    request = service.files().export_media(fileId=file_id, mimeType="text/csv")
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fd=fh, request=request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
    print(file_name)
    print("Download progress: {0}".format(int(status.progress() * 100))) 


    fh.seek(0)
    with open(os.path.join(downloadPath,file_name), 'wb') as f:
        f.write(fh.read())
        f.close()

# loop through 
for file in filelist:
    fileId = file["id"]
    #get file name
    fileget = service.files().get(fileId=fileId).execute()
    file_name = f'{fileget.get("name")}.csv'
    download_file(fileId,file_name)
    


