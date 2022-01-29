#Import Dependencies
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.runtime.client_request import ClientRequest
from office365.runtime.http.request_options import RequestOptions
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
from office365.runtime.auth.client_credential import ClientCredential
from office365.runtime.auth.user_credential import UserCredential
import pandas as pd
import datetime

#Read in credentials for Sharepoint Authentication
with open('./Credentials.txt') as creds:
    lines = creds.read().splitlines()
creds.close

def download_files(row):
    #Read in columns URL and File Name
    file_url = row['ServerRelativeUrl']
    filenm = row['Name']
    #Concatentate to save to directory
    download_prefix = "C:/Temp/"
    download_path = "".join([download_prefix,filenm])
    #Let the magic happen
    with open(download_path, "wb") as local_file:
        file = ctx.web.get_file_by_server_relative_url(file_url).download(local_file).execute_query()
    
    return(print("[Ok] file has been downloaded into: {0}".format(download_path)))

tenant_url = "https://{company}.sharepoint.com/sites/TestTeam/"
folder_url = "/sites/TestTeam/Shared Documents/"

#Authenticate base tenant url
ctx_auth = AuthenticationContext(tenant_url)
user_credentials = UserCredential(lines[0],lines[1])
ctx = ClientContext(tenant_url).with_credentials(user_credentials)
web = ctx.web
ctx.load(web)
ctx.execute_query()
print("Web title: {0}".format(web.properties['Title']))

#Access Folder
libraryRoot = ctx.web.get_folder_by_server_relative_path(folder_url)
ctx.load(libraryRoot)
ctx.execute_query()

#Load files
files = libraryRoot.files
ctx.load(files)
ctx.execute_query()

#Declare data frame column names
df_files = pd.DataFrame(columns = ['Name', 'ServerRelativeUrl', 'TimeLastModified', 'ModTime'])
#Loop through files and pull name,url,mod time
for myfile in files:
    #use mod_time to get in better date format
    mod_time = datetime.datetime.strptime(myfile.properties['TimeLastModified'], '%Y-%m-%dT%H:%M:%SZ')  
    #create a dict of all of the info to add into dataframe and then append to dataframe
    dict = {'Name': myfile.properties['Name'], 'ServerRelativeUrl': myfile.properties['ServerRelativeUrl'], 'TimeLastModified': myfile.properties['TimeLastModified'], 'ModTime': mod_time}
    df_files = df_files.append(dict, ignore_index= True )

pd.set_option('display.max_colwidth', None)
df_files

#run apply function
df_files = df_files.apply(download_files, axis=1)
