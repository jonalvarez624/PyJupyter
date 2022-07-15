#Import Dependencies
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.runtime.client_request import ClientRequest
from office365.runtime.http.request_options import RequestOptions
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
from office365.runtime.auth.client_credential import ClientCredential
from office365.runtime.auth.user_credential import UserCredential
from simple_salesforce import Salesforce
import requests
import pandas as pd
from io import StringIO
import os
import pyodbc
from sqlalchemy import create_engine
import time

try:
    #read in db credentials
    with open('//oneabbott.com/ADFS9/DEPT/DEPT/Dcbc_215/eCommerce/Databases/Source Files/ANPST_PASS.txt') as creds:
        anpst = creds.read().splitlines()
    creds.close

    #Connect to ANPST
    conn = pyodbc.connect('Driver={ODBC Driver 17 for SQL Server};Server=WQ00391p;Database=ANPST;UID=ANPST_USER;PWD='+anpst[0])
    cursor = conn.cursor()

    #Create SQL engine - Need to Use ODBC Driver 17 for SQL Server in order to improve speed
    engine = create_engine("mssql+pyodbc://{user}:{pw}@WQ00391p/{db}?driver=ODBC+Driver+17+for+SQL+Server"
                           .format(user="ANPST_USER",
                                   pw=anpst[0],
                                   db="ANPST"),
                           #This speeds up insert and updates statements 10 fold
                           fast_executemany=True)

    #read in salesforce credentials from central location
    with open('//oneabbott.com/ADFS9/DEPT/DEPT/Dcbc_215/eCommerce/Databases/Source Files/SALESFORCE_PASS.txt') as creds:
        lines = creds.read().splitlines()
    creds.close

    #authenticate and connect to salesforce
    sf = Salesforce(username=lines[0],password=lines[1], security_token=lines[2])

    #Download Account Territory csv from report
    sf_instance = 'https://anretail-prod.lightning.force.com/' #AN Salesforce Instance URL
    reportId = '00O4X000009k6zRUAQ' # add report id
    export = '?isdtp=p1&export=1&enc=UTF-8&xf=csv'
    sfUrl = sf_instance + reportId + export
    response = requests.get(sfUrl, headers=sf.headers, cookies={'sid': sf.session_id})
    download_report = response.content.decode('utf-8')
    df1 = pd.read_csv(StringIO(download_report))
    df1.columns = df1.columns.str.replace(' ', '_')
    df1 = ( df1.drop(columns=['CBC'])
           .rename(columns={'Parent_Account:_Account_Name':'Parent_Account_Name'})
           )

    #Download Internal Users csv from report
    sf_instance = 'https://anretail-prod.lightning.force.com/' #AN Salesforce Instance URL
    reportId = '00O4X000009kCFjUAM' # add report id
    export = '?isdtp=p1&export=1&enc=UTF-8&xf=csv'
    sfUrl = sf_instance + reportId + export
    response = requests.get(sfUrl, headers=sf.headers, cookies={'sid': sf.session_id})
    download_report = response.content.decode('utf-8')
    df2 = pd.read_csv(StringIO(download_report))

    # Clean up the Internal Users report to match Account Territory data
    df2.columns = df2.columns.str.replace(' ', '_')
    df2 = (df2.drop(columns=['First_Name', 'Last_Name', 'User_ID', 'Department'])
           .assign(CBC_Territory_Number= "9999999",HQ_Territory= "9999999")
           .assign(Account_Name = "Internal",Parent_Account_Name = "Internal",HQ_Regional_Sales_Director= "Jason Smith")
           .rename(columns={"Title":"Role_in_Territory"}, errors="raise")
           )

    #Union dataframes and filter out external partners aka advantagesolutions emails
    finaldf = pd.concat([df2, df1])
    finaldf = finaldf[~finaldf['Email'].str.contains("advantagesolution")]
    print('Total Salesforce Users Row Count is:', len(finaldf.index))

    #Upload Salesforce Territory assignments to ANPST
    finaldf.to_sql("RLS_CBC_Account_Territory", engine, index=False,if_exists="replace",schema="dbo")

    #Check if new users need to be added to ANPST RPT tables
    dfp = pd.read_sql("SELECT EmailAddress FROM RPT_Person", engine)
    dfp = dfp["EmailAddress"].str.lower()  #merge is case sensitive in pandas
    df_merged = pd.merge(dfp,finaldf, how='outer', left_on='EmailAddress', right_on='Email',indicator=True)
    df_new_users = df_merged[(df_merged._merge == "right_only")]
    df_new_users = df_new_users[['Full_Name','Email']].drop_duplicates()
    df_new_users = (df_new_users.rename(columns={"Full_Name":"PersonName","Email":"EmailAddress"}, errors="raise")
                                .assign(Active = 1)
                    )

    print('New Email Address Users Row Count is:', len(df_new_users.index))

    #Add new users to RPT_Person table in ANPST
    if len(df_new_users.index) > 0:
        print('Uploading',len(df_new_users.index),'new records to RPT_Person')
        df_new_users.to_sql("RPT_Person", engine, index=False,if_exists="append",schema="dbo")

    #After updating the RPT_Person table build query to update RPT_Group_Membership Table
    query = open('//oneabbott.com/ADFS9/DEPT/DEPT/Dcbc_215/eCommerce/Databases/Source Files/SALESFORCE_UPDATE_GROUP_MEMBERSHIP.sql', 'r')
    df_membership = pd.read_sql_query(query.read(), engine)

    print('New Group Membership Row Count is:', len(df_membership.index))

    #Add new memberships to RPT_Group_Membership table in ANPST
    if len(df_membership.index) > 0:
        print('Uploading',len(df_membership.index),'new records to RPT_Group_Membership')
        df_membership.to_sql("RPT_Group_Membership", engine, index=False,if_exists="append",schema="dbo")

    # Read in credentials for Sharepoint Authentication
    with open('//oneabbott.com/ADFS9/DEPT/DEPT/Dcbc_215/eCommerce/Databases/Source Files/JA_CREDS.txt') as creds:
        lines = creds.read().splitlines()
    creds.close

    # Authenticate for new sharepoint site to read in manual list
    tenant_url = "https://abbott.sharepoint.com/sites/US-ANPD-RGM/"
    file_url = "/sites/US-ANPD-RGM/Shared Documents/General/Reporting/Row_Level_Security_Manual_Userlist.xlsx"
    filenm = "Row_Level_Security_Manual_Userlist.xlsx"

    # Authenticate base tenant url
    ctx_auth = AuthenticationContext(tenant_url)
    user_credentials = UserCredential(lines[0], lines[1])
    ctx = ClientContext(tenant_url).with_credentials(user_credentials)
    web = ctx.web
    ctx.load(web)
    ctx.execute_query()
    print("Web title: {0}".format(web.properties['Title']))

    # Concatentate to save to directory
    download_prefix = "//oneabbott.com/ADFS9/DEPT/Data/AN_Retail_Advanced_Analytics/RGM/RLS/"
    download_path = "".join([download_prefix, filenm])
    # Let the magic happen
    with open(download_path, "wb") as local_file:
        file = ctx.web.get_file_by_server_relative_url(file_url).download(local_file).execute_query()

    print("[Ok] file has been downloaded into: {0}".format(download_path))

    #read in excel to df
    df_rls = pd.read_excel(download_path)
    df_rls["Email"] = df_rls["Email"].str.lower()
    #Check for new users from manual list that need to be added to RLS table
    db_rls = pd.read_sql("SELECT DISTINCT Email FROM RLS_CBC_Account_Territory", engine)
    rls_merged = pd.merge(df_rls, db_rls, how='outer', left_on='Email', right_on='Email', indicator=True)
    rls_new_users = rls_merged[(rls_merged._merge == "left_only")]
    rls_new_users = rls_new_users.drop('_merge', 1)
    rls_new_users = rls_new_users.drop('Unnamed: 10', 1)

    print('New RLS Users Row Count is:', len(rls_new_users.index))

    # Add new users to RLS table in ANPST
    if len(rls_new_users.index) > 0:
        print('Uploading', len(rls_new_users.index), 'new records to RLS_CBC_Account_Territory')
        rls_new_users.to_sql("RLS_CBC_Account_Territory", engine, index=False, if_exists="append", schema="dbo")

    print('Script Complete:', time.strftime("%m-%d-%Y %I:%M %p"))

except Exception as e: print(e)

finally:
    #Close all connections
    response.close()
    engine.dispose()
    cursor.close()