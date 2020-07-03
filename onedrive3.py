import requests
import os.path
import os
import re
import pandas as pd
import json

# - - - - - - - - - - - - - - - - - - - - - - - - 

master_path = 'C:\\programs\\excel-onedrive\\excel\\'

base_url = 'https://graph.microsoft.com/v1.0/me/drive/items/'

url11 = '/workbook/tables(\'1\')'
url12 = '/workbook/tables(\'2\')'
url21 = '/workbook/worksheets/scan_info/tables/add'
url22 = '/workbook/worksheets/vuln_data/tables/add'
url31 = '/workbook/worksheets(\'scan_info\')/range(address=\'scan_info!A1:B1\')'
url32 = '/workbook/worksheets(\'vuln_data\')/range(address=\'vuln_data!A1:B1\')'
url41 = '/workbook/tables(\'1\')/rows'
url42 = '/workbook/tables(\'2\')/rows'
# - - - - - - - - - - - - - - - - - - - - - - - - 

token = 'eyJ0eXAiOiJKV1QiLCJub25jZSI6IlROQzRQNXZxcmdrRzFnVlBIOXQ5SUZvc3JuOGZOclFSS3BmbWdtcXR4a0kiLCJhbGciOiJSUzI1NiIsIng1dCI6IlNzWnNCTmhaY0YzUTlTNHRycFFCVEJ5TlJSSSIsImtpZCI6IlNzWnNCTmhaY0YzUTlTNHRycFFCVEJ5TlJSSSJ9.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTAwMDAtYzAwMC0wMDAwMDAwMDAwMDAiLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC84Y2M0MzRkNy05N2QwLTQ3ZDMtYjVjNS0xNGZlMGUzM2UzNGIvIiwiaWF0IjoxNTkzNzA5ODY1LCJuYmYiOjE1OTM3MDk4NjUsImV4cCI6MTU5MzcxMzc2NSwiYWNjdCI6MCwiYWNyIjoiMSIsImFpbyI6IkUyQmdZTmhnTUh1ZTdnTS95eWt1MmhuMVYvUVcyZTdNMFZ6elptVzB5TGFrQnptZlhlc0IiLCJhcHBfZGlzcGxheW5hbWUiOiJHcmFwaCBleHBsb3JlciAob2ZmaWNpYWwgc2l0ZSkiLCJhcHBpZCI6ImRlOGJjOGI1LWQ5ZjktNDhiMS1hOGFkLWI3NDhkYTcyNTA2NCIsImFwcGlkYWNyIjoiMCIsImZhbWlseV9uYW1lIjoiU2VuYW5heWFrZSIsImdpdmVuX25hbWUiOiJEdW1pZHUiLCJpcGFkZHIiOiIxNzUuMTU3LjE4OC4xMTIiLCJuYW1lIjoiU2VuYW5heWFrZSwgRHVtaWR1Iiwib2lkIjoiNTI4MmY0OWYtNjBkOC00OWU4LThmYzAtMmZmZDNiZjIwNjY1Iiwib25wcmVtX3NpZCI6IlMtMS01LTIxLTEwODUwMzEyMTQtMjAwMDQ3ODM1NC04Mzk1MjIxMTUtOTUxOTk0IiwicGxhdGYiOiIzIiwicHVpZCI6IjEwMDMyMDAwNTM1RjhCRjEiLCJzY3AiOiJDYWxlbmRhcnMuUmVhZFdyaXRlIENoYXQuUmVhZCBDb250YWN0cy5SZWFkV3JpdGUgRGV2aWNlTWFuYWdlbWVudEFwcHMuUmVhZC5BbGwgRGV2aWNlTWFuYWdlbWVudEFwcHMuUmVhZFdyaXRlLkFsbCBEZXZpY2VNYW5hZ2VtZW50Q29uZmlndXJhdGlvbi5SZWFkLkFsbCBEZXZpY2VNYW5hZ2VtZW50Q29uZmlndXJhdGlvbi5SZWFkV3JpdGUuQWxsIERldmljZU1hbmFnZW1lbnRNYW5hZ2VkRGV2aWNlcy5Qcml2aWxlZ2VkT3BlcmF0aW9ucy5BbGwgRGV2aWNlTWFuYWdlbWVudE1hbmFnZWREZXZpY2VzLlJlYWQuQWxsIERldmljZU1hbmFnZW1lbnRNYW5hZ2VkRGV2aWNlcy5SZWFkV3JpdGUuQWxsIERldmljZU1hbmFnZW1lbnRSQkFDLlJlYWQuQWxsIERldmljZU1hbmFnZW1lbnRSQkFDLlJlYWRXcml0ZS5BbGwgRGV2aWNlTWFuYWdlbWVudFNlcnZpY2VDb25maWcuUmVhZC5BbGwgRGV2aWNlTWFuYWdlbWVudFNlcnZpY2VDb25maWcuUmVhZFdyaXRlLkFsbCBEaXJlY3RvcnkuQWNjZXNzQXNVc2VyLkFsbCBEaXJlY3RvcnkuUmVhZC5BbGwgRGlyZWN0b3J5LlJlYWRXcml0ZS5BbGwgRXh0ZXJuYWxJdGVtLlJlYWQuQWxsIEZpbGVzLlJlYWQgRmlsZXMuUmVhZC5BbGwgRmlsZXMuUmVhZFdyaXRlLkFsbCBGaWxlcy5SZWFkV3JpdGUuQXBwRm9sZGVyIEdyb3VwLlJlYWQuQWxsIEdyb3VwLlJlYWRXcml0ZS5BbGwgSWRlbnRpdHlSaXNrRXZlbnQuUmVhZC5BbGwgTWFpbC5SZWFkV3JpdGUgTWFpbGJveFNldHRpbmdzLlJlYWRXcml0ZSBOb3Rlcy5SZWFkV3JpdGUuQWxsIG9wZW5pZCBQZW9wbGUuUmVhZCBwcm9maWxlIFJlcG9ydHMuUmVhZC5BbGwgU2l0ZXMuUmVhZFdyaXRlLkFsbCBUYXNrcy5SZWFkV3JpdGUgVXNlci5SZWFkIFVzZXIuUmVhZEJhc2ljLkFsbCBVc2VyLlJlYWRXcml0ZSBVc2VyLlJlYWRXcml0ZS5BbGwgZW1haWwiLCJzaWduaW5fc3RhdGUiOlsia21zaSJdLCJzdWIiOiI2cXI2bVk0RDJMT1FwUnZ5ZjB4WEU3c1BrSTZPTTl1MVV4WnRSd2QxVm84IiwidGVuYW50X3JlZ2lvbl9zY29wZSI6Ik5BIiwidGlkIjoiOGNjNDM0ZDctOTdkMC00N2QzLWI1YzUtMTRmZTBlMzNlMzRiIiwidW5pcXVlX25hbWUiOiJkdW1pZHUuc2VuYW5heWFrZUBwZWFyc29uLmNvbSIsInVwbiI6ImR1bWlkdS5zZW5hbmF5YWtlQHBlYXJzb24uY29tIiwidXRpIjoiN0thREVVT05Fa2EtRzVQTUpSVmRBQSIsInZlciI6IjEuMCIsInhtc19zdCI6eyJzdWIiOiJBbm5vMmlwYklrUmRqaHJCYzlqaG1TbXI0NlBxamlWSUdxMzMtVUdtOG80In0sInhtc190Y2R0IjoxMzQxNTE0MDkzfQ.UH0d_jSRBNHlnQeAVh4AViXoWjBsrmzHJZtwHBGPBuUK9AOnanDSCZKCnn8SjLJaRknOXwcDVb6nm3b5s_BpVtpbp5WO8nslDkuDeASx_pfbk3aLs2-CxmJBO9JH4hh3h6vTZxDBr1Cgg1detssdsE7Er72QX0GvDlxM4xXUuqrOMWHb__AFAn8a-mFZWR0KIvzW9CNzv9y5JjA6kQojx6q9Y-Ws4wIjwPlpORicHRUNpfV_XHmdd6iiJlyHNpBvnIqk5XftjuhNGG6LR8CGY5q7IRPEdot4CU483YMJbyyreLYijR0mkkkpoMyoo6H4CUQpeIUoOfGsWiQWRRePDw'

headers = {'Authorization': 'Bearer {}'.format(token)}
# - - - - - - - - - - - - - - - - - - - - - - - - 


f = open ("C:\\programs\\excel-onedrive\\excel\\master_file.csv", "r")     # open source excel folder and gogole sheet id list.
f1 = f.readlines()

for line in f1:
    line = line.rstrip()                                            # rstrip - remove the new line.
    a = line.split(",")                                             # split the line from "comma".
    excel_name,sheetid = a[0],a[1]
    
    
    latest_excel_file = os.path.join(master_path, '%s.xlsx' % (excel_name)) 
    print (latest_excel_file)
    
    # - - - - - - - - - - - - - - - - - - - - - - - -     
    
    df = pd.read_excel('%s' % latest_excel_file, sheet_name='Scan Information')
    #print (df)

    df1 = df.to_json(orient='values')
    #print (df1)

    df2 = '{"index":0,"values":'+df1+'}'
    print (df2)

    payload1 = json.loads(df2)
    #print (payload1)
    # - - - - - - - - - - - - - - - - - - - - - - - - 

    # delete existing table in the scan_info tab
    urlA = (base_url + sheetid + url11) 
    r1 = requests.delete(urlA, headers = headers).text
    print (r1)
    # print ("scan info delete table done")

    # create a table in the excel sheet in onedrive
    urlB = (base_url + sheetid + url21) 
    payload11 = {'address': 'A1:B2'}
    r2 = requests.post(urlB, headers = headers, json = payload11).text
    # print (r2)
    # print ("scan info create table done")

    # update the table headers
    urlC = (base_url + sheetid + url31) 
    payload12 = {"values":[["id","name"]]}
    print (payload12)
    r3 = requests.patch(urlC, headers = headers, json = payload12).text
    # print (r3)
    # print ("scan info update table headers")

    # update data in the table
    urlD = (base_url + sheetid + url41) 
    r4 = requests.post(urlD, headers = headers, json = payload1).text
    # print (r4)
    # print ("scan info update table data")
    # = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =



    df = pd.read_excel('%s' % latest_excel_file, sheet_name='Vulnerability Data')
    #print (df)

    df1 = df.to_json(orient='values')
    #print (df1)

    df2 = '{"index":0,"values":'+df1+'}'
    #print (df2)

    payload2 = json.loads(df2)
    #print (payload2)
    # - - - - - - - - - - - - - - - - - - - - - - - - 


    # delete existing table in the vuln_data tab
    urlE = (base_url + sheetid + url12) 
    r5 = requests.delete(urlE, headers = headers).text
    # print (r5)
    # print ("vuln data delete table done")

    # create a table in the excel sheet in onedrive
    urlF = (base_url + sheetid + url22) 
    payload21 = {'address': 'A1:B2'}
    r6 = requests.post(urlF, headers = headers, json = payload21).text
    # print (r6)
    #print ("vuln data create table done")

    # update the table headers
    urlG = (base_url + sheetid + url32) 
    payload22 = {"values":[["id","name"]]}
    r7 = requests.patch(urlG, headers = headers, json = payload22).text
    # print (r7)
    # print ("vuln data update table headers")

    # update data in the table
    urlH = (base_url + sheetid + url42) 
    r8 = requests.post(urlH, headers = headers, json = payload2).text
    # print (r8)
    # print ("vuln data update table data")
    # = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
    print (excel_name, "done")
