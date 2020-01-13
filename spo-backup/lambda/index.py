''' 
  This function probes (for items) multiple sharepoint lists from different 
  sites which are provided in the sites_list.txt file. The items are 
  stored in CSV file and stored in S3 bucket
  
  Author: Oluwatise Olasebikan
'''
import json
import requests
import list_to_bucket
import boto3
import os
import adal
client = boto3.client('secretsmanager')
recipient = os.environ['recipient']
cc_recipient = os.environ['ccRecipient']

# get authenticated to perform actions on microsoft's apps
def sign_in():
  # response = client.get_parameter(Name="sharepoint_purge_and_backup", WithDecryption=True)
  # response = response['Parameter']['Value'].split(';')

  response = client.get_secret_value(SecretId='sharepoint_backup_secret')
  response = response['SecretString'].split(';')
  
  cert_key = response[0]
  client_id = response[1]
  client_secret = response[2]
  tenant_id = response[4]
  thumbprint = response[5]
  user_id = response[6]

  resource = "https://harborcapitaladvisors.sharepoint.com"
  url = "https://login.microsoftonline.com/"+ tenant_id + "/oauth2/v2.0/token"
  
  context = adal.AuthenticationContext('https://login.microsoftonline.com/'+tenant_id)
  token = context.acquire_token_with_client_certificate( "https://harborcapitaladvisors.sharepoint.com",
      client_id,  
      cert_key, 
      thumbprint)

  mail_token = context.acquire_token_with_client_certificate( "https://graph.microsoft.com",
      client_id,  
      cert_key, 
      thumbprint)

  return token, mail_token,resource,user_id


def send_email(token, message,user_id):  
  email_structure = {
    "message": {
      "subject": "SharePoint List Backup",
      "body": {
        "contentType": "Text",
        "content": message
      },
      "toRecipients": [
        {
          "emailAddress": {
            "address": recipient
          }
        }
      ],"ccRecipients": [
      {
        "emailAddress": {
          "address": cc_recipient
        }
      }
    ]
    },
    "saveToSentItems": "false"
  }
  request_url = "https://graph.microsoft.com/v1.0/users/"+user_id+"/sendMail"
  # print("\n\n URL IS " + str(request_url))
  headers = { 
  'User-Agent' : 'agent',
  'Authorization' : 'Bearer {0}'.format(token["accessToken"]),
  'Accept' : 'application/json',
  'Content-Type' : 'application/json'
  }
  
  send_req = requests.post(request_url, data = json.dumps(email_structure), headers = headers)
  if send_req:
    print("email sending is successful")
  else:
    print("email sending is NOT successful")

def handler(event, context):
  token, mail_token, resource,user_id = sign_in()
  path_sharepoint, is_all_success = list_to_bucket.do_job(token,resource)
  
  
  
  if is_all_success:
    message = "Successfully backed up the SharePoint List to CSV. View the CSV file(s) below: "
  else:
    message = "Sharepoint List to CSV backup was NOT successful. Please check Logs and view the CSV file(s): "
    
  for path in path_sharepoint:
      message = message + " \n"+path
  send_email(mail_token, message,user_id)
  return {
      "Result": message
  }