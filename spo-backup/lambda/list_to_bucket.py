import requests
import logging
import time
import csv
import json
import boto3
import random
import os
from datetime import datetime
import string
import re
from botocore.exceptions import ClientError
s3 = boto3.resource('s3')
region = os.environ['AWS_REGION']


is_all_success = True                      # will become false if any request fails
bucket_sharepoint = os.environ['bucketSharepoint']   # S3 Bucket to store the sharepoint CSV file in
date_regex = "[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]T..:..:..[A-Z]"
date_pattern = re.compile(date_regex)

'''GET the lists to probe, from the file "sites_lists.txt".
 NOTE: The file must end with a newline 
 NOTE: sites_list structure is {site1:[list1, list2], site2:[list1, list2]...}
'''
def get_lists_from_file(filename):
    sites_lists = {}
    file = open(filename, 'r')
    for line in file:
        slash_pos = line.rfind('/')
        site = line[:slash_pos]
        list = line[slash_pos+1:-1]
        if(site in sites_lists):
            sites_lists[site].append(list)
        else:
            array = []
            array.append(list)
            sites_lists[site] = array
    file.close()
    return sites_lists

''' Create the $select clause for lookup fields, by using the first item only '''
def get_lookup_clause(url, header):
  top_pos = url.find('$top=')
  one_item = requests.get(url=url[:top_pos]+'$top=1', headers=header)
  one_item = one_item.json()['value'][0]
  print(one_item)
  keyword = "&$select="
  expands = "&$expand="
  pos = url.find('/items')
#   url = url[:pos]+"/fields?$filter=Hidden eq false and ReadOnlyField eq false"
  url = url[:pos]+"/fields?$filter=Hidden eq false&$top=100"
  keywords_req = requests.get(url = url, headers=header)
#   print(keywords_req.text)
  field_objs = keywords_req.json()['value']
  all_fields = {}
  for field in field_objs:
    name = field['EntityPropertyName']
    if (name in one_item.keys() and one_item[name] == None) or name.startswith('OData__'):
        continue
    if 'LookupField' in field and name+"Id" in one_item.keys():
      keyword=keyword+name+"/Title,"+name+"Id,"
      expands = expands+name+","
      all_fields[field['Title']] = name
    elif not 'LookupField' in field and name in one_item.keys():
        keyword=keyword+name+','
        all_fields[field['Title']] = name
  # remove trailing commas
  keyword = keyword[:-1]
  expands = expands[:-1]
  
  keyword = keyword+expands
  print(keyword)
  return keyword, all_fields


def remove_innerJsons(arr):
  for obj in arr:
    for key,value in obj.items():
      if type(value) is dict:
        obj[key] = obj[key]['Title']
  return arr

''' Get items from all lists that are in sites_lists 
    (which is a dictionary of arrays/lists) 
    Returns a list that looks like:
    all_items_list ==> [ [{jsonObj1_fromList1}, {jsonObj2_fromList1}], [{jsonObj1_fromList2}, {jsonObj2_fromList2}]....] 
'''
def get_items(resource, header, sites_lists, site_type):
    resource = resource+"/sites/{0}/_api/web/lists/getbytitle('{1}')/items?$top=1000"
    all_items_list = []
    list_columns = []
    for site_name, lists_list in sites_lists.items():
        for list_name in lists_list:
            select_keyword, columns = get_lookup_clause(resource.format(site_name, lists_list[0]), header)
            list_columns.append(columns)
            url = resource.format(site_name, list_name)+select_keyword
            response = requests.get(url = url, headers = header)
            # print("url is==> "+url)
            # print("Response is===>> "+str(response))
            items_list = remove_innerJsons(response.json()['value'])
            all_items_list.append(items_list)
            list_columns.append(columns)
    # print("\n\n")
    # print(all_items_list)
    return all_items_list, list_columns

def generate_unique_filename(file_path):
    millis = int(round(time.time() * 1000))
    rand_str = ''.join(random.choice(string.ascii_lowercase) for i in range(4))  #2 character string
    path = file_path+ str(millis)+ '_'+ rand_str+".csv"
    return path

''' Convert date to form "Jan-6-20" '''
def to_readable_date(date):
    date = date[:10]
    words = datetime.strptime(date, "%Y-%m-%d").strftime("%b-%d %y")
    return words
    
''' items_list => [ [{jsonObj1_fromList1}, {jsonObj2_fromList1}], [{jsonObj1_fromList2}, {jsonObj2_fromList2}]....] '''
def store_in_csv(items_list, list_columns, file_path):
    file_paths = []
    count = 0
    for item in items_list:
        fieldnames = list(list_columns[count].values())
        fieldnames = ['Id' if field=='ID' else field for field in fieldnames]
        headers = list(list_columns[count].keys())
        path = generate_unique_filename(file_path)
        csv_file = open(path, 'w')
        csv_writer = csv.DictWriter(csv_file, fieldnames=(headers))
        csv_writer.writeheader()
        
        csv_writer = csv.DictWriter(csv_file, fieldnames=(fieldnames), restval="", extrasaction='ignore')
        for cell in item:
            # make each date-type data, more readable
            for key,val in cell.items():
                if date_pattern.match(str(val)):
                    cell[key] = to_readable_date(val)
            # write the row
            csv_writer.writerow(cell)
        csv_file.close
        file_paths.append(path)
        count = count+1
    return file_paths
    
    
''' Get the sharepoint items, store them in CSV file and
    store the file in S3 buckets   '''
def do_job(token, resource):
    global is_all_success
    sites_list= get_lists_from_file('sites_lists.txt')
    # rest_api_endpoint = resource+"/sites/"+site_name+"/_api/web/lists/getbytitle('"+ list_name+"')/items"
    
    header = { 
    'User-Agent' : 'agent',
    'Authorization' : 'Bearer {0}'.format(token["accessToken"]),
    'Accept' : 'application/json',
    'Content-Type' : 'application/json'
    }


    items_list, list_columns = get_items(resource, header, sites_list, 'list')
    #put in csv file and upload to S3 bucket
    paths_sharepoint = store_in_csv(items_list, list_columns, "/tmp/sharepoint_data_") 
    try:
        for path in paths_sharepoint:
            s3.meta.client.upload_file(path, bucket_sharepoint, path[5:])
      
    except ClientError as e:
        print("S3 bucket Error: "+str(e.response))
        is_all_success = False

    all_uploaded_paths = []
    s3_obj_path = "https://{}.s3.{}.amazonaws.com/{}"
    for s3FileName in paths_sharepoint:
        sharepoint_obj_path = s3_obj_path.format(bucket_sharepoint,region,s3FileName[5:])
        all_uploaded_paths.append(sharepoint_obj_path)
        
    return all_uploaded_paths, is_all_success