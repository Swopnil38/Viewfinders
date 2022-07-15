import csv
from email import header
import os
import re
from venv import create
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from numpy import average
import pandas as pd
import xlsxwriter

from tabulate import tabulate
from flask import Flask

app = Flask(__name__)

SCOPES = ['https://www.googleapis.com/auth/yt-analytics.readonly']

API_SERVICE_NAME = 'youtubeAnalytics'
API_VERSION = 'v2'
CLIENT_SECRETS_FILE = 'C:\\Users\\Swopil\\Downloads\\clientdetails4.json'

def get_service():
  flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRETS_FILE, SCOPES)
  credentials = flow.run_local_server()
  return build(API_SERVICE_NAME, API_VERSION, credentials = credentials)

def execute_api_request(client_library_function, **kwargs):
  response = client_library_function(
    **kwargs
  ).execute()
  return response

def create_table(table, headers=None):
    if headers:
        headerstring = "\t{}\t" * len(headers)
        print(headerstring.format(*headers))

    rowstring = "\t{}\t" * len(table[0])

    for row in table:
        print(rowstring.format(*row))


@app.route('/')
def execute():

    youtubeAnalytics = get_service()

    result1 = execute_api_request(
        youtubeAnalytics.reports().query,
        ids='channel==MINE',
        startDate='2019-01-01',
        endDate='2022-07-15',
        metrics='averageViewDuration,views',
        dimensions='day',
        sort='day'
    )
    result2 = execute_api_request(
        youtubeAnalytics.reports().query,
        ids='channel==MINE',
        startDate='2019-01-01',
        endDate='2022-07-15',
        dimensions='ageGroup,gender',
        metrics='viewerPercentage',
        sort='gender,ageGroup'
    )
    result3 = execute_api_request(
        youtubeAnalytics.reports().query,
        ids='channel==MINE',
        startDate='2019-01-01',
        endDate='2022-07-15fla',
        dimensions='country',
        metrics='views,estimatedMinutesWatched,averageViewDuration,averageViewPercentage',
        sort='country'
    )
    #print(result)
    headers2 = ['date', 'avg.viewDuration', 'views']
    headers3 = ['ageGroup','Gender','viewerPercentage']
    headers4 = ['Country','views','est.Minute','avg.ViewDurstion','avg.viewPer']
    

    workbook = xlsxwriter.Workbook('YoutubeData.xlsx')
    worksheet1 = workbook.add_worksheet("GenderWiseAge")
    worksheet2 = workbook.add_worksheet("Geography")
    worksheet3 = workbook.add_worksheet("view")
    
    worksheet1.write_row(0,0,headers3)
    for row_num, row_data in enumerate(result2['rows']):
        print("1")
        for col_num, col_data in enumerate(row_data):
            worksheet1.write(row_num+1, col_num, col_data)

    
    worksheet2.write_row(0,0,headers4)
    for row_num, row_data in enumerate(result3['rows']):
        for col_num, col_data in enumerate(row_data):
            worksheet2.write(row_num+1, col_num, col_data)

    worksheet3.write_row(0,0,headers2)
    for row_num, row_data in enumerate(result1['rows']):
        for col_num, col_data in enumerate(row_data):
            worksheet3.write(row_num+1, col_num, col_data)
            
    workbook.close()
      
