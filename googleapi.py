import os
import google.oauth2.credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google_auth_oauthlib.flow import InstalledAppFlow
from google_auth_oauthlib.tool import In
SCOPES = ['https://www.googleapis.com/auth/yt-analytics.readonly']

API_SERVICE_NAME = 'youtubeAnalytics'
API_VERSION = 'v3'
CLIENT_SECRETS_FILE = 'C:\\Users\\Swopil\\Downloads\\clientdetails.json'
def get_service():
  flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRETS_FILE, SCOPES)
  credentials = flow.run_console()
  #return build(API_SERVICE_NAME, API_VERSION, credentials = credentials)
  return build(API_SERVICE_NAME, API_VERSION, developerKey='AIzaSyCxI0_D5lQ1UoAKTOjzb7MMvGgW3z14_is')

def execute_api_request(client_library_function, **kwargs):
  response = client_library_function(
    **kwargs
  ).execute()

  print(response)

if __name__ == '__main__':
  # Disable OAuthlib's HTTPs verification when running locally.
  # *DO NOT* leave this option enabled when running in production.
  os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '0'

  youtubeAnalytics = get_service()
  execute_api_request(
      youtubeAnalytics.reports().query,
      ids='channel==MINE',
      startDate='2017-01-01',
      endDate='2017-12-31',
      metrics='estimatedMinutesWatched,views,likes,subscribersGained',
      dimensions='day',
      sort='day'
  )