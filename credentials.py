##############
# credentials.py
# handles gspread credentials
##############

# imports
import gspread
import json
from oauth2client.client import SignedJwtAssertionCredentials \
    as JwtCredentials

# main class
class Credentials(object):

    ## constructor
    ### needs a path to a keyfile
    ### (Google service account JSON file)
    def __init__(self, keyfile):
        jsonKey = json.load(open(keyfile))
        scope = ['https://spreadsheets.google.com/feeds',
                 'https://www.googleapis.com/auth/drive'] 
                # allows creation of new sheets
        credentials = JwtCredentials(jsonKey['client_email'],
                          jsonKey['private_key'].encode(),
                          scope)
        self.gc = gpsread.authorize(credentials)

    ## getClient
    ### retrieves gspread Client
    ### ensures client isn't expired
    def getClient(self):
        if self.gc.auth.access_token_expired:
            self.gc.login()
        return self.gc