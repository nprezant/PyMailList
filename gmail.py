
# for the authenticator
from os import path, replace
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
from apiclient import errors

# for the message class
import base64
from email.mime.text import MIMEText

class Authenicator:
    '''
    Authenticates and manages a google web service
    '''

    def __init__(self, credential_path=None, token_path=None):
        self._SCOPES = ('https://www.googleapis.com/auth/gmail.send ' + 'https://www.googleapis.com/auth/gmail.readonly')
        self.credential_path = credential_path
        self.token_path = token_path
        self.service = None
        self.profile = None
        

    def start(self):
        '''
        gets the user credentials
        creates service to gmail account based on SCOPES
        also gets the profile data of the user
        '''
        creds = self._credentials()
        self.service = self._build(creds)
        self.profile = self._get_profile()


    def start_from_local(self):
        '''
        gets the users credentials to start
        the service from local storage if possible
        '''
        pass


    def restart(self):
        '''
        remove the authentication token
        to force a user login
        '''
        self.remove()
        self.start()


    def _credentials(self):
        '''
        get the credentials from storage if possible
        if not, get it from the login page
        '''
        store = file.Storage(self.token_path) # 'token.json'
        creds = store.get()
        if not creds or creds.invalid:
            flow = client.flow_from_clientsecrets(self.credential_path, self._SCOPES) # 'credentials.json'
            creds = tools.run_flow(flow, store)
        return creds

    
    def _build(self, creds):
        '''
        builds the gmail service
        '''
        return build('gmail', 'v1', http=creds.authorize(Http()))

    def _get_profile(self):
        '''
        gets the profile data from the user
        return is a dictionary of the profile data
        '''
        return self.service.users().getProfile(userId='me').execute()

    def remove(self):
        '''
        removes the token so a fresh log in will be forced
        (actually it just renames the old token, doesn't actually remove it)
        '''
        try:
            (root, ext) = path.splitext(self.token_path)
            replace(self.token_path, root+'-old'+ext)
        except FileNotFoundError:
            print(f'{self.token_path} file not found')


class Message:
    '''
    Holds an email message.
    Can also create the message

    Then you can change the object
    properties and recreate it
    (e.g. with a different send-to address)
    '''


    def __init__(self):
        self.subject:str = None
        self.body:str = None
        self.to:str = None # 'me'
        self.sender:str = None
        self.body_type:str = None # 'plain' or 'html'


    def create(self, to, sender, subject, body, body_type):
        '''
        creates a base64 encoded message object
        makes it based on the inputs
        '''
        self.subject = subject
        self.body = body
        self.to = to # 'me'
        self.sender = sender
        self.body_type = body_type # 'plain' or 'html'
        return self.recreate()

    def recreate(self):
        '''
        creates a message base64 encoded message object
        based on already existing object properties
        '''
        message = MIMEText(self.body, self.body_type)
        message['to'] = self.to
        message['from'] = self.sender
        message['subject'] = self.subject
        raw = base64.urlsafe_b64encode(message.as_bytes())
        raw = raw.decode()
        self.message = {'raw': raw}
        return self.message


class Emailer:
    '''
    Can send messages.
    Holds the authentication service.
    Also holds a single message
    '''

    def __init__(self, service):
        self.service = service
        self.message = Message()

    def send(self):
        '''
        safely sends the message that already
        exists within the object
        if there is an error, it will print it and return False, error
        if the message sends successfully, it will return True, sent_message
        '''
        try:
            sent_msg = (self.service.users().messages().send(
                userId=self.message.sender, 
                body=self.message.message
                ).execute())
            success = True, sent_msg
        except errors.HttpError as e:
            print('An error occurred: {}'.format(e))
            success = False, e
        finally:
            return success