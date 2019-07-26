# ----------------------------------------------------------------------------#
# (C) British Crown Copyright 2019 Met Office.                                #
# Author: Steve Wardle                                                        #
#                                                                             #
# This file is part of OWA Checker.                                           #
# OWA Checker is free software: you can redistribute it and/or modify it      #
# under the terms of the Modified BSD License, as published by the            #
# Open Source Initiative.                                                     #
#                                                                             #
# OWA Checker is distributed in the hope that it will be useful,              #
# but WITHOUT ANY WARRANTY; without even the implied warranty of              #
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the               #
# Modified BSD License for more details.                                      #
#                                                                             #
# You should have received a copy of the Modified BSD License                 #
# along with OWA Checker...                                                   #
# If not, see <http://opensource.org/licenses/BSD-3-Clause>                   #
# ----------------------------------------------------------------------------#
import os
import urllib
import time
import requests
import o365_api
import re
from BaseHTTPServer import HTTPServer, BaseHTTPRequestHandler

# Client ID and secret
CLIENT_ID = os.environ.get("OWA_CHECKER_CLIENT_ID", None)
CLIENT_SECRET = os.environ.get("OWA_CHECKER_CLIENT_SECRET", None)

if CLIENT_ID is None or CLIENT_SECRET is None:
    msg = ("Microsoft Azure Client ID and Client Secret/Password "
           "not found (please specify via environment variables "
           "OWA_CHECKER_CLIENT_ID and OWA_CHECKER_CLIENT_SECRET "
           "before starting OWA checker)")
    raise ValueError(msg)

# The OAuth authority
AUTHORITY = 'https://login.microsoftonline.com'

# The authorize URL that initiates the OAuth2 client credential
# flow for admin consent
AUTHORIZE_URL = '{0}{1}'.format(AUTHORITY, '/common/oauth2/v2.0/authorize?{0}')

# The token issuing endpoint
TOKEN_URL = '{0}{1}'.format(AUTHORITY, '/common/oauth2/v2.0/token')

# The scopes required by the app
SCOPES = ['Calendars.Read',
          'Mail.Read',
          'openid',
          'offline_access',
          'User.Read',
          'User.ReadBasic.All',
          ]

# Persistent dictionary that stores the tokens for the current session
OWACHECKER_SESSION = {'access_token': '',
                      'refresh_token': '',
                      'expires_at': '',
                      }

# Details for the local server used to provide the authentication
REDIRECT_PORT = 1234
REDIRECT_URI = "http://localhost:{0}".format(REDIRECT_PORT)

# Token caching
TOKEN_CACHE_PATH = os.path.join(os.environ['HOME'], ".owa_check")

# Token expiry threshold in minutes - this will be subtracted from the
# actual expiry time to avoid letting it get too close to expiring
TOKEN_EXPIRY_THRESHOLD = 300


def get_signin_url():
    """
    Given a redirection location construct the appropriate signin URL which
    the user should be directed to in order to allow access to the app

    """
    params = {'tenant': 'common',
              'client_id': CLIENT_ID,
              'redirect_uri': REDIRECT_URI,
              'response_type': 'code',
              'scope': ' '.join(str(i) for i in SCOPES)
              }

    signin_url = AUTHORIZE_URL.format(urllib.urlencode(params))

    return signin_url


def get_token_from_code(auth_code):
    """
    Given the code which should have been returned via the initial
    authentication process, requests an access token for the API

    """
    post_data = {'grant_type': 'authorization_code',
                 'code': auth_code,
                 'redirect_uri': REDIRECT_URI,
                 'scope': ' '.join(str(i) for i in SCOPES),
                 'client_id': CLIENT_ID,
                 'client_secret': CLIENT_SECRET
                 }

    r = requests.post(TOKEN_URL, data=post_data)

    if (r.status_code == requests.codes.ok):
        tokens = r.json()
        # Save tokens to the session dictionary
        OWACHECKER_SESSION['access_token'] = tokens['access_token']
        OWACHECKER_SESSION['refresh_token'] = tokens['refresh_token']
        # Calculate expiry time in absolute terms (use epoch... strip
        # off some time to avoid letting it get too close to the wire)
        expiration = (
            int(time.time()) + tokens['expires_in'] - TOKEN_EXPIRY_THRESHOLD)
        OWACHECKER_SESSION['expires_at'] = expiration
        # Cache the new token
        cache_refresh_token()
    else:
        raise ValueError("Could not retrieve token from code: "
                         + r.text)


def get_token_from_refresh_token():
    """
    Given a special refresh token (which may be requested as part of the
    initial authentication), request an access token for the API

    """
    refresh_token = OWACHECKER_SESSION['refresh_token']

    post_data = {'grant_type': 'refresh_token',
                 'refresh_token': refresh_token,
                 'redirect_uri': REDIRECT_URI,
                 'scope': ' '.join(str(i) for i in SCOPES),
                 'client_id': CLIENT_ID,
                 'client_secret': CLIENT_SECRET
                 }

    r = requests.post(TOKEN_URL, data=post_data)

    if (r.status_code == requests.codes.ok):
        tokens = r.json()
        # Save tokens to the session dictionary
        OWACHECKER_SESSION['access_token'] = tokens['access_token']
        OWACHECKER_SESSION['refresh_token'] = tokens['refresh_token']
        # Calculate expiry time in absolute terms (use epoch... strip
        # off some time to avoid letting it get too close to the wire)
        expiration = (
            int(time.time()) + tokens['expires_in'] - TOKEN_EXPIRY_THRESHOLD)
        OWACHECKER_SESSION['expires_at'] = expiration
        # Cache the new token
        cache_refresh_token()
    else:
        raise ValueError("Could not retrieve token from refresh token: "
                         + r.text)


def get_access_token():
    """
    Returns an active access token for use in calls to the API; if the
    current token has expired a new one will be requested using the refresh
    token

    """
    current_token = OWACHECKER_SESSION['access_token']
    expiration = OWACHECKER_SESSION['expires_at']
    now = int(time.time())
    if (current_token == "" or now > expiration):
        # Token expired
        get_token_from_refresh_token()
        return OWACHECKER_SESSION['access_token']
    else:
        # Token still valid
        return current_token


def cache_refresh_token():
    """
    Writes the refresh token to a cache in the user's home directory

    """
    # If the path doesn't already exist, create it, otherwise
    # just ensure it is private
    if not os.path.exists(TOKEN_CACHE_PATH):
        os.mkdir(TOKEN_CACHE_PATH, 0o700)
    else:
        os.chmod(TOKEN_CACHE_PATH, 0o700)

    cache_file = os.path.join(TOKEN_CACHE_PATH, "refresh")
    with open(cache_file, "w") as cfile:
        cfile.write(OWACHECKER_SESSION['refresh_token'])


def load_refresh_token():
    """
    Reads the refresh token from the cache in the user's home directory

    """
    # Check if the file exists and if it does read in the token
    cache_file = os.path.join(TOKEN_CACHE_PATH, "refresh")
    if os.path.exists(cache_file):
        with open(cache_file, "r") as cfile:
            OWACHECKER_SESSION['refresh_token'] = cfile.read()


class LoginHTTPHandler(BaseHTTPRequestHandler):
    def _display(self):
        self.send_response(200)
        self.send_header('Content-type', 'text/html')
        self.end_headers()

    def _redirect(self, new_url):
        self.send_response(301)
        self.send_header('Location', new_url)
        self.end_headers()

    def do_GET(self):
        search = re.search(r"code=(.*)&", self.path)
        if search:
            # Check for arriving as a redirect from Microsoft (which
            # will be accompanied by the auth code)
            get_token_from_code(search.group(1))

            # Test that the returned token works by retreiving the user's
            # email address from their contact details
            user = o365_api.get_user_info(get_access_token())

            if type(user) is not dict:
                self._display()
                self.wfile.write(
                    "<html><body><h1>Error Signing in!</h1></body></html>")
                return

            self._display()
            self.wfile.write("""
                 <html><body><h1>Signed in as: {0}</h1>
                 <p>You can now close this tab/browser</p>
                 </body></html>"""
                             .format(user['mail']))
            self.server.auth_complete = True
        else:
            # If we haven't arrived here with the code, redirect to
            # the signin page
            auth_url = get_signin_url()
            self._redirect(auth_url)

    # Suppress printing of stdout; comment this out for debugging
    def log_message(self, *args):
        return


def run(server_class=HTTPServer, handler_class=LoginHTTPHandler):
    server_address = ('', REDIRECT_PORT)
    httpd = server_class(server_address, handler_class)

    while not hasattr(httpd, "auth_complete"):
        httpd.handle_request()


if __name__ == "__main__":
    print "Starting Server... please navigate to: " + REDIRECT_URI
    run()
    if OWACHECKER_SESSION["access_token"] is not "":
        print "Successfully signed in!"
        print "You can now re-run without the --auth argument"
        print "to launch the checker. Note that if your desktop"
        print "password changes, or your token expires for some"
        print "other reason you will need to re-run the"
        print "authentication process"
    else:
        print "Failed to sign in..."
