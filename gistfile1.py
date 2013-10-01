__author__ = 'archeg'

import httplib
import urllib
import urllib2
import re


def URLRequest(url, params, headers, method="GET"):
    if method == "POST":
        return urllib2.Request(url, data=urllib.urlencode(params), headers=headers)
    else:
        return urllib2.Request(url + "?" + urllib.urlencode(params))

def setNote(row, column, value, tokens, spreadSheetId):
    """ Sets a note to the given cell
        Args:
            row: row coordinate
            column: column coordinate
            value: new value of the note (will be overriden)
            tokens: dictionary of auth tokens, returned by GetAuthTokens
            spreadSheetId: id of the spreadsheet. Be aware that this is not the key of the spreadsheet and cannot be got from url directly. Use GetSpreadsheetId function
    """
    cookies = "SID=%s; HSID=%s; SSID=%s;" % (tokens['SID'], tokens['HSID'], tokens['SSID'])
    print cookies
    headers = {
    "Content-Type":  "application/x-www-form-urlencoded;",
    "Cookie": cookies,
    "X-Same-Domain":	"trix"
    }
    data = {
                "action":	9,
                "atyp":	501,
                "ecol":	column,
                "erow":	row,
                "gid":	0, # TODO: Get Worksheet id
               "scol":	column,
                "srow":	row,
                "v":	value}
    url = "https://docs.google.com/a/ciklum.com/spreadsheet/edit/action9?id=" + spreadSheetId
    a = urllib2.Request(url, data=urllib.urlencode(data), headers=headers)

    res = urllib2.urlopen(a)
    b = res.read().decode()
    print b


CRLF = '\r\n'

# The following are used for authentication functions.
GAIA_HOST = 'www.google.com'
LOGIN_URI = '/accounts/ServiceLoginAuth'
LOGIN_URL = 'https://www.google.com/accounts/ClientLogin'
SERVICELOGIN_URL = 'https://www.google.com/accounts/ServiceLogin'
HOST = "accounts.google.com"
SERVICE = 'wise'

CLIENT_NAME = 'Example client to put notes on the cells'

# Example taken from https://developers.google.com/cloud-print/docs/pythonCode
def GetCookie(cookie_key, cookie_string):
    """Extract the cookie value from a set-cookie string.

    Args:
      cookie_key: string, cookie identifier.
      cookie_string: string, from a set-cookie command.
    Returns:
      string, value of cookie.
    """
    print('Getting cookie from %s' % cookie_string)
    id_string = cookie_key + '='
    cookie_crumbs = cookie_string.split(';')
    for c in cookie_crumbs:
      if id_string in c:
        cookie = c.split(id_string)
        return cookie[1]
    return None

# Example taken from https://developers.google.com/cloud-print/docs/pythonCode
def GaiaLogin(email, password):
    """Login to gaia using HTTP post to the gaia login page.

    Args:
      email: string,
      password: string
    Returns:
      dictionary of authentication tokens.
    """
    tokens = {}
    cookie_keys = ['SID', 'LSID', 'HSID', 'SSID']
    email = email.replace('+', '%2B')

    # Load sign in page to retrieve Galx
    opener = urllib2.build_opener(urllib2.HTTPCookieProcessor())
    urllib2.install_opener(opener)
    login_page_contents = opener.open(SERVICELOGIN_URL).read()
    galx_match_obj = re.search(r'name="GALX"\s*value="([^"]+)"', login_page_contents, re.IGNORECASE)
    galx_cookie = galx_match_obj.group(1) if galx_match_obj.group(1) is not None else ''

    # Forming a PUT request to login and retrieve auth data
    form = "GALX=%s&Email=%s&Passwd=%s" % (galx_cookie, email, password)

    login = httplib.HTTPS(GAIA_HOST, 443)
    login.putrequest('POST', LOGIN_URI)
    login.putheader('Host', HOST)
    login.putheader('content-type', 'application/x-www-form-urlencoded')
    login.putheader('content-length', str(len(form)))
    login.putheader('Cookie', 'GALX=%s' % galx_cookie)
    print('Sent POST content: %s' % form)
    login.endheaders()
    print('HTTP POST to https://%s%s' % (GAIA_HOST, LOGIN_URI))
    login.send(form)

    (errcode, errmsg, headers) = login.getreply()
    login_output = login.getfile()
    login_output.close()
    login.close()
    print('Login complete.')

    if errcode != 302:
      print('Gaia HTTP post returned %d, expected 302' % errcode)
      print('Message: %s' % errmsg)

    for line in str(headers).split('\r\n'):
      if not line: continue
      (name, content) = line.split(':', 1)
      if name.lower() == 'set-cookie':
        for k in cookie_keys:
          if content.strip().startswith(k):
            tokens[k] = GetCookie(k, content)

    if not tokens:
      print('No cookies received, check post parameters.')
      return None
    else:
      print('Received the following authorization tokens.')
      for t in tokens:
        print(t)
      return tokens

# Example taken from https://developers.google.com/cloud-print/docs/pythonCode
# Auth token is not needed for this particular script, but could be useful overall, so I kept it.
def GetAuthTokens(email, password):
    """Assign login credentials from GAIA accounts service.

    Args:
      email: Email address of the Google account to use.
      password: Cleartext password of the email account.
    Returns:
      dictionary containing Auth token.
    """
    # First get GAIA login credentials using our GaiaLogin method.
    tokens = GaiaLogin(email, password)

    # We still need to get the Auth token.
    params = {'accountType': 'GOOGLE',
              'Email': email,
              'Passwd': password,
              'service': SERVICE,
              'source': CLIENT_NAME}
    stream = urllib.urlopen(LOGIN_URL, urllib.urlencode(params))

    for line in stream:
      if line.strip().startswith('Auth='):
        tokens['Auth'] = line.strip().replace('Auth=', '')

    return tokens

# Definitely not the best way to get spreadsheet id, it should be possible to do that simpler.
def GetSpreadsheetId(key, tokens):
    """ Opens the spreadshet page and parses it's id from it's content
        Args:
            key: spreadsheet key, got from url
            tokens: dictionary of authentication tokens
        Returns:
            real id of the spreadsheet
    """

    print("Getting spreadsheet id")
    opener = urllib2.build_opener(urllib2.HTTPCookieProcessor())
    urllib2.install_opener(opener)

    cookies = "SID=%s; HSID=%s; SSID=%s;" % (tokens['SID'], tokens['HSID'], tokens['SSID'])
    opener.addheaders = [
        ("Cookie", cookies),
        ("user-agent", "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/27.0.1453.94 Safari/537.36")
    ]
    pageContents = opener.open("https://docs.google.com/spreadsheet/ccc?&key=" + key).read()

    match_obj = re.search(r'"id":"(\w*\.\w*\.\w*)"', pageContents)
    spreadSheetId = match_obj.group(1) if match_obj.group(1) is not None else ''

    return spreadSheetId

# Here it goes:
email = "anton.burnashev@gmail.com"
pwd = "burn25cerebrum"
column = 2
row = 10
noteValue = 'Hello notes!'
spreadSheetKey = "0AkgM6iO_6dprdE9COGJlZVdtT3R3cjU5VGVnUk1hR2c" # Can be got from spreadsheet URL by looking at &key= parameter
tokens = GetAuthTokens(email, pwd)
setNote(row, column, noteValue, tokens, GetSpreadsheetId(spreadSheetKey, tokens))
