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
import requests
import uuid
import json
from datetime import datetime, timedelta

GRAPH_ENDPOINT = 'https://graph.microsoft.com/v1.0{0}'


class o365Error(ValueError):
    """
    Raised when the API fails to return anything (could be for a
    number of reasons; the response will be attached)

    """
    pass


# Generic API Sending
def make_api_call(method, url, token, user_email,
                  payload=None, parameters=None, timeout=None):
    # Send these headers with all API calls
    headers = {'User-Agent': 'OWA Checker',
               'Authorization': 'Bearer {0}'.format(token),
               'Accept': 'application/json',
               'X-AnchorMailbox': user_email}

    # Use these headers to instrument calls. Makes it easier
    # to correlate requests and responses in case of problems
    # and is a recommended best practice.
    request_id = str(uuid.uuid4())
    instrumentation = {'client-request-id': request_id,
                       'return-client-request-id': 'true'}

    headers.update(instrumentation)

    response = None

    if (method.upper() == 'GET'):
        response = requests.get(url,
                                headers=headers,
                                params=parameters,
                                timeout=timeout)
    elif (method.upper() == 'DELETE'):
        response = requests.delete(url,
                                   headers=headers,
                                   params=parameters)
    elif (method.upper() == 'PATCH'):
        headers.update({'Content-Type': 'application/json'})
        response = requests.patch(url,
                                  headers=headers,
                                  data=json.dumps(payload),
                                  params=parameters)
    elif (method.upper() == 'POST'):
        headers.update({'Content-Type': 'application/json'})
        response = requests.post(url,
                                 headers=headers,
                                 data=json.dumps(payload),
                                 params=parameters)
    return response


def get_user_info(access_token):
    """
    Returns the logged-in user's username and email address

    """
    get_me_url = GRAPH_ENDPOINT.format('/me')

    # Use OData query parameters to control the results
    #  - Only return the displayName and mail fields
    query_parameters = {'$select': 'displayName,mail'}

    try:
        r = make_api_call('GET', get_me_url, access_token,
                          "", parameters=query_parameters)
    except Exception as err:
        raise o365Error(err)

    if (r.status_code == requests.codes.ok):
        return r.json()
    else:
        raise o365Error("{0}: {1}".format(r.status_code, r.text))


def get_new_messages(access_token, user_email, last_seen=None, folders=None):
    """
    Returns unread messages in the logged-in user's Inbox; note that
    the 365 API limits the maximum amount of results that are returned
    to 10 by default.

    If provided with a "last_seen" time (a string in the same format as
    returned in the JSON) the call will only return unread messages
    which arrived *more recently* than that time.  Otherwise it will
    return any unread messages (up to 10, see above!)

    """
    if folders is None:
        folders = ["inbox"]
    else:
        folders = [folder.lower() for folder in folders]

    if last_seen is None:
        # If no last-seen time was provided, get all unread messages
        query_parameters = {
            '$filter': 'isRead eq false',
            '$select': 'receivedDateTime,subject,from',
            '$orderby': 'receivedDateTime DESC'}
    else:
        # Otherwise retrieve only those since the last-seen message
        query_parameters = {
            '$filter': ('isRead eq false and receivedDateTime gt {0}'
                        .format(last_seen)),
            '$select': 'receivedDateTime,subject,from',
            '$orderby': 'receivedDateTime DESC'}

    folder_ids = []
    get_folders_url = GRAPH_ENDPOINT.format('/me/mailfolders')
    next_page = True
    while next_page:
        # Get the folder ids
        try:
            r = make_api_call('GET', get_folders_url, access_token,
                              user_email)
        except Exception as err:
            raise o365Error(err)

        if (r.status_code == requests.codes.ok):
            output = r.json()
            for folder in output["value"]:
                if folder["displayName"].lower() in folders:
                    folder_ids.append(folder["id"])

            if len(folder_ids) == len(folders):
                break

            if '@odata.nextLink' in output:
                get_folders_url = output['@odata.nextLink']
            else:
                next_page = False

    messages = []
    for folder_id in folder_ids:
        # Make the call and return the result
        get_messages_url = GRAPH_ENDPOINT.format(
            "/me/mailfolders/{0}/messages".format(folder_id))
        try:
            r = make_api_call('GET', get_messages_url, access_token,
                              user_email, parameters=query_parameters)
        except Exception as err:
            raise o365Error(err)

        if (r.status_code == requests.codes.ok):
            output = r.json()
            if output is not None:
                messages += output["value"]
        else:
            raise o365Error("{0}: {1}".format(r.status_code, r.text))

    return messages


def get_num_messages(access_token, user_email, folders=None):
    """
    Returns the total number of unread messages in the logged-in
    user's Inbox.  This should be used in preference to counting the
    result of "get_new_messages" since it isn't affected by the limit
    on returned values (see "get_new_messages" for details)

    """
    # Base url for messages
    get_messages_url = GRAPH_ENDPOINT.format('/me/mailfolders')

    # If no list of folders supplied, default to just the Inbox
    if folders is None:
        folders = ['inbox']
    else:
        folders = [folder.lower() for folder in folders]

    # Logical stores if there are more pages of results to return
    next_page = True
    message_count = 0
    while next_page:
        try:
            r = make_api_call('GET', get_messages_url, access_token,
                              user_email)
        except Exception as err:
            raise o365Error(err)

        if (r.status_code == requests.codes.ok):
            output = r.json()

            for folder in output['value']:
                if folder['displayName'].lower() in folders:
                    message_count += folder['unreadItemCount']

            if '@odata.nextLink' in output:
                get_messages_url = output['@odata.nextLink']
            else:
                next_page = False

        else:
            raise o365Error("{0}: {1}".format(r.status_code, r.text))

    return message_count


def get_week_events(access_token, user_email):
    """
    Returns all calendar events on the logged-in user's calendar starting
    from the current time and for the next week; this is arbitrary but
    considered a long enough time to hopefully not miss meetings with
    the longest expected reminder duration

    """
    today_start = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:00.0000000")
    next_week_start = (
        datetime.utcnow()
        + timedelta(days=7)).strftime("%Y-%m-%dT00:00:00.0000000")

    url_query = ("?startDateTime={0}&endDateTime={1}"
                 .format(today_start, next_week_start))

    get_events_url = GRAPH_ENDPOINT.format('/me/calendarView' + url_query)

    # Use OData query parameters to control the results
    query_parameters = {
        '$select': ('isReminderOn, reminderMinutesBeforeStart,'
                    'subject,start,location,isCancelled'),
        '$orderby': 'start/dateTime ASC'}

    try:
        r = make_api_call('GET', get_events_url, access_token,
                          user_email, parameters=query_parameters)
    except Exception as err:
        raise o365Error(err)

    if (r.status_code == requests.codes.ok):
        return r.json()
    else:
        raise o365Error("{0}: {1}".format(r.status_code, r.text))


def get_user_portrait(access_token, user_email, size='64x64'):
    """
    Returns the user's portrait image (if they have one!)

    """
    get_portrait_url = GRAPH_ENDPOINT.format(
        "/users/{0}/photos/{1}/$value".format(user_email, size))

    try:
        r = make_api_call('GET', get_portrait_url,
                          access_token, user_email, timeout=2)
    except Exception as err:
        raise o365Error(err)

    if (r.status_code == requests.codes.ok):
        return r.content
    else:
        raise o365Error("{0}: {1}".format(r.status_code, r.text))
