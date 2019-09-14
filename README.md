(C) British Crown Copyright 2019 Met Office.

Author: Steve Wardle

OWA Checker
-----------

OWA Checker is a tool to provide native popup notifications for mail and
calendar events for users using Microsoft's OWA (Outlook Web Access) - the
online version of Outlook


Prerequisites (Required)
------------------------
 * Python (2.7 currently, 3+ support pending investigation)
 * gtk (for interfacing with GNOME)
 * gobject (for interfacing with GNOME)


Prerequisites (Optional)
------------------------
 * wmctrl (Command-line utility for switching workspaces)
 * xset (Command-line utility used for blinking scroll lock LED)

Microsoft Azure Setup
---------------------
Before you can use OWA Checker you need to create a new application profile
for it within Microsoft's Azure system.  The steps here may vary depending on
your site but will be something like:

 1. Navigate to portal.azure.com and sign in using your account (which
    should redirect to your own site's infrastructure)
 2. The app name doesn't matter, call it something like "OWA Checker"
 3. Under Authentication use "Web" and Redirect URI of "http://localhost:1234"
 4. In the Overview, make a note of the "Application (client) ID"
 5. Under Certificates, make a note of the "Client Secret" (create one if
    needed)
 6. Under API Permissions, ensure that the app has the following:
     * Calendars.Read
     * Mail.Read
     * User.Read
     * User.ReadBasic.All
     * offline_access
     * openid

NOTE - if any of the above will not work, it is likely that you will need to
       ask your local platforms team for help here - some of the settings need
       and Admin level account to apply them.  In my case I added such a user
       as an additional "Owner" of the app and had them apply these changes


Checker Setup
-------------
With the above done you will need to create a wrapper script to launch the
checker which provides it with the app information.  The form of the wrapper
script should be something like this:

    #!/bin/bash

    # Authentication and site-specific settings (from Azure)
    export OWA_CHECKER_CLIENT_ID=
    export OWA_CHECKER_CLIENT_SECRET=
    export OWA_CHECKER_DOMAIN=

    exec /installation/path/to/owa/checker/app/owa_checker "$@"

Populate the environment variables accordingly with the Application ID and 
secret. For the domain this is your local site's email domain (i.e. the part
of your email addresses following the @ sign)


Running the Checker
-------------------
The first time you run the checker you should provide the extra argument:

    owa_checker --auth

This should prompt you to navigate in a browser to http://localhost:1234
and doing so should prompt you to login with your credentials.  Once this 
step is done you can re-launch the checker without the argument to begin
using it (note that you can then continue to launch it this way and should
not need to authenticate again unless there is a problem)
