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
import sys
import subprocess
import time
import signal
import logging
import logging.handlers
import o365_api
import oauth2
import stream_logging
from datetime import datetime, timedelta
from pytz import timezone
from ConfigParser import ConfigParser
from status_icon import OWAConfig, OWAErrorDialog, CONFIG_FILE_DIR

# Get the path where the checker has been installed
OWA_CHECKER_PATH = os.path.dirname(os.path.realpath(sys.argv[0]))

# Retry maximum (how many times the same call to the API is
# allowed to fail before bringing down the checker)
MAX_RETRIES = 5

# Main timing loop settings - the check interval (the check for new mail)
# is in seconds, and the check for issuing of reminders or checking for
# new/updated calendar events are each multiples of the check interval
CHECK_INTERVAL = 5
REMINDER_MULTIPLE = 1
CALENDAR_MULTIPLE = 12

# How many mail popups should be issue in one call before resorting to
# a popup that says "+ further messages"
MAIL_POPUP_LIMIT = 4

# The number of times communication with the status icon should be retried
# if it fails to respond to stdin communication
MAX_STATUS_ICON_RETRIES = 2


class OWAError(Exception):
    """
    Main exception class, which will issue a warning dialog when something
    goes wrong - so that the user knows it has happened
    """
    def __init__(self, message, owa_instance=None):

        # Popup a warning message
        OWAErrorDialog(message)

        # If the checker is running clear up before crashing
        if owa_instance is not None:
            owa_instance.cleanup()


class OWACheck(object):
    """
    Outlook Web Access checker class, handles the connections and the checking
    of the mail and calendar
    """
    def __init__(self, logger):
        """
        Setup the checker

        """
        # Save a reference to the logger
        self.logger = logger

        # Initialise the notification icon, this is a separate python process
        # which looks after the icon, it is given the PID of this process so
        # that it can send a SIGTERM signal when the user clicks "quit" to
        # bring down this process
        self.logger.info('Running, PID: {0}'.format(os.getpid()))
        self.statusicon_init()

        # Try to load the user's refresh token
        oauth2.load_refresh_token()
        if oauth2.OWACHECKER_SESSION['refresh_token'] == "":
            msg = ("Unable to load user refresh token; please run the "
                   "setup utility before this checker")
            self.logger.error(msg)
            raise OWAError(msg)
        try:
            oauth2.get_token_from_refresh_token()
        except ValueError:
            msg = ("Unable to load user access token; please run the "
                   "setup utility before this checker")
            self.logger.error(msg)
            raise OWAError(msg)

        # Now test the o365 api by retrieving the user's details
        try:
            self.user = o365_api.get_user_info(oauth2.get_access_token())
        except o365_api.o365Error as err:
            self.logger.error(err)
            self.cleanup()

        # Get the user's config settings
        self.config = OWAConfig()

        self.check_mail_last = None
        self.reminders = {}
        self.n_msgs_last = 0
        self.retries = 0

        self.logger.info("Successfully initialised OWA Checker!")

    def check_mail(self, quiet=False):
        """
        Check for new mail and notify the user if needed

        """
        # If the user's config file has changed, re-read it
        cf_mtime = os.path.getmtime(self.config.file)
        if cf_mtime > self.config.mtime:
            self.config.read_config()
            self.config.mtime = cf_mtime

        # Get the total number of unread messages
        try:
            n_msgs = o365_api.get_num_messages(
                oauth2.get_access_token(),
                self.user['mail'],
                folders=self.config.folder_list.split("::"))
            self.retries = 0
        except o365_api.o365Error as err:
            self.retries += 1
            self.logger.error("Failed to get num messages (retry {0})"
                              .format(self.retries))
            self.logger.error(err)
            if self.retries <= MAX_RETRIES:
                return
            else:
                self.cleanup()

        # Notify the user of the number of unread messages by updating the icon
        self.statusicon_communicate("status:n_msgs:{0:d}\n".format(n_msgs))

        # If the number of messages found has fallen, switch off the LED
        if n_msgs < self.n_msgs_last:
            self.statusicon_communicate("status:blink:0\n")

        # Save the amount of messages for the next check
        self.n_msgs_last = n_msgs

        # Get the new messages
        messages = None
        if n_msgs > 0:
            try:
                messages = o365_api.get_new_messages(
                    oauth2.get_access_token(),
                    self.user['mail'],
                    last_seen=self.check_mail_last,
                    folders=self.config.folder_list.split("::"))
                self.retries = 0
            except o365_api.o365Error as err:
                self.retries += 1
                self.logger.error("Failed to get new messages (retry {0})"
                                  .format(self.retries))
                self.logger.error(err)
                if self.retries <= MAX_RETRIES:
                    return
                else:
                    self.cleanup()

        # Now loop through the messages issuing popups
        if not quiet and messages:
            # Limit the maximum number of popups
            for imessage, message in enumerate(messages):

                if (self.check_mail_last is not None
                        and (message['receivedDateTime']
                             == self.check_mail_last)):
                    continue

                # Only print messages up to a given limit
                if imessage == MAIL_POPUP_LIMIT:
                    break

                # Extract the relevant information
                sender = message['from']['emailAddress']
                subject = message['subject']

                # Handle case where there is no subject
                if subject is None:
                    subject = "(No Subject)"

                # Switch on the blinker (if active)
                self.statusicon_communicate("status:blink:1\n")

                # Send a command to the icon to issue the popup
                self.statusicon_communicate(
                    "popup:{0:s}:{1:s}:{2:s}\n"
                    .format(sender['name'].replace(":", "\\:"),
                            subject.replace(":", "\\:"),
                            sender.get("address", "(No Address)")))

            # If there were messages past the popup limit, finish off with an
            # extra popup stating that there are more
            if len(messages) >= MAIL_POPUP_LIMIT:
                self.statusicon_communicate(
                    "popup:Plus further message(s)...:\n")

        # Save the time of the most recent email (ensuring we never
        # repeat or miss a new email)
        if messages:
            most_recent = messages[0]['receivedDateTime']
            for message in messages[1:]:
                if message['receivedDateTime'] > most_recent:
                    most_recent = message['receivedDateTime']

            if (self.check_mail_last is None
                    or most_recent > self.check_mail_last):
                self.check_mail_last = most_recent

        # Finally activate the status icon to show the popup etc.
        self.statusicon_communicate('\n\n')

    def check_calendar(self):
        """
        Check the user's calendar and save or update any pending events

        """
        current_time = datetime.now(timezone("UTC"))
        # If the user's config file has changed, re-read it
        cf_mtime = os.path.getmtime(self.config.file)
        if cf_mtime > self.config.mtime:
            self.config.read_config()
            self.config.mtime = cf_mtime

        # Get the user's calendar appointments
        try:
            events = o365_api.get_week_events(
                oauth2.get_access_token(), self.user['mail'])
            self.retries = 0
        except o365_api.o365Error as err:
            self.retries += 1
            self.logger.error("Failed to get events (retry {0})"
                              .format(self.retries))
            self.logger.error(err)
            if self.retries <= MAX_RETRIES:
                return
            else:
                self.cleanup()

        # Loop through the events - first take a copy of the current reminders,
        # events will be removed from this copy as they are seen to detect
        # cancelled events later
        reminders = self.reminders.copy()

        for event in events["value"]:
            # Ignore any with no reminder setting, or that have been cancelled
            if not event["isReminderOn"] or event["isCancelled"]:
                continue

            start = event["start"]["dateTime"]
            reminder = event["reminderMinutesBeforeStart"]
            subject = event["subject"]
            location = event["location"]["displayName"]

            # Convert the start time to a UTC datetime object
            start = datetime(*(list(time.strptime(
                start[:start.rindex(".")],
                "%Y-%m-%dT%H:%M:%S")[0:6]) + [0, timezone("UTC")]))

            # Calculate the reminder time
            reminder_time = start - timedelta(minutes=reminder)

            # If the event is already in the reminders cache, we have seen
            # it before, but if it has been set to None that means the
            # notification has been issued, so we should avoid adding it
            # to the reminders dict again
            if event["id"] in self.reminders:
                # Remove this event from the copied dictionary
                reminders.pop(event["id"])
                if self.reminders[event["id"]][0] is True:
                    continue

            # Reaching this point means either the event is new or is an
            # existing event which has not yet been issued; so create or
            # update its value here

            # The API will sometimes return events earlier than the requested
            # time (no idea why), especially when the checker is first run.
            # So simply catch events that are already in the past here
            if current_time > start:
                continue

            self.reminders[event["id"]] = [
                False, reminder_time, start, subject, location]

        # If the copied dictionary still contains events, they must have been
        # cancelled, or have expired, so remove them from the main dictionary
        if events["value"]:
            for reminder in reminders:
                self.reminders.pop(reminder)

    def issue_reminders(self):
        """
        Check if any reminders are due to be issued and issue a calendar
        reminder

        """
        current_time = datetime.now(timezone("UTC"))
        for event_id in self.reminders.keys():
            issued, reminder_time, start, subject, location = (
                self.reminders[event_id])

            # Skip this event if it was already issued
            if issued:
                continue

            # Check if the reminder needs to be issued
            if current_time >= reminder_time:
                # Construct the command to the other script; note that by using
                # the current time here instead of the reminder time we can
                # ensure the time is always accurate (if for example the
                # checker is launched *after* the reminder was due to start)
                reminder_mins = start - current_time
                reminder_mins = int(
                    reminder_mins.days * 1440
                    + reminder_mins.seconds / 60 + 1)

                command = [
                    "python",
                    os.path.join(OWA_CHECKER_PATH, "calendar_popup.py"),
                    "{0:d}".format(reminder_mins)]

                # Setup the display environment and spawn the command, note
                # that we use environment variables in the subshell to pass
                # the location and subject; otherwise these potentially
                # private details would be viewable in the process viewer
                temp_env = os.environ.copy()
                temp_env["DISPLAY"] = self.config.cal_display
                if subject is None or subject == "":
                    subject = "Untitled Meeting"
                temp_env["OWA_MEETING_TITLE"] = subject
                if location is None or location == "":
                    location = "No Location Set"
                temp_env["OWA_MEETING_LOCATION"] = location

                subprocess.Popen(command, env=temp_env)

                # Since this reminder has now been issued, update the entry
                # to avoid it issuing a second notification
                self.reminders[event_id][0] = True

    def cleanup(self, signal=None, stack=None):
        """
        Signal handler should the checker be killed or interrupted; this
        will bring down the status icon, which will in-turn kill this
        process (as if the user clicked "quit" in the menu) - ensuring both
        OWA processes get brought down together

        """
        self.logger.debug("Received kill signal, attempting to close nicely")

        try:
            self.statusicon_communicate("quit\n")
            self.statusicon_communicate("\n")
        except Exception:
            self.logger.error("Problem sending quit signal to status icon")
        os._exit(0)

    def statusicon_init(self):
        """
        Initialise the status icon subprocess which handles the display of
        the (mail) popups, the notification area icon and the configuration
        menu

        """
        self.statusicon = subprocess.Popen(
            ['python', os.path.join(OWA_CHECKER_PATH, 'status_icon.py')],
            stdin=subprocess.PIPE)

    def statusicon_communicate(self, command):
        """
        Send a command to the status icon process.  If for any reason the
        pipe connected to the process has broken (usually if the icon
        crashes or similar) attempt to restart it

        """
        retries = 0
        while retries < MAX_STATUS_ICON_RETRIES:
            try:
                self.statusicon.stdin.write(command)
                break
            except IOError as err:
                if err.strerror == 'Broken pipe':
                    self.logger.error(
                        "Connection to status icon broken, "
                        "attempting to restart it")
                    self.statusicon_init()
            retries += 1


def main_loop(logger):
    """
    Sets up the checker and begins the main loop to start monitoring for
    messages and events

    """
    # Setup OWA checker
    owa = OWACheck(logger)

    # Set logging level based on user preference
    if owa.config.debuglog:
        logger.setLevel(logging.DEBUG)
    else:
        logger.setLevel(logging.INFO)

    # Setup trapping for signals, to allow the checker to be
    # killed nicely and tidy up after itself
    catch = ['SIGTERM', 'SIGINT', 'SIGABRT', 'SIGHUP', 'SIGQUIT']
    for sig in catch:
        signum = getattr(signal, sig)
        signal.signal(signum, owa.cleanup)

    # Setup the counters for the calendar and reminder checks
    calendar_count = CALENDAR_MULTIPLE
    reminder_count = REMINDER_MULTIPLE

    # Begin the loop
    first_pass = True
    while True:
        # Increment the counters
        calendar_count += 1
        reminder_count += 1

        # Check the user's email every cycle
        try:
            owa.check_mail(quiet=first_pass)
            if first_pass:
                first_pass = False
        except Exception:
            msg = "Failed to check mail"
            logger.error(msg, exc_info=True)
            raise OWAError(msg, owa)

        # If this is a calendar cycle also check the calendar (note this is
        # greater-than or equal to account for the unlikely event of a
        # skipped check)
        if calendar_count >= CALENDAR_MULTIPLE:
            try:
                owa.check_calendar()
                calendar_count = 0
            except Exception:
                msg = "Failed to check calendar"
                logger.error(msg, exc_info=True)
                raise OWAError(msg, owa)

        # If this is a reminder cycle also check to see if any meeting
        # reminders need to be sent
        if reminder_count >= REMINDER_MULTIPLE:
            try:
                owa.issue_reminders()
                reminder_count = 0
            except Exception:
                msg = "Failed to issue reminders"
                logger.error(msg, exc_info=True)
                raise OWAError(msg, owa)

        # Rather than simply sleeping the amount of the interval,
        # keep the time in sync with the clock and sleep until the
        # next time the given interval is hit
        sleep_secs = CHECK_INTERVAL - (time.time() % CHECK_INTERVAL)
        time.sleep(sleep_secs)


if __name__ == "__main__":

    # Create a logger
    log_dir = os.path.join(CONFIG_FILE_DIR)
    if not os.path.exists(log_dir):
        os.mkdir(log_dir)

    logger = stream_logging.create_logger(
        os.path.join(log_dir, "owa_checker.log"))

    # Daemonize the process - start by forking
    try:
        pid = os.fork()
        if pid > 0:
            # Close parent
            os._exit(0)
    except OSError as error:
        msg = (
            'Unable to fork (1). Error: {0:d} ({1:s})'.format
            (error.errno, error.strerror))
        raise OWAError(msg)

    # Decouple from parent environment - move to root to ensure the CWD
    # of the process cannot dissapear, make this process a session leader
    # and clear the umask
    os.chdir("/")
    os.setsid()
    prev_umask = os.umask(0)

    # Fork the process again - due to the above this process will not
    # be a session leader (because its parent is)
    try:
        pid = os.fork()
        if pid > 0:
            # Close parent
            os._exit(0)
    except OSError as error:
        msg = (
            'Unable to fork (2). Error: {0:d} ({1:s})'.format
            (error.errno, error.strerror))
        raise OWAError(msg)

    # Reset umask
    os.umask(prev_umask)

    # Redirect /dev/null to stdin
    dvnl = file(os.devnull, 'r')
    os.dup2(dvnl.fileno(), sys.stdin.fileno())

    # Begin looping
    main_loop(logger)
