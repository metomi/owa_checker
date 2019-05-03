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
import re
import sys
import gtk
import cgi
import gobject
import subprocess
import stream_logging
import tempfile
import socket
from status_icon import CONFIG_FILE_DIR

# Get the path where the checker has been installed
OWA_CHECKER_PATH = os.path.dirname(os.path.realpath(sys.argv[0]))

# The hostname
HOSTNAME = socket.gethostname()

# Choices for snooze (how many minutes to snooze for)
SNOOZE_DROPDOWN_CHOICES = [5, 10, 15, 30, 60, 120, 240, 360]

# How many seconds each "tick" should decrement the timer
COUNTDOWN_TICK_SECS = 60

# How many ticks must pass after the event time for the reminder
# to close itself without intervention from the user
GIVE_UP_TICKS = 60

# Various options to control appearance of reminder
REMINDER_BUTTON_PADDING = 2
REMINDER_VERTICAL_SPACING = 5
REMINDER_HORIZONTAL_SPACING = 10
REMINDER_LOGO_SIZE = 150
REMINDER_OUTER_PADDING = 2


class CalendarPopup(object):
    """
    Defines a calendar popup window, which indicates the approaching
    start time of a meeting or event

    """
    def __init__(self, description, location, minutes_total, logger):

        self.logger = logger
        self.logger.info('Running, PID: {0}'.format(os.getpid()))

        self.popup = gtk.Window(gtk.WINDOW_POPUP)
        self.popup.set_position(gtk.WIN_POS_CENTER)
        self.minutes_total = minutes_total

        # Calendar popup Icon
        icon = gtk.Image()
        pixbuf = gtk.gdk.pixbuf_new_from_file(
            os.path.join(OWA_CHECKER_PATH, "icons", "owa_logo.png"))
        sclbuf = pixbuf.scale_simple(
            REMINDER_LOGO_SIZE, REMINDER_LOGO_SIZE, gtk.gdk.INTERP_BILINEAR)
        icon.set_from_pixbuf(sclbuf)

        # Description (of the calendar appointment)
        descr_label = gtk.Label()
        description = cgi.escape(description)
        descr_label.set_markup("<b>{0:s}</b>".format(description))
        descr_label.set_line_wrap(True)
        descr_label.set_width_chars(40)
        descr_label.set_justify(gtk.JUSTIFY_CENTER)

        # Timer which counts down until appointment
        self.timer_label = gtk.Label(
            "In {0:d} minutes".format(self.minutes_total))

        # Location (of the calendar appointment)
        locn_label = gtk.Label()
        location = cgi.escape(location)
        locn_label.set_markup("{0:s}".format(location))
        locn_label.set_line_wrap(True)
        locn_label.set_justify(gtk.JUSTIFY_CENTER)

        # Button which closes the popup and cancels the reminders
        button_d = gtk.Button("Dismiss")
        button_d.connect("clicked", self.dismiss)

        # Button which hides the popup until the next reminder time
        button_s = gtk.Button("Snooze")
        button_s.connect("clicked", self.snooze)
        snooze_label = gtk.Label()
        snooze_label.set_markup(" for ")
        self.snooze_dropdown = gtk.combo_box_new_text()
        self.snooze_dropdown_choices = SNOOZE_DROPDOWN_CHOICES

        for mins in self.snooze_dropdown_choices:
            self.snooze_dropdown.append_text("{0} Minutes".format(mins))
        self.snooze_dropdown.set_active(0)

        # Pack the 2 buttons side-by-side
        hbox_buttons = gtk.HBox()
        hbox_buttons.pack_start(
            button_d, expand=True, fill=True,
            padding=REMINDER_BUTTON_PADDING)
        hbox_buttons.pack_start(
            button_s, expand=True, fill=True,
            padding=REMINDER_BUTTON_PADDING)
        hbox_buttons.pack_start(
            snooze_label, expand=True, fill=True,
            padding=REMINDER_BUTTON_PADDING)
        hbox_buttons.pack_start(
            self.snooze_dropdown, expand=True, fill=True,
            padding=REMINDER_BUTTON_PADDING)

        # Pack the description, timer, location and buttons vertically
        vbox = gtk.VBox()
        vbox.pack_start(
            descr_label, expand=True, fill=True,
            padding=REMINDER_VERTICAL_SPACING)
        vbox.pack_start(
            self.timer_label, expand=True, fill=True,
            padding=REMINDER_VERTICAL_SPACING)
        vbox.pack_start(
            locn_label, expand=True, fill=True,
            padding=REMINDER_VERTICAL_SPACING)
        vbox.pack_start(
            hbox_buttons, expand=True, fill=False,
            padding=REMINDER_VERTICAL_SPACING)

        # Pack the list of labels and buttons side-by-side the icon
        hbox_main = gtk.HBox()
        hbox_main.pack_start(
            icon, expand=False, fill=True,
            padding=REMINDER_HORIZONTAL_SPACING)
        hbox_main.pack_start(
            vbox, expand=False, fill=True,
            padding=REMINDER_HORIZONTAL_SPACING)
        hbox_main.set_border_width(REMINDER_HORIZONTAL_SPACING)

        # Put this main box into an outer frame, since the nice border
        # it provides looks a little nicer than a frame-less window
        outer_frame = gtk.Frame()
        outer_frame.set_shadow_type(gtk.SHADOW_IN)
        outer_frame.set_border_width(REMINDER_OUTER_PADDING)
        outer_frame.add(hbox_main)

        self.popup.add(outer_frame)

        # Before activating the popup, do a quick check to see if the
        # screensaver is active (because a gtk.WINDOW_POPUP window will
        # appear *above* the screensaver, which we don't want)
        if self.screensaver_isactive():
            self.screensaver = True
        else:
            self.screensaver = False
            self.popup.show_all()

        self.timer = self.minutes_total
        self.snoozed = False

        # Start the screensaver monitor, to watch for when the screensaver
        # is active and to take action when it is de-activated
        gobject.timeout_add(10000, self.screensaver_check)

        # Start the countdown timer (count down every minute)
        gobject.timeout_add(COUNTDOWN_TICK_SECS * 1000, self.decrement_timer)

        gtk.main()

    def screensaver_isactive(self):
        """
        Checks if the screensaver (screen-lock) is currently active; this
        is important as it's actually possible for the popup to show above
        the lock screen!

        """
        # Note that the below method does not work on the VDI machines, as
        # they have the lock disabled, so only need to do this check on non
        # VDI machines
        if HOSTNAME.startswith("vld"):
            return False
        else:
            try:
                saver_check = subprocess.Popen(
                    ['gsettings', 'get', 'org.gnome.desktop.lockdown',
                     'disable-lock-screen'],
                    stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            except OSError:
                # If the command isn't found or fails, respond as if the
                # screen is not locked (but log it as an error)
                self.logger.error("Failed to check for screen lock")
                return False

            stdout, stderr = saver_check.communicate()
            if saver_check.returncode != 0:
                self.logger.error(stderr)
                return False
            return "true" in stdout

    def dismiss(self, event):
        """
        Called if the user clicks the "dismiss" button

        """
        self.popup.destroy()
        gtk.main_quit()

    def snooze(self, event):
        """
        Called if the user clicks the "snooze" button; the windows is
        hidden from view after the time in the dropdown is saved

        """
        self.popup.hide()
        self.snooze_interval = (
            self.snooze_dropdown_choices[
                self.snooze_dropdown.get_active()])
        self.snooze_count = 1
        self.snoozed = True

    def screensaver_check(self):
        """
        Checks if the screensave is active, and shows the popup if it
        has been de-activated since the last check (so that someone
        returning to their desk with an impending meeting will be
        notified about it immediately)

        """
        if self.screensaver_isactive():
            self.screensaver = True
        else:
            if (self.screensaver
                    and self.snoozed
                    and self.snooze_count >= self.snooze_interval):
                self.popup.show_all()
            self.screensaver = False
        return True

    def decrement_timer(self):
        """
        Called once each minute to progress the countdown to the event

        """
        self.timer -= 1

        min_or_mins = "minutes"
        if abs(self.timer) == 1:
            min_or_mins = "minute"

        if self.timer > 0:
            # If there is still time remaining before the appointment,
            # update the popup text accordingly
            self.timer_label.set_text(
                "In {0:d} {1:s}".format(self.timer, min_or_mins))

            if self.snoozed:
                if self.snooze_count >= self.snooze_interval:
                    # If the user previously pressed snooze to hide the window,
                    # make it visible again if it has reached the snooze time
                    if not self.screensaver:
                        self.popup.show()
                else:
                    # Increment the counter
                    self.snooze_count += 1

        elif self.timer == 0:
            # If the timer has finished counting down the meeting is now, make
            # sure the window is re-shown if it was previously hidden with
            # snooze
            self.timer_label.set_text("Now!")
            if self.snoozed and not self.screensaver:
                self.popup.show()
        else:
            # If the timer has gone below 0 then continue to remind the user
            # with a different message
            self.timer_label.set_text(
                "{0:d} {1:s} ago!!!".format(-self.timer, min_or_mins))

            if self.snoozed and not self.screensaver:
                # If the meeting has started just re-show the popup
                # every minute
                self.popup.show()

            if self.timer == -GIVE_UP_TICKS:
                # If it's been more than an hour since the meeting started,
                # assume that for whatever reason the reminder is no longer
                # needed
                self.dismiss(None)

        # The gobject requires that the timer function returns True
        return True


if __name__ == "__main__":

    # Create a logger
    log_dir = os.path.join(CONFIG_FILE_DIR)
    if not os.path.exists(log_dir):
        os.mkdir(log_dir)

    # There could be multiple of these, so make sure they're unique
    log_file = tempfile.mktemp(
        prefix="calendar_popup_", suffix=".log", dir=log_dir)

    logger = stream_logging.create_logger(log_file)

    # Get the description and location (note these are passed via
    # environment variables set by the checker; to avoid the possibly
    # private subject and location of the meeting being shown in
    # process viewers)
    description = os.environ.get("OWA_MEETING_TITLE", "Unknown")
    location = os.environ.get("OWA_MEETING_LOCATION", "Unknown")

    # The minutes are passed directly
    minutes_total = int(sys.argv[1])

    # Issue the popup
    popup = CalendarPopup(
        description, location, minutes_total, logger)

    # Remove the log file after successful execution
    os.remove(log_file)
