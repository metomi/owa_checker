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
import subprocess
import gobject
import o365_api
import oauth2
import stream_logging
import socket
import warnings
from blink import ScrollBlink
from ConfigParser import ConfigParser

# Get the path where the checker has been installed
OWA_CHECKER_PATH = os.path.dirname(os.path.realpath(sys.argv[0]))

# Get the local domain (used to decide whether to try and display
# user avatars or not)
LOCAL_DOMAIN = os.environ.get("OWA_CHECKER_DOMAIN", None)

if LOCAL_DOMAIN is None:
    msg = ("Local domain not found; OWA Checker will not display "
           "any user avatars (to remedy this, provide your site's "
           "domain via environment variable OWA_CHECKER_DOMAIN "
           "before starting OWA checker)")
    warnings.warn(msg)

# The hostname
HOSTNAME = socket.gethostname()

# And the dimensions of the screen/s
screen = gtk.gdk.Screen()
N_SCREENS = screen.get_display().get_n_screens()
SCREEN_X, SCREEN_Y = screen.get_width(), screen.get_height()

CONFIG_FILE_DIR = os.path.join(os.environ['HOME'], ".owa_check")

# Values used in the popup "about" dialog
ABOUT_DIALOG_VERSION = "2.2"
ABOUT_DIALOG_COPYRIGHT = "Met Office Crown Copyright 2019"

# Various sizing values for the popups
POPUP_SCREEN_EDGE_PADDING = 5
POPUP_GNOME3_TOPBAR_PADDING = 25
POPUP_GNOME3_BOTTOMBAR_PADDING = 30
POPUP_AVATAR_SIZE = 50
POPUP_TEXT_WIDTH = 300
POPUP_TEXT_PADDING = 5
POPUP_AVATAR_PADDING = 5
POPUP_SEPARATION = 10

# Various sizing values for the "configuration" dialog
CONFIG_OPTION_WIDTH = 20
CONFIG_BORDER_WIDTH = 20
CONFIG_LOGO_SIZE = 30
CONFIG_OPTION_PADDING = 2
CONFIG_SECTION_SEPARATION = 10


class OWAConfig(ConfigParser):
    """
    Defines the configuration options which are stored in a per-user
    file and used to control various aspects of the checker

    """

    def __init__(self, *args):
        self.file = os.path.join(CONFIG_FILE_DIR, "owa.conf")

        # Setup default options
        self.timeoutms = 7000
        self.showface = 1
        self.blink = 0
        self.mail_pos = "top-right"
        self.mail_opacity = 90
        self.workspace = 0
        self.mail_display = ":0.0"
        self.cal_display = ":0.0"
        self.folder_list = "inbox"
        self.debuglog = 0

        ConfigParser.__init__(self, *args)
        if os.path.exists(self.file):
            self.read_config()
        else:
            self.save_config()
        self.mtime = os.path.getmtime(self.file)

    def read_config(self):
        """
        Read in any options which are found in the config file. By not
        assuming any option is there new options can be easily added or
        removed later without breaking anything when the application tries
        to load a configuration file from an earlier version

        """
        OWAConfig.read(self, self.file)

        opt = (self, 'mail_checker', 'timeout_seconds')
        if OWAConfig.has_option(*opt):
            self.timeoutms = int(OWAConfig.get(*opt)) * 1000

        opt = (self, 'mail_checker', 'show_faces')
        if OWAConfig.has_option(*opt):
            self.showface = int(OWAConfig.get(*opt))

        opt = (self, 'mail_checker', 'blink')
        if OWAConfig.has_option(*opt):
            self.blink = int(OWAConfig.get(*opt))

        opt = (self, 'mail_checker', 'workspace')
        if OWAConfig.has_option(*opt):
            self.workspace = int(OWAConfig.get(*opt))

        opt = (self, 'mail_checker', 'position')
        if OWAConfig.has_option(*opt):
            self.mail_pos = OWAConfig.get(*opt)

        opt = (self, 'mail_checker', 'opacity')
        if OWAConfig.has_option(*opt):
            self.mail_opacity = int(OWAConfig.get(*opt)) / 100.0

        opt = (self, 'mail_checker', 'folder_list')
        if OWAConfig.has_option(*opt):
            self.folder_list = OWAConfig.get(*opt)

        opt = (self, 'mail_checker', 'debuglog')
        if OWAConfig.has_option(*opt):
            self.debuglog = int(OWAConfig.get(*opt))

        # Some options should only be read in if the user has multiple DISPLAYs
        if N_SCREENS > 1:
            opt = (self, 'calendar_notifier', 'notify_display')
            if OWAConfig.has_option(*opt):
                self.cal_display = OWAConfig.get(*opt)

            opt = (self, 'mail_checker', 'display')
            if OWAConfig.has_option(*opt):
                self.mail_display = OWAConfig.get(*opt)

    def save_config(self):
        """
        Save the configuration options to the configuration file

        """
        if not OWAConfig.has_section(self, 'mail_checker'):
            OWAConfig.add_section(self, 'mail_checker')
        if not OWAConfig.has_section(self, 'calendar_notifier'):
            OWAConfig.add_section(self, 'calendar_notifier')

        OWAConfig.set(
            self, 'mail_checker', 'timeout_seconds', self.timeoutms / 1000)
        OWAConfig.set(
            self, 'mail_checker', 'show_faces', self.showface)
        OWAConfig.set(
            self, 'mail_checker', 'blink', self.blink)
        OWAConfig.set(
            self, 'mail_checker', 'workspace', self.workspace)
        OWAConfig.set(
            self, 'mail_checker', 'position', self.mail_pos)
        OWAConfig.set(
            self, 'mail_checker', 'opacity',
            int(round(self.mail_opacity * 100.0)))
        OWAConfig.set(
            self, 'mail_checker', 'folder_list', self.folder_list)
        OWAConfig.set(
            self, 'mail_checker', 'debuglog', self.debuglog)

        # Some options only need to be saved if the user has multiple DISPLAYs
        if N_SCREENS > 1:
            OWAConfig.set(
                self, 'calendar_notifier', 'notify_display', self.cal_display)
            OWAConfig.set(
                self, 'mail_checker', 'display', self.mail_display)

        with open(self.file, 'w') as fp:
            OWAConfig.write(self, fp)


def OWAErrorDialog(message):
    """
    A simple error message dialog to popup in case of an error, since
    we won't have a terminal to report to

    """
    dialog = gtk.MessageDialog(parent=None,
                               flags=0,
                               type=gtk.MESSAGE_WARNING,
                               buttons=gtk.BUTTONS_OK,
                               message_format=None)
    dialog.set_position(gtk.WIN_POS_CENTER)
    dialog.set_border_width(10)
    dialog.set_decorated(False)
    dialog.set_keep_above(True)
    dialog.set_markup(
        "<b>OWA Checker has encountered a problem</b>\n\n"
        "<i>" + message + "</i>\n\n"
        "Logs written to {0:s}"
        .format(os.path.join(CONFIG_FILE_DIR, "owa_checker.log")))
    dialog.run()


class OWAStatusIcon(object):
    """
    Defines and controls the notification area icon, the configuration
    and about windows, and the popup notifications

    """
    def __init__(self, logger):
        self.logger = logger
        self.logger.info('Running, PID: {0}'.format(os.getpid()))

        # Get the parent PID (for quitting) and initial configuration
        self.parent = os.getppid()
        self.logger.info('From OWA Checker with PID: {0}'.format(self.parent))
        self.config = OWAConfig()

        # Create a connection to the API
        self.o365 = False
        try:
            oauth2.load_refresh_token()
            oauth2.get_token_from_refresh_token()
            self.o365 = True
        except Exception:
            pass

        # Setup the Status Icon
        self.init_StatusIcon()
        self.n_msgs_shown = 0

        # Create the blinker, disable it immediately if it isn't wanted
        self.blinker = ScrollBlink(blink_speed_ms=1000)
        self.blinker.blinking = False
        self.blinker.active = (False, True)[self.config.blink]

        # Setup the silence setting
        self.silenced = False

        # Setup wmctrl
        self.set_wmctrl_command(self.config.workspace)

        # Setup popup manager
        self.init_PopupManager(self.config.mail_pos)

        # Watch stdin for any inputs, which is how the icon is controlled
        gobject.io_add_watch(sys.stdin, gobject.IO_IN, self.receive_stdin)

        self.logger.info("Successfully initialised OWA Checker Status Icon!")

    def set_wmctrl_command(self, owa_desktop):
        """
        Wmctrl is a utility which automates various aspects of window manager
        interaction; here construct an appropriate command to switch to the
        given workspace

        """
        temp_env = os.environ.copy()
        temp_env["DISPLAY"] = self.config.mail_display
        self.wmctrl_command = [
            "wmctrl",
            "-s",
            "{0:d}".format(owa_desktop)]

    def receive_stdin(self, stdin, cb_condition):
        """
        Parses the STDIN arriving from the main process, and issues the
        appropriate responses

        """
        # Lines should be of the linux standard colon-delimited format, if they
        # aren't then just ignore them, note that this function must return
        # True as per the documentation for gobject
        line = stdin.readline().replace("\n", "")
        while line:
            # Split up the line, ignoring escaped colons
            split_line = re.split(r'(?<!\\):', line)

            # The first element gives the component to change
            component = split_line[0]

            # Changing the number of messages or LED state of the icon
            if component == "status":
                action = split_line[1]
                value = split_line[2]

                if action == "n_msgs":
                    self.change_icon(int(value))
                if action == "blink":
                    self.blink(int(value))

            elif component == "popup":
                info = split_line[1:]
                self.send_popup(*info)

            # Command to terminate process completely
            elif component == "quit":
                self.quit(parent_kill=False)

            line = stdin.readline().replace("\n", "")

        return True

    def quit(self, event=None, parent_kill=True):
        """
        Shuts down the icon; this will also attempt to kill the parent
        process (since when running as part of the checker; the checker is
        daemonised and the user does not interact with it directly

        """
        if parent_kill:
            subprocess.Popen(["kill", "{0:d}".format(self.parent)])
        self.blinker.blinking = False
        gtk.main_quit()

    def init_StatusIcon(self):
        """
        Creates the GTK status icon itself, and attach the left and right
        click actions to it

        """
        self.statusicon = gtk.StatusIcon()
        self.statusicon.set_from_file(
            os.path.join(OWA_CHECKER_PATH, "icons", "envelope.png"))
        self.statusicon.connect("popup-menu", self.open_menu)
        self.statusicon.connect("activate", self.go_to_desktop)

    def open_menu(self, icon, button, time):
        """
        The right click action for the status icon; launches a menu with
        various options

        """
        menu = gtk.Menu()
        menu.set_border_width(5)

        config = gtk.MenuItem("Configure")
        about = gtk.MenuItem("About")
        quit = gtk.MenuItem("Quit")

        config.connect("activate", self.show_config_dialog)
        about.connect("activate", self.show_about_dialog)
        quit.connect("activate", self.quit)

        silenced = gtk.CheckMenuItem("Silence Mail")
        silenced.connect("button-press-event", self.toggle_silenced)
        silenced.set_active(self.silenced)

        menu.append(config)
        menu.append(silenced)
        menu.append(about)
        menu.append(quit)

        menu.show_all()

        menu.popup(None, None,
                   gtk.status_icon_position_menu,
                   button, time, self.statusicon)

    def toggle_silenced(self, widget, data=None):
        """
        Controller for the "silence" switch; toggles the switch state and
        the controlling value

        """
        if widget.get_active():
            widget.set_active(False)
            self.silenced = False
        else:
            widget.set_active(True)
            self.silenced = True

    def show_about_dialog(self, widget):
        """
        Defines and shows the "about" dialog for the app

        """
        about_dialog = gtk.AboutDialog()

        about_dialog.set_destroy_with_parent(True)
        about_dialog.set_name("OWA Checker")
        about_dialog.set_version(ABOUT_DIALOG_VERSION)
        about_dialog.set_copyright(ABOUT_DIALOG_COPYRIGHT)
        about_dialog.set_comments(
            "A Replacement for the mail popups and calendar "
            "notifications normally found in Microsoft Outlook "
            "but missing from Outlook Web Access\n\n"
            "Written by Steve Wardle")

        pixbuf = gtk.gdk.pixbuf_new_from_file(
            os.path.join(OWA_CHECKER_PATH, "icons", "owa_logo.png"))
        sclbuf = pixbuf.scale_simple(100, 100, gtk.gdk.INTERP_BILINEAR)
        about_dialog.set_logo(sclbuf)
        about_dialog.set_icon(sclbuf)

        about_dialog.run()
        about_dialog.destroy()

    def show_config_dialog(self, widget):
        """
        Opens the configuration window

        """
        self.config_menu()

    def change_icon(self, n_msgs):
        """
        Update the icon to reflect a new number of unread messages

        """
        self.n_msgs_shown = n_msgs
        if n_msgs == 0:
            self.blinker.blinking = False
            self.statusicon.set_from_file(
                os.path.join(OWA_CHECKER_PATH, "icons", "envelope.png"))
        elif n_msgs <= 99:
            self.statusicon.set_from_file(
                os.path.join(OWA_CHECKER_PATH, "icons",
                             "envelope_i_{0:02d}.png".format(n_msgs)))
        else:
            self.statusicon.set_from_file(
                os.path.join(OWA_CHECKER_PATH, "icons", "envelope_i_99p.png"))

        self.statusicon.set_tooltip("{0:d} unread messages".format(n_msgs))

    def blink(self, state):
        """
        Toggle the Scroll Lock LED blinking

        """
        if state == 1:
            self.blinker.blinking = True
        if state == 0:
            self.blinker.blinking = False

    def go_to_desktop(self, event):
        """
        Runs the Wmctrl command to switch desktop

        """
        temp_env = os.environ.copy()
        temp_env["DISPLAY"] = self.config.mail_display
        switch = subprocess.Popen(self.wmctrl_command, env=temp_env)

        # De-activate the blinking LED if the user does this, assuming they
        # have acknowledged the new messages
        self.blinker.blinking = False

    def recalc_offset(self, position):
        """
        Work out the bounding co-ordinates of the desktop area, accounting
        for panels etc.

        """
        # Separate the positioning text and use it to define the co-ords
        self.position_v, self.position_h = position.split("-", 1)

        if self.position_v == "top":
            self.offset_y = (POPUP_GNOME3_TOPBAR_PADDING
                             + POPUP_SCREEN_EDGE_PADDING)
        else:
            self.offset_y = (POPUP_GNOME3_BOTTOMBAR_PADDING
                             + POPUP_SCREEN_EDGE_PADDING)

        if self.position_h == "left":
            self.offset_x = POPUP_SCREEN_EDGE_PADDING
        else:
            self.offset_x = POPUP_SCREEN_EDGE_PADDING

    def init_PopupManager(self, position="top-right"):
        """
        Defines the popup for mail messages - this is an expanding top-level
        window which adjusts its position on the fly to behave as intended

        """
        self.popup = gtk.Window(gtk.WINDOW_POPUP)
        self.popup.set_decorated(False)
        self.popup.set_keep_above(True)
        self.popup.set_opacity(self.config.mail_opacity)
        self.popup.move(1, 1)
        self.popup.set_border_width(0)
        self.recalc_offset(position)

        # This main VBox is the main body of the PopupManager, it is this that
        # will be repeatedly packed + unpacked to give the appearance of new
        # popups appearing and dissapearing
        self.master_vbox = gtk.VBox()

        # Since each popup isn't technically a separate window like it was in
        # notify-send, this frame makes it feel slightly more defined at the
        # edges
        outer_frame = gtk.Frame()
        outer_frame.set_shadow_type(gtk.SHADOW_IN)
        outer_frame.set_border_width(1)
        outer_frame.add(self.master_vbox)

        # Put the frame inside the PopupManager, and connect a method to the
        # configure event - this will get triggered whenever the size of the
        # manager changes
        self.popup.add(outer_frame)
        self.popup.connect("configure_event", self.configure_event)
        self.popup.show_all()
        self.popup.hide()

    def configure_event(self, event, something):
        """
        A configure event gets triggered in this case whenever the size of
        the popup gets changed (and we use this as a way of pinging it to
        adjust it's position accordingly)

        """
        x, y = self.popup.get_size()

        # If the PopupManager is at the bottom of the screen then the position
        # for it must move up and down as popups are issued and dismissed since
        # GTK windows and widgets always extend downwards when packing (and we
        # want the impression of it packing upwards instead)
        if self.position_v == "bottom":
            self.pos_y = SCREEN_Y - y - self.offset_y
        else:
            self.pos_y = self.offset_y

        # This just ensures that the popup is position at either edge of the
        # screen area
        if self.position_h == "right":
            self.pos_x = SCREEN_X - x - self.offset_x
        elif self.position_h == "middle-right":
            self.pos_x = int(SCREEN_X / 2.0) + self.offset_x
        elif self.position_h == "middle-left":
            self.pos_x = int(SCREEN_X / 2.0) - x - self.offset_x
        elif self.position_h == "left":
            self.pos_x = self.offset_x

        self.popup.move(self.pos_x, self.pos_y)

    def shrink(self):
        """
        This is a method of pinging the PopupManager in order to trigger
        the configure event (by attempting to resize it)

        """
        self.popup.resize(1, 1)

        # If there aren't any popups inside the Manager it should be hidden
        if len(self.master_vbox.get_children()) == 0:
            self.popup.hide()

    def send_popup(self, sender, title, uname=None, timeout=None):
        """
        Issue a popup given appropriate values for the sender and message
        title.  If specified a full username may be used to try and obtain
        a portrait for the popup to display.  The timeout controls how many
        seconds the popup should remain visible

        """
        # Note that the below method does not work on the VDI machines, as
        # they have the lock disabled, so only need to do this check on non
        # VDI machines
        if not HOSTNAME.startswith("vld"):
            saver_check = subprocess.Popen(
                ['gsettings', 'get', 'org.gnome.desktop.lockdown',
                 'disable-lock-screen'],
                stdout=subprocess.PIPE, stderr=subprocess.PIPE)

            stdout, stderr = saver_check.communicate()
            if "true" in stdout:
                return

        # If the user has activated the "silence mail" setting, they don't want
        # to be disturbed by popups
        if self.silenced:
            return

        # Just in case the supplied image is not readable, fall-back to the
        # default logo (which is also used if no image is supplied)
        pixbuf = None
        if self.config.showface and uname not in (None, "(No Address)"):
            username, domain = uname.split("@")
            if LOCAL_DOMAIN is not None and domain == LOCAL_DOMAIN:
                # Try to take it from the o365 api
                if pixbuf is None and self.o365:
                    try:
                        binary = o365_api.get_user_portrait(
                            oauth2.get_access_token(), uname)
                    except o365_api.o365Error:
                        binary = None

                    if binary is not None:
                        loader = gtk.gdk.PixbufLoader("jpeg")
                        loader.write(binary)
                        loader.close()
                        pixbuf = loader.get_pixbuf()

                # If that didn't work, try a .face file
                if pixbuf is None:
                    face_abbrev = "~{0:s}/.face".format(username)
                    face_expand = os.path.expanduser(face_abbrev)
                    if os.access(face_expand, os.R_OK):
                        pixbuf = gtk.gdk.pixbuf_new_from_file(face_expand)

        # And if nothing worked, use the default logo
        if pixbuf is None:
            pixbuf = gtk.gdk.pixbuf_new_from_file(
                os.path.join(OWA_CHECKER_PATH, "icons", "owa_logo.png"))

        if not timeout:
            timeout = int(self.config.timeoutms / 1000)
        else:
            timeout = int(timeout)

        sender = sender.strip()
        title = title.strip()

        # Format the icon, keep its height fixed and adjust width to preserve
        # its aspect ratio
        icon = gtk.Image()
        aspect_width = (
            int(round(float(
                pixbuf.get_width())
                / float(pixbuf.get_height()) * POPUP_AVATAR_SIZE)))
        if aspect_width > POPUP_AVATAR_SIZE:
            aspect_width = POPUP_AVATAR_SIZE
        sclbuf = pixbuf.scale_simple(
            aspect_width, POPUP_AVATAR_SIZE, gtk.gdk.INTERP_BILINEAR)
        icon.set_from_pixbuf(sclbuf)
        icon.set_size_request(POPUP_AVATAR_SIZE, POPUP_AVATAR_SIZE)

        # Label for the sender, bold, left aligned, fixed width
        sender_label = gtk.Label()
        sender = sender.replace("\\:", ":")
        sender = cgi.escape(sender)
        sender_label.set_markup("<b>{0:s}</b>".format(sender))
        sender_label.set_line_wrap(True)
        sender_label.set_alignment(0, 0.5)
        sender_label.set_size_request(POPUP_TEXT_WIDTH, -1)

        # Label for the message title, as above but not bold
        title_label = gtk.Label()
        title = title.replace("\\:", ":")
        title = cgi.escape(title)
        title_label.set_markup("{0:s}".format(title))
        title_label.set_line_wrap(True)
        title_label.set_alignment(0, 0.5)
        title_label.set_size_request(POPUP_TEXT_WIDTH, -1)

        # A container to group the message sender and title text
        text_vbox = gtk.VBox()
        text_vbox.pack_start(
            sender_label, expand=True, fill=True, padding=POPUP_TEXT_PADDING)
        text_vbox.pack_start(
            title_label, expand=True, fill=True, padding=POPUP_TEXT_PADDING)

        # A container to group the text box and icon
        body_hbox = gtk.HBox()
        body_hbox.pack_start(
            icon, expand=False, fill=True, padding=POPUP_AVATAR_PADDING)
        body_hbox.pack_start(
            gtk.VSeparator(), expand=False, fill=True)
        body_hbox.pack_start(
            text_vbox, expand=False, fill=True, padding=POPUP_AVATAR_PADDING)

        # A container to add the separator lines above and below
        outer_vbox = gtk.VBox()
        outer_vbox.pack_start(
            gtk.HSeparator(), expand=False, fill=True)
        outer_vbox.pack_start(
            body_hbox, expand=False, fill=True, padding=POPUP_SEPARATION)
        outer_vbox.pack_start(
            gtk.HSeparator(), expand=False, fill=True)

        # Finally an EventBox to contain the whole thing, this will allow
        # the popup to recieve actions when clicked on
        eventbox = gtk.EventBox()
        eventbox.add(outer_vbox)
        eventbox.show_all()

        # The packing direction depends on which way the popups should expand
        if self.position_v == "top":
            self.master_vbox.pack_start(eventbox)
        elif self.position_v == "bottom":
            self.master_vbox.pack_end(eventbox)

        # Make the popup manager visible if it was not already
        self.popup.show()

        # When the popup is removed, it must be destroyed and the manager must
        # be collapsed back down (and possibly hidden)  A lambda function won't
        # do here due to the number of arguments but here we define a bespoke
        # function for this newly created popup
        def destroy_and_shrink(*args, **kwargs):
            if len(args) > 0:
                args[0].destroy()
                self.shrink()
            elif "widget" in kwargs:
                kwargs["widget"].destroy()
                self.shrink()

        # Connect the above function so that it fires when the user clicks on
        # the popup, or when the timeout is reached (if it is set to a non-zero
        # value)
        eventbox.connect('button_press_event', destroy_and_shrink)
        if timeout > 0:
            gobject.timeout_add_seconds(timeout, destroy_and_shrink, eventbox)

    def conf_callback(self, widget, event=None, data=None):
        """
        Event handler for updates to the configuration window

        """
        # Checkboxes, just get the states
        if event == "check_face":
            self.config.showface = int(widget.get_active())
        elif event == "check_blink":
            self.config.blink = int(widget.get_active())
        elif event == "check_debug":
            self.config.debuglog = int(widget.get_active())

        elif event in [
                "entry_timeout", "entry_workspace", "entry_opacity"]:
            # Don't allow non-numerical entries in the text boxes
            # by insta-removing them when they get typed (it works!)
            text = widget.get_text().strip()
            widget.set_text(''.join([i for i in text if i in '0123456789']))
            text = widget.get_text().strip()
            if text != "":
                if event == "entry_timeout":
                    self.config.timeoutms = int(text) * 1000
                if event == "entry_opacity":
                    self.config.mail_opacity = int(text) / 100.0
                elif event == "entry_workspace":
                    self.config.workspace = int(text)
        elif event in ["entry_position"]:
            self.config.mail_pos = widget.get_active_text().strip()
        elif event in ["entry_mail_display"]:
            self.config.mail_display = widget.get_text().strip()
        elif event in ["entry_cal_display"]:
            self.config.cal_display = widget.get_text().strip()
        elif event in ["entry_folder_list"]:
            self.config.folder_list = widget.get_text().strip()

        # Either save or discard the changes depending on what the user
        # chooses
        elif event == "button_apply":
            if self.test_statusicon:
                self.test_statusicon.kill()
                self.blinker.turn_off_led()
            if self.test_cal_popup:
                self.test_cal_popup.kill()
            self.config.save_config()
            self.test_mail(reset=True)
            self.conf.destroy()

        elif event == "button_cancel" or data == "button_cancel":
            if self.test_statusicon:
                self.test_statusicon.kill()
                self.blinker.turn_off_led()
            if self.test_cal_popup:
                self.test_cal_popup.kill()
            self.config.read_config()
            self.test_mail(reset=True)
            self.conf.destroy()

    def config_menu(self):
        """
        Defines and shows the configuration dialog

        """
        self.conf = gtk.Window(gtk.WINDOW_TOPLEVEL)
        self.conf.set_position(gtk.WIN_POS_CENTER)
        self.conf.set_border_width(CONFIG_BORDER_WIDTH)
        self.conf.set_decorated(True)
        self.conf.set_keep_above(True)
        self.conf.set_title("OWA Checker Configuration")
        self.conf.set_icon_from_file(
            os.path.join(OWA_CHECKER_PATH, "icons", "owa_logo.png"))
        self.conf.connect("delete-event", self.conf_callback, "button_cancel")

        # Information describing the purpose of the dialog
        info_label_1 = gtk.Label()
        info_label_1.set_markup('<b>OWA Checker Configuration</b>')
        info_label_2 = gtk.Label()
        info_label_2.set_markup(
            '<i>Hover the mouse over headings for details</i>')
        info_vbox = gtk.VBox()
        info_vbox.pack_start(
            info_label_1, expand=True, fill=True,
            padding=CONFIG_OPTION_PADDING)
        info_vbox.pack_start(
            info_label_2, expand=True, fill=True,
            padding=CONFIG_OPTION_PADDING)

        # Create the icon to use in the dialog
        icon = gtk.Image()
        pixbuf = gtk.gdk.pixbuf_new_from_file(
            os.path.join(OWA_CHECKER_PATH, "icons", "owa_logo.png"))
        sclbuf = pixbuf.scale_simple(
            CONFIG_LOGO_SIZE, CONFIG_LOGO_SIZE, gtk.gdk.INTERP_BILINEAR)
        icon.set_from_pixbuf(sclbuf)

        # Pack the above into a container for the dialog header
        hbox_head = gtk.HBox()
        hbox_head.pack_start(icon, expand=True, fill=True,
                             padding=CONFIG_OPTION_PADDING)
        hbox_head.pack_start(info_vbox, expand=True, fill=True,
                             padding=CONFIG_OPTION_PADDING)

        # Create the heading for the mail popup section
        mail_info_label = gtk.Label()
        mail_info_label.set_markup('<b>Mail Popups</b>')
        button_test_mail = gtk.Button("Click here to test mail popup")
        self.test_statusicon = None
        self.test_mail_sent = 0
        button_test_mail.connect("clicked", self.test_mail)
        vbox_mail_info = gtk.VBox()
        vbox_mail_info.pack_start(
            mail_info_label, expand=True, fill=True,
            padding=CONFIG_OPTION_PADDING)
        vbox_mail_info.pack_start(
            button_test_mail, expand=True, fill=True,
            padding=CONFIG_OPTION_PADDING)

        # Checkbox for the faces option
        button_faces = gtk.CheckButton("Show faces")
        button_faces.connect("toggled", self.conf_callback, "check_face")
        button_faces.set_tooltip_markup(
            "Display the senders .face (if it exists)")
        button_faces.set_active((False, True)[self.config.showface])

        # Checkbox for the LED blink option
        button_blink = gtk.CheckButton("Blink LED")
        button_blink.connect("toggled", self.conf_callback, "check_blink")
        button_blink.set_tooltip_markup(
            "Blink the scroll-lock LED when you have unread emails waiting")
        button_blink.set_active((False, True)[self.config.blink])

        # Checkbox for the faces option
        button_debug = gtk.CheckButton("Debug Logging")
        button_debug.connect("toggled", self.conf_callback, "check_debug")
        button_debug.set_tooltip_markup(
            "Turn on full logging to logfile. <b>Note:</b> This "
            "only takes effect after restarting the checker")
        button_debug.set_active((False, True)[self.config.debuglog])

        # Pack the checkboxes side-by-side
        hbox_faceblink = gtk.HBox()
        hbox_faceblink.pack_start(
            button_faces, expand=True, fill=False,
            padding=CONFIG_OPTION_PADDING)
        hbox_faceblink.pack_start(
            button_blink, expand=True, fill=False,
            padding=CONFIG_OPTION_PADDING)
        hbox_faceblink.pack_start(
            button_debug, expand=True, fill=False,
            padding=CONFIG_OPTION_PADDING)

        # Entry for the mail popup timeout
        timeout_label = gtk.Label("Popup Timeout:")
        timeout_label.set_width_chars(CONFIG_OPTION_WIDTH)
        timeout_entry = gtk.Entry()
        timeout_entry.set_text(str(self.config.timeoutms / 1000))
        timeout_entry.connect("changed", self.conf_callback, "entry_timeout")
        tooltip = (
            "How long the email popup stays on the screen in seconds "
            "(0 = until clicked)")
        timeout_label.set_tooltip_markup(tooltip)
        timeout_entry.set_tooltip_markup(tooltip)

        # Entry for the mail popup position
        mail_pos_label = gtk.Label("Popup Position:")
        mail_pos_label.set_width_chars(CONFIG_OPTION_WIDTH)
        mail_pos_entry = gtk.combo_box_new_text()
        positions = ["top-right",
                     "top-left",
                     "bottom-right",
                     "bottom-left",
                     "top-middle-right",
                     "top-middle-left",
                     "bottom-middle-right",
                     "bottom-middle-left"]
        for pos in positions:
            mail_pos_entry.append_text(pos)
        mail_pos_entry.set_active(positions.index(self.config.mail_pos))
        mail_pos_entry.connect("changed", self.conf_callback, "entry_position")
        tooltip = "Where on the screen the mail popup appears"
        mail_pos_label.set_tooltip_markup(tooltip)
        mail_pos_entry.set_tooltip_markup(tooltip)

        # Entry for the mail popup opacity
        opacity_label = gtk.Label("Popup Opacity:")
        opacity_label.set_width_chars(CONFIG_OPTION_WIDTH)
        opacity_entry = gtk.Entry()
        opacity_entry.set_text(str(int(self.config.mail_opacity * 100)))
        opacity_entry.connect("changed", self.conf_callback, "entry_opacity")
        tooltip = (
            "How opaque/transparent the mail popup is (in % opacity)")
        opacity_label.set_tooltip_markup(tooltip)
        opacity_entry.set_tooltip_markup(tooltip)

        # Entry for the index of the OWA workspace
        workspace_label = gtk.Label("OWA Workspace:")
        workspace_label.set_width_chars(CONFIG_OPTION_WIDTH)
        workspace_entry = gtk.Entry()
        workspace_entry.set_text(str(self.config.workspace))
        workspace_entry.connect(
            "changed", self.conf_callback, "entry_workspace")
        tooltip = "Which workspace left-clicking on the envelope sends you to"
        workspace_label.set_tooltip_markup(tooltip)
        workspace_entry.set_tooltip_markup(tooltip)

        # Entry for the DISPLAY to use for mail popups
        mail_display_label = gtk.Label("OWA Display:")
        mail_display_label.set_width_chars(CONFIG_OPTION_WIDTH)
        mail_display_entry = gtk.Entry()
        mail_display_entry.set_text(str(self.config.mail_display))
        mail_display_entry.connect(
            "changed", self.conf_callback, "entry_mail_display")
        if N_SCREENS > 1:
            tooltip = "Which display OWA resides on (for desktop switching)"
        else:
            tooltip = "Not relevant when only using one X-server"
            mail_display_label.set_sensitive(False)
            mail_display_entry.set_sensitive(False)
        mail_display_label.set_tooltip_markup(tooltip)
        mail_display_entry.set_tooltip_markup(tooltip)

        # Entry for the list of folders to be checkerd
        folder_list_label = gtk.Label("Folder list:")
        folder_list_label.set_width_chars(CONFIG_OPTION_WIDTH)
        folder_list_entry = gtk.Entry()
        folder_list_entry.set_text(str(self.config.folder_list))
        folder_list_entry.connect(
            "changed", self.conf_callback, "entry_folder_list")
        tooltip = ("List of folder names to check for mail "
                   "(double-colon separated e.g. folder1::folder2)")
        folder_list_label.set_tooltip_markup(tooltip)
        folder_list_entry.set_tooltip_markup(tooltip)

        # Create the heading for the calendar reminder section
        cal_info_label = gtk.Label()
        cal_info_label.set_markup('<b>Calendar Reminders</b>')
        button_test_calendar = gtk.Button(
            "Click here to test calendar reminder")
        self.test_cal_popup = None
        button_test_calendar.connect("clicked", self.test_calendar)
        vbox_cal_info = gtk.VBox()
        vbox_cal_info.pack_start(
            cal_info_label, expand=True, fill=True,
            padding=CONFIG_OPTION_PADDING)
        vbox_cal_info.pack_start(
            button_test_calendar, expand=True, fill=True,
            padding=CONFIG_OPTION_PADDING)

        # Entry for the display to use for calendar popups
        cal_display_label = gtk.Label("Reminder Display:")
        cal_display_label.set_width_chars(CONFIG_OPTION_WIDTH)
        cal_display_entry = gtk.Entry()
        cal_display_entry.set_text(str(self.config.cal_display))
        cal_display_entry.connect(
            "changed", self.conf_callback, "entry_cal_display")
        if N_SCREENS > 1:
            tooltip = "Which display to use for the calendar popup"
        else:
            tooltip = "Not relevant when only using one X-server"
            cal_display_label.set_sensitive(False)
            cal_display_entry.set_sensitive(False)
        cal_display_label.set_tooltip_markup(tooltip)
        cal_display_entry.set_tooltip_markup(tooltip)

        # Apply and cancel buttons
        button_apply = gtk.Button("Apply")
        button_apply.connect("clicked", self.conf_callback, "button_apply")
        button_cancel = gtk.Button("Cancel")
        button_cancel.connect("clicked", self.conf_callback, "button_cancel")

        # Pack the buttons together
        hbox_confirm = gtk.HBox()
        hbox_confirm.pack_start(
            button_apply, expand=True, fill=True,
            padding=CONFIG_OPTION_PADDING)
        hbox_confirm.pack_start(
            button_cancel, expand=True, fill=True,
            padding=CONFIG_OPTION_PADDING)

        # Pack the mail labels
        vbox_mail_left = gtk.VBox()
        vbox_mail_left.pack_start(
            timeout_label, expand=True, fill=True,
            padding=CONFIG_OPTION_PADDING)
        vbox_mail_left.pack_start(
            mail_pos_label, expand=True, fill=True,
            padding=CONFIG_OPTION_PADDING)
        vbox_mail_left.pack_start(
            workspace_label, expand=True, fill=True,
            padding=CONFIG_OPTION_PADDING)
        vbox_mail_left.pack_start(
            mail_display_label, expand=True, fill=True,
            padding=CONFIG_OPTION_PADDING)
        vbox_mail_left.pack_start(
            opacity_label, expand=True, fill=True,
            padding=CONFIG_OPTION_PADDING)
        vbox_mail_left.pack_start(
            folder_list_label, expand=True, fill=True,
            padding=CONFIG_OPTION_PADDING)

        # Pack the calendar labels
        vbox_cal_left = gtk.VBox()
        vbox_cal_left.pack_start(
            cal_display_label, expand=True, fill=True,
            padding=CONFIG_OPTION_PADDING)

        # Pack the mail entry areas
        vbox_mail_right = gtk.VBox()
        vbox_mail_right.pack_start(
            timeout_entry, expand=True, fill=True,
            padding=CONFIG_OPTION_PADDING)
        vbox_mail_right.pack_start(
            mail_pos_entry, expand=True, fill=True,
            padding=CONFIG_OPTION_PADDING)
        vbox_mail_right.pack_start(
            workspace_entry, expand=True, fill=True,
            padding=CONFIG_OPTION_PADDING)
        vbox_mail_right.pack_start(
            mail_display_entry, expand=True, fill=True,
            padding=CONFIG_OPTION_PADDING)
        vbox_mail_right.pack_start(
            opacity_entry, expand=True, fill=True,
            padding=CONFIG_OPTION_PADDING)
        vbox_mail_right.pack_start(
            folder_list_entry, expand=True, fill=True,
            padding=CONFIG_OPTION_PADDING)

        # Pack the calendar entry areas
        vbox_cal_right = gtk.VBox()
        vbox_cal_right.pack_start(
            cal_display_entry, expand=True, fill=True,
            padding=CONFIG_OPTION_PADDING)

        # Combine the mail labels and entry areas
        hbox_mail = gtk.HBox()
        hbox_mail.pack_start(
            vbox_mail_left, expand=True, fill=True,
            padding=CONFIG_OPTION_PADDING)
        hbox_mail.pack_start(
            vbox_mail_right, expand=True, fill=True,
            padding=CONFIG_OPTION_PADDING)

        # Combine the calendar label and entry areas
        hbox_cal = gtk.HBox()
        hbox_cal.pack_start(
            vbox_cal_left, expand=True, fill=True,
            padding=CONFIG_OPTION_PADDING)
        hbox_cal.pack_start(
            vbox_cal_right, expand=True, fill=True,
            padding=CONFIG_OPTION_PADDING)

        # Put all of the above together for the final result
        vbox_master = gtk.VBox()
        vbox_master.pack_start(
            hbox_head, expand=True, fill=True,
            padding=CONFIG_OPTION_PADDING)
        vbox_master.pack_start(
            gtk.HSeparator(), expand=False, fill=True,
            padding=CONFIG_SECTION_SEPARATION)
        vbox_master.pack_start(
            vbox_mail_info, expand=True, fill=True,
            padding=CONFIG_OPTION_PADDING)
        vbox_master.pack_start(
            hbox_faceblink, expand=True, fill=True,
            padding=CONFIG_OPTION_PADDING)
        vbox_master.pack_start(
            hbox_mail, expand=True, fill=True,
            padding=CONFIG_OPTION_PADDING)
        vbox_master.pack_start(
            gtk.HSeparator(), expand=False, fill=True,
            padding=CONFIG_SECTION_SEPARATION)
        vbox_master.pack_start(
            vbox_cal_info, expand=True, fill=True,
            padding=CONFIG_OPTION_PADDING)
        vbox_master.pack_start(
            hbox_cal, expand=True, fill=True,
            padding=CONFIG_OPTION_PADDING)
        vbox_master.pack_start(
            gtk.HSeparator(), expand=False, fill=True,
            padding=CONFIG_SECTION_SEPARATION)
        vbox_master.pack_start(
            hbox_confirm, expand=True, fill=True,
            padding=CONFIG_OPTION_PADDING)

        # Add to the configuration window, also hide the maximize and minimize
        # buttons from the frame using the type hint and start the main loop
        self.conf.add(vbox_master)
        self.conf.set_type_hint(gtk.gdk.WINDOW_TYPE_HINT_MENU)
        self.conf.show_all()

    def test_calendar(self, event=None):
        """
        This may be called from the configuration menu in order to test
        out the calendar reminder - it just issues a reminder using the
        configuration options currently set

        """
        # Only allow one test instance at a time
        if self.test_cal_popup:
            self.test_cal_popup.kill()

        temp_env = os.environ.copy()
        temp_env["DISPLAY"] = self.config.cal_display
        temp_env["OWA_MEETING_TITLE"] = "OWA Checker Calendar Test"
        temp_env["OWA_MEETING_LOCATION"] = "Meeting Location"
        self.test_cal_popup = subprocess.Popen(
            ["python",
             os.path.join(OWA_CHECKER_PATH, "calendar_popup.py"),
             "15"], env=temp_env)

    def test_mail(self, event=None, reset=False):
        """
        This may be called from the configuration menu in order to test
        out the mail popups - it issues a popup using the configuration
        options currently set

        """
        self.set_wmctrl_command(self.config.workspace)
        self.recalc_offset(self.config.mail_pos)
        self.popup.set_opacity(float(self.config.mail_opacity))
        self.blinker.active = (False, True)[self.config.blink]

        if not reset:
            self.test_mail_sent += 1
            self.blink(1)
            self.change_icon(self.n_msgs_shown + 1)
            self.send_popup("OWA Checker",
                            "Test Message {0}".format(self.test_mail_sent),
                            timeout=int(self.config.timeoutms / 1000))
        else:
            self.blink(0)
            self.change_icon(self.n_msgs_shown - self.test_mail_sent)
            self.test_mail_sent = 0


if __name__ == "__main__":

    # Create a logger
    log_dir = os.path.join(CONFIG_FILE_DIR)
    if not os.path.exists(log_dir):
        os.mkdir(log_dir)

    logger = stream_logging.create_logger(
        os.path.join(log_dir, "status_icon.log"))

    icon = OWAStatusIcon(logger)
    gtk.main()
