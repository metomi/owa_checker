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
import sys
import gtk
import gobject
import subprocess


class ScrollBlink(object):
    """
    This provides a class which can control the scroll-lock LED, making
    it blink as an indication to the user

    """
    def __init__(self, blink_speed_ms=500):
        """
        Setup controlling variables and register the timed call to
        switch the LED state

        """
        self.active = True
        self.blinking = True
        self.led_on = True

        gobject.timeout_add(blink_speed_ms, self.blink)

    def blink(self):
        """
        Toggles the LED to turn either on or off depending on its
        current state

        """
        # If the process has been deactivated, this should make sure
        # the final call turns the led off
        if not self.active:
            self.blinking = False

        # If we don't want to blink
        if not self.blinking:
            # But the led is currently on
            if self.led_on:
                # Make sure it stops blinking in the "off" state
                self.turn_off_led()
                self.led_on = False

            # Once the led is off either wait for the next pass or
            # close the blinker if active has been set to False
            return self.active

        # Otherwise, toggle the LED
        if self.led_on:
            self.turn_off_led()
            self.led_on = False
        else:
            self.turn_on_led()
            self.led_on = True

        # Return True here (rather than self.active) since it protects
        # against the state changing while the led is on
        return True

    @staticmethod
    def turn_on_led():
        subprocess.Popen(["xset", "led", "named", "Scroll Lock"])

    @staticmethod
    def turn_off_led():
        subprocess.Popen(["xset", "-led", "named", "Scroll Lock"])


if __name__ == "__main__":
    # For testing purposes - this file can be run as a script and
    # will blink the LED at the speed passed to it
    speed = int(sys.argv[1])
    blinker = ScrollBlink(speed)
