import os, sys, inspect, thread, time
import win32api
import win32con
from win32con import *
import time
from win32api import GetSystemMetrics

sys.path.append("../lib")
src_dir = os.path.dirname(inspect.getfile(inspect.currentframe()))
arch_dir = '../lib/x64' if sys.maxsize > 2**32 else '../lib/x86'
sys.path.insert(0, os.path.abspath(os.path.join(src_dir, arch_dir)))

import Leap

pinch_power_threshold = 1
grab_power_threshold = 1

# Computer screen dimensions
screen_width = GetSystemMetrics(0)
screen_height = GetSystemMetrics(1)

# Leap control boundaries (measured from origin)
scale = 0.1
x_bound = int(screen_width * scale)
y_bound = int(screen_height * scale)

# Array for calculating hand position moving averages
arr_len = 20
x_array = [0] * arr_len
y_array = [0] * arr_len

left_pressed = 0
right_pressed = 0

x_val = 0
y_val = 0

init_scroll_pos_y = 0
enable_scroll = 0

thumb_detect = 0
thumb_action = 0

yaw_tol = -0.2

def mov_average(array, new_value):
    global arr_len
    del array[0]
    array.append(new_value)
    return sum(array)/arr_len

def map(x, in_min, in_max, out_min, out_max):
    return (x - in_min) * (out_max - out_min) / (in_max - in_min) + out_min;

def leftClick():
    global left_pressed
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,x_val,y_val,0,0)
    left_pressed = 1
    time.sleep(0.05)

def leftRelease():
    global left_pressed
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,x_val,y_val,0,0)
    left_pressed = 0
    time.sleep(0.05)

def rightClick():
    global right_pressed
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN,x_val,y_val,0,0)
    right_pressed = 1
    time.sleep(0.05)

def rightRelease():
    global right_pressed
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP,x_val,y_val,0,0)
    right_pressed = 0
    time.sleep(0.05)

def scrollUp():
    #Scroll one up
    win32api.mouse_event(MOUSEEVENTF_WHEEL, x_val, y_val, 1, 0)

def scrollDown():
    #Scroll one down
    win32api.mouse_event(MOUSEEVENTF_WHEEL, x_val, y_val, -1, 0)

class SampleListener(Leap.Listener):
    finger_names = ['Thumb', 'Index', 'Middle', 'Ring', 'Pinky']
    bone_names = ['Metacarpal', 'Proximal', 'Intermediate', 'Distal']

    def on_init(self, controller):
        print "Initialized"

    def on_connect(self, controller):
        print "Connected"

    def on_disconnect(self, controller):
        # Note: not dispatched when running in a debugger.
        print "Disconnected"

    def on_exit(self, controller):
        print "Exited"

    def on_frame(self, controller):
        # Get the most recent frame and report some basic information
        frame = controller.frame()

        # Give control back to mouse if no hand is detected
        if not frame.hands:
            return

        # Get hands
        hand = frame.hands[0]
        global x_val
        global y_val
        global left_pressed
        global right_pressed

        global enable_scroll
        global init_scroll_pos_y

        global thumb_detect
        global thumb_action

        global grab_strength

        # Map hand position to cursor position
        new_x = map(int(hand.palm_position[0]), -x_bound, x_bound, \
                    0, screen_width)
        x_val = mov_average(x_array, new_x)

        new_y = map(int(hand.palm_position[2]), -y_bound, y_bound, \
                    0, screen_height)
        y_val = mov_average(y_array, new_y)

        win32api.SetCursorPos((x_val, y_val))

        #print hand.direction[2]
        if (hand.pinch_strength >= pinch_power_threshold and left_pressed == 0 and
            hand.direction[2] < yaw_tol):
            leftClick()

        if (hand.pinch_strength < pinch_power_threshold and left_pressed == 1 and
            hand.direction[2] < yaw_tol):
            leftRelease()

        print hand.pinch_strength
        if ( (hand.direction[2] > yaw_tol) and right_pressed == 0
            and (hand.pinch_strength >= pinch_power_threshold)):
            rightClick()

        if ( hand.direction[2] > yaw_tol and right_pressed == 1
            and (hand.pinch_strength < pinch_power_threshold)):
            rightRelease()


        if (hand.grab_strength >= grab_power_threshold and enable_scroll == 0):
            global enable_scroll
            init_scroll_pos_y = y_val
            enable_scroll = 1

        if (hand.grab_strength < grab_power_threshold):
            global enable_scroll
            init_scroll_pos_y = 0
            enable_scroll = 0

        if (enable_scroll == 1):
            if ( (y_val - init_scroll_pos_y) > 75 ):
                scrollDown()
                init_scroll_pos_y = y_val
            if ( (y_val - init_scroll_pos_y) < -75 ):
                scrollUp()
                init_scroll_pos_y = y_val


def main():
    # Create a sample listener and controller
    listener = SampleListener()
    controller = Leap.Controller()

    # Have the sample listener receive events from the controller
    controller.add_listener(listener)

    # Keep this process running until Enter is pressed
    print "Press Enter to quit..."
    try:
        sys.stdin.readline()
    except KeyboardInterrupt:
        pass
    finally:
        # Remove the sample listener when done
        controller.remove_listener(listener)


if __name__ == "__main__":
    main()
