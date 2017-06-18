import os, sys, inspect, thread, time
import win32api
import win32con
import time
from win32api import GetSystemMetrics

sys.path.append("../lib")
src_dir = os.path.dirname(inspect.getfile(inspect.currentframe()))
# Windows and Linux
arch_dir = '../lib/x64' if sys.maxsize > 2**32 else '../lib/x86'
sys.path.insert(0, os.path.abspath(os.path.join(src_dir, arch_dir)))
import Leap

pinch_power_threshold = 0.9
grab_power_threshold = 0.9

screen_width = GetSystemMetrics(0)
screen_height = GetSystemMetrics(1)

left_pressed = 0
right_pressed = 0

x_val = 0
y_val = 0

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

def rightCLick():
    global right_pressed
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN,x_val,y_val,0,0)
    right_pressed = 1
    time.sleep(0.05)

def rightRelease():
    global right_pressed
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP,x_val,y_val,0,0)
    right_pressed = 0
    time.sleep(0.05)

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

        # Get hands
        for hand in frame.hands:
            global x_val
            global y_val
            global left_pressed
            global right_pressed

            handType = "Left hand" if hand.is_left else "Right hand"

            #print " position: %s" % (hand.palm_position)
            print "x - %s" % (hand.palm_position[0])
            print "y - %s" % (-hand.palm_position[2])
            win32api.SetCursorPos((map(int(hand.palm_position[0]), -400, 400, 0, screen_width),
                                   map(int(hand.palm_position[2]), -400, 400, 0, screen_height)))
            x_val = map(int(hand.palm_position[0]), 0, 400, 0, screen_width)
            y_val = map(int(hand.palm_position[2]), 0, 400, 0, screen_height)

            if (hand.pinch_strength > pinch_power_threshold and left_pressed == 0):
                leftClick()

            if (hand.pinch_strength <= pinch_power_threshold and left_pressed == 1):
                leftRelease()

            if (hand.grab_strength > grab_power_threshold and right_pressed == 0):
                rightCLick()

            if (hand.grab_strength <= grab_power_threshold and right_pressed == 1):
                rightRelease()

        if not frame.hands.is_empty:
            print ""

        for gesture in frame.gestures():
            print "gesture dir: %d, gesture pos: %d, gesture speed: %d" % (gesture.direction, gesture.position, gesture.speed)
            if gesture.type is Leap.Gesture.TYPE_SWIPE:
                swipe = Leap.SwipeGesture(gesture)
                print "hello"


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
