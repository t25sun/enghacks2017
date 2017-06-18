import os, sys, inspect, thread, time
import win32api
from win32api import GetSystemMetrics
import win32con
from win32con import *
import pyperclip
import time
import speech_recognition as sr

sys.path.append("../lib")
src_dir = os.path.dirname(inspect.getfile(inspect.currentframe()))
arch_dir = '../lib/x64' if sys.maxsize > 2**32 else '../lib/x86'
sys.path.insert(0, os.path.abspath(os.path.join(src_dir, arch_dir)))

import Leap

pinch_power_threshold = 0.95
grab_power_threshold = 1

#Computer screen dimensions
screen_width = GetSystemMetrics(0)
screen_height = GetSystemMetrics(1)

#Leap control boundaries (measured from origin)
scale = 0.1
x_bound = int(screen_width * scale)
y_bound = int(screen_height * scale)

#Array for calculating hand position moving averages
arr_len = 20
x_array = [0] * arr_len
y_array = [0] * arr_len

#Left/right click parameters
left_pressed = 0
right_pressed = 0

#Cursor position
x_val = 0
y_val = 0

#Velocity parameters
last_vel = 0
vel = 0

#Scroll enable parameters
init_scroll_pos_y = 0
enable_scroll = 0

#Yaw tolerance
yaw_tol = -0.2

#win32api keyboard codes
VK_CODE = {'backspace':0x08, 'enter':0x0D, 'ctrl':0x11, 'v':0x56}

#Functions to activate key actions    
def press(key):
    '''
    one press, one release.
    accepts as many arguments as you want. e.g. press('left_arrow', 'a','b').
    '''
    win32api.keybd_event(VK_CODE[key], 0,0,0)
    time.sleep(.05)
    win32api.keybd_event(VK_CODE[key],0 ,win32con.KEYEVENTF_KEYUP ,0)

def pressAndHold(key):
    '''
    press and hold. Do NOT release.
    accepts as many arguments as you want.
    e.g. pressAndHold('left_arrow', 'a','b').
    '''
    win32api.keybd_event(VK_CODE[key], 0,0,0)
    time.sleep(.05)
    
def release(key):
    '''
    release depressed keys
    accepts as many arguments as you want.
    e.g. release('left_arrow', 'a','b').
    '''
    win32api.keybd_event(VK_CODE[key],0 ,win32con.KEYEVENTF_KEYUP ,0)

	
def active_listen():
    '''Recogizes speech, copies it to the clipboard and pastes in the current
    cursor location'''
    r = sr.Recognizer()
    with sr.Microphone() as src:
	audio = r.listen(src)
    msg = ''
    
    try:
	msg = r.recognize_google(audio)
	print(msg.lower())
	pyperclip.copy(msg.lower())
	pressAndHold('ctrl')
	press('v')
	release('ctrl')
	
    except sr.UnknownValueError:
	print("Google Speech Recognition could not understand audio")
	pass
    except sr.RequestError as e:
	print("Could not request results from Google STT; {0}".format(e))
	pass
    except:
	print("Unknown exception occurred!")
	pass
	
def mov_average(array, new_value):
    '''Calculates moving average of the last arr_len data values'''
    global arr_len
    del array[0]
    array.append(new_value)
    return sum(array)/arr_len

def map(x, in_min, in_max, out_min, out_max):
    '''Mapping function'''
    return (x - in_min) * (out_max - out_min) / (in_max - in_min) + out_min;

#Cursor clicking functions
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

#Scrolling functions    
def scrollUp():
    #Scroll one up
    for i in range(80):
	win32api.mouse_event(MOUSEEVENTF_WHEEL, x_val, y_val, -1, 0)

def scrollDown():
    #Scroll one down
    for i in range(80):
	win32api.mouse_event(MOUSEEVENTF_WHEEL, x_val, y_val, 1, 0)    

class SampleListener(Leap.Listener):
    '''Function that polls data from Leap and activates corresponding 
    functions'''
    finger_names = ['Thumb', 'Index', 'Middle', 'Ring', 'Pinky']
    bone_names = ['Metacarpal', 'Proximal', 'Intermediate', 'Distal']

    def on_init(self, controller):
        print("Initialized")

    def on_connect(self, controller):
        print("Connected")

    def on_disconnect(self, controller):
        # Note: not dispatched when running in a debugger.
        print("Disconnected")

    def on_exit(self, controller):
        print("Exited")

    def on_frame(self, controller):
        # Get the most recent frame and report some basic information
        frame = controller.frame()
        
        # Give control back to mouse if no hand is detected
        if not frame.hands:
            return
	
	if len(frame.hands) == 2:
	    leftHand = frame.hands[0]
	    rightHand = frame.hands[1]
	    if (leftHand.grab_strength >= grab_power_threshold and
	        rightHand.grab_strength >= grab_power_threshold):
		active_listen()

        # Get hands
        hand = frame.hands[0]
        global x_val, y_val
	global vel, last_vel
        global left_pressed, right_pressed
	global enable_scroll, init_scroll_pos_y	
        
        # Map hand position to cursor position
        new_x = map(int(hand.palm_position[0]), -x_bound, x_bound, \
                    0, screen_width)
        x_val = mov_average(x_array, new_x)
        
        new_y = map(int(hand.palm_position[2]), -y_bound, y_bound, \
                    0, screen_height)
        y_val = mov_average(y_array, new_y)
        
        win32api.SetCursorPos((x_val, y_val))
 
        last_vel = vel
	vel = hand.palm_velocity[0]
        
        if (hand.pinch_strength > pinch_power_threshold and left_pressed == 0):
            leftClick()

        if (hand.pinch_strength <= pinch_power_threshold and left_pressed == 1):
            leftRelease()

	if (hand.grab_strength >= grab_power_threshold and enable_scroll == 0):
	    init_scroll_pos_y = y_val
	    enable_scroll = 1

	if (hand.grab_strength < grab_power_threshold):
	    init_scroll_pos_y = 0
	    enable_scroll = 0

	if (enable_scroll == 1):
	    if ( (y_val - init_scroll_pos_y) > 50 ):
		scrollDown()
		init_scroll_pos_y = y_val
	    if ( (y_val - init_scroll_pos_y) < -50 ):
		scrollUp()
		init_scroll_pos_y = y_val
    
	if ((hand.direction[2] > yaw_tol) and right_pressed == 0 
	    and (hand.pinch_strength >= pinch_power_threshold)):
	    rightClick()
	
	if ((hand.direction[2] > yaw_tol) and right_pressed == 1
            and (hand.pinch_strength < pinch_power_threshold)):
	    rightRelease()

        if (vel > 1000 and last_vel < 1000):
	    press('enter')
	    
	if (vel < -1000 and last_vel > -1000):
	    press('backspace')	

def main():
    # Create a sample listener and controller
    listener = SampleListener()
    controller = Leap.Controller()

    # Have the sample listener receive events from the controller
    controller.add_listener(listener)
    
    # Keep this process running until Enter is pressed
    print("Press Enter to quit...")
    try:
        sys.stdin.readline()
    except KeyboardInterrupt:
        pass
    finally:
        # Remove the sample listener when done
        controller.remove_listener(listener)


if __name__ == "__main__":
    main()