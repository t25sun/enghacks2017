import speech_recognition as sr
import os
def active_listen():
    while (True):
        r = sr.Recognizer()
        with sr.Microphone() as src:
         	audio = r.listen(src)
        msg = ''
        try:
            msg = r.recognize_google(audio)
    	    print msg.lower()
        except sr.UnknownValueError:
            print "Google Speech Recognition could not understand audio"
            pass
        except sr.RequestError as e:
            print "Could not request results from Google STT; {0}".format(e)
            pass
        except:
            print "Unknown exception occurred!"
            pass

active_listen()
