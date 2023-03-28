import os
import subprocess
import sys
import time
from windows_tools.installed_software import get_installed_software
import wmi #get running programes
# importing the pygame library
import pygame
import pygame.camera
import pyttsx3
import speech_recognition as sr
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import cv2
import mediapipe as mp
from math import hypot
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import numpy as np



options = Options()
options.add_experimental_option('detach', True)


engine = pyttsx3.init()

def controlingSoundVolumeCV():
    cap = cv2.VideoCapture(0)  # Checks for camera

    mpHands = mp.solutions.hands  # detects hand/finger
    hands = mpHands.Hands()  # complete the initialization configuration of hands
    mpDraw = mp.solutions.drawing_utils

    # To access speaker through the library pycaw
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    volbar = 400
    volper = 0

    volMin, volMax = volume.GetVolumeRange()[:2]

    while True:
        success, img = cap.read()  # If camera works capture an image
        imgRGB = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)  # Convert to rgb

        # Collection of gesture information
        results = hands.process(imgRGB)  # completes the image processing.

        lmList = []  # empty list
        if results.multi_hand_landmarks:  # list of all hands detected.
            # By accessing the list, we can get the information of each hand's corresponding flag bit
            for handlandmark in results.multi_hand_landmarks:
                for id, lm in enumerate(handlandmark.landmark):  # adding counter and returning it
                    # Get finger joint points
                    h, w, _ = img.shape
                    cx, cy = int(lm.x * w), int(lm.y * h)
                    lmList.append([id, cx, cy])  # adding to the empty list 'lmList'
                mpDraw.draw_landmarks(img, handlandmark, mpHands.HAND_CONNECTIONS)

        if lmList != []:
            # getting the value at a point
            # x      #y
            x1, y1 = lmList[4][1], lmList[4][2]  # thumb
            x2, y2 = lmList[8][1], lmList[8][2]  # index finger
            # creating circle at the tips of thumb and index finger
            cv2.circle(img, (x1, y1), 13, (255, 0, 0), cv2.FILLED)  # image #fingers #radius #rgb
            cv2.circle(img, (x2, y2), 13, (255, 0, 0), cv2.FILLED)  # image #fingers #radius #rgb
            cv2.line(img, (x1, y1), (x2, y2), (255, 0, 0), 3)  # create a line b/w tips of index finger and thumb

            length = hypot(x2 - x1, y2 - y1)  # distance b/w tips using hypotenuse
            # from numpy we find our length,by converting hand range in terms of volume range ie b/w -63.5 to 0
            vol = np.interp(length, [30, 350], [volMin, volMax])
            volbar = np.interp(length, [30, 350], [400, 150])
            volper = np.interp(length, [30, 350], [0, 100])

            print(vol, int(length))
            volume.SetMasterVolumeLevel(vol, None)

            # Hand range 30 - 350
            # Volume range -63.5 - 0.0
            # creating volume bar for volume level
            cv2.rectangle(img, (50, 150), (85, 400), (0, 0, 255),
                          4)  # vid ,initial position ,ending position ,rgb ,thickness
            cv2.rectangle(img, (50, int(volbar)), (85, 400), (0, 0, 255), cv2.FILLED)
            cv2.putText(img, f"{int(volper)}%", (10, 40), cv2.FONT_ITALIC, 1, (0, 255, 98), 3)
            # tell the volume percentage ,location,font of text,length,rgb color,thickness
        cv2.imshow('Image', img)  # Show the video
        if cv2.waitKey(1) & 0xff == ord(' '):  # By using spacebar delay will stop
            break

    cap.release()  # stop cam
    cv2.destroyAllWindows()  # close window





#covert Text to speech

# Import the required module for text
# to speech conversion

# This module is imported so that we can
# play the converted audio

# The text that you want to convert to audio
txtbeforeLiten = 'Please say somthing'
txtafterLiten = 'sir you said'

# Language in which you want to convert
language = 'en'

# Passing the text and language to the engine,
# here we have marked slow=False. Which tells
# the module that the converted audio should

def takingImage():
    # initializing  the camera
    pygame.camera.init()
    # make the list of all available cameras
    camlist = pygame.camera.list_cameras()
    # if camera is detected or not
    if camlist:

        # initializing the cam variable with default camera
        cam = pygame.camera.Camera(camlist[0], (640, 480))

        # opening the camera
        cam.start()

        # capturing the single image
        image = cam.get_image()

        # saving the image
        pygame.image.save(image, "camImage.jpg")

        listenfun("Image taked do you want to show it")
    # if camera is not detected the moving to else part
    else:
        print("No camera on current device")

def showImage():
    # activate the pygame library .
    pygame.init()
    X = 600
    Y = 600

    # create the display surface object
    # of specific dimension..e(X, Y).
    scrn = pygame.display.set_mode((X, Y))

    # set the pygame window name
    pygame.display.set_caption('image')

    # create a surface object, image is drawn on it.
    imp = pygame.image.load("camImage.jpg").convert()

    # Using blit to copy content from one surface to other
    scrn.blit(imp, (0, 0))

    # paint screen one time
    pygame.display.flip()
    status = True
    while (status):

        # iterate over the list of Event objects
        # that was returned by pygame.event.get() method.
        for i in pygame.event.get():

            # if event object type is QUIT
            # then quitting the pygame
            # and program both.
            if i.type == pygame.QUIT:
                status = False

    # deactivates the pygame library
    pygame.quit()
    listenfun("Do you want any thing else ali basha")

def getSofwareLocation(softwareName):
    print(os.system('where is '+softwareName))

def closePrograme(name):
    # creating a forever loop
    while 1:
        subprocess.call("TASKKILL /F /IM "+name+".exe", shell=True)
        time.sleep(10)

def getRunningProgrames():
    # Initializing the wmi constructor
    f = wmi.WMI()

    # Printing the header for the later columns
    print("pid   Process name")

    # Iterating through all the running processes
    for process in f.Win32_Process():
        # Displaying the P_ID and P_Name of the process
        print(f"{process.ProcessId:<10} {process.Name}")

def getSystemInstalledSoftware(softwareName):
    softwareChosedName=" "
    for software in get_installed_software():
        # if softwareName in software["name"]:
        if (software["name"].find(softwareName)==0):
            engine.say("i found "+software["name"]+" do you want to open")
            engine.setProperty('rate', 0)
            engine.setProperty('volume', 0.2)
            engine.runAndWait()
            print(software["name"])
            print(type(software["name"]))
            reco=getrecognize()
            if(reco.find("yes")==0):
                print("done")
                softwareChosedName=software["name"]
                break
            # print(getSofwareLocation(software['name']))
    getSofwareLocation(softwareChosedName)

def getOreder(text):
    if "google" in text.lower():
        driver = webdriver.Chrome()
        print("opening google..")
        if "open" in text.lower():
            driver.get("https://www.google.com")
            listenfun("Do you want something else ")
            time.sleep(60)
        elif "close" in text.lower():
            driver.close()
            listenfun("Do you want something else ")
    elif "facebook" in text.lower():
        print("opening facebook..")
        driver = webdriver.Chrome()
        driver.get("https://www.facebook.com")
        listenfun("Do you want any thing else")
        time.sleep(60)

    elif "open notepad" in text.lower():
        print("opening notepad..")
        subprocess.Popen('C:\\Windows\\System32\\notepad.exe')
        listenfun("Do you want any thing else")
    elif "close notepad" in text.lower():
        print("closing notepad..")
        closePrograme("notepad")
        listenfun("Do you want any thing else")
    elif "image" in text.lower():
        print("Taking Image..")
        takingImage()
    elif "yes" in text.lower():
        print("showing Image..")
        showImage()
    elif "exit" in text.lower():
        print("Exit..")
        engine.say("ok goodbye")
        engine.setProperty('rate', 0)
        engine.setProperty('volume', 0.2)
        engine.runAndWait()
        sys.exit()
    elif "computer" in text.lower():
        print("opening sound volume level")
        say("ok you will be in control with sound level with your finger")
        controlingSoundVolumeCV()
    else:
        print("re-listen")
        listenfun("I do not understand you  say it again")


def listenfun(text):
    r=sr.Recognizer()
    with sr.Microphone() as source:
        print(text)
        say(text)
        r.adjust_for_ambient_noise(source)
        audio=r.listen(source)
        print("Recognizing Now .... ")

        try:
            text=r.recognize_google(audio)
            print("you said "+text.lower())
            getOreder(text)
        except Exception as e:
            print("Error : " +str(e))
            listenfun("I do not understand you  say it again")

def getrecognize():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.adjust_for_ambient_noise(source)
        audio = r.listen(source)
        print("Recognizing Now .... ")

        try:
            text = r.recognize_google(audio)
            print("you said " + text.lower())
        except Exception as e:
            print("Error : " + str(e))

    return text.lower();
def say(text):
    print("from saying method")
    engine.say(text)
    engine.setProperty('rate', 0)
    engine.setProperty('volume', 0.2)
    engine.runAndWait()

listenfun("what do you want")


