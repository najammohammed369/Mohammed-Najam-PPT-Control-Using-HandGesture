# This app will use your built-in webcam to control your slide's presentation.
# For a one-handed presentation, use Gesture 1 (thumbs up) to go to the previous slide
# and Gesture 2 (whole hand pointing up) to go to the next slide.

import win32com.client
from cvzone.HandTrackingModule import HandDetector
import cv2
import numpy as np
import aspose.slides
import aspose.pydrawing

Application = win32com.client.Dispatch("PowerPoint.Application")
Presentation = Application.Presentations.Open("C:\\PycharmProjects\\HandGestureRecognition\\Demo PPt.pptx")
print(Presentation.Name)
Presentation.SlideShowSettings.Run()

# Parameters
width, height = 1080, 720

# Camera Setup
cap = cv2.VideoCapture(0)
cap.set(3, width)
cap.set(4, height)

# Hand Detector
detectorHand = HandDetector(detectionCon=0.5, maxHands=1)

# Variables
imgList = []
hs, ws = int(180 * 1), int(270 * 1)
gestureThreshold = 400
delay = 30
buttonPressed = False
buttonCounter = 0
counter = 0
drawMode = False
imgNumber = 20
delayCounter = 0
annotations = [[]]
annotationNumber = 0
annotationStart = False

while True:
    # Get image frame
    success, img = cap.read()
    img = cv2.flip(img, 1)

    # Find the hand and its landmarks
    hands, img = detectorHand.findHands(img)  # with draw
    horizontalLine = cv2.line(img, (gestureThreshold, gestureThreshold), (width, gestureThreshold),
                              (255, 0, 0), 10)
    verticalLine = cv2.line(img, (gestureThreshold, 0), (gestureThreshold, gestureThreshold),
                            (255, 0, 0), 10)
    print(annotationNumber)

    if hands and buttonPressed is False:  # If hand is detected
        hand = hands[0]
        fingers = detectorHand.fingersUp(hand)  # List of which fingers are up
        cx, cy = hand["center"]
        lmList = hand["lmList"]  # List of 21 Landmark points

        xVal = int(np.interp(lmList[8][0], [width // 2, width - 200], [0, width - 200]))
        yVal = int(np.interp(lmList[8][1], [150, height - 200], [0, height - 200]))
        indexFinger = xVal, yVal

        if cy <= gestureThreshold:  # If hand is at the height of the face
            if cx >= gestureThreshold:

                if fingers == [0, 0, 0, 0, 1]:  # first gesture to move on to next page
                    print("Next")
                    annotationStart = False
                    if imgNumber > 0:
                        buttonPressed = True
                        Presentation.SlideShowWindow.View.Next()
                        imgNumber += 1
                        annotations = [[]]
                        annotationNumber = 0
                    else:
                        print("this is the last page")

                if fingers == [1, 0, 0, 0, 0]:  # Second gesture to move on to previous page
                    print("Previous")
                    annotationStart = False
                    if imgNumber > 0:
                        buttonPressed = True
                        Presentation.SlideShowWindow.View.Previous()
                        imgNumber -= 1
                        annotations = [[]]
                        annotationNumber = -1
                    else:
                        print("this is the first page")

            if fingers == [0, 1, 1, 0, 0]:
                print("Highlight")
                cv2.circle(img, indexFinger, 12, (0, 0, 255), cv2.FILLED)

    else:
        annotationStart = False

    if buttonPressed:
        counter += 1
        if counter > delay:
            counter = 0
            buttonPressed = False

    for i, annotation in enumerate(annotations):
        for j in range(len(annotation)):
            if j != 0:
                cv2.line(img, annotation[j - 1], annotation[j], (0, 0, 200), 12)

    cv2.imshow("Image", img)

    key = cv2.waitKey(1)
    if key == ord('q'):
        break
