import cv2
import numpy as np
import aspose.slides as slides
import aspose.pydrawing as drawing
from cvzone.HandTrackingModule import HandDetector

# Load the PowerPoint file using Aspose.Slides
pptx = slides.Presentation("C:\\PycharmProjects\\HandGestureRecognition\\Demo PPt.pptx")

# Parameters
width, height = 1080, 720

# Camera Setup
cap = cv2.VideoCapture(0)
cap.set(3, width)
cap.set(4, height)

# Hand Detector
detectorHand = HandDetector(detectionCon=0.5, maxHands=1)

# Variables
gestureThreshold = 400
drawMode = False
annotationPoints = []
currentSlide = 0

while True:
    # Get image frame from webcam
    success, img = cap.read()
    img = cv2.flip(img, 1)

    # Find the hand and its landmarks
    hands, img = detectorHand.findHands(img)  # with draw

    if hands:
        hand = hands[0]
        fingers = detectorHand.fingersUp(hand)  # List of which fingers are up
        lmList = hand["lmList"]  # List of 21 Landmark points
        cx, cy = hand["center"]

        # Gesture for drawing (index and thumb pinching)
        if fingers == [0, 1, 1, 0, 0]:
            drawMode = True
            indexFinger = lmList[8]
            annotationPoints.append(indexFinger)  # Store points for drawing

            # Draw on the image as feedback
            cv2.circle(img, (indexFinger[0], indexFinger[1]), 12, (0, 0, 255), cv2.FILLED)
        else:
            if drawMode and len(annotationPoints) > 1:
                # Drawing on PowerPoint Slide
                slide = pptx.slides[currentSlide]
                shape = slide.shapes.add_auto_shape(
                    slides.ShapeType.LINE,
                    annotationPoints[0][0], annotationPoints[0][1],
                    annotationPoints[-1][0], annotationPoints[-1][1]
                )
                shape.line_format.width = 5
                shape.line_format.fill_format.solid_fill_color.color = drawing.Color.red

                # Reset annotation points after drawing
                annotationPoints = []
                drawMode = False

    # Show the webcam image with detection feedback
    cv2.imshow("Image", img)

    key = cv2.waitKey(1)
    if key == ord('q'):
        break

# Save the modified presentation with annotations
pptx.save("C:\\PycharmProjects\\HandGestureRecognition\\Annotated_Presentation.pptx")
