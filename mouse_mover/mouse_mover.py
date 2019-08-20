# This script moves the mouse pointer 8 pixels along X axis and back

import pyautogui
import time



# Time duration to delay the movement (given in seconds)
delayDuration = 5

# Movement distance (given in pixels)
movementDistance = 50


# Loop to move the mouse
while True:
    try:
        # Starting position of mouse
        startingPosition = pyautogui.position()
        currentPosition = pyautogui.position()
        
        if(currentPosition == startingPosition):
            pyautogui.moveRel(movementDistance, 0)
            newPosition = pyautogui.position()
        else:
            pyautogui.moveRel(-movementDistance, 0)
            newPosition = pyautogui.position()

        # Delay the loop for n seconds
        time.sleep(delayDuration)
    except Exception as e:
        print('ese')