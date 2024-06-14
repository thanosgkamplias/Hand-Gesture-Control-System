import cv2
import time
import math
import numpy as np
import pyautogui
import autopy
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import mediapipe as mp

class handDetector():
    def __init__(self, mode=False, maxHands=2, detectionCon=0.5, trackCon=0.5):
        self.mode = mode
        self.maxHands = maxHands
        self.detectionCon = detectionCon
        self.trackCon = trackCon

        self.mpHands = mp.solutions.hands
        self.hands = self.mpHands.Hands(self.mode, self.maxHands, self.detectionCon, self.trackCon)
        self.mpDraw = mp.solutions.drawing_utils

    def findHands(self, img, draw=True):
        imgRGB = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
        self.results = self.hands.process(imgRGB)

        if self.results.multi_hand_landmarks:
            for handLms in self.results.multi_hand_landmarks:
                if draw:
                    self.mpDraw.draw_landmarks(img, handLms, self.mpHands.HAND_CONNECTIONS)
        return img

    def findPosition(self, img, handNo=0, draw=True, color=(255, 0, 255), z_axis=False):
        lmList = []
        if self.results.multi_hand_landmarks:
            myHand = self.results.multi_hand_landmarks[handNo]
            for id, lm in enumerate(myHand.landmark):
                h, w, c = img.shape
                if z_axis == False:
                    cx, cy = int(lm.x * w), int(lm.y * h)
                    lmList.append([id, cx, cy])
                elif z_axis:
                    cx, cy, cz = int(lm.x * w), int(lm.y * h), round(lm.z, 3)
                    lmList.append([id, cx, cy, cz])

                if draw:
                    cv2.circle(img, (cx, cy), 5, color, cv2.FILLED)

        return lmList

def putText(img, mode, loc=(250, 450), color=(0, 255, 255)):
    cv2.putText(img, str(mode), loc, cv2.FONT_HERSHEY_COMPLEX_SMALL, 3, color, 3)

def main():
    wCam, hCam = 640, 480  # Reduced resolution
    cap = cv2.VideoCapture(0)
    cap.set(3, wCam)
    cap.set(4, hCam)
    cap.set(5, 60)

    actual_wCam = cap.get(3)
    actual_hCam = cap.get(4)
    pTime = 0

    detector = handDetector(maxHands=1, detectionCon=0.85, trackCon=0.8)

    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    volRange = volume.GetVolumeRange()

    minVol = -63
    maxVol = volRange[1]
    hmin = 50
    hmax = 200
    volBar = 400
    volPer = 0
    vol = 0
    color = (0, 215, 255)
    tipIds = [4, 8, 12, 16, 20]
    mode = ''
    active = 0

    pyautogui.FAILSAFE = True

    try:
        while True:
            success, img = cap.read()
            img = detector.findHands(img)
            lmList = detector.findPosition(img, z_axis=True, draw=False)

            fingers = []

            if len(lmList) != 0:
                if lmList[tipIds[0]][1] > lmList[tipIds[0 - 1]][1]:
                    if lmList[tipIds[0]][1] >= lmList[tipIds[0] - 1][1]:
                        fingers.append(1)
                    else:
                        fingers.append(0)
                elif lmList[tipIds[0]][1] < lmList[tipIds[0 - 1]][1]:
                    if lmList[tipIds[0]][1] <= lmList[tipIds[0] - 1][1]:
                        fingers.append(1)
                    else:
                        fingers.append(0)

                for id in range(1, 5):
                    if lmList[tipIds[id]][2] < lmList[tipIds[id] - 2][2]:
                        fingers.append(1)
                    else:
                        fingers.append(0)

                if (fingers == [0, 0, 0, 0, 0]) and (active == 0):
                    mode = 'N'
                elif (fingers == [0, 1, 0, 0, 0] or fingers == [0, 1, 1, 0, 0]) and (active == 0):
                    mode = 'Scroll'
                    active = 1
                elif (fingers == [1, 1, 0, 0, 0]) and (active == 0):
                    mode = 'Volume'
                    active = 1
                elif (fingers == [1, 1, 1, 1, 1]) and (active == 0):
                    mode = 'Cursor'
                    active = 1

            if mode == 'Scroll':
                active = 1
                cv2.rectangle(img, (200, 410), (245, 460), (255, 255, 255), cv2.FILLED)
                if len(lmList) != 0:
                    if fingers == [0, 1, 0, 0, 0]:
                        putText(img, mode='U', loc=(200, 455), color=(0, 255, 0))  # Pass img here
                        pyautogui.scroll(300)
                    if fingers == [0, 1, 1, 0, 0]:
                        putText(img, mode='D', loc=(200, 455), color=(0, 0, 255))  # Pass img here
                        pyautogui.scroll(-300)
                    elif fingers == [0, 0, 0, 0, 0]:
                        active = 0
                        mode = 'N'

            if mode == 'Volume':
                active = 1
                if len(lmList) != 0:
                    if fingers[-1] == 1:
                        active = 0
                        mode = 'N'
                    else:
                        x1, y1 = lmList[4][1], lmList[4][2]
                        x2, y2 = lmList[8][1], lmList[8][2]
                        cx, cy = (x1 + x2) // 2, (y1 + y2) // 2
                        cv2.circle(img, (x1, y1), 10, color, cv2.FILLED)
                        cv2.circle(img, (x2, y2), 10, color, cv2.FILLED)
                        cv2.line(img, (x1, y1), (x2, y2), color, 3)
                        cv2.circle(img, (cx, cy), 8, color, cv2.FILLED)

                        length = math.hypot(x2 - x1, y2 - y1)

                        vol = np.interp(length, [hmin, hmax], [minVol, maxVol])
                        volBar = np.interp(vol, [minVol, maxVol], [400, 150])
                        volPer = np.interp(vol, [minVol, maxVol], [0, 100])

                        volume.SetMasterVolumeLevel(vol, None)

                        if length < 50:
                            cv2.circle(img, (cx, cy), 11, (0, 0, 255), cv2.FILLED)

                        cv2.rectangle(img, (30, 150), (55, 400), (209, 206, 0), 3)
                        cv2.rectangle(img, (30, int(volBar)), (55, 400), (215, 255, 127), cv2.FILLED)
                        cv2.putText(img, f'{int(volPer)}%', (25, 430), cv2.FONT_HERSHEY_COMPLEX, 0.9, (209, 206, 0), 3)

            if mode == 'Cursor':
                active = 1
                cv2.rectangle(img, (110, 20), (620, 350), (255, 255, 255), 3)

                if fingers[1:] == [0, 0, 0, 0]:
                    active = 0
                    mode = 'N'
                else:
                    if len(lmList) != 0:
                        x1, y1 = lmList[8][1], lmList[8][2]
                        w, h = autopy.screen.size()
                        X = int(np.interp(x1, [110, 620], [0, w - 1]))
                        Y = int(np.interp(y1, [20, 350], [0, h - 1]))
                        cv2.circle(img, (lmList[8][1], lmList[8][2]), 7, (255, 255, 255), cv2.FILLED)
                        cv2.circle(img, (lmList[4][1], lmList[4][2]), 10, (0, 255, 0), cv2.FILLED)

                        if X % 2 != 0:
                            X = X - X % 2
                        if Y % 2 != 0:
                            Y = Y - Y % 2

                        autopy.mouse.move(X, Y)

                        if fingers[0] == 0:
                            cv2.circle(img, (lmList[4][1], lmList[4][2]), 10, (0, 0, 255), cv2.FILLED)
                            pyautogui.click()

            cTime = time.time()
            fps = 1 / ((cTime + 0.01) - pTime)
            pTime = cTime

            cv2.putText(img, f'FPS:{int(fps)}', (480, 50), cv2.FONT_ITALIC, 1, (255, 0, 0), 2)
            cv2.imshow('Hand LiveFeed', img)

            if cv2.waitKey(1) & 0xFF == ord('q'):
                break
    except Exception as e:
        print(f"An error occurred: {str(e)}")
    finally:
        cap.release()
        cv2.destroyAllWindows()

if __name__ == "__main__":
    main()