import pyautogui
import time

pyautogui.PAUSE = 0.3
time.sleep(5)
l = pyautogui.position()
print(l)