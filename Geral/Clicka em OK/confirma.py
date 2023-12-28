from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img
from PIL import ImageGrab
from functools import partial
ImageGrab.grab = partial(ImageGrab.grab, all_screens=True)

def run():

    while True:
        if _find_img('ok1.png', conf=0.9):
            _click_img('ok1.png', conf=0.9)
        if _find_img('ok2.png', conf=0.9):
            _click_img('ok2.png', conf=0.9)

if __name__ == '__main__':
    run()