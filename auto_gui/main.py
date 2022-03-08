# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

"""运行之后，随便点击任意可输入文本的窗口
    都能输入指定 的内容"""


import time
import random
import pyautogui
import pyperclip as pc

# Press the green button in the gutter to run the script.
if __name__ == '__main__':

    # """英文输出"""
    # with open('tell.txt', 'r', encoding='u8') as f:
    #     lines = f.readlines()  # 每一行都存到lines
    # while True:
    #     pyautogui.typewrite(random.choice(lines))
    #     pyautogui.press('enter')
    #     time.sleep(random.randint(1,3))

    """中文输出"""
    with open('tell_c.txt', 'r', encoding='u8') as f:
        lines = f.readlines()  # 每一行都存到lines
    while True:
        # pyauto 不能识别中文
        # 换用 ctl c ,v 复制粘贴的方法
        # pyautogui.typewrite(random.choice(lines))
        x = random.choice(lines)
        pc.copy(x)
        pyautogui.hotkey('ctrl','v')
        pyautogui.press('enter')
        time.sleep(random.randint(1,3))


# See PyCharm help at https://www.jetbrains.com/help/pycharm/
