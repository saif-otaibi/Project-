# import os

# pid = os.getpid()
# print(pid)

# os.system("taskkill /f /PID 5792")

# os.system("taskkill /f /im WINWORD.EXE")

# =============================================================================
import os
import hashlib
import time

with open('C:\\Users\\jojo\\Desktop\\test\\honey.txt', 'rb') as f:
    content = f.read()
    print(content)

hashOfHoney = hashlib.md5(content).hexdigest()
# print(hashOfHoney)

while True:
    with open('C:\\Users\\jojo\\Desktop\\test\\honey.txt', 'rb') as f:
        content = f.read()
    hashAfter = hashlib.md5(content).hexdigest()
    # print('BEFORE = ' + hashOfHoney)
    # print('AFTER  = ' + hashAfter)
    # print(content)
    if hashOfHoney != hashAfter:
        the_output = os.popen("taskkill /f /im GASFinal.exe").read()
        break
        # os.system("taskkill /f /im WINWORD.EXE")
    time.sleep(0.001)
