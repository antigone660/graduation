import time
import threading
import keyboard
import win32com.client as client
import pythoncom
import operation

title = "中国象棋2017"
exit_flag = False
sentence = []
# 键盘输入回调函数

class KeyboardThread(threading.Thread):
    def __init__(self):
        super().__init__()
        self.sentence = []

    def run(self):
        keyboard.on_release_key('esc', self.stop)
        keyboard.on_release_key('enter', self.add_sentence)
        keyboard.on_release(self.add_char)

        keyboard.wait('esc')

    def add_char(self, event):
        if event.name != 'enter' and event.name != 'esc':
            self.sentence.append(event.name)

    def add_sentence(self, event):
        sentence_str = ''.join(self.sentence)
        if sentence_str[:4] == 'left':
            sentence_str = sentence_str[8:]
        operation.move_window_to_foreground(title)
        print('收到一句话：', sentence_str)
        self.sentence = []

    def stop(self, event):
        keyboard.unhook_all()
        #print('程序已退出')

# 定时截图线程
def capture_thread():
    while not exit_flag:
        # 执行截图操作
        global title
        operation.capture_window(title,'data/dh.png')

        # 间隔一段时间后再次执行截图操作
        time.sleep(5)

keyboard_thread = KeyboardThread()
keyboard_thread.start()



# 启动定时截图线程
cap_thread = threading.Thread(target=capture_thread)
cap_thread.start()

keyboard_thread.join()
exit_flag = True
cap_thread.join()
# 程序结束
print('程序已退出')


