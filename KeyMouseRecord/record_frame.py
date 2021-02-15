from PySide6.QtUiTools import QUiLoader
from PySide6.QtCore import QFile, QIODevice
import win32api
import pyWinhook
import json
import sys
import time
import datetime
import os
import threading
import win32con



frame = None


def create():
    global frame
    if not frame:
        frame = RecordFrame()
    return frame


def current_time():
    return int(time.time() * 1000)


def generate_script_path():
    time_format = datetime.datetime.now().strftime('%Y%m%d%H%M%S')
    if not os.path.exists("scripts"):
        os.mkdir("scripts")
    file_path = os.getcwd()
    return os.path.join(os.getcwd(), "scripts", f"{time_format}.txt")


def list_scripts():
    scripts = os.listdir('scripts')
    scripts.sort(reverse=True)
    return scripts


class RecordFrame:
    HOT_KEYS = ("F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10", "F11", "F12")
    all_mouse_messages = ('mouse left down', 'mouse left up', 'mouse right down', 'mouse right up', 'mouse move')
    all_key_messages = ("key down", "key up")
    current_mill_time = 0
    delay = 0
    run_times = 0
    start_hot_key = None
    stop_hot_key = None
    record = []

    def __init__(self):

        ui_file_name = "record_frame.ui"
        ui_file = QFile(ui_file_name)
        if not ui_file.open(QIODevice.ReadOnly):
            print("can not open file " + ui_file_name)
            sys.exit(-1)
        self.window = QUiLoader().load(ui_file)
        ui_file.close()
        self._component_bind()
        self.window.show()
        self.status = "Ready"
        self.record = []
        self.current_mill_time = current_time()
        self.delay = int(self.window.record_time_stepper.value())
        self.run_times = int(self.window.run_time_stepper.value())
        self.start_hot_key = self.window.run_hot_key.currentText()
        self.stop_hot_key = self.window.stop_hot_key.currentText()
        # self.window.record_script
        self.window.run_hot_key.addItems(self.HOT_KEYS)
        self.window.run_hot_key.setCurrentIndex(8)
        self.window.stop_hot_key.addItems(self.HOT_KEYS)
        self.window.stop_hot_key.setCurrentIndex(9)
        self._refresh_scripts()
        self.window.status.setText(self.status)

        self.hookManager = pyWinhook.HookManager()
        self.hookManager.MouseAll = self._mouse_move_handler
        self.hookManager.KeyAll = self._keyboard_click_handler
        self.hookManager.HookKeyboard()
        self.hookManager.HookMouse()
        #pythoncom.PumpMessages()

    def _component_bind(self):
        self.window.record_btn.clicked.connect(self._record_handler)
        self.window.script_run_btn.clicked.connect(self._script_run_handler)
        self.window.script_delete_btn.clicked.connect(self._script_delete_handler)
        self.window.record_time_stepper.valueChanged.connect(self._delay_changed_handler)
        self.window.run_time_stepper.valueChanged.connect(self._run_times_changed_handler)
        self.window.run_hot_key.currentTextChanged.connect(self._run_hot_key_changed_handler)
        self.window.stop_hot_key.currentTextChanged.connect(self._stop_hot_key_changed_handler)

    def _record_handler(self):
        print(f"_record_handler ")
        if self.status != "Recording":
            self.status = "Recording"
            self.record = []
        else:
            self.status = "Finish"
        self.change_btn_status(self.status)

        if self.status == "Finish":
            output = json.dumps(self.record, indent=1)
            # output = output.replace('\n   ', '').replace('\n  ', '')
            # output = output.replace('\n ]', ']')
            print(output)
            try:
                file = open(generate_script_path(), 'w', encoding='utf8')
                file.write(output)
                file.close()
            except Exception as e:
                print(e)
            self.record = []
            self._refresh_scripts()

    def _script_run_handler(self):
        print(f"_script_run_handler ")
        if self.status != "Running":
            self.status = "Running"
            self.change_btn_status(self.status)
            self._run_script()
        else:
            self.status = "Finish"
            self.change_btn_status(self.status)
            self._stop_script()


    def _script_delete_handler(self):
        path = os.path.join(os.getcwd(), "scripts", self.selected_script)
        if os.path.exists(path):
            os.remove(path)
        self._refresh_scripts()

    def _keyboard_click_handler(self, event):
        print(f"_keyboard_click_handler ")
        # 监听键盘事件
        message = event.MessageName.replace(' sys ', ' ')
        key_name = event.Key
        if key_name:
            key_name.upper()
        key_ID = event.KeyID
        key_extended = event.Extended
        if self.status != "Recording" and key_name != self.start_hot_key and key_name != self.stop_hot_key:
            return True
        elif key_name == self.start_hot_key and (self.status == "Ready" or self.status == "Finish"):
            self._run_script()
            return True
        elif key_name == self.stop_hot_key and self.status == "Running":
            self._stop_script()
            return True
        elif self.status == "Recording":
            delay_time = current_time() - self.current_mill_time
            key_info = (key_name, key_ID, key_extended)
            self.record.append(["keyboard", delay_time, message, key_info])
            self.current_mill_time = current_time()
        # 同鼠标事件监听函数的返回值
        return True

    def _run_script(self):
        print(f"_run_script = ")
        self.status = "Running"
        self.change_btn_status(self.status)
        print(f"_run_script = {self.selected_script}")
        ScriptRunThread("script_run", self.run_times, self.selected_script, self.window.remaining_num, self).start()

    def _stop_script(self):
        print(f"_stop_script = ")

        self.status = "intercepted"
        self.change_btn_status(self.status)
        print(f"_stop_script = {self.selected_script}")
        pass

    def _mouse_move_handler(self, event):
       # print(f"_mouse_move_handler = ")
        # 监听鼠标事件
        pos = win32api.GetCursorPos()
        self.window.point_x.setText(str(pos[0]))
        self.window.point_y.setText(str(pos[1]))

        if event.MessageName not in self.all_mouse_messages:
            return True

        delay_time = current_time() - self.current_mill_time
        if event.MessageName == "mouse move" and delay_time < self.delay:
            return True
        if not self.record:
            delay_time = 0

        # print(delay_time, event.MessageName, pos)
        self.record.append(["mouse", delay_time, event.MessageName, pos])
        self.current_mill_time = current_time()
        return True

    def _delay_changed_handler(self, event):
        self.delay = event
        print(f"self.delay = {self.delay}")

    def _run_times_changed_handler(self, event):
        self.run_times = event
        print(f"self.run_times = {self.run_times}")

    def _run_hot_key_changed_handler(self, event):
        self.start_hot_key = event
        print(f"self.start_hot_key = {self.start_hot_key}")

    def _stop_hot_key_changed_handler(self, event):
        self.stop_hot_key = event
        print(f"self.stop_hot_key = {self.stop_hot_key}")

    def change_btn_status(self, status="Ready"):
        print(f"_change_btn_status  {status}")
        status_message = None
        if status == "Ready":
            status_message = u'Ready'
            self.window.status.setStyleSheet('color: black;')
            self.window.record_btn.setEnabled(True)
            self.window.script_run_btn.setEnabled(True)
            self.window.script_delete_btn.setEnabled(True)
        elif status == "Recording":
            status_message = u'脚本录制中'
            self.window.record_btn.setText(u"停止")
            self.window.status.setStyleSheet('color: red;')
            self.window.record_btn.setEnabled(True)
            self.window.script_run_btn.setEnabled(False)
            self.window.script_delete_btn.setEnabled(False)
        elif status == "Running":
            status_message = u'脚本执行中'
            self.window.script_run_btn.setText(u"结束")
            self.window.status.setStyleSheet('color: red;')
            self.window.record_btn.setEnabled(False)
            self.window.script_run_btn.setEnabled(True)
            self.window.script_delete_btn.setEnabled(False)
        elif status == "Finish":
            status_message = u'Finish'
            self.window.record_btn.setText(u"录制")
            self.window.script_run_btn.setText(u"运行")
            self.window.status.setStyleSheet('color: black;')
            self.window.record_btn.setEnabled(True)
            self.window.script_run_btn.setEnabled(True)
            self.window.script_delete_btn.setEnabled(True)
        else:
            status_message = u'intercepted'
            self.window.record_btn.setText(u"录制")
            self.window.script_run_btn.setText(u"运行")
            self.window.status.setStyleSheet('color: black;')
            self.window.record_btn.setEnabled(True)
            self.window.script_run_btn.setEnabled(True)
            self.window.script_delete_btn.setEnabled(True)
        self.window.status.setText(status_message)

    def _refresh_scripts(self):
        self.scripts = list_scripts()
        self.window.record_script.clear()
        self.window.record_script.addItems(self.scripts)
        self.selected_script = self.window.record_script.currentText()


class ScriptRunThread(threading.Thread):

    def __init__(self, thread_name, run_times, script, run_times_label, frame):
        threading.Thread.__init__(self)
        self.thread_name = thread_name
        self.run_times = run_times
        self.script = script
        self.run_times_label = run_times_label
        self.frame = frame

    def run(self):
        current_run_time = 1
        while current_run_time <= self.run_times or self.run_times == 0:
            print(f"run = ")
            print(self.script)
            if self.frame.window.status.text() == "intercepted":
                break
            self.script_run()
            self.run_times_label.setText(str(current_run_time))
            current_run_time += 1
        self.frame.change_btn_status( "Finish")
    def script_run(self):
        print(f"script_run ")
        file = open(os.path.join(os.getcwd(), "scripts", self.script), "r", encoding="utf8")
        script = json.load(file)
        for cmd in script:
            print(cmd)
            if self.frame.window.status.text() == "intercepted":
                break
            kind, delay, event,info = cmd
            time.sleep(int(delay)/1000)
            if kind == "mouse":
                if event == "mouse left down":
                    win32api.mouse_event(win32con.MOUSEEVENTF_ABSOLUTE + win32con.MOUSEEVENTF_LEFTDOWN, info[0], info[1], 0, 0)
                elif event == "mouse left up":
                    win32api.mouse_event(win32con.MOUSEEVENTF_ABSOLUTE + win32con.MOUSEEVENTF_LEFTUP, info[0], info[1], 0, 0)
                elif event == "mouse right down":
                    win32api.mouse_event(win32con.MOUSEEVENTF_ABSOLUTE + win32con.MOUSEEVENTF_RIGHTDOWN, info[0], info[1], 0, 0)
                elif event == "mouse right up":
                    win32api.mouse_event(win32con.MOUSEEVENTF_ABSOLUTE + win32con.MOUSEEVENTF_RIGHTUP, info[0], info[1], 0, 0)
                elif event == "mouse move":
                    win32api.SetCursorPos(info)
                else:
                    pass

            elif kind == "keyboard":
                if event == "key down":
                    win32api.keybd_event(info[1], 0, win32con.KEYEVENTF_EXTENDEDKEY, 0)
                elif event == "key up":
                    win32api.keybd_event(info[1], 0, win32con.KEYEVENTF_EXTENDEDKEY + win32con.KEYEVENTF_KEYUP, 0)
                else:
                    pass
            else:
                pass