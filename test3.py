from win32com.client import Dispatch
import os
from win32gui import FindWindow


class Operation:

    def __init__(self, dm, hwnd):
        self.dm = dm
        self.hwnd = hwnd
        # self.dm.Reg('注册码', '')
        # self.dm.register(r"RegDll.dll",
        #                    r"dm.dll")
        print(self.dm.Ver())
        self.bind()

    def bind(self):
        self.dm.BindWindowEx(self.hwnd, "normal", "normal", "normal", "", 0)
        self.dm.SetSimMode(0)
        self.dm.EnableRealKeypad(1)
        self.dm.EnableRealMouse(2, 20, 30)
        self.dm.SetKeypadDelay("normal", 70)
        self.dm.SetClientSize(self.hwnd, 596, 446)
        print(self.dm.GetClientSize(self.hwnd))
        print('绑定成功')


def regsvr():
    try:
        dm_1 = Dispatch('dm.dmsoft')
    except Exception:
        os.system(r'regsvr32 /s %s\dm.dll' % os.getcwd())
        dm_1 = Dispatch('dm.dmsoft')
    print(dm_1.Ver())
    return dm_1


if __name__ == '__main__':
    window_id = FindWindow('Notepad++', None)
    dm_main = regsvr()
    operation = Operation(dm_main, window_id)
    print("window_id:",window_id)
    dm_main.BindWindowEx(window_id, "normal", "normal", "normal", "", 0)
    # 设置鼠标的前台模拟方式，有需求的话可以切换。
    dm_main.SetSimMode(0)
    dm_main.EnableRealKeypad(1)
    # 设置键盘的仿真，即按下按键和放开按键的间隔随机而定（有函数作用范围，现在未讲，可以忽略）。
    dm_main.EnableRealMouse(2, 20, 30)
    # 设置鼠标的仿真，鼠标动作模拟真实操作, 带移动轨迹, 以及点击延时随机。
    dm_main.SetKeypadDelay("normal", 70)
    # 设置键盘按下放松的随机区间。
    dm_main.SetClientSize(window_id, 596, 446)
    # 设置窗口内容区域大小（什么是窗口内容区？就是除去窗口上面显示窗口类名和关闭窗口的条形剩下的区域）。

