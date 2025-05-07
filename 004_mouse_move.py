from dmMain import Operation
from win32gui import FindWindow
from dmMain import regsvr
# 教程链接地址
# https://www.cnblogs.com/100-rzsyztd?page=3
def close_txt_1(operation_1):
    # 随便点击两次观察光标
    operation_1.leftclick(100, 10, 5, 5)
    operation_1.leftclick(100, 180, 5, 5)


def close_txt_2(operation_1):
    # 组合键  按下alt（不松） 然后按下f（松）  接着按下x（松） 松开alt
    # operation_1.keypress(187)
    operation_1.keydown(18)
    operation_1.keypress(70)
    operation_1.keypress(88)
    operation_1.keyup(18)


if __name__ == '__main__':
    window_id = FindWindow('Notepad++', None)
    dm_main = regsvr()
    operation = Operation(dm_main, window_id)
    close_txt_1(operation)
    close_txt_2(operation)