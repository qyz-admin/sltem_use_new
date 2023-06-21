
import PySimpleGUI as sg

# 设置主题
# sg.theme('DarkAmber')

# 界面布局
layout = [
            [sg.Text('个人信息')],
            [sg.Text('姓名'), sg.InputText()],
            [sg.Text('年龄'), sg.InputText(key="input_age")],
            [sg.Button('提交'), sg.Button('取消')]
        ]

# 设置主窗口
window = sg.Window('填写信息', layout)

# 事件循环，监听事件的事件名和对应的值
while True:
    event, values = window.read()

    if event == sg.WIN_CLOSED:
        break
    elif event == "提交":
        print(event)  # 提交
        print(values)  # {0: 'pan', 'input_age': '18'}
    elif event == "取消":
        break

# 关闭窗口
window.close()



# import PySimpleGUI as sg
#
# layout = [[sg.Button('Click Me')]]
#
# window = sg.Window('My Second Window', layout)
#
# while True:
#     event, values = window.read()
#     if event == sg.WIN_CLOSED:
#         break
#     print('You clicked the button')
#
# window.close()