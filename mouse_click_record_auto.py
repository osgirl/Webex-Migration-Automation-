import win32api
import win32con
import win32gui
import time

def click(x,y):
    win32api.SetCursorPos((x,y))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,x,y,0,0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,x,y,0,0)
mouse_clicks = []
mouse_points = []

state_left = win32api.GetKeyState(0x01)  # Left button down = 0 or 1. Button up = -127 or -128
state_right = win32api.GetKeyState(0x02)  # Right button down = 0 or 1. Button up = -127 or -128
try:
    while True:

        mouse_points.append(list(win32gui.GetCursorPos()))
        a = win32api.GetKeyState(0x01)
        b = win32api.GetKeyState(0x02)
        
        
        if a != state_left:  # Button state changed
            state_left = a
            print(a)
            if a < 0:
                mouse_clicks.append(1)
            else:
                mouse_clicks.append(0)
        elif b != state_right:  # Button state changed
            state_right = b
            if b < 0:
                mouse_clicks.append(2)
            else:
                mouse_clicks.append(0)
        else:
            mouse_clicks.append(0)

        time.sleep(0.01)

except KeyboardInterrupt:
    pass

for i in range(len(mouse_points)):
    time.sleep(.01)
    win32api.SetCursorPos((mouse_points[i][0], mouse_points[i][1]))
    if mouse_clicks[i] == 1:
        click(mouse_points[i][0], mouse_points[i][1])
        




   