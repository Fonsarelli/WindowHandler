
# Libraries
import win32gui
import win32api
import win32con

# Get handle and title of each visible window
def getAllWindows():

    # Initial empty array of handles
    handles = []

    # Add visible windows to array of handles
    win32gui.EnumWindows(lambda hwnd, ctx : handles.append([hwnd, win32gui.GetWindowText(hwnd)]) if (win32gui.IsWindowVisible(hwnd)) else None, None)

    # Return array of handles
    return handles

# Get last active specified window handle
def getWindowHandle(title):

    # Initial empty window handle
    window = None

    # Get all visible windows
    handles = getAllWindows()

    # Find first specified window
    for handle in handles:
        if (handle[1].find(title) != -1):
            window = handle[0]
            break

    # Return window handle
    return window

# Send a click to specified window and position
def sendClickToWindow(hwnd, pos=[0 ,0], type="left"):

    if (type == "left"):
        # Send left click to specified position
        win32api.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, win32api.MAKELONG(pos[0], pos[1])) 
        win32api.SendMessage(hwnd, win32con.WM_LBUTTONUP, win32con.MK_LBUTTON, win32api.MAKELONG(pos[0], pos[1]))
    elif (type == "right"):
        # Send right click to specified position
        win32api.SendMessage(hwnd, win32con.WM_RBUTTONDOWN, win32con.MK_RBUTTON, win32api.MAKELONG(pos[0], pos[1])) 
        win32api.SendMessage(hwnd, win32con.WM_RBUTTONUP, win32con.MK_RBUTTON, win32api.MAKELONG(pos[0], pos[1]))
    elif (type == "middle"):
    # Send right click to specified position
        win32api.SendMessage(hwnd, win32con.WM_MBUTTONDOWN, win32con.MK_MBUTTON, win32api.MAKELONG(pos[0], pos[1])) 
        win32api.SendMessage(hwnd, win32con.WM_MBUTTONUP, win32con.MK_MBUTTON, win32api.MAKELONG(pos[0], pos[1]))
    else:
        return False

# Send a keystroke to specified window and position
def sendKeystrokeToWindow(hwnd, key, type="down"):

    if (type == "down"):
        # Send specified down keystroke to specified position
        win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, key, 0)
    elif (type == "up"):
        # Send specified up keystroke to specified position
        win32api.PostMessage(hwnd, win32con.WM_KEYUP, key, 0)
    else:
        return False
    
# Send a character to specified window
def sendCharToWindow(hwnd, char):

    # Send specified character
    win32api.PostMessage(hwnd, win32con.WM_CHAR, ord(char), 0)
