Attribute VB_Name = "Module1"
'如果模块和窗体API声明重复可以删除窗体里的 也可以不删除
'yardloun制作  HOOK、透明、锁鼠标用法来源网络 可以自己仔细琢磨一番 会用就行
Private Declare Function StartMaskKey Lib "MaskKey" (lpdwVirtualKey As Long, ByVal nLength As Long, Optional ByVal bDisableKeyboard As Boolean = False) As Long
Private Declare Function StopMaskKey Lib "MaskKey" () As Long

Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, ByVal lpvSource As Long, ByVal cbCopy As Long)
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const WS_EX_LAYERED = &H80000
Public Const GWL_EXSTYLE = (-20)
Public Const LWA_ALPHA = &H2
Public Const LWA_COLORKEY = &H1
Dim t As Long
Dim tms As Long, kg As Boolean
Public Type KEYMSGS
    vKey As Long
    sKey As Long
    flag As Long
    time As Long
End Type
Public Const WH_KEYBOARD_LL = 13
Public Const VK_LWIN = &H5B
Public Const VK_RWIN = &H5C
Public Const VK_CONTROL = &H11
Public Const VK_SHIFT = &H10
Public Const HC_ACTION = 0
Public Const HC_SYSMODALOFF = 5
Public Const HC_SYSMODALON = 4
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_SYSKEYDOWN = &H104
Public Const WM_SYSKEYUP = &H105
Public pwd As String
Public P As KEYMSGS
Public lHook As Long















'还有一段可以禁用Ctrl+Esc Alt + Esc Alt+Tab三组热键的
'Private Const WH_KEYBOARD_LL = 13& ''enables monitoring of keyboard
''input events about to be posted
''in a thread input queue
'Private Const HC_ACTION = 0& ''wParam and lParam parameters
''contain information about a
''keyboard message
Private Const LLKHF_EXTENDED = &H1& ''test the extended-key flag
Private Const LLKHF_INJECTED = &H10& ''test the event-injected flag
Private Const LLKHF_ALTDOWN = &H20& ''test the context code
Private Const LLKHF_UP = &H80& ''test the transition-state flag
Private Const VK_TAB = &H9 ''virtual key constants
'Private Const VK_CONTROL = &H11
Private Const VK_ESCAPE = &H1B
Private Type KBDLLHOOKSTRUCT
vkCode As Long ''a virtual-key code in the range 1 to 254
scanCode As Long ''hardware scan code for the key
flags As Long ''specifies the extended-key flag,
''event-injected flag, context code,
''and transition-state flag
time As Long ''time stamp for this message
dwExtraInfo As Long ''extra info associated with the message
End Type




Private Declare Function GetAsyncKeyState Lib "user32" _
(ByVal vKey As Long) As Integer
Private m_hDllKbdHook As Long ''private variable holding
''the handle to the hook procedure
Public Sub Main()
''set and obtain the handle to the keyboard hook
m_hDllKbdHook = SetWindowsHookEx(WH_KEYBOARD_LL, _
AddressOf LowLevelKeyboardProc, _
App.hInstance, _
0&)
If m_hDllKbdHook <> 0 Then
MsgBox "Ctrl+Esc, Alt+Tab and Alt+Esc are blocked. " & _
"Click OK to quit and re-enable the keys.", _
vbOKOnly Or vbInformation, _
"Keyboard Hook Active"
Call UnhookWindowsHookEx(m_hDllKbdHook)
Else
MsgBox "Failed to install low-level keyboard hook - " & Err.LastDllError
End If
End Sub
Public Function LowLevelKeyboardProc(ByVal nCode As Long, _
ByVal wParam As Long, _
ByVal lParam As Long) As Long
Static kbdllhs As KBDLLHOOKSTRUCT


If nCode = HC_ACTION Then
Call CopyMemory(kbdllhs, ByVal lParam, Len(kbdllhs))

''Ctrl+Esc --------------
If (kbdllhs.vkCode = VK_ESCAPE) And _
CBool(GetAsyncKeyState(VK_CONTROL) _
And &H8000) Then
Debug.Print "Ctrl+Esc blocked"
LowLevelKeyboardProc = 1
Exit Function
End If ''kbdllhs.vkCode = VK_ESCAPE
''Ctrl+Alt --------------
If (kbdllhs.vkCode = VK_CONTROL) And CBool(kbdllhs.flags And _
LLKHF_ALTDOWN) Then
Debug.Print "Ctrl+Alt blocked"
LowLevelKeyboardProc = 1
Exit Function
End If ''kbdllhs.vkCode = VK_ESCAPE
''Alt+Tab --------------
If (kbdllhs.vkCode = VK_TAB) And _
CBool(kbdllhs.flags And _
LLKHF_ALTDOWN) Then
Debug.Print "Alt+Tab blocked"
LowLevelKeyboardProc = 1
Exit Function
End If ''kbdllhs.vkCode = VK_TAB
''Alt+Esc --------------
If (kbdllhs.vkCode = VK_ESCAPE) And _
CBool(kbdllhs.flags And _
LLKHF_ALTDOWN) Then
Debug.Print "Alt+Esc blocked"
LowLevelKeyboardProc = 1
Exit Function
End If ''kbdllhs.vkCode = VK_ESCAPE
''Lwin --------------
If (kbdllhs.vkCode = VK_LWIN) And _
CBool(kbdllhs.flags) Then
Debug.Print "Lwin blocked"
LowLevelKeyboardProc = 1
Exit Function
End If ''kbdllhs.vkCode = VK_LWIN
''Rwin --------------
If (kbdllhs.vkCode = VK_RWIN) And _
CBool(kbdllhs.flags) Then
Debug.Print "Rwin blocked"
LowLevelKeyboardProc = 1
Exit Function
End If ''kbdllhs.vkCode = VK_RWIN
End If ''nCode = HC_ACTION
LowLevelKeyboardProc = CallNextHookEx(m_hDllKbdHook, _
nCode, _
wParam, _
lParam)
End Function








