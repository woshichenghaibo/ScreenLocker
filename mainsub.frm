VERSION 5.00
Begin VB.Form mainsub 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4875
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H8000000A&
   Icon            =   "mainsub.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "mainsub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Const WS_EX_LAYERED = &H80000
Const GWL_EXSTYLE = (-20)
Const LWA_ALPHA = &H2
Const LWA_COLORKEY = &H1
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()



Private Sub Form_Click()
    Form1.Show
End Sub

Private Sub Form_Load()
    Call AddHook
    '锁住任务管理器
    Dim sTmp As String * 50
    Dim abc, bcd As String
    Dim length As Long
    length = GetSystemDirectory(sTmp, 50) '获取系统目录
    abc = Left(sTmp, length)
    bcd = abc & "\taskmgr.exe" '打开而不执行一个程序(任务管理器)
    Open bcd For Input Lock Read Write As #305 '以达到锁定的目的
    '设置透明
    Dim sty As Long
    sty = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
    sty = sty Or WS_EX_LAYERED
    SetWindowLong Me.hWnd, GWL_EXSTYLE, sty
    SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3
    SetLayeredWindowAttributes Me.hWnd, 1, 200, LWA_ALPHA Or LWA_COLORKEY
End Sub
Private Sub form_Initialize()
    InitCommonControls
End Sub

Private Sub AddHook()
    lHook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf LowLevelKeyboardProc, App.hInstance, 0)
End Sub

Private Sub DelHook()
    UnhookWindowsHookEx lHook
End Sub
Private Sub Form_Unload(Cancel As Integer)
'别忘了要卸载hook
    Call DelHook
End Sub
