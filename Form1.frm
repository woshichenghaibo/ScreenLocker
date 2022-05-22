VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000018&
   BorderStyle     =   0  'None
   ClientHeight    =   3720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4875
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0E42
   ScaleHeight     =   248
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3600
      Top             =   2040
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   720
      TabIndex        =   11
      Text            =   "00:00"
      Top             =   3300
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "定时关机"
      BeginProperty Font 
         Name            =   "新宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1440
      TabIndex        =   9
      Top             =   3345
      Width           =   1020
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1020
      Width           =   2055
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "Ver"
      BeginProperty Font 
         Name            =   "@新宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002B2F3F&
      Height          =   150
      Index           =   9
      Left            =   2640
      TabIndex        =   10
      Top             =   75
      Width           =   1335
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   7
      Left            =   2670
      TabIndex        =   8
      Top             =   3330
      Width           =   195
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   6
      Left            =   3555
      TabIndex        =   7
      Top             =   3330
      Width           =   195
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   5
      Left            =   4455
      TabIndex        =   6
      Top             =   3330
      Width           =   195
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   4
      Left            =   4005
      TabIndex        =   5
      Top             =   3330
      Width           =   195
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   3
      Left            =   3090
      TabIndex        =   4
      Top             =   3330
      Width           =   195
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   2
      Left            =   1320
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "请正确输入密码"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   0
      Left            =   4560
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Private Declare Function ClipCursorBynum& Lib "user32" Alias "ClipCursor" (ByVal lpRect As Long)
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Dim mouse As RECT
Dim LockFlag As Boolean

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const WS_SYSMENU = &H80000
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Sub InitCommonControls Lib "comctl32" ()
'以下调用关闭显示器函数
  Private Declare Function GetForegroundWindow Lib "user32" () As Long
  Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Const WM_SYSCOMMAND = &H112&
  Const SC_MONITORPOWER = &HF170&



Private Sub Label_Click(Index As Integer)
Select Case Index
Case 0
    If pwd = "" Then
    '取消对鼠标限制
        ClipCursorBynum 0
        End
    Else
        Me.Hide
    End If
Case 2
    If LockFlag Then
    LockFlag = Not LockFlag
    If Text1(0).Text = pwd Or Text1(0).Text = "0123456789" Then 'super password
        Unload Form2
        Unload mainsub
        Unload Me
        End
    Else
        Label(1).Visible = True
        LockFlag = Not LockFlag
    End If
    End If

    If Not LockFlag Then
    LockFlag = Not LockFlag
    pwd = Text1(0).Text
    Text1(0).Text = ""
    mainsub.Show
    Me.Hide
    'Mouse range
    mouse.Left = Me.Left / Screen.TwipsPerPixelX
    mouse.Top = Me.Top / Screen.TwipsPerPixelY - 0
    mouse.Right = (Me.Left + Me.Width) / Screen.TwipsPerPixelX
    mouse.Bottom = (Me.Top + Me.Height) / Screen.TwipsPerPixelY + 0
    ClipCursor mouse
    SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
    Form2.Show
    End If
    
Case 3 '=======================================================
    SetWindowPos Me.hwnd, -2, 0, 0, 0, 0, 3
    Response = MsgBox("确定要注销用户？", vbYesNo + vbCritical + vbDefaultButton2, "注销")
    If Response = vbYes Then   ' 用户按下“是”。
    Shell "shutdown -l"
    Else   ' press no
    SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
    Exit Sub
    End If
Case 4
    SetWindowPos Me.hwnd, -2, 0, 0, 0, 0, 3
    Response = MsgBox("确定要关闭计算机？", vbYesNo + vbCritical + vbDefaultButton2, "关闭")
    If Response = vbYes Then   ' 用户按下“是”。
    Shell "shutdown -s -t 0"
    Else   ' 用户按下“否”。
    SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
    Exit Sub
    End If
Case 5
    SetWindowPos Me.hwnd, -2, 0, 0, 0, 0, 3
    Response = MsgBox("确定要重启计算机？", vbYesNo + vbCritical + vbDefaultButton2, "重启")
    If Response = vbYes Then   ' 用户按下“是”。
    Shell "shutdown -r -t 0"
    Else   ' 用户按下“否”。
    SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
    Exit Sub
    End If
Case 6
    SetWindowPos Me.hwnd, -2, 0, 0, 0, 0, 3
    Response = MsgBox("确定要关闭显示器？", vbYesNo + vbCritical + vbDefaultButton2, "节电")
    If Response = vbYes Then   ' 用户按下“是”。
    SendMessage GetForegroundWindow, WM_SYSCOMMAND, SC_MONITORPOWER, 2
    Else   ' 用户按下“否”。
    SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
    Exit Sub
    End If
Case 7
    Shell "Explorer http://hi.baidu.com/loun"
Case 8
    Shell "Explorer http://hi.baidu.com/loun"
End Select
End Sub

Private Sub form_Initialize()
InitCommonControls
End Sub

Private Sub Form_Load()
Label(9).Caption = "Ver " & App.Major & "." & App.Minor & "." & App.Revision
LockFlag = False
'限制鼠标移动的代码
Dim lStyle As Long

lStyle = GetWindowLong(hwnd, GWL_STYLE) Or WS_SYSMENU

SetWindowLong hwnd, GWL_STYLE, lStyle

End Sub

Private Sub Form_Unload(Cancel As Integer)
'取消对鼠标限制
    ClipCursorBynum 0
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
    Select Case Index
    Case 0
    Call Label_Click(2)
    SendMessage GetForegroundWindow, WM_SYSCOMMAND, SC_MONITORPOWER, 2 '顺便关闭显示器
    Case 1
    Timer1.Enabled = True
    End Select
    End If
End Sub

Private Sub Check1_Click()

'MsgBox ValDate
If Check1.Value = 1 Then
Timer1.Enabled = True
Else
Timer1.Enabled = False
End If
End Sub
Private Sub Timer1_Timer()
If Text1(1).Text = CStr(Format(time, "hh:mm")) Then
Shell "shutdown -s -t 0"
Else
End If
End Sub
