VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   1110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7035
   BeginProperty Font 
      Name            =   "黑体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   1110
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "当前屏幕已被锁定"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
'设置透明
Dim sty As Long
Me.BackColor = RGB(66, 66, 66)
sty = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
sty = sty Or WS_EX_LAYERED
SetWindowLong Me.hwnd, GWL_EXSTYLE, sty
SetLayeredWindowAttributes Me.hwnd, RGB(66, 66, 66), 255, LWA_ALPHA Or LWA_COLORKEY
End Sub
