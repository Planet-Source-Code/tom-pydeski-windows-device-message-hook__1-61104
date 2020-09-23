VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Tom Pydeski's Device Change Message Hook"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   9735
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   50
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   50
      Width           =   9375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Text1.Left = 50
Text1.Top = 50
SetWinMess
Dim NotificationFilter As DEV_BROADCAST_DEVICEINTERFACE
With NotificationFilter
    .dbcc_size = Len(NotificationFilter)
    .dbcc_devicetype = DBT_DEVTYP_DEVICEINTERFACE
End With
HookWin Me.hWnd
hDevNotify = RegisterDeviceNotification(Me.hWnd, NotificationFilter, DEVICE_NOTIFY_WINDOW_HANDLE Or DEVICE_NOTIFY_ALL_INTERFACE_CLASSES)
Text1.SelStart = Len(Text1.Text)
Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call UnregisterDeviceNotification(hDevNotify)
UnHookWin
End Sub

Private Sub Form_Resize()
If WindowState = vbMinimized Then Exit Sub
Me.Text1.Width = (Me.Width - (2 * Text1.Left)) - 100
Me.Text1.Height = (Me.Height - (2 * Text1.Top)) - 500
End Sub

