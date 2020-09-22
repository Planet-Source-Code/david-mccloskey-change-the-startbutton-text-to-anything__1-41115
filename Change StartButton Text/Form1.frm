VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   6645
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   315
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   525
      Width           =   4320
   End
   Begin VB.CommandButton Command1 
      Caption         =   "set text"
      Height          =   540
      Left            =   4830
      TabIndex        =   0
      Top             =   420
      Width           =   1170
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   645
      Left            =   1470
      TabIndex        =   2
      Top             =   2310
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim wnd As Long, tWnd As Long


Private Sub Command1_Click()
    Dim dummy() As Byte
    
    ReDim dummy(Len(Text1) + 1)
    
    dummy = StringToByteArray(Text1.Text)
    
    Call SendMessage(wnd, WM_SETTEXT, 0&, dummy(0))
    
End Sub

Private Sub Form_Load()
    'Find the taskbar window , Shell_TrayWnd
    tWnd = FindWindow("Shell_TrayWnd", "")
    
    '5 stands for GW_CHILD or GW_MAX
    wnd = GetWindow(tWnd, 5)
    
    'Start button child hwnd = 196668
    Label1.Caption = wnd
End Sub
