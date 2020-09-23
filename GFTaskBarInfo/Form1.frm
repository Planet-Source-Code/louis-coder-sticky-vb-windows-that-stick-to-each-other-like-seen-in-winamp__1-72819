VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'Kein
   Caption         =   "Form1"
   ClientHeight    =   3225
   ClientLeft      =   0
   ClientTop       =   -120
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   60
      Top             =   60
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'(c)2001, 2004 by Louis. Test form for GFTaskBarInfomod.
'
'Downloaded from www.louis-coder.com.
'Use the GFTaskBarInfomod functions to get information about the Windows task bar position and
'size, so that windows of your programs can be placed and sized correctly (if they have BorderStyle
'0 for example).

Private Sub Form_Resize()
    'on error resume next
    Me.Cls 'reset
    Me.Line (0, 0)-(Me.Width - Screen.TwipsPerPixelX, Me.Height - Screen.TwipsPerPixelY), 0, B 'to make window edges clearly visible
End Sub

Private Sub Timer1_Timer()
    'on error resume next 'resizes/moves window when task bar is dragged around.
    Dim WindowXPos As Long
    Dim WindowYPos As Long
    Dim WindowXSize As Long
    Dim WindowYSize As Long
    'begin
    Call GFTaskBarInfo_GetWindowPosSize(WindowXPos, WindowYPos, WindowXSize, WindowYSize)
    Call Me.Move(WindowXPos, WindowYPos, WindowXSize, WindowYSize)
End Sub

Private Sub Command1_Click()
    'on error resume next
    Unload Me
End Sub

