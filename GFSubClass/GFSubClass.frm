VERSION 5.00
Begin VB.Form Testfrm 
   Caption         =   "Form1"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   8055
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   60
      Top             =   3600
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Re-Subclass Testfrm2 and child controls"
      Height          =   315
      Left            =   60
      TabIndex        =   10
      Top             =   780
      Width           =   4575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Hide frm2"
      Height          =   375
      Left            =   3540
      TabIndex        =   9
      Top             =   420
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Show frm2"
      Height          =   375
      Left            =   3540
      TabIndex        =   8
      Top             =   60
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   5910
      Left            =   4740
      TabIndex        =   7
      Top             =   60
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   0
      Left            =   60
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   3240
      Width           =   4575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Disable Testfrm2"
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      ToolTipText     =   "send messages to Testfrm only"
      Top             =   420
      Width           =   1635
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Enable Testfrm2"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      ToolTipText     =   "send messages to Testfrm2, too"
      Top             =   60
      Width           =   1635
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   60
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1140
      Width           =   4575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "UnSubClass"
      Height          =   375
      Left            =   60
      TabIndex        =   2
      Top             =   420
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SubClass"
      Height          =   375
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   1635
   End
   Begin VB.PictureBox Picture1 
      Height          =   1635
      Left            =   60
      ScaleHeight     =   1575
      ScaleWidth      =   4515
      TabIndex        =   0
      Top             =   1500
      Width           =   4575
   End
End
Attribute VB_Name = "Testfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'(c)2001, 2002 by Louis. Use to quickly subclass one or more controls/windows.

Private Sub Form_Load()
    'on error resume next
    Dim Temp As Long
'    For Temp = 1 To 383
'        Load Text2(Temp)
'        Text2(Temp).Visible = True
'        Text2(Temp).Top = Text2(0).Top + Temp * 0.5! * Screen.TwipsPerPixelY
'        Call GFSubClassmod.GFSubClass(Me.Text2(Temp), "Text2(" + CStr(Temp) + ")", Me, True)
'    Next Temp
    For Temp = 1 To 10000
        List1.AddItem CStr(Temp) 'test if ruckling when scolling
    Next Temp
End Sub

Private Sub Command1_Click()
    'on error resume next
    Call GFSubClassmod.GFSubClass(Text1, "Text1", Testfrm, True)
    Call GFSubClassmod.GFSubClass(Text1, "Text2", Testfrm, True)
    Call GFSubClassmod.GFSubClass(Testfrm2.Text1, "Testfrm2.Text1", Testfrm, True)
End Sub

Private Sub Command2_Click()
    'on error resume next
    Call GFSubClassmod.GFSubClass_Terminate
End Sub

Public Sub GFSubClassWindowProc(ByVal SourceDescription As String, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef ReturnValue As Long, ByRef ReturnValueUsedFlag As Boolean)
    'on error resume next
    Debug.Print "Testfrm received message"
    Debug.Print SourceDescription
End Sub

Private Sub Command3_Click()
    'on error resume next
    'Call GFSubClassmod.GFSubClass(Text1, "Text1", Testfrm2, True)
    Call GFSubClassmod.GFSubClass_UnSubclass("Text1", Testfrm2)
End Sub

Private Sub Command4_Click()
    'on error resume next
    'Call GFSubClassmod.GFSubClass(Text1, "Text1", Testfrm2, False)
    Call GFSubClassmod.GFSubClass(Text1, "Text1", Testfrm2, True)
End Sub

Private Sub Command5_Click()
    'on error resume next
    Testfrm2.Show
End Sub

Private Sub Command6_Click()
    'on error resume next
    Testfrm2.Hide
End Sub

Private Sub Command7_Click()
    'on error resume next
    Call GFSubClass_ReSubClassByTargetObjectDescriptionPrefix("Testfrm2")
End Sub

Private Sub Timer1_Timer()
    'on error resume next
    Debug.Print Testfrm2.Text1.hWnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'on error resume next
    Call GFSubClassmod.GFSubClass_Terminate
End Sub

