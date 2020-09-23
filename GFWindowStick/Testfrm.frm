VERSION 5.00
Begin VB.Form Testfrm 
   Caption         =   "Testfrm"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4635
   LinkTopic       =   "Form4"
   ScaleHeight     =   3195
   ScaleWidth      =   4635
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command4 
      Caption         =   "Show 3"
      Height          =   435
      Left            =   3420
      TabIndex        =   3
      Top             =   2220
      Width           =   1155
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Show 2"
      Height          =   435
      Left            =   3420
      TabIndex        =   2
      Top             =   1740
      Width           =   1155
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Show 1"
      Height          =   435
      Left            =   3420
      TabIndex        =   1
      Top             =   1260
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Restore Stick Type"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   2760
      Width           =   2175
   End
End
Attribute VB_Name = "Testfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'(c)2001, 2004 by Louis. Makes windows (slave windows) sticky to one special window (master window).
'
'Downloaded from www.louis-coder.com.
'Add GFWindowStickfrm and all modules to your project. Then you can use
'the sticky effects like demonstrated in this form (Testfrm).
'Sample usage: Toricxs (www.toricxs.com).
'
'ProgramGetMousePos[X, Y]
Private Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINTAPI) As Long
'ProgramGetMousePos[X, Y]
Private Type POINTAPI
    x As Long
    y As Long
End Type
'FormMoveStruct
Private Type FormMoveStruct
    FormMoveEnabledFlag As Boolean
    FormMoveDeltaXPos As Long
    FormMoveDeltaYPos As Long
End Type
Dim FormMoveStructVar As FormMoveStruct

Private Sub Command1_Click()
    'on error resume next
    Dim XPos As Long
    Dim YPos As Long
    'begin
    If GFWindowStickfrm.GetSlaveWindowPosBest("Form1", XPos, YPos) = True Then 'always check return value
        Call Form1.Move(XPos, YPos)
    End If
    If GFWindowStickfrm.GetSlaveWindowPosBest("Form2", XPos, YPos) = True Then 'always check return value
        Call Form2.Move(XPos, YPos)
    End If
    If GFWindowStickfrm.GetSlaveWindowPosBest("Form3", XPos, YPos) = True Then 'always check return value
        Call Form3.Move(XPos, YPos)
    End If
End Sub

Private Sub Command2_Click()
    'on error resume next
    Form1.Enabled = True
    Form1.Visible = True
    Form1.Refresh
End Sub

Private Sub Command3_Click()
    'on error resume next
    Form2.Enabled = True
    Form2.Visible = True
    Form2.Refresh
End Sub

Private Sub Command4_Click()
    'on error resume next
    Form3.Enabled = True
    Form3.Visible = True
    Form3.Refresh
End Sub

Private Sub Form_Load()
    'on error resume next
    With GFWindowStickfrm
        'NOTE: under WinXP (NT?) we can only create a sub key in Software\, not in the HKEY_LOCAL_MACHINE root key.
        Call .GFWindowStick_Initialize(HKEY_LOCAL_MACHINE, "Software\GFWindowStick Test", "Testfrm", Testfrm)
        Call .GFWindowStick_AddWindow("Testfrm", Testfrm)
        Call .GFWindowStick_AddWindow("Form1", Form1)
        Call .GFWindowStick_AddWindow("Form2", Form2)
        Call .GFWindowStick_AddWindow("Form3", Form3)
    End With
    Form1.Show
    Form2.Show
    Form3.Show
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'on error resume next
    FormMoveStructVar.FormMoveEnabledFlag = True
    FormMoveStructVar.FormMoveDeltaXPos = ProgramGetMousePosX
    FormMoveStructVar.FormMoveDeltaYPos = ProgramGetMousePosY
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'on error resume next
    If FormMoveStructVar.FormMoveEnabledFlag = True Then
        Testfrm.Left = Testfrm.Left + Screen.TwipsPerPixelX * (ProgramGetMousePosX - FormMoveStructVar.FormMoveDeltaXPos)
        Testfrm.Top = Testfrm.Top + Screen.TwipsPerPixelY * (ProgramGetMousePosY - FormMoveStructVar.FormMoveDeltaYPos)
        FormMoveStructVar.FormMoveDeltaXPos = ProgramGetMousePosX
        FormMoveStructVar.FormMoveDeltaYPos = ProgramGetMousePosY
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'on error resume next
    FormMoveStructVar.FormMoveEnabledFlag = False
End Sub

Public Function ProgramGetMousePosX() As Long
    On Error Resume Next 'the format is: pixels
    Dim ProgramGetMousePosXTemp As Long
    Dim CurrentMousePos As POINTAPI
    ProgramGetMousePosXTemp = GetCursorPos(CurrentMousePos)
    ProgramGetMousePosX = CurrentMousePos.x
End Function

Public Function ProgramGetMousePosY() As Long
    On Error Resume Next 'the format is: pixels
    Dim ProgramGetMousePosYTemp As Long
    Dim CurrentMousePos As POINTAPI
    ProgramGetMousePosYTemp = GetCursorPos(CurrentMousePos)
    ProgramGetMousePosY = CurrentMousePos.y
End Function

Private Sub Form_Unload(Cancel As Integer)
    'on error resume next
    Call GFSubClass_Terminate
    End
End Sub

