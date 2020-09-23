VERSION 5.00
Begin VB.Form Testfrm2 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4635
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   60
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   60
      Width           =   4515
   End
End
Attribute VB_Name = "Testfrm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'(c)2001 by Louis.

Public Sub GFSubClassWindowProc(ByVal SourceDescription As String, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef ReturnValue As Long, ByRef ReturnValueUsedFlag As Boolean)
    'on error resume next
    Debug.Print "Test2frm received message"
    Debug.Print SourceDescription
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'on error resume next
    Call GFSubClass_ReSubClass_UnSubClassByTargetObjectDescriptionPrefix("Testfrm2")
End Sub
