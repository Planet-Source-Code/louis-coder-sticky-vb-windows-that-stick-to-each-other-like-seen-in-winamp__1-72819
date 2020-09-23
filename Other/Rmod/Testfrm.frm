VERSION 5.00
Begin VB.Form Testfrm 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4635
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1380
      TabIndex        =   6
      Top             =   1560
      Width           =   1155
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   1155
   End
   Begin VB.CommandButton Command5 
      Caption         =   "RegGetKeyValue"
      Height          =   375
      Left            =   2700
      TabIndex        =   4
      Top             =   1500
      Width           =   1875
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Create Extended Sub Key List"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2340
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Create Key Value"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Create Sub Key List"
      Height          =   375
      Left            =   2700
      TabIndex        =   1
      Top             =   2340
      Width           =   1875
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create Key Value List"
      Height          =   375
      Left            =   2700
      TabIndex        =   0
      Top             =   2760
      Width           =   1875
   End
End
Attribute VB_Name = "Testfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'(c)2001 by Louis. Test form for Rmod.
'NOTE: use the Rmod in any project to access the Win95/98 Registry.

Private Sub Command1_Click()
    'on error resume next
    Dim RegValueNumber As Integer
    Dim RegValueNameArray() As String
    Dim RegValueValueArray() As String
    Dim Temp As Long
    'begin
    Debug.Print Rmod.RegGetKeyValueList(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\", RegValueNumber, RegValueNameArray(), RegValueValueArray())
    'display result
    For Temp = 1 To RegValueNumber
        Debug.Print RegValueNameArray(Temp)
        Debug.Print RegValueValueArray(Temp)
    Next Temp
End Sub

Private Sub Command2_Click()
    'on error resume next
    Dim RegSubKeyNumber As Integer
    Dim RegSubKeyNameArray() As String
    Dim Temp As Long
    'begin
    Debug.Print Rmod.RegGetSubKeyList(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\", RegSubKeyNumber, RegSubKeyNameArray())
    'display result
    For Temp = 1 To RegSubKeyNumber
        Debug.Print RegSubKeyNameArray(Temp)
    Next Temp
End Sub

Private Sub Command3_Click()
    'on errorr esume next
    Call RegCreateSubKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Rmod_Test\SubKey")
    Call RegSetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Rmod_Test\SubKey", "Test", "Test", REG_SZ)
End Sub

Private Sub Command4_Click()
    'on error resume next
    Dim RegSubKeyNumber As Integer
    Dim RegSubKeyNameArray() As String
    Dim Temp As Long
    'begin
    Debug.Print Rmod.RegGetSubKeyListEx(HKEY_LOCAL_MACHINE, "SOFTWARE\", RegSubKeyNumber, RegSubKeyNameArray(), True)
    'display result
    For Temp = 1 To RegSubKeyNumber
        Debug.Print RegSubKeyNameArray(Temp)
    Next Temp
End Sub

Private Sub Command5_Click()
    'on error resume next
    Debug.Print Rmod.RegGetKeyValue(HKEY_LOCAL_MACHINE, Text1.Text, Text2.Text)
End Sub
