Attribute VB_Name = "GFTaskBarInfomod"
Option Explicit
'(c)2001, 2004 by Louis. Stuff to place/size windows correctly without covering the stupid task bar.
'Downloaded from www.louis-coder.com.
'
'GFGetTaskBar[Height/Width]
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
'GFGetTaskBar[Height/Width]
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Sub GFTaskBarInfo_GetWindowPosSize(ByRef WindowXPos As Long, ByRef WindowYPos As Long, ByRef WindowXSize As Long, ByRef WindowYSize As Long)
    'on error resume next 'format: twips
    '
    'NOTE: the initialized values are to be used to place/size
    'any window so that it looks like maximized.
    'Use Form1.Move(WindowXPos, WindowYPos, WindowXSize, WindowYSize)
    'to correctly place/size a window.
    '
    If GFTaskBarInfo_GetTaskBarLeft > (Screen.Width / 2!) Then
        WindowXPos = 0
    Else
        WindowXPos = GFTaskBarInfo_GetTaskBarWidth
    End If
    If GFTaskBarInfo_GetTaskBarTop > (Screen.Height / 2!) Then
        WindowYPos = 0
    Else
        WindowYPos = GFTaskBarInfo_GetTaskBarHeight
    End If
    WindowXSize = Screen.Width - GFTaskBarInfo_GetTaskBarWidth
    WindowYSize = Screen.Height - GFTaskBarInfo_GetTaskBarHeight
    Exit Sub
End Sub

Public Sub GFTaskBarInfo_GetVisibleScreenArea(ByRef AreaLeft As Long, ByRef AreaTop As Long, ByRef AreaRight As Long, ByRef AreaBottom As Long)
    'on error resume next 'returns rectangular area that's not covered by the stupid task bar; format: twips
    Dim WindowXPos As Long
    Dim WindowYPos As Long
    Dim WindowXSize As Long
    Dim WindowYSize As Long
    'begin
    Call GFTaskBarInfo_GetWindowPosSize(WindowXPos, WindowYPos, WindowXSize, WindowYSize)
    AreaLeft = WindowXPos
    AreaTop = WindowYPos
    AreaRight = WindowXPos + WindowXSize - Screen.TwipsPerPixelX
    AreaBottom = WindowYPos + WindowYSize - Screen.TwipsPerPixelY
End Sub

Public Function GFTaskBarInfo_GetTaskBarLeft() As Long
    'on error resume next 'returns task bar left position in twips
    Dim TaskBarhWnd As Long
    Dim RECTVar As RECT
    'preset
    GFTaskBarInfo_GetTaskBarLeft = 0 'preset
    'begin
    TaskBarhWnd = FindWindow("Shell_traywnd", "")
    If TaskBarhWnd = 0 Then Exit Function 'verify
    Call GetWindowRect(TaskBarhWnd, RECTVar)
    GFTaskBarInfo_GetTaskBarLeft = RECTVar.Left * Screen.TwipsPerPixelX 'format: twips
    Exit Function
End Function

Private Function GFTaskBarInfo_GetTaskBarTop() As Long
    'on error resume next 'returns task bar top position in twips
    Dim TaskBarhWnd As Long
    Dim RECTVar As RECT
    'preset
    GFTaskBarInfo_GetTaskBarTop = 0 'preset
    'begin
    TaskBarhWnd = FindWindow("Shell_traywnd", "")
    If TaskBarhWnd = 0 Then Exit Function 'verify
    Call GetWindowRect(TaskBarhWnd, RECTVar)
    GFTaskBarInfo_GetTaskBarTop = RECTVar.Top * Screen.TwipsPerPixelY 'format: twips
    Exit Function
End Function

Public Function GFTaskBarInfo_GetTaskBarWidth() As Long
    'on error resume next 'returns task bar width in twips or 0 if the task bar has screen width or if no task bar was found
    Dim TaskBarhWnd As Long
    Dim RECTVar As RECT
    'preset
    GFTaskBarInfo_GetTaskBarWidth = 0 'preset
    'begin
    TaskBarhWnd = FindWindow("Shell_traywnd", "")
    If TaskBarhWnd = 0 Then Exit Function 'error
    Call GetWindowRect(TaskBarhWnd, RECTVar)
    If (RECTVar.Right - RECTVar.Left + 1) < (Screen.Width / Screen.TwipsPerPixelX) Then
        'NOTE: two twips must be substracted as GetWindowRect returns 2 twips too much for every value (!?).
        GFTaskBarInfo_GetTaskBarWidth = (RECTVar.Right - RECTVar.Left + 0 - 2) * Screen.TwipsPerPixelX 'format: twips
        If GFTaskBarInfo_GetTaskBarWidth > Screen.Width Then GFTaskBarInfo_GetTaskBarWidth = Screen.Width 'important
        Exit Function
    Else
        Exit Function
    End If
End Function

Public Function GFTaskBarInfo_GetTaskBarHeight() As Long
    'on error resume next 'returns task bar height in twips or 0 if the task bar has screen height or if no task bar was found
    Dim TaskBarhWnd As Long
    Dim RECTVar As RECT
    'preset
    GFTaskBarInfo_GetTaskBarHeight = 0 'preset
    'begin
    TaskBarhWnd = FindWindow("Shell_traywnd", "")
    If TaskBarhWnd = 0 Then Exit Function 'error
    Call GetWindowRect(TaskBarhWnd, RECTVar)
    If (RECTVar.Bottom - RECTVar.Top + 1) < (Screen.Height / Screen.TwipsPerPixelX) Then
        'NOTE: two twips must be substracted as GetWindowRect returns 2 twips too much for every value (!?).
        GFTaskBarInfo_GetTaskBarHeight = (RECTVar.Bottom - RECTVar.Top + 0 - 2) * Screen.TwipsPerPixelY 'format: twips
        If GFTaskBarInfo_GetTaskBarHeight > Screen.Height Then GFTaskBarInfo_GetTaskBarHeight = Screen.Height 'important
        Exit Function
    Else
        Exit Function
    End If
End Function

