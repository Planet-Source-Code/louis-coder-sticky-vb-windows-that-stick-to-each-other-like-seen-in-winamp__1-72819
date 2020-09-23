VERSION 5.00
Begin VB.Form GFWindowStickfrm 
   BorderStyle     =   0  'Kein
   Caption         =   "GFWindowStickfrm"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   -60
   ClientWidth     =   4635
   LinkTopic       =   "Form4"
   ScaleHeight     =   3195
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
End
Attribute VB_Name = "GFWindowStickfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'(c)2001-2003 by Louis.
'
'NOTE: this project is not finished yet. 31.12.2002: yes it is
'NOTE: damn this code, it never wants to work right. 31.12.2002: yes it wants to
'
'GetWindowStateChange
Private Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long
'GFMoveMinimizedWindow
'Private Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
'other
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
'GFSubClassWindowProc
Private Const WM_MOVE = &H3
Private Const WM_SIZE = &H5
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_WINDOWPOSCHANGING = &H46
Private Const WM_WINDOWPOSCHANGED = &H47
Private Const WM_PAINT = &HF
Private Const WM_ERASEBKGND = &H14
Private Const WM_NCPAINT = &H85
'WindowStateChange constants
Private Const WINDOWSTATECHANGE_NOCHANGE As Integer = 0
Private Const WINDOWSTATECHANGE_WASMINIMIZED As Integer = 1
Private Const WINDOWSTATECHANGE_WASRESTORED As Integer = 2
Private Const WINDOWSTATECHANGE_WASMAXIMIZED As Integer = 3
'StickTypeBitConstants - used when writing StickTypeBits to registry
Private Const STICKY As Long = 1
Private Const STICKY_INDIRECT As Long = 2
Private Const STICKY_AT_TOP As Long = 16
Private Const STICKY_AT_BOTTOM As Long = 32
Private Const STICKY_AT_LEFT As Long = 64
Private Const STICKY_AT_RIGHT As Long = 128
Private Const TOP_HEIGHT_STICKY As Long = 256 '.Top values of slave and master window are equal
Private Const BOTTOM_HEIGHT_STICKY As Long = 512 'botton edge of slave and master window are at same height
Private Const LEFT_WIDTH_STICKY As Long = 1024
Private Const RIGHT_WIDTH_STICKY As Long = 2048
'NOTE: some of the bits 17-32 are reserved for a distance value (see code for further information).
'Version
Private Const Version As String = "v1.2"
'GFSubClassWindowProc
Private Type WINDOWPOS
    hwnd As Long
    hWndInsertAfter As Long
    x As Long
    y As Long
    CX As Long
    CY As Long
    Flags As Long
End Type
'GFMoveMinimizedWindow
Private Type POINTAPI
    x As Long
    y As Long
End Type
'GFMoveMinimizedWindow
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'GFMoveMinimizedWindow
Private Type WINDOWPLACEMENT
    Length As Long
    Flags As Long
    showCmd As Long
    ptMinPosition As POINTAPI
    ptMaxPosition As POINTAPI
    rcNormalPosition As RECT
End Type
'GFWindowStickStruct - general data
Private Type GFWindowStickStruct
    RegMainKey As Long
    RegRootKey As String
    GFWindowStickSystemEnabledFlag As Boolean
    MasterWindowName As String
    MasterWindowObject As Form
    MasterWindowLeftOld As Long
    MasterWindowTopOld As Long
    MasterWindowWidthOld As Long
    MasterWindowHeightOld As Long
    UserMoveWindowIndex As Integer 'which window is currently moved through the mouse
End Type
Dim GFWindowStickStructVar As GFWindowStickStruct
'WindowStickStruct - stores window data
Private Type WindowStickStruct
    WindowName As String
    WindowObject As Form
    WindowTopOld As Long
    WindowLeftOld As Long
    WindowWidthOld As Long
    WindowHeightOld As Long
    IsWindowAtTopFlag As Boolean
    IsWindowAtBottomFlag As Boolean
    IsWindowAtLeftFlag As Boolean
    IsWindowAtRightFlag As Boolean
    IsWindowTopHeightStickyFlag As Boolean
    IsWindowBottomHeightStickyFlag As Boolean
    IsWindowLeftWidthStickyFlag As Boolean
    IsWindowRightWidthStickyFlag As Boolean
    IsWindowStickyFlag As Boolean
    IsWindowStickyIndirectFlag As Boolean 'if the window sticks at one or more other slave windows that stick at the master window
    IsIconicFlagOld As Boolean
    IsZoomedFlagOld As Boolean
    AboutMovingFlag As Boolean 'if window is during the moving process
End Type
Dim WindowStickStructNumber As Integer
Dim WindowStickStructArray() As WindowStickStruct
'WindowSizeStruct
Private Type WindowSizeStruct
    WindowStickStructIndex As Integer 'index of window that is sized
End Type
Dim WindowSizeStructVar As WindowSizeStruct
'MasterWindowMoveResultStruct
Private Type MasterWindowMoveResultStruct
    XPosFixedFlag As Boolean
    XPos As Long 'format: pixels
    YPosFixedFlag As Boolean
    YPos As Long 'format: pixels
End Type
'SlaveWindowMoveResultStruct - same as MasterWindowMoveResultStruct
Private Type SlaveWindowMoveResultStruct
    XPosFixedFlag As Boolean
    XPos As Long 'format: pixels
    YPosFixedFlag As Boolean
    YPos As Long 'format: pixels
End Type
'StickTypeBitStruct
Private Type StickTypeBitStruct
    SlaveWindowName As String
    SlaveWindowStickTypeBits As Long
End Type
Dim StickTypeBitStructNumber As Integer
Dim StickTypeBitStructArray() As StickTypeBitStruct
'GFMoveMinimizedWindow
Private Const WPF_SETMINPOSITION = &H1
Private Const SW_SHOWNA = 8
'other consts
Private Const UNDEFINED As Long = -256& ^ 3& 'something that will never become a window pos
'other
Dim MasterWindowIndex As Integer 'for caching
Dim MasterWindowMovingFlag As Boolean

Private Sub Form_Load()
    'on error resume next
    'do nothing
End Sub

'************************************INTERFACE SUBS************************************
'NOTE: the target project can temporary disable the whole GFWindowStick system
'through calling GFWindowStickSystem_Disable.

Public Sub GFWindowStick_Initialize(ByVal RegMainKey As Long, ByVal RegRootKey As String, ByVal MasterWindowName As String, ByRef MasterWindowObject As Form)
    'on error resume next 'call first
    If Not (Right$(RegRootKey, 1) = "\") Then RegRootKey = RegRootKey + "\" 'verify
    GFWindowStickStructVar.RegMainKey = RegMainKey
    GFWindowStickStructVar.RegRootKey = RegRootKey 'another sub key will be added by system (pass e.g. 'Software\MyApp\')
    GFWindowStickStructVar.GFWindowStickSystemEnabledFlag = True 'preset
    GFWindowStickStructVar.MasterWindowName = MasterWindowName
    Set GFWindowStickStructVar.MasterWindowObject = MasterWindowObject
    GFWindowStickStructVar.MasterWindowLeftOld = MasterWindowObject.Left
    GFWindowStickStructVar.MasterWindowTopOld = MasterWindowObject.Top
    GFWindowStickStructVar.MasterWindowWidthOld = MasterWindowObject.Width
    GFWindowStickStructVar.MasterWindowHeightOld = MasterWindowObject.Height
    Call StickTypeBitStructFromReg(StickTypeBitStructNumber, StickTypeBitStructArray())
End Sub

'NOTE: if the GFWindowStickSystem is disabled then incoming messages are not processed any more.
'That means the stick type bits of any window will not be changed until the system is enabled again.
'Temporary disable the system if a form is temporary maximized.
'The target project should restore all slave window's position when the system is reenabled.
'
Public Sub GFWindowStickSystem_Enable()
    'on error resume next
    If GFWindowStickStructVar.GFWindowStickSystemEnabledFlag = False Then
        GFWindowStickStructVar.GFWindowStickSystemEnabledFlag = True 'process incoming messages
    End If
    '
    'NOTE: tests showed that the old master window position must be updated in any case,
    'no matter if the GFWindowStickSystem was already enabled or not.
    '
    If (GetMasterWindowIndex) Then 'verify (important)
        Dim WindowObject As Form
        Set WindowObject = WindowStickStructArray(GetMasterWindowIndex).WindowObject
        GFWindowStickStructVar.MasterWindowLeftOld = WindowObject.Left 'WindowObject is the master window in this case
        GFWindowStickStructVar.MasterWindowTopOld = WindowObject.Top
        GFWindowStickStructVar.MasterWindowWidthOld = WindowObject.Width
        GFWindowStickStructVar.MasterWindowHeightOld = WindowObject.Height
    End If
End Sub

Public Function GFWindowStickSystem_Enabled() As Boolean
    'on error resume next 'use to determinate if the GFWindowSticksystem is enabled
    GFWindowStickSystem_Enabled = GFWindowStickStructVar.GFWindowStickSystemEnabledFlag
End Function

Public Sub GFWindowStickSystem_Disable()
    'on error resume next
    If GFWindowStickStructVar.GFWindowStickSystemEnabledFlag = True Then
        GFWindowStickStructVar.GFWindowStickSystemEnabledFlag = False 'don't process incoming messages
    End If
End Sub

Public Sub GFWindowStick_AddWindow(ByVal WindowName As String, ByRef WindowObject As Form)
    'on error resume next
    Dim TempObject As Object
    If Not (WindowStickStructNumber = 32766) Then 'verify
        WindowStickStructNumber = WindowStickStructNumber + 1
    Else
        Exit Sub 'error
    End If
    ReDim Preserve WindowStickStructArray(1 To WindowStickStructNumber) As WindowStickStruct
    WindowStickStructArray(WindowStickStructNumber).WindowName = WindowName
    Set WindowStickStructArray(WindowStickStructNumber).WindowObject = WindowObject
    'Set TempObject = WindowObject
    Call GFSubClass(WindowObject, WindowName + "(GFWindowStick)", Me, True) 'TempObject (ehm, what was this good for?)
    'IMPORTANT: add "(GFWindowStick)" at end of name as the GFSubClass code needs a special naming format to work.
End Sub

Public Sub GFWindowStick_UpdateWindowReference(ByVal WindowName As String, ByRef WindowObjectNew As Form)
    'on error resume next 'call if a form has been unloaded and loaded again; pass e.g. "Form1", Me
    Dim WindowStickStructIndex As Integer
    'begin
    WindowStickStructIndex = GetWindowStickStructIndex(WindowName + "(GFWindowStick)")
    If (WindowStickStructIndex) Then 'verify
        Set WindowStickStructArray(WindowStickStructIndex).WindowObject = WindowObjectNew
    Else
        MsgBox "internal error in GFWindowStick_UpdateWindowReference(): passed value invalid !", vbOKOnly + vbExclamation
    End If
End Sub

'********************************END OF INTERFACE SUBS*********************************
'*************************************CALLBACK SUBS************************************
'NOTE: a WM_WINDOWPOSCHANGING message is sent before the related window
'is moved or sized. The future position/size of the window can be manipulated through
'changing values of the WINDOWPOS structure.

Public Sub GFSubClassWindowProc(ByVal SourceDescription As String, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef ReturnValue As Long, ByRef ReturnValueUsedFlag As Boolean)
    'on error Resume Next
    'verify
    '
    If Msg = WM_PAINT Then Exit Sub 'increase speed
    If Msg = WM_NCPAINT Then Exit Sub
    If Msg = WM_ERASEBKGND Then Exit Sub
    If GFWindowStickStructVar.GFWindowStickSystemEnabledFlag = False Then Exit Sub 'do nothing
    '
    Select Case Msg 'filter to increase speed
    Case 561, WM_LBUTTONDOWN, 562, WM_LBUTTONUP, WM_WINDOWPOSCHANGING, WM_MOVE, WM_SIZE
        '
        'NOTE: declare vars in GFSubClassWindowProc() subs only if necessary to increase
        'message processing speed (otherwise we get a window-salad).
        '
        Dim WindowStickStructIndex As Integer
        Dim TempMasterWindowMoveResultStruct As MasterWindowMoveResultStruct
        Dim TempSlaveWindowMoveResultStruct As SlaveWindowMoveResultStruct
        Dim XPos As Long
        Dim YPos As Long
        Dim MasterWindowIndex As Integer
        Dim WINDOWPOSVar As WINDOWPOS
        Dim StructLoop As Integer
    Case Else
        Exit Sub 'increase speed
    End Select
    'preset
    WindowStickStructIndex = GetWindowStickStructIndex(SourceDescription)
    If WindowStickStructIndex = 0 Then Exit Sub 'nothing to do
    'If (WindowStickStructArray(WindowStickStructIndex).WindowObject.Enabled = False) Or (WindowStickStructArray(WindowStickStructIndex).WindowObject.Visible = False) Then _
    '    Exit Sub 'do nothing when target project is loaded (user must move the window)
    'NOTE: tests showed that also disabled and hidden windows should be processed.
    'begin
    Select Case Msg
    Case 561 'mouse down over frame
        '
        'NOTE: WindowSizeStructVar.WindowStickStructIndex is used to avoid that
        'a form that is currently resized is also made stick to a form.
        'WindowSizeStructVar.WindowStickStructIndex is set when a WM_SIZE message arrives
        'and reset to zero when the user releases the left mouse button.
        'But a WM_SIZE message also arrives when a form is initialized (no mouse button release),
        'that's why we must reset WindowSizeStructVar.WindowStickStructIndex below to avoid
        'that the GFWindowStick system thinks that a form is resized although it is moved
        '(a WM_SIZE message will arrive instantly if the form is really resized).
        '
        GFWindowStickStructVar.UserMoveWindowIndex = WindowStickStructIndex
        WindowSizeStructVar.WindowStickStructIndex = 0 'reset
        If IsMasterWindow(SourceDescription) = True Then
            Call SlaveWindow_FreezePosition
        End If
    Case WM_LBUTTONDOWN 'if no frame is existing
        GFWindowStickStructVar.UserMoveWindowIndex = WindowStickStructIndex
        WindowSizeStructVar.WindowStickStructIndex = 0 'reset
        If IsMasterWindow(SourceDescription) = True Then
            Call SlaveWindow_FreezePosition
        End If
    Case 562 'mouse up over frame
        Call StickTypeBitStruct_Update(StickTypeBitStructNumber, StickTypeBitStructArray(), WindowStickStructNumber, WindowStickStructArray())
        Call StickTypeBitStructToReg(StickTypeBitStructNumber, StickTypeBitStructArray()) 'save changes
        GFWindowStickStructVar.UserMoveWindowIndex = 0 'reset
        WindowSizeStructVar.WindowStickStructIndex = 0 'reset (if not it doesn't matter anyway (if 'Esc' pressed))
    Case WM_LBUTTONUP 'if no frame is existing
        Call StickTypeBitStruct_Update(StickTypeBitStructNumber, StickTypeBitStructArray(), WindowStickStructNumber, WindowStickStructArray())
        Call StickTypeBitStructToReg(StickTypeBitStructNumber, StickTypeBitStructArray()) 'save changes
        GFWindowStickStructVar.UserMoveWindowIndex = 0 'reset
        WindowSizeStructVar.WindowStickStructIndex = 0 'reset (if not it doesn't matter anyway (if 'Esc' pressed))
    Case WM_WINDOWPOSCHANGING
        'Debug.Print Time$ & "->" & WindowStickStructArray(WindowStickStructIndex).WindowName & " ICONIC: " & IsIconic(WindowStickStructArray(WindowStickStructIndex).WindowObject.hwnd) & " ZOOMED: " & IsZoomed(WindowStickStructArray(WindowStickStructIndex).WindowObject.hwnd)
        If Not (WindowStickStructIndex = WindowSizeStructVar.WindowStickStructIndex) Then
            If (WindowStickStructArray(WindowStickStructIndex).AboutMovingFlag = False) And _
                (MasterWindowMovingFlag = False) Then
                WindowStickStructArray(WindowStickStructIndex).AboutMovingFlag = True
                If IsMasterWindow(SourceDescription) = True Then
                    MasterWindowMovingFlag = True 'important to avoid that slave windows move themselves when being moved around with master window
                    '
                    StructLoop = GetMasterWindowIndex
                    If (IsIconic(WindowStickStructArray(StructLoop).WindowObject.hwnd) = 0) And _
                        (IsZoomed(WindowStickStructArray(StructLoop).WindowObject.hwnd) = 0) Then 'verify
                        '
                        'NOTE: a window that is minimized is Iconic() before the message arrives
                        '(tested).
                        '
                        'If Not (StructLoop = WindowStickStructIndex) Then
                            'stick moved window to screen edges
                            Call CopyMemory(WINDOWPOSVar, ByVal lParam, Len(WINDOWPOSVar))
                            TempMasterWindowMoveResultStruct = MasterWindow_Move2( _
                                WINDOWPOSVar.x * Screen.TwipsPerPixelX, WINDOWPOSVar.x * Screen.TwipsPerPixelX + WindowStickStructArray(WindowStickStructIndex).WindowObject.Width - TX(1), _
                                WINDOWPOSVar.y * Screen.TwipsPerPixelY, WINDOWPOSVar.y * Screen.TwipsPerPixelY + WindowStickStructArray(WindowStickStructIndex).WindowObject.Height - TY(1))
                            Call CopyMemory(WINDOWPOSVar, ByVal lParam, Len(WINDOWPOSVar))
                            If TempMasterWindowMoveResultStruct.XPosFixedFlag = True Then
                                WINDOWPOSVar.x = TempMasterWindowMoveResultStruct.XPos
                            End If
                            If TempMasterWindowMoveResultStruct.YPosFixedFlag = True Then
                                WINDOWPOSVar.y = TempMasterWindowMoveResultStruct.YPos
                            End If
                            Call CopyMemory(ByVal lParam, WINDOWPOSVar, Len(WINDOWPOSVar))
                        'End If
                    End If
                    '
                    'Call MasterWindow_Move(SourceDescription, WindowStickStructArray(WindowStickStructIndex).WindowObject)
                    MasterWindowMovingFlag = False 'reset
                Else
                    '
                    'NOTE: enable the out-commented stuff to allow sticking slave windows
                    'to other slave windows. This is not recommended as IsWindowStickyIndirect()
                    'does not work, and indirectly sticked slave windows could not be moved through
                    'the master window.
                    '
                    'For StructLoop = 1 To WindowStickStructNumber
                        'If WindowStickStructArray(StructLoop).WindowObject.Visible = True Then 'verify (important)
                        'vs.
                        '
                        'stick moved window to screen edges if not currently sticky (also if MasterWindow iconic or zoomed)
                        If WindowStickStructArray(WindowStickStructIndex).IsWindowStickyFlag = False Then
                            Call CopyMemory(WINDOWPOSVar, ByVal lParam, Len(WINDOWPOSVar))
                            TempMasterWindowMoveResultStruct = MasterWindow_Move2( _
                                WINDOWPOSVar.x * Screen.TwipsPerPixelX, WINDOWPOSVar.x * Screen.TwipsPerPixelX + WindowStickStructArray(WindowStickStructIndex).WindowObject.Width - TX(1), _
                                WINDOWPOSVar.y * Screen.TwipsPerPixelY, WINDOWPOSVar.y * Screen.TwipsPerPixelY + WindowStickStructArray(WindowStickStructIndex).WindowObject.Height - TY(1))
                            Call CopyMemory(WINDOWPOSVar, ByVal lParam, Len(WINDOWPOSVar))
                            If TempMasterWindowMoveResultStruct.XPosFixedFlag = True Then
                                WINDOWPOSVar.x = TempMasterWindowMoveResultStruct.XPos
                            End If
                            If TempMasterWindowMoveResultStruct.YPosFixedFlag = True Then
                                WINDOWPOSVar.y = TempMasterWindowMoveResultStruct.YPos
                            End If
                            Call CopyMemory(ByVal lParam, WINDOWPOSVar, Len(WINDOWPOSVar))
                        End If
                        '
                        StructLoop = GetMasterWindowIndex
                        If (IsIconic(WindowStickStructArray(StructLoop).WindowObject.hwnd) = 0) And _
                            (IsZoomed(WindowStickStructArray(StructLoop).WindowObject.hwnd) = 0) Then 'verify
                            '
                            'NOTE: a window that is minimized is Iconic() before the message arrives
                            '(tested).
                            '
                            If Not (StructLoop = WindowStickStructIndex) Then
                                'stick moved window to any other window (except itself)
                                Call CopyMemory(WINDOWPOSVar, ByVal lParam, Len(WINDOWPOSVar))
                                TempSlaveWindowMoveResultStruct = SlaveWindow_Move( _
                                    SourceDescription, WindowStickStructArray(WindowStickStructIndex).WindowObject, _
                                    WindowStickStructArray(StructLoop).WindowObject.Left, _
                                    WindowStickStructArray(StructLoop).WindowObject.Left + WindowStickStructArray(StructLoop).WindowObject.Width - TX(1), _
                                    WindowStickStructArray(StructLoop).WindowObject.Top, _
                                    WindowStickStructArray(StructLoop).WindowObject.Top + WindowStickStructArray(StructLoop).WindowObject.Height - TY(1), _
                                    WINDOWPOSVar.x * Screen.TwipsPerPixelX, WINDOWPOSVar.x * Screen.TwipsPerPixelX + WindowStickStructArray(WindowStickStructIndex).WindowObject.Width - TX(1), _
                                    WINDOWPOSVar.y * Screen.TwipsPerPixelY, WINDOWPOSVar.y * Screen.TwipsPerPixelY + WindowStickStructArray(WindowStickStructIndex).WindowObject.Height - TY(1))
                                Call CopyMemory(WINDOWPOSVar, ByVal lParam, Len(WINDOWPOSVar))
                                If TempSlaveWindowMoveResultStruct.XPosFixedFlag = True Then
                                    WINDOWPOSVar.x = TempSlaveWindowMoveResultStruct.XPos
                                End If
                                If TempSlaveWindowMoveResultStruct.YPosFixedFlag = True Then
                                    WINDOWPOSVar.y = TempSlaveWindowMoveResultStruct.YPos
                                End If
                                Call CopyMemory(ByVal lParam, WINDOWPOSVar, Len(WINDOWPOSVar))
                                If (TempSlaveWindowMoveResultStruct.XPosFixedFlag = False) And (TempSlaveWindowMoveResultStruct.YPosFixedFlag = False) Then
                                    If (WindowStickStructArray(WindowStickStructIndex).WindowObject.WindowState = vbNormal) And _
                                        (WindowStickStructIndex = GFWindowStickStructVar.UserMoveWindowIndex) Then 'verify (do not reset sticky flag if minimized or not moved by user (very important))
                                        '
                                        'NOTE: the sticky flag is updated not before the user releases the mouse button.
                                        'If a window is dragged from the master window to the screen edge then it would
                                        'not instantly become sticky to the screen edge.
                                        '
                                        WindowStickStructArray(WindowStickStructIndex).IsWindowStickyFlag = False 'reset
                                    End If
                                End If
                                If (TempSlaveWindowMoveResultStruct.XPosFixedFlag = True) Or (TempSlaveWindowMoveResultStruct.YPosFixedFlag = True) Then
                                    'Exit For 'slave window has been made stuck once, exit
                                End If
                            End If
                        End If
                    'Next StructLoop
                End If
                WindowStickStructArray(WindowStickStructIndex).AboutMovingFlag = False 'reset
            End If
        End If
'        If IsMasterWindow(SourceDescription) = False Then 'for some reason the bullshit does not work, is slow and does not move the window before being restored
'            If (IsIconic(WindowStickStructArray(WindowStickStructIndex).WindowObject.hwnd)) Then
'                '
'                'NOTE: we move the window before it becomes visible (looks much better).
'                'As we cannot use any VB functions to move a minimized window we must use
'                'a General Function specially made for that purpose.
'                '
'                If GetSlaveWindowPosBest(WindowStickStructArray(WindowStickStructIndex).WindowName, XPos, YPos) = True Then 'do not use SourceDescription
'                    If WindowStickStructArray(WindowStickStructIndex).AboutMovingFlag = False Then 'verify (avoid endless loop, important)
'                        WindowStickStructArray(WindowStickStructIndex).AboutMovingFlag = True
'                        Debug.Print WindowStickStructArray(WindowStickStructIndex).WindowName, XPos / Screen.TwipsPerPixelX, YPos / Screen.TwipsPerPixelY
'                        Call GFMoveMinimizedWindow(WindowStickStructArray(WindowStickStructIndex).WindowObject.hwnd, XPos / Screen.TwipsPerPixelX, YPos / Screen.TwipsPerPixelY)
'                        WindowStickStructArray(WindowStickStructIndex).AboutMovingFlag = False
'                    End If
'                End If
'            End If
'        End If
    Case WM_MOVE
        If IsMasterWindow(SourceDescription) = True Then
            MasterWindowMovingFlag = True 'important to avoid that slave windows move themselves when being moved around with master window
            Call MasterWindow_Move(SourceDescription, WindowStickStructArray(WindowStickStructIndex).WindowObject)
            MasterWindowMovingFlag = False 'reset
        Else
        End If
    Case WM_SIZE
        If IsMasterWindow(SourceDescription) = True Then
            Select Case GetWindowStateChange(GetMasterWindowIndex)
            Case WINDOWSTATECHANGE_WASRESTORED
                For StructLoop = 1 To WindowStickStructNumber
                    If WindowStickStructArray(StructLoop).IsWindowStickyFlag = True Then
                        WindowStickStructArray(StructLoop).WindowObject.WindowState = vbNormal
                        WindowStickStructArray(StructLoop).WindowObject.Refresh 'looks better (window content drawn instantly, not only frame)
                    End If
                Next StructLoop
            Case WINDOWSTATECHANGE_WASMINIMIZED
                For StructLoop = 1 To WindowStickStructNumber
                    If WindowStickStructArray(StructLoop).IsWindowStickyFlag = True Then
                        WindowStickStructArray(StructLoop).WindowObject.WindowState = vbMinimized
                        WindowStickStructArray(StructLoop).WindowObject.Refresh 'looks better (window content drawn instantly, not only frame)
                    End If
                Next StructLoop
            Case WINDOWSTATECHANGE_WASMAXIMIZED
                'do not update stick type bits or size of any slave window
            Case WINDOWSTATECHANGE_NOCHANGE
                WindowSizeStructVar.WindowStickStructIndex = WindowStickStructIndex 'avoid 'nervous' window (when size is changed only, do not set index when a window is maximized, restored or minimized)
                Call MasterWindow_Size(SourceDescription, WindowStickStructArray(WindowStickStructIndex).WindowObject)
            End Select
        Else
            '
            'NOTE: a WM_SIZE message also arrives when a window was
            'minimized, maximized or restored.
            '
            Select Case GetWindowStateChange(WindowStickStructIndex)
            Case WINDOWSTATECHANGE_WASRESTORED
                If WindowStickStructArray(WindowStickStructIndex).IsWindowStickyFlag = True Then 'if restored window is not sticky then the other windows are not restored, too (indirectly through restoring the master window)
                    MasterWindowIndex = GetMasterWindowIndex
                    If MasterWindowIndex = 0 Then Exit Sub 'verify
                    If Not (WindowStickStructArray(MasterWindowIndex).WindowObject.WindowState = vbNormal) Then
                        WindowStickStructArray(MasterWindowIndex).WindowObject.WindowState = vbNormal
                        WindowStickStructArray(MasterWindowIndex).WindowObject.Refresh 'looks better (window content drawn instantly, not only frame)
                    End If
                    If GetSlaveWindowPosBest(WindowStickStructArray(WindowStickStructIndex).WindowName, XPos, YPos) = True Then 'do not use SourceDescription
                        Call WindowStickStructArray(WindowStickStructIndex).WindowObject.Move(XPos, YPos)
                        Call WindowStickStructArray(WindowStickStructIndex).WindowObject.Refresh 'looks better (window content drawn instantly, not only frame)
                    End If
                End If
            Case WINDOWSTATECHANGE_WASMAXIMIZED
                'do not update stick type bits
            Case WINDOWSTATECHANGE_WASMINIMIZED
                'do not update stick type bits
            Case WINDOWSTATECHANGE_NOCHANGE
                WindowSizeStructVar.WindowStickStructIndex = WindowStickStructIndex 'avoid 'nervous' window (when size is changed only, do not set index when a window is maximized, restored or minimized)
                Call SlaveWindow_Size(SourceDescription, WindowStickStructArray(WindowStickStructIndex).WindowObject)
            End Select
        End If
    End Select
    Exit Sub
End Sub

'*********************************END OF CALLBACK SUBS*********************************
'************************************MASTER WINDOW*************************************

Private Function GetMasterWindowIndex()
    'on error resume next 'returns index of master window (name got out of structure, needn't to be passed) or 0 for error
    Dim StructLoop As Integer
    'verify
    If Not (MasterWindowIndex = 0) Then 'once the master window index is set it isn't changed any more and can be cached
        GetMasterWindowIndex = MasterWindowIndex
        Exit Function
    End If
    'begin
    For StructLoop = 1 To WindowStickStructNumber
        If Len(WindowStickStructArray(StructLoop).WindowName) = Len(GFWindowStickStructVar.MasterWindowName) Then 'check first to increase speed
            If WindowStickStructArray(StructLoop).WindowName = GFWindowStickStructVar.MasterWindowName Then
                MasterWindowIndex = StructLoop 'cache master window index
                GetMasterWindowIndex = StructLoop 'ok
                Exit Function
            End If
        End If
    Next StructLoop
    GetMasterWindowIndex = 0 'error (should not happen)
    Exit Function
End Function

Private Function IsMasterWindow(ByVal SourceDescription As String) As Boolean
    'on error resume next 'call this function out of GFSubClassWindowProc()
    If SourceDescription = GFWindowStickStructVar.MasterWindowName + "(GFWindowStick)" Then
        IsMasterWindow = True
    Else
        IsMasterWindow = False
    End If
End Function

Private Sub MasterWindow_Move(ByVal WindowName As String, ByVal WindowObject As Form)
    'on error resume next
    Dim DeltaTop As Long
    Dim DeltaLeft As Long
    Dim StructLoop As Integer
    '
    'NOTE: all slave windows are moved around together with the master window.
    'NOTE: the WM_MOVE message is sent after (!) the related window has been moved.
    '
    'preset
    DeltaLeft = GFWindowStickStructVar.MasterWindowObject.Left - GFWindowStickStructVar.MasterWindowLeftOld
    DeltaTop = GFWindowStickStructVar.MasterWindowObject.Top - GFWindowStickStructVar.MasterWindowTopOld
    'begin
    For StructLoop = 1 To WindowStickStructNumber
        If Not (WindowStickStructArray(StructLoop).WindowName + "(GFWindowStick)" = WindowName) Then 'do not move master window
            If (WindowStickStructArray(StructLoop).IsWindowStickyFlag = True) Or _
                (WindowStickStructArray(StructLoop).IsWindowStickyIndirectFlag = True) Then
                If (IsIconic(WindowStickStructArray(StructLoop).WindowObject.hwnd) = 0) And _
                    (IsZoomed(WindowStickStructArray(StructLoop).WindowObject.hwnd) = 0) Then 'verify
                    WindowStickStructArray(StructLoop).WindowObject.Left = WindowStickStructArray(StructLoop).WindowObject.Left + DeltaLeft
                    WindowStickStructArray(StructLoop).WindowObject.Top = WindowStickStructArray(StructLoop).WindowObject.Top + DeltaTop
                End If
            End If
        End If
    Next StructLoop
    GFWindowStickStructVar.MasterWindowLeftOld = WindowObject.Left 'WindowObject is the master window in this case
    GFWindowStickStructVar.MasterWindowTopOld = WindowObject.Top
    GFWindowStickStructVar.MasterWindowWidthOld = WindowObject.Width
    GFWindowStickStructVar.MasterWindowHeightOld = WindowObject.Height
End Sub

Private Sub MasterWindow_Size(ByVal WindowName As String, ByVal WindowObject As Form)
    'on error resume next
    Dim DeltaHeight As Long
    Dim DeltaWidth As Long
    Dim StructLoop As Integer
    '
    'NOTE: this function verifies the slave windows still stick at the master window after resizing the master window.
    'NOTE: the WM_SIZE message is sent after (!) the related window has been sized.
    '
    'preset
    DeltaWidth = GFWindowStickStructVar.MasterWindowObject.Width - GFWindowStickStructVar.MasterWindowWidthOld
    DeltaHeight = GFWindowStickStructVar.MasterWindowObject.Height - GFWindowStickStructVar.MasterWindowHeightOld
    'begin
    For StructLoop = 1 To WindowStickStructNumber
        If Not (WindowStickStructArray(StructLoop).WindowName + "(GFWindowStick)" = WindowName) Then 'do not size master window
            If (IsIconic(WindowStickStructArray(StructLoop).WindowObject.hwnd) = 0) And _
                (IsZoomed(WindowStickStructArray(StructLoop).WindowObject.hwnd) = 0) Then 'verify
                If WindowStickStructArray(StructLoop).IsWindowAtTopFlag = True Then
                    'do nothing
                End If
                If WindowStickStructArray(StructLoop).IsWindowAtBottomFlag = True Then
                    WindowStickStructArray(StructLoop).WindowObject.Top = WindowStickStructArray(StructLoop).WindowObject.Top + DeltaHeight
                End If
                If WindowStickStructArray(StructLoop).IsWindowAtLeftFlag = True Then
                    'do nothing
                End If
                If WindowStickStructArray(StructLoop).IsWindowAtRightFlag = True Then
                    WindowStickStructArray(StructLoop).WindowObject.Left = WindowStickStructArray(StructLoop).WindowObject.Left + DeltaWidth
                End If
            End If
        End If
    Next StructLoop
    GFWindowStickStructVar.MasterWindowLeftOld = WindowObject.Left 'WindowObject is the master window in this case
    GFWindowStickStructVar.MasterWindowTopOld = WindowObject.Top
    GFWindowStickStructVar.MasterWindowWidthOld = WindowObject.Width
    GFWindowStickStructVar.MasterWindowHeightOld = WindowObject.Height
End Sub

Private Function MasterWindow_Move2(ByVal MasterWindowLeft As Long, ByVal MasterWindowRight As Long, ByVal MasterWindowTop As Long, ByVal MasterWindowBottom As Long) As MasterWindowMoveResultStruct
    'on error resume next
    Const STICKDISTANCEMAX As Long = 8
    Dim AreaLeft As Long 'visible screen area
    Dim AreaTop As Long
    Dim AreaRight As Long
    Dim AreaBottom As Long
    '
    'NOTE: this sub checks if the master window is near enough to one of the four screen edges
    'to be docked (made sticky). GFWindowStick uses GFTaksBarInfofrm functions to
    'determinate the visible screen area.
    '
    'The calling sub must set this window x/y pos through manipulating the parameters of the received
    'WM_WINDOWPOSCHANGING message.
    '
    'NOTE: do not call WindowObject.Top etc. but make calling sub set WindowObject's window position.
    '
    'NOTE: a master window can currently be made stuck to a screen edge only.
    'The master window position is currently NOT saved in registry.
    '
    'preset
    Call GFTaskBarInfo_GetVisibleScreenArea(AreaLeft, AreaTop, AreaRight, AreaBottom)
    'begin
    If (Abs(MasterWindowLeft - AreaLeft) < (STICKDISTANCEMAX * Screen.TwipsPerPixelX)) Then
        MasterWindow_Move2.XPosFixedFlag = True
        MasterWindow_Move2.XPos = (AreaLeft) / Screen.TwipsPerPixelX
    End If
    If (Abs(MasterWindowRight - AreaRight) < (STICKDISTANCEMAX * Screen.TwipsPerPixelX)) Then
        MasterWindow_Move2.XPosFixedFlag = True
        MasterWindow_Move2.XPos = (AreaRight - (MasterWindowRight - MasterWindowLeft + 0 * Screen.TwipsPerPixelX)) / Screen.TwipsPerPixelX
    End If
    If (Abs(MasterWindowTop - AreaTop) < (STICKDISTANCEMAX * Screen.TwipsPerPixelY)) Then
        MasterWindow_Move2.YPosFixedFlag = True
        MasterWindow_Move2.YPos = (AreaTop) / Screen.TwipsPerPixelY
    End If
    If (Abs(MasterWindowBottom - AreaBottom) < (STICKDISTANCEMAX * Screen.TwipsPerPixelY)) Then
        MasterWindow_Move2.YPosFixedFlag = True
        MasterWindow_Move2.YPos = (AreaBottom - (MasterWindowBottom - MasterWindowTop + 0 * Screen.TwipsPerPixelY)) / Screen.TwipsPerPixelY
    End If
End Function

'********************************END OF MASTER WINDOW**********************************
'************************************SLAVE WINDOW**************************************
'***SLAVE WINDOW POSING***
'NOTE: the SlaveWindow posing subs/functions can be used by the target project
'to restore a slave window's position (i.e. restore the relation to the MasterWindow).
'Example: the target project should use the SlaveWindow posing subs/function
'whenever a form is opened (teh form poses 'itself').
'
'Private Sub Form_Load
'   'on error resume next
'   Dim XPos As Long
'   Dim YPos As Long
'   If GFWindowStickfrm.GetSlaveWindowPosBest("Form1", XPos, YPos) = True Then 'always check return value
'        Call Form1.Move(XPos, YPos)
'   End If
'[...]

Public Function GetSlaveWindowPosBest(ByVal SlaveWindowName As String, ByRef XPos As Long, ByRef YPos As Long) As Boolean
    'on error resume next 'returns True if returned position is valid, False if there was an error
    'Dim StickTypeBitStructNumber As Integer 'global
    'Dim StickTypeBitStructArray() As StickTypeBitStruct 'global
    Dim MasterWindowIndex As Integer
    Dim SlaveWindowIndex As Integer
    Dim DistancePercentage As Integer
    Dim StickTypeBits As Long
    Dim StructLoop As Integer
    '
    'NOTE: this function returns the screen coordinates (in twips)
    'where the passed slave window should be placed at to be sticky
    'in the way it was sticky the last time the program was running
    '(i.e. like the last time StickTypeBitsToReg was called).
    'If this function returnes False then the slave window must not be moved.
    '
    'verify
    If GFWindowStickStructVar.GFWindowStickSystemEnabledFlag = False Then GoTo Error:
    'preset
    Call StickTypeBitStructFromReg(StickTypeBitStructNumber, StickTypeBitStructArray())
    MasterWindowIndex = GetMasterWindowIndex
    If MasterWindowIndex = 0 Then GoTo Error: 'verify
    SlaveWindowIndex = GetWindowStickStructIndex(SlaveWindowName + "(GFWindowStick)") 'for getting index the string "(GFWindowStick)" must be added (important when receiving messages of subclassing)
    If SlaveWindowIndex = 0 Then GoTo Error: 'verify
    'verify
    'NOTE: the following two lines lead to errors in a test project, don't use them (see also SetSlaveWindowPosBest()).
    'If (WindowStickStructArray(MasterWindowIndex).WindowObject.WindowState = vbNormal) Then GoTo Error: 'calculation would fail (tested)
    'If (WindowStickStructArray(SlaveWindowIndex).WindowObject.WindowState = vbNormal) Then GoTo Error: 'calculation would fail (tested)
    'preset
    YPos = UNDEFINED 'preset
    XPos = UNDEFINED 'preset
    'begin
    For StructLoop = 1 To StickTypeBitStructNumber
        If StickTypeBitStructArray(StructLoop).SlaveWindowName = SlaveWindowName Then
            StickTypeBits = StickTypeBitStructArray(StructLoop).SlaveWindowStickTypeBits
            If (StickTypeBits And STICKY) = 0 Then
                GetSlaveWindowPosBest = False 'error
                Exit Function 'important
            End If
            'NOTE: the handling of STICKY_INDIRECT is not supported.
            If (StickTypeBits And STICKY_AT_TOP) = STICKY_AT_TOP Then
                YPos = WindowStickStructArray(MasterWindowIndex).WindowObject.Top - WindowStickStructArray(SlaveWindowIndex).WindowObject.Height
            End If
            If (StickTypeBits And STICKY_AT_BOTTOM) = STICKY_AT_BOTTOM Then
                YPos = WindowStickStructArray(MasterWindowIndex).WindowObject.Top + WindowStickStructArray(MasterWindowIndex).WindowObject.Height
            End If
            If (StickTypeBits And STICKY_AT_LEFT) = STICKY_AT_LEFT Then
                XPos = WindowStickStructArray(MasterWindowIndex).WindowObject.Left - WindowStickStructArray(SlaveWindowIndex).WindowObject.Width
            End If
            If (StickTypeBits And STICKY_AT_RIGHT) = STICKY_AT_RIGHT Then
                XPos = WindowStickStructArray(MasterWindowIndex).WindowObject.Left + WindowStickStructArray(MasterWindowIndex).WindowObject.Width
            End If
            If (StickTypeBits And TOP_HEIGHT_STICKY) = TOP_HEIGHT_STICKY Then
                YPos = WindowStickStructArray(MasterWindowIndex).WindowObject.Top
            End If
            If (StickTypeBits And BOTTOM_HEIGHT_STICKY) = BOTTOM_HEIGHT_STICKY Then
                YPos = WindowStickStructArray(MasterWindowIndex).WindowObject.Top + WindowStickStructArray(MasterWindowIndex).WindowObject.Height - WindowStickStructArray(SlaveWindowIndex).WindowObject.Height 'vector-run
            End If
            If (StickTypeBits And LEFT_WIDTH_STICKY) = LEFT_WIDTH_STICKY Then
                XPos = WindowStickStructArray(MasterWindowIndex).WindowObject.Left
            End If
            If (StickTypeBits And RIGHT_WIDTH_STICKY) = RIGHT_WIDTH_STICKY Then
                XPos = WindowStickStructArray(MasterWindowIndex).WindowObject.Left + WindowStickStructArray(MasterWindowIndex).WindowObject.Width - WindowStickStructArray(SlaveWindowIndex).WindowObject.Width 'vector-run
            End If
            If XPos = UNDEFINED Then
                Call CopyMemory(ByVal VarPtr(DistancePercentage), ByVal (VarPtr(StickTypeBits) + 2), 2)
                XPos = CLng(((CSng(DistancePercentage) * WindowStickStructArray(MasterWindowIndex).WindowObject.Width) / 100!) + _
                     WindowStickStructArray(MasterWindowIndex).WindowObject.Left - (WindowStickStructArray(SlaveWindowIndex).WindowObject.Width / 2!))
            End If
            If YPos = UNDEFINED Then
                Call CopyMemory(ByVal VarPtr(DistancePercentage), ByVal (VarPtr(StickTypeBits) + 2), 2)
                YPos = CLng(((CSng(DistancePercentage) * WindowStickStructArray(MasterWindowIndex).WindowObject.Height) / 100!) + _
                     WindowStickStructArray(MasterWindowIndex).WindowObject.Top - (WindowStickStructArray(SlaveWindowIndex).WindowObject.Height / 2!))
            End If
            GetSlaveWindowPosBest = True 'ok
            Exit Function
        End If
    Next StructLoop
    'NOTE: function should already have been left in loop (except SlaveWindowName or registry entry is invalid).
    XPos = 0 'reset
    YPos = 0 'reset
    GetSlaveWindowPosBest = False 'error
    Exit Function
Error:
    XPos = 0 'reset (error)
    YPos = 0 'reset (error)
    GetSlaveWindowPosBest = False 'error
    Exit Function
End Function

Public Function SetSlaveWindowPosBest(ByVal SlaveWindowName As String, ByVal XPos As Long, ByVal YPos As Long) As Boolean
    'on error resume next 'returns True if window has been moved, False if not
    Dim SlaveWindowIndex As Integer
    'preset
    SlaveWindowIndex = GetWindowStickStructIndex(SlaveWindowName + "(GFWindowStick)") 'for getting index the string "(GFWindowStick)" must be added (important when receiving messages of subclassing)
    If SlaveWindowIndex = 0 Then GoTo Error: 'verify
    'verify
    If Not (WindowStickStructArray(SlaveWindowIndex).WindowObject.WindowState = vbNormal) Then
        '
        'NOTE: do not restore the window or there will be a recursion (tested).
        'If the slave window is minimized then the passed position calculated
        'by GetSlaveWindowPosBest() is not true anyway (tested).
        '
        GoTo Error:
    End If
    'begin
    WindowStickStructArray(SlaveWindowIndex).AboutMovingFlag = False 'reset (verify)
    If Not ((WindowStickStructArray(SlaveWindowIndex).WindowObject.Left = XPos) And (WindowStickStructArray(SlaveWindowIndex).WindowObject.Top = YPos)) Then 'verify
        Call WindowStickStructArray(SlaveWindowIndex).WindowObject.Move(XPos, YPos)
        Call SlaveWindow_FreezePosition 'important as moved slave window could have become sticky
    Else
        Call SlaveWindow_FreezePosition 'important as moved slave window could have become sticky
    End If
    SetSlaveWindowPosBest = True 'yeah
    Exit Function
Error:
    SetSlaveWindowPosBest = False 'error
    Exit Function
End Function

'***END OF SLAVE WINDOW POSING***

Private Function SlaveWindow_Move(ByVal WindowName As String, ByVal WindowObject As Form, ByVal MasterWindowLeft As Long, ByVal MasterWindowRight As Long, ByVal MasterWindowTop As Long, ByVal MasterWindowBottom As Long, ByVal SlaveWindowLeft As Long, ByVal SlaveWindowRight As Long, ByVal SlaveWindowTop As Long, ByVal SlaveWindowBottom As Long) As SlaveWindowMoveResultStruct
    'on error resume next
    Const STICKDISTANCEMAX As Long = 16
    '
    'NOTE: this sub checks if a slave window is near enough at the master window to
    'become sticky. If this is the case (the slave window is sticky), the new slave window x/y pos is returned.
    'The calling sub must set this window x/y pos through manipulating the parameters of the received
    'WM_WINDOWPOSCHANGING message.
    '
    'NOTE: do not call WindowObject.Top etc. but make calling sub set WindowObject's window position.
    '
    'NOTE: a slave window can currently be made sticky to the master window only
    '(but not to another slave window and also not to any screen edge).
    '
    'begin; fit to master window edge
    '
    If Abs(SlaveWindowBottom - MasterWindowTop) <= (STICKDISTANCEMAX * Screen.TwipsPerPixelY) Then
        If Not ((SlaveWindowRight < (MasterWindowLeft - STICKDISTANCEMAX * Screen.TwipsPerPixelX)) Or _
            (SlaveWindowLeft > (MasterWindowRight + STICKDISTANCEMAX * Screen.TwipsPerPixelX))) Then
            SlaveWindow_Move.YPos = (MasterWindowTop - WindowObject.Height) / Screen.TwipsPerPixelY
            SlaveWindow_Move.YPosFixedFlag = True
        End If
    End If
    If Abs(SlaveWindowTop - MasterWindowBottom) <= (STICKDISTANCEMAX * Screen.TwipsPerPixelY) Then
        If Not ((SlaveWindowRight < (MasterWindowLeft - STICKDISTANCEMAX * Screen.TwipsPerPixelX)) Or _
            (SlaveWindowLeft > (MasterWindowRight + STICKDISTANCEMAX * Screen.TwipsPerPixelX))) Then
            SlaveWindow_Move.YPos = (MasterWindowBottom / Screen.TwipsPerPixelY) + 1
            SlaveWindow_Move.YPosFixedFlag = True
        End If
    End If
    If Abs(SlaveWindowLeft - MasterWindowRight) <= (STICKDISTANCEMAX * Screen.TwipsPerPixelY) Then
        If Not ((SlaveWindowBottom < (MasterWindowTop - STICKDISTANCEMAX * Screen.TwipsPerPixelY)) Or _
            (SlaveWindowTop > (MasterWindowBottom + STICKDISTANCEMAX * Screen.TwipsPerPixelY))) Then
            SlaveWindow_Move.XPos = (MasterWindowRight / Screen.TwipsPerPixelX) + 1
            SlaveWindow_Move.XPosFixedFlag = True
        End If
    End If
    If Abs(SlaveWindowRight - MasterWindowLeft) <= (STICKDISTANCEMAX * Screen.TwipsPerPixelY) Then
        If Not ((SlaveWindowBottom < (MasterWindowTop - STICKDISTANCEMAX * Screen.TwipsPerPixelY)) Or _
            (SlaveWindowTop > (MasterWindowBottom + STICKDISTANCEMAX * Screen.TwipsPerPixelY))) Then
            SlaveWindow_Move.XPos = (MasterWindowLeft - WindowObject.Width) / Screen.TwipsPerPixelX
            SlaveWindow_Move.XPosFixedFlag = True
        End If
    End If
    If (SlaveWindow_Move.XPosFixedFlag = True) Or (SlaveWindow_Move.YPosFixedFlag = True) Then
        'fit to same master window x/y pos
        If Abs(SlaveWindowTop - MasterWindowTop) <= (STICKDISTANCEMAX * Screen.TwipsPerPixelY) Then
            SlaveWindow_Move.YPos = (MasterWindowTop) / Screen.TwipsPerPixelY
            SlaveWindow_Move.YPosFixedFlag = True
        End If
        If Abs(SlaveWindowLeft - MasterWindowLeft) <= (STICKDISTANCEMAX * Screen.TwipsPerPixelY) Then
            SlaveWindow_Move.XPos = (MasterWindowLeft) / Screen.TwipsPerPixelX
            SlaveWindow_Move.XPosFixedFlag = True
        End If
        If Abs(SlaveWindowRight - MasterWindowRight) <= (STICKDISTANCEMAX * Screen.TwipsPerPixelY) Then
            SlaveWindow_Move.XPos = (MasterWindowRight - WindowObject.Width) / Screen.TwipsPerPixelX + 1
            SlaveWindow_Move.XPosFixedFlag = True
        End If
        If Abs(SlaveWindowBottom - MasterWindowBottom) <= (STICKDISTANCEMAX * Screen.TwipsPerPixelY) Then
            SlaveWindow_Move.YPos = (MasterWindowBottom - WindowObject.Height) / Screen.TwipsPerPixelY + 1
            SlaveWindow_Move.YPosFixedFlag = True
        End If
    End If
End Function

Private Sub SlaveWindow_Size(ByVal WindowName As String, ByVal WindowObject As Form)
    'on error resume next
    'do nothing
End Sub

Private Sub SlaveWindow_FreezePosition()
    'on error resume next
    Dim MasterWindowIndex As Integer
    Dim StructLoop As Integer
    '
    'NOTE: you cannot use i.e. IsWindowSticky() directly when the
    'master window is moved as the gap resulting out of the move
    'would lead to a false result, the relation master-slave
    'window is set when moving begins and cannot change during
    'the move.
    '
    'preset
    MasterWindowIndex = GetMasterWindowIndex
    'begin
    For StructLoop = 1 To WindowStickStructNumber
        If Not (StructLoop = MasterWindowIndex) Then
            If (IsFormLoaded(WindowStickStructArray(StructLoop).WindowObject)) Then 'verify, don't load unnecessarily to save memory
                If WindowStickStructArray(StructLoop).WindowObject.WindowState = vbNormal Then 'verify (important)
                    WindowStickStructArray(StructLoop).IsWindowAtLeftFlag = IsWindowAtLeft(MasterWindowIndex, StructLoop)
                    WindowStickStructArray(StructLoop).IsWindowAtRightFlag = IsWindowAtRight(MasterWindowIndex, StructLoop)
                    WindowStickStructArray(StructLoop).IsWindowAtTopFlag = IsWindowAtTop(MasterWindowIndex, StructLoop)
                    WindowStickStructArray(StructLoop).IsWindowAtBottomFlag = IsWindowAtBottom(MasterWindowIndex, StructLoop)
                    WindowStickStructArray(StructLoop).IsWindowTopHeightStickyFlag = IsWindowTopHeightSticky(MasterWindowIndex, StructLoop)
                    WindowStickStructArray(StructLoop).IsWindowBottomHeightStickyFlag = IsWindowBottomHeightSticky(MasterWindowIndex, StructLoop)
                    WindowStickStructArray(StructLoop).IsWindowLeftWidthStickyFlag = IsWindowLeftWidthSticky(MasterWindowIndex, StructLoop)
                    WindowStickStructArray(StructLoop).IsWindowRightWidthStickyFlag = IsWindowRightWidthSticky(MasterWindowIndex, StructLoop)
                    WindowStickStructArray(StructLoop).IsWindowStickyFlag = IsWindowSticky(MasterWindowIndex, StructLoop)
                    WindowStickStructArray(StructLoop).IsWindowStickyIndirectFlag = IsWindowStickyIndirect(StructLoop)
                End If
            End If
        End If
    Next StructLoop
End Sub

'********************************END OF SLAVE WINDOW***********************************
'**************************************STICK TYPES*************************************
'NOTE: the following code is used to determinate and load and save the stick type of a window.

Private Sub StickTypeBitStructToReg(ByRef StickTypeBitStructNumber As Integer, ByRef StickTypeBitStructArray() As StickTypeBitStruct)
    'on error resume next
    Dim StructLoop As Integer
    '
    'NOTE: this sub does the following:
    '-write the name of the master window in registry
    '-write the names of the slave windows and its stick type bits into registry.
    'Through the registry entries it is possible to restore the current stick type.
    'Note that this registry system does not allow saving stick states for
    'indirect sticks, as it is supposed that all slave windows are sticky at the
    'master window (if they are sticky anyway).
    '
    'preset
    Call Rmod.RegDeleteSubKey(GFWindowStickStructVar.RegMainKey, GFWindowStickStructVar.RegRootKey + "GFWindowStick")
    Call Rmod.RegCreateSubKey(GFWindowStickStructVar.RegMainKey, GFWindowStickStructVar.RegRootKey + "GFWindowStick")
    'begin
    Call Rmod.RegSetKeyValue(GFWindowStickStructVar.RegMainKey, GFWindowStickStructVar.RegRootKey + "GFWindowStick", "master window name", CVar(GFWindowStickStructVar.MasterWindowName), REG_SZ)
    For StructLoop = 1 To StickTypeBitStructNumber
        'NOTE: to make code easier the master window name and also its stick type bits are written into registry (ignore when reading).
        'NOTE: under WinXP we can only create a sub key in (highest level) HKEY_LOCAL_MACHINE\Software\, not in HKEY_LOCAL_MACHINE\.
        Call Rmod.RegSetKeyValue(GFWindowStickStructVar.RegMainKey, GFWindowStickStructVar.RegRootKey + "GFWindowStick", "slave window name " + LTrim$(Str$(StructLoop)), CVar(StickTypeBitStructArray(StructLoop).SlaveWindowName), REG_SZ)
        Call Rmod.RegSetKeyValue(GFWindowStickStructVar.RegMainKey, GFWindowStickStructVar.RegRootKey + "GFWindowStick", "slave window stick type bits " + LTrim$(Str$(StructLoop)), CVar(StickTypeBitStructArray(StructLoop).SlaveWindowStickTypeBits), REG_SZ)
    Next StructLoop
End Sub

Private Sub StickTypeBitStructFromReg(ByRef StickTypeBitStructNumber As Integer, ByRef StickTypeBitStructArray() As StickTypeBitStruct)
    'on error resume next
    Dim MasterWindowName As String
    Dim StructLoop As Integer
    Dim Tempstr$
    '
    'NOTE: this sub does the following:
    '-read the master window name
    '-read all slave window names and stick type bits into structure (skip master window)
    '
    'reset
    StickTypeBitStructNumber = 0 'reset
    ReDim StickTypeBitStructArray(1 To 1) As StickTypeBitStruct
    'begin
    Rmod.RegGetKeyValueErrorFlag = False 'reset
    MasterWindowName = Rmod.RegGetKeyValue(GFWindowStickStructVar.RegMainKey, GFWindowStickStructVar.RegRootKey + "GFWindowStick", "master window name")
    If Rmod.RegGetKeyValueErrorFlag = True Then Exit Sub 'error
    '
    For StructLoop = 1 To 32766
        '
        Rmod.RegGetKeyValueErrorFlag = False 'reset
        Tempstr$ = Rmod.RegGetKeyValue(GFWindowStickStructVar.RegMainKey, GFWindowStickStructVar.RegRootKey + "GFWindowStick", "slave window name " + LTrim$(Str$(StructLoop)))
        If Rmod.RegGetKeyValueErrorFlag = True Then
            Exit Sub 'error
        Else
            StickTypeBitStructNumber = StickTypeBitStructNumber + 1 'cannot exceed 32767
            ReDim Preserve StickTypeBitStructArray(1 To StickTypeBitStructNumber) As StickTypeBitStruct
            StickTypeBitStructArray(StickTypeBitStructNumber).SlaveWindowName = Tempstr$
        End If
        '
        Rmod.RegGetKeyValueErrorFlag = False 'reset
        Tempstr$ = Rmod.RegGetKeyValue(GFWindowStickStructVar.RegMainKey, GFWindowStickStructVar.RegRootKey + "GFWindowStick", "slave window stick type bits " + LTrim$(Str$(StructLoop)))
        'If Not ((Val(Left$(Tempstr$, 8)) > 32767&) Or (Val(Left$(Tempstr$, 8)) < 0&)) Then 'verify (avoid overflow) 'no! (now Long value used)
            StickTypeBitStructArray(StickTypeBitStructNumber).SlaveWindowStickTypeBits = Val(Tempstr$)
        'Else
        '    StickTypeBitStructArray(StickTypeBitStructNumber).SlaveWindowStickTypeBits = 0 'error
        'End If
        '
    Next StructLoop
End Sub

Private Sub StickTypeBitStruct_Update(ByRef StickTypeBitStructNumber As Integer, ByRef StickTypeBitStructArray() As StickTypeBitStruct, ByVal WindowStickStructNumber As Integer, ByRef WindowStickStructArray() As WindowStickStruct)
    'on error resume next 'call to update the stick type bits of the passed window
    Dim SlaveWindowIndex As Integer
    Dim StructLoop As Integer
    'reset
    '
    'NOTE: resize array only if necessary to allow not saving stick type bits
    'of a minimized or maximized window.
    '
    'StickTypeBitStructNumber = 0 'reset
    'ReDim StickTypeBitStructArray(1 To 1) As StickTypeBitStruct
    'begin
    For StructLoop = 1 To WindowStickStructNumber
        If Not (StructLoop = GetMasterWindowIndex) Then
            '
            SlaveWindowIndex = SlaveWindowIndex + 1 'NOT equal to StructLoop (because of master window)
            If SlaveWindowIndex > StickTypeBitStructNumber Then 'can only happen if no rgistry entries existing
                StickTypeBitStructNumber = SlaveWindowIndex
                ReDim Preserve StickTypeBitStructArray(1 To StickTypeBitStructNumber) As StickTypeBitStruct
            End If
            '
            If (IsIconic(WindowStickStructArray(StructLoop).WindowObject.hwnd) = 0) And _
                (IsZoomed(WindowStickStructArray(StructLoop).WindowObject.hwnd) = 0) And _
                (WindowStickStructArray(StructLoop).WindowObject.Enabled = True) And _
                (WindowStickStructArray(StructLoop).WindowObject.Visible = True) Then 'verify
                '
                'NOTE: do not change stick type bits when the related window is not in normal state.
                'NOTE: we must not update any stick type bits when the related window
                'is not enabled and visible as otherwise the total garbage will be written
                'into the registry (tested).
                '
                StickTypeBitStructArray(SlaveWindowIndex).SlaveWindowName = WindowStickStructArray(StructLoop).WindowName
                StickTypeBitStructArray(SlaveWindowIndex).SlaveWindowStickTypeBits = GetWindowStickTypeBits(StructLoop)
            Else
                'save window name only, leave stick type bits in structure unchanged
                StickTypeBitStructArray(SlaveWindowIndex).SlaveWindowName = WindowStickStructArray(StructLoop).WindowName
            End If
            '
        End If
    Next StructLoop
End Sub

Private Function GetWindowStickTypeBits(ByVal WindowStickStructIndex As Integer) As Long
    'on error resume next 'returns an integer value with special bits set, depending on stick type of passed window
    Dim DistancePercentage As Integer
    'verify
    If (WindowStickStructIndex < 1) Or (WindowStickStructIndex > WindowStickStructNumber) Then 'verify
        MsgBox "internal error in GetStickyTypeBits() (GFWindowStick): passed value invalid !", vbOKOnly + vbExclamation
        GetWindowStickTypeBits = 0 'reset (error)
        Exit Function
    End If
    Call SlaveWindow_FreezePosition 'update IsWindowAt[...]Flags etc.
    'begin
    If WindowStickStructArray(WindowStickStructIndex).IsWindowStickyFlag = True Then _
        GetWindowStickTypeBits = GetWindowStickTypeBits Or STICKY
    If WindowStickStructArray(WindowStickStructIndex).IsWindowStickyIndirectFlag = True Then _
        GetWindowStickTypeBits = GetWindowStickTypeBits Or STICKY_INDIRECT 'although not supported yet
    If WindowStickStructArray(WindowStickStructIndex).IsWindowAtTopFlag = True Then _
        GetWindowStickTypeBits = GetWindowStickTypeBits Or STICKY_AT_TOP
    If WindowStickStructArray(WindowStickStructIndex).IsWindowAtBottomFlag = True Then _
        GetWindowStickTypeBits = GetWindowStickTypeBits Or STICKY_AT_BOTTOM
    If WindowStickStructArray(WindowStickStructIndex).IsWindowAtLeftFlag = True Then _
        GetWindowStickTypeBits = GetWindowStickTypeBits Or STICKY_AT_LEFT
    If WindowStickStructArray(WindowStickStructIndex).IsWindowAtRightFlag = True Then _
        GetWindowStickTypeBits = GetWindowStickTypeBits Or STICKY_AT_RIGHT
    If WindowStickStructArray(WindowStickStructIndex).IsWindowTopHeightStickyFlag = True Then _
        GetWindowStickTypeBits = GetWindowStickTypeBits Or TOP_HEIGHT_STICKY
    If WindowStickStructArray(WindowStickStructIndex).IsWindowBottomHeightStickyFlag = True Then _
        GetWindowStickTypeBits = GetWindowStickTypeBits Or BOTTOM_HEIGHT_STICKY
    If WindowStickStructArray(WindowStickStructIndex).IsWindowLeftWidthStickyFlag = True Then _
        GetWindowStickTypeBits = GetWindowStickTypeBits Or LEFT_WIDTH_STICKY
    If WindowStickStructArray(WindowStickStructIndex).IsWindowRightWidthStickyFlag = True Then _
        GetWindowStickTypeBits = GetWindowStickTypeBits Or RIGHT_WIDTH_STICKY
    '
    'NOTE: if window is not sticky in any master window corner, then save distance from
    'top/left master window corner to slave window top/left edge.
    'The distance is converted to a percentage and saved in the bits 9-16 of the
    'window stick type bits.
    '
    If (GetWindowStickTypeBits And STICKY) Then 'check if any distance is to be saved
        If (GetWindowStickTypeBits And STICKY_AT_TOP) Or (GetWindowStickTypeBits And STICKY_AT_BOTTOM) Or _
            (GetWindowStickTypeBits And TOP_HEIGHT_STICKY) Or (GetWindowStickTypeBits And BOTTOM_HEIGHT_STICKY) Then
            If Not ((GetWindowStickTypeBits And STICKY_AT_LEFT) Or (GetWindowStickTypeBits And STICKY_AT_RIGHT) Or _
                (GetWindowStickTypeBits And LEFT_WIDTH_STICKY) Or (GetWindowStickTypeBits And RIGHT_WIDTH_STICKY)) Then
                'calculate distance left master window edge <-> slave window x center
                DistancePercentage = CInt( _
                    ((WindowStickStructArray(WindowStickStructIndex).WindowObject.Left + _
                      WindowStickStructArray(WindowStickStructIndex).WindowObject.Width / 2) - _
                      WindowStickStructArray(GetMasterWindowIndex).WindowObject.Left) / _
                      WindowStickStructArray(GetMasterWindowIndex).WindowObject.Width * 100!)
            End If
        End If
        If (GetWindowStickTypeBits And STICKY_AT_LEFT) Or (GetWindowStickTypeBits And STICKY_AT_RIGHT) Or _
            (GetWindowStickTypeBits And LEFT_WIDTH_STICKY) Or (GetWindowStickTypeBits And RIGHT_WIDTH_STICKY) Then
            If Not ((GetWindowStickTypeBits And STICKY_AT_TOP) Or (GetWindowStickTypeBits And STICKY_AT_BOTTOM) Or _
                (GetWindowStickTypeBits And TOP_HEIGHT_STICKY) Or (GetWindowStickTypeBits And BOTTOM_HEIGHT_STICKY)) Then
                'calculate distance top master window edge <-> slave window y center
                DistancePercentage = CInt( _
                    ((WindowStickStructArray(WindowStickStructIndex).WindowObject.Top + _
                      WindowStickStructArray(WindowStickStructIndex).WindowObject.Height / 2) - _
                      WindowStickStructArray(GetMasterWindowIndex).WindowObject.Top) / _
                      WindowStickStructArray(GetMasterWindowIndex).WindowObject.Height * 100!)
            End If
        End If
    End If
    Call CopyMemory(ByVal (VarPtr(GetWindowStickTypeBits) + 2), ByVal VarPtr(DistancePercentage), 2)
End Function

Private Function IsWindowAtLeft(ByVal MasterWindowIndex As Integer, ByVal SlaveWindowIndex As Integer) As Boolean
    'on error resume next 'returns True if the slave window has the special position related to the master window
    'verify
    If (SlaveWindowIndex < 1) Or (SlaveWindowIndex > WindowStickStructNumber) Then 'verify
        IsWindowAtLeft = False
        Exit Function 'error
    End If
    'begin
    'NOTE: for all IsWindowAt[...] a gap of one pixel is tolerated to set flags right also in the case of rounding errors etc.
    If Abs(WindowStickStructArray(SlaveWindowIndex).WindowObject.Left - (WindowStickStructArray(MasterWindowIndex).WindowObject.Left - WindowStickStructArray(SlaveWindowIndex).WindowObject.Width)) <= Screen.TwipsPerPixelX Then
        IsWindowAtLeft = True
    Else
        IsWindowAtLeft = False
    End If
End Function

Private Function IsWindowAtTop(ByVal MasterWindowIndex As Integer, ByVal SlaveWindowIndex As Integer) As Boolean
    'on error resume next 'returns True if the slave window has the special position related to the master window
    'verify
    If (SlaveWindowIndex < 1) Or (SlaveWindowIndex > WindowStickStructNumber) Then 'verify
        IsWindowAtTop = False
        Exit Function 'error
    End If
    'begin
    If Abs(WindowStickStructArray(SlaveWindowIndex).WindowObject.Top - (WindowStickStructArray(MasterWindowIndex).WindowObject.Top - WindowStickStructArray(SlaveWindowIndex).WindowObject.Height)) <= Screen.TwipsPerPixelY Then
        IsWindowAtTop = True
    Else
        IsWindowAtTop = False
    End If
End Function

Private Function IsWindowAtRight(ByVal MasterWindowIndex As Integer, ByVal SlaveWindowIndex As Integer) As Boolean
    'on error resume next 'returns True if the slave window has the special position related to the master window
    'verify
    If (SlaveWindowIndex < 1) Or (SlaveWindowIndex > WindowStickStructNumber) Then 'verify
        IsWindowAtRight = False
        Exit Function 'error
    End If
    'begin
    If Abs(WindowStickStructArray(SlaveWindowIndex).WindowObject.Left - (WindowStickStructArray(MasterWindowIndex).WindowObject.Left + WindowStickStructArray(MasterWindowIndex).WindowObject.Width)) <= Screen.TwipsPerPixelX Then
        IsWindowAtRight = True
    Else
        IsWindowAtRight = False
    End If
End Function

Private Function IsWindowAtBottom(ByVal MasterWindowIndex As Integer, ByVal SlaveWindowIndex As Integer) As Boolean
    'on error resume next 'returns True if the slave window has the special position related to the master window
    'verify
    If (SlaveWindowIndex < 1) Or (SlaveWindowIndex > WindowStickStructNumber) Then 'verify
        IsWindowAtBottom = False
        Exit Function 'error
    End If
    'begin
    If Abs(WindowStickStructArray(SlaveWindowIndex).WindowObject.Top - (WindowStickStructArray(MasterWindowIndex).WindowObject.Top + WindowStickStructArray(MasterWindowIndex).WindowObject.Height)) <= Screen.TwipsPerPixelY Then 'one pixel space allowed
        IsWindowAtBottom = True
    Else
        IsWindowAtBottom = False
    End If
End Function

Private Function IsWindowTopHeightSticky(ByVal MasterWindowIndex As Integer, ByVal SlaveWindowIndex As Integer) As Boolean
    'on error resume next 'returns True if the slave window has the special position related to the master window
    'verify
    If (SlaveWindowIndex < 1) Or (SlaveWindowIndex > WindowStickStructNumber) Then 'verify
        IsWindowTopHeightSticky = False
        Exit Function 'error
    End If
    'begin
    If Abs(WindowStickStructArray(SlaveWindowIndex).WindowObject.Top - WindowStickStructArray(MasterWindowIndex).WindowObject.Top) <= Screen.TwipsPerPixelY Then 'one pixel space allowed
        'NOTE: a window must be sticky at right or left to be really top height sticky.
        If (IsWindowAtLeft(MasterWindowIndex, SlaveWindowIndex) = True) Or (IsWindowAtRight(MasterWindowIndex, SlaveWindowIndex) = True) Then
            IsWindowTopHeightSticky = True
        Else
            IsWindowTopHeightSticky = False
        End If
    Else
        IsWindowTopHeightSticky = False
    End If
End Function

Private Function IsWindowBottomHeightSticky(ByVal MasterWindowIndex As Integer, ByVal SlaveWindowIndex As Integer) As Boolean
    'on error resume next 'returns True if the slave window has the special position related to the master window
    'verify
    If (SlaveWindowIndex < 1) Or (SlaveWindowIndex > WindowStickStructNumber) Then 'verify
        IsWindowBottomHeightSticky = False
        Exit Function 'error
    End If
    'begin
    If Abs((WindowStickStructArray(SlaveWindowIndex).WindowObject.Top + WindowStickStructArray(SlaveWindowIndex).WindowObject.Height) - (WindowStickStructArray(MasterWindowIndex).WindowObject.Top + WindowStickStructArray(MasterWindowIndex).WindowObject.Height)) <= Screen.TwipsPerPixelY Then 'one pixel space allowed
        'NOTE: a window must be sticky at right or left to be really bottom height sticky.
        If (IsWindowAtLeft(MasterWindowIndex, SlaveWindowIndex) = True) Or (IsWindowAtRight(MasterWindowIndex, SlaveWindowIndex) = True) Then
            IsWindowBottomHeightSticky = True
        Else
            IsWindowBottomHeightSticky = False
        End If
    Else
        IsWindowBottomHeightSticky = False
    End If
End Function

Private Function IsWindowLeftWidthSticky(ByVal MasterWindowIndex As Integer, ByVal SlaveWindowIndex As Integer) As Boolean
    'on error resume next 'returns True if the slave window has the special position related to the master window
    'verify
    If (SlaveWindowIndex < 1) Or (SlaveWindowIndex > WindowStickStructNumber) Then 'verify
        IsWindowLeftWidthSticky = False
        Exit Function 'error
    End If
    'begin
    If Abs(WindowStickStructArray(SlaveWindowIndex).WindowObject.Left - WindowStickStructArray(MasterWindowIndex).WindowObject.Left) <= Screen.TwipsPerPixelX Then 'one pixel space allowed
        'NOTE: a window must be sticky at top or bottom to be really left height sticky.
        If (IsWindowAtTop(MasterWindowIndex, SlaveWindowIndex) = True) Or (IsWindowAtBottom(MasterWindowIndex, SlaveWindowIndex) = True) Then
            IsWindowLeftWidthSticky = True
        Else
            IsWindowLeftWidthSticky = False
        End If
    Else
        IsWindowLeftWidthSticky = False
    End If
End Function

Private Function IsWindowRightWidthSticky(ByVal MasterWindowIndex As Integer, ByVal SlaveWindowIndex As Integer) As Boolean
    'on error resume next 'returns True if the slave window has the special position related to the master window
    'verify
    If (SlaveWindowIndex < 1) Or (SlaveWindowIndex > WindowStickStructNumber) Then 'verify
        IsWindowRightWidthSticky = False
        Exit Function 'error
    End If
    'begin
    If Abs((WindowStickStructArray(SlaveWindowIndex).WindowObject.Left + WindowStickStructArray(SlaveWindowIndex).WindowObject.Width) - (WindowStickStructArray(MasterWindowIndex).WindowObject.Left + WindowStickStructArray(MasterWindowIndex).WindowObject.Width)) <= Screen.TwipsPerPixelX Then 'one pixel space allowed
        'NOTE: a window must be sticky at top or bottom to be really right height sticky.
        If (IsWindowAtTop(MasterWindowIndex, SlaveWindowIndex) = True) Or (IsWindowAtBottom(MasterWindowIndex, SlaveWindowIndex) = True) Then
            IsWindowRightWidthSticky = True
        Else
            IsWindowRightWidthSticky = False
        End If
    Else
        IsWindowRightWidthSticky = False
    End If
End Function

Public Function IsWindowSticky_Public(ByVal WindowName As String) As Boolean
    'on error resume next 'can be used by the target project to determinate is a slave window is sticky to the master window
    Dim MasterWindowIndex As Integer
    Dim SlaveWindowIndex As Integer
    'preset
    MasterWindowIndex = GetMasterWindowIndex
    If MasterWindowIndex = 0 Then GoTo Error: 'should not happen
    SlaveWindowIndex = GetWindowStickStructIndex(WindowName + "(GFWindowStick)")
    If SlaveWindowIndex = 0 Then GoTo Error:
    'begin
    IsWindowSticky_Public = IsWindowSticky(MasterWindowIndex, SlaveWindowIndex)
    Exit Function
Error:
    IsWindowSticky_Public = False 'error
    Exit Function
End Function

Private Function IsWindowSticky(ByVal MasterWindowIndex As Integer, ByVal SlaveWindowIndex As Integer) As Boolean
    'on error resume next
    If GFWindowStickStructVar.GFWindowStickSystemEnabledFlag = True Then
        If IsWindowAtTop(MasterWindowIndex, SlaveWindowIndex) = True Then IsWindowSticky = True: Exit Function
        If IsWindowAtBottom(MasterWindowIndex, SlaveWindowIndex) = True Then IsWindowSticky = True: Exit Function
        If IsWindowAtLeft(MasterWindowIndex, SlaveWindowIndex) = True Then IsWindowSticky = True: Exit Function
        If IsWindowAtRight(MasterWindowIndex, SlaveWindowIndex) = True Then IsWindowSticky = True: Exit Function
        If IsWindowTopHeightSticky(MasterWindowIndex, SlaveWindowIndex) = True Then IsWindowSticky = True: Exit Function
        If IsWindowBottomHeightSticky(MasterWindowIndex, SlaveWindowIndex) = True Then IsWindowSticky = True: Exit Function
        If IsWindowLeftWidthSticky(MasterWindowIndex, SlaveWindowIndex) = True Then IsWindowSticky = True: Exit Function
        If IsWindowRightWidthSticky(MasterWindowIndex, SlaveWindowIndex) = True Then IsWindowSticky = True: Exit Function
    Else
        IsWindowSticky = False
    End If
End Function

Private Function IsWindowStickyIndirect(ByVal SlaveWindowIndexPassed As Integer) As Boolean
    'on error resume next 'retruns True if slave window sticks either directly at master window, or if there are one or more sticks slave windows between
    Dim SlaveWindowIndex As Integer
    Dim SlaveWindowIndexOld As Integer
    Dim StructLoop As Integer
    '
    'NOTE: there may be only one 'longer path' to the master window,
    'if the current window is surrounded by slave windows only this
    'function will hang up (at the moment, must be improved).
    '
    IsWindowStickyIndirect = False 'fuck!
    Exit Function 'this function does not work (damn it!)
    'preset
    SlaveWindowIndex = SlaveWindowIndexPassed
    SlaveWindowIndexOld = 0
    IsWindowStickyIndirect = False 'preset
    'begin
ReDo:
    For StructLoop = 1 To WindowStickStructNumber
        If Not ((StructLoop = SlaveWindowIndex) Or (StructLoop = SlaveWindowIndexOld)) Then
            If IsWindowSticky(SlaveWindowIndex, StructLoop) = True Then
                If StructLoop = GetMasterWindowIndex() Then
                    IsWindowStickyIndirect = True
                    Exit Function
                End If
            Else
                SlaveWindowIndexOld = SlaveWindowIndex 'avoid 'going back'
                SlaveWindowIndex = StructLoop
                GoTo ReDo:
            End If
        End If
    Next StructLoop
End Function

'**********************************END OF STICK TYPES**********************************
'***********************************GENERAL FUNCTIONS**********************************

Private Function GFMoveMinimizedWindow(ByVal WindowHandle As Long, ByVal WindowXPosNew As Long, ByVal WindowYPosNew As Long) As Boolean
    'on error resume next 'window will appear at the given position when restored (size stays unchanged); returns True if successful, False in case of an error
    Dim WINDOWPLACEMENTVar As WINDOWPLACEMENT
    Dim WindowHeightUnchanged As Long
    Dim WindowWidthUnchanged As Long
    'verify
    If IsIconic(WindowHandle) = 0& Then
        GFMoveMinimizedWindow = False 'error
        Exit Function
    End If
    'preset
    WINDOWPLACEMENTVar.Length = Len(WINDOWPLACEMENTVar)
    'begin
    Call GetWindowPlacement(WindowHandle, WINDOWPLACEMENTVar)
    '
    WindowWidthUnchanged = WINDOWPLACEMENTVar.rcNormalPosition.Right - WINDOWPLACEMENTVar.rcNormalPosition.Left
    WindowHeightUnchanged = WINDOWPLACEMENTVar.rcNormalPosition.Bottom - WINDOWPLACEMENTVar.rcNormalPosition.Top
    '
    WINDOWPLACEMENTVar.Flags = WPF_SETMINPOSITION
    WINDOWPLACEMENTVar.showCmd = SW_SHOWNA
    '
    WINDOWPLACEMENTVar.rcNormalPosition.Left = WindowXPosNew
    WINDOWPLACEMENTVar.rcNormalPosition.Top = WindowYPosNew
    WINDOWPLACEMENTVar.rcNormalPosition.Right = WindowXPosNew + WindowWidthUnchanged
    WINDOWPLACEMENTVar.rcNormalPosition.Bottom = WindowYPosNew + WindowHeightUnchanged
    '
    GFMoveMinimizedWindow = CBool(SetWindowPlacement(WindowHandle, WINDOWPLACEMENTVar))
    Exit Function
End Function

'*******************************END OF GENERAL FUNCTIONS*******************************
'****************************************OTHER*****************************************

Private Function GetWindowStickStructIndex(ByVal SourceDescription As String) As Integer
    'on error resume next 'returns index or 0 for error; call this functio out of GFSubClassWindowProc()
    Dim SourceDescriptionLength As Long
    Dim StructLoop As Integer
    'preset
    GetWindowStickStructIndex = 0 'preset (error)
    'verify
    If Not (Right$(SourceDescription, 15) = "(GFWindowStick)") Then
        Exit Function 'error
    Else
        SourceDescription = Left$(SourceDescription, Len(SourceDescription) - 15)
    End If
    'begin
    SourceDescriptionLength = Len(SourceDescription)
    For StructLoop = 1 To WindowStickStructNumber
        If SourceDescriptionLength = Len(WindowStickStructArray(StructLoop).WindowName) Then
            If SourceDescription = WindowStickStructArray(StructLoop).WindowName Then
                GetWindowStickStructIndex = StructLoop
                Exit Function 'ok
            End If
        End If
    Next StructLoop
End Function

Private Function GetWindowStateChange(ByVal WindowStickStructIndex As Integer) As Integer
    'on error resume next 'returns a WINDOWSTATECHANGE constant
    '
    'NOTE: call this function only only per message processing as Old-flags
    'are set, calling this function twice will lead to returning false values.
    '
    If Not ((WindowStickStructIndex < 1) Or (WindowStickStructIndex > WindowStickStructNumber)) Then 'verify
        If (IsIconic(WindowStickStructArray(WindowStickStructIndex).WindowObject.hwnd) = 0) And (IsZoomed(WindowStickStructArray(WindowStickStructIndex).WindowObject.hwnd) = 0) Then
            If (WindowStickStructArray(WindowStickStructIndex).IsIconicFlagOld = True) Or (WindowStickStructArray(WindowStickStructIndex).IsZoomedFlagOld = True) Then
                GetWindowStateChange = WINDOWSTATECHANGE_WASRESTORED
                GoTo Jump:
            End If
        End If
        If (IsIconic(WindowStickStructArray(WindowStickStructIndex).WindowObject.hwnd)) Then
            If (WindowStickStructArray(WindowStickStructIndex).IsIconicFlagOld = False) Then
                GetWindowStateChange = WINDOWSTATECHANGE_WASMINIMIZED
                GoTo Jump:
            End If
        End If
        If (IsZoomed(WindowStickStructArray(WindowStickStructIndex).WindowObject.hwnd)) Then
            If (WindowStickStructArray(WindowStickStructIndex).IsZoomedFlagOld = False) Then
                GetWindowStateChange = WINDOWSTATECHANGE_WASMAXIMIZED
                GoTo Jump:
            End If
        End If
        GetWindowStateChange = WINDOWSTATECHANGE_NOCHANGE 'no change existing
Jump:
        WindowStickStructArray(WindowStickStructIndex).IsIconicFlagOld = CBool(IsIconic(WindowStickStructArray(WindowStickStructIndex).WindowObject.hwnd))
        WindowStickStructArray(WindowStickStructIndex).IsZoomedFlagOld = CBool(IsZoomed(WindowStickStructArray(WindowStickStructIndex).WindowObject.hwnd))
    Else
        GetWindowStateChange = WINDOWSTATECHANGE_NOCHANGE 'error
        Exit Function
    End If
End Function

Private Function IsFormLoaded(ByRef FormObject As Form) As Boolean
    'on error resume next 'check return value of this function before accessing Form.Visible or so to avoid permanent reloading
    Dim FormLoop As Integer
    'begin
    For FormLoop = 0 To Forms.Count - 1
        If Forms(FormLoop) Is FormObject Then
            IsFormLoaded = True
            Exit Function
        End If
    Next FormLoop
    IsFormLoaded = False
    Exit Function
End Function

Private Function TX(ByVal PixelsX As Long) As Long
    'on error resume next
    TX = PixelsX * Screen.TwipsPerPixelX
End Function

Private Function TY(ByVal PixelsY As Long) As Long
    'on error resume next
    TY = PixelsY * Screen.TwipsPerPixelY
End Function

'************************************END OF OTHER**************************************

Public Sub GFWindowStick_Terminate()
    'on error resume next
    'NOTE: the target project must call GFSubClass_Terminate when unloading.
End Sub

