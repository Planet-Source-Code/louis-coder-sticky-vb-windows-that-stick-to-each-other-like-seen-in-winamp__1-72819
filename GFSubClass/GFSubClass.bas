Attribute VB_Name = "GFSubClassmod"
Option Explicit
'(c)2001-2003 by Louis. Use to subclass one or more controls/windows.
'
'THIS MODULE IS PLUG-IN CODE, DO NOT CHANGE!
'
#Const GFSubClassSystemDisabledFlag = False 'enable for target project debugging only
#Const GFSubClassWindowProcExEnabledFlag = False
#Const JamLockEnabledFlag = False 'see code annotations
Const JamLockMsg As Long = 31 'see code annotations
'
'NOTE: about GFSubClassWindowProcEx():
'In the MP3 Renamer 2 project mysterious slow-downs appeared when
'calling TargetFormArray().GFSubClassWindowProc().
'Therefore GFSubClassWindowProcEx() was implemented.
'If the related switch is enabled, GFSubClassmod will call
'Mfrm.GFSubClassWindowProcEx(), this sub has the task to call
'the GFSubClassWindowProc() of the form whose hWnd was passed
'as argument.
'In this way no reference must be saved in an object var and thus
'the slow-downs should not appear.
'
'Public Sub GFSubClassWindowProcEx(ByVal TargetFormName As String, ByVal SourceDescription As String, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef ReturnValue As Long, ByRef ReturnValueUsedFlag As Boolean)
'   'on error resume next
'End Sub
'
'NOTE: about JamLock ((c)2001 by Louis):
'If Msg JamLockMsg arrives three (coherent) times the GFSubClass code
'does not call any call back sub any more, then VB is able to open an error message box.
'Insert Debug.Print Msg in GFSubClassProc() to determinate JamLockMsg.
'
'NOTE: the target form (must not be GFSubClassmod) must contain the following sub:
'Public Sub GFSubClassWindowProc(ByVal SourceDescription As String, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef ReturnValue As Long, ByRef ReturnValueUsedFlag As Boolean)
'    'on error resume next
'End Sub
'
'It is valid to assign an own target form/module to every subclassed control/window.
'It is also valid to subclass an object several times, assigning different target forms
'(the GFSubClass system will call all target forms in the order they have been added).
'
'BUG: something's wrong inside here: if a VB control is subclassed
'under two different names and one name is removed and then re-added,
'no messages will be filtered for the re-added control any more.
'
'IMPORTANT: all GFSubClassWindowProc() subs of the target project should
'be left immediately if the passed message is not to be processed.
'Do not perform any action in those subs that is not required.
'Declare vars only when required, declare no vars if passed message is not
'to be processed.
'
'GFSubClass
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
'GFSubClass_GetParent
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
'GFSubClass
Private Const GWL_WNDPROC = (-4)
Private Const GWL_USERDATA = (-21)
'GFSubClassStruct - Information about subclassed controls/windows
Private Type GFSubClassStruct
    TargetObjectDescription As String
    TargetObjecthWndOld As Long
    TargetObjecthWndOldPassedFlag As Boolean 'if hWnd was PASSED to GFSubClass()
    TargetObjecthWndNew As Long
    TargetObjectSubClassedFlag As Boolean 'if subclassing was enabled
    TargetNumber As Integer
    TargetObjectSubClassEnabledFlagArray() As Boolean 'if message is to be forwarded to current target form
    TargetObject As Object 'for verifying that no control is subclassed twice
    TargetFormIndexArray() As Integer
    TargetFormNameArray() As String
End Type
Dim GFSubClassStructNumber As Integer
Dim GFSubClassStructArray() As GFSubClassStruct
'TargetForm
Dim TargetFormNumber As Integer
Dim TargetFormArray() As Object
'MessageRestoreStruct - see code for details
Private Type MessageRestoreStruct
    SourceDescription As String
    hwnd As Long
    Msg As Long
    wParam As Long
    lParam As Long
End Type
Dim MessageRestoreStructNumber As Integer
Dim MessageRestoreStructArray() As MessageRestoreStruct
Dim MessageRestore_CheckCalledFlag As Boolean
Dim MessageRestore_BroadcastMsgCalledFlag As Boolean
'JamLock
Dim JamLockMsg1 As Long
Dim JamLockMsg2 As Long
Dim JamLockMsg3 As Long
Dim JamLockMsgPointer As Integer

'*************************************SUB CLASSING*************************************

Public Sub GFSubClass(ByRef TargetObject As Object, ByVal TargetObjectDescription As String, ByRef TargetForm As Object, ByVal SubClassEnabledFlag As Boolean, Optional ByVal TargetObjecthWnd As Long = 0, Optional ByVal CallTargetFormAtFirstFlag As Boolean = False)
    'on error Resume Next 'if TargetObjecthWnd is not 0, TargetObject can be Nothing (use if a reference to TargetObject is not available)
    Dim GFSubClassStructPointer As Integer
    Dim GFSubClassTargetPointer As Integer
    Dim StructLoop As Integer
    '
    'IMPORTANT: TargetObject/TargetObjectHandle and TargetObjectDescription must always
    'be related to each other, that means it is not valid to subclass e.g. TAGfrm.TAGListView
    'once under the name "TAGListView" and then under the name "TAGfrm.TAGListView"
    'as then the subclass procedure will start calling itself when a message from TAGListView
    'arrives.
    '
    'verify
    #If GFSubClassSystemDisabledFlag = True Then
        Exit Sub
    #End If
    'begin
    GFSubClassStructPointer = GetGFSubClassStructPointer(TargetObjectDescription)
    If GFSubClassStructPointer = 0 Then
        '
        'NOTE: the same object may be subclassed several times under
        'different names.
        '
        'create new array element
        If Not (GFSubClassStructNumber = 32766) Then 'verify
            GFSubClassStructNumber = GFSubClassStructNumber + 1
            GFSubClassStructPointer = GFSubClassStructNumber
        Else
            MsgBox "internal error in GFSubClass(): overflow !", vbOKOnly + vbExclamation
            Exit Sub 'error
        End If
        ReDim Preserve GFSubClassStructArray(1 To GFSubClassStructNumber) As GFSubClassStruct
        GFSubClassStructArray(GFSubClassStructNumber).TargetObjectDescription = TargetObjectDescription
        GFSubClassStructArray(GFSubClassStructNumber).TargetNumber = 1 'preset
        GFSubClassTargetPointer = 1
        'add target form
        ReDim Preserve GFSubClassStructArray(GFSubClassStructPointer).TargetObjectSubClassEnabledFlagArray(1 To GFSubClassStructArray(GFSubClassStructPointer).TargetNumber) As Boolean
        ReDim Preserve GFSubClassStructArray(GFSubClassStructPointer).TargetFormIndexArray(1 To GFSubClassStructArray(GFSubClassStructPointer).TargetNumber) As Integer
        ReDim Preserve GFSubClassStructArray(GFSubClassStructPointer).TargetFormNameArray(1 To GFSubClassStructArray(GFSubClassStructPointer).TargetNumber) As String
        GFSubClassStructArray(GFSubClassStructPointer).TargetObjectSubClassEnabledFlagArray(GFSubClassStructArray(GFSubClassStructPointer).TargetNumber) = SubClassEnabledFlag
        Set GFSubClassStructArray(GFSubClassStructPointer).TargetObject = TargetObject
        Call TargetForm_Add(TargetForm)
        GFSubClassStructArray(GFSubClassStructPointer).TargetFormIndexArray(GFSubClassStructArray(GFSubClassStructPointer).TargetNumber) = GetTargetFormIndex(TargetForm)
        GFSubClassStructArray(GFSubClassStructPointer).TargetFormNameArray(GFSubClassStructArray(GFSubClassStructPointer).TargetNumber) = TargetForm.Name
        'enable subclassing
        '
        'NOTE: an object is subclassed until GFSubClass_Terminate is called,
        'but the GFSubClass system does not call the call back sub of the target form
        'if the related SubClassEnabledFlag is set to False.
        '
        If TargetObjecthWnd = 0 Then
            GFSubClassStructArray(GFSubClassStructPointer).TargetObjecthWndOld = TargetObject.hwnd
            GFSubClassStructArray(GFSubClassStructPointer).TargetObjecthWndOldPassedFlag = False
        Else
            'NOTE: it is possible to pass the hWnd of an object only if it was not created by VB.
            GFSubClassStructArray(GFSubClassStructPointer).TargetObjecthWndOld = TargetObjecthWnd
            GFSubClassStructArray(GFSubClassStructPointer).TargetObjecthWndOldPassedFlag = True
        End If
        'For StructLoop = 1 To (GFSubClassStructNumber - 1) 'exclude currently added element; this code doesn't work correctly
        '    If GFSubClassStructArray(StructLoop).TargetObjecthWndOld = GFSubClassStructArray(GFSubClassStructPointer).TargetObjecthWndOld Then
        '        If (GFSubClassStructArray(StructLoop).TargetObject Is GFSubClassStructArray(GFSubClassStructPointer).TargetObject) Or _
        '            (GFSubClassStructArray(StructLoop).TargetObject Is Nothing) Or (GFSubClassStructArray(GFSubClassStructPointer).TargetObject Is Nothing) Then
        '            'transfer data from already subclassed control to just created control
        '            GFSubClassStructArray(GFSubClassStructPointer).TargetObjecthWndNew = GFSubClassStructArray(StructLoop).TargetObjecthWndNew
        '            GFSubClassStructArray(GFSubClassStructPointer).TargetObjectSubClassedFlag = True
        '            GoTo AlreadySubClassed:
        '        End If
        '    End If
        'Next StructLoop
        If (GetWindowLong(GFSubClassStructArray(GFSubClassStructPointer).TargetObjecthWndOld, GWL_USERDATA)) Then 'the relation between already-subclased marking and control is definite
            GFSubClassStructArray(GFSubClassStructPointer).TargetObjecthWndNew = GetWindowLong(GFSubClassStructArray(GFSubClassStructPointer).TargetObjecthWndOld, GWL_USERDATA)
            GFSubClassStructArray(GFSubClassStructPointer).TargetObjectSubClassedFlag = True
            GoTo AlreadySubClassed:
        End If
        If GFSubClassStructArray(GFSubClassStructPointer).TargetObjectSubClassedFlag = False Then 'True if a control was subclassed twice under two different names
            '
            'NOTE: if an object is subclassed twice (called SetWindowLong() two times) then endless loops
            'in the message system will appear, leading to a program crash or through serious slow-downs.
            '
            GFSubClassStructArray(GFSubClassStructPointer).TargetObjecthWndNew = SetWindowLong(GFSubClassStructArray(GFSubClassStructPointer).TargetObjecthWndOld, GWL_WNDPROC, AddressOf GFSubClassmod.GFSubClassWindowProc)
            Call SetWindowLong(GFSubClassStructArray(GFSubClassStructPointer).TargetObjecthWndOld, GWL_USERDATA, GFSubClassStructArray(GFSubClassStructPointer).TargetObjecthWndNew)
            GFSubClassStructArray(GFSubClassStructPointer).TargetObjectSubClassedFlag = True
        End If
AlreadySubClassed:
        'end of enabling subclassing
    Else
        GFSubClassTargetPointer = GetTargetPointer(GFSubClassStructPointer, TargetForm)
        If GFSubClassTargetPointer = 0 Then
            'add a target form to an existing array element
            If Not (GFSubClassStructArray(GFSubClassStructPointer).TargetNumber = 32766) Then 'verify
                GFSubClassStructArray(GFSubClassStructPointer).TargetNumber = GFSubClassStructArray(GFSubClassStructPointer).TargetNumber + 1
                GFSubClassTargetPointer = GFSubClassStructArray(GFSubClassStructPointer).TargetNumber
            Else
                MsgBox "internal error in GFSubClass(): overflow (2) !", vbOKOnly + vbExclamation
                Exit Sub 'error
            End If
            ReDim Preserve GFSubClassStructArray(GFSubClassStructPointer).TargetObjectSubClassEnabledFlagArray(1 To GFSubClassStructArray(GFSubClassStructPointer).TargetNumber) As Boolean
            ReDim Preserve GFSubClassStructArray(GFSubClassStructPointer).TargetFormIndexArray(1 To GFSubClassStructArray(GFSubClassStructPointer).TargetNumber) As Integer
            ReDim Preserve GFSubClassStructArray(GFSubClassStructPointer).TargetFormNameArray(1 To GFSubClassStructArray(GFSubClassStructPointer).TargetNumber) As String
            If CallTargetFormAtFirstFlag = False Then
                GFSubClassStructArray(GFSubClassStructPointer).TargetObjectSubClassEnabledFlagArray(GFSubClassStructArray(GFSubClassStructPointer).TargetNumber) = SubClassEnabledFlag
                Call TargetForm_Add(TargetForm)
                GFSubClassStructArray(GFSubClassStructPointer).TargetFormIndexArray(GFSubClassStructArray(GFSubClassStructPointer).TargetNumber) = GetTargetFormIndex(TargetForm)
                GFSubClassStructArray(GFSubClassStructPointer).TargetFormNameArray(GFSubClassStructArray(GFSubClassStructPointer).TargetNumber) = TargetForm.Name
            Else
                'NOTE: the newly added target form will be called at first.
                GFSubClassStructArray(GFSubClassStructPointer).TargetObjectSubClassEnabledFlagArray(GFSubClassStructArray(GFSubClassStructPointer).TargetNumber) = SubClassEnabledFlag
                For StructLoop = GFSubClassStructArray(GFSubClassStructPointer).TargetNumber To 1 Step (-1)
                    If Not (StructLoop = 1) Then
                        GFSubClassStructArray(GFSubClassStructPointer).TargetFormIndexArray(StructLoop) = GFSubClassStructArray(GFSubClassStructPointer).TargetFormIndexArray(StructLoop - 1)
                        GFSubClassStructArray(GFSubClassStructPointer).TargetFormNameArray(StructLoop) = GFSubClassStructArray(GFSubClassStructPointer).TargetFormNameArray(StructLoop - 1)
                    Else
                        Call TargetForm_Add(TargetForm)
                        GFSubClassStructArray(GFSubClassStructPointer).TargetFormIndexArray(1) = GetTargetFormIndex(TargetForm)
                        GFSubClassStructArray(GFSubClassStructPointer).TargetFormNameArray(1) = TargetForm.Name
                    End If
                Next StructLoop
            End If
        Else
            'enable/disable sub class enabled flag
            GFSubClassStructArray(GFSubClassStructPointer).TargetObjectSubClassEnabledFlagArray(GFSubClassTargetPointer) = SubClassEnabledFlag
        End If
    End If
    Exit Sub
End Sub

Private Function GetTargetObjectCount(ByRef TargetObject As Object) As Integer
    'on error resume next 'returns how often a control to subclass appears in GFSubClassStructArray()
    Dim ObjectCount As Integer
    Dim StructLoop As Integer
    '
    'NOTE: use this function to determinate if SetWindowLong()
    'must be used or if it has already been used.
    '
    'begin
    For StructLoop = 1 To GFSubClassStructNumber
        If GFSubClassStructArray(StructLoop).TargetObject Is TargetObject Then
            ObjectCount = ObjectCount + 1
        End If
    Next StructLoop
    GetTargetObjectCount = ObjectCount
End Function

Public Sub GFSubClass_UnSubclass(ByVal TargetObjectDescription As String, ByRef TargetForm As Object)
    'on error resume next
    Dim GFSubClassStructPointer As Integer
    Dim GFSubClassTargetPointer As Integer
    Dim StructLoop As Integer
    'begin
    GFSubClassStructPointer = GetGFSubClassStructPointer(TargetObjectDescription)
    If Not (GFSubClassStructPointer = 0) Then
        GFSubClassTargetPointer = GetTargetPointer(GFSubClassStructPointer, TargetForm)
        If Not (GFSubClassTargetPointer = 0) Then 'verify
            Select Case GFSubClassStructArray(GFSubClassStructPointer).TargetNumber
            Case 0, 1 '0 should not happen
                'disable the one and only target form, unsubclass object
                If GFSubClassStructArray(GFSubClassStructPointer).TargetObjectSubClassedFlag = True Then
                    'object has been subclassed
                    GFSubClassStructArray(GFSubClassStructPointer).TargetObjectSubClassedFlag = False 'reset
                    If GetTargetObjectCount(GFSubClassStructArray(GFSubClassStructPointer).TargetObject) = 1 Then
                        '
                        'NOTE: use SetWindowLong() only if object to subclass is not
                        'still registered under a different name.
                        '
                        Call SetWindowLong(GFSubClassStructArray(GFSubClassStructPointer).TargetObjecthWndOld, GWL_WNDPROC, GFSubClassStructArray(GFSubClassStructPointer).TargetObjecthWndNew)
                        Call SetWindowLong(GFSubClassStructArray(GFSubClassStructPointer).TargetObjecthWndOld, GWL_USERDATA, 0) 'reset
                    End If
                End If
                'remove object from structure
                For StructLoop = GFSubClassStructPointer To GFSubClassStructNumber
                    If Not (StructLoop = GFSubClassStructNumber) Then
                        GFSubClassStructArray(StructLoop) = GFSubClassStructArray(StructLoop + 1)
                    Else
                        GFSubClassStructNumber = GFSubClassStructNumber - 1
                        StructLoop = GFSubClassStructNumber 'StructLoop is not used any more
                        If StructLoop < 1 Then StructLoop = 1 'verify
                        ReDim Preserve GFSubClassStructArray(1 To StructLoop) As GFSubClassStruct
                        Exit For
                    End If
                Next StructLoop
            Case Else
                'disable one target form (remove it from current control's target form structure)
                For StructLoop = GFSubClassTargetPointer To GFSubClassStructArray(GFSubClassStructPointer).TargetNumber
                    If Not (StructLoop = GFSubClassStructArray(GFSubClassStructPointer).TargetNumber) Then
                        GFSubClassStructArray(GFSubClassStructPointer).TargetFormIndexArray(StructLoop) = GFSubClassStructArray(GFSubClassStructPointer).TargetFormIndexArray(StructLoop + 1)
                        GFSubClassStructArray(GFSubClassStructPointer).TargetFormNameArray(StructLoop) = GFSubClassStructArray(GFSubClassStructPointer).TargetFormNameArray(StructLoop + 1)
                    Else
                        GFSubClassStructArray(GFSubClassStructPointer).TargetNumber = GFSubClassStructArray(GFSubClassStructPointer).TargetNumber - 1
                        StructLoop = GFSubClassStructArray(GFSubClassStructPointer).TargetNumber
                        If StructLoop < 1 Then StructLoop = 1 'verify
                        ReDim Preserve GFSubClassStructArray(GFSubClassStructPointer).TargetFormIndexArray(1 To StructLoop) As Integer
                        ReDim Preserve GFSubClassStructArray(GFSubClassStructPointer).TargetFormNameArray(1 To StructLoop) As String
                        Exit For
                    End If
                Next StructLoop
            End Select
        Else 'target form invalid
            'MsgBox "internal error in GFSubClass_Terminate(): passed value invalid !", vbOKOnly + vbExclamation
        End If
    Else 'control name invalid
        'MsgBox "internal error in GFSubClass_Terminate(): passed value invalid !", vbOKOnly + vbExclamation
    End If
End Sub

Public Sub GFSubClass_Terminate()
    'on error Resume Next 'must be called through Form_Unload() when project is quit
    Dim StructLoop As Integer
    'begin
    For StructLoop = 1 To GFSubClassStructNumber
        If GFSubClassStructArray(StructLoop).TargetObjectSubClassedFlag = True Then
            'object has been subclassed
            GFSubClassStructArray(StructLoop).TargetObjectSubClassedFlag = False 'reset
            Call SetWindowLong(GFSubClassStructArray(StructLoop).TargetObjecthWndOld, GWL_WNDPROC, GFSubClassStructArray(StructLoop).TargetObjecthWndNew)
            Call SetWindowLong(GFSubClassStructArray(StructLoop).TargetObjecthWndOld, GWL_USERDATA, 0) 'reset
        End If
    Next StructLoop
    'reset
    GFSubClassStructNumber = 0 'reset
    ReDim GFSubClassStructArray(1 To 1) As GFSubClassStruct 'reset
End Sub

Public Function GFSubClass_IsSubClassed(ByVal TargetObjectDescription As String, ByRef TargetForm As Object) As Boolean
    'on error resume next 'returns True if passed object is subclassed, False if not
    Dim GFSubClassStructPointer As Integer
    '
    'NOTE: if TargetForm is Nothing, then this function returns True if messages
    'of the passed control are forwarded to any form of the target project,
    'if TargetForm is a reference to a form then this function returns True if messages
    'are forwarded to TargetForm.
    '
    'preset
    GFSubClassStructPointer = GetGFSubClassStructPointer(TargetObjectDescription)
    'begin
    If (GFSubClassStructPointer) Then
        If TargetForm Is Nothing Then
            GFSubClass_IsSubClassed = True
        Else
            If (GetTargetPointer(GFSubClassStructPointer, TargetForm)) Then
                GFSubClass_IsSubClassed = True
            Else
                GFSubClass_IsSubClassed = False
            End If
        End If
    Else
        GFSubClass_IsSubClassed = False
    End If
End Function

Public Function GFSubClass_GetParent(ByVal ChildhWnd As Long) As Long
    'on error resume next 'usable when messages are sent to the parent of a control
    GFSubClass_GetParent = GetParent(ChildhWnd)
End Function

Public Sub GFSubClass_ShowTargetFormNames(ByVal TargetObjectDescription As String)
    'on error resume next 'impllemented for debugging only, cannot have any function in a compiled executable
    Dim GFSubClassStructPointer As Integer
    Dim TargetLoop As Integer
    'preset
    GFSubClassStructPointer = GetGFSubClassStructPointer(TargetObjectDescription)
    If GFSubClassStructPointer = 0 Then
        Debug.Print "NO TARGET FORMS FOR " + TargetObjectDescription + ", OBJECT NOT SUBCLASSED."
        Exit Sub
    End If
    'begin
    Debug.Print "TARGET FORMS FOR " + TargetObjectDescription + ":"
    For TargetLoop = 1 To GFSubClassStructArray(GFSubClassStructPointer).TargetNumber
        Debug.Print TargetFormArray(GFSubClassStructArray(GFSubClassStructPointer).TargetFormIndexArray(TargetLoop)).Name 'every form object has a name
    Next TargetLoop
    Debug.Print "END OF TARGET FORMS."
    Exit Sub
End Sub

Private Function GetGFSubClassStructPointer(ByVal TargetObjectDescription As String) As Integer
    'on error Resume Next 'returns struct index or 0 for error
    Dim StructLoop As Integer
    For StructLoop = 1 To GFSubClassStructNumber
        If Len(GFSubClassStructArray(StructLoop).TargetObjectDescription) = Len(TargetObjectDescription) Then 'check first to increase speed
            If GFSubClassStructArray(StructLoop).TargetObjectDescription = TargetObjectDescription Then
                GetGFSubClassStructPointer = StructLoop 'ok
                Exit Function
            End If
        End If
    Next StructLoop
    GetGFSubClassStructPointer = 0 'error
    Exit Function
End Function

Private Function GetTargetPointer(ByVal GFSubClassStructPointer As Integer, ByRef TargetForm As Object) As Integer
    'on error Resume Next 'returns target form index or 0 for object not existing
    Dim TargetLoop As Integer
    'verify
    If (GFSubClassStructPointer < 1) Or (GFSubClassStructPointer > GFSubClassStructNumber) Then
        GetTargetPointer = 0 'error
        Exit Function
    End If
    'begin
    For TargetLoop = 1 To GFSubClassStructArray(GFSubClassStructPointer).TargetNumber
        If TargetFormArray(GFSubClassStructArray(GFSubClassStructPointer).TargetFormIndexArray(TargetLoop)) Is TargetForm Then
            GetTargetPointer = TargetLoop 'ok
            Exit Function
        End If
    Next TargetLoop
    GetTargetPointer = 0 'error
    Exit Function
End Function

Public Function GFSubClassWindowProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    'on error Resume Next 'module code of GFSubClass
    Dim GFSubClassStructPointer As Integer
    Dim ReturnValue As Long
    Dim ReturnValueUsedFlag As Boolean
    Dim TargetLoop As Integer
    Dim StructLoop As Integer
    'begin
    '
    'NOTE: passed hWnd is the old handle of the subclassed control/window.
    'NOTE: the form that receives the message at last has the highest priority
    'in setting the return value.
    '
    #If JamLockEnabledFlag = True Then
        JamLockMsgPointer = JamLockMsgPointer + 1
        If JamLockMsgPointer > 3 Then JamLockMsgPointer = 1
        Select Case JamLockMsgPointer
        Case 1
            JamLockMsg1 = Msg
        Case 2
            JamLockMsg2 = Msg
        Case 3
            JamLockMsg3 = Msg
        End Select
        If (JamLockMsg1 = JamLockMsg2) And (JamLockMsg1 = JamLockMsg3) And (JamLockMsg1 = JamLockMsg) Then
            Debug.Print "JAM!"
            JamLockMsgPointer = -1
        End If
    #End If
    '
    For StructLoop = 1 To GFSubClassStructNumber
        If GFSubClassStructArray(StructLoop).TargetObjecthWndOld = hwnd Then
            GFSubClassStructPointer = StructLoop
            If JamLockMsgPointer = -1 Then GoTo Jam: 'after setting GFSubClassStructPointer
            '
            'NOTE: the same object could be subclassed twice under two names,
            'then the message must be forwarded twice.
            '
            'NOTE: the message is sent to all registered target forms.
            'The one that was first registered receives the message at first.
            '
            'NOTE: the following loop code is also implemented in MessageRestore_Process.
            '
            For TargetLoop = 1 To GFSubClassStructArray(GFSubClassStructPointer).TargetNumber
                '
                If Not ((GFSubClassStructPointer < 1) Or (GFSubClassStructPointer > GFSubClassStructNumber)) Then 'verify (important, for every target loop)
                '
                    If (GFSubClassStructArray(GFSubClassStructPointer).TargetObjectSubClassEnabledFlagArray(TargetLoop)) Then
                        'target form is enabled for receiving messages, send message
                        #If GFSubClassWindowProcExEnabledFlag = True Then
                            Call Mfrm.GFSubClassWindowProcEx( _
                                GFSubClassStructArray(GFSubClassStructPointer).TargetFormNameArray(TargetLoop), _
                                GFSubClassStructArray(StructLoop).TargetObjectDescription, _
                                hwnd, Msg, wParam, lParam, ReturnValue, ReturnValueUsedFlag)
                        #Else
                            Call TargetFormArray(GFSubClassStructArray(GFSubClassStructPointer).TargetFormIndexArray(TargetLoop)).GFSubClassWindowProc( _
                                GFSubClassStructArray(StructLoop).TargetObjectDescription, _
                                hwnd, Msg, wParam, lParam, ReturnValue, ReturnValueUsedFlag)
                        #End If
                    End If
                End If
            Next TargetLoop
        End If
    Next StructLoop
    '
    If Not ((GFSubClassStructPointer < 1) Or (GFSubClassStructPointer > GFSubClassStructNumber)) Then 'verify (important)
        If ReturnValueUsedFlag = False Then
Jam:
            GFSubClassWindowProc = CallWindowProc(GFSubClassStructArray(GFSubClassStructPointer).TargetObjecthWndNew, hwnd, Msg, wParam, lParam)
        Else
            '
            'NOTE: for some means it is important to return a special value
            '(e.g. 0 when processed the WM_DROPFILES message) to prevent
            'the OS from crashing.
            '
            GFSubClassWindowProc = ReturnValue
        End If
    End If
    #If JamLockEnabledFlag = True Then
        If Not (JamLockMsgPointer = -1) Then Call MessageRestore_Process
    #Else
        Call MessageRestore_Process
    #End If
End Function

'*********************************END OF SUB CLASSING**********************************
'**************************************RESUBCLASS**************************************
'NOTE: you can automatically re-subclass a form that has been unloaded and loaded again.
'This feature is important for e.g. the GFSkinEngine.

Public Sub GFSubClass_ReSubClassByTargetObjectDescriptionPrefix(ByVal TargetObjectDescriptionPrefix As String)
    'on error resume next
    Dim StructLoop As Integer
    '
    'NOTE: this sub automatically re-subclasses ALL controls of a defined form,
    'including the form itself. The controls must have the form name in their names,
    'e.g. "Extrafrm", "Extrafrm.ExtraFrame" etc.
    'Use this sub when a form has been unloaded and loaded again, and it would be
    'too complicated to do all the GFSubClass() calls again.
    '
    'begin
    For StructLoop = 1 To GFSubClassStructNumber
        If InStr(1, GFSubClassStructArray(StructLoop).TargetObjectDescription, TargetObjectDescriptionPrefix, vbBinaryCompare) = 1 Then
            'Debug.Print "RESUBCLASS: " + GFSubClassStructArray(StructLoop).TargetObjectDescription
            'If GFSubClassStructArray(StructLoop).TargetObjectSubClassedFlag = False Then 'the same object may be subclassed under two names, but update both references to object
                If GFSubClassStructArray(StructLoop).TargetObjecthWndOldPassedFlag = False Then
                    '
                    'NOTE: if an object is subclassed twice (called SetWindowLong() two times) then endless loops
                    'in the message system will appear, leading to a program crash or through serious slow-downs.
                    '
                    Set GFSubClassStructArray(StructLoop).TargetObject = GetObjectByName(GFSubClassStructArray(StructLoop).TargetObjectDescription)
                    If Not (GFSubClassStructArray(StructLoop).TargetObject Is Nothing) Then 'verify
                        'important, update hWnd, old one is not valid any more after form unloading
                        GFSubClassStructArray(StructLoop).TargetObjecthWndOld = GFSubClassStructArray(StructLoop).TargetObject.hwnd
                        If (GetWindowLong(GFSubClassStructArray(StructLoop).TargetObjecthWndOld, GWL_USERDATA)) Then 'check if control is 'physically' subclassed
                            'already subclassed under an other name
                            GFSubClassStructArray(StructLoop).TargetObjecthWndNew = GetWindowLong(GFSubClassStructArray(StructLoop).TargetObjecthWndOld, GWL_USERDATA)
                            GFSubClassStructArray(StructLoop).TargetObjectSubClassedFlag = True
                        Else
                            'not subclassed yet, subclass now
                            GFSubClassStructArray(StructLoop).TargetObjecthWndNew = SetWindowLong(GFSubClassStructArray(StructLoop).TargetObject.hwnd, GWL_WNDPROC, AddressOf GFSubClassmod.GFSubClassWindowProc)
                            Call SetWindowLong(GFSubClassStructArray(StructLoop).TargetObject.hwnd, GWL_USERDATA, GFSubClassStructArray(StructLoop).TargetObjecthWndNew)
                            GFSubClassStructArray(StructLoop).TargetObjectSubClassedFlag = True
                        End If
                    Else
                        'Debug.Print "OBJECT NOT FOUND"
                        'object cannot be re-subclassed
                    End If
                    '
                Else
                    'Debug.Print "HWND PASSED"
                    'object cannot be subclassed
                End If
            'Else
            '    Debug.Print "ALREADY SUBCLASSED"
            'End If
        End If
    Next StructLoop
End Sub

Private Function GetObjectByName(ByVal ObjectName As String) As Object
    On Error Resume Next 'important
    Dim Control As Object
    Dim ControlName As String
    Dim ControlLoop As Integer
    Dim FormName As String
    Dim FormLoop As Integer
    Dim Temp As Long
    'preset
    '
    'NOTE: the GFWindowStick system uses the TargetObjectName format
    'Form.Control(GFWindowStick). Other users of GFSubClassmod should also
    'append extended data in brackets if necessary. The GFSubClass code
    'will find the correct target object also with extended data in brackets.
    '
    Temp = InStr(1, ObjectName, "(", vbBinaryCompare)
    If (Temp > 0) Then
        ObjectName = Left$(ObjectName, Temp - 1)
    End If
    'begin
    For FormLoop = 0 To Forms.Count - 1
        FormName = "" 'reset
        FormName = Forms(FormLoop).Name 'if fails (object has no .Name property) then FormName stays ""
        If FormName = ObjectName Then
            Set GetObjectByName = Forms(FormLoop)
            Exit Function
        End If
        If FormName = Left$(ObjectName, Len(FormName)) Then 'don't search form controls if form wrong
            For Each Control In Forms(FormLoop).Controls 'from VB help
                ControlName = "" 'reset
                ControlName = Control.Name 'if fails (object has no .Name property) then ControlName stays ""
                If Forms(FormLoop).Name + "." + ControlName = ObjectName Then 'Form.Control naming system expected
                    Set GetObjectByName = Control
                    Exit Function
                End If
            Next
        End If
'        For ControlLoop = 1 To Forms(FormLoop).Controls.Count - 1 'failed (?!?)
'            ControlName = "" 'reset
'            ControlName = Forms(FormLoop).Controls(ControlLoop).Name 'if fails (object has no .Name property) then ControlName stays ""
'            If ControlName = ObjectName Then
'                Set GetObjectByName = Forms(FormLoop).Controls(ControlLoop)
'                Exit Function
'            End If
'        Next ControlLoop
    Next FormLoop
    Set GetObjectByName = Nothing
    Exit Function
End Function

Public Sub GFSubClass_ReSubClass_UnSubClassByTargetObjectDescriptionPrefix(ByVal TargetObjectDescriptionPrefix As String)
    'on error resume next
    Dim StructLoop As Integer
    '
    'NOTE: this sub automatically re-subclasses ALL controls of a defined form,
    'including the form itself. The controls must have the form name in their names,
    'e.g. "Extrafrm", "Extrafrm.ExtraFrame" etc.
    'Use this sub when a form has been unloaded and loaded again, and it would be
    'too complicated to do all the GFSubClass() calls again.
    '
    'begin
    For StructLoop = 1 To GFSubClassStructNumber
        If InStr(1, GFSubClassStructArray(StructLoop).TargetObjectDescription, TargetObjectDescriptionPrefix, vbBinaryCompare) = 1 Then
            If GFSubClassStructArray(StructLoop).TargetObjectSubClassedFlag = True Then 'verify
                'object has been subclassed
                GFSubClassStructArray(StructLoop).TargetObjectSubClassedFlag = False 'reset
                Call SetWindowLong(GFSubClassStructArray(StructLoop).TargetObjecthWndOld, GWL_WNDPROC, GFSubClassStructArray(StructLoop).TargetObjecthWndNew)
                Call SetWindowLong(GFSubClassStructArray(StructLoop).TargetObjecthWndOld, GWL_USERDATA, 0) 'reset
            End If
        End If
    Next StructLoop
End Sub

'**********************************END OF RESUBCLASS***********************************
'**************************************TARGETFORM**************************************
'NOTE: informaer times we called
'GFSubClassStructArray(GFSubClassStructPointer).TargetFormArray(TargetLoop).GFSubClassWindowProc().
'As this lead to mysterious slow downs when 'jumping' to GFSubClassWindowProc() we now call
'TargetFormArray(GFSubClassStructArray(GFSubClassStructPointer).TargetFormIndexArray(TargetLoop)).GFSubClassWindowProc().

Private Sub TargetForm_Add(ByRef TargetForm As Object)
    'on error resume next
    'verify
    If (GetTargetFormIndex(TargetForm)) Then Exit Sub 'already added
    'begin
    If Not (TargetFormNumber = 32766) Then
        TargetFormNumber = TargetFormNumber + 1
        ReDim Preserve TargetFormArray(1 To TargetFormNumber) As Object
        Set TargetFormArray(TargetFormNumber) = TargetForm
    Else
        MsgBox "internal error in TargetForm_Add() (GFSubClass): overflow !", vbOKOnly + vbExclamation
    End If
End Sub

Private Function GetTargetFormIndex(ByRef TargetForm As Object) As Integer
    'on error resume next 'returns TargetFormArray() index or 0 for error
    Dim StructLoop As Integer
    'begin
    For StructLoop = 1 To TargetFormNumber
        If TargetFormArray(StructLoop) Is TargetForm Then
            GetTargetFormIndex = StructLoop 'ok
            Exit Function
        End If
    Next StructLoop
    GetTargetFormIndex = 0 'error
    Exit Function
End Function

'**********************************END OF TARGETFORM***********************************
'***********************************MESSAGE RESTORE************************************
'NOTE: the target project can call MessageRestore_AddMsg() to simulate the arriving
'of a message when GFSubClassWindowProc() is left.
'This can be very useful if code e.g. expects WM_LBUTTONUP and WM_LBUTTONDOWN
'mesages appearing always as a pair, but a WMLBUTTONUP message would not
'arrive as the target project opens another window at the WM_LBUTTONDOWN message.

Public Sub MessageRestore_AddMsg(ByVal SourceDescription As String, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long)
    'on error resume next
    If Not (MessageRestoreStructNumber = 32766) Then 'verify
        MessageRestoreStructNumber = MessageRestoreStructNumber + 1
    Else
        MsgBox "internal error in MessageRestore_AddMsg(): overflow !", vbOKOnly + vbExclamation
        Exit Sub
    End If
    ReDim Preserve MessageRestoreStructArray(1 To MessageRestoreStructNumber) As MessageRestoreStruct
    MessageRestoreStructArray(MessageRestoreStructNumber).SourceDescription = SourceDescription
    MessageRestoreStructArray(MessageRestoreStructNumber).hwnd = hwnd
    MessageRestoreStructArray(MessageRestoreStructNumber).Msg = Msg
    MessageRestoreStructArray(MessageRestoreStructNumber).wParam = wParam
    MessageRestoreStructArray(MessageRestoreStructNumber).lParam = lParam
End Sub

Public Sub MessageRestore_Process()
    'on error resume next
    Dim ReturnValue As Long
    Dim ReturnValueUsedFlag As Boolean
    Dim GFSubClassStructPointer As Integer
    Dim TargetLoop As Integer
    Dim StructLoop As Integer
    'verify
    If MessageRestore_CheckCalledFlag = True Then
        Exit Sub
    Else
        MessageRestore_CheckCalledFlag = True
    End If
    'begin
    For StructLoop = 1 To MessageRestoreStructNumber
        GFSubClassStructPointer = GetGFSubClassStructPointer( _
            MessageRestoreStructArray(StructLoop).SourceDescription)
        If GFSubClassStructPointer = 0 Then GoTo Jump: 'verify
        For TargetLoop = 1 To GFSubClassStructArray(GFSubClassStructPointer).TargetNumber
            If GFSubClassStructArray(GFSubClassStructPointer).TargetObjectSubClassEnabledFlagArray(TargetLoop) = True Then
                'target form is enabled for receiving messages, send message
                #If GFSubClassWindowProcExEnabledFlag = True Then
                    Call Mfrm.GFSubClassWindowProcEx( _
                        GFSubClassStructArray(GFSubClassStructPointer).TargetFormNameArray(TargetLoop), _
                        MessageRestoreStructArray(StructLoop).SourceDescription, _
                        MessageRestoreStructArray(StructLoop).hwnd, _
                        MessageRestoreStructArray(StructLoop).Msg, _
                        MessageRestoreStructArray(StructLoop).wParam, _
                        MessageRestoreStructArray(StructLoop).lParam, _
                        ReturnValue, ReturnValueUsedFlag)
                #Else
                    Call TargetFormArray(GFSubClassStructArray(GFSubClassStructPointer).TargetFormIndexArray(TargetLoop)).GFSubClassWindowProc( _
                        MessageRestoreStructArray(StructLoop).SourceDescription, _
                        MessageRestoreStructArray(StructLoop).hwnd, _
                        MessageRestoreStructArray(StructLoop).Msg, _
                        MessageRestoreStructArray(StructLoop).wParam, _
                        MessageRestoreStructArray(StructLoop).lParam, _
                        ReturnValue, ReturnValueUsedFlag)
                #End If
            End If
        Next TargetLoop
Jump:
    Next StructLoop
    'reset
    MessageRestore_CheckCalledFlag = False 'reset
    If (MessageRestoreStructNumber) Then
        MessageRestoreStructNumber = 0 'reset
        ReDim MessageRestoreStructArray(1 To 1) As MessageRestoreStruct
    End If
End Sub

Public Sub MessageRestore_BroadcastMsg(ByVal SourceDescription As String, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef ReturnValue As Long, ByRef ReturnValueUsedFlag As Boolean)
    'on error resume next 'return value forwarded
    Dim GFSubClassStructPointer As Integer
    Dim TargetLoop As Integer
    Dim StructLoop As Integer
    '
    'NOTE: the target project can call this sub to instantly send a message within the current project.
    'The message will be sent instantly. Use this sub to send e.g. a WM_CANCELMODE message,
    'which is usually used to abort any action if a pop up menu is opened
    '(send this message if e.g. a MsgBox is opened by the target project).
    '
    'verify
    If MessageRestore_BroadcastMsgCalledFlag = True Then
        Exit Sub
    Else
        MessageRestore_BroadcastMsgCalledFlag = True
    End If
    'begin
    GFSubClassStructPointer = GetGFSubClassStructPointer(SourceDescription)
    If GFSubClassStructPointer = 0 Then GoTo Jump: 'verify
    For TargetLoop = 1 To GFSubClassStructArray(GFSubClassStructPointer).TargetNumber
        If GFSubClassStructArray(GFSubClassStructPointer).TargetObjectSubClassEnabledFlagArray(TargetLoop) = True Then
            'target form is enabled for receiving messages, send message
            #If GFSubClassWindowProcExEnabledFlag = True Then
                Call Mfrm.GFSubClassWindowProcEx( _
                    GFSubClassStructArray(GFSubClassStructPointer).TargetFormNameArray(TargetLoop), _
                    SourceDescription, _
                    hwnd, _
                    Msg, _
                    wParam, _
                    lParam, _
                    ReturnValue, _
                    ReturnValueUsedFlag)
            #Else
                Call TargetFormArray(GFSubClassStructArray(GFSubClassStructPointer).TargetFormIndexArray(TargetLoop)).GFSubClassWindowProc( _
                    SourceDescription, _
                    hwnd, _
                    Msg, _
                    wParam, _
                    lParam, _
                    ReturnValue, ReturnValueUsedFlag)
            #End If
        End If
    Next TargetLoop
Jump:
    'reset
    MessageRestore_BroadcastMsgCalledFlag = False 'reset
End Sub

'*******************************END OF MESSAGE RESTORE*********************************

