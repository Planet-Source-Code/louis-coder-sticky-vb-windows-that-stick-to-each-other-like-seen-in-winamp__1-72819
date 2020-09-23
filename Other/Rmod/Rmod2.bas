Attribute VB_Name = "Rmod"
Option Explicit
'(c)1999 - 2002 by Louis.
'
'THIS MODULE IS PLUG-IN CODE, DO NOT CHANGE!
'
'NOTE: this is the second version of Rmod with bugs removed.
'Use this module as a general function (module).
'
'RegGetKeyValue
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
'GetRegKeyValueList; source: www.matthart.com (registry.zip)
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
'GetRegSubKeyList
'IMPORTANT: use BYVAL Reserved As Long and not just Reserved As Long or the function will fail on WinXP!
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
'GFSubKeyList
Dim GFSubKeyListNumber As Integer
Dim GFSubKeyListArray() As String
Dim GFSubKeyDir_RegMainKey As Long
Dim GFSubKeyDir_RegSubKey As String 'current (complete) sub key name
Dim GFSubKeyDir_ListCount As Integer '"SubKeyDir" (replaces GFDirectoryListDir)
Dim GFSubKeyDir_List() As String '"SubKeyDir" (replaces GFDirectoryListDir)
'RegSetKeyValue[String/Long/Byte]
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
'RegCreateSubKey
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
'RegDeleteKeyValue
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
'RegDeleteSubKey
'NOTE: this function is public as it deletes the whole key, not only one value (as the module functions do).
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
'other
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
'RegMainKey constants
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
'RegValueType constants
Public Const REG_SZ = 1&
Public Const REG_BINARY = 3&
Public Const REG_DWORD = 4&
'Error constants
Private Const ERROR_SUCCESS As Long = 0
Private Const ERROR_NO_MORE_ITEMS As Long = 259
'Key access (and related) constants
Private Const SYNCHRONIZE = &H100000
Private Const READ_CONTROL = &H20000
Private Const READ_WRITE = 2
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const KEY_SET_VALUE = &H2
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_EVENT = &H1
Private Const KEY_NOTIFY = &H10
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
'RegGetSubKeyList
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
'RegCreateSubKey
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
'error flags (reset before calling related function)
Public RegGetKeyValueErrorFlag As Boolean
Public RegSetKeyValueErrorFlag As Boolean
Public RegSetKeyValueCreateSubKeyCalledFlag As Boolean
Public RegCreateSubKeyErrorFlag As Boolean
Public RegDeleteValueErrorFlag As Boolean
Public RegDeleteKeyErrorFlag As Boolean
'
'NOTE: the functions in this module are used to manipulate any
'Win95/98 Registry setting.
'
'The following data describes the location of a registry value:
'RegMainKey: constant numeric value, e.g. HKEY_LOCAL_MACHINE
'RegSubKey: string containing name of sub key, e.g. Software\Microsoft\
'RegValueName: string containing the name related to the value to get/set
'RegValueValue: value to get/set (type sometimes included in variable name)
'
'Note that you must differ between three value types: String, Long and Byte.
'Not all functions support other types than String.
'
'********************************REGISTRY GET FUNCTIONS********************************
'NOTE: the following subs/functions are used to retrieve registry data.

Public Function RegGetKeyValue(ByVal RegMainKey As Long, ByVal RegSubKey As String, ByVal RegValueName As String) As String
    'on error Resume Next 'use to obtain a key value
    Dim RegKeyHandle As Long
    Dim RegValueType As Long
    Dim RegValueString As String
    Dim RegValueStringLength As Long
    Dim Temp As Long
    Dim Tempsngl!
    Dim Tempstr$
    'preset
    RegValueStringLength = 128000
    RegValueString = String$(RegValueStringLength, Chr$(0))
    'begin
    Temp = RegOpenKeyEx(RegMainKey, RegSubKey, 0&, KEY_READ, RegKeyHandle)
    If Temp = ERROR_SUCCESS Then 'verify
        RegGetKeyValueErrorFlag = False 'ok
    Else
        GoTo Error:
    End If
    Temp = RegQueryValueEx(RegKeyHandle, RegValueName, ByVal 0&, RegValueType, ByVal RegValueString, RegValueStringLength)
    If Temp = ERROR_SUCCESS Then 'verify
        Select Case RegValueType
        Case REG_DWORD
            'return the number as a string
            Call CopyMemory(Temp, ByVal RegValueString, 4)
            RegGetKeyValue = LTrim$(Str$(Temp))
        Case REG_BINARY
            'return the binary data like in a hex editor
            For Temp = 1 To RegValueStringLength 'VC help says size of null termination is excluded
                If Len(Mid$(RegValueString, Temp, 1)) Then 'verify (RegValueStringLength is gabage, Windows sucks)
                    Tempstr$ = Hex$(Asc(Mid$(RegValueString, Temp, 1)))
                    If Len(Tempstr$) = 1 Then Tempstr$ = "0" + Tempstr$
                    RegGetKeyValue = RegGetKeyValue + Tempstr$ + " "
                End If
            Next Temp
            RegGetKeyValue = Trim$(RegGetKeyValue)
        Case Else 'e.g. REG_SZ
            'return the string just as string (without null termination)
            If (InStr(1, RegValueString, Chr$(0), vbBinaryCompare) - 1) Then 'verify
                'NOTE: we do not process the value of RegValueStringLength as it isn't always correct (tested).
                RegGetKeyValue = Left$(RegValueString, (InStr(1, RegValueString, Chr$(0), vbBinaryCompare) - 1)) 'cut null-termination
            End If
        End Select
    Else
        GoTo Error:
    End If
    Call RegCloseKey(RegKeyHandle)
    RegGetKeyValueErrorFlag = False
    Exit Function
Error:
    Call RegCloseKey(RegKeyHandle) 'make sure handle is closed
    RegGetKeyValueErrorFlag = True
    RegGetKeyValue = "" 'reset (error)
    Exit Function
End Function

Public Function RegGetKeyValueList(ByVal RegMainKey As Long, ByVal RegSubKey As String, ByRef RegValueNumber As Integer, ByRef RegValueNameArray() As String, ByRef RegValueValueArray() As String) As Boolean
    'on error Resume Next 'returns True for success, False for error; initializes passed array with all value names and values related to the passed registry key
    Dim RegKeyHandle As Long
    Dim RegDataType As Long
    Dim RegValueIndex As Integer
    Dim RegValueName As String
    Dim RegValueNameLength As Long
    Dim RegValueSetting As String
    Dim RegValueSettingLength As Long
    Dim Temp1 As Long
    Dim Temp2 As Long
    Dim Tempsngl!
    'reset
    RegValueNumber = 0 'reset
    ReDim RegValueNameArray(1 To 1) As String
    ReDim RegValueValueArray(1 To 1) As String
    'preset
    RegValueIndex = 0 'reset; used to number values in registry key
    RegValueNameLength = 1024 'preset
    RegValueName = String$(RegValueNameLength, Chr$(0))
    RegValueSettingLength = 1024 'preset
    RegValueSetting = String$(RegValueNameLength, Chr$(0))
    'begin
    Temp1 = Rmod.RegOpenKeyEx(RegMainKey, RegSubKey, 0, KEY_READ, RegKeyHandle)
    If Not (Temp1 = ERROR_SUCCESS) Then 'error opening key (maybe does not exist)
        GoTo Error:
    End If
    Temp2 = Rmod.RegEnumValue(RegKeyHandle, RegValueIndex, RegValueName, RegValueNameLength, 0, RegDataType, RegValueSetting, RegValueSettingLength)
    If Not ((Temp2 = ERROR_SUCCESS) Or (Temp2 = ERROR_NO_MORE_ITEMS)) Then GoTo Error:
    'NOTE: to not change RegValueName and RegValueSetting, or the program will crash (pointer usage).
    Do While Not ((Temp2 = ERROR_NO_MORE_ITEMS) Or (RegValueIndex = 32766)) 'avoid endless loop through overflow error
        Select Case RegDataType
        Case 3 'REG_BINARY
            '[not supported] (did not work, see NN99 code or search Internet if binary values need to be read)
        Case 4 'REG_DWORD
            '[not supported] (did not work, see NN99 code or search Internet if DWORD values need to be read)
        Case Else 'STRING
            'nothing to do
        End Select
        RegValueNumber = RegValueNumber + 1
        ReDim Preserve RegValueNameArray(1 To RegValueNumber) As String
        ReDim Preserve RegValueValueArray(1 To RegValueNumber) As String
        If Not (InStr(1, RegValueName, Chr$(0), vbBinaryCompare) = 0) Then 'verify
            RegValueNameArray(RegValueNumber) = Left$(RegValueName, InStr(1, RegValueName, Chr$(0), vbBinaryCompare) - 1) 'do not use length value returned by API function (garbage)
        Else
            RegValueNameArray(RegValueNumber) = RegValueName
        End If
        If Not (InStr(1, RegValueSetting, Chr$(0), vbBinaryCompare) = 0) Then 'verify
            RegValueValueArray(RegValueNumber) = Left$(RegValueSetting, InStr(1, RegValueSetting, Chr$(0), vbBinaryCompare) - 1) 'do not use length value returned by API function (garbage)
        Else
            RegValueValueArray(RegValueNumber) = RegValueSetting
        End If
        RegValueNameLength = 1024 'reset
        RegValueSettingLength = 1024 'reset
        RegValueIndex = RegValueIndex + 1
        Temp2 = RegEnumValue(RegKeyHandle, RegValueIndex, RegValueName, RegValueNameLength, 0, RegDataType, RegValueSetting, RegValueSettingLength)
        If Not ((Temp2 = ERROR_SUCCESS) Or (Temp2 = ERROR_NO_MORE_ITEMS)) Then GoTo Error:
    Loop
    Call Rmod.RegCloseKey(RegKeyHandle)
    RegGetKeyValueList = True 'ok
    Exit Function
Error:
    RegGetKeyValueList = False 'error
    Exit Function
End Function

Public Function RegGetSubKeyList(ByVal RegMainKey As Long, ByVal RegSubKey As String, ByRef RegSubKeyNumber As Integer, ByRef RegSubKeyNameArray() As String) As Boolean
    'on error Resume Next 'returns True for success, False for error; initializes passed array with all sub key names (full sub key name, e.g. 'SOFTWARE\Microsoft\Windows\')
    Dim RegKeyHandle As Long
    Dim RegSubKeyIndex As Integer
    Dim RegSubKeyName As String
    Dim RegSubKeyNameLength As Long
    Dim RegClassName As String
    Dim RegClassNameLength As Long
    Dim Temp1 As Long
    Dim Temp2 As Long
    Dim Tempsngl!
    Dim TempFILETIME As FILETIME
    'reset
    RegSubKeyNumber = 0 'reset
    ReDim RegSubKeyNameArray(1 To 1) As String
    ReDim RegValueValueArray(1 To 1) As String
    'verify
    If Not (Right(RegSubKey, 1) = "\") Then RegSubKey = RegSubKey + "\"
    'preset
    RegSubKeyIndex = 0 'reset; used to number values in registry key
    RegSubKeyNameLength = 1024 'preset
    RegSubKeyName = String$(RegSubKeyNameLength, Chr$(0))
    RegClassNameLength = 1024 'preset
    RegClassName = String$(RegSubKeyNameLength, Chr$(0))
    'begin
    Temp1 = Rmod.RegOpenKeyEx(RegMainKey, RegSubKey, 0, KEY_READ, RegKeyHandle)
    If Not (Temp1 = ERROR_SUCCESS) Then 'error opening key (maybe does not exist)
        GoTo Error:
    End If
    Temp2 = Rmod.RegEnumKeyEx(RegKeyHandle, RegSubKeyIndex, RegSubKeyName, RegSubKeyNameLength, 0, RegClassName, RegClassNameLength, TempFILETIME)
    If Not ((Temp2 = ERROR_SUCCESS) Or (Temp2 = ERROR_NO_MORE_ITEMS)) Then GoTo Error:
    'NOTE: to not change RegSubKeyName and RegClassName, or the program will crash (pointer usage).
    Do While Not ((Temp2 = ERROR_NO_MORE_ITEMS) Or (RegSubKeyNumber = 32766)) 'avoid endless loop through overflow error
        RegSubKeyNumber = RegSubKeyNumber + 1
        ReDim Preserve RegSubKeyNameArray(1 To RegSubKeyNumber) As String
        ReDim Preserve RegValueValueArray(1 To RegSubKeyNumber) As String
        RegSubKeyNameArray(RegSubKeyNumber) = RegSubKey + Left$(RegSubKeyName, RegSubKeyNameLength)
        If Not (Right$(RegSubKeyNameArray(RegSubKeyNumber), 1) = "\") Then 'verify
            RegSubKeyNameArray(RegSubKeyNumber) = RegSubKeyNameArray(RegSubKeyNumber) + "\"
        End If
        RegSubKeyNameLength = 1024 'reset
        RegClassNameLength = 1024 'preset
        RegSubKeyIndex = RegSubKeyIndex + 1
        Temp2 = RegEnumKeyEx(RegKeyHandle, RegSubKeyIndex, RegSubKeyName, RegSubKeyNameLength, 0, RegClassName, RegClassNameLength, TempFILETIME)
        If Not ((Temp2 = ERROR_SUCCESS) Or (Temp2 = ERROR_NO_MORE_ITEMS)) Then GoTo Error:
    Loop
    Call Rmod.RegCloseKey(RegKeyHandle)
    RegGetSubKeyList = True 'ok
    Exit Function
Error:
    RegGetSubKeyList = False 'error
    Exit Function
End Function

Public Function RegGetSubKeyListEx(ByVal RegMainKey As Long, ByVal RegSubKey As String, ByRef RegSubKeyNumber As Integer, ByRef RegSubKeyNameArray() As String, ByVal AddScanStartKeyFlag As Boolean) As Boolean
    'on error resume next 'returns True for success, False for error
    '
    'NOTE: this function calls GFSubKeyList_Create(), a function created out of
    'GFDirectoryList and RegGetSubKeyList.
    'Recursive searching is used to list ALL sub keys of the passed start key.
    'Note that the technique is the same as when scanning directories,
    'but the VB directory list box has been replaces through a piece of code
    'that represents the 'sub key dir box'.
    'Note that creating the sub key list may take a while, even on fast machines.
    '
    RegGetSubKeyListEx = GFSubKeyList_Create(RegMainKey, RegSubKey, AddScanStartKeyFlag, RegSubKeyNameArray(), RegSubKeyNumber)
End Function

Private Function GFSubKeyList_Create(ByVal ScanMainKey As Long, ByVal ScanStartKey As String, ByVal AddScanStartKeyFlag As Boolean, ByRef GFSubKeyListArrayPassed() As String, ByRef GFSubKeyListNumberPassed As Integer) As Boolean
    On Error GoTo Error: 'returns True if seccessful, False if error (i.e. ScanStartKey is invalid)
    Dim Temp As Long
    '
    'NOTE: this function initalizes the passed array with all sub keys of ScanStartKey.
    '
    'reset
    GFSubKeyListNumber = 0 'reset
    ReDim GFSubKeyListArray(1 To 1) As String 'reset
    'scan beginning in ScanStartKey
    If AddScanStartKeyFlag = True Then
        Call GFSubKeyList_AddItem(ScanStartKey)
    End If
    'REGKEYDIR
    GFSubKeyDir_RegMainKey = ScanMainKey
    GFSubKeyDir_RegSubKey = ScanStartKey
    Call RegGetSubKeyList(GFSubKeyDir_RegMainKey, GFSubKeyDir_RegSubKey, GFSubKeyDir_ListCount, GFSubKeyDir_List())
    'END OF REGKEYDIR
    Call GFSubKeyList_Scan("")
    'transfer scan result
    GFSubKeyListNumberPassed = GFSubKeyListNumber
    If Not (GFSubKeyListNumberPassed = 0) Then 'verify
        ReDim GFSubKeyListArrayPassed(1 To GFSubKeyListNumberPassed) As String
    Else
        ReDim GFSubKeyListArrayPassed(1 To 1) As String
    End If
    For Temp = 1 To GFSubKeyListNumberPassed
        GFSubKeyListArrayPassed(Temp) = GFSubKeyListArray(Temp)
    Next Temp
    GFSubKeyList_Create = True 'ok
    Exit Function
Error:
    GFSubKeyList_Create = False 'error
    Exit Function
End Function

Private Sub GFSubKeyList_Scan(ByRef ScanDirOld As String)
    On Error GoTo Error: 'important (avoid endless loop)
    Dim ScanDirNumberTotal As Integer
    Dim ScanParentDir As String
    'begin
    ScanDirNumberTotal = GFSubKeyDir_ListCount
    Do While ScanDirNumberTotal > 0
        ScanParentDir = GFSubKeyDir_RegSubKey
        If GFSubKeyDir_ListCount > 0 Then
            'REGKEYDIR
            GFSubKeyDir_RegMainKey = GFSubKeyDir_RegMainKey
            GFSubKeyDir_RegSubKey = GFSubKeyDir_List(ScanDirNumberTotal)
            Call RegGetSubKeyList(GFSubKeyDir_RegMainKey, GFSubKeyDir_RegSubKey, GFSubKeyDir_ListCount, GFSubKeyDir_List())
            'END OF REGKEYDIR
            Call GFSubKeyList_AddItem(GFSubKeyDir_RegSubKey)
            Call GFSubKeyList_Scan(ScanParentDir)
        End If
        ScanDirNumberTotal = ScanDirNumberTotal - 1
    Loop
    If Not (ScanDirOld = "") Then
        'REGKEYDIR
        GFSubKeyDir_RegMainKey = GFSubKeyDir_RegMainKey
        GFSubKeyDir_RegSubKey = ScanDirOld
        Call RegGetSubKeyList(GFSubKeyDir_RegMainKey, GFSubKeyDir_RegSubKey, GFSubKeyDir_ListCount, GFSubKeyDir_List())
        'END OF REGKEYDIR
    End If
    Exit Sub
Error:
    'do nothing
    Exit Sub
End Sub

Private Sub GFSubKeyList_AddItem(ByVal SubKeyName As String)
    'on error resume next
    If Not (GFSubKeyListNumber = 32766) Then 'verify
        GFSubKeyListNumber = GFSubKeyListNumber + 1
    Else
        Exit Sub 'error
    End If
    If Not (Right$(SubKeyName, 1) = "\") Then SubKeyName = SubKeyName + "\" 'verify
    ReDim Preserve GFSubKeyListArray(1 To GFSubKeyListNumber) As String
    GFSubKeyListArray(GFSubKeyListNumber) = SubKeyName
End Sub

'****************************END OF REGISTRY GET FUNCTIONS*****************************
'********************************REGISTRY SET FUNCTIONS********************************
'NOTE: the following subs/functions are used to set registry data.

Public Sub RegSetKeyValue(ByVal RegMainKey As Long, ByVal RegSubKey As String, ByVal RegValueName As String, ByVal RegValueVariant As Variant, ByVal RegValueType As Long)
    'on error Resume Next 'set a registry key value; if key is not existing, it will be created
    Dim RegKeyHandle As Long
    Dim Temp As Long
    'preset
    RegSetKeyValueCreateSubKeyCalledFlag = False
    'begin
    Temp = RegOpenKeyEx(RegMainKey, RegSubKey, 0&, KEY_WRITE, RegKeyHandle)
    If Temp = ERROR_SUCCESS Then 'verify
        'ok
    Else
        Call RegCreateSubKey(RegMainKey, RegSubKey)
        RegSetKeyValueCreateSubKeyCalledFlag = True
        Temp = RegOpenKeyEx(RegMainKey, RegSubKey, 0&, KEY_WRITE, RegKeyHandle)
    End If
    If Temp = ERROR_SUCCESS Then 'verify
        Select Case RegValueType
        Case REG_SZ
            Temp = RegSetKeyValueString(RegKeyHandle, RegValueName, RegValueType, CStr(RegValueVariant))
        Case REG_DWORD
            Temp = RegSetKeyValueLong(RegKeyHandle, RegValueName, RegValueType, CLng(RegValueVariant))
        Case REG_BINARY
            Temp = RegSetKeyValueByte(RegKeyHandle, RegValueName, RegValueType, CStr(RegValueVariant))
        End Select
    Else
        GoTo Error:
    End If
    If Temp = ERROR_SUCCESS Then 'verify
        'ok
    Else
        GoTo Error:
    End If
    Call RegCloseKey(RegKeyHandle)
    RegSetKeyValueErrorFlag = False
    Exit Sub
Error:
    Call RegCloseKey(RegKeyHandle) 'make sure handle is closed
    RegSetKeyValueErrorFlag = True
    Exit Sub
End Sub

Private Function RegSetKeyValueString(ByVal RegKeyHandle As Long, ByVal RegValueName As String, ByVal RegValueType As Long, ByVal RegValueValue As String) As Long
    'on error Resume Next
    RegSetKeyValueString = RegSetValueEx(RegKeyHandle, RegValueName, 0&, RegValueType, ByVal RegValueValue, Len(RegValueValue))
End Function

Private Function RegSetKeyValueLong(ByVal RegKeyHandle As Long, ByVal RegValueName As String, ByVal RegValueType As Long, ByVal RegValueValue As Long) As Long
    'on error Resume Next
    RegSetKeyValueLong = RegSetValueEx(RegKeyHandle, RegValueName, 0&, RegValueType, RegValueValue, 4)
End Function

Private Function RegSetKeyValueByte(ByVal RegKeyHandle As Long, ByVal RegValueName As String, ByVal RegValueType As Long, ByVal RegValueValue As String) As Long
    'on error Resume Next
    Dim RegValueValueArray() As Byte
    If Not (Len(RegValueValue) = 0) Then 'verify
        ReDim RegValueValueArray(1 To Len(RegValueValue)) As Byte
        Call CopyMemory(RegValueValueArray(1), ByVal RegValueValue, Len(RegValueValue))
        RegSetKeyValueByte = RegSetValueEx(RegKeyHandle, RegValueName, 0&, RegValueType, RegValueValueArray(1), UBound(RegValueValueArray()))
    Else
        ReDim RegValueValueArray(1 To 1) As Byte
        RegSetKeyValueByte = RegSetValueEx(RegKeyHandle, RegValueName, 0&, RegValueType, RegValueValueArray(1), 0)
    End If
End Function

'****************************END OF REGISTRY SET FUNCTIONS*****************************
'******************************REGISTRY CREATE FUNCTIONS*******************************
'NOTE: the following subs/functions are used to create data that can be set.

Public Sub RegCreateSubKey(ByVal RegMainKey As Long, ByVal RegSubKey As String)
    'on error Resume Next 'several sub keys can be created at once, values will not be erased when creating an existing sub key
    Dim RegKeyHandle As Long
    Dim TempSECURITY_ATTRIBUTES As SECURITY_ATTRIBUTES
    Dim Temp1 As Long
    Dim Temp2 As Long
    'begin
    Temp1 = RegCreateKeyEx(RegMainKey, RegSubKey, 0&, "", 0&, KEY_WRITE, TempSECURITY_ATTRIBUTES, Temp2, RegKeyHandle)
    If Temp1 = ERROR_SUCCESS Then 'verify
        'ok
    Else
        GoTo Error:
    End If
    Call RegCloseKey(RegKeyHandle)
    RegCreateSubKeyErrorFlag = False
    Exit Sub
Error:
    Call RegCloseKey(RegKeyHandle) 'make sure handle is closed
    RegCreateSubKeyErrorFlag = True
    Exit Sub
End Sub

'**************************END OF REGISTRY CREATE FUNCTIONS****************************
'******************************REGISTRY DELETE FUNCTIONS*******************************
'NOTE: teh following subs/functions are used to erase registry data.

Public Sub RegDeleteKeyValue(ByVal RegMainKey As Long, ByVal RegSubKey As String, ByVal RegValueName As String)
    'on error Resume Next 'deletes a value name and setting
    Dim RegKeyHandle As Long
    Dim Temp As Long
    'begin
    Temp = RegOpenKeyEx(RegMainKey, RegSubKey, 0&, KEY_WRITE, RegKeyHandle)
    Temp = RegDeleteValue(RegKeyHandle, RegValueName)
    If Temp = ERROR_SUCCESS Then 'verify
        RegDeleteValueErrorFlag = False
    Else
        RegDeleteValueErrorFlag = True
    End If
    Call RegCloseKey(RegKeyHandle)
End Sub

Public Sub RegDeleteSubKey(ByVal RegMainKey As Long, ByVal RegSubKey As String)
    'on error Resume Next 'Win95/98: deletes sub key and all further sub keys, WinNT: cannot delete further sub keys
    Dim Temp As Long
    'begin
    'NOTE: sub key to delete mustn't have further sub keys
    Temp = RegDeleteKey(RegMainKey, RegSubKey)
    If Temp = ERROR_SUCCESS Then 'verify
        RegDeleteKeyErrorFlag = False
    Else
        RegDeleteKeyErrorFlag = True
    End If
End Sub

Public Function RegDeleteSubKeyEx(ByVal RegMainKey As Long, ByVal RegSubKey As String) As Integer
    'on error resume next 'returns the number of deleted sub keys
    Dim RegSubKeyNumber As Integer
    Dim RegSubKeyArray() As String
    Dim RegSubKeyLengthMax As Long
    Dim RegSubKeyDeleteNumber As Integer 'number of deleted sub keys
    Dim Loop1 As Integer
    Dim Loop2 As Integer
    Dim Tempstr$
    '
    'NOTE: this function deletes the passed Registry key and all sub keys.
    '
    'preset
    Call Rmod.RegGetSubKeyListEx(RegMainKey, RegSubKey, RegSubKeyNumber, RegSubKeyArray(), True)
    'begin
    '
    'NOTE: now we sort the sub keys by their length.
    '
    Loop2 = 1 'preset
    Do
        RegSubKeyLengthMax = 0 'reset
        For Loop1 = Loop2 To RegSubKeyNumber
            If Len(RegSubKeyArray(Loop1)) > RegSubKeyLengthMax Then
                RegSubKeyLengthMax = Len(RegSubKeyArray(Loop1))
            End If
        Next Loop1
        If RegSubKeyLengthMax = 0 Then Exit Do
        For Loop1 = Loop2 To RegSubKeyNumber
            If Len(RegSubKeyArray(Loop1)) = RegSubKeyLengthMax Then
                If Not (Loop1 = Loop2) Then 'verify swapping is necessary
                    Tempstr$ = RegSubKeyArray(Loop2)
                    RegSubKeyArray(Loop2) = RegSubKeyArray(Loop1)
                    RegSubKeyArray(Loop1) = Tempstr$
                End If
                Loop2 = Loop2 + 1
            End If
        Next Loop1
    Loop
    '
    'NOTE: now we delete the sub keys.
    '
    For Loop1 = 1 To RegSubKeyNumber
        Rmod.RegDeleteKeyErrorFlag = False 'reset
        Call Rmod.RegDeleteSubKey(RegMainKey, RegSubKeyArray(Loop1))
        If Rmod.RegDeleteKeyErrorFlag = False Then
            RegSubKeyDeleteNumber = RegSubKeyDeleteNumber + 1
        End If
    Next Loop1
    '
    RegDeleteSubKeyEx = RegSubKeyDeleteNumber
End Function

'**************************END OF REGISTRY DELETE FUNCTIONS****************************
'*******************************REGISTRY HELP FUNCTIONS********************************
'NOTE: the following subs/functions are mainly type conversion functions.

Public Function GetRegMainKey(ByVal RegKey As String) As Long
    'on error resume next 'returns value of main key constant or 0 for error
    If UCase$(Left$(RegKey, 17)) = "HKEY_CLASSES_ROOT" Then
        GetRegMainKey = HKEY_CLASSES_ROOT
        Exit Function
    End If
    If UCase$(Left$(RegKey, 17)) = "HKEY_CURRENT_USER" Then
        GetRegMainKey = HKEY_CURRENT_USER
        Exit Function
    End If
    If UCase$(Left$(RegKey, 18)) = "HKEY_LOCAL_MACHINE" Then
        GetRegMainKey = HKEY_LOCAL_MACHINE
        Exit Function
    End If
    If UCase$(Left$(RegKey, 10)) = "HKEY_USERS" Then
        GetRegMainKey = HKEY_USERS
        Exit Function
    End If
    GetRegMainKey = 0 'error (not all keys are supported yet)
    Exit Function
End Function

Public Function GetRegMainKeyName(ByVal RegMainKey As Long) As String
    'on error resume next 'opposite of GetRegMainKey()
    Select Case RegMainKey 'note that not all main keys are supported
    Case HKEY_CLASSES_ROOT
        GetRegMainKeyName = "HKEY_CLASSES_ROOT"
    Case HKEY_CURRENT_USER
        GetRegMainKeyName = "HKEY_CURRENT_USER"
    Case HKEY_LOCAL_MACHINE
        GetRegMainKeyName = "HKEY_LOCAL_MACHINE"
    Case HKEY_USERS
        GetRegMainKeyName = "HKEY_USERS"
    Case Else
        GetRegMainKeyName = "" 'error
    End Select
End Function

Public Function GetRegSubKey(ByVal RegKey As String) As String
    'on error resume next 'returns 'real' sub key name (without main key constant) or "" for error
    Dim Temp As Long
    'verify
    If GetRegMainKey(RegKey) = 0 Then
        GetRegSubKey = "" 'reset (error)
        Exit Function
    End If
    For Temp = 1 To Len(RegKey)
        If Mid$(RegKey, Temp, 1) = "\" Then
            GetRegSubKey = Right$(RegKey, Len(RegKey) - Temp)
            Exit Function 'ok
        End If
    Next Temp
    GetRegSubKey = "" 'error
    Exit Function
End Function

'***************************END OF REGISTRY HELP FUNCTIONS*****************************
'****************************************OTHER*****************************************

Private Function MIN(ByVal Value1 As Long, ByVal Value2 As Long) As Long
    'on error Resume Next 'use for i.e. CopyMemory(a(1), ByVal b, MIN(UBound(a()), Len(b))
    If Value1 < Value2 Then
        MIN = Value1
    Else
        MIN = Value2
    End If
End Function

Private Function MAX(ByVal Value1 As Long, ByVal Value2 As Long) As Long
    'on error Resume Next 'use in combination with ReDim()
    If Value1 > Value2 Then
        MAX = Value1
    Else
        MAX = Value2
    End If
End Function

'***END OF RMOD***

