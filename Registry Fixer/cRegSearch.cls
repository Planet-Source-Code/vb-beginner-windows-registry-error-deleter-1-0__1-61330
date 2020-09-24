VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRegSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Sorry, this isn't my registry class so I can't really make any comments, except for obvious ones


' Class for searching Windows Registry
'
' Written by Arkadiy Olovyannikov (ark@fesma.ru)
' Copyright 2001 by Arkadiy Olovyannikov
'
' This software is FREEWARE. You may use it as you see fit for
' your own projects but you may not re-sell the original or the
' source code.
'
' No warranty express or implied, is given as to the use of this
' program. Use at your own risk.

Option Explicit

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long

Enum ROOT_KEYS
    HKEY_ALL = &H0&
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_CURRENT_USER = &H80000001
    HKEY_DYN_DATA = &H80000006
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_USERS = &H80000003
End Enum

Enum SEARCH_FLAGS
    KEY_NAME = 0
    VALUE_NAME = 1
    VALUE_VALUE = 2
    WHOLE_STRING = 4
End Enum

Enum FOUND_WHERE
    FOUND_IN_KEY_NAME
    FOUND_IN_VALUE_NAME
    FOUND_IN_VALUE_VALUE
End Enum

Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const SYNCHRONIZE = &H100000
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Const KEY_READ = &H20019    ' ((READ_CONTROL Or KEY_QUERY_VALUE Or
' KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not
' SYNCHRONIZE))

Private Const ERROR_SUCCESS = 0&
Private Const ERR_MORE_DATA = 234&
Private Const ERROR_NO_MORE_ITEMS = 259&

Private Const REG_NONE = 0
Private Const REG_SZ = 1
Private Const REG_EXPAND_SZ = 2
Private Const REG_BINARY = 3
Private Const REG_DWORD = 4
Private Const REG_DWORD_LITTLE_ENDIAN = 4
Private Const REG_DWORD_BIG_ENDIAN = 5
Private Const REG_LINK = 6
Private Const REG_MULTI_SZ = 7
Private Const REG_RESOURCE_LIST = 8
Private Const REG_FULL_RESOURCE_DESCRIPTOR = 9
Private Const REG_RESOURCE_REQUIREMENTS_LIST = 10

Private Const MAX_KEY_SIZE = 260
Private Const MAX_VALUE_SIZE = 4096

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)

Public Event SearchFound(ByVal sRootKey As String, ByVal sKey As String, ByVal sValue As Variant, ByVal lFound As FOUND_WHERE)
Public Event SearchFinished(ByVal lReason As Long)
Public Event SearchKeyChanged(ByVal sFullKeyName As String)

Private mvarRootKey As ROOT_KEYS
Private mvarSearchFlags As SEARCH_FLAGS
Private mvarSearchString As String
Private mvarSubKey As String

Dim lStopSearch As Long

Public Property Let SubKey(ByVal vData As String)
    mvarSubKey = vData
End Property

Public Property Let SearchString(ByVal vData As String)
    mvarSearchString = vData
End Property

Public Property Let SearchFlags(ByVal vData As SEARCH_FLAGS)
    mvarSearchFlags = vData
End Property

Public Property Let RootKey(ByVal vData As ROOT_KEYS)
    mvarRootKey = vData
End Property

Public Sub DoSearch()
    If mvarRootKey <> HKEY_ALL Then
        If (mvarSearchFlags And VALUE_NAME) = VALUE_NAME Or (mvarSearchFlags And VALUE_VALUE) = VALUE_VALUE Then
            Call EnumRegValues(mvarRootKey, mvarSubKey)
        End If
        Call EnumRegKeys(mvarRootKey, mvarSubKey)
    Else
        Call EnumRegKeys(HKEY_CLASSES_ROOT, mvarSubKey)
        If lStopSearch Then GoTo Search_Terminated
        Call EnumRegKeys(HKEY_CURRENT_USER, mvarSubKey)
        If lStopSearch Then GoTo Search_Terminated
        Call EnumRegKeys(HKEY_LOCAL_MACHINE, mvarSubKey)
        If lStopSearch Then GoTo Search_Terminated
        Call EnumRegKeys(HKEY_USERS, mvarSubKey)
        If lStopSearch Then GoTo Search_Terminated
        Call EnumRegKeys(HKEY_PERFORMANCE_DATA, mvarSubKey)
        If lStopSearch Then GoTo Search_Terminated
        Call EnumRegKeys(HKEY_CURRENT_CONFIG, mvarSubKey)
        If lStopSearch Then GoTo Search_Terminated
        Call EnumRegKeys(HKEY_DYN_DATA, mvarSubKey)
    End If
Search_Terminated:
    RaiseEvent SearchFinished(lStopSearch)
    lStopSearch = 0
End Sub

Public Sub StopSearch()
    lStopSearch = 1
End Sub

Private Sub EnumRegKeys(ByVal lKeyRoot As Long, ByVal sSubKey As String)
    Dim curidx As Long
    Dim KeyName As String
    Dim hKey As Long
    Dim sTemp As String
    If lStopSearch Then Exit Sub
    On Error GoTo ErrEnum
    If RegOpenKeyEx(lKeyRoot, sSubKey, 0, KEY_READ, hKey) Then Exit Sub
    Do
        DoEvents
        KeyName = Space$(MAX_KEY_SIZE)
        If RegEnumKey(hKey, curidx, KeyName, MAX_KEY_SIZE) <> ERROR_SUCCESS Then Exit Do
        curidx = curidx + 1
        KeyName = TrimNull(KeyName)
        If sSubKey <> "" Then
            sTemp = sSubKey & "\" & KeyName
        Else
            sTemp = KeyName
        End If
        '****************************************************
        'This event is used for showing currently viewing key.
        'Usually you don't need this.
        'To increase performance, remove this event
        If lStopSearch = 0 Then RaiseEvent SearchKeyChanged(RootKeyName(lKeyRoot) & "\" & sTemp)
        '****************************************************
        If (mvarSearchFlags And KEY_NAME) = KEY_NAME Then
            If CheckMatching(KeyName) Then
                RaiseEvent SearchFound(RootKeyName(lKeyRoot), sTemp, "*", FOUND_IN_KEY_NAME)
            End If
        End If
        If (mvarSearchFlags And VALUE_NAME) = VALUE_NAME Or (mvarSearchFlags And VALUE_VALUE) = VALUE_VALUE Then
            Call EnumRegValues(lKeyRoot, sTemp)
        End If
        Call EnumRegKeys(lKeyRoot, sTemp)
    Loop
ErrEnum:
    If Err Then lStopSearch = Err
    RegCloseKey hKey
End Sub

Private Sub EnumRegValues(ByVal lKeyRoot As Long, ByVal sSubKey As String)
    Dim curidx As Long, ValueName As String, ValueValue As String
    Dim hKey As Long
    Dim lType As Long
    Dim arrData() As Byte
    Dim cbDataSize As Long
    If lStopSearch Then Exit Sub
    On Error GoTo ErrEnum
    If RegOpenKeyEx(lKeyRoot, sSubKey, 0, KEY_READ, hKey) Then Exit Sub
    Do
        ValueName = String(MAX_KEY_SIZE, 0)
        cbDataSize = MAX_VALUE_SIZE
        ReDim arrData(cbDataSize - 1)
        If RegEnumValue(hKey, curidx, ValueName, MAX_KEY_SIZE, ByVal 0&, lType, arrData(0), cbDataSize) <> ERROR_SUCCESS Then Exit Do
        If cbDataSize < 1 Then cbDataSize = 1
        ReDim Preserve arrData(cbDataSize - 1)
        ValueName = TrimNull(ValueName)
        If (mvarSearchFlags And VALUE_NAME) = VALUE_NAME Then
            If CheckMatching(ValueName) Then RaiseEvent SearchFound(RootKeyName(lKeyRoot), sSubKey & "\" & ValueName, GetRegData(lType, arrData), FOUND_IN_VALUE_NAME)
        End If
        If (mvarSearchFlags And VALUE_VALUE) = VALUE_VALUE Then
            ValueValue = TrimNull(GetRegData(lType, arrData))
            If CheckMatching(ValueValue) Then
                RaiseEvent SearchFound(RootKeyName(lKeyRoot), sSubKey & "\" & ValueName, ValueValue, FOUND_IN_VALUE_VALUE)
            End If
        End If
        curidx = curidx + 1
    Loop
ErrEnum:
    If Err Then lStopSearch = Err
    RegCloseKey hKey
End Sub

Private Function TrimNull(startstr As String) As String
    Dim pos As Integer
    pos = InStr(startstr, Chr$(0))
    If pos Then
        TrimNull = Left$(startstr, pos - 1)
        Exit Function
    End If
    TrimNull = startstr
End Function

Private Function CheckMatching(ByVal sCheck As String) As Boolean
    If (mvarSearchFlags And WHOLE_STRING) = WHOLE_STRING Then
        CheckMatching = (UCase(sCheck) = UCase(mvarSearchString))
    Else
        CheckMatching = InStr(1, sCheck, mvarSearchString, vbTextCompare)
    End If
End Function

Private Function GetRegData(ByVal lType As Long, abData() As Byte) As String
    Dim lData As Long, i As Long
    Dim sTemp As String
    sTemp = ""
    Select Case lType
    Case REG_SZ, REG_MULTI_SZ
        GetRegData = TrimNull(StrConv(abData, vbUnicode))
    Case REG_DWORD
        CopyMem lData, abData(0), 4&
        GetRegData = "0x" & Format(Hex(lData), "00000000") & "(" & lData & ")"
    Case REG_BINARY
        For i = 0 To UBound(abData)
            sTemp = sTemp & Right("00" & Hex(abData(i)), 2) & " "
        Next i
        GetRegData = Left(sTemp, Len(sTemp) - 1)
    Case Else
        GetRegData = "Temporarily unsupported"
    End Select
End Function

Private Function RootKeyName(lKey As Long) As String
    Select Case lKey
    Case HKEY_CLASSES_ROOT
        RootKeyName = "HKEY_CLASSES_ROOT"
    Case HKEY_CURRENT_USER
        RootKeyName = "HKEY_CURRENT_USER"
    Case HKEY_LOCAL_MACHINE
        RootKeyName = "HKEY_LOCAL_MACHINE"
    Case HKEY_USERS
        RootKeyName = "HKEY_USERS"
    Case HKEY_PERFORMANCE_DATA
        RootKeyName = "HKEY_PERFORMANCE_DATA"
    Case HKEY_CURRENT_CONFIG
        RootKeyName = "HKEY_CURRENT_CONFIG"
    Case HKEY_DYN_DATA
        RootKeyName = "HKEY_DYN_DATA"
    End Select
End Function

Private Sub Class_Initialize()
    mvarRootKey = HKEY_ALL
    mvarSubKey = ""
    mvarSearchString = ""
End Sub

Private Sub Class_Terminate()
    lStopSearch = 1
End Sub
