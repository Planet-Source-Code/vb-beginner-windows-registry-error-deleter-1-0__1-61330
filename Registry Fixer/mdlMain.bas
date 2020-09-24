Attribute VB_Name = "mdlMain"
Option Explicit

'Start file searcher
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Public Const MAX_PATH = 260
Public Const MAXDWORD = &HFFFF
Public Const INVALID_HANDLE_VALUE = -1
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100

Type FILETIME
dwLowDateTime As Long
dwHighDateTime As Long
End Type

Type WIN32_FIND_DATA
dwFileAttributes As Long
ftCreationTime As FILETIME
ftLastAccessTime As FILETIME
ftLastWriteTime As FILETIME
nFileSizeHigh As Long
nFileSizeLow As Long
dwReserved0 As Long
dwReserved1 As Long
cFileName As String * MAX_PATH
cAlternate As String * 14
End Type
'End file searcher


Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Public Const HKEY_ALL = &H0&
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_DYN_DATA = &H80000006
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_USERS = &H80000003

Public Const SYNCHRONIZE = &H100000
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

'Check if a path or file exists
Public Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

'For ListView AutoSize
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const LVM_FIRST = &H1000


Public Function StripNulls(OriginalStr As String) As String
If (InStr(OriginalStr, Chr(0)) > 0) Then
OriginalStr = Left(OriginalStr, _
InStr(OriginalStr, Chr(0)) - 1)
End If
StripNulls = OriginalStr
End Function

'Reverse a string
Public Function ReverseString(TheString As String) As String
    Dim i As Integer
    For i = 1 To Len(TheString)
        ReverseString = ReverseString & Mid(Right$(TheString, i), 1, 1)
    Next
End Function

'Create a backup of the registry, I would have used the "regedit.exe /e" command but it takes too long.
Public Function BackupReg()
    Dim i As Integer
    Dim TheKey As String
    Dim TheValue As String
    Dim DefaultValue As Boolean
    Dim BackupFilename As String
    If FileorFolderExists(App.Path & "\RegBackups") = False Then MkDir App.Path & "\RegBackups"
    Do Until FileorFolderExists(App.Path & "\RegBackups\Backup #" & i & " (" & Replace(Replace(Now, "/", "-"), ":", ";") & ").reg") = False
    i = i + 1
    Loop
    BackupFilename = App.Path & "\RegBackups\Backup #" & i & " (" & Replace(Replace(Now, "/", "-"), ":", ";") & ").reg"
    MsgBox BackupFilename
    Open BackupFilename For Output As #1
    Print #1, "REGEDIT4" & vbCrLf
    'Loops through all the checked items and saves the values into C:\Backup.reg
    For i = 1 To frmMain.lvwRegErrors.ListItems.Count
        If frmMain.lvwRegErrors.ListItems.Item(i).Checked = True Then
            TheKey = ReverseString(frmMain.lvwRegErrors.ListItems.Item(i).SubItems(1) & "\" & frmMain.lvwRegErrors.ListItems.Item(i).SubItems(2))
            'I'm not sure, but I think that is the value ends with a "\", then it's the default value for that key
            If Right$(TheKey, 1) = "\" Then DefaultValue = True: TheKey = Mid(TheKey, 2)
            TheValue = Chr(34) & Replace(ReverseString(Mid(TheKey, 1, InStr(1, TheKey, "\") - 1)), "\", "\\") & Chr(34)
            TheKey = ReverseString(Mid(TheKey, InStr(1, TheKey, "\") + 1))
            If DefaultValue = True Then TheValue = "@"
            Print #1, "[" & TheKey & "]" & vbCrLf
            Print #1, TheValue & "=" & Chr(34) & frmMain.lvwRegErrors.ListItems.Item(i).SubItems(3) & Chr(34) & vbCrLf
        End If
    Next
    Close #1
End Function

Public Sub LV_AutoSizeColumn(LV As ListView, Optional Column As ColumnHeader = Nothing)
 Dim C As ColumnHeader
 If Column Is Nothing Then
  For Each C In LV.ColumnHeaders
   SendMessage LV.hWnd, LVM_FIRST + 30, C.Index - 1, -1
  Next
 Else
  SendMessage LV.hWnd, LVM_FIRST + 30, Column.Index - 1, -1
 End If
 LV.Refresh
End Sub

Public Function FindFilesAPI(Path As String, SearchStr As String, FileCount As Integer, DirCount As Integer)
Dim FileName As String
Dim DirName As String
Dim dirNames() As String
Dim nDir As Integer
Dim i As Integer
Dim hSearch As Long
Dim WFD As WIN32_FIND_DATA
Dim Cont As Integer
If Right(Path, 1) <> "\" Then Path = Path & "\"
nDir = 0
ReDim dirNames(nDir)
Cont = True
hSearch = FindFirstFile(Path & "*", WFD)
If hSearch <> INVALID_HANDLE_VALUE Then
Do While Cont
DirName = StripNulls(WFD.cFileName)
If (DirName <> ".") And (DirName <> "..") Then
If GetFileAttributes(Path & DirName) And FILE_ATTRIBUTE_DIRECTORY Then
dirNames(nDir) = DirName
DirCount = DirCount + 1
nDir = nDir + 1
ReDim Preserve dirNames(nDir)
End If
End If
Cont = FindNextFile(hSearch, WFD)
Loop
Cont = FindClose(hSearch)
End If
hSearch = FindFirstFile(Path & SearchStr, WFD)
Cont = True
If hSearch <> INVALID_HANDLE_VALUE Then
While Cont
FileName = StripNulls(WFD.cFileName)
If (FileName <> ".") And (FileName <> "..") Then
FindFilesAPI = FindFilesAPI + (WFD.nFileSizeHigh * MAXDWORD) + WFD.nFileSizeLow
FileCount = FileCount + 1
frmSearch.lstFiles.AddItem Path & FileName
End If
Cont = FindNextFile(hSearch, WFD)
Wend
Cont = FindClose(hSearch)
End If
If nDir > 0 Then
For i = 0 To nDir - 1
FindFilesAPI = FindFilesAPI + FindFilesAPI(Path & dirNames(i) & "\", SearchStr, FileCount, DirCount)
Next i
End If
End Function

'A DeleteValue function
Public Sub DeleteValue(ROOTKEYS As ROOT_KEYS, Path As String, sKey As String)
    Dim ValKey As String
    Dim SecKey As String, SlashPos As Single
    SlashPos = InStrRev(Path, "\", compare:=vbTextCompare)
    SecKey = Left(Path, SlashPos - 1)    'This will retreive the section key that I need
    ValKey = Right(Path, Len(Path) - SlashPos)    'This will retreive the ValueKey that I need to delete
    DeleteValue2 ROOTKEYS, SecKey, ValKey
End Sub

'Another DeleteValue function
Public Sub DeleteValue2(hKey As ROOT_KEYS, strPath As String, strValue As String)
    Dim Ret
    RegCreateKey hKey, strPath, Ret
    RegDeleteValue Ret, strValue
    RegCloseKey Ret
End Sub

'Returns the long value of the string entered.
Public Function GetClassKey(cls As String) As ROOT_KEYS
    Select Case cls
    Case "HKEY_ALL"
        GetClassKey = HKEY_ALL
    Case "HKEY_CLASSES_ROOT"
        GetClassKey = HKEY_CLASSES_ROOT
    Case "HKEY_CURRENT_USER"
        GetClassKey = HKEY_CURRENT_USER
    Case "HKEY_LOCAL_MACHINE"
        GetClassKey = HKEY_LOCAL_MACHINE
    Case "HKEY_USERS"
        GetClassKey = HKEY_USERS
    Case "HKEY_PERFORMANCE_DATA"
        GetClassKey = HKEY_PERFORMANCE_DATA
    Case "HKEY_CURRENT_CONFIG"
        GetClassKey = HKEY_CURRENT_CONFIG
    Case "HKEY_DYN_DATA"
        GetClassKey = HKEY_DYN_DATA
    End Select
End Function

'Checks if a folder or file exists
Public Function FileorFolderExists(FolderOrFilename As String) As Boolean
    If PathFileExists(FolderOrFilename) = 1 Then
        FileorFolderExists = True
    ElseIf PathFileExists(FolderOrFilename) = 0 Then
        FileorFolderExists = False
    End If
End Function

