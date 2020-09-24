VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "Windows Registry Error Deleter 1.0"
   ClientHeight    =   4035
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   8985
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4035
   ScaleWidth      =   8985
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvwRegErrors 
      Height          =   2175
      Left            =   0
      TabIndex        =   3
      Top             =   1800
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   3836
      View            =   2
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdStartStop 
      Caption         =   "&Start Scan"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   8955
   End
   Begin VB.Label lblCurrentKey 
      BorderStyle     =   1  'Fixed Single
      Height          =   795
      Left            =   1680
      TabIndex        =   2
      Top             =   960
      Width           =   7155
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblStatus 
      Caption         =   "Searching Key:"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1515
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuStartStop 
         Caption         =   "&Start Scan"
      End
      Begin VB.Menu mnuSeperator0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "&Restore Registry Backup"
      End
      Begin VB.Menu mnuSeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuRepair 
      Caption         =   "------------------>&Repair<------------------"
      Begin VB.Menu mnuCheckAll 
         Caption         =   "&Check All Items"
      End
      Begin VB.Menu mnuUncheckAll 
         Caption         =   "&Uncheck All Items"
      End
      Begin VB.Menu mnuSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFix 
         Caption         =   "&Delete All Checked Items"
      End
      Begin VB.Menu mnuSeperator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "&Search for Missing File, Manually (Experts Only)"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp2 
         Caption         =   "&Help"
      End
      Begin VB.Menu mnuSeperator4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents cReg As cRegSearch
Attribute cReg.VB_VarHelpID = -1

'Stop or start scanning for errors
Private Sub cmdStartStop_Click()
    If cmdStartStop.Caption = "&Start Scan" Then mnuStartStop.Caption = "&Stop Scan"
    If cmdStartStop.Caption = "&Stop Scan" Then
        Caption = "Exiting..."
        cReg.StopSearch
        Exit Sub
    End If
    If lvwRegErrors.ListItems.Count > 0 Then mnuRepair.Visible = True
    cmdStartStop.Caption = "&Stop Scan"
    If lvwRegErrors.Visible = False Then
        Top = Top / 2
        Height = Height * 2
        lvwRegErrors.Visible = True
    End If
    lvwRegErrors.ListItems.Clear
    lblStatus = "Searching key:"
    lblCurrentKey = ""

    cReg.RootKey = 0
    cReg.SubKey = ""
    cReg.SearchFlags = KEY_NAME * 0 + VALUE_NAME * 1 + VALUE_VALUE * 1 + WHOLE_STRING * 0
    cReg.SearchString = "C:\"
    Caption = "Scanning..."
    cReg.DoSearch
    If lvwRegErrors.ListItems.Count = 0 Then mnuRepair.Visible = False
End Sub

'The search is finished
Private Sub cReg_SearchFinished(ByVal lReason As Long)
    If lReason = 0 Then
        lblCurrentKey = "Done!"
    ElseIf lReason = 1 Then
        lblCurrentKey = "Terminated by user!"
    Else
        lblCurrentKey = "An Error occured! Err number = " & lReason
        'Err.Raise lReason
    End If
    cmdStartStop.Caption = "&Start Scan"
    mnuRepair.Visible = True
    lblStatus = "Search result:"
    Caption = "Finished Scanning (" & lvwRegErrors.ListItems.Count & " errors found)"
End Sub

'If a registry error is found
Private Sub cReg_SearchFound(ByVal sRootKey As String, ByVal sKey As String, ByVal sValue As Variant, ByVal lFound As FOUND_WHERE)
    Dim sTemp As String
    Dim FileorPath As String
    Dim lvItm As ListItem
    Select Case lFound
    Case FOUND_IN_KEY_NAME
        sTemp = "KEY_NAME"
    Case FOUND_IN_VALUE_NAME
        sTemp = "VALUE NAME"
    Case FOUND_IN_VALUE_VALUE
        sTemp = "VALUE VALUE"
    End Select

    'Fix up the file or path so that it's compatible with the FileorFolderExists function
    FileorPath = sValue

    'Find the start of the path or filename (Example:"h6j65ej(C:\Test)")
    If InStr(1, FileorPath, "C:\") Then FileorPath = Mid(FileorPath, InStr(1, FileorPath, "C:\"))
    If InStr(1, FileorPath, "c:\") Then FileorPath = Mid(FileorPath, InStr(1, FileorPath, "c:\"))

    'Remove everything after the path. This definitely doesn't work for all values.
    '(Example:"C:\blablablablablablablabla?5784846\84585")
    If InStr(1, FileorPath, "/") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, "/") - 1)
    If InStr(1, FileorPath, "*") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, "*") - 1)
    If InStr(1, FileorPath, "?") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, "?") - 1)
    If InStr(1, FileorPath, Chr(34)) > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, Chr(34)) - 1)
    If InStr(1, FileorPath, "<") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, "<") - 1)
    If InStr(1, FileorPath, ">") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ">") - 1)
    If InStr(1, FileorPath, "|") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, "|") - 1)
    If InStr(1, FileorPath, ",") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ",") - 1)
    If InStr(1, FileorPath, "(") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, "(") - 1)
    If InStr(1, FileorPath, ";") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ";") - 1)

    'I don 't know if it's just my computer, but some registry values somehow didn't contain "C:\"
    If InStr(1, FileorPath, "C:\") = 0 Then FileorPath = "C:\"

    'Remove everything before the path or file. The same as the other one except this is for specific extensions
    '(Example:"C:\lalalalalalalala\idfjb.dll\50")
    If InStr(1, FileorPath, ".EXE ") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".EXE ") + 3)
    If InStr(1, FileorPath, ".exe ") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".exe ") + 3)
    If InStr(1, FileorPath, ".SYS ") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".SYS ") + 3)
    If InStr(1, FileorPath, ".sys ") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".sys ") + 3)
    If InStr(1, FileorPath, ".EXE\") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".EXE\") + 3)
    If InStr(1, FileorPath, ".exe\") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".exe\") + 3)
    If InStr(1, FileorPath, ".DLL\") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".DLL\") + 3)
    If InStr(1, FileorPath, ".dll\") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".dll\") + 3)
    If InStr(1, FileorPath, ".OCX\") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".OCX\") + 3)
    If InStr(1, FileorPath, ".ocx\") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".ocx\") + 3)
    If InStr(1, FileorPath, "*") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, "*") - 1)

    '%1 is used for file associations
    '(Example:"C:\WINDOWS\NOTEPAD.EXE %1")
    FileorPath = Replace(FileorPath, " %1", "")

    'Check if the value is an invalid path or file
    'If it is, then it adds the value to lvwRegErrors and displays the current number of errors, so far.
    If FileorFolderExists(FileorPath) = False Then
        With lvwRegErrors
            Set lvItm = .ListItems.Add(, , sTemp)
            lvItm.SubItems(1) = sRootKey
            lvItm.SubItems(2) = sKey
            lvItm.SubItems(3) = sValue
        End With
        LV_AutoSizeColumn lvwRegErrors
        Me.Caption = "Windows Registry Error Deleter 1.0 (" & lvwRegErrors.ListItems.Count & " errors found)"
        lblStatus.Caption = "Searching Key:" & vbCrLf & "(" & lvwRegErrors.ListItems.Count & " errors found)"
    End If
    
    Set lvItm = Nothing
End Sub

'I don't know if I should remove it
Private Sub cReg_SearchKeyChanged(ByVal sFullKeyName As String)
'Note: This event cause a lot of printing.
'To increase performance remove this event.
    If Me.WindowState <> vbMinimized Then lblCurrentKey = sFullKeyName
End Sub

'Setup everything
Private Sub Form_Load()
    mnuRepair.Visible = False
    With lvwRegErrors
        .View = lvwReport
        .ColumnHeaders.Add , , "Found at:"
        .ColumnHeaders.Add , , "RootKey"
        .ColumnHeaders.Add , , "SubKey"
        .ColumnHeaders.Add , , "Value"
    End With

    Me.Height = (Me.Height - Me.ScaleHeight) + lvwRegErrors.Top
    
    Me.Move Me.Left, Me.Top, Screen.Width / 1.5, Screen.Height / 1.5
    Me.Move (Screen.Width / 2) - (Me.ScaleWidth / 2), (Screen.Height / 2) - (Me.ScaleHeight / 2)
    Set cReg = New cRegSearch
End Sub

'Resize the controls if the form is resized
Private Sub Form_Resize()
    On Error GoTo ERROR_HANDLER
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.WindowState = vbMaximized Then lvwRegErrors.Visible = True
    cmdStartStop.Move 0, cmdStartStop.Top, Me.ScaleWidth
    cmdStartStop.Left = Me.ScaleWidth - cmdStartStop.Width
    lblCurrentKey.Width = cmdStartStop.Left + cmdStartStop.Width - lblCurrentKey.Left
    lvwRegErrors.Move 0, lblCurrentKey.Top + lblCurrentKey.Height, Me.ScaleWidth, Me.ScaleHeight - 1800
    lvwRegErrors.ColumnHeaders(3).Width = (lvwRegErrors.Width - lvwRegErrors.ColumnHeaders(1).Width * 2) / 2 - 600
    lvwRegErrors.ColumnHeaders(4).Width = lvwRegErrors.ColumnHeaders(3).Width
    LV_AutoSizeColumn lvwRegErrors
    Exit Sub
ERROR_HANDLER:
End Sub

'I'm not sure if this is necessary, but I guess it's just to clean up and exit this program
Private Sub Form_Unload(Cancel As Integer)
    cReg.StopSearch
    Set cReg = Nothing
End Sub

'If you select multiple items, they will be checked if their unchecked and unchecked if their checked
Private Sub lvwRegErrors_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked = False Then
        Item.Checked = True
    Else
        Item.Checked = False
    End If
End Sub

'If the the right clicks on lvwRegErrors, mnuRepair become visible
Private Sub lvwRegErrors_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And mnuRepair.Visible = True Then
        PopupMenu mnuRepair
    End If
End Sub

'About menu
Private Sub mnuAbout_Click()
    MsgBox "VB Registry Fixer 1.0" & vbCrLf & vbCrLf & "If you use this program, you take full responsibility for any damages this program may do to your computer.", vbInformation, "VB Registry Fixer 1.0"
End Sub

Private Sub mnuExit_Click()
End
End Sub

'Creates a registry backup of all the values about to be deleted and then deletes them
Private Sub mnuFix_Click()
    Dim i As Integer, nLoop As Single, m As Single
    Dim NOTRemoved As Integer
    Dim BackupFilename As String
    On Error Resume Next

'I don't think this is necessary, but if the registry backup takes a while, this program tells the user to wait.
    lblCurrentKey.FontSize = 24
    lblCurrentKey.FontBold = True
    lblCurrentKey.Caption = "Creating Registry Backup..."
    BackupReg
    lblCurrentKey.FontSize = 8
    lblCurrentKey.FontBold = False
    lblCurrentKey.Caption = ""

    Do Until FileorFolderExists(App.Path & "\RegBackup\Backups #" & i & " (" & Replace(Replace(Now, "/", "-"), ":", ";") & ").reg") = False
    BackupFilename = App.Path & "\RegBackup\Backups #" & i & " (" & Replace(Replace(Now, "/", "-"), ":", ";") & ").reg"
    i = i + 1
    Loop
    
    'Tell the user that this program has created a backup and and to restore the registry if the user's computer acts abnormal
    MsgBox "This program has created a backup of all of the registry values that are about to be deleted. If you experience problems after using this, keep pressing F8 when you start up your computer and select Safe Mode and open up " & BackupFilename, vbInformation, "Important"
    
    'Loop through every item in lvwRegErrors
    For i = 1 To lvwRegErrors.ListItems.Count
        'If the item is checked
        If lvwRegErrors.ListItems.Item(i).Checked = True Then
            'Delete the registry error and mark the item as removed
            DeleteValue GetClassKey(lvwRegErrors.ListItems.Item(i).SubItems(1)), lvwRegErrors.ListItems.Item(i).SubItems(2), lvwRegErrors.ListItems.Item(i).SubItems(3)
            lvwRegErrors.ListItems.Item(i).Text = "REMOVED"
            NOTRemoved = NOTRemoved + 1
        End If    'If you remove the if...then line above then also remove this line.
    Next
    
    'Tell the user how many items that were not removed

    MsgBox "VB Registry Fixer has successfully fixed your registry. There were " & lvwRegErrors.ListItems.Count - NOTRemoved & " registry values that were NOT removed."
End Sub

'Help menu
Private Sub mnuHelp2_Click()
    MsgBox "Step 1 - Click Start Scan" & vbCrLf & vbCrLf & _
    "Step 2 - When the scan is finished, check all the items on the list that you want to delete. I highly recommend that you look carefully for what items you want to remove and not just check all of them." & vbCrLf & vbCrLf & _
    "Step 3 - Right click the list and click 'Delete All Checked Items'", vbInformation, "Help"
End Sub

'Checked all items in lvwRegErrors
Private Sub mnuCheckAll_Click()
    Dim i As Integer
    For i = 1 To lvwRegErrors.ListItems.Count
        lvwRegErrors.ListItems.Item(i).Checked = True
    Next
End Sub

Private Sub mnuRestore_Click()
frmRestore.Show vbModal
End Sub

Private Sub mnuSearch_Click()
frmSearch.Show vbModal
End Sub

'Uncheck all checked items in lvwRegErrors
Private Sub mnuUncheckAll_Click()
    Dim i As Integer
    For i = 1 To lvwRegErrors.ListItems.Count
        lvwRegErrors.ListItems.Item(i).Checked = False
    Next
End Sub

'Start or stop the scan
Private Sub mnuStartStop_Click()
    If mnuStartStop.Caption = "&Start Scan" Then
        cmdStartStop_Click
        mnuStartStop.Caption = "&Stop Scan"
    End If
    
    If mnuStartStop.Caption = "&Stop Scan" Then
        cmdStartStop_Click
        mnuStartStop.Caption = "&Start Scan"
    End If
End Sub
