VERSION 5.00
Begin VB.Form frmSearch 
   Caption         =   "Search for Missing File"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   6945
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   240
      Width           =   6495
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.ListBox lstFiles 
      Height          =   2010
      Left            =   0
      TabIndex        =   0
      Top             =   2160
      Width           =   6855
   End
   Begin VB.Label lblFile 
      Caption         =   "File to Search For:"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSearch_Click()
Dim SearchPath As String, FindStr As String
Dim FileSize As Long
Dim NumFiles As Integer, NumDirs As Integer
Screen.MousePointer = vbHourglass
lstFiles.Clear
SearchPath = "C:\"
FindStr = txtFile.Text
FileSize = FindFilesAPI(SearchPath, FindStr, NumFiles, NumDirs)
MsgBox NumFiles & " Files found in " & NumDirs + 1 & " Directories"
MsgBox "Size of files found under " & SearchPath & " = " & Format(FileSize, "#,###,###,##0") & " Bytes"
Screen.MousePointer = vbDefault
End Sub

