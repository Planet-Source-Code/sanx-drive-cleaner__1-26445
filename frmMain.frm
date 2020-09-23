VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Drive Cleaner"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8100
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   8100
   StartUpPosition =   3  'Windows Default
   Begin VB.DirListBox dirList 
      BackColor       =   &H00C0FFFF&
      Height          =   2565
      Left            =   120
      TabIndex        =   15
      Top             =   840
      Width           =   2655
   End
   Begin VB.CheckBox chkTempDir 
      Caption         =   "Temporary Directory"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton butDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   6000
      TabIndex        =   13
      Top             =   6000
      Width           =   1935
   End
   Begin VB.CheckBox chkZERO 
      Caption         =   "Zero-byte Files"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CheckBox chkTILDA 
      Caption         =   "~*.*"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CheckBox chkBAK 
      Caption         =   "*.bak"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CheckBox chkTMP 
      Caption         =   "*.tmp"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CheckBox chkSelectAll 
      Caption         =   "Select All"
      Height          =   255
      Left            =   2880
      TabIndex        =   5
      Top             =   6000
      Width           =   1095
   End
   Begin VB.ListBox lstDir 
      Height          =   255
      Left            =   5520
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   6480
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton butSearch 
      Caption         =   "Search"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   6000
      Width           =   1935
   End
   Begin VB.ListBox lstFound 
      BackColor       =   &H00C0FFFF&
      Height          =   5460
      Left            =   2880
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   480
      Width           =   5055
   End
   Begin MSComctlLib.StatusBar barMain 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   6480
      Width           =   8100
      _ExtentX        =   14288
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.DriveListBox drvList 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2655
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   2760
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Label Label3 
      Caption         =   "Files to look for:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   2760
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label Label2 
      Caption         =   "Files found:"
      Height          =   255
      Left            =   2880
      TabIndex        =   7
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Drive:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This application is copyright Sanx, 2001. This application is
'offered as freeware, and as such, you may copy, modify and
'distribute it without conditions, provided this copyright
'notice remains.
'http://www.sanx.org

Private Sub butDelete_Click()

On Error Resume Next

response = MsgBox("Are you sure you wish to delete all selected files?", vbYesNo + vbCritical, "Drive Cleaner")

If response <> 6 Then Exit Sub

For temp = 0 To (lstFound.ListCount - 1)
    If lstFound.Selected(temp) = True Then
        delFile = lstFound.List(temp)
        barMain.SimpleText = "Deleting: " + delFile
        Kill delFile
    End If
Next

barMain.SimpleText = "Finished deleting."

chkSelectAll.Value = 0

End Sub

Private Sub butSearch_Click()

lstDir.Clear
lstFound.Clear
lstFound.Refresh
ScanDir
FileFind

If chkTempDir.Value = 1 Then
    FindTempDir
End If

End Sub

Private Function GetTempPath()

GetTempPath = Environ("TEMP")

End Function

Private Sub FindTempDir()

barMain.SimpleText = "Scanning temporary directories..."

lstDir.Clear

MyPath = GetTempPath() + "\"
lstDir.AddItem GetTempPath()
MyName = Dir(MyPath, vbDirectory)
Do While MyName <> ""
    If MyName <> "." And MyName <> ".." Then
        If (GetAttr(MyPath & MyName) And vbDirectory) = vbDirectory Then
            lstDir.AddItem MyPath & MyName
        End If
    End If
    MyName = Dir
Loop

entryCount = 0
Do While entryCount < lstDir.ListCount

MyPath = lstDir.List(entryCount) + "\"
MyName = Dir(MyPath, vbDirectory)
Do While MyName <> ""
    If MyName <> "." And MyName <> ".." Then
        If (GetAttr(MyPath & MyName) And vbDirectory) = vbDirectory Then
            lstDir.AddItem MyPath & MyName
        End If
    End If
    MyName = Dir
Loop
entryCount = entryCount + 1
Loop

entryCount = 0
Do While entryCount < lstDir.ListCount
barMain.SimpleText = "Scanning " + lstDir.List(entryCount) + " ..."
MyPath = lstDir.List(entryCount) + "\*.*"
MyName = Dir(MyPath, vbNormal)
Do While MyName <> ""
    If MyName <> "." And MyName <> ".." Then
        If (GetAttr(lstDir.List(entryCount) & "\" & MyName) And vbNormal) = vbNormal Then
            lstFound.AddItem lstDir.List(entryCount) & "\" & MyName
        End If
    End If
    MyName = Dir
Loop
entryCount = entryCount + 1
Loop

barMain.SimpleText = "Finished temporary directory scan."

End Sub

Private Sub ScanDir()

On Error GoTo ErrorFound

barMain.SimpleText = "Building directory tree..."

MyPath = dirList.Path + "\"
MyName = Dir(MyPath, vbDirectory)
Do While MyName <> ""
    If MyName <> "." And MyName <> ".." Then
        If (GetAttr(MyPath & MyName) And vbDirectory) = vbDirectory Then
            lstDir.AddItem MyPath & MyName
        End If
    End If
    MyName = Dir
Loop

entryCount = 0
Do While entryCount < lstDir.ListCount

MyPath = lstDir.List(entryCount) + "\"
MyName = Dir(MyPath, vbDirectory)
Do While MyName <> ""
    If MyName <> "." And MyName <> ".." Then
        If (GetAttr(MyPath & MyName) And vbDirectory) = vbDirectory Then
            lstDir.AddItem MyPath & MyName
        End If
    End If
    MyName = Dir
Loop
entryCount = entryCount + 1
Loop

barMain.SimpleText = "Finished directory scan."

Exit Sub

ErrorFound:

MsgBox "Error: " + Err.Number + vbCrLf + Err.Description + vbCrLf + "Application will now close.", vbCritical + vbApplicationModal + vbOKOnly, "Fatal Error Encountered"
End

End Sub

Private Sub FileFind()

If chkTMP.Value = 1 Then
    FindTMP
End If

If chkBAK.Value = 1 Then
    FindBAK
End If

If chkTILDA.Value = 1 Then
    FindTILDA
End If

If chkZERO.Value = 1 Then
    FindZero
End If

barMain.SimpleText = "Finished file scan."

End Sub


Private Sub FindZero()

entryCount = 0
Do While entryCount < lstDir.ListCount
barMain.SimpleText = "Scanning " + lstDir.List(entryCount) + " ..."
MyPath = lstDir.List(entryCount) + "\*.*"
MyName = Dir(MyPath, vbNormal)
Do While MyName <> ""
    If MyName <> "." And MyName <> ".." Then
        If FileLen(lstDir.List(entryCount) & "\" & MyName) = 0 Then
            lstFound.AddItem lstDir.List(entryCount) & "\" & MyName
        End If
    End If
    MyName = Dir
Loop
entryCount = entryCount + 1
Loop

End Sub
Private Sub FindTILDA()

entryCount = 0
Do While entryCount < lstDir.ListCount
barMain.SimpleText = "Scanning " + lstDir.List(entryCount) + " ..."
MyPath = lstDir.List(entryCount) + "\~*.*"
MyName = Dir(MyPath, vbNormal)
Do While MyName <> ""
    If MyName <> "." And MyName <> ".." Then
        If (GetAttr(lstDir.List(entryCount) & "\" & MyName) And vbNormal) = vbNormal Then
            lstFound.AddItem lstDir.List(entryCount) & "\" & MyName
        End If
    End If
    MyName = Dir
Loop
entryCount = entryCount + 1
Loop

End Sub
Private Sub FindBAK()

entryCount = 0
Do While entryCount < lstDir.ListCount
barMain.SimpleText = "Scanning " + lstDir.List(entryCount) + " ..."
MyPath = lstDir.List(entryCount) + "\*.bak"
MyName = Dir(MyPath, vbNormal)
Do While MyName <> ""
    If MyName <> "." And MyName <> ".." Then
        If (GetAttr(lstDir.List(entryCount) & "\" & MyName) And vbNormal) = vbNormal Then
            lstFound.AddItem lstDir.List(entryCount) & "\" & MyName
        End If
    End If
    MyName = Dir
Loop
entryCount = entryCount + 1
Loop

End Sub
Private Sub FindTMP()

entryCount = 0
Do While entryCount < lstDir.ListCount
barMain.SimpleText = "Scanning " + lstDir.List(entryCount) + " ..."
MyPath = lstDir.List(entryCount) + "\*.tmp"
MyName = Dir(MyPath, vbNormal)
Do While MyName <> ""
    If MyName <> "." And MyName <> ".." Then
        If (GetAttr(lstDir.List(entryCount) & "\" & MyName) And vbNormal) = vbNormal Then
            lstFound.AddItem lstDir.List(entryCount) & "\" & MyName
        End If
    End If
    MyName = Dir
Loop
entryCount = entryCount + 1
Loop

End Sub

Private Sub chkSelectAll_Click()

If chkSelectAll.Value = 1 Then
    For temp = 0 To lstFound.ListCount - 1
        lstFound.Selected(temp) = True
    Next
Else
    For temp = 0 To lstFound.ListCount - 1
        lstFound.Selected(temp) = False
    Next
End If

End Sub

Private Sub drvList_Change()

dirList.Path = Left$(drvList.Drive, 2) + "\"

End Sub

Private Sub Form_Load()

dirList.Path = Left$(drvList.Drive, 2) + "\"

End Sub
