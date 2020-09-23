VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Image Dupe Remover"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7830
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   7830
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Duplicate files found"
      Height          =   1575
      Left            =   2040
      TabIndex        =   6
      Top             =   4440
      Width           =   5655
      Begin VB.OptionButton Option2 
         Caption         =   "Erase any duplicates found."
         Height          =   255
         Left            =   3120
         TabIndex        =   9
         Top             =   1230
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Move duplicates into new folder"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1230
         Value           =   -1  'True
         Width           =   2775
      End
      Begin VB.ListBox lstDupes 
         Height          =   840
         Left            =   120
         TabIndex        =   7
         Top             =   270
         Width           =   5415
      End
   End
   Begin VB.CheckBox chkPreview 
      Caption         =   "Show Preview"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   4420
      Value           =   1  'Checked
      Width           =   1320
   End
   Begin MSComctlLib.ListView lvMain 
      Height          =   2775
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   4895
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Filename"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "CRC"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "File Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Date Time"
         Object.Width           =   3599
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Directory"
         Object.Width           =   4305
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Directory To Scan"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin VB.CheckBox Check1 
         Caption         =   "Scan Sub Folders"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse"
         Height          =   495
         Left            =   5040
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdStartDir 
         Caption         =   "Start"
         Enabled         =   0   'False
         Height          =   495
         Left            =   6240
         MousePointer    =   1  'Arrow
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtDirectory 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   4575
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1575
      Left            =   120
      TabIndex        =   5
      Top             =   4440
      Width           =   1815
      Begin VB.Image imgPreview 
         Height          =   1215
         Left            =   120
         Stretch         =   -1  'True
         Top             =   260
         Width           =   1575
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Files that have been compared"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Label lblStatus 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3840
      TabIndex        =   10
      Top             =   1320
      Width           =   3855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' Written by L. "Mike" Trivette
' Please send me comments at mtrivette@yahoo.com
'
'
'
'

Option Explicit

Public m_CRC As clsCRC
Dim iFCount As Long
Dim EscapeKey As Boolean

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const BIF_BROWSEFORCOMPUTER = &H1000
Private Const MAX_PATH = 260

Private Declare Function SHBrowseForFolder Lib "shell32" _
    (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" _
    (ByVal pidList As Long, _
    ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
    (ByVal lpString1 As String, ByVal _
    lpString2 As String) As Long
    
Private Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Private Sub chkPreview_Click()
    If chkPreview.Value = 1 Then
        imgPreview.Visible = True
    Else
        imgPreview.Visible = False
    End If
End Sub

Private Sub cmdBrowse_Click()
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo
    Dim RecVal
    
    ' Start process of working directory over
    ' and clean up controls and variables.
    EscapeKey = False
    iFCount = 0
    szTitle = "Select Folder"
    cmdStartDir.Enabled = False
    
    ' Get folder from user
    With tBrowseInfo
        .hwndOwner = Me.hWnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    
    lpIDList = SHBrowseForFolder(tBrowseInfo)

    If (lpIDList) Then
        cmdStartDir.Enabled = True
        
        ' Clean up and format the directory chosen
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        If Right(sBuffer, 1) <> "\" Then
            sBuffer = sBuffer & "\"
        End If
            
        txtDirectory.Text = sBuffer
        lvMain.ListItems.Clear
        lstDupes.Clear
    Else
    
    End If
End Sub

Private Sub cmdStartDir_Click()
    Dim RecVal
    
   
        
    ' Check the status of the start buttonand act accordingly
    ' escape from the process loop here is need be
    If cmdStartDir.Caption <> "Start" Then
        EscapeKey = True
        cmdStartDir.Caption = "Start"
        Exit Sub
    End If
    
    ' Start process of working directory over
    ' and clean up controls and variables.
    EscapeKey = False
    iFCount = 0
    cmdBrowse.Enabled = False
    
    ' Use user chooses folder then work over folder
    If txtDirectory.Text <> "" Then
        ' Clear controls for GUI
        cmdStartDir.Caption = "Stop"
        lvMain.ListItems.Clear
        lstDupes.Clear

        lvMain.MousePointer = 11
        'list and process all files found in folder
        'search_files sBuffer, RecVal
        FileList txtDirectory.Text
        lvMain.MousePointer = 0 'Reset mouse status to default
    Else
        cmdStartDir.Enabled = True
        
    End If
    cmdBrowse.Enabled = True
    cmdStartDir.Caption = "Start" ' Reset button to reset and ready to run again
    Beep ' give audiable sound to let user know the process is done
    
End Sub



Sub process_file(fname As String)
    On Error Resume Next
    Static te As String
    Dim LI As ListItem
    Dim test As Integer
    Dim f As String
    Dim StrSource As String
    Dim strDest As String
    Dim ftitle As String
    
    If fname = "" Then Exit Sub
    
    ftitle = GetFileName(fname)
    
    m_CRC.Algorithm = CRC32
    f = Hex(m_CRC.CalculateFile(fname))

    If Compare(f, fname) = False Then
        Set LI = lvMain.ListItems.Add(, , ftitle)
        LI.SubItems(1) = f
        LI.SubItems(2) = FileLen(fname) & " bytes"
        LI.SubItems(3) = FileDateTime(fname)
        LI.SubItems(4) = Replace(fname, ftitle, "")
        imgPreview.Picture = LoadPicture(Replace(fname, ftitle, "") & ftitle)
        LI.SubItems(5) = imgPreview.Height / 15 & " x " & imgPreview.Width / 15
    Else
        makedir txtDirectory & "\duplicate"
        If Option1.Value = True Then
            strDest = txtDirectory & "\duplicate\" & ftitle
            FileCopy fname, strDest
        End If
        Set LI = lvMain.ListItems.Add(, , ftitle)
        LI.SubItems(1) = f
        LI.SubItems(2) = FileLen(fname) & " bytes"
        LI.SubItems(3) = FileDateTime(fname)
        LI.SubItems(4) = Replace(fname, ftitle, "")
        imgPreview.Picture = LoadPicture(Replace(fname, ftitle, "") & ftitle)
        LI.SubItems(5) = imgPreview.Height / 15 & " x " & imgPreview.Width / 15
        Kill fname
    End If
    
    ' Count the number of files processed
    ' Doevents every 30 files
    iFCount = iFCount + 1
    If iFCount Mod 30 = 0 Then
        DoEvents
    End If
        
    lblStatus.Caption = iFCount & " files checked."
End Sub

Private Function Compare(obj As String, fname As String) As Boolean
    ' This is a very sloppy way of comparing the files.
    ' This will definately need some improving. (but it does work)
    Dim i As Long
    Dim Action As String
    Compare = False
    
    If Option1.Value = True Then
        Action = " [MOVED]"
    Else
        Action = " [DELETED]"
    End If
        
    If lvMain.ListItems.Count < 1 Then Exit Function
    For i = 2 To lvMain.ListItems.Count
        If lvMain.ListItems(i).SubItems(1) = obj Then
            Compare = True
            lstDupes.AddItem fname & " = " & lvMain.ListItems(i).SubItems(4) & lvMain.ListItems(i).Text & Action
            Exit Function
        End If
    Next i
End Function



''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''
''
''       Form subs are here...
''
''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
    Set m_CRC = New clsCRC
End Sub

Private Sub Form_Unload(Cancel As Integer)
    EscapeKey = True
    End
End Sub

Private Sub lstDupes_Click()
    On Error Resume Next
    imgPreview.Picture = LoadPicture(txtDirectory & "duplicate_images\" & lstDupes.Text)
End Sub

''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''
''                                    ''
''  Listview functions below here...  ''
''                                    ''
''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''
Private Sub lvMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvMain.Sorted = True
    lvMain.SortKey = ColumnHeader.Index - 1
    If lvMain.SortOrder = lvwAscending Then
        lvMain.SortOrder = lvwDescending
    Else
        lvMain.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    imgPreview.Picture = LoadPicture(lvMain.SelectedItem.SubItems(4) & lvMain.SelectedItem.Text)
End Sub

Private Sub lvMain_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Delete file if "DEL" key is pressed
    Dim response As Integer
    
    If KeyCode = 46 Then
        response = MsgBox("Are you sure you want to delete " & lvMain.SelectedItem.SubItems(4) & lvMain.SelectedItem.Text, vbQuestion + vbYesNo, "Caution")
        If response = 6 Then
            Kill lvMain.SelectedItem.SubItems(4) & lvMain.SelectedItem.Text
            lvMain.ListItems.Remove (lvMain.SelectedItem.Index)
        End If
    End If
End Sub

Private Sub makedir(strDirName As String)
    ' Wanted to keep this seperate from
    ' main loop for debug reasons
    On Error Resume Next
    MkDir strDirName
End Sub

Private Sub txtDirectory_Click()
    cmdStartDir_Click
End Sub

Function GetFileName(path As String) As String
Dim i As Long
    For i = (Len(path)) To 1 Step -1
        If Mid(path, i, 1) = "\" Then
            GetFileName = Mid(path, i + 1, Len(path) - i + 1)
            Exit For
        End If
    Next
End Function

Private Function FileList(ByVal pathname As String, Optional DirCount As Long, Optional FileCount As Long) As String
    Dim folds(1000) As String
    Dim i As Long
    Dim fold As String
    
    If Check1.Value = False Then
        GetFiles pathname
        Exit Function
    End If
    
    fold = Dir(txtDirectory, vbDirectory)
    'MsgBox txtDirectory
    'GetFiles txtDirectory
    While fold > ""
         If fold <> "." And fold <> ".." Then
            If GetAttr(txtDirectory + fold) And vbDirectory Then
              ' a folder
              i = i + 1
              folds(i) = txtDirectory & fold & "\"
              'GetFiles txtDirectory & fold & "\"
            Else
              ' a file
            End If
        End If
    fold = Dir
    Wend
    
    GetFiles pathname
    For i = 1 To i
    If InStr(1, folds(i), "duplicate") = 0 Then GetFiles folds(i)
    'MsgBox folds(i)
    Next i
    

End Function

Private Sub GetFiles(pathname As String)
    Dim first As String
    first = Dir(pathname & "\*.*", vbNormal)
    'MsgBox first
    'process_file pathname & first
    
    Do While RTrim(first) <> "" And EscapeKey = False
        'MsgBox first
        If first <> "" Then process_file pathname & first
        first = Dir
    Loop
End Sub




