VERSION 5.00
Object = "{A1FF59A3-0995-11CF-B3E8-00608C82AA8D}#1.6#0"; "PIXEZOCX.OCX"
Begin VB.Form frmScannedDocument 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Scanned Document Name"
   ClientHeight    =   3720
   ClientLeft      =   3870
   ClientTop       =   6285
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3720
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "Schema Activation"
      Height          =   1812
      Left            =   5292
      TabIndex        =   15
      Top             =   648
      Width           =   2580
      Begin VB.Frame Frame4 
         Caption         =   "Schema Name"
         Height          =   564
         Left            =   132
         TabIndex        =   18
         Top             =   732
         Width           =   2376
         Begin PixezocxLib.PixEzScanControl PixEzScanControl8 
            Height          =   264
            Left            =   120
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   240
            Width           =   2256
            _Version        =   65542
            _ExtentX        =   3969
            _ExtentY        =   476
            _StockProps     =   77
            ControlType     =   6
            Data            =   0
            Caption         =   ""
         End
      End
      Begin PixezocxLib.PixEzScanControl PixEzScanControl6 
         Height          =   396
         Left            =   1440
         TabIndex        =   20
         Top             =   1320
         Width           =   1068
         _Version        =   65542
         _ExtentX        =   1879
         _ExtentY        =   688
         _StockProps     =   77
         ControlType     =   8
         Data            =   0
         Caption         =   "&Edit Schema..."
      End
      Begin PixezocxLib.PixEzScanControl PixEzScanControl7 
         Height          =   300
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1080
         _Version        =   65542
         _ExtentX        =   1905
         _ExtentY        =   529
         _StockProps     =   77
         ControlType     =   14
         Data            =   0
         Caption         =   "&Multi-Page"
      End
      Begin PixezocxLib.PixEzScanControl PixEzScanControl5 
         Height          =   300
         Left            =   1200
         TabIndex        =   17
         Top             =   240
         Width           =   1212
         _Version        =   65542
         _ExtentX        =   2143
         _ExtentY        =   529
         _StockProps     =   77
         ControlType     =   7
         Data            =   0
         Caption         =   "&Use Schema"
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "File Format Properties"
      Height          =   780
      Left            =   36
      TabIndex        =   10
      Top             =   2856
      Width           =   5256
      Begin PixezocxLib.PixEzScanControl PixEzScanControl1 
         Height          =   315
         Left            =   2520
         TabIndex        =   14
         Top             =   480
         Width           =   2010
         _Version        =   65542
         _ExtentX        =   3545
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   16777215
         ControlType     =   13
         Data            =   0
         Caption         =   ""
      End
      Begin PixezocxLib.PixEzScanControl cboColorFormat 
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   2025
         _Version        =   65542
         _ExtentX        =   3572
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   16777215
         ControlType     =   12
         Data            =   0
         Caption         =   ""
      End
      Begin VB.Label lblCompression 
         Caption         =   "&Compression:"
         Height          =   204
         Left            =   2568
         TabIndex        =   13
         Top             =   192
         Width           =   1020
      End
      Begin VB.Label lblColorFormat 
         Caption         =   "Color &Format:"
         Height          =   204
         Left            =   36
         TabIndex        =   11
         Top             =   192
         Width           =   1008
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   396
      Left            =   6624
      TabIndex        =   23
      Top             =   3240
      Width           =   1260
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   396
      Left            =   5340
      TabIndex        =   22
      Top             =   3240
      Width           =   1260
   End
   Begin VB.DriveListBox drvList 
      Height          =   288
      Left            =   3108
      TabIndex        =   9
      Top             =   2520
      Width           =   2124
   End
   Begin VB.FileListBox filList 
      Height          =   1065
      Left            =   84
      Pattern         =   "*.TIF"
      TabIndex        =   2
      Top             =   696
      Width           =   2928
   End
   Begin VB.TextBox txtSearch 
      Height          =   300
      Left            =   84
      TabIndex        =   1
      Text            =   "*.TIF"
      Top             =   300
      Width           =   2928
   End
   Begin VB.DirListBox dirList 
      Height          =   1536
      Left            =   3096
      TabIndex        =   5
      Top             =   696
      Width           =   2124
   End
   Begin PixezocxLib.PixEzScanControl PixEzScanControl2 
      Height          =   300
      Left            =   5280
      TabIndex        =   21
      Top             =   2640
      Width           =   1680
      _Version        =   65542
      _ExtentX        =   2963
      _ExtentY        =   529
      _StockProps     =   77
      ControlType     =   0
      Data            =   0
      Caption         =   "Warn on Over&write"
   End
   Begin PixezocxLib.PixEzScanControl ctlFileTypes 
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   2910
      _Version        =   65542
      _ExtentX        =   5133
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   16777215
      ControlType     =   11
      Data            =   0
      Caption         =   ""
   End
   Begin VB.Label Label9 
      Caption         =   "File &Name:"
      Height          =   204
      Left            =   84
      TabIndex        =   0
      Top             =   48
      Width           =   1944
   End
   Begin VB.Label Label8 
      Caption         =   "List Files of &Type:"
      Height          =   204
      Left            =   72
      TabIndex        =   6
      Top             =   2316
      Width           =   1332
   End
   Begin VB.Label Label7 
      Caption         =   "Dri&ves:"
      Height          =   204
      Left            =   3096
      TabIndex        =   8
      Top             =   2316
      Width           =   624
   End
   Begin VB.Label Label1 
      Caption         =   "&Directories:"
      Height          =   204
      Left            =   3096
      TabIndex        =   3
      Top             =   48
      Width           =   1164
   End
   Begin VB.Label lblDirName 
      Height          =   192
      Left            =   3096
      TabIndex        =   4
      Top             =   300
      Width           =   4716
   End
End
Attribute VB_Name = "frmScannedDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ctrActivePixEzImage As PixEzImage
Private savedScanFileRoot As String
Private lReturn As Long
Private firsttime As Integer
Private lSavedScanPackaging As Long
Private lSavedScanPack As Long
Private SavedScanString As String
Private ScanDocSavedPath As String
Option Explicit

Private Sub ctlFileTypes_Click()
    Dim tmpString As String
    Dim tmpCurrentPattern As String
    tmpCurrentPattern = filList.Pattern
    On Error GoTo ErrorHandler
    tmpString = Right$(ctlFileTypes.Caption, 6)
    If (Left(tmpString, 1) = "*") Then
        filList.Pattern = Left$(tmpString, 5)
    End If
    If (Len(txtSearch.Text) = 0 Or InStr(1, txtSearch.Text, "*") <> 0) Then
      ' tmpCurrentPattern <> filList.Pattern) and ) Then txtSearch.Text = tmpCurrentPattern
        txtSearch.Text = filList.Pattern
        txtSearch.SelStart = 0          ' Highlight the current entry.
        txtSearch.SelLength = Len(txtSearch.Text)
    End If
    Exit Sub
ErrorHandler:
   filList.Pattern = tmpCurrentPattern
End Sub


Private Sub dirList_GotFocus()
    cmdOK.Default = False
End Sub

Private Sub filList_Click()
   txtSearch.Text = Trim(filList.List(filList.ListIndex))
   'cmdOk.SetFocus
End Sub

Private Sub filList_DblClick()
    txtSearch.Text = Trim(filList.List(filList.ListIndex))
    cmdOK_Click
End Sub

Private Sub filList_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmdOK_Click
   End If
End Sub

Private Sub Form_Load()
   PixLocalizeForm Me
   ' center form
   '
   Me.Move (Screen.width - Me.width) / 2, _
        (Screen.Height - Me.Height) / 2
End Sub

Public Function StartForm(ctrPixEzImage As PixEzImage) As Long
    Dim ScanPrecedence, ScanPackaging, ScanWarnOnOverwrite, ScanColorformat, ScanCompression, ScanMultipage, ScanUseSchema As Long
    Dim ScanFileSchema As String

    Set ctrActivePixEzImage = ctrPixEzImage
    savedScanFileRoot = ctrActivePixEzImage.ScanFileRoot
    lSavedScanPackaging = ctrActivePixEzImage.ScanPackaging
#If Win16 Then
    ctrActivePixEzImage.ScanAllowLongNames = 0
#ElseIf Win32 Then
    ctrActivePixEzImage.ScanAllowLongNames = 1
#End If
    If firsttime Then
        ctrActivePixEzImage.ScanPackaging = lSavedScanPack
        txtSearch = SavedScanString
        filList.Pattern = SavedScanString
        On Error Resume Next
        dirList.Path = ScanDocSavedPath
        If Err Then
            dirList.Path = App.Path
        End If
        drvList.Drive = dirList.Path
    Else
         dirList.Path = CurDir: drvList.Drive = dirList.Path ' Reset the path.
         ScanDocSavedPath = dirList.Path
    End If
    
    firsttime = True
    ctlFileTypes_Click
    filList.Pattern = txtSearch.Text
    'dirList.Path = CurDir: drvList.Drive = dirList.Path ' Reset the path.
    lblDirName = dirList.Path
    
    ' Updated this section to discard the changes on Cancel
    ScanPrecedence = ctrActivePixEzImage.ScanPrecedence
    ScanPackaging = ctrActivePixEzImage.ScanPackaging
    ScanWarnOnOverwrite = ctrActivePixEzImage.ScanWarnOnOverwrite
    ScanColorformat = ctrActivePixEzImage.ScanColorformat
    ScanCompression = ctrActivePixEzImage.ScanCompression
    ScanFileSchema = ctrActivePixEzImage.ScanFileSchema
    ScanMultipage = ctrActivePixEzImage.ScanMultipage
    ScanUseSchema = ctrActivePixEzImage.ScanUseSchema
   
    Me.Show 1
    StartForm = lReturn
    If (lReturn = 0) Then ' Back to Original value
        ctrActivePixEzImage.ScanPrecedence = ScanPrecedence
        ctrActivePixEzImage.ScanPackaging = ScanPackaging
        ctrActivePixEzImage.ScanWarnOnOverwrite = ScanWarnOnOverwrite
        ctrActivePixEzImage.ScanColorformat = ScanColorformat
        ctrActivePixEzImage.ScanCompression = ScanCompression
        ctrActivePixEzImage.ScanFileSchema = ScanFileSchema
        ctrActivePixEzImage.ScanMultipage = ScanMultipage
        ctrActivePixEzImage.ScanUseSchema = ScanUseSchema
    End If
    
    On Error Resume Next
    lSavedScanPack = ctrActivePixEzImage.ScanPackaging
    SavedScanString = filList.Pattern
    ScanDocSavedPath = dirList.Path
    Unload Me
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Cancel = True       'any non-zero value stops the unload
        cmdCancel_Click     ' if user closed thru control box
                            '  treatit like it is cancel.
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set ctrActivePixEzImage = Nothing
End Sub

Private Sub cmdCancel_Click()
    ctrActivePixEzImage.ScanFileRoot = savedScanFileRoot
    Me.Hide
    lReturn = 0
End Sub

Private Sub cmdOK_Click()
    Dim MsgString As String
    If (ctrActivePixEzImage.ScanUseSchema = 1 And ctrActivePixEzImage.ScanFileSchema = "") Then
        MsgString = EW2W("No Schema Specified!!") + Chr(13) + EW2W("Please Specify Schema or Uncheck Use Schema")
        MsgBox MsgString, vbOKOnly + vbCritical, EW2W("Schema Error")
        Exit Sub
    End If
    
    Dim strFullFileName As String
    Dim strFileName As String
    Dim FileName As String
    FileName = Trim(txtSearch.Text)
    If (FileName <> "" And InStr(1, FileName, "*") = 0 And InStr(1, FileName, "?") = 0) Then
        If (Mid(FileName, 2, 1) = ":" Or Left(FileName, 2) = "\\") Then  'Network path support
           strFullFileName = FileName
        Else ' It is not full file name
            If (Right$(lblDirName, 1) <> "\") Then
              strFullFileName = lblDirName & "\" & txtSearch.Text
            Else
              strFullFileName = lblDirName & txtSearch.Text
            End If
        End If
#If 0 Then  ' we now do this in SaveWarnOnOverwrite event callback
        If (ctrActivePixEzImage.ScanWarnOnOverwrite) Then
              strFileName = Dir$(strFullFileName)
              If (strFileName <> "" And UCase(strFileName) = UCase(Right(strFullFileName, Len(strFileName)))) Then
                    Dim strMsgString As String
                    strMsgString = EW2W("File Already Exists ") & Chr(13) & Chr(10) & EW2W("Do you want to overwrite?") & Chr(13) & Chr(10)
                    If (vbYes = MsgBox(strMsgString, vbYesNo, EW2W("File Exist Error"))) Then
                        GoTo SetFileName
                    Else
                        Exit Sub
                    End If
              End If
        End If
#End If

SetFileName:
        On Error Resume Next
'        filList.fileName = fileName
        Dim nLen As String
        nLen = Len(FileName)
        If (FileName = ".." Or Len(filList.FileName) = 0 And (Right(FileName, 1) = "\" Or Right(FileName, 1) = ":" Or filList.Path = FileName)) Then
         txtSearch_KeyPress (13)
         Exit Sub
        End If
        ctrActivePixEzImage.ScanFileName = strFullFileName 'filList.Path & "\\" & filList.fileName
        Me.Hide
        lReturn = 1
   Else
        MsgBox EW2W("Invalid File Name")
   End If
End Sub

Private Sub dirList_Change()
    ' Update the file list box to synchronize with the directory list box.
    lblDirName = dirList.Path
    filList.Path = dirList.Path
End Sub
Private Sub DirList_LostFocus()
    'dirList.Path = dirList.List(dirList.ListIndex)
    dirList.ListIndex = -1
    cmdOK.Default = True
End Sub
Private Sub dirList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dirList.Path = dirList.List(dirList.ListIndex)
    End If
End Sub

Private Sub drvList_Change()
    On Error GoTo DriveHandler
    dirList.Path = drvList.Drive
    Exit Sub

DriveHandler:
    drvList.Drive = dirList.Path
    Exit Sub
End Sub

Private Sub txtSearch_Change()
    On Error Resume Next
    If (InStr(1, txtSearch.Text, ".") = 0 And InStr(1, txtSearch.Text, ":") = 0 And InStr(1, txtSearch.Text, "/") = 0 And InStr(1, txtSearch.Text, "\") = 0) Then
        ctrActivePixEzImage.ScanFileRoot = txtSearch.Text 'filList.Path & "\\" & filList.fileName
    Else
        ctrActivePixEzImage.ScanFileRoot = savedScanFileRoot
    End If
End Sub

Private Sub txtSearch_GotFocus()
    txtSearch.SelStart = 0          ' Highlight the current entry.
    txtSearch.SelLength = Len(txtSearch.Text)
    cmdOK.Default = False
End Sub

Public Sub OnImageChange()
'    MsgBox "OnChange"
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    Dim FileName As String
    Dim tmpfileName As String
    Dim Path As String
    
    If KeyAscii <> 13 Then
      Exit Sub
    End If
    
    KeyAscii = 0                     ' for the beep
    
    tmpfileName = txtSearch.Text
    If tmpfileName = "" Then
        Exit Sub
    End If
    
    'This string could be one of the following
        ' a pathname (directory name)(full or relative)
        ' a wildcard pattern (with or without a full path name)
        ' an existing filename (with or without the path)
        ' a non-existing filename (with or without the path)
        
    'All but the last one can be handled by setting it to the filList.Filename
    'For the lastcase we set it anyway, and catch the rebound.
    '    If it is a full path name - change to that directory display just the file _
    '    in the txtSearch and await the user to accept or do something else.
    '    If it is an ordinary filename in current directory just wait for the user_
    '    to click Ok and accept it.
    
    Err.Clear
    On Error Resume Next
    filList.FileName = tmpfileName
    If (Err = 53) Then ' file not found
        cmdOK_Click
        Exit Sub
    End If
    'filListBox rejected the string
    If Err Then
        Path = GetPath(tmpfileName)  ' try once to extract a valid path from the string
        If Path <> "" Then
            On Error Resume Next
            filList.FileName = Path
            If Err Then    'Not a valid path.Leave the string in the text box
                Debug.Print ""
            Else
                dirList.Path = filList.Path
                drvList.Drive = dirList.Path
                txtSearch.Text = Right(tmpfileName, Len(tmpfileName) - Len(Path) - 1)
                Exit Sub
            End If
        End If
        
        txtSearch.Text = tmpfileName
        txtSearch_GotFocus
        Exit Sub
    End If
    
    'filList box has accepted.
    On Error GoTo ErrorHandler
    dirList.Path = filList.Path
    drvList.Drive = dirList.Path
    If (filList.FileName = "") Then
        If (filList.Pattern <> tmpfileName And filList.Path <> Left(tmpfileName, Len(filList.Path))) Then
            ctlFileTypes_Click
        Else
         txtSearch.Text = "" ' Update for the bug, valid dir clear edit box
        End If
    End If
    FileName = Dir$(tmpfileName)
    If (FileName <> "" And UCase(FileName) = UCase(Right(tmpfileName, Len(FileName)))) Then
            txtSearch_GotFocus
    Else
         txtSearch.Text = "" ' Update for the bug, valid dir clear edit box
    End If
    Exit Sub
   
ErrorHandler:
    Select Case Err.Number  ' Evaluate error number.
        Case 76 ' "File already open" error.
            MsgBox EW2W("Path is Invalid..."), vbOKOnly
            Exit Sub
        Case Else
            ' Handle other situations here...
            Resume Next ' Resume execution at same line
    End Select
    Resume Next ' Resume execution at same line

End Sub

Private Sub txtSearch_LostFocus()
    On Error Resume Next
    If (InStr(1, txtSearch.Text, ".") = 0 And InStr(1, txtSearch.Text, ":") = 0 And InStr(1, txtSearch.Text, "/") = 0 And InStr(1, txtSearch.Text, "\") = 0) Then
        ctrActivePixEzImage.ScanFileRoot = txtSearch.Text 'filList.Path & "\\" & filList.fileName
    Else
        ctrActivePixEzImage.ScanFileRoot = savedScanFileRoot
    End If
    cmdOK.Default = True
End Sub
Public Function GetPath(Path As String) As String
    If InStr(1, Path, "\") = 0 Then
        GetPath = ""
        Exit Function
    End If
    
    Dim i As Integer
    
    i = Len(Path)
    While (Mid(Path, i, 1) <> "\" And i <> 0)
      i = i - 1
    Wend
    
    GetPath = Left(Path, i - 1)
End Function















