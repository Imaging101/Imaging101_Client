VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmImport 
   Caption         =   "Import"
   ClientHeight    =   8115
   ClientLeft      =   255
   ClientTop       =   540
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   ScaleHeight     =   8115
   ScaleWidth      =   10590
   Begin VB.CommandButton cmdSample 
      Caption         =   "&Sample"
      Enabled         =   0   'False
      Height          =   855
      Left            =   6840
      Picture         =   "frmImport.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Left            =   3240
      Top             =   0
   End
   Begin VB.PictureBox picImaging101Logo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   8160
      Picture         =   "frmImport.frx":08CA
      ScaleHeight     =   735
      ScaleWidth      =   2415
      TabIndex        =   15
      Top             =   0
      Width           =   2415
   End
   Begin VB.TextBox txtDocumentRECID 
      Enabled         =   0   'False
      Height          =   405
      Left            =   8520
      TabIndex        =   13
      Text            =   "000001056"
      Top             =   6240
      Width           =   1935
   End
   Begin VB.TextBox txtFileRoom 
      Height          =   285
      Left            =   1680
      TabIndex        =   11
      Text            =   "ImportUser"
      Top             =   4320
      Width           =   2655
   End
   Begin VB.TextBox txtApplicationRECID 
      Height          =   285
      Left            =   1680
      TabIndex        =   10
      Text            =   "67"
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox txtSourceUNCFilePath 
      Height          =   285
      Left            =   1680
      TabIndex        =   9
      Text            =   "C:\IMAGING101DEMO\SAMPLES\BI-TONAL\00001.tif"
      Top             =   3600
      Width           =   8775
   End
   Begin VB.TextBox txtItemCounterPass1 
      Height          =   375
      Left            =   8520
      TabIndex        =   7
      Text            =   "1"
      Top             =   5520
      Width           =   1335
   End
   Begin VB.TextBox txtFilePattern 
      Height          =   315
      Left            =   4440
      TabIndex        =   5
      Text            =   "*.TXT"
      Top             =   1080
      Width           =   3495
   End
   Begin VB.FileListBox File1 
      Height          =   1260
      Left            =   4440
      TabIndex        =   4
      Top             =   1560
      Width           =   4215
   End
   Begin VB.TextBox txtInputLine 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3120
      Width           =   10335
   End
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   4215
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   4215
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "&Import"
      Height          =   855
      Left            =   8640
      Picture         =   "frmImport.frx":16C3
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7080
      Width           =   1335
   End
   Begin MSMask.MaskEdBox mebIndexValues 
      Height          =   375
      HelpContextID   =   1
      Index           =   0
      Left            =   2400
      TabIndex        =   22
      Top             =   4920
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   15
      Mask            =   "130160100890009"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mebIndexValues 
      Height          =   375
      HelpContextID   =   1
      Index           =   1
      Left            =   2400
      TabIndex        =   37
      Top             =   5280
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   6
      Mask            =   "158843"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtFieldIsRequiredForCommit 
      Height          =   285
      Index           =   0
      Left            =   2400
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtFieldName 
      Height          =   285
      Index           =   0
      Left            =   5640
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtFieldSize 
      Height          =   285
      Index           =   0
      Left            =   5280
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtFieldType 
      Height          =   285
      Index           =   0
      Left            =   4920
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtBatchFieldsRECID 
      Height          =   285
      Index           =   0
      Left            =   2760
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtFieldsRECID 
      Height          =   285
      Index           =   0
      Left            =   3120
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtFieldDefaultValue 
      Height          =   285
      Index           =   0
      Left            =   3480
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtFieldHighValue 
      Height          =   285
      Index           =   0
      Left            =   4200
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtFieldLowValue 
      Height          =   285
      Index           =   0
      Left            =   3840
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtFieldIsSticky 
      Height          =   285
      Index           =   0
      Left            =   4560
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSMask.MaskEdBox mebIndexValues 
      Height          =   375
      HelpContextID   =   1
      Index           =   2
      Left            =   2400
      TabIndex        =   39
      Top             =   5640
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "2006"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mebIndexValues 
      Height          =   375
      HelpContextID   =   1
      Index           =   3
      Left            =   2400
      TabIndex        =   41
      Top             =   6000
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   49
      Mask            =   "Foerst, James C ,   (H), Foerst , Joyce J ,   (W)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mebIndexValues 
      Height          =   375
      HelpContextID   =   1
      Index           =   4
      Left            =   2400
      TabIndex        =   43
      Top             =   6360
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   10
      Mask            =   "Exemptions"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mebIndexValues 
      Height          =   375
      HelpContextID   =   1
      Index           =   5
      Left            =   2400
      TabIndex        =   45
      Top             =   6720
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   11
      Mask            =   "Application"
      PromptChar      =   "_"
   End
   Begin VB.Label lblFieldDescription 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Document Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   46
      Top             =   6720
      Width           =   2175
   End
   Begin VB.Label lblFieldDescription 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Document Group"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   44
      Top             =   6360
      Width           =   2175
   End
   Begin VB.Label lblFieldDescription 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Owner Names"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   42
      Top             =   6000
      Width           =   2175
   End
   Begin VB.Label lblFieldDescription 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   40
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Label lblFieldDescription 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Exemption ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   38
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Label Label10 
      Caption         =   "Drive"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackColor       =   &H0080FFFF&
      Caption         =   "Application Fields"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   4680
      Width           =   5895
   End
   Begin VB.Label lblFieldDescription 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Parcel ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   23
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label Label8 
      Caption         =   "UserID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "ApplicationRECID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "FilePath"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Input Line"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblServer 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Import Files To Archive"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   225
      Left            =   8040
      TabIndex        =   17
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      ForeColor       =   &H00808000&
      Height          =   225
      Left            =   8280
      TabIndex        =   16
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Last Document RECID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   14
      Top             =   6000
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Import to Archive"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Items Processed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   8
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "File Pattern"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   6
      Top             =   840
      Width           =   1695
   End
End
Attribute VB_Name = "frmImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim txtFullPathName As String
    Dim txtDestinationFilename As String
    Dim txtFileDestination As String
    Dim SourceUNCFilePath As String
    Dim FileName As String
                    
    Dim UNCFilePath  As String
    
    Dim Fileroom As String
    Dim Filecabinet As String
    Dim DocumentDate As String
    Dim DateAdded As String
    Dim DocumentType  As String
    Dim Folder  As String
    Dim FolderDescription  As String
    Dim DocumentSubType   As String
    Dim DocumentExpireDate  As String
    Dim DocumentNote   As String
                    
    Dim BatchID As String
    Dim PageCount As Long
    Dim DocumentRECID As Double
    
    Dim arrFullDirSections() As String
    Dim txtFullDir As String
    Dim txtNewDir As String
    Dim flagMultiPageDocument As Boolean
    Dim flagDocumentInProcess As Boolean
    



Public Sub cmdImport_Click()

    '*** SET the BATCH ID to the current Date and Time
    txtBatchID = Format(Now, "yyyy-mm-dd hh:mm:ss AM/PM")
    
    '*** CHECK FOR NO FILE SELECTED
    If File1.FileName = "" Then
        Result = MsgBox("Please select a File to process.", vbOKOnly)
        Exit Sub
    End If
    
    txtInputFilePath = File1.path + "\" + File1.FileName

    '*** CHECK FOR RE-START
    txtOutputTempFilePath = Left(txtInputFilePath, Len(txtInputFilePath) - 4) + ".TX2"
    If subFileExists(txtOutputTempFilePath) Then
        Result = MsgBox("I noticed that the previous run did not complete properly.  Would you like to continue where you left off?", vbYesNo)
        If Result = vbNo Then
            Result = MsgBox("You may want to call PC Networks to see why my files are in an inconsistent state!", vbOKOnly)
            blnContinueLastRun = False
        Else
            blnContinueLastRun = True
        End If
    End If
    
    
    Open txtInputFilePath For Input As #1   ' Open file for input.
    txtOutputTempFilePath = Left(txtInputFilePath, Len(txtInputFilePath) - 4) + ".TX2"
    
    ' Reset Item Counter
    txtItemCounterPass1 = 0
    
    
    If blnContinueLastRun = True Then
        Open txtOutputTempFilePath For Input As #2 ' Open LST file
        ' Walk down the TX2 file to
        '    find the Last Line Processed in the previous run
        numLineCount = 1
        Do While Not EOF(2)
            numLineCount = numLineCount + 1
            txtItemCounterPass1 = numLineCount
            DoEvents
            Line Input #2, INPUTLINE
            If Trim(INPUTLINE) <> "" Then    'Check for blank line
                Line Input #1, INPUTLINE    ' Read line of data.
            End If
        Loop
''        Line Input #1, InputLine    ' Read ONE MORE line...
        


        ' Changed from "Output" to "Append" so we don't
        '  the contents of the TX2 file
        Close #2
        Open txtOutputTempFilePath For Append As #2 ' Open LST file
    Else
        Open txtOutputTempFilePath For Output As #2 ' Open LST file
    End If
    
    '*** CONNECT TO CONTROL DATABASE TABLE
'    frmImport.Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Password=;Data Source=C:\WorkArea\Jacob\Program\Imaging101\Imaging101.MDB"
'    frmImport.Adodc1.RecordSource = "select BatchID, DocumentRECID From Control"
'    frmImport.Adodc1.Refresh

    
'    Text1.Text = ""
    txtMaxWidth = 0
    txtPageLines = 0
    txtCharString = ""
    numMonth = ""
    txtImageFileName = ""
    txtPreviousBatch = ""
    
    
    
    While Not EOF(1) ' Check for end of file.
        Line Input #1, INPUTLINE    ' Read line of data.
        
        ' Skip processing if Header or Blank lines
        If Trim(INPUTLINE) = "ImageFile~SSN" Or Trim(INPUTLINE) = "" Then
            GoTo GetNextLine
        End If
        
        ' Add a "Delimiter" at the end to make sure we get the LAST field
        INPUTLINE = INPUTLINE + frmConfig.txtInputFileDelimiter
        
        ' Replace Old Delimiters
        intCountDelim = 0
        intHoldPrevPos = 0
       
            For numCharCount = 1 To Len(INPUTLINE)
                If Mid(INPUTLINE, numCharCount, 1) = frmConfig.txtInputFileDelimiter Then
                    
                    intCountDelim = intCountDelim + 1
                    txtHoldField = Mid(INPUTLINE, intHoldPrevPos + 1, numCharCount - intHoldPrevPos - 1)
                    
                    ' Zero-fill Field
                    Select Case intCountDelim
                        Case 1    '
                            txtSourceUNCFilePath = txtHoldField
                        Case 2    '
                            txtFileRoom = txtHoldField
                        Case 3    '
                            txtFileCabinet = txtHoldField
                        Case 4    '
                            txtDateAdded = txtHoldField
                        Case 5    '
                            txtDocumentDate = txtHoldField
                        Case 6    '
                            txtDocumentType = txtHoldField
                        Case 7    '
                            txtFolder = txtHoldField
                        Case 8    '
                            txtFolderDescription = txtHoldField
                        Case 9    '
                            txtPageCount = txtHoldField
                        Case 10   '
                        Case Else
'                            txtImageFileName = txtHoldField
'                            txtDocumentSubType = txtHoldField
'                            txtDocumentExpireDate = txtHoldField
'                            txtDocumentNote = txtHoldField
                    End Select
                        
                  intHoldPrevPos = numCharCount
                
                End If
            
            Next
            
            txtDateAdded = Format(Now(), "yyyy-mm-dd")


            '***  Break Out File Parts
'            lenImageFileName = Len(txtImageFileName)
'            lenImageFileName = lenImageFileName - 13
'            txtImageFilePath = Left(txtImageFileName, lenImageFileName)
'            txtImageFile = Right(txtImageFileName, 12)
'            txtImageBatch = Right(txtImageFilePath, 8)
'            txtOutputFile = txtImageFilePath + "\P" + Right(txtImageFile, 11)
            
            
            ' Display Single Output Line being processed
            frmImport.txtInputLine = INPUTLINE
            
             ' Write Output Line to Output File
            OUTPUTLINE = Left(INPUTLINE, Len(INPUTLINE) - 1)
            Print #2, OUTPUTLINE
            OUTPUTLINE = ""
            

           DoEvents


        '******************************************************************
        '*** BEGIN SPICER OPEN / INSERT SECTION
        
        txtFullPathName = txtSourceUNCFilePath
            
        DoEvents
 
        
        CheckForNextPageOfMultiPageDocument
        CheckForLastPageOfDocument
        CheckForFirstPageOfDocument
        
        '*** END   SPICER OPEN / INSERT SECTION
        '******************************************************************
        
        
            
            txtItemCounterPass1 = txtItemCounterPass1 + 1
            DoEvents
    
GetNextLine:

    Wend
    
    CheckForLastPageOfDocument
    
    
    txtInputLine = "***  IMPORT COMPLETE ***"
    
    On Error Resume Next
    Close #1
    Close #2
    
    On Error GoTo 0
    
    Set frmViewForm = Nothing
    Me.Hide
    
    

End Sub
        
        
                
Public Sub CheckForNextPageOfMultiPageDocument()
        If txtPageCount > 1 Then
                    
                '****************************************************************
                '***
                '*** N E X T   P A G E   O F   A   MULTI-PAGE   D O C U M E N T
                '***
                '***             IMPORT AND APPEND IT
                '***
                '****************************************************************
                
                    Set docContents = frmViewForm.SpicerDoc1.object
                   
                    flagMultiPageDocument = True
                    
'                            PageCount = frmImport.txtPageCount
                            
                   ' Import the selected file at the end of the document
                   '    IN_NEWPAGE_BEFORE (0)
                   '    IN_NEWPAGE_AFTER (1)
                   '    IN_NEWPAGE_BEGIN (2)
                   '    IN_NEWPAGE_END (3)
                
                   docContents.ImportPage 0, 0, IN_NEWPAGE_END, Left(FileName, Len(FileName) - 4) & "_" & txtPageCount, txtFullPathName
                
        End If
        
End Sub
        
Public Sub CheckForLastPageOfDocument()

        If (txtPageCount = 1 And flagDocumentInProcess) Or EOF(1) = True Then
        
                '***********************************************************
                '***
                '*** L A S T   P A G E   P R O C E S S E D  - SAVE & RESET
                '***
                '***********************************************************
                
                
                On Error GoTo 0
                
        
                '*** ADD RECORD TO THE DATABASE
                frmImaging101Retrieve.Adodc1.Recordset.AddNew
                  
                
                BatchID = frmImport.txtBatchID
                PageCount = txtPageCount
                
                frmImaging101Retrieve.Adodc1.Recordset!BatchID = BatchID
                
                frmImaging101Retrieve.Adodc1.Recordset!DocumentRECID = DocumentRECID
                frmImaging101Retrieve.Adodc1.Recordset!UNCFilePath = UNCFilePath
                frmImaging101Retrieve.Adodc1.Recordset!FileName = FileName
                frmImaging101Retrieve.Adodc1.Recordset!Fileroom = Fileroom
                frmImaging101Retrieve.Adodc1.Recordset!Filecabinet = Filecabinet
                frmImaging101Retrieve.Adodc1.Recordset!DateAdded = DateAdded
                frmImaging101Retrieve.Adodc1.Recordset!DocumentDate = DocumentDate
                frmImaging101Retrieve.Adodc1.Recordset!DocumentType = DocumentType
                frmImaging101Retrieve.Adodc1.Recordset!Folder = Folder
                frmImaging101Retrieve.Adodc1.Recordset!FolderDescription = FolderDescription
                frmImaging101Retrieve.Adodc1.Recordset!DocumentSubType = DocumentSubType
                frmImaging101Retrieve.Adodc1.Recordset!DocumentExpireDate = DocumentExpireDate
                frmImaging101Retrieve.Adodc1.Recordset!DocumentNote = DocumentNote
                frmImaging101Retrieve.Adodc1.Recordset!PageCount = PageCount
                
                '***  RETURN TO NORMAL ERROR TRAPPING
                On Error GoTo 0
                
                    
                If UCase(Right(SourceUNCFilePath, 3)) = "TIF" And flagMultiPageDocument = True Then
                    ' EXPORT the Multi-Page TIF Document
                    Dim docSave As IDocSave
                    Set docContents = frmViewForm.SpicerDoc1.object
    
                    ' Set the object variable for the IDocSave interface to the Document Control object
                    Set docSave = frmViewForm.SpicerDoc1.object

                    docSave.Export docContents.RootID, True, API_MPAGE_TIFF, txtFileDestination, FileName
            
                    ' Set CloseDocument to "True" to check if the document has been changed.
                     docContents.CloseDocument False
                    
                    ' De-initialize the object variables
                    Set docSave = Nothing
                    Set docContents = Nothing
                    Set ActivePage = Nothing
                    
                Else
                    '*** MOVE and RENAME the Source File to the Destination Location
                    FileCopy SourceUNCFilePath, txtFileDestination
''                    ' SAVE the file in it's ORIGINAL Format DOESN'T WORK!!!
''                    ' Only supports exporting to few specific formats!!!
''                    docSave.Save docContents.RootID, True, , txtfiledestination, FileName
                End If
                
                
                ' UPDATE / INSERT RECORD IN DATABASE
                frmImaging101Retrieve.Adodc1.Recordset.Update
                
                
                ' RESET flagMultiPageDocument
                flagMultiPageDocument = False
                
                ' RESET flagDocumentInProcess
                flagDocumentInProcess = False
    
                
        End If
            
End Sub
        
Public Sub CheckForFirstPageOfDocument()

        If txtPageCount = 1 And flagDocumentInProcess = False Then
            
                '*****************************************************
                '***
                '*** N E W   D O C U M E N T   -  OPEN FIRST PAGE
                '***
                '*****************************************************
                
                flagDocumentInProcess = True
            
                '*** Check if file Exists
                If Not subFileExists(txtFullPathName) Then
                    Result = MsgBox("SORRY! I can't find file:" + vbNewLine + txtFullPathName + vbNewLine + "PLEASE CONTACT PC NETWORKS", vbCritical)
                    Exit Sub
                End If
                
    '            Dim docContents As IDocContents
    '            Dim ActivePage As IActivePage
    '            Dim frmViewForm As ChildForm1
                
                Set frmViewForm = New ChildForm1
                frmViewForm.Caption = txtFullPathName
                frmViewForm.Show
            
                
                '*** Close any open documents before opening a new one.
                ' Set the object variable for the IDocContents interface to the Document Control object
                Set docContents = frmViewForm.SpicerDoc1.object
                ' Close the document in the SpicerDoc1 control and
                ' Set CloseDocument to "True" to check if the document has been changed.
                 docContents.CloseDocument False
                ' De-initialize the object variable
                Set docContents = Nothing
                   
            '    ChildForm1.Show
                
                Set docContents = frmViewForm.SpicerDoc1.object
                Set ActivePage = frmViewForm.SpicerView1.object
             
                docContents.OpenFile txtFullPathName
                ActivePage.BindToDocumentControl frmViewForm.SpicerDoc1.object
                
                    'These assignments where necessary because VB
                    '  displays errors if we try to assign the
                    '  form fields directly to the adoPrimaryRS fields.
                    
                    ' Increment DocumentRECID
                    frmImport.Adodc1.Recordset!DocumentRECID = frmImport.Adodc1.Recordset!DocumentRECID + 1
                    frmImport.Adodc1.Recordset.Update
                    DocumentRECID = frmImport.Adodc1.Recordset!DocumentRECID
                    frmImport.Adodc1.Recordset.Requery
        
                    
                    ' The UNCFilePath NO Longer needs the full Path since we now have added
                    '  txtRootDirToStoreObjects configuration value to allow us to build the
                    '  full path to the document at time of retrieval!
                    UNCFilePath = "\" & txtFileRoom & "\" & txtFileCabinet
                    
                    Fileroom = frmImport.txtFileRoom
                    Filecabinet = frmImport.txtFileCabinet
                    DocumentDate = frmImport.txtDocumentDate
                    DateAdded = frmImport.txtDateAdded
                    DocumentType = frmImport.txtDocumentType
                    Folder = frmImport.txtFolder
                    FolderDescription = frmImport.txtFolderDescription
                    DocumentSubType = frmImport.txtDocumentSubType
                    DocumentExpireDate = frmImport.txtDocumentExpireDate
                    DocumentNote = frmImport.txtDocumentNote
                    
                    ' Create a concatenated DestinationFileName
                    txtDestinationFilename = Fileroom & "_" & Filecabinet & "_" & DocumentRECID & "." & Right(txtSourceUNCFilePath, 3)
                    FileName = txtDestinationFilename     'should be DocumentRECID
                    txtFileDestination = frmConfig.txtRootDirToStoreObjects & UNCFilePath & "\" & txtDestinationFilename
                
                    SourceUNCFilePath = txtSourceUNCFilePath
                    
            
                     '********************************************************
                     '*** CREATE DESTINATION DIRECTORY STRUCTURE FOR FILE
                     '********************************************************
                     
                     On Error Resume Next ' Ignore errors
                     
                    '*** Create the Root Directory Structure
                     
                     ' Remove Right Backslash if necessary
                     If Right(frmConfig.txtRootDirToStoreObjects, 1) = "\" Then
                         frmConfig.txtRootDirToStoreObjects = Left(frmConfig.txtRootDirToStoreObjects, Len(frmConfig.txtRootDirToStoreObjects) - 1)
                     End If
                     
                     txtFullDir = frmConfig.txtRootDirToStoreObjects & "\" & Fileroom & "\" & Filecabinet
    ''                    txtFullDir = frmConfig.txtRootDirToStoreObjects & "\" & txtFileRoom & "\" & txtFileCabinet & "\" & DocumentType & "_" & DocumentDate
                     
                     txtNewDir = ""
                     
                     '*** SEE IF DIRECTORY EXISTS
                     If Trim(Dir(txtFullDir, vbDirectory)) = "" Then
                     
                         '*** SPLIT THE DIRECTORY STRUCTURE INTO EACH SECTION
                          '    because the MKDIR statement can't make a subdirectory if
                          '    its root doesn't exist.
                          arrFullDirSections = Split(txtFullDir, "\")
                          '-- Loop through array to Build each Subdirectory
                          For iCounter = LBound(arrFullDirSections) To UBound(arrFullDirSections)
                              txtNewDir = txtNewDir & arrFullDirSections(iCounter) & "\"
                              MkDir txtNewDir
                          Next
                     
                     End If
                     
                     '***  RETURN TO NORMAL ERROR TRAPPING
                     On Error GoTo 0
                     
                     
            End If
                
        
End Sub

Private Sub cmdSample_Click()

    '*** CHECK FOR NO FILE SELECTED
    If File1.FileName = "" Then
        Result = MsgBox("Please select a File to process.", vbOKOnly)
        Exit Sub
    End If
    
    txtInputFilePath = File1.path + "\" + File1.FileName

    '*** CHECK FOR RE-START
    txtOutputTempFilePath = Left(txtInputFilePath, Len(txtInputFilePath) - 4) + ".TX2"
    If subFileExists(txtOutputTempFilePath) Then
        Result = MsgBox("I noticed that the previous run did not complete properly.  Would you like to continue where you left off?", vbYesNo)
        If Result = vbNo Then
            Result = MsgBox("You may want to call PC Networks to see why my files are in an inconsistent state!", vbOKOnly)
            blnContinueLastRun = False
        Else
            blnContinueLastRun = True
        End If
    End If
    
    
    Open txtInputFilePath For Input As #1   ' Open file for input.
    txtOutputTempFilePath = Left(txtInputFilePath, Len(txtInputFilePath) - 4) + ".TX2"
    
    ' Reset Item Counter
    txtItemCounterPass1 = 0
    
    
    If blnContinueLastRun = True Then
        Open txtOutputTempFilePath For Input As #2 ' Open LST file
        ' Walk down the TX2 file to
        '    find the Last Line Processed in the previous run
        numLineCount = 1
        Do While Not EOF(2)
            numLineCount = numLineCount + 1
            txtItemCounterPass1 = numLineCount
            DoEvents
            Line Input #2, INPUTLINE
            If Trim(INPUTLINE) <> "" Then    'Check for blank line
                Line Input #1, INPUTLINE    ' Read line of data.
            End If
        Loop
''        Line Input #1, InputLine    ' Read ONE MORE line...
        


        ' Changed from "Output" to "Append" so we don't
        '  the contents of the TX2 file
        Close #2
        Open txtOutputTempFilePath For Append As #2 ' Open LST file
    Else
        Open txtOutputTempFilePath For Output As #2 ' Open LST file
    End If
    
    '*** CONNECT TO CONTROL DATABASE TABLE
    frmImport.Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Password=;Data Source=C:\WorkArea\Jacob\Program\Imaging101\Imaging101.MDB"
    frmImport.Adodc1.RecordSource = "select BatchID, DocumentRECID From Control"
    frmImport.Adodc1.Refresh

    
'    Text1.Text = ""
    txtMaxWidth = 0
    txtPageLines = 0
    txtCharString = ""
    numMonth = ""
    txtImageFileName = ""
    txtPreviousBatch = ""
    
    
    
    While Not EOF(1) ' Check for end of file.
        Line Input #1, INPUTLINE    ' Read line of data.
        
        ' Skip processing if Header or Blank lines
        If Trim(INPUTLINE) = "ImageFile~SSN" Or Trim(INPUTLINE) = "" Then
            GoTo GetNextLine
        End If
        
        ' Add a "Delimiter" at the end to make sure we get the LAST field
        INPUTLINE = INPUTLINE + frmConfig.txtInputFileDelimiter
        
        ' Replace Old Delimiters
        intCountDelim = 0
        intHoldPrevPos = 0
       
            For numCharCount = 1 To Len(INPUTLINE)
                If Mid(INPUTLINE, numCharCount, 1) = frmConfig.txtInputFileDelimiter Then
                    
                    intCountDelim = intCountDelim + 1
                    txtHoldField = Mid(INPUTLINE, intHoldPrevPos + 1, numCharCount - intHoldPrevPos - 1)
                    
                    ' Zero-fill Field
                    Select Case intCountDelim
                        Case 1    '
                            txtSourceUNCFilePath = txtHoldField
                        Case 2    '
                            txtFileRoom = txtHoldField
                        Case 3    '
                            txtFileCabinet = txtHoldField
                        Case 4    '
                            txtDateAdded = txtHoldField
                        Case 5    '
                            txtDocumentDate = txtHoldField
                        Case 6    '
                            txtDocumentType = txtHoldField
                        Case 7    '
                            txtFolder = txtHoldField
                        Case 8    '
                            txtFolderDescription = txtHoldField
                        Case 9    '
                            txtPageCount = txtHoldField
                        Case 10   '
                        Case Else
'                            txtImageFileName = txtHoldField
'                            txtDocumentSubType = txtHoldField
'                            txtDocumentExpireDate = txtHoldField
'                            txtDocumentNote = txtHoldField
                    End Select
                        
                  intHoldPrevPos = numCharCount
                
                End If
            
            Next
    '*** LOAD FIELD DEFINITIONS
    subLoadFieldDefinitions
    
End Sub

Private Sub Drive1_Change()
    Dir1.path = Drive1.Drive
    File1.path = Dir1.path
    File1.Pattern = frmImport.txtFilePattern
    File1.Refresh

End Sub

Private Sub Form_Load()
    
    bolImportFormLoadComplete = False
    
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    
    '*** Disable buttons to prevent users from Clicking on them
    '    prior to the form being ready
    cmdImport.Enabled = True
    
    
    Dir1.path = Drive1.Drive
    File1.path = Dir1.path
    File1.Pattern = frmImport.txtFilePattern
    File1.Refresh
    
End Sub


Private Sub txtFilePattern_Change()
    Dir1.path = Drive1.Drive
    File1.path = Dir1.path
    File1.Pattern = frmImport.txtFilePattern

End Sub


Private Sub Timer1_Timer()
    'This timer is simply to bypass a VB Error:  "Unable to unload within this context (Error 365)"
    ' attempting to "Destroy" the fields when switching Applications
    'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/vbenlr98/html/vamsgldcantunloadhere.asp
    'apparently VB won't let you unload certain objects in certain contexts like the "_Click" event
    ' or any event or Sub that this event calls!!!
    
    'Disable this timer now
    Timer1.Enabled = False
    

    'Now load the new field definitiona
    Call subLoadFieldDefinitions
    
    '*** Re-enable buttons
    cmdFind.Enabled = True
    cmdClear.Enabled = True
    cmdHelp.Enabled = True
    bolImportFormLoadComplete = True
    
End Sub

Sub subLoadFieldDefinitions()


    Me.Enabled = False
    
    '*** THIS SUBROUTINE LOADS ALL THE APPLICATION FIELD DEFINITION INFORMATION
    '***  INCLUDING FIELD FORMAT VALUES INTO AN ARRAY.
    
    Set Con = New ADODB.Connection
    Con.Open RegImaging101ConnectionString
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = Con
    
    
    rs.Source = "Select * from I101Fields WHERE ApplicationRECID = " & txtApplicationRECID & " ORDER BY FieldOrderBatch"
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LockType = adLockReadOnly
    
   On Error GoTo ERROR_TRAP
    
    'sql statement to select items on the drop down list
''    ssql = "Select * from I101Fields where ApplicationRECID = " & txtApplicationRECID
''    rs.Open ssql, Con
    Con.Errors.Clear
    rs.Open
    

''    Debug.Print rs.PageCount
''    Debug.Print rs.RecordCount
''    Debug.Print rs.AbsolutePage
''    Debug.Print rs.AbsolutePosition
    
    
    rs.MoveFirst
    
    
    '*** DESTROY FIELDS ARRAYS
   On Error Resume Next
    
    For intIndex = 1 To lblFieldDescription.Count - 1
        Unload lblFieldDescription(intIndex)
        Unload mebIndexValues(intIndex)
        Unload txtBatchFieldsRECID(intIndex)
        Unload txtFieldsRECID(intIndex)
        Unload txtFieldDefaultValue(intIndex)
        Unload txtFieldLowValue(intIndex)
        Unload txtFieldHighValue(intIndex)
        Unload txtFieldIsSticky(intIndex)
        Unload txtFieldType(intIndex)
        Unload txtFieldSize(intIndex)
        Unload txtFieldName(intIndex)
        Unload txtFieldIsRequiredForCommit(intIndex)
'        Unload txtFieldIsRequiredForSplit(intIndex)
'        Unload txtFieldSplitBatches(intIndex)
'        Unload cmdFieldDropDown(intIndex)
    Next
   
   On Error GoTo ERROR_TRAP
        
    
    'RE-Size Form Based on How many fields we Expect if more than 10
'    If rs.RecordCount > 10 Then
'        'Increase the size of the form by the number of Fields we expect
        Dim intNewHeight As Integer
        intNewHeight = lblFieldDescription(0).Top + (lblFieldDescription(0).Height * rs.RecordCount) + 700
'        If intNewHeight > 6000 Then
            Me.Height = intNewHeight
'        Else
'            Me.Height = 6000
'        End If
'    End If
    
    '*** intFieldIndex allows us to Add the Second Date field to search through
    '     we set it to (-1) to make sure we start at Zero (0) in the Loop
    Dim intFieldIndex As Integer
    Dim bolSecondPass As Boolean
    intFieldIndex = -1
    
    '*** intTabCounter allows us to number the TAB ORDER of the fields and buttons properly
    '     we set it to (1) to make sure we start after the last FIXED/Pre-defined field
    intTabCounter = 1
    
    For intIndex = 0 To rs.RecordCount - 1
    
        'Initialize the bolFirstPass flag to track fields we want to create duplicates of...
        bolFirstPass = True

CREATE_FIELD_OBJECTS:

            intFieldIndex = intFieldIndex + 1
        
        '* Create Field Objects - BEGIN
            If intFieldIndex > 0 Then
                Load lblFieldDescription(intFieldIndex)
'                Set lblFieldDescription(intFieldIndex).Container = Frame2
                lblFieldDescription(intFieldIndex).Top = lblFieldDescription(intFieldIndex - 1).Top + lblFieldDescription(intFieldIndex - 1).Height
                lblFieldDescription(intFieldIndex).Enabled = True
                lblFieldDescription(intFieldIndex).Visible = True
                lblFieldDescription(intFieldIndex).Caption = ""
                
                Load mebIndexValues(intFieldIndex)
'                Set mebIndexValues(intFieldIndex).Container = Frame2
                mebIndexValues(intFieldIndex).Top = mebIndexValues(intFieldIndex - 1).Top + mebIndexValues(intFieldIndex - 1).Height
                mebIndexValues(intFieldIndex).Enabled = True
                mebIndexValues(intFieldIndex).Visible = True
                intTabCounter = intTabCounter + 1
                mebIndexValues(intFieldIndex).TabStop = True
                mebIndexValues(intFieldIndex).TabIndex = intTabCounter
                mebIndexValues(intFieldIndex).Text = ""
                
                Load txtBatchFieldsRECID(intFieldIndex)
'                Set txtBatchFieldsRECID(intFieldIndex).Container = Frame2
                txtBatchFieldsRECID(intFieldIndex).Top = txtBatchFieldsRECID(intFieldIndex - 1).Top + txtBatchFieldsRECID(intFieldIndex - 1).Height
                txtBatchFieldsRECID(intFieldIndex).Enabled = True
                txtBatchFieldsRECID(intFieldIndex).Visible = False

                Load txtFieldsRECID(intFieldIndex)
'                Set txtFieldsRECID(intFieldIndex).Container = Frame2
                txtFieldsRECID(intFieldIndex).Top = txtFieldsRECID(intFieldIndex - 1).Top + txtFieldsRECID(intFieldIndex - 1).Height
                txtFieldsRECID(intFieldIndex).Enabled = True
                txtFieldsRECID(intFieldIndex).Visible = False

                Load txtFieldDefaultValue(intFieldIndex)
'                Set txtFieldDefaultValue(intFieldIndex).Container = Frame2
                txtFieldDefaultValue(intFieldIndex).Top = txtFieldDefaultValue(intFieldIndex - 1).Top + txtFieldDefaultValue(intFieldIndex - 1).Height
                txtFieldDefaultValue(intFieldIndex).Enabled = True
                txtFieldDefaultValue(intFieldIndex).Visible = False
                txtFieldDefaultValue(intFieldIndex).Text = ""
                
                Load txtFieldLowValue(intFieldIndex)
'                Set txtFieldLowValue(intFieldIndex).Container = Frame2
                txtFieldLowValue(intFieldIndex).Top = txtFieldLowValue(intFieldIndex - 1).Top + txtFieldLowValue(intFieldIndex - 1).Height
                txtFieldLowValue(intFieldIndex).Enabled = True
                txtFieldLowValue(intFieldIndex).Visible = False
                txtFieldLowValue(intFieldIndex).Text = ""
            
                Load txtFieldHighValue(intFieldIndex)
'                Set txtFieldHighValue(intFieldIndex).Container = Frame2
                txtFieldHighValue(intFieldIndex).Top = txtFieldHighValue(intFieldIndex - 1).Top + txtFieldHighValue(intFieldIndex - 1).Height
                txtFieldHighValue(intFieldIndex).Enabled = True
                txtFieldHighValue(intFieldIndex).Visible = False
                txtFieldHighValue(intFieldIndex).Text = ""
                
                Load txtFieldIsSticky(intFieldIndex)
'                Set txtFieldIsSticky(intFieldIndex).Container = Frame2
                txtFieldIsSticky(intFieldIndex).Top = txtFieldIsSticky(intFieldIndex - 1).Top + txtFieldIsSticky(intFieldIndex - 1).Height
                txtFieldIsSticky(intFieldIndex).Enabled = True
                txtFieldIsSticky(intFieldIndex).Visible = False
                txtFieldIsSticky(intFieldIndex).Text = ""
            
                Load txtFieldType(intFieldIndex)
'                Set txtFieldType(intFieldIndex).Container = Frame2
                txtFieldType(intFieldIndex).Top = txtFieldType(intFieldIndex - 1).Top + txtFieldType(intFieldIndex - 1).Height
                txtFieldType(intFieldIndex).Enabled = True
                txtFieldType(intFieldIndex).Visible = False
                txtFieldType(intFieldIndex).Text = ""
            
                Load txtFieldSize(intFieldIndex)
'                Set txtFieldSize(intFieldIndex).Container = Frame2
                txtFieldSize(intFieldIndex).Top = txtFieldSize(intFieldIndex - 1).Top + txtFieldSize(intFieldIndex - 1).Height
                txtFieldSize(intFieldIndex).Enabled = True
                txtFieldSize(intFieldIndex).Visible = False
                txtFieldSize(intFieldIndex).Text = ""
                
                Load txtFieldName(intFieldIndex)
'                Set txtFieldName(intFieldIndex).Container = Frame2
                txtFieldName(intFieldIndex).Top = txtFieldName(intFieldIndex - 1).Top + txtFieldName(intFieldIndex - 1).Height
                txtFieldName(intFieldIndex).Enabled = True
                txtFieldName(intFieldIndex).Visible = False
                txtFieldName(intFieldIndex).Text = ""
            
                Load txtFieldIsRequiredForCommit(intFieldIndex)
'                Set txtFieldIsRequiredForCommit(intFieldIndex).Container = Frame2
                txtFieldIsRequiredForCommit(intFieldIndex).Top = txtFieldIsRequiredForCommit(intFieldIndex - 1).Top + txtFieldIsRequiredForCommit(intFieldIndex - 1).Height
                txtFieldIsRequiredForCommit(intFieldIndex).Enabled = True
                txtFieldIsRequiredForCommit(intFieldIndex).Visible = False
                txtFieldIsRequiredForCommit(intFieldIndex).Text = ""
            
'                '*** Create the DROP-DOWN Button
'                Load cmdFieldDropDown(intFieldIndex)
''                Set FieldDropDownList(intFieldIndex).Container = Frame2
'                cmdFieldDropDown(intFieldIndex).Top = mebIndexValues(intFieldIndex - 1).Top + mebIndexValues(intFieldIndex - 1).Height
'                cmdFieldDropDown(intFieldIndex).Enabled = True
'                intTabCounter = intTabCounter + 1
'                cmdFieldDropDown(intFieldIndex).TabStop = True
'                cmdFieldDropDown(intFieldIndex).TabIndex = intTabCounter
'                'Make the DropDownList button VISIBLE only if Checked for the current field
'                If rs.Fields!FieldDropDownList = vbChecked Then
'                    cmdFieldDropDown(intFieldIndex).Visible = True
'                Else
'                    cmdFieldDropDown(intFieldIndex).Visible = False
'                End If
                
            Else 'intFieldIndex <=  0
            
                'Make the DropDownList button VISIBLE only if Checked for the current field
'                If rs.Fields!FieldDropDownList = vbChecked Then
'                    cmdFieldDropDown(intFieldIndex).Visible = True
'                Else
'                    cmdFieldDropDown(intFieldIndex).Visible = False
'                End If
            
            End If
        '* Create Field Objects - END

        
        'Clear any Values carried over from the first (Master) field
        lblFieldDescription(intFieldIndex) = ""
        mebIndexValues(intFieldIndex).Mask = ""
        mebIndexValues(intFieldIndex).Format = ""
        mebIndexValues(intFieldIndex).Text = ""
    
        '* Assign Field Values
        txtFieldsRECID(intFieldIndex) = rs.Fields!FieldsRECID
        If (IsNull(rs.Fields!FieldNameForInput)) Or (rs.Fields!FieldNameForInput <> "") Then
            lblFieldDescription(intFieldIndex) = rs.Fields!FieldNameForInput
        Else
            lblFieldDescription(intFieldIndex) = rs.Fields!FieldName
        End If
        
        '* If setting up a Range Field - append the text "(Thru)"
        If bolFirstPass = False Then
            lblFieldDescription(intFieldIndex) = lblFieldDescription(intFieldIndex) & " (Thru)"
        End If
        
        If Not IsNull(rs.Fields!FieldMask) Then mebIndexValues(intFieldIndex).Mask = rs.Fields!FieldMask
        If Not IsNull(rs.Fields!FieldFormat) Then mebIndexValues(intFieldIndex).Format = rs.Fields!FieldFormat
        
        If Not IsNull(rs.Fields!FieldDefaultValue) Then txtFieldDefaultValue(intFieldIndex) = rs.Fields!FieldDefaultValue
        If Not IsNull(rs.Fields!FieldLowValue) Then txtFieldLowValue(intFieldIndex) = rs.Fields!FieldLowValue
        If Not IsNull(rs.Fields!FieldHighValue) Then txtFieldHighValue(intFieldIndex) = rs.Fields!FieldHighValue
        If Not IsNull(rs.Fields!FieldIsSticky) Then txtFieldIsSticky(intFieldIndex) = rs.Fields!FieldIsSticky
        If Not IsNull(rs.Fields!FieldType) Then txtFieldType(intFieldIndex) = rs.Fields!FieldType
        If Not IsNull(rs.Fields!FieldSize) Then txtFieldSize(intFieldIndex) = rs.Fields!FieldSize
        If Not IsNull(rs.Fields!FieldName) Then txtFieldName(intFieldIndex) = rs.Fields!FieldName
        If Not IsNull(rs.Fields!FieldIsRequiredForCommit) Then
            txtFieldIsRequiredForCommit(intFieldIndex) = rs.Fields!FieldIsRequiredForCommit
'            If txtFieldIsRequiredForCommit(intFieldIndex) = vbChecked Then
'                'Show that field is required!
'                lblFieldDescription(intFieldIndex) = lblFieldDescription(intFieldIndex)
'                lblFieldDescription(intFieldIndex).ForeColor = vbRed
'            Else
'                lblFieldDescription(intFieldIndex).ForeColor = vbNormal
'            End If
        End If
        
        If txtFieldType(intFieldIndex).Text = "Date" And bolFirstPass = True Then
            'Increase Form Height to account for this extra field
            Me.Height = Me.Height + lblFieldDescription(0).Height
            'Make sure we don't create it more than once.
            bolFirstPass = False
            'Create another Date field identical to this one
            GoTo CREATE_FIELD_OBJECTS
        End If
            
        rs.MoveNext
    Next
    
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    Con.Close
    Set Con = Nothing
    
    Me.Show
    
    DoEvents
    
    '*** Re-enable buttons
    cmdFind.Enabled = True
    cmdClear.Enabled = True
    cmdPackage.Enabled = False
    cmdHelp.Enabled = True
    bolSearchFormLoadComplete = True
    Me.Enabled = True
    
    
Exit Sub
    
ERROR_TRAP:
    ' 94 = Invalid use of Null
    If Err.Number = 94 Then
        Err.Clear
        Resume Next
    End If
    Result = MsgBox("LoadFieldFormats - Error: " & Err.Number & " - " & Err.Description, vbOK)
    Err.Clear
End Sub
