VERSION 5.00
Begin VB.Form frmImaging101ExportOptions 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Export Options"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   8235
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   8235
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFullPathForPdfReaderFiles 
      Height          =   405
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   26
      Top             =   5040
      Width           =   7575
   End
   Begin VB.CommandButton cmdBrowseForPdfReaderDirectory 
      BackColor       =   &H00FFFFFF&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5040
      Width           =   375
   End
   Begin VB.CheckBox chkIncludePdfReader 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Include PDF Reader?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   24
      Top             =   3120
      Width           =   2535
   End
   Begin VB.DirListBox dirHtmlDirectoryList 
      Height          =   288
      Left            =   4560
      TabIndex        =   23
      Top             =   1800
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.ComboBox cmbHTMLsourceDir 
      Height          =   288
      Left            =   480
      TabIndex        =   21
      Top             =   1560
      Width           =   7575
   End
   Begin VB.CommandButton cmdBrowseForPDFDirectory 
      BackColor       =   &H00FFFFFF&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4320
      Width           =   375
   End
   Begin VB.CommandButton cmdBrowseForHTMLDirectory 
      BackColor       =   &H00FFFFFF&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2280
      Width           =   375
   End
   Begin VB.CheckBox chkExportPDFtoCD 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Export PDF to CD?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   18
      Top             =   3120
      Width           =   2055
   End
   Begin VB.CheckBox chkExportHTMLtoCD 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Export HTML to CD?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   15
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox txtFullPathForHTMLexport 
      Height          =   405
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   2280
      Width           =   7575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   8175
      Begin VB.PictureBox picImaging101Logo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   6600
         Picture         =   "frmImaging101ExportOptions.frx":0000
         ScaleHeight     =   375
         ScaleWidth      =   1455
         TabIndex        =   28
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   840
         Picture         =   "frmImaging101ExportOptions.frx":0693
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   0
         Picture         =   "frmImaging101ExportOptions.frx":0AD5
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton cmdExportSelected 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Export"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1680
         Picture         =   "frmImaging101ExportOptions.frx":105F
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Export Selected Documents"
         Top             =   0
         Width           =   855
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C07000&
         Height          =   225
         Left            =   6600
         TabIndex        =   10
         Top             =   480
         Width           =   1245
      End
   End
   Begin VB.OptionButton optCombineIntoSinglePDF 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Combine all Documents into a &SINGLE PDF Document"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   5
      Top             =   3720
      Width           =   5055
   End
   Begin VB.OptionButton optBreakPDFbyDocgroup 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Break up PDF by &DocGroup"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   4
      Top             =   3480
      Value           =   -1  'True
      Width           =   2895
   End
   Begin VB.CheckBox chkExportToPDF 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Export to &PDF?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CheckBox chkExportToHTML 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Export to &HTML?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox txtFullPathForPDFexport 
      Height          =   405
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   4320
      Width           =   7575
   End
   Begin VB.TextBox txtCDBurnDriveforHTML 
      Height          =   285
      Left            =   7200
      TabIndex        =   13
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtCDBurnDriveforPDF 
      Height          =   285
      Left            =   7080
      TabIndex        =   16
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label lblFullPathForPdfReaderFiles 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Full ROOT Directory Path for PDF Reader Source Files"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   27
      Top             =   4800
      Width           =   4815
   End
   Begin VB.Label lblHTMLsourceDir 
      BackColor       =   &H00FFFFFF&
      Caption         =   "HTML Source Directory"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   22
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label lblCDBurnDriveForPDF 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CD Burning Drive for PDF"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   17
      Top             =   4080
      Width           =   2535
   End
   Begin VB.Label lblCDBurnDriveForHTML 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CD Burning Drive"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      TabIndex        =   14
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label lblFullPathForHTMLexport 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Full Directory Path for HTML Export"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   12
      Top             =   2040
      Width           =   3135
   End
   Begin VB.Label lblFullPathForPDFexport 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Full Directory Path for PDF Export"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   4080
      Width           =   3375
   End
End
Attribute VB_Name = "frmImaging101ExportOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub chkExportHTMLtoCD_Click()
    
    If chkExportHTMLtoCD.Value = vbChecked Then
        lblCDBurnDriveForHTML.Visible = True
        txtCDBurnDriveforHTML.Visible = True
        createCDBurner
    Else
        lblCDBurnDriveForHTML.Visible = False
        txtCDBurnDriveforHTML.Visible = False
    End If

End Sub

Private Sub chkExportPDFtoCD_Click()

    If chkExportPDFtoCD.Value = vbChecked Then
        lblCDBurnDriveForPDF.Visible = True
        txtCDBurnDriveforPDF.Visible = True
        createCDBurner
    Else
        lblCDBurnDriveForPDF.Visible = False
        txtCDBurnDriveforPDF.Visible = False
    End If

End Sub

Private Sub chkExportToHTML_Click()
    
    If chkExportToHTML.Value = vbChecked Then
    
        chkExportHTMLtoCD.enabled = True
        lblCDBurnDriveForHTML.Visible = True
        chkExportHTMLtoCD.enabled = True
        lblFullPathForHTMLexport.enabled = True
        txtFullPathForHTMLexport.enabled = True
        cmbHTMLsourceDir.enabled = True
        cmdBrowseForHTMLDirectory.enabled = True
        cmbHTMLsourceDir.enabled = True
        lblHTMLsourceDir.enabled = True
        If cmbHTMLsourceDir.Text = "" Then
            'Set to the first
        End If
        
    Else
    
        chkExportHTMLtoCD.enabled = False
        lblCDBurnDriveForHTML.Visible = False
        chkExportHTMLtoCD.enabled = False
        lblFullPathForHTMLexport.enabled = False
        txtFullPathForHTMLexport.enabled = False
        chkExportHTMLtoCD.Value = Unchecked
        cmdBrowseForHTMLDirectory.enabled = False
        cmbHTMLsourceDir.enabled = False
        lblHTMLsourceDir.enabled = False
'        cmbEntity.enabled = False

    End If
    
    chkExportHTMLtoCD_Click
    
End Sub

Private Sub chkExportToPDF_Click()

    If chkExportToPDF.Value = vbChecked Then
        chkExportPDFtoCD.enabled = True
        lblCDBurnDriveForPDF.Visible = True
        lblCDBurnDriveForPDF.enabled = False
        
        chkExportPDFtoCD.enabled = True
        optBreakPDFbyDocgroup.enabled = True
        optCombineIntoSinglePDF.enabled = True
        lblFullPathForPDFexport.enabled = True
        txtFullPathForPDFexport.enabled = True
        
    Else
    
        chkExportPDFtoCD.enabled = False
        lblCDBurnDriveForPDF.Visible = False
        lblCDBurnDriveForPDF.enabled = False
        optBreakPDFbyDocgroup.enabled = False
        optCombineIntoSinglePDF.enabled = False
        lblFullPathForPDFexport.enabled = False
        txtFullPathForPDFexport.enabled = False
        chkExportPDFtoCD.Value = Unchecked
    
    End If

    chkExportPDFtoCD_Click
    
End Sub

Private Sub chkIncludePdfReader_Click()

    If chkIncludePdfReader.Value = vbChecked Then
        lblFullPathForPdfReaderFiles.Visible = True
        txtFullPathForPdfReaderFiles.Visible = True
        cmdBrowseForPdfReaderDirectory.Visible = True
    Else
        lblFullPathForPdfReaderFiles.Visible = False
        txtFullPathForPdfReaderFiles.Visible = False
        cmdBrowseForPdfReaderDirectory.Visible = False
    End If
    
End Sub

Private Sub cmdBrowseForHTMLDirectory_Click()
    
    txtFullPathForHTMLexport.Text = funcBrowseForDirectory

End Sub

Private Sub cmdBrowseForPDFDirectory_Click()

    txtFullPathForPDFexport.Text = funcBrowseForDirectory

End Sub

Private Sub cmdBrowseForPdfReaderDirectory_Click()

    txtFullPathForPdfReaderFiles.Text = funcBrowseForDirectory

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdExportSelected_Click()

    'Check if No options selected or operation canceled
    If (frmImaging101ExportOptions.chkExportToHTML.Value Or frmImaging101ExportOptions.chkExportToPDF.Value) <> vbChecked Then
        MsgBox "OK... Seriously now... Do you really expect me to export with NO Export Options Selected?", vbQuestion, "No Export Options"
        Exit Sub
    End If

    '*** DO NOT UNLOAD this Form HERE... we need this form for the Options & Parameters.
    '    We will unload it when the frmImaging101Retrieve form is unloaded.
'    Me.Hide
    frmImaging101Retrieve.subExportSelected_Run
    
End Sub




Private Sub cmdSave_Click()

    funcGetSetUserSettings "SET", "ExportToHTML", chkExportToHTML.Value
    funcGetSetUserSettings "SET", "HTMLsourceDir", cmbHTMLsourceDir.Text
    funcGetSetUserSettings "SET", "ExportHTMLtoCD", chkExportHTMLtoCD.Value
    funcGetSetUserSettings "SET", "FullPathForHTMLexport", txtFullPathForHTMLexport.Text
    
    funcGetSetUserSettings "SET", "ExportToPDF", chkExportToPDF.Value
    funcGetSetUserSettings "SET", "IncludePdfReader", chkIncludePdfReader.Value
    funcGetSetUserSettings "SET", "ExportPDFtoCD", chkExportPDFtoCD.Value
    funcGetSetUserSettings "SET", "FullPathForPDFexport", txtFullPathForPDFexport.Text
    funcGetSetUserSettings "SET", "FullPathForPdfReaderFiles", txtFullPathForPdfReaderFiles.Text

    funcGetSetUserSettings "SET", "BreakPDFbyDocgroup", optBreakPDFbyDocgroup.Value
    funcGetSetUserSettings "SET", "CombineIntoSinglePDF", optCombineIntoSinglePDF.Value


End Sub

Private Sub Form_Load()


    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision

    ' Get Position settings from the registry
    On Error Resume Next
    Me.Top = VBGetPrivateProfileString(RegAppname, "frmImaging101ExportOptions.Top", RegFileName)
    Me.Left = VBGetPrivateProfileString(RegAppname, "frmImaging101ExportOptions.Left", RegFileName)


    '*********************************************************
    'Load the Combo Box with a list of Subdirectories
    
    dirHtmlDirectoryList.Path = funcGetFieldFromDB(RegImaging101ConnectionString, "I101Applications", "ApplicationName = '" & frmImaging101Search.txtApplicationName & "'", "RootDirectoryPathForHtmlSource") & ""
    
    
    For i = 0 To dirHtmlDirectoryList.ListCount - 1
        cmbHTMLsourceDir.AddItem dirHtmlDirectoryList.List(i)
    Next
    
    
    '*******************************************************
    'Load Saved Export Options for the Current User
    
    chkExportToHTML.Value = funcGetSetUserSettings("GET", "ExportToHTML", "")
    cmbHTMLsourceDir.Text = funcGetSetUserSettings("GET", "HTMLsourceDir", "")
    chkExportHTMLtoCD.Value = funcGetSetUserSettings("GET", "ExportHTMLtoCD", "")
    txtFullPathForHTMLexport.Text = funcGetSetUserSettings("GET", "FullPathForHTMLexport", "")
    
    chkExportToPDF.Value = funcGetSetUserSettings("GET", "ExportToPDF", "")
    chkIncludePdfReader.Value = funcGetSetUserSettings("GET", "IncludePdfReader", "")
    chkExportPDFtoCD.Value = funcGetSetUserSettings("GET", "ExportPDFtoCD", "")
    txtFullPathForPDFexport.Text = funcGetSetUserSettings("GET", "FullPathForPDFexport", "")
    txtFullPathForPdfReaderFiles.Text = funcGetSetUserSettings("GET", "FullPathForPdfReaderFiles", "")
    
    optBreakPDFbyDocgroup.Value = funcGetSetUserSettings("GET", "BreakPDFbyDocgroup", "")
    optCombineIntoSinglePDF.Value = funcGetSetUserSettings("GET", "CombineIntoSinglePDF", "")
    


    
    '****************************************
    'Initialize Export controls & fields
    Call chkExportToHTML_Click
    Call chkExportToPDF_Click
    Call chkIncludePdfReader_Click
    
'    txtDirectoryPathForPDF.Text = App.Path & "\PDF"

End Sub

Private Sub createCDBurner()
   
   Set m_cSimpleCDBurner = New cSimpleCDBurner
   On Error GoTo ErrorHandler
   m_cSimpleCDBurner.Initialise Me.hwnd
   
   If (m_cSimpleCDBurner.HasRecordableDrive) Then
      
        'See if Export to HTML to CD is checked
        If chkExportHTMLtoCD.Value = vbChecked Then
            txtCDBurnDriveforHTML.Text = m_cSimpleCDBurner.RecorderDriveLetter
            txtFullPathForHTMLexport.Text = m_cSimpleCDBurner.BurnStagingAreaFolder
        End If

        'See if Export to PDF to CD is checked
        If chkExportPDFtoCD.Value = vbChecked Then
            txtCDBurnDriveforPDF.Text = m_cSimpleCDBurner.RecorderDriveLetter
            txtFullPathForPDFexport.Text = m_cSimpleCDBurner.BurnStagingAreaFolder
        End If

'      showStagingFiles
      
'      enableControl txtStagingArea, True
'      enableControl txtDrive, True
'      enableControl tvwFiles, True
'      enableControl cmdAdd, True
'      enableControl cmdRemove, True
'      enableControl cmdRefresh, True
'      enableControl cmdBurn, True
      
   Else
        MsgBox "No Recordable Drive found.", vbInformation, "No Recordable Drive"
        txtCDBurnDriveforHTML.Text = "N/A"
        txtCDBurnDriveforPDF.Text = "N/A"
        
        chkExportHTMLtoCD.Value = vbUnchecked
        chkExportHTMLtoCD.enabled = False
        
        chkExportPDFtoCD.Value = vbUnchecked
        chkExportPDFtoCD.enabled = False
   End If

Exit Sub

ErrorHandler:
'   txtDrive.Text = "CD Burner Interface not initialised"
'   txtStagingArea.Text = "N/A"
   MsgBox "Failed to initialise the CD Burner!  This could mean either you DON'T have a CD Burner/Writer or it is not working properly.", vbExclamation
   Exit Sub

End Sub

Private Sub Form_Resize()

  On Error Resume Next
  
    Frame1.width = Me.width
    picImaging101Logo.Left = Me.ScaleWidth - picImaging101Logo.width - 10
    lblVersion.Left = picImaging101Logo.Left

End Sub

Private Sub Form_Unload(Cancel As Integer)

    'Save Form settings to the registry, except if minimized (i.e.- Top or Left < 0)
    If Me.Top >= 0 And Me.Left >= 0 Then
        result = WritePrivateProfileString(RegAppname, "frmImaging101ExportOptions.Top", Me.Top, RegFileName)
        result = WritePrivateProfileString(RegAppname, "frmImaging101ExportOptions.Left", Me.Left, RegFileName)
    End If

End Sub

