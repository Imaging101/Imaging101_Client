VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmImaging101Modify 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Modify Index Properties Form - Imaging101"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8190
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   8190
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtIndexValuesOld 
      BackColor       =   &H00C0FFFF&
      Height          =   330
      Index           =   0
      Left            =   2040
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   28
      Text            =   "frmImaging101Modify.frx":0000
      Top             =   2280
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox txtIndexValues 
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   0
      Left            =   4920
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   23
      Text            =   "frmImaging101Modify.frx":0014
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox txtDocumentRECID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Left            =   6240
      TabIndex        =   21
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtFieldIsSticky 
      Height          =   285
      Index           =   0
      Left            =   3240
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1860
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtFieldLowValue 
      Height          =   285
      Index           =   0
      Left            =   2520
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1860
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtFieldHighValue 
      Height          =   285
      Index           =   0
      Left            =   2880
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1860
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtFieldDefaultValue 
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1860
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtFieldsRECID 
      Height          =   285
      Index           =   0
      Left            =   1800
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1860
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtBatchFieldsRECID 
      Height          =   285
      Index           =   0
      Left            =   1440
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1860
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtFieldType 
      Height          =   285
      Index           =   0
      Left            =   3600
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1860
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtFieldSize 
      Height          =   285
      Index           =   0
      Left            =   3960
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1860
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtFieldName 
      Height          =   285
      Index           =   0
      Left            =   4320
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1860
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtFieldIsRequiredForCommit 
      Height          =   285
      Index           =   0
      Left            =   1080
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1860
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtApplicationRECID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Left            =   4800
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtApplicationName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Left            =   1200
      TabIndex        =   4
      Top             =   840
      Width           =   3375
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
      TabIndex        =   1
      Top             =   0
      Width           =   8175
      Begin VB.CommandButton cmdHelp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Help"
         Enabled         =   0   'False
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
         Left            =   2520
         Picture         =   "frmImaging101Modify.frx":0023
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton cmdOpenSelected 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Open"
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
         Picture         =   "frmImaging101Modify.frx":08ED
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Clear"
         Enabled         =   0   'False
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
         Picture         =   "frmImaging101Modify.frx":1057
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   0
         Width           =   855
      End
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
         Picture         =   "frmImaging101Modify.frx":1499
         ScaleHeight     =   375
         ScaleWidth      =   1455
         TabIndex        =   25
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Update"
         Default         =   -1  'True
         Enabled         =   0   'False
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
         Picture         =   "frmImaging101Modify.frx":1B2C
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Left            =   6480
         TabIndex        =   26
         Top             =   480
         Width           =   1365
      End
   End
   Begin VB.CommandButton cmdFieldDropDown 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   7680
      Picture         =   "frmImaging101Modify.frx":1F6E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSMask.MaskEdBox mebIndexValuesOld 
      Height          =   330
      HelpContextID   =   1
      Index           =   0
      Left            =   2040
      TabIndex        =   18
      Top             =   1560
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      PromptInclude   =   0   'False
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mebIndexValues 
      Height          =   330
      HelpContextID   =   1
      Index           =   0
      Left            =   4920
      TabIndex        =   24
      Top             =   1560
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   16777215
      PromptInclude   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "New Value"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4920
      TabIndex        =   20
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Curent Value"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   2040
      TabIndex        =   19
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lblSelectApplication 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Application"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblFieldDescription 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   1560
      Width           =   1815
   End
End
Attribute VB_Name = "frmImaging101Modify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

    '****************************
    '*** Declarations
    Dim con As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim ssql As String
    Dim cmd As ADODB.Command
    Dim bolErrorOccured As Boolean
    Dim bolFormLoadComplete As Boolean
    Dim intHoldFocusIndex As Integer
    Dim bolWaitForUpdate As Boolean
    
    
    

Private Sub cmdClear_Click()

    Dim strHoldFieldMask As String
    
    For intIndex = 0 To lblFieldDescription.Count - 1
        strHoldFieldMask = mebIndexValues(intIndex).Mask
        mebIndexValues(intIndex).Mask = ""
        mebIndexValues(intIndex).Text = ""
        mebIndexValues(intIndex).Mask = strHoldFieldMask
    Next
    'Set focus on the first input field
'    mebIndexValues(0).SetFocus

End Sub

Private Sub cmdFieldDropDown_Click(Index As Integer)

    frmDropDownList.Caption = txtFieldName(Index)
    frmDropDownList.funcPopulateDropDown txtApplicationName, txtFieldName(Index), mebIndexValues(Index)
    
    txtIndexValues(Index) = mebIndexValues(Index)
    
    'Show the list in as Modal AFTER populating it.  Otherwise it stops processing and won't populate.
    frmDropDownList.Show vbModal, Me
    

End Sub


Private Sub cmdOpenSelected_Click()

            frmImaging101Retrieve.ListView1_DblClick
            
            Me.SetFocus

End Sub

Private Sub cmdUpdate_Click()
    
    '*** VALIDATE THE FIELD - Mainly for valid Date!
    '    *** THERE SEEMS TO BE A BUG IN VB
    '        If you TAB out of a field, it will execute the Validate sub
    '        If you Press [ENTER]... the Default Key takes precedence and
    '            BYPASSES the Validate sub.
    mebIndexValues_Validate intHoldFocusIndex, False
    
    If bolErrorOccured Then
        Exit Sub
    End If
    
    Call subSaveDocumentPageValues
    
    bolWaitForUpdate = False
    
End Sub

Private Sub Form_Activate()
    
'    If txtFieldType(0) = "LongText" Then
'        txtIndexValues(0).SetFocus
'    Else
'        mebIndexValues(0).SetFocus
'    End If
    cmdOpenSelected.SetFocus
    
End Sub

Private Sub Form_Load()
    
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
    

    bolFormLoadComplete = False
    
    '*** Disable buttons to prevent users from Clicking on them
    '    prior to the form being ready
    cmdUpdate.enabled = False
    cmdClear.enabled = False
    cmdClear.Visible = False
    cmdHelp.enabled = False
    
    ' Get Position settings from the registry
    On Error Resume Next
    Me.Top = VBGetPrivateProfileString(RegAppname, "frmImaging101Modify.Top", RegFileName)
    Me.Left = VBGetPrivateProfileString(RegAppname, "frmImaging101Modify.Left", RegFileName)
    Me.width = VBGetPrivateProfileString(RegAppname, "frmImaging101Modify.Width", RegFileName)
    Me.Height = VBGetPrivateProfileString(RegAppname, "frmImaging101Modify.Height", RegFileName)
    On Error GoTo 0
    

'''''''    ' Get Database Connections settings from the registry
'''''''    On Error Resume Next
'''''''    RegImaging101ConnectionType = VBGetPrivateProfileString("DATABASE", "frmImaging101Retrieve.Adodc1.ConnectionType", RegFileName)
'''''''    RegImaging101ConnectionString = VBGetPrivateProfileString("DATABASE", "frmImaging101Retrieve.Adodc1.ConnectionString." & RegImaging101ConnectionType, RegFileName)
'''''''    On Error GoTo 0
    
    On Error GoTo FORM_LOAD_ERROR
    
    '***************************************
    '*** GET APPLICATION INFO
    txtApplicationName = frmImaging101Search.cmbApplicationList.Text
    txtApplicationRECID = frmImaging101Search.txtApplicationRECID.Text

'    Dim lstIndex As Integer
'    lstIndex = frmImaging101Retrieve.ListView1.SelectedItem.Index
'    txtDocumentRECID = frmImaging101Retrieve.ListView1.ListItems(lstIndex).Text
    

    '***************************************
    '*** Now load the new field definitiona
    Call subLoadFieldDefinitions
'    Call subGetFieldValues
    
''    Me.Visible = True
'    Me.Show vbModal
'
'    mebIndexValues(Index).SetFocus
    
    '*** Re-enable buttons
    cmdUpdate.enabled = True
    cmdClear.enabled = False
    cmdClear.Visible = False
    cmdHelp.enabled = True
    bolFormLoadComplete = True

    
Exit Sub



FORM_LOAD_ERROR:
    result = MsgBox("FORM_LOAD_ERROR: " & Err.Number & " - " & Err.Description, vbOKCancel)
    Err.Clear
    If result = vbOK Then
        'Try again
        Resume
    Else
        Unload Me
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Save Form settings to the registry, except if minimized (i.e.- Top or Left < 0)
    If Me.Top >= 0 And Me.Left >= 0 Then
        result = WritePrivateProfileString(RegAppname, "frmImaging101Modify.Top", Me.Top, RegFileName)
        result = WritePrivateProfileString(RegAppname, "frmImaging101Modify.Left", Me.Left, RegFileName)
        result = WritePrivateProfileString(RegAppname, "frmImaging101Modify.Width", Me.width, RegFileName)
        result = WritePrivateProfileString(RegAppname, "frmImaging101Modify.Height", Me.Height, RegFileName)
'''        Result = WritePrivateProfileString(RegAppname, "frmImaging101BatchList.Caption", Me.Caption, RegFileName)
    End If


End Sub


Sub subLoadFieldDefinitions()


    '*** THIS SUBROUTINE LOADS ALL THE APPLICATION FIELD DEFINITION INFORMATION
    '***  INCLUDING FIELD FORMAT VALUES INTO AN ARRAY.
    
    Set con = New ADODB.Connection
    con.Open RegImaging101ConnectionString
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    
    
    rs.Source = "Select * from I101Fields WHERE ApplicationRECID = " & txtApplicationRECID & " ORDER BY FieldOrderBatch"
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LOCKTYPE = adLockReadOnly
    
   On Error GoTo ERROR_TRAP
    
    'sql statement to select items on the drop down list
''    ssql = "Select * from I101Fields where ApplicationRECID = " & txtApplicationRECID
''    rs.Open ssql, Con
    con.Errors.Clear
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
        Unload cmdFieldDropDown(intIndex)
        
        Unload mebIndexValuesOld(intIndex)
        Unload txtIndexValuesOld(intIndex)
        
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
        
            Dim intFieldTop As Integer
            
        '* Create Field Objects - BEGIN
            If intFieldIndex > 0 Then
                Load lblFieldDescription(intFieldIndex)
                Load mebIndexValuesOld(intFieldIndex)
                Load mebIndexValues(intFieldIndex)
                Load txtIndexValuesOld(intFieldIndex)
                Load txtIndexValues(intFieldIndex)
                Load txtBatchFieldsRECID(intFieldIndex)
                Load txtFieldsRECID(intFieldIndex)
                Load txtFieldDefaultValue(intFieldIndex)
                Load txtFieldLowValue(intFieldIndex)
                Load txtFieldHighValue(intFieldIndex)
                Load txtFieldIsSticky(intFieldIndex)
                Load txtFieldType(intFieldIndex)
                Load txtFieldSize(intFieldIndex)
                Load txtFieldName(intFieldIndex)
                Load txtFieldIsRequiredForCommit(intFieldIndex)
'                Load txtFieldSearchCondition(intFieldIndex)
'                Load cboFieldSearchCondition(intFieldIndex)
                Load cmdFieldDropDown(intFieldIndex)
                'Set the top to slightly below the previous field
                intFieldTop = lblFieldDescription(intFieldIndex - 1).Top + lblFieldDescription(intFieldIndex - 1).Height + intFieldSpacing
            Else
                'Set top to where the first field is
                intFieldTop = lblFieldDescription(intFieldIndex).Top
            End If
        '* Create Field Objects - END
        

'                Set lblFieldDescription(intFieldIndex).Container = Frame2
                lblFieldDescription(intFieldIndex).Top = intFieldTop
                lblFieldDescription(intFieldIndex).enabled = True
                lblFieldDescription(intFieldIndex).Visible = True
                lblFieldDescription(intFieldIndex).Caption = ""
                
                
'                Set mebIndexValuesOld(intFieldIndex).Container = Frame2
                mebIndexValuesOld(intFieldIndex).Top = intFieldTop
                mebIndexValuesOld(intFieldIndex).enabled = False
                mebIndexValuesOld(intFieldIndex).Visible = False
                mebIndexValuesOld(intFieldIndex).TabStop = False
                mebIndexValuesOld(intFieldIndex).Text = ""
                
'                Set mebIndexValues(intFieldIndex).Container = Frame2
                mebIndexValues(intFieldIndex).Top = intFieldTop
                mebIndexValues(intFieldIndex).enabled = True
                mebIndexValues(intFieldIndex).Visible = False
                mebIndexValues(intFieldIndex).TabIndex = intIndex + 1
                mebIndexValues(intFieldIndex).Text = ""
                
                
'                Set txtIndexValuesOld(intFieldIndex).Container = Frame2
                txtIndexValuesOld(intFieldIndex).Top = intFieldTop
                txtIndexValuesOld(intFieldIndex).enabled = False
                txtIndexValuesOld(intFieldIndex).Visible = False
                txtIndexValuesOld(intFieldIndex).TabStop = False
                txtIndexValuesOld(intFieldIndex).Text = ""
                
'                Set txtIndexValues(intFieldIndex).Container = Frame2
                txtIndexValues(intFieldIndex).Top = intFieldTop
                txtIndexValues(intFieldIndex).enabled = True
                txtIndexValues(intFieldIndex).Visible = False
                txtIndexValues(intFieldIndex).TabIndex = intIndex + 1
                txtIndexValues(intFieldIndex).Text = ""

'                Set txtBatchFieldsRECID(intFieldIndex).Container = Frame2
                txtBatchFieldsRECID(intFieldIndex).enabled = True
                txtBatchFieldsRECID(intFieldIndex).Visible = False

'                Set txtFieldsRECID(intFieldIndex).Container = Frame2
                txtFieldsRECID(intFieldIndex).enabled = True
                txtFieldsRECID(intFieldIndex).Visible = False

'                Set txtFieldDefaultValue(intFieldIndex).Container = Frame2
                txtFieldDefaultValue(intFieldIndex).enabled = True
                txtFieldDefaultValue(intFieldIndex).Visible = False
                txtFieldDefaultValue(intFieldIndex).Text = ""
                
'                Set txtFieldLowValue(intFieldIndex).Container = Frame2
                txtFieldLowValue(intFieldIndex).enabled = True
                txtFieldLowValue(intFieldIndex).Visible = False
                txtFieldLowValue(intFieldIndex).Text = ""
            
'                Set txtFieldHighValue(intFieldIndex).Container = Frame2
                txtFieldHighValue(intFieldIndex).enabled = True
                txtFieldHighValue(intFieldIndex).Visible = False
                txtFieldHighValue(intFieldIndex).Text = ""
                
'                Set txtFieldIsSticky(intFieldIndex).Container = Frame2
                txtFieldIsSticky(intFieldIndex).enabled = True
                txtFieldIsSticky(intFieldIndex).Visible = False
                txtFieldIsSticky(intFieldIndex).Text = ""
            
'                Set txtFieldType(intFieldIndex).Container = Frame2
                txtFieldType(intFieldIndex).enabled = True
                txtFieldType(intFieldIndex).Visible = False
                txtFieldType(intFieldIndex).Text = ""
            
'                Set txtFieldSize(intFieldIndex).Container = Frame2
                txtFieldSize(intFieldIndex).enabled = True
                txtFieldSize(intFieldIndex).Visible = False
                txtFieldSize(intFieldIndex).Text = ""
                
'                Set txtFieldName(intFieldIndex).Container = Frame2
                txtFieldName(intFieldIndex).enabled = True
                txtFieldName(intFieldIndex).Visible = False
                txtFieldName(intFieldIndex).Text = ""
            
'                Set txtFieldIsRequiredForCommit(intFieldIndex).Container = Frame2
                txtFieldIsRequiredForCommit(intFieldIndex).enabled = True
                txtFieldIsRequiredForCommit(intFieldIndex).Visible = False
                txtFieldIsRequiredForCommit(intFieldIndex).Text = ""
            
        
                '*** Create the DROP-DOWN Button
'                Set FieldDropDownList(intFieldIndex).Container = Frame2
                cmdFieldDropDown(intFieldIndex).Top = intFieldTop
                cmdFieldDropDown(intFieldIndex).enabled = True
                intTabCounter = intTabCounter + 1
                cmdFieldDropDown(intFieldIndex).TabStop = False
                'Make the DropDownList button VISIBLE only if Checked for the current field
                If rs.Fields!FieldDropDownList = vbChecked Then
                    cmdFieldDropDown(intFieldIndex).Visible = True
                Else
                    cmdFieldDropDown(intFieldIndex).Visible = False
                End If
                

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
        
        '10/7/2012 - Jacob - Moved this down so txtFieldType test will work
        '*** Determine whether to use the Text or Masked Edit Control
        If txtFieldType(intFieldIndex) = "LongText" Then
            mebIndexValues(intFieldIndex).TabStop = False
            txtIndexValuesOld(intFieldIndex).Visible = True
            txtIndexValues(intFieldIndex).TabStop = True
            txtIndexValues(intFieldIndex).Visible = True
            txtIndexValues(intFieldIndex).TabIndex = intFieldIndex + 1
        Else
            txtIndexValues(intFieldIndex).TabStop = False
            mebIndexValuesOld(intFieldIndex).Visible = True
            mebIndexValues(intFieldIndex).TabStop = True
            mebIndexValues(intFieldIndex).Visible = True
            mebIndexValues(intFieldIndex).TabIndex = intFieldIndex + 1
        End If
            
        rs.MoveNext
    Next
    
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    con.Close
    Set con = Nothing
    
'    Me.Show
'    Me.Visible = True
    
    DoEvents
    
    
Exit Sub
    
ERROR_TRAP:
    ' 94 = Invalid use of Null
    If Err.Number = 94 Then
        Err.Clear
        Resume Next
    End If
    result = MsgBox("LoadFieldFormats - Error: " & Err.Number & " - " & Err.Description, vbOK)
    Err.Clear
End Sub


Private Sub Form_Resize()

    Frame1.width = Me.width
    picImaging101Logo.Left = Me.ScaleWidth - picImaging101Logo.width - 10

End Sub

Private Sub Form_Unload(Cancel As Integer)
 
   
   frmImaging101Search.Show
   frmImaging101Retrieve.Show
   
   frmImaging101Search.enabled = True
   frmImaging101Retrieve.enabled = True
   
   '7/22/2005 - Jacob - Disabled the Refresh of the Result Set after a Modify
'   frmImaging101Search.cmdFind_Click


 
End Sub

Private Sub mebIndexValues_GotFocus(Index As Integer)

    
    '* If Date is Blank, then Errors - Copy the Date to the Thru field
    If (txtFieldType(Index) = "Date") And (Trim(mebIndexValues(Index).Text) = "") Then
        If InStr(1, lblFieldDescription(Index), "(Thru)") > 0 Then
            '* By design, the Thru date field is always immediatelly after the From
            mebIndexValues(Index).Text = mebIndexValues(Index - 1).Text
        End If
    End If
    
    '*** Highlight the Field
    mebIndexValues.item(Index).SelStart = 0
'    mebIndexValues.item(Index).SelLength = Len(mebIndexValues.item(Index).Text)
    mebIndexValues.item(Index).SelLength = 99
    
    intHoldFocusIndex = Index
    
End Sub

Private Sub mebIndexValues_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If bolFormLoadComplete <> True Then
        Exit Sub
    End If
    
    
    'Catch Enter key
    If KeyAscii = 13 Then
        cmdUpdate_Click
    End If

    If KeyAscii = Asc("[") And frmImaging101Retrieve.Visible = True Then
        frmImaging101Modify.SetFocus
        'Cancel the Keypress by setting the KeyAscii to Zero (0)
        KeyAscii = 0
    End If
    
    If KeyAscii = Asc("]") And frmImaging101Retrieve.Visible = True Then
        frmImaging101Retrieve.SetFocus
        'Cancel the Keypress by setting the KeyAscii to Zero (0)
        KeyAscii = 0
    End If
    
    If KeyAscii = Asc("\") And MainMDIForm.Visible = True Then
        MainMDIForm.SetFocus
        'Cancel the Keypress by setting the KeyAscii to Zero (0)
        KeyAscii = 0
    End If

End Sub

Private Sub mebIndexValues_Validate(Index As Integer, Cancel As Boolean)
    
    bolErrorOccured = False
    
    On Error GoTo ERROR_TRAP
    
    If (txtFieldType(Index) = "Date") And (Trim(mebIndexValues(Index).Text) <> "") Then
        '* Remove the "Prompt" characters
        strDateFormatted = Replace(Trim(mebIndexValues(Index).FormattedText), "_", "")
        strDateFormatted = CDate(strDateFormatted)
    End If
            
Exit Sub

ERROR_TRAP:

    If Err.Number = 13 Then
        result = MsgBox("Field Format Error: " & Err.Number & " - " & Err.Description & vbCrLf & "PLEASE CHECK YOUR INPUT.", vbOK)
        Me.SetFocus
        mebIndexValues(Index).SetFocus
        bolErrorOccured = True
        Err.Clear
        'Prevent moving to the next field
        Cancel = True
        'Force the Field to Highlight
        mebIndexValues_GotFocus (Index)
'        Exit Sub
    End If


End Sub

Private Sub subGetFieldValues()

    Dim intFieldIndex As Integer

    On Error Resume Next
    
    
    ' Set Connection Properties
    Set con = New ADODB.Connection
    con.Open RegImaging101ConnectionString
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    
    
    con.Errors.Clear
    
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LOCKTYPE = adLockReadOnly
    
    rs.Source = "Select * from " & txtApplicationName & " WHERE DocumentRECID = " & txtDocumentRECID
    rs.Open
    
    rs.MoveFirst
    
    '*** Display Field Values
    For intIndex = 0 To mebIndexValues.Count - 1
    
        If txtFieldType(intIndex).Text = "LongText" Then
            'Use the TEXT Control
            
            If (Not IsNull(rs.Fields("" + txtFieldName(intIndex) + "")) And (rs.Fields("" + txtFieldName(intIndex) + "") <> "")) Then
                ' If the field is not empty, Set field value
                txtIndexValuesOld(intIndex).Text = rs.Fields("" + txtFieldName(intIndex) + "") & ""
                txtIndexValues(intIndex).Text = rs.Fields("" + txtFieldName(intIndex) + "") & ""
            End If
            
        Else
        
            'Make modified values "sticky"
            
                If Trim(mebIndexValues(intIndex).Format) = "" Then
                   mebIndexValuesOld(intIndex).Text = rs.Fields("" + txtFieldName(intIndex) + "") & ""
                    If mebIndexValuesOld(intIndex).Text <> "" Then
                        mebIndexValues(intIndex).Text = rs.Fields("" + txtFieldName(intIndex) + "") & ""
                    End If
                Else
                    mebIndexValuesOld(intIndex).Text = Format(rs.Fields("" + txtFieldName(intIndex) + ""), mebIndexValues(intIndex).Format) & ""
                    If mebIndexValuesOld(intIndex).Text <> "" Then
                        mebIndexValues(intIndex).Text = Format(rs.Fields("" + txtFieldName(intIndex) + ""), mebIndexValues(intIndex).Format) & ""
                    End If
                End If
            
        End If
        
    Next
    
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    con.Close
    Set con = Nothing
    

      
End Sub

Private Sub subSaveDocumentPageValues()
    
    Dim intFieldIndex As Integer

    On Error Resume Next
    
    ' Set Connection Properties
    Set con = New ADODB.Connection
    con.Open RegImaging101ConnectionString
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    
    con.Errors.Clear
    
    '*  Begin Transaction
    con.BeginTrans
    
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenForwardOnly
    rs.LOCKTYPE = adLockOptimistic
    
    rs.Source = "Select * from " & txtApplicationName & " WHERE DocumentRECID = " & txtDocumentRECID
    rs.Open
    
    
    rs.MoveFirst
    
    '*** Cycle through Field Values
    For intIndex = 0 To mebIndexValues.Count - 1
            
            ' Get Field Index by comparing the Form fieldname with the DB result set fieldname
            For intFieldIndex = 0 To rs.Fields.Count - 1
                If rs.Fields(intFieldIndex).name = txtFieldName(intIndex) Then
                    Exit For
                End If
            Next
            
            
'            'Save the field value ONLY if it has been Changed!
'            If Trim(mebIndexValues(intIndex).FormattedText) <> Trim(mebIndexValuesOld(intIndex).FormattedText) Then
'                ' If the field is empty, Set to Null value
'                If Trim(mebIndexValues(intIndex)) = "" Then
'                    rs.Fields(intFieldIndex) = Null
'                ' If the field is Date, Format as Date value
'                ElseIf txtFieldType(intIndex) = "Date" Then
'                    strDateFormatted = Replace(Trim(mebIndexValues(intIndex).FormattedText), "_", "")
'                    strDateFormatted = CDate(strDateFormatted)
'                    rs.Fields(intFieldIndex) = strDateFormatted
'                ' If the field is Numeric, convert to Long
'                ElseIf txtFieldType(intIndex) = "Numeric" Then
'                    rs.Fields(intFieldIndex) = CDbl(Format(mebIndexValues(intIndex).FormattedText, mebIndexValues(intIndex).Format))
'                ' If the field is Currency, convert to Currency
'                ElseIf txtFieldType(intIndex) = "Currency" Then
'                    rs.Fields(intFieldIndex) = CCur(Format(mebIndexValues(intIndex).FormattedText, mebIndexValues(intIndex).Format))
'                Else
'                    'Use the FieldMask value to format the data
'                    '  also Trim it save only up to the defined length of the field
'                    If Trim(mebIndexValues(intIndex).Format) = "" Then
'                        'NO Format -- Simply Trim and Clip the Field
'                        rs.Fields(intFieldIndex) = Left(Trim(mebIndexValues(intIndex).Text), txtFieldSize(intIndex))
'                    Else
'                        rs.Fields(intFieldIndex) = Left(Trim(Format(mebIndexValues(intIndex).Text, mebIndexValues(intIndex).Format)), txtFieldSize(intIndex))
'                    End If
'                End If
'
'            End If
            
'            '* If field flagged as questionable, flag as red
'            If mebIndexValues(intIndex).Text = txtQuestionable Then
'                mebIndexValues(intIndex).Font.Bold = True
'                mebIndexValues(intIndex).ForeColor = vbRed
'            Else
'                mebIndexValues(intIndex).Font.Bold = False
'                mebIndexValues(intIndex).ForeColor = vbNormal
'            End If
        
            If (Not IsNull(mebIndexValues(intIndex))) _
                    Or (mebIndexValues(intIndex) <> "") _
                    Or (Not IsNull(txtIndexValues(intIndex))) _
                    Or (txtIndexValues(intIndex) <> "") _
                    Then
                
                
                '*** Check if the txtFieldType is set to LongText...
                '    this will handle saving the value of the the TextBox control
                '    instead of the mebIndexValues Masked Edit control as needed.
                '    This is because the Masked Edit control has a MAX size of 64 Char.
                If txtFieldType(intIndex) = "LongText" Then
                    'Save the TEXT Control Value
                    rs.Fields(intFieldIndex) = Trim(txtIndexValues(intIndex).Text)
                
                Else
                    'Save the MASKED EDIT Control Value
                
                    ' If the field is empty, Set to Null value
                    If Trim(mebIndexValues(intIndex)) = "" Then
                        rs.Fields(intFieldIndex) = Null
                    ' If the field is Date, Format as Date value
                    ElseIf txtFieldType(intIndex) = "Date" Then
                        rs.Fields(intFieldIndex) = Format(mebIndexValues(intIndex).FormattedText, mebIndexValues(intIndex).Format)
                        If Err.Number <> 0 Then
                            mebIndexValues(intIndex).SetFocus
                            mebIndexValues(intIndex).ForeColor = vbRed
                        End If
                    ' If the field is Numeric, convert to Long
                    ElseIf txtFieldType(intIndex) = "Numeric" Then
                        rs.Fields(intFieldIndex) = CDbl(Format(mebIndexValues(intIndex).FormattedText, mebIndexValues(intIndex).Format))
                    ' If the field is Currency, convert to Currency
                    ElseIf txtFieldType(intIndex) = "Currency" Then
                        rs.Fields(intFieldIndex) = CCur(Format(mebIndexValues(intIndex).FormattedText, mebIndexValues(intIndex).Format))
                    Else
                        'Use the FieldMask value to format the data
                        '  also Trim it save only up to the defined length of the field
                        If Trim(mebIndexValues(intIndex).Format) = "" Then
                            'NO Format -- Simply Trim and Clip the Field
                            rs.Fields(intFieldIndex) = Left(Trim(mebIndexValues(intIndex).Text), txtFieldSize(intIndex))
                        Else
                            rs.Fields(intFieldIndex) = Left(Trim(Format(mebIndexValues(intIndex).Text, mebIndexValues(intIndex).Format)), txtFieldSize(intIndex))
                        End If
                    End If
                
                End If '  txtFieldType(intIndex) = "LongText"
                
            End If
            
    Next
    
    
    ' Update and Commit the Transactions
    rs.Update
    con.CommitTrans
    
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    con.Close
    Set con = Nothing


End Sub


Private Sub txtIndexValues_Change(Index As Integer)
    If Not bolIndexFormLoadComplete Then
        Exit Sub
    End If
    
    If Len(txtIndexValues(Index).Text) > Int(txtFieldSize(Index).Text) Then
        MsgBox "Exceeded field size of " & txtFieldSize(Index).Text & " characters!" & _
                vbCrLf & "Truncating to defined size.", vbOKOnly
        txtIndexValues(Index).Text = Left(txtIndexValues(Index).Text, Int(txtFieldSize(Index).Text))
    End If
    
End Sub

Private Sub txtIndexValues_DblClick(Index As Integer)
    
    'Re-size the Text field.
    If txtIndexValues(Index).Height = txtIndexValues(0).Height Then
        txtIndexValues(Index).BackColor = vbCyan
        txtIndexValues(Index).Refresh
        txtIndexValues(Index).Height = txtIndexValues(0).Height * 3
    Else
        txtIndexValues(Index).BackColor = vbWhite ' Light Red
        txtIndexValues(Index).Height = txtIndexValues(0).Height

    End If
    
End Sub


Public Sub subModifyRecords()

    For i = 1 To frmImaging101Retrieve.ListView1.ListItems.Count
    
        If frmImaging101Retrieve.ListView1.ListItems(i).Selected = True Then
        
            funcWriteToDebugLog Me.name, frmImaging101Retrieve.ListView1.ListItems(i).Text
            frmImaging101Retrieve.ListView1.ListItems(i).Selected = True   ' Force item selection
            
            ReleaseCapture
            
            '1/19/2011 - Jacob - Moved to cmdOpenSelected_Click()
'            frmImaging101Retrieve.ListView1_DblClick
            
            Me.SetFocus
            
            Dim lstIndex As Integer
            lstIndex = frmImaging101Retrieve.ListView1.SelectedItem.Index
            txtDocumentRECID = frmImaging101Retrieve.ListView1.ListItems(lstIndex).Text
            

            Call subGetFieldValues
            
            bolWaitForUpdate = True
            
            While bolWaitForUpdate = True
                DoEvents
            Wend
            
            
            
        End If
    Next

    Unload Me
    
End Sub

Private Sub txtIndexValues_GotFocus(Index As Integer)

        txtIndexValuesOld(Index).Refresh
        txtIndexValuesOld(Index).Height = txtIndexValues(Index).Height * 3
        
        txtIndexValues(Index).BackColor = vbYellow
        txtIndexValues(Index).Refresh
        txtIndexValues(Index).Height = txtIndexValues(Index).Height * 3
        
        Dim intFieldDepth As Integer
        intFieldDepth = txtIndexValues(Index).Top + txtIndexValues(Index).Height
        
        If intFieldDepth > Me.ScaleHeight Then
            Me.Height = Me.Height + 200
            DoEvents
        End If
        
End Sub

Private Sub txtIndexValues_LostFocus(Index As Integer)

        txtIndexValuesOld(Index).Height = mebIndexValues(Index).Height
        
        txtIndexValues(Index).BackColor = vbWhite
        txtIndexValues(Index).Height = mebIndexValues(Index).Height
        txtIndexValues(Index).SelStart = 0
        txtIndexValues(Index).SelLength = 0
        
End Sub
