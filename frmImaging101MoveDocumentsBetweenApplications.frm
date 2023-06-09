VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmImaging101MoveDocumentsBetweenApplications 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Move Documents Between Applications"
   ClientHeight    =   7575
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11535
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
   ScaleHeight     =   7575
   ScaleWidth      =   11535
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   615
      Left            =   0
      TabIndex        =   13
      Top             =   6960
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   1085
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.TextBox txtSourceFieldName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   2640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtDestinationFieldName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   6000
      TabIndex        =   14
      Top             =   2640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtSourceApplication 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   720
      Width           =   4335
   End
   Begin VB.CommandButton cmdMoveDocsBetweenApplications 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Move documents "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9480
      Picture         =   "frmImaging101MoveDocumentsBetweenApplications.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6720
      Width           =   2055
   End
   Begin VB.Timer TimerDestination 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5760
      Top             =   0
   End
   Begin VB.TextBox txtDestinationApplicationRECID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
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
      Left            =   5880
      TabIndex        =   9
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtDestinationApplicationName 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
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
      Left            =   7200
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtSourceApplicationRECID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
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
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtSourceApplicationName 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
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
      Left            =   1440
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
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
      Left            =   9840
      Picture         =   "frmImaging101MoveDocumentsBetweenApplications.frx":0442
      ScaleHeight     =   375
      ScaleWidth      =   1455
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin VB.ComboBox cmbSourceFieldNameForInput 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   2250
      Width           =   2895
   End
   Begin VB.ComboBox cmbDestinationApplicationList 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   5880
      TabIndex        =   1
      Top             =   720
      Width           =   4335
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      ForeColor       =   &H00C07000&
      Height          =   225
      Left            =   9840
      TabIndex        =   26
      Top             =   480
      Width           =   1245
   End
   Begin VB.Label lblWarning 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   720
      TabIndex        =   25
      Top             =   1320
      Width           =   7815
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Field Name for Input"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Field Name for Input"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   255
      Left            =   5880
      TabIndex        =   23
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label lblDestinationFieldSize 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   9960
      TabIndex        =   22
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label lblDestinationFieldType 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   8760
      TabIndex        =   21
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lblSourceFieldSize 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4320
      TabIndex        =   20
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label lblSourceFieldType 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   3120
      TabIndex        =   19
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "To this Destination Field"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5880
      TabIndex        =   18
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Copy values from this SOURCE Field"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
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
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label lblArrows 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   ">>>"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   495
      Index           =   0
      Left            =   5280
      TabIndex        =   16
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label lblDestinationFieldNameForInput 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   6000
      TabIndex        =   10
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Move Documents Between Applications"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select &DESTINATION Application"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   5880
      TabIndex        =   2
      Top             =   480
      Width           =   3492
   End
   Begin VB.Label lblApplication 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&SOURCE Application"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2892
   End
End
Attribute VB_Name = "frmImaging101MoveDocumentsBetweenApplications"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmbDestinationApplicationList_Click()

    If Trim(txtSourceApplicationRECID = "") Then
        MsgBox "Please select the SOURCE first!", vbInformation
        Exit Sub
    End If
    
    ' Get the Application to Commit Batches to
    Set con = New ADODB.Connection
    con.Open RegImaging101ConnectionString
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    
        
    rs.Source = "Select ApplicationRECID,ApplicationName, MaxItemsToRetrieve from I101Applications " & _
                " WHERE ApplicationName= '" & cmbDestinationApplicationList.Text & "'"
                
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LOCKTYPE = adLockReadOnly
    
    con.Errors.Clear
    
    rs.Open
    
    If Not (rs.EOF Or rs.BOF) Then
        txtDestinationApplicationRECID = rs!ApplicationRECID
        txtDestinationApplicationName = rs!ApplicationName

    End If
    
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    con.Close
    Set con = Nothing


    TimerDestination.enabled = True
    
    




End Sub

Private Sub cmbSourceFieldNameForInput_Click(Index As Integer)


    If Trim(cmbSourceFieldNameForInput(Index)) = "" Then
        txtSourceFieldName(Index) = ""
        lblSourceFieldType(Index) = ""
        lblSourceFieldSize(Index) = ""
    
        Exit Sub
    End If
    
    txtSourceFieldName(Index) = funcGetFieldFromDB(RegImaging101ConnectionString, "I101Fields", "ApplicationRECID = " & txtSourceApplicationRECID & " AND FieldNameForInput = '" & cmbSourceFieldNameForInput(Index) & "'", "FieldName")
    lblSourceFieldType(Index) = funcGetFieldFromDB(RegImaging101ConnectionString, "I101Fields", "ApplicationRECID = " & txtSourceApplicationRECID & " AND FieldNameForInput = '" & cmbSourceFieldNameForInput(Index) & "'", "FieldType")
    lblSourceFieldSize(Index) = funcGetFieldFromDB(RegImaging101ConnectionString, "I101Fields", "ApplicationRECID = " & txtSourceApplicationRECID & " AND FieldNameForInput = '" & cmbSourceFieldNameForInput(Index) & "'", "FieldSize")

    If lblSourceFieldSize(Index) = "" Then lblSourceFieldSize(Index) = "0"
    If lblDestinationFieldSize(Index) = "" Then lblDestinationFieldSize(Index) = "0"
    
    If UCase(lblSourceFieldType(Index)) <> "DATE" _
    And CInt(lblSourceFieldSize(Index)) > CInt(lblDestinationFieldSize(Index)) _
    Then
        cmbSourceFieldNameForInput(Index).ForeColor = vbRed
        lblSourceFieldType(Index).ForeColor = vbRed
        lblSourceFieldSize(Index).ForeColor = vbRed
        
        lblWarning.Caption = "WARNING:  Fields marked in RED may LOSE DATA !!!"
    Else
        cmbSourceFieldNameForInput(Index).ForeColor = vbBlack
        lblSourceFieldType(Index).ForeColor = vbBlack
        lblSourceFieldSize(Index).ForeColor = vbBlack
    End If
    
    
End Sub

Private Sub cmbSourceFieldNameForInput_DropDown(Index As Integer)

    funcFillList cmbSourceFieldNameForInput(Index), RegImaging101ConnectionString, "I101Fields", _
                "FieldNameForInput", "ApplicationRECID = " & txtSourceApplicationRECID & _
                " AND FieldType = '" & lblDestinationFieldType(Index) & "'", True, True

End Sub

Private Sub cmbSourceApplicationList_Click()


    
    ' Get the Application to Commit Batches to
    Set con = New ADODB.Connection
    con.Open RegImaging101ConnectionString
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    
        
    rs.Source = "Select ApplicationRECID,ApplicationName, MaxItemsToRetrieve from I101Applications WHERE ApplicationName= '" & cmbSourceApplicationList.Text & "'"
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LOCKTYPE = adLockReadOnly
    
    con.Errors.Clear
    
    rs.Open
    
    If Not (rs.EOF Or rs.BOF) Then
        txtSourceApplicationRECID = rs!ApplicationRECID
        txtSourceApplicationName = rs!ApplicationName

    End If
    
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    con.Close
    Set con = Nothing





End Sub

Private Sub cmbSourceField_DropDown(Index As Integer)

    funcFillList cmbSourceFieldNameForInput(Index), RegImaging101ConnectionString, "I101Fields", "FieldNameForInput", "ApplicationRECID = " & txtSourceApplicationRECID, True, True

End Sub

Private Sub cmdMoveDocsBetweenApplications_Click()

    'Go ahead and move the documents.
    frmImaging101Retrieve.subMoveDocsBetweenApplications

End Sub

Private Sub Form_Load()


    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision

    txtSourceApplication.Text = frmImaging101Search.txtApplicationName
    txtSourceApplicationRECID.Text = frmImaging101Search.txtApplicationRECID
    
    '***************************************
    '*** LOAD APPLICATION LIST DROP-DOWN
    
    Set con = New ADODB.Connection
    con.Open RegImaging101ConnectionString
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    
'*** Changed the Load to work with Security
'    rs.Source = "Select ApplicationName, ApplicationRECID from I101Applications ORDER BY ApplicationName"
    rs.Source = ""
    rs.Source = rs.Source & "Select * "
    rs.Source = rs.Source & " FROM I101Applications, I101SecurityApplications"
    rs.Source = rs.Source & " WHERE I101Applications.ApplicationRECID = I101SecurityApplications.ApplicationRECID And I101SecurityApplications.SecurityRECID = " & gsecSecurityRECID
    rs.Source = rs.Source & " AND I101Applications.ApplicationName <> '" & txtSourceApplication & "'"

    rs.Source = rs.Source & " ORDER BY ApplicationName"

    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LOCKTYPE = adLockReadOnly
    
    con.Errors.Clear
    
    rs.Open
    rs.MoveFirst
    
    For intIndex = 0 To rs.RecordCount - 1
    
        cmbDestinationApplicationList.AddItem rs.Fields!ApplicationName
        cmbDestinationApplicationList.ItemData(intIndex) = rs.Fields!ApplicationRECID
        
        rs.MoveNext
    Next
        
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    con.Close
    Set con = Nothing

    '****************************



End Sub



Private Sub Form_Unload(Cancel As Integer)

    frmImaging101Search.Show
    frmImaging101Retrieve.Show

End Sub

Private Sub TimerDestination_Timer()

    TimerDestination.enabled = False
    
    
    cmbSourceFieldNameForInput(0).Text = ""
    
    For i = 1 To cmbSourceFieldNameForInput.UBound
        Unload cmbSourceFieldNameForInput(i)
        Unload lblSourceFieldType(i)
        Unload lblSourceFieldSize(i)
        Unload txtSourceFieldName(i)
        
        Unload lblDestinationFieldNameForInput(i)
        Unload lblDestinationFieldType(i)
        Unload lblDestinationFieldSize(i)
        Unload txtDestinationFieldName(i)
        
        Unload lblArrows(i)
    Next

    '*** Declarations
'    Dim rs As ADODB.Recordset
'    Dim con As ADODB.Connection
    Dim ssql As String

    Set con = New ADODB.Connection
    Set rs = New ADODB.Recordset
    con.Open RegImaging101ConnectionString

    'sql statement to select items on the drop down list
    ssql = "SELECT *  FROM I101Fields WHERE ApplicationRECID = " & txtDestinationApplicationRECID
    rs.Open ssql, con

    Dim intIndex As Integer
    On Error Resume Next
    
    While Not rs.EOF
    
        Dim intFieldSpacing As Integer
        
        intFieldSpacing = 100
        
        If intIndex > 0 Then
            'Set Up SOURCE Fields
            Load cmbSourceFieldNameForInput(intIndex)
            cmbSourceFieldNameForInput(intIndex).Top = cmbSourceFieldNameForInput(intIndex - 1).Top + cmbSourceFieldNameForInput(intIndex - 1).Height + intFieldSpacing
            cmbSourceFieldNameForInput(intIndex).Visible = True
            
            Load lblSourceFieldType(intIndex)
            lblSourceFieldType(intIndex).Top = cmbSourceFieldNameForInput(intIndex).Top
            lblSourceFieldType(intIndex).Visible = True
            
            Load lblSourceFieldSize(intIndex)
            lblSourceFieldSize(intIndex).Top = cmbSourceFieldNameForInput(intIndex).Top
            lblSourceFieldSize(intIndex).Visible = True
            
            Load txtSourceFieldName(intIndex)
            
            'Set Up Arrows
            Load lblArrows(intIndex)
            lblArrows(intIndex).Top = cmbSourceFieldNameForInput(intIndex).Top
            lblArrows(intIndex).Visible = True
            
            'Set up Destination fields
            Load lblDestinationFieldNameForInput(intIndex)
            lblDestinationFieldNameForInput(intIndex).Top = cmbSourceFieldNameForInput(intIndex).Top
            lblDestinationFieldNameForInput(intIndex).Visible = True
            
            Load lblDestinationFieldType(intIndex)
            lblDestinationFieldType(intIndex).Top = cmbSourceFieldNameForInput(intIndex).Top
            lblDestinationFieldType(intIndex).Visible = True
            
            Load lblDestinationFieldSize(intIndex)
            lblDestinationFieldSize(intIndex).Top = cmbSourceFieldNameForInput(intIndex).Top
            lblDestinationFieldSize(intIndex).Visible = True
            
            Load txtDestinationFieldName(intIndex)
            
            
        End If
        
        'Resize Form
        Dim a As Integer
                
        a = cmbSourceFieldNameForInput(intIndex).Top + cmbSourceFieldNameForInput(intIndex).Height + intFieldSpacing
        
        If a >= cmdMoveDocsBetweenApplications.Top Then
            Dim intAmountToGrow As Integer
            intAmountToGrow = cmbSourceFieldNameForInput(intIndex).Height + intFieldSpacing
            Me.Height = Me.Height + intAmountToGrow
            cmdMoveDocsBetweenApplications.Top = cmdMoveDocsBetweenApplications.Top + intAmountToGrow
            ProgressBar1.Top = ProgressBar1.Top + intAmountToGrow
        End If
        
        lblDestinationFieldNameForInput(intIndex).Caption = rs("FieldNameForInput")
        txtDestinationFieldName(intIndex).Text = rs("FieldName")
        lblDestinationFieldType(intIndex).Caption = rs("FieldType")
        lblDestinationFieldSize(intIndex).Caption = rs("FieldSize")
        rs.MoveNext
        
        intIndex = intIndex + 1
    Wend

    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    con.Close
    Set con = Nothing
    
    


End Sub

