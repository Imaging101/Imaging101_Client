VERSION 5.00
Begin VB.Form I101MoveDocumentsBetweenApplications 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Move Documents Between Applications"
   ClientHeight    =   7575
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMoveDocsBetweenApplications 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Move documents "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4800
      Picture         =   "I101MoveDocumentsBetweenApplications.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6600
      Width           =   2055
   End
   Begin VB.Timer TimerDestination 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4080
      Top             =   1440
   End
   Begin VB.TextBox txtDestinationApplicationRECID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Left            =   360
      TabIndex        =   11
      Top             =   2280
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   10
      Top             =   2280
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   1320
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   8
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picImaging101Logo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5280
      Picture         =   "I101MoveDocumentsBetweenApplications.frx":0442
      ScaleHeight     =   495
      ScaleWidth      =   1695
      TabIndex        =   5
      Top             =   0
      Width           =   1695
   End
   Begin VB.ComboBox cmbDestinationField 
      Height          =   315
      Index           =   0
      Left            =   2880
      TabIndex        =   4
      Top             =   2640
      Width           =   3495
   End
   Begin VB.ComboBox cmbDestinationApplicationList 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1920
      Width           =   4335
   End
   Begin VB.ComboBox cmbSourceApplicationList 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   4335
   End
   Begin VB.Label lblFieldDescription 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   12
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Move Documents Between Applications"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   4935
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      ForeColor       =   &H00008000&
      Height          =   225
      Left            =   5280
      TabIndex        =   6
      Top             =   480
      Width           =   1605
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select &DESTINATION Application"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label lblSelectApplication 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select &SOURCE Application"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   2535
   End
End
Attribute VB_Name = "I101MoveDocumentsBetweenApplications"
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
    
        
    rs.Source = "Select ApplicationRECID,ApplicationName, MaxItemsToRetrieve from I101Applications WHERE ApplicationName= '" & cmbDestinationApplicationList.Text & "'"
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LockType = adLockReadOnly
    
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

Private Sub cmbDestinationField_DropDown(Index As Integer)

    funcFillList cmbDestinationField(Index), RegImaging101ConnectionString, "I101Fields", "FieldNameForInput", "ApplicationRECID = " & txtSourceApplicationRECID, True, True

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
    rs.LockType = adLockReadOnly
    
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

    funcFillList cmbSourceField(Index), RegImaging101ConnectionString, "I101Fields", "FieldNameForInput", "ApplicationRECID = " & txtSourceApplicationRECID, True, True

End Sub

Private Sub Form_Load()


    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision


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
    rs.Source = rs.Source & " ORDER BY ApplicationName"

    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LockType = adLockReadOnly
    
    con.Errors.Clear
    
    rs.Open
    rs.MoveFirst
    
    For intIndex = 0 To rs.RecordCount - 1
    
        cmbSourceApplicationList.AddItem rs.Fields!ApplicationName
        cmbSourceApplicationList.ItemData(intIndex) = rs.Fields!ApplicationRECID
        
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



Private Sub TimerDestination_Timer()

    TimerDestination.enabled = False
    
    
    cmbDestinationField(0).Text = ""
    
    For i = 1 To cmbDestinationField.UBound
        Unload cmbDestinationField(i)
    Next

    '*** Declarations
'    Dim rs As ADODB.Recordset
'    Dim con As ADODB.Connection
    Dim ssql As String

    Set con = New ADODB.Connection
    Set rs = New ADODB.Recordset
    con.Open RegImaging101ConnectionString

    'sql statement to select items on the drop down list
    ssql = "SELECT FieldName  FROM I101Fields WHERE ApplicationRECID = " & txtDestinationApplicationRECID
    rs.Open ssql, con

    Dim intIndex As Integer
    On Error Resume Next
    
    While Not rs.EOF
    
        
        If intIndex > 0 Then
            Load lblFieldDescription(intIndex)
            lblFieldDescription(intIndex).Top = cmbDestinationField(intIndex - 1).Top + cmbDestinationField(intIndex - 1).Height + 100
            lblFieldDescription(intIndex).Visible = True
            
            Load cmbDestinationField(intIndex)
            cmbDestinationField(intIndex).Top = cmbDestinationField(intIndex - 1).Top + cmbDestinationField(intIndex - 1).Height + 100
            cmbDestinationField(intIndex).Visible = True
        End If
        
        lblFieldDescription(intIndex).Caption = rs("FieldName")
        
        rs.MoveNext
        
        intIndex = intIndex + 1
    Wend

    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    con.Close
    Set con = Nothing
    
    


End Sub

