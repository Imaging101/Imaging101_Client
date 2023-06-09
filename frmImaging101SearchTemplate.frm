VERSION 5.00
Begin VB.Form frmImaging101SearchTemplate 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Retrieval Search Template Form - Imaging101"
   ClientHeight    =   7125
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12390
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   12390
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtUsersTotal 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   10455
      TabIndex        =   29
      Top             =   1635
      Width           =   405
   End
   Begin VB.TextBox txtUsersSelected 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   11790
      TabIndex        =   27
      Top             =   1635
      Width           =   405
   End
   Begin VB.TextBox txtCurrentlyLoadedTemplateName 
      Height          =   285
      Left            =   3840
      TabIndex        =   26
      Text            =   "txtCurrentlyLoadedTemplateName"
      Top             =   1560
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.TextBox txtSearchTemplateRECID 
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
      Left            =   6480
      TabIndex        =   25
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdFieldList 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select Fields"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   7320
      Picture         =   "frmImaging101SearchTemplate.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   960
      Width           =   1155
   End
   Begin VB.CommandButton cmdEqualSign 
      BackColor       =   &H00FFFFFF&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8505
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   960
      Width           =   555
   End
   Begin VB.CommandButton cmdLessThanSign 
      BackColor       =   &H00FFFFFF&
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8490
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1275
      Width           =   555
   End
   Begin VB.CommandButton cmdNotEqualSign 
      BackColor       =   &H00FFFFFF&
      Caption         =   "<>"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8985
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   960
      Width           =   555
   End
   Begin VB.CommandButton cmdGreaterThanSign 
      BackColor       =   &H00FFFFFF&
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8970
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1275
      Width           =   555
   End
   Begin VB.CommandButton cmdLike 
      BackColor       =   &H00FFFFFF&
      Caption         =   "LIKE"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9465
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton cmdLeftParentheses 
      BackColor       =   &H00FFFFFF&
      Caption         =   "("
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10275
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   960
      Width           =   555
   End
   Begin VB.CommandButton cmdAND 
      BackColor       =   &H00FFFFFF&
      Caption         =   "AND"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10830
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   960
      Width           =   555
   End
   Begin VB.Timer Timer1 
      Left            =   6600
      Top             =   120
   End
   Begin VB.ListBox lstUserList 
      Height          =   5010
      ItemData        =   "frmImaging101SearchTemplate.frx":058A
      Left            =   9720
      List            =   "frmImaging101SearchTemplate.frx":058C
      Style           =   1  'Checkbox
      TabIndex        =   13
      Top             =   1875
      Width           =   2535
   End
   Begin VB.ComboBox cmbApplicationList 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   3495
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
      Left            =   2760
      TabIndex        =   9
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtApplicationName 
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
      Left            =   1920
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox cmbSearchTemplateList 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3720
      TabIndex        =   7
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   735
      Left            =   -15
      TabIndex        =   1
      Top             =   -30
      Width           =   6255
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
         Left            =   0
         Picture         =   "frmImaging101SearchTemplate.frx":058E
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   15
         Width           =   855
      End
      Begin VB.PictureBox picImaging101Logo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   4800
         Picture         =   "frmImaging101SearchTemplate.frx":0E58
         ScaleHeight     =   405
         ScaleWidth      =   1440
         TabIndex        =   5
         Top             =   120
         Width           =   1440
      End
      Begin VB.CommandButton cmdHelp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Help"
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
         Left            =   3420
         Picture         =   "frmImaging101SearchTemplate.frx":14EB
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
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
         Left            =   855
         Picture         =   "frmImaging101SearchTemplate.frx":1DB5
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
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C07000&
         Height          =   225
         Left            =   4800
         TabIndex        =   6
         Top             =   480
         Width           =   1245
      End
   End
   Begin VB.TextBox txtWhereFreehand 
      DragIcon        =   "frmImaging101SearchTemplate.frx":21F7
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5190
      Left            =   75
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1815
      Width           =   9585
   End
   Begin VB.CommandButton cmdNotLike 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NOT LIKE"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9450
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   1275
      Width           =   870
   End
   Begin VB.CommandButton cmdRightParentheses 
      BackColor       =   &H00FFFFFF&
      Caption         =   ")"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10290
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1275
      Width           =   555
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OR"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10815
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1275
      Width           =   555
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Selected:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   10965
      TabIndex        =   30
      Top             =   1635
      Width           =   780
   End
   Begin VB.Label Label33 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Users:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   9750
      TabIndex        =   28
      Top             =   1635
      Width           =   600
   End
   Begin VB.Label lblSelectApplication 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select &Application"
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
      TabIndex        =   12
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lblSelectSavedSearch 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select or Type in Search Template"
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
      Left            =   3720
      TabIndex        =   11
      Top             =   960
      Width           =   2655
   End
End
Attribute VB_Name = "frmImaging101SearchTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmbApplicationList_Click()

    txtApplicationRECID = cmbApplicationList.ItemData(cmbApplicationList.ListIndex)
    txtApplicationName = cmbApplicationList.Text

End Sub

Private Sub cmbApplicationList_DropDown()

    '*************************************************************
    '*** LOAD APPLICATION LIST DROP-DOWN
    
    Set con = New ADODB.Connection
    
    con.ConnectionTimeout = 120
    con.CommandTimeout = 600
        
    con.Open RegImaging101ConnectionString
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    
'*** Changed the Load to work with Security
    rs.Source = ""
    rs.Source = rs.Source & "Select * "
    rs.Source = rs.Source & " FROM I101Applications, I101SecurityApplications"
    rs.Source = rs.Source & " WHERE I101Applications.ApplicationRECID = I101SecurityApplications.ApplicationRECID And I101SecurityApplications.SecurityRECID = " & gsecSecurityRECID
    rs.Source = rs.Source & " ORDER BY ApplicationName"

    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LOCKTYPE = adLockReadOnly
    
    con.Errors.Clear
    
    rs.Open
    rs.MoveFirst
    
    For intIndex = 0 To rs.RecordCount - 1
        cmbApplicationList.AddItem rs.Fields!ApplicationName
        cmbApplicationList.ItemData(intIndex) = rs.Fields!ApplicationRECID
        rs.MoveNext
    Next
        
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    con.Close
    Set con = Nothing

    '****************************




    
End Sub



Private Sub cmbSearchTemplateList_Click()

    Me.MousePointer = vbHourglass
    
    'Store the selected Search Template
    funcGetSetUserSettings "SET", "SearchTemplate", cmbSearchTemplateList
    
    txtCurrentlyLoadedTemplateName = cmbSearchTemplateList
    
    txtSearchTemplateRECID = cmbSearchTemplateList.ItemData(cmbSearchTemplateList.ListIndex)
    
    txtWhereFreehand = funcGetFieldFromDB(RegImaging101ConnectionString, "I101SearchTemplates", "SearchTemplateRECID = " & txtSearchTemplateRECID, "WhereFreehand")
    
    
    '*** Get Users
    Dim i As Long
    Dim strNames As String
    Dim strSecurityRECID As String

    lstUserList.Visible = False

    For i = 0 To lstUserList.ListCount - 1    'loop through the items in the ListBox
        
        strSecurityRECID = funcGetFieldFromDB(RegImaging101ConnectionString, "I101SearchTemplateUsers", "SearchTemplateRECID = " & txtSearchTemplateRECID & " AND SecurityRECID = " & lstUserList.ItemData(i), "SecurityRECID")
        
        
        ' if the item is selected(checked)
        If strSecurityRECID <> "" Then
        
            lstUserList.Selected(i) = True
        
        Else
            
            lstUserList.Selected(i) = False


        End If
        
    Next

    txtUsersTotal = lstUserList.ListCount
    txtUsersSelected = lstUserList.SelCount

        lstUserList.Refresh
        lstUserList.ListIndex = 0
        lstUserList.Visible = True
        
    Me.MousePointer = vbNormal

End Sub

Private Sub cmbSearchTemplateList_DropDown()

    If Trim(cmbApplicationList) = "" Then
        result = MsgBox("Please select and APPLICATION and try again.", vbInformation, "No Application Selected")
        Exit Sub
    End If


    '******************************************************************
    '*** LOAD SEARCH TEMPLATE LIST DROP-DOWN
    
    Set con = New ADODB.Connection
    
    con.ConnectionTimeout = 120
    con.CommandTimeout = 600
        
    con.Open RegImaging101ConnectionString
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    
    rs.Source = ""
    rs.Source = rs.Source & "Select * "
    rs.Source = rs.Source & " FROM I101SearchTemplates"
    rs.Source = rs.Source & " WHERE ApplicationRECID = " & txtApplicationRECID
    rs.Source = rs.Source & " ORDER BY SearchTemplateName"

    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LOCKTYPE = adLockReadOnly
    
    con.Errors.Clear
    
    rs.Open
    
    If rs.RecordCount > 0 Then
    
        rs.MoveFirst
        
        cmbSearchTemplateList.Clear
        
        For intIndex = 0 To rs.RecordCount - 1
            cmbSearchTemplateList.AddItem rs.Fields!SearchTemplateName
            cmbSearchTemplateList.ItemData(intIndex) = rs.Fields!SearchTemplateRECID
            rs.MoveNext
        Next
        
    Else
        result = MsgBox("No Search Templates defined yet for this Application!", vbInformation, "No Search Templates")
        
    End If
    
        
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    con.Close
    Set con = Nothing

    '****************************

End Sub

Private Sub cmdAND_Click()

    subAddTextTotxtWhereFreehand "AND"
    
End Sub



Private Sub cmdDelete_Click()
    
    Dim strcommandtext As String
    
    '*** Delete User Records
    strcommandtext = "DELETE FROM I101SearchTemplateUsers WHERE SearchTemplateRECID = " & txtSearchTemplateRECID
    result = funcRunSQLCommand(RegImaging101ConnectionString, strcommandtext)
    
    '*** Delete Template Records
    strcommandtext = "DELETE FROM I101SearchTemplates WHERE SearchTemplateRECID = " & txtSearchTemplateRECID
    result = funcRunSQLCommand(RegImaging101ConnectionString, strcommandtext)

    result = MsgBox("Search Template [" & cmbSearchTemplateList & "] DELETED.", vbInformation, "Search Template Deleted")

End Sub

Private Sub cmdEqualSign_Click()

    subAddTextTotxtWhereFreehand "="
    
End Sub

Private Sub cmdFieldList_Click()

    '*** POPULATE the FieldList Drop-Down
    
'    select column_name from INFORMATION_SCHEMA.COLUMNS
'    where table_name = 'CLIENT_FILES'
'    order by column_name

    frmDropDownList.Show
    frmDropDownList.Caption = "Field List"
    
    'CLEAR the List
    frmDropDownList.lstDropDownsList.Clear
    
    'ADD Special Fields - surrounded with Single-Quotes
    frmDropDownList.lstDropDownsList.AddItem "'{LoggedInUserID}'"
    frmDropDownList.lstDropDownsList.AddItem "'{CurrentDate}'"
    
    'FILL the List
    funcFillList frmDropDownList.lstDropDownsList, RegImaging101ConnectionString, "INFORMATION_SCHEMA.COLUMNS", "column_name", "table_name = '" & cmbApplicationList & "'", False, False
    

    
    frmImaging101SearchTemplate.SetFocus
    frmImaging101SearchTemplate.txtWhereFreehand.SetFocus
    
End Sub



Private Sub cmdGreaterThanSign_Click()

    subAddTextTotxtWhereFreehand ">"
    
End Sub

Private Sub cmdLeftParentheses_Click()

    subAddTextTotxtWhereFreehand "("
    
End Sub

Private Sub cmdLessThanSign_Click()

    subAddTextTotxtWhereFreehand "<"
    
End Sub

Private Sub cmdLike_Click()

    subAddTextTotxtWhereFreehand "LIKE"
    
End Sub

Private Sub cmdNotEqualSign_Click()

    subAddTextTotxtWhereFreehand "<>"
    
End Sub

Private Sub cmdNotLike_Click()

    subAddTextTotxtWhereFreehand "NOT LIKE"
    
End Sub

Private Sub cmdRightParentheses_Click()

    subAddTextTotxtWhereFreehand ")"
    
End Sub

Private Sub cmdSave_Click()

    subSaveSearchTemplate

End Sub

Private Sub Command1_Click()

    subAddTextTotxtWhereFreehand "OR"
    
End Sub

Private Sub Form_Load()

    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
    
    subSecurityLoadUserIDs

End Sub

Private Sub subSecurityLoadUserIDs()

    '*** CLEAR the USERLIST & Combos
    lstUserList.Clear
    
    '*************************************************************
    '*** LOAD UserID's   - BEGIN
    
    txtActionBeforeError = "Connect to Imaging101 DB"
    
    Set con = New ADODB.Connection
    con.Open RegImaging101ConnectionString

    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    
    rs.Source = "Select * from I101Security ORDER BY UserName"
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LOCKTYPE = adLockReadOnly
    
    rs.Open
    
    txtActionBeforeError = "Populate UserID List"
    
    For intIndex = 0 To rs.RecordCount - 1
        lstUserList.AddItem rs.Fields("UserName")
        lstUserList.ItemData(lstUserList.ListCount - 1) = rs.Fields("SecurityRECID")
        rs.MoveNext
        DoEvents
    Next
        
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    
    '*** LOAD UserID's   - END
    '*************************************************************

End Sub

Private Sub Form_Resize()

        Frame1.width = Me.width
        picImaging101Logo.Left = Me.ScaleWidth - picImaging101Logo.width - 10
        lblVersion.Left = picImaging101Logo.Left
        
        txtWhereFreehand.Height = Me.Height - txtWhereFreehand.Top - Frame1.Height
        lstUserList.Height = Me.Height - lstUserList.Top - Frame1.Height
        lstUserList.width = Me.ScaleWidth - lstUserList.Left - 50
        
End Sub

Private Sub subAddTextTotxtWhereFreehand(ByVal txtTextToAdd As String)

        Dim intCursorPos As Integer
        Dim intFieldLen As Integer
        Dim txtTextLeft As String
        Dim txtTextRight As String
    
        intCursorPos = frmImaging101SearchTemplate.txtWhereFreehand.SelStart
        intFieldLen = Len(txtTextToAdd) + 2 'to account for spaces before and after
        txtTextLeft = Left(frmImaging101SearchTemplate.txtWhereFreehand.Text, intCursorPos)
        txtTextRight = Right(frmImaging101SearchTemplate.txtWhereFreehand.Text, Len(frmImaging101SearchTemplate.txtWhereFreehand.Text) - intCursorPos)

        txtWhereFreehand.Text = txtTextLeft & " " & txtTextToAdd & " " & txtTextRight
        Me.SetFocus
        txtWhereFreehand.SetFocus
        txtWhereFreehand.SelStart = intCursorPos + intFieldLen

End Sub


Private Sub subSaveSearchTemplate()

    If lstUserList.SelCount > 0 Then
'        MsgBox lstUserList.SelCount & " items are selected"
    Else
        MsgBox "Sorry,no users are selected !"
        Exit Sub
    End If
    
    Dim strcommandtext As String

    '*** Check if Search Record Already Exists
    If Trim(txtSearchTemplateRECID) <> "" And cmbSearchTemplateList = txtCurrentlyLoadedTemplateName Then
    
        result = MsgBox("Do you want to OVERWRITE this Search Temlplate?" & vbCrLf & "If NOT, then Click 'NO', then edit the Search Template Name.", vbYesNo, "Overwrite Search Template?")
        If result = vbNo Then
            Exit Sub
        End If
        
        'UPDATE Search Template Record
        'Fix Error if Single Quotes are used in Lookup value
        'because SQL-Server will return Runtime Error "Incorrect Syntax".
        'Replace Single Quotes with TWO Single Quotes

        strcommandtext = "UPDATE I101SearchTemplates "
        strcommandtext = strcommandtext & "SET SearchTemplateRECID = " & txtSearchTemplateRECID
        strcommandtext = strcommandtext & ",      ApplicationRECID = " & txtApplicationRECID
        strcommandtext = strcommandtext & ",      SearchTemplateName = '" & cmbSearchTemplateList & "'"
        strcommandtext = strcommandtext & ",      WhereFreehand = '" & Replace(txtWhereFreehand, "'", "''") & "'"
        strcommandtext = strcommandtext & " WHERE  SearchTemplateRECID = " & txtSearchTemplateRECID

        result = funcRunSQLCommand(RegImaging101ConnectionString, strcommandtext)
        
    Else
    
        'Search Template Record does NOT exist, CREATE it.
        
        txtSearchTemplateRECID = funcGetNextControlNumber(RegImaging101ConnectionString, "I101Control", "SearchTemplateRECID")
        
        'Fix Error if Single Quotes are used in Lookup value
        'because SQL-Server will return Runtime Error "Incorrect Syntax".
        'Replace Single Quotes with TWO Single Quotes

        strcommandtext = "INSERT INTO I101SearchTemplates (SearchTemplateRECID, ApplicationRECID, SearchTemplateName, WhereFreehand) "
        strcommandtext = strcommandtext & " VALUES ( " & txtSearchTemplateRECID & ", " & txtApplicationRECID & ", '" & cmbSearchTemplateList & "', '" & Replace(txtWhereFreehand, "'", "''") & "')"
        
        result = funcRunSQLCommand(RegImaging101ConnectionString, strcommandtext)
        
        'Hold the new name
        txtCurrentlyLoadedTemplateName = cmbSearchTemplateList
        
    End If
    
    '*** Delete ALL User Records
    strcommandtext = "DELETE FROM I101SearchTemplateUsers WHERE SearchTemplateRECID = " & txtSearchTemplateRECID
    result = funcRunSQLCommand(RegImaging101ConnectionString, strcommandtext)

    
     Dim i As Long
     Dim strNames As String

    For i = 0 To lstUserList.ListCount - 1    'loop through the items in the ListBox
    
        If lstUserList.Selected(i) = True Then    ' if the item is selected(checked)
        
            strNames = strNames & lstUserList.ItemData(i) & "  =  " & lstUserList.List(i) & vbCrLf
            
            strcommandtext = "INSERT INTO I101SearchTemplateUsers (SearchTemplateRECID, SecurityRECID) "
            strcommandtext = strcommandtext & " VALUES ( " & txtSearchTemplateRECID & ", " & lstUserList.ItemData(i) & ")"

            result = funcRunSQLCommand(RegImaging101ConnectionString, strcommandtext)

        End If
        
    Next
    
'    MsgBox strNames        ' display the item
    
End Sub
