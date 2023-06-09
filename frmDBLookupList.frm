VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmTableLookupFields 
   Caption         =   "Document Types - Imaging101"
   ClientHeight    =   7704
   ClientLeft      =   60
   ClientTop       =   636
   ClientWidth     =   11832
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   7.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7704
   ScaleWidth      =   11832
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbApplicationList 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   1800
      TabIndex        =   6
      Top             =   960
      Width           =   4935
   End
   Begin VB.ComboBox cmbOrderBy 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      ItemData        =   "frmDBLookupList.frx":0000
      Left            =   1800
      List            =   "frmDBLookupList.frx":002B
      TabIndex        =   7
      Text            =   "APPLICATION,AREA,DOCGROUP,DOCTYPE,FORMDESC"
      Top             =   1320
      Width           =   6735
   End
   Begin VB.PictureBox picFields 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2460
      Left            =   0
      ScaleHeight     =   2460
      ScaleWidth      =   11832
      TabIndex        =   17
      Top             =   4875
      Width           =   11835
      Begin VB.TextBox txtRECID 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         Height          =   285
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1335
      End
      Begin VB.ComboBox cmbApplications 
         Height          =   288
         Left            =   2040
         TabIndex        =   0
         Top             =   120
         Width           =   4455
      End
      Begin VB.ComboBox cmbAreas 
         Height          =   288
         Left            =   2040
         TabIndex        =   1
         Top             =   480
         Width           =   4455
      End
      Begin VB.ComboBox cmbDocumentGroups 
         Height          =   288
         Left            =   2040
         TabIndex        =   2
         Top             =   840
         Width           =   5895
      End
      Begin VB.ComboBox cmbDocumentTypes 
         Height          =   288
         Left            =   2040
         TabIndex        =   3
         Top             =   1200
         Width           =   5895
      End
      Begin VB.TextBox txtPages 
         Height          =   285
         Left            =   2040
         TabIndex        =   5
         Top             =   1920
         Width           =   1575
      End
      Begin VB.ComboBox cmbFormDescriptions 
         Height          =   288
         Left            =   2040
         TabIndex        =   4
         Top             =   1560
         Width           =   5895
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         Caption         =   "Record ID"
         Height          =   255
         Left            =   5520
         TabIndex        =   27
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000016&
         Caption         =   "Application"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000016&
         Caption         =   "Area"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000016&
         Caption         =   "Document Group"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000016&
         Caption         =   "Document Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000016&
         Caption         =   "Form Description"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000016&
         Caption         =   "Average # of Pages"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1920
         Width           =   1815
      End
   End
   Begin VB.PictureBox picImaging101Logo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   564
      Left            =   9600
      Picture         =   "frmDBLookupList.frx":0241
      ScaleHeight     =   564
      ScaleWidth      =   1836
      TabIndex        =   16
      Top             =   0
      Width           =   1836
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   7335
      Width           =   11835
      _ExtentX        =   20870
      _ExtentY        =   656
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame frameButtons 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   -120
      TabIndex        =   10
      Top             =   -120
      Width           =   9615
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3060
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmDBLookupList.frx":103A
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1020
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   0
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmDBLookupList.frx":147C
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1020
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1020
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmDBLookupList.frx":18BE
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1020
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2040
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmDBLookupList.frx":1D00
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1020
      End
      Begin VB.CommandButton cmdClearFields 
         Caption         =   "&Clear Fields"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4080
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmDBLookupList.frx":2142
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1020
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2775
      Left            =   0
      TabIndex        =   8
      Top             =   1800
      Width           =   11415
      _ExtentX        =   20130
      _ExtentY        =   4890
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label lblSelectApplication 
      Caption         =   "Select &Application"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Order By"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      ForeColor       =   &H00C00000&
      Height          =   228
      Left            =   9720
      TabIndex        =   9
      Top             =   720
      Width           =   1812
   End
End
Attribute VB_Name = "frmTableLookupFields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

    '****************************
    '*** Declarations
    Dim con As ADODB.Connection
    Dim ssql As String
    Dim cmd As ADODB.Command
    
    Dim mode As String


Private Sub cmbApplicationList_Click()

    subPopulateListview
    
End Sub

Private Sub cmbApplicationList_DropDown()

    funcFillList Me.cmbApplicationList, RegImaging101ConnectionString, "DOCTYPES", "APPLICATION", "", False, True

End Sub

Private Sub cmbApplications_DropDown()

    funcFillList Me.cmbApplications, RegImaging101ConnectionString, "DOCTYPES", "APPLICATION", "", False, True

End Sub

Private Sub cmbAreas_DropDown()

    funcFillList Me.cmbAreas, RegImaging101ConnectionString, "DOCTYPES", "AREA", "", False, True

End Sub

Private Sub cmbDocumentTypes_DropDown()
    
    funcFillList Me.cmbDocumentTypes, RegImaging101ConnectionString, "DOCTYPES", "DOCTYPE", "", False, True

End Sub

Private Sub cmbFormDescriptions_DropDown()

    funcFillList Me.cmbFormDescriptions, RegImaging101ConnectionString, "DOCTYPES", "FORMDESC", "", False, True

End Sub

Private Sub cmbDocumentGroups_DropDown()

    funcFillList Me.cmbDocumentGroups, RegImaging101ConnectionString, "DOCTYPES", "DOCGROUP", "", False, True

End Sub

Private Sub cmbFormDescriptions_GotFocus()

    If Trim(cmbFormDescriptions.Text = "") Then
        cmbFormDescriptions.Text = cmbDocumentTypes.Text
    End If

End Sub

Private Sub cmbOrderBy_Click()

    subPopulateListview
    
End Sub

Private Sub cmdAdd_Click()

    'Add New Transaction Record
'    subClearForm
    cmdAdd.enabled = False
    cmdDelete.enabled = False
    cmdUpdate.enabled = False

    If Trim(cmbApplications.Text & cmbAreas.Text & cmbDocumentGroups.Text & cmbDocumentTypes.Text) = "" Then
        MsgBox "Cannot Create an Empty Record... Please fill in the appropriate fields and click [Add] !", vbInformation, "No Field Values Assigned"
        StatusBar1.Panels(1).Text = "Empty Record NOT Added."
        cmdAdd.enabled = True
        Exit Sub
    End If
    
    StatusBar1.Panels(1).Text = "Adding New Document Type Record."
    subAddTransaction
    
    cmdAdd.enabled = True
    cmdClearFields.Visible = True
    cmdUpdate.enabled = True
    cmdDelete.enabled = True
    
End Sub

Private Sub cmdClearFields_Click()

    subClearForm
    cmdAdd.enabled = True
    cmdUpdate.enabled = False
    cmdDelete.enabled = False
    
End Sub


Private Sub cmdDelete_Click()
    
    Dim result As String
    
    result = MsgBox("Are you sure you wish to DELETE record #: " & txtRECID & " ?" & _
                    vbCrLf & " Application: " & cmbApplications.Text & _
                    vbCrLf & " Area         : " & cmbAreas.Text & _
                    vbCrLf & " Doc Group : " & cmbDocumentGroups.Text & _
                    vbCrLf & " Doc Type   : " & cmbDocumentTypes.Text & _
                    vbCrLf & " Form Desc : " & cmbFormDescriptions.Text _
                    , vbYesNo, "Delete Transaction")
                    
    If result = vbNo Then
        StatusBar1.Panels(1).Text = "Delete CANCELLED!"
        Exit Sub
    End If

    StatusBar1.Panels(1).Text = "Deleting Document Type Record."
    subDeleteTransaction
    subPopulateListview
    
    cmdDelete.enabled = False
    cmdUpdate.enabled = False
    
    
End Sub





Private Sub cmdRefresh_Click()

    'Refresh the Transaction List & Combos
    StatusBar1.Panels(1).Text = "Refreshing..."
    
    subPopulateListview
    
    StatusBar1.Panels(1).Text = "Refresh Complete"

End Sub


Private Sub cmdUpdate_Click()


        StatusBar1.Panels(1).Text = "Updating Record"
        subUpdateTransaction
        
        subPopulateListview
        
        cmdClearFields.Visible = True
        cmdUpdate.enabled = False
        cmdDelete.enabled = False
        
        
End Sub




Private Sub Form_Load()

    cmdDelete.enabled = False
    cmdUpdate.enabled = False
    

    cmbApplicationList.Text = frmConfig.txtApplicationName
    cmbApplicationList_DropDown
    
    subPopulateListview

    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision

End Sub

Private Sub Form_Resize()
  On Error Resume Next
  'This will resize the grid when the form is resized
  picFields.Top = Me.ScaleHeight - picFields.Height - StatusBar1.Height
  picFields.width = Me.ScaleWidth - picFields.Left - 50
  ListView1.Height = Me.ScaleHeight - ListView1.Top - picFields.Height - StatusBar1.Height - 50
  ListView1.width = Me.ScaleWidth - ListView1.Left - 50
  StatusBar1.Panels(1).width = StatusBar1.width
End Sub

Public Sub subPopulateListview()

    
    ListView1.ListItems.Clear
        
     '*** Declarations -- MOVED TO MODULE LEVEL TOP

    Set con = New ADODB.Connection
    con.Open RegImaging101ConnectionString
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Dim strSQL As String
    Dim strSELECT As String
    Dim strWhere As String
    Dim strORDERBY As String
      
    strSELECT = "Select ID,APPLICATION,AREA,DOCGROUP,DOCTYPE,FORMDESC,PAGES From DOCTYPES "
    
    If Trim(cmbApplicationList.Text) <> "" Then
      strWhere = " WHERE APPLICATION = '" & cmbApplicationList.Text & "'"
    End If
    
    If cmbOrderBy.Text <> "" Then
      strORDERBY = " Order by " & cmbOrderBy.Text
    End If
    
    strSQL = strSELECT & " " & strWhere & " " & strORDERBY
        
    With rs
        .ActiveConnection = con
        .CursorLocation = adUseServer
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .Source = strSQL
    End With
        
    con.Errors.Clear
    rs.Open
        
    
   
    '*** Setup Up ListView properties - BEGIN
    
    ListView1.Visible = False
    ListView1.View = lvwReport
    ListView1.ColumnHeaders.Clear
    '    Column widths are in PIXELS!
    ListView1.MultiSelect = True
    ListView1.FullRowSelect = True
    ListView1.LabelEdit = lvwManual
    ListView1.AllowColumnReorder = True
    
        '***  SET COLUMN HEADINGS
        For intListIndex = 0 To rs.Fields.count - 1
            ListView1.ColumnHeaders.Add , , rs.Fields.item(intListIndex).name, Len(rs.Fields.item(intListIndex).name) * 150, lvwColumnLeft
        Next
                
    On Error Resume Next
    
    If Not rs.EOF Then
        rs.MoveFirst
    End If
    
    While Not rs.EOF
            For intListIndex = 0 To rs.Fields.count - 1
                If intListIndex = 0 Then
                    If Not IsNull(rs.Fields.item(intListIndex).Value) Then
                        Set lstItem = ListView1.ListItems.Add(, , rs.Fields.item(intListIndex).Value)
                    End If
                Else
            
                        '* This null check is to make sure we don't Skip fields caused by an error.
                        If Not IsNull(rs.Fields.item(intListIndex).Value) Then
                            ' Not null... show value
                            Select Case rs.Fields.item(intListIndex).Type
                                Case adDBTimeStamp
                                    Set lstSubItem = lstItem.ListSubItems.Add(, , Format(rs.Fields.item(intListIndex).Value, "yyyy/mm/dd AMPM hh:mm:ss "))
'                                Case adInteger
'                                    Set lstSubItem = lstItem.ListSubItems.Add(, , Right("      " & Format(rs.Fields.item(intListIndex).Value, "##,###"), 6))
'                                Case adNumeric, adDouble
'                                    Set lstSubItem = lstItem.ListSubItems.Add(, , Right("          " & Format(rs.Fields.item(intListIndex).Value, "##,###,###"), 10))
                                Case adCurrency
                                    Set lstSubItem = lstItem.ListSubItems.Add(, , Right("                " & Format(rs.Fields.item(intListIndex).Value, "$##,###,##0.00"), 14))
                                Case Else
                                    Set lstSubItem = lstItem.ListSubItems.Add(, , rs.Fields.item(intListIndex).Value)
                            End Select
                            
'                            If rs.Fields.item(intListIndex).Type = adDBTimeStamp Then
'                                Set lstSubItem = lstItem.ListSubItems.Add(, , Format(rs.Fields.item(intListIndex).Value, "yyyy/mm/dd AMPM hh:mm:ss "))
'                            Else
'                                Set lstSubItem = lstItem.ListSubItems.Add(, , rs.Fields.item(intListIndex).Value)
'                            End If
                        Else
                            ' Null... show empty string
                            Set lstSubItem = lstItem.ListSubItems.Add(, , "")
                        End If
                        
                        If rs.Fields.item(intListIndex).name = "BatchNotes" Then
                            lstItem.ListSubItems(intListIndex).ForeColor = vbRed
                       End If
                End If
            Next
        rs.MoveNext
    Wend
    On Error GoTo 0
    
    ' AutoSize ALL Columns
    Dim i As Integer, lparam As Long
    UseHeader = True
    If UseHeader = False Then
        lparam = LVSCW_AUTOSIZE
    Else
        lparam = LVSCW_AUTOSIZE_USEHEADER
    End If
    For i = 0 To ListView1.ColumnHeaders.count - 1
        SendMessage ListView1.hwnd, LVM_SETCOLUMNWIDTH, i, ByVal lparam
    Next
    
    ' Hide the RecordID's
    ListView1.ColumnHeaders(1).width = 0
'    ListView1.ColumnHeaders(2).Width = 0
    
'    ' Size the Key fields to a standard size
'    ListView1.ColumnHeaders(3).Width = 3000
    
    ListView1.Visible = True

    '*** Setup Up ListView properties - END
    
    ' Disable Buttons until at least ONE ROW is selected

    rs.Close
    Set rs = Nothing
    

End Sub


Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    If ListView1.SortOrder = lvwAscending Then
        ListView1.SortOrder = lvwDescending
    Else
        ListView1.SortOrder = lvwAscending
    End If
    ' Set the Sort Column
    ListView1.SortKey = ColumnHeader.Index - 1
    ' Sort It!
    ListView1.Sorted = True
    
End Sub

Private Sub AutoSizeColumns(Listview As Listview, Optional ByVal UseHeader As Boolean = False)
  Dim i As Integer, lparam As Long
  If UseHeader = False Then
      lparam = LVSCW_AUTOSIZE
  Else
      lparam = LVSCW_AUTOSIZE_USEHEADER
  End If
  For i = 0 To Listview.ColumnHeaders.count - 1
      SendMessage Listview.hwnd, LVM_SETCOLUMNWIDTH, i, ByVal lparam
  Next
End Sub


Private Sub subClearForm()

        ' Clear the data entry form.
        cmbApplications.Text() = ""
        cmbAreas.Text() = ""
        cmbDocumentGroups.Text() = ""
        cmbDocumentTypes.Text() = ""
        cmbFormDescriptions.Text() = ""
        txtPages.Text() = ""

End Sub

Private Sub subDeleteTransaction()
    ' This sub is used to delete the product record from the database
    ' when the user clicks the delete button

    On Error GoTo ERROR_HANDLER
    
    Set con = New ADODB.Connection
    con.Open RegImaging101ConnectionString
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    With rs
        .ActiveConnection = con
        .CursorLocation = adUseServer
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
    End With


        ' Build Select statement to query product information from the products
        ' table
        rs.Source = "DELETE FROM DOCTYPES " & _
                 "WHERE ID = " & CDbl(txtRECID.Text)

        Err.Clear
        
        rs.Open
        
        ' Close and Clean up objects
        'The Delete call automatically closes the RecordSet (rs)
        con.Close
        Set rs = Nothing
        Set con = Nothing
        
        subClearForm
        cmdDelete.enabled = False
        cmdUpdate.enabled = False
        
        StatusBar1.Panels(1).Text = "Document Type Record Deleted."
    
Exit Sub

ERROR_HANDLER:

    MsgBox "An Error Occured while attempting to DELETE a record... Please contact Technical Support!", vbCritical, "Delete Record Error"
    StatusBar1.Panels(1).Text = "Error Occured!  Document Type Record NOT Deleted."

    Set rs = Nothing
    Set con = Nothing

        
End Sub

Private Sub subPopulateForm()
    
    
    
    Dim lstIndex As Long
    
    lstIndex = Me.ListView1.SelectedItem.Index
    ' Get Main Item
    txtRECID.Text = Me.ListView1.ListItems(lstIndex).Text

    Set con = New ADODB.Connection
    con.Open RegImaging101ConnectionString
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    With rs
        .ActiveConnection = con
        .CursorLocation = adUseServer
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
    End With


        ' Build Select statement to query product information from the products
        ' table
        rs.Source = "SELECT * " & _
                 "FROM DOCTYPES " & _
                 "WHERE ID = " & txtRECID.Text


    rs.Open
        
    If rs.RecordCount > 0 Then
    
        ' Populate form with the data
        txtRECID = rs.Fields("ID")
        cmbApplications = rs.Fields("APPLICATION") & ""
        cmbAreas = rs.Fields("AREA") & ""
        cmbDocumentGroups = rs.Fields("DOCGROUP") & ""
        cmbDocumentTypes.Text = rs.Fields("DOCTYPE") & ""
        cmbFormDescriptions.Text = rs.Fields("FORMDESC") & ""
        txtPages.Text = rs.Fields("PAGES") & ""
    End If

    rs.Close
    Set rs = Nothing
    
    cmdAdd.enabled = True
    cmdDelete.enabled = True
    cmdUpdate.enabled = True
    
    StatusBar1.Panels(1).Text = "To Modify this record, make the required changes and Click the [Update] button.  To Create a duplicate click [Add]."
    
End Sub




Private Sub ListView1_ItemClick(ByVal item As MSComctlLib.ListItem)
    
    If mode = "Add" Then
        result = MsgBox("You were Adding a record... this will clear what you have typed!  Are you sure you wish to display the selected item?", vbYesNo, "Select Transaction")
        If result = vbNo Then
            Exit Sub
        End If
    End If
    
    subPopulateForm
    
End Sub


Private Sub subAddTransaction()

    On Error GoTo ERROR_HANDLER
    
    Set con = New ADODB.Connection
    con.Open RegImaging101ConnectionString
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    With rs
        .ActiveConnection = con
        .CursorLocation = adUseServer
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Source = "DOCTYPES"
        
        .Open
        .AddNew
    
'        txtRECID = rs.Fields("ID")
        .Fields("APPLICATION") = PrepareStr2(cmbApplications.Text)
        .Fields("AREA") = PrepareStr2(cmbAreas.Text)
        .Fields("DOCGROUP") = PrepareStr2(cmbDocumentGroups.Text)
        .Fields("DOCTYPE") = PrepareStr2(cmbDocumentTypes.Text)
        .Fields("FORMDESC") = PrepareStr2(cmbFormDescriptions.Text)
        .Fields("PAGES") = PrepareStr2(txtPages.Text)
            
        .Update
        
        ' Close and Clean up objects
        .Close
    End With
    
    
    con.Close
    Set rs = Nothing
    Set con = Nothing
    
    ' Refresh Product List
    subPopulateListview
    
    StatusBar1.Panels(1).Text = "Document Type Record Added."

    
Exit Sub

ERROR_HANDLER:

    MsgBox "An Error Occured while attempting to ADD a record... please check your fields to make sure you have entered valid values!", vbCritical, "Add Record Error"
    StatusBar1.Panels(1).Text = "Error Occured!  Document Type Record NOT Added."

    Set rs = Nothing
    Set con = Nothing

End Sub
    
    
Private Function PrepareStr2(ByVal strValue As String) As String
    ' This function accepts a string and creates a string that can
    ' be used in a SQL statement by adding single quotes around
    ' it and handling empty values.
    If Trim(strValue) = "" Then
        PrepareStr2 = " "
    Else
        PrepareStr2 = Trim(strValue)
    End If
End Function


Private Sub subUpdateTransaction()
    
    On Error GoTo ERROR_HANDLER
    
    ' This sub is used to update and existing record with values
    ' from the form.
    Dim strSQL As String
    Dim intRowsAffected As Integer

    ' Validate form values.

    Set con = New ADODB.Connection
    con.Open RegImaging101ConnectionString
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    With rs
        .ActiveConnection = con
        .CursorLocation = adUseServer
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Source = "SELECT * " & _
                  "FROM DOCTYPES " & _
                  "WHERE ID = " & txtRECID.Text
        
        .Open
        
        If .RecordCount = 1 Then
    '        txtRECID = rs.Fields("ID")
            .Fields("APPLICATION") = PrepareStr2(cmbApplications.Text)
            .Fields("AREA") = PrepareStr2(cmbAreas.Text)
            .Fields("DOCGROUP") = PrepareStr2(cmbDocumentGroups.Text)
            .Fields("DOCTYPE") = PrepareStr2(cmbDocumentTypes.Text)
            .Fields("FORMDESC") = PrepareStr2(cmbFormDescriptions.Text)
            .Fields("PAGES") = PrepareStr2(txtPages.Text)
                
            .Update
            StatusBar1.Panels(1).Text = "Record Updated!"
        Else
            StatusBar1.Panels(1).Text = "Record NOT FOUND... NOT Updated!"
        End If
        
        ' Close and Clean up objects
        .Close
    End With

    ' Close and Clean up objects
    con.Close
    Set rs = Nothing
    Set con = Nothing
    
    subPopulateListview
    
    
Exit Sub

ERROR_HANDLER:

    MsgBox "An Error Occured while attempting to UPDATE a record... please check your fields to make sure you have entered valid values!", vbCritical, "Add Record Error"
    StatusBar1.Panels(1).Text = "Error Occured!  Document Type Record NOT Added."

    Set rs = Nothing
    Set con = Nothing

    
End Sub

Private Sub subLoadApplicationDropDown()

    '***************************************
    '*** LOAD APPLICATION LIST DROP-DOWN
    
    cmbApplicationList.Clear
    
    Dim con As ADODB.Connection
    Set con = New ADODB.Connection
    con.Open RegImaging101ConnectionString
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    
    rs.Source = "Select ApplicationName, ApplicationRECID from I101Applications ORDER BY ApplicationName"
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LockType = adLockReadOnly
    
    con.Errors.Clear
    
    rs.Open
    rs.MoveFirst
    
    cmbApplicationList.AddItem ""
    
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

