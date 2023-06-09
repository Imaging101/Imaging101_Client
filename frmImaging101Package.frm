VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmImaging101Package 
   BackColor       =   &H80000013&
   Caption         =   "Package Documents - Imaging101"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   7935
   Icon            =   "frmImaging101Package.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDocPackageSendFormat 
      Height          =   285
      Left            =   1560
      TabIndex        =   9
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox txtDocPackageSendMode 
      Height          =   285
      Left            =   1560
      TabIndex        =   8
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox txtDocPackageNotes 
      Height          =   645
      Left            =   1560
      TabIndex        =   7
      Top             =   1560
      Width           =   6135
   End
   Begin VB.ComboBox cmbPackageList 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   7695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5535
      Begin VB.PictureBox picImaging101Logo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   3840
         Picture         =   "frmImaging101Package.frx":0CCA
         ScaleHeight     =   495
         ScaleWidth      =   1695
         TabIndex        =   14
         Top             =   0
         Width           =   1695
      End
      Begin VB.CommandButton cmdPackageIt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Package"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   735
         Left            =   0
         Picture         =   "frmImaging101Package.frx":175A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton cmdHelp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Help"
         Height          =   735
         Left            =   960
         Picture         =   "frmImaging101Package.frx":2424
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   855
      End
      Begin VB.Label lblVersion 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         ForeColor       =   &H00008000&
         Height          =   225
         Left            =   3840
         TabIndex        =   5
         Top             =   480
         Width           =   1605
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4560
      Top             =   2880
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      ConnectMode     =   1
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   1
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3615
      Left            =   0
      TabIndex        =   13
      Top             =   3240
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   6376
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Send Format"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      TabIndex        =   12
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Send Mode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      TabIndex        =   11
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Package Notes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      TabIndex        =   10
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Package Documents"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      TabIndex        =   6
      Top             =   3000
      Width           =   2895
   End
   Begin VB.Label lblSelectApplication 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Select &Package"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
End
Attribute VB_Name = "frmImaging101Package"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmbPackageList_Click()

    subGetHeaderFields

    subBuildSelectStatement
    
    subPopulateListview

    cmdPackageIt.enabled = True

End Sub

Private Sub cmdPackageIt_Click()

    ' Generate Package
    
    

End Sub

Private Sub Form_Load()

    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision

    '***************************************
    '*** LOAD APPLICATION LIST DROP-DOWN
    
    Dim con As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    Set con = New ADODB.Connection
    con.Open RegImaging101ConnectionString
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    
    rs.Source = ""
    rs.Source = rs.Source & "Select * "
    rs.Source = rs.Source & " FROM I101DocPackage"
    rs.Source = rs.Source & " WHERE I101DocPackage.ApplicationRECID = " & frmImaging101Search.txtApplicationRECID
    rs.Source = rs.Source & " ORDER BY DocPackageDescription"

    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LockType = adLockReadOnly
    
    con.Errors.Clear
    
    rs.Open
    
    If rs.RecordCount > 0 Then
    
        rs.MoveFirst
        
        For intIndex = 0 To rs.RecordCount - 1
            cmbPackageList.AddItem rs.Fields!DocPackageDescription
            cmbPackageList.ItemData(intIndex) = rs.Fields!DocPackageRECID
            rs.MoveNext
        Next
        
        
    Else
        
        MsgBox "Sorry... No Packages Exist for this Application.", vbInformation, "No Packages Available"
    
    End If
    

    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    con.Close
    Set con = Nothing
    
    '****************************

End Sub

Private Sub Form_Resize()

    Frame1.width = Me.width
    picImaging101Logo.Left = Me.ScaleWidth - picImaging101Logo.width - 10
    lblVersion.Left = picImaging101Logo.Left

End Sub


Public Sub subPopulateListview()

    On Error GoTo ERROR_TRAP
    
    '*** Setup Up ListView properties - BEGIN
    
    ListView1.ListItems.Clear
    
    ListView1.Visible = False
    ListView1.View = lvwReport
    ListView1.ColumnHeaders.Clear
    '    Column widths are in PIXELS!
    ListView1.MultiSelect = True
    ListView1.FullRowSelect = True
    ListView1.LabelEdit = lvwManual
    ListView1.AllowColumnReorder = True
    
        '***  SET COLUMN HEADINGS
        For intListIndex = 0 To Adodc1.Recordset.Fields.Count - 1
            ListView1.ColumnHeaders.Add , , Adodc1.Recordset.Fields.item(intListIndex).name, Len(Adodc1.Recordset.Fields.item(intListIndex).name) * 150, lvwColumnLeft
            
            
        Next
                
'    On Error Resume Next
'
'    Adodc1.Recordset.MoveFirst

    While Not Adodc1.Recordset.EOF
            For intListIndex = 0 To Adodc1.Recordset.Fields.Count - 1
                If intListIndex = 0 Then
                    If Not IsNull(Adodc1.Recordset.Fields.item(intListIndex).Value) Then
                        Set lstItem = ListView1.ListItems.Add(, , Adodc1.Recordset.Fields.item(intListIndex).Value)
                    End If
                Else
            
                    '* This null check is to make sure we don't Skip fields caused by an error.
                    If Not IsNull(Adodc1.Recordset.Fields.item(intListIndex).Value) Then
                        ' Not null... show value
                        
                        Select Case Adodc1.Recordset.Fields.item(intListIndex).Type
                            Case adNumeric, adInteger, adDouble, adSingle, adSmallInt
                                '*** FORCE RIGHT ALIGNMENT OF NUMBERS
                                Set lstSubItem = lstItem.ListSubItems.Add(, , Right("             " & Trim(CStr(Adodc1.Recordset.Fields.item(intListIndex).Value)), 12))
                            Case adDBTimeStamp
                                Set lstSubItem = lstItem.ListSubItems.Add(, , Format(Adodc1.Recordset.Fields.item(intListIndex).Value, "yyyy/mm/dd"))
                            Case Else
                                Set lstSubItem = lstItem.ListSubItems.Add(, , Adodc1.Recordset.Fields.item(intListIndex).Value)
                        End Select
                        
                    Else
                        ' Null... show empty string
                        Set lstSubItem = lstItem.ListSubItems.Add(, , "")
                    End If
                    
                End If
            Next
        Adodc1.Recordset.MoveNext
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
    For i = 0 To ListView1.ColumnHeaders.Count - 1
        SendMessage ListView1.hwnd, LVM_SETCOLUMNWIDTH, i, ByVal lparam
    Next
    
    ' Hide the RecordID's
    ListView1.ColumnHeaders(1).width = 0
    ListView1.ColumnHeaders(2).width = 0
    ListView1.ColumnHeaders(3).width = 2000
    ListView1.ColumnHeaders(4).width = 5000
    
    
    ListView1.Visible = True

    '*** Setup Up ListView properties - END
    
    
Exit Sub
    
ERROR_TRAP:

    MsgBox "subPopulateListview ERROR: " & Err.Number & " - " & vbCrLf & Err.Description & "  DURING ACTION: (" & txtActionBeforeError & ")", vbExclamation
    
End Sub


Private Sub subBuildSelectStatement()

    Dim txtFilterStatement As String
    Dim txtFieldsList As String
    Dim txtOrderByList As String
    Dim strDateFormatted As String
    
    On Error GoTo ERROR_TRAP
    
    '*** Clear variables
    txtFilterStatement = ""
    txtFieldsList = ""
    txtOrderByList = ""
    
    'RESET the Error occured flag
    bolErrorOccured = False
    
    
    '*** WE ARE SETTING THE frmImaging101Retrieve (Search Results List) FORM CONTROLS
   Adodc1.ConnectionString = RegImaging101ConnectionString
   Adodc1.RecordSource = "SELECT " & _
                        " DocPackageDetailRECID, " & _
                        " DocPackageRECID, " & _
                        " DocPackageDetailOrder, " & _
                        " DOCTYPE " & _
                        " FROM I101DocPackageDetail" & _
                        " WHERE DocPackageRECID = " & cmbPackageList.ItemData(cmbPackageList.ListIndex) & _
                        " ORDER BY DocPackageDetailOrder"

   Adodc1.Refresh
   
    
   
Exit Sub
    
ERROR_TRAP:
    ' 94 = Invalid use of Null
    If Err.Number = 94 Then
        Err.Clear
        Resume Next
    End If
    
    result = MsgBox("subBuildSelectStatement - Error: " & Err.Number & " - " & Err.Description, vbOK)
    Err.Clear
    
End Sub

Private Sub subGetHeaderFields()

    Set con = New ADODB.Connection
    con.Open RegImaging101ConnectionString
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    
    rs.Source = ""
    rs.Source = rs.Source & "Select * "
    rs.Source = rs.Source & " FROM I101DocPackage"
    rs.Source = rs.Source & " WHERE I101DocPackage.DocPackageRECID = " & cmbPackageList.ItemData(cmbPackageList.ListIndex)

    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LockType = adLockReadOnly
    
    con.Errors.Clear
    
    rs.Open
    rs.MoveFirst
    
    txtDocPackageNotes = rs.Fields("DocPackageNotes")
    txtDocPackageSendMode = rs.Fields("DocPackageSendMode")
    txtDocPackageSendFormat = rs.Fields("DocPackageSendFormat")
    
        
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    con.Close
    Set con = Nothing

End Sub

