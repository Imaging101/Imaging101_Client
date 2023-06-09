VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLookupList 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Table Lookup"
   ClientHeight    =   3405
   ClientLeft      =   330
   ClientTop       =   3960
   ClientWidth     =   7995
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLookupList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3405
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkHighlightLookpFieldAfterNextPage 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Highlight the Lookup Field after [Next Page]"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   360
      Width           =   3495
   End
   Begin VB.CommandButton cmdNextImage 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Next Page"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6750
      Picture         =   "frmLookupList.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Find"
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
      Height          =   615
      Left            =   5280
      Picture         =   "frmLookupList.frx":07CC
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox txtTableLookupField 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2295
   End
   Begin VB.CommandButton cmdClearFields 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6015
      Picture         =   "frmLookupList.frx":0D56
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox txtRowDescription 
      BackColor       =   &H00E0E9EF&
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
      Height          =   375
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   2775
   End
   Begin MSDataGridLib.DataGrid grdDataGrid 
      Bindings        =   "frmLookupList.frx":0EA0
      Height          =   2055
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   3625
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16777215
      HeadLines       =   1
      RowHeight       =   20
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         Locked          =   -1  'True
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdodcLookupList 
      Height          =   330
      Index           =   0
      Left            =   120
      Top             =   3000
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   582
      ConnectMode     =   1
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   1
      CommandType     =   1
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
      Caption         =   "AdodcLookupList"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdodcLookupList 
      Height          =   330
      Index           =   1
      Left            =   3480
      Top             =   3000
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   582
      ConnectMode     =   1
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   1
      CommandType     =   1
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
      Caption         =   "AdodcLookupList"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frmLookupList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'JR - 7/12/2001:  Added Support for AccountNumberBegin/intAccountNumberLength INI file parameters!
'                   convert the Strings to Longs using "CDbl()"
'JR - 7/20/2001:  Added Support for AccountNumberSearch INI file parameters
'                    to Begin Searching the DB after This number of characters.
    
Dim moIndexing As Object
Dim strAccountTableName As String

Dim strAccountNumberEcField As String
Dim strAccountNumberDbField As String
Dim strAccountNameEcField As String
Dim strAccountNameDbField As String
Dim strAccountCaseEcField As String
Dim strAccountCaseDbField As String
Dim strAccountCaseIdEcField As String
Dim strAccountCaseIdDbField As String
Dim strAccountTicketNumberEcField As String
Dim strAccountTicketNumberDbField As String

Dim strAccountNumberEcFieldBegin As String
Dim strAccountNumberEcFieldLength As String
Dim strAccountNumberSearchChars As String


'*** SELECT Statement Variables - Added 1/22/2004 by Jacob
Dim strFieldListForSelect
Dim strFieldListForWhere
Dim strFieldListForOrderBy
Dim intIndex As Integer
Dim strFieldToSelectAfterLookupClick As String
Dim strFieldToSelectAfterDocListClick As String
Dim bolAutoLookupOnBatchLoad As Boolean

'*** Index Population ARRAY Variables - Added 1/22/2004 by Jacob
Dim arrApplicationFieldNameForInput() As String
Dim arrLookupTableFieldName() As String
Dim arrColumnTitle() As String
Dim arrShowColumn() As String
Dim arrAllowLookup() As String
Dim arrTrimFieldBeginChar() As String
Dim arrTrimFieldLength() As String
Dim arrTreatAsNumeric() As String
Dim arrFieldSearchCondition() As String

Dim RegLookupListConnectionString As String
Dim RegLookupListConnectionType As String
Dim RegLookupDBTableIsOnSQLServer As String
Dim RegLookupListWhereClause As String

Dim arrLookupDBConnectionString(2) As String
Dim arrLookupDBTableName(2) As String
Dim arrLookupDBTableIsOnSQLServer(2) As String
Dim arrLookupDBWhereClause(2) As String

Dim dblCaseIdCutoff As Double

Dim arrAdodcLookupList(2) As ADODB.Connection
Dim arrAdodcLookupListRS(2) As ADODB.Recordset



Dim intSiteIdIndex As Integer

Dim strApplicationCommitBatchTo As String

Dim strDataBaseType(2) As String


Dim bolFormLoaded As Boolean
Dim bolWalkingDownDataGrid As Boolean


Private Const HWND_TOPMOST = -1
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long






Public Sub cmdFind_Click()

    

    'Bail Out if Nothing Entered
    If txtTableLookupField.Text = "" Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    
    On Error GoTo ERROR_HANDLER
    
'*** MOVED THIS SECTION TO THE FORM LOAD TO SPEED UP SEARCHES BY CONNECTING ONLY ONCE.
''''    '****************************************
''''    '*** Determine the Database Type
''''    If InStr(UCase(RegLookupListConnectionString), "SQLOLEDB.1") Then
''''        '*** SQL-Server/MSDE
''''        strDataBaseType = "SQL-Server/MSDE"
''''    ElseIf InStr(UCase(RegLookupListConnectionString), "MICROSOFT.JET") Then
''''        '*** MS-Acess/Jet DB's
''''        strDataBaseType = "MS-Acess/Jet"
''''    Else
''''        '*** ALL Other DB's
''''        '*** Do NOT make it Case-Insensitive... it KILLS Performance on Some DB's
''''        '    like Sentai's Progress DB
''''        strDataBaseType = "OTHER"
''''    End If
''''
''''
''''    AdodcLookupList(intSiteIdIndex).Recordset.Close
''''    AdodcLookupList(intSiteIdIndex).Recordset.Open
    
    'CLEAR the Where variable
    strFieldListForWhere = ""
    
    'Fix Error if Single Quotes are used in Lookup value
    'because SQL-Server will return Runtime Error "Incorrect Syntax".
    ' Replace Single Quotes with TWO Single Quotes
    txtTableLookupField = Replace(txtTableLookupField, "'", "''")
    
    
    'DEFAULT to Site A
    intSiteIdIndex = 0

    If strApplicationCommitBatchTo = "TTC" Then
        If IsNumeric(txtTableLookupField) Then
            If CDbl(txtTableLookupField) >= dblCaseIdCutoff Then
                intSiteIdIndex = 1
                grdDataGrid.BackColor = vbYellow
            Else
                grdDataGrid.BackColor = vbWhite
            End If
        End If
    End If

    Set arrAdodcLookupListRS(intSiteIdIndex) = New ADODB.Recordset

    
    For intIndex = 1 To UBound(arrApplicationFieldNameForInput)
        
        If UCase(arrAllowLookup(intIndex)) = vbChecked Then
            
                '*** 2008-11-19 - Jacob: Added to allow User-defined Search Conditions
                Dim strFieldSearchCondition As String
                Dim strFieldWildCardBegin As String
                Dim strFieldWildCardEnd As String
                
                Select Case arrFieldSearchCondition(intIndex)
                    Case "Contains"
                        strFieldSearchCondition = "LIKE"
                        strFieldWildCardBegin = "%"
                        strFieldWildCardEnd = "%"
                    Case "Begins With"
                        strFieldSearchCondition = "LIKE"
                        strFieldWildCardBegin = ""
                        strFieldWildCardEnd = "%"
                    Case Else
                        strFieldSearchCondition = arrFieldSearchCondition(intIndex)
                        strFieldWildCardBegin = ""
                        strFieldWildCardEnd = ""
                
                End Select
            
            
            'Prepare DISPLAY FIELDS portion of SELECT Statement
            If strFieldListForWhere = "" Then
                'Set up First Field
                '*** JR 6/28/2005 - Added the UPPER and UCASE features to be able to FIND
                '                   regardless of the CASE.
                '*** JR 2/16/2006 - Added check for MS ACCESS DB which uses UCASE instead of UPPER
                '*** JR 9/5/2006 -  Changed IF to Handle the Sentai/Progress DB
                Select Case strDataBaseType(intSiteIdIndex)
                    Case "SQL-Server/MSDE"
                        '*** SQL-Server/MSDE
                        strFieldListForWhere = "UPPER(" & arrLookupTableFieldName(intIndex) & ") " & strFieldSearchCondition & " '" & strFieldWildCardBegin & UCase(Trim(txtTableLookupField)) & strFieldWildCardEnd & "'"
                    Case "MS-Acess/Jet"
                        '*** MS-Acess/Jet DB's
                        strFieldListForWhere = "UCASE(" & arrLookupTableFieldName(intIndex) & ") " & strFieldSearchCondition & " '" & strFieldWildCardBegin & UCase(Trim(txtTableLookupField)) & strFieldWildCardEnd & "'"
                    Case Else
                        '*** ALL Other DB's
                        '*** Do NOT make it Case-Insensitive... it KILLS Performance on Some DB's
                        '    like Sentai's Progress DB
                        '*** 1/31/2007 Jacob - Added Field and Test for Numeric.
                        '                      Progress DB cannot handle LIKE with numerics.
                        If arrTreatAsNumeric(intIndex) = "" Then
                            arrTreatAsNumeric(intIndex) = 0
                        End If
                        
                        If arrTreatAsNumeric(intIndex) = True Then
                            If IsNumeric(Trim(txtTableLookupField)) Then
                                strFieldListForWhere = arrLookupTableFieldName(intIndex) & " = " & Trim(txtTableLookupField)
                            Else
'                                MsgBox "Sorry... search value for (" & arrLookupTableFieldName(intIndex) & ") MUST be Numeric!", vbInformation, "Invalid Search Value"
                            End If
                        Else
                            strFieldListForWhere = arrLookupTableFieldName(intIndex) & " " & strFieldSearchCondition & " '" & strFieldWildCardBegin & Trim(txtTableLookupField) & strFieldWildCardEnd & "'"
                        End If
                End Select
                
            Else
            
                'Insert commas before additional fields
                '*** JR 6/28/2005 - Added the UPPER and UCASE features to be able to FIND
                '                   regardless of the CASE.
                Select Case strDataBaseType(intSiteIdIndex)
                    Case "SQL-Server/MSDE"
                        '*** SQL-Server/MSDE
                        strFieldListForWhere = strFieldListForWhere & " OR " & "UPPER(" & arrLookupTableFieldName(intIndex) & ")" & " LIKE '" & UCase(Trim(txtTableLookupField)) & "%'"
                    Case "MS-Acess/Jet"
                        '*** MS-Acess/Jet DB's
                        strFieldListForWhere = strFieldListForWhere & " OR " & "UCASE(" & arrLookupTableFieldName(intIndex) & ")" & " LIKE '" & UCase(Trim(txtTableLookupField)) & "%'"
                    Case Else
                        '*** ALL Other DB's
                        
                        If arrTreatAsNumeric(intIndex) = "" Then
                            arrTreatAsNumeric(intIndex) = 0
                        End If

                        '*** Do NOT make it Case-Insensitive... it KILLS Performance on Some DB's
                        '    like Sentai's Progress DB
                            '*** Do NOT make it Case-Insensitive... it KILLS Performance
                        '*** 1/31/2007 Jacob - Added Field and Test for Numeric.
                        '                      Progress DB cannot handle LIKE with numerics.
                        If arrTreatAsNumeric(intIndex) = True Then
                            If IsNumeric(Trim(txtTableLookupField)) Then
                                strFieldListForWhere = strFieldListForWhere & " OR " & arrLookupTableFieldName(intIndex) & " = " & Trim(txtTableLookupField)
                            Else
'                                MsgBox "Sorry... search value for (" & arrLookupTableFieldName(intIndex) & ") MUST be Numeric!", vbInformation, "Invalid Search Value"
                            End If
                        Else
'                            strFieldListForWhere = strFieldListForWhere & " OR " & arrLookupTableFieldName(intIndex) & " LIKE '" & Trim(txtTableLookupField) & "%'"
                            strFieldListForWhere = strFieldListForWhere & " OR " & arrLookupTableFieldName(intIndex) & " " & strFieldSearchCondition & " '" & strFieldWildCardBegin & Trim(txtTableLookupField) & strFieldWildCardEnd & "'"
                        End If
                End Select
                
            End If
        End If
 
    Next
    
    '*** JR 2/1/2007 - Added if NO Where parameters, skip search
    If Trim(strFieldListForWhere) <> "" Then
    
        
        '*** If a Default Application WHERE Clause has been set, add it now
        If Trim(arrLookupDBWhereClause(intSiteIdIndex)) <> "" Then
            If strFieldListForWhere = "" Then
                strFieldListForWhere = Trim(arrLookupDBWhereClause(intSiteIdIndex))
            Else
                strFieldListForWhere = Trim(arrLookupDBWhereClause(intSiteIdIndex)) & " AND ( " & strFieldListForWhere & " )"
            End If
        End If
    
        strFieldListForWhere = Replace(strFieldListForWhere, "*", "%")
        
        
        '*** JR 12/20/2004 - ADDED "SET CONCAT_NULL_YIELDS_NULL OFF;" TO MAKE SURE WE GET A VALUE
        '                    IF ONE OF THE FIELDS IS NULL
        '*** JR 6/27/2005 - Added the LookupDBTableIsOnSQLServer field and IF to allow for
        '                    DB's that don't support the CONCAT_NULL... statement.
        '*** JR 9/5/2006  - Changed logic to match the WHERE clause above and automatically detect DB Type
    '    If RegLookupDBTableIsOnSQLServer = vbChecked Then
        Select Case strDataBaseType(intSiteIdIndex)
            Case "SQL-Server/MSDE"
                arrAdodcLookupListRS(intSiteIdIndex).Source = "SET CONCAT_NULL_YIELDS_NULL OFF; " & _
                                                " SELECT DISTINCT " & strFieldListForSelect & _
                                                " FROM " & arrLookupDBTableName(intSiteIdIndex) & _
                                                " WHERE " & strFieldListForWhere & _
                                                " ORDER BY " & strFieldListForOrderBy
            Case "MS-Acess/Jet"
                arrAdodcLookupListRS(intSiteIdIndex).Source = "SELECT DISTINCT " & strFieldListForSelect & _
                                                " FROM " & arrLookupDBTableName(intSiteIdIndex) & _
                                                " WHERE " & strFieldListForWhere & _
                                                " ORDER BY " & strFieldListForOrderBy
            Case Else
                '*** NO "ORDER BY"
                arrAdodcLookupListRS(intSiteIdIndex).Source = "SELECT DISTINCT " & strFieldListForSelect & _
                                                " FROM " & arrLookupDBTableName(intSiteIdIndex) & _
                                                " WHERE " & strFieldListForWhere
            End Select
        
        
        funcWriteToDebugLog Me.name, ""
        funcWriteToDebugLog Me.name, arrAdodcLookupListRS(intSiteIdIndex).Source
        funcWriteToDebugLog Me.name, ""
        Debug.Print arrAdodcLookupListRS(intSiteIdIndex).Source
        
        
        '********************************************************
        '*** REFRESH THE DATAGRID
        arrAdodcLookupListRS(intSiteIdIndex).CursorLocation = adUseClient
        arrAdodcLookupListRS(intSiteIdIndex).LOCKTYPE = adLockReadOnly
        arrAdodcLookupListRS(intSiteIdIndex).CursorType = adOpenStatic
        arrAdodcLookupListRS(intSiteIdIndex).ActiveConnection = arrAdodcLookupList(intSiteIdIndex)
        arrAdodcLookupListRS(intSiteIdIndex).Open
        
        Set grdDataGrid.DataSource = arrAdodcLookupListRS(intSiteIdIndex).DataSource
        
'        AdodcLookupList(intSiteIdIndex).Refresh

        grdDataGrid.Visible = True
        DoEvents
        
        '********************************************************
        '*** MAKE SURE ALL COLUMNS ARE LOCKED - NOT EDITABLE
        For intColumn = 0 To grdDataGrid.Columns.Count - 1
            grdDataGrid.Columns.item(intColumn).Locked = False
        Next
        
        '********************************************************
        '*** SET WIDTH OF COLUMNS TO FIT DATA
        If arrAdodcLookupListRS(intSiteIdIndex).RecordCount > 0 Then
            bolWalkingDownDataGrid = True
            arrAdodcLookupListRS(intSiteIdIndex).MoveFirst
            'Walk down the Rows in the Recordset
            While Not arrAdodcLookupListRS(intSiteIdIndex).EOF

                    'Walk across Colums in the grdDataGrid
                    For i = 0 To (grdDataGrid.Columns.Count - 1)
                    
                        X = (Len(arrAdodcLookupListRS(intSiteIdIndex).Fields(i).Value) * 128) + 60
                        If X > grdDataGrid.Columns.item(i).width Then
                            grdDataGrid.Columns.item(i).width = X
                        End If
                    
                    Next
                    
                    arrAdodcLookupListRS(intSiteIdIndex).MoveNext

            Wend
            bolWalkingDownDataGrid = False
            arrAdodcLookupListRS(intSiteIdIndex).MoveFirst

        End If
        
        

        '********************************************************
        '*** If we only found ONE record... Automatically copy the
        '*** values to the appropriate indexing fields and set focus
        '*** to the frmIndex
        '*** Otherwise, set focus to the Accountnumber again
        If arrAdodcLookupListRS(intSiteIdIndex).RecordCount = 1 Then
            
            subGetDataGridCellValues
            
            If bolAIM_Command_AddFile Then
                frmImaging101Search.SetFocus
            ElseIf bolBatchScanningModule = True Then
                Imaging101ScanMainPix.txtBatchName.SetFocus
            Else
                frmIndex.SetFocus
            End If
            
        Else
            txtTableLookupField.SetFocus
        End If
    
    End If 'Trim(strFieldListForWhere) <> ""
    
    '********************************************************
    '*** POPULATE THE LISTVIEW
'''    subListViewPopulate

    Screen.MousePointer = vbNormal
    

Exit Sub
    
ERROR_HANDLER:
        
    Screen.MousePointer = vbNormal

    ' Err 91 = List is empty
    If Err.Number = 91 Then
        Resume Next
    End If
    funcQuickMessage "SHOW", "Lookup Field Change - Error: " & Err.Number & " - " & Err.Description & " [" & Err.Source & "] - (" & RegLookupListConnectionString & ") - (" & arrAdodcLookupListRS(intSiteIdIndex).Source & ")"
    
End Sub



Private Sub cmdNextImage_Click()

    frmIndex.cmdNextImage_Click

End Sub

Private Sub Form_Activate()


    bolFormLoaded = True
'    txtTableLookupField.SetFocus

    'If NO Lookup Fields are Available - Unload the Lookup List Form
    If gNoLookupFieldsAvailable = True Then
        Unload Me
    End If
    
    If bolAIM_Command_AddFile = True Or bolBatchScanningModule = True Then
        cmdNextImage.Visible = False
    Else
        cmdNextImage.Visible = True
    End If
    
End Sub



Private Sub Form_Load()

    'If Batch is in ReadOnly Mode... Get Out of Here
    If gOpenBatchInReadOnlyMode = True Then
        Exit Sub
    End If


    bolFormLoaded = False
    
    'Default to Site A
    intSiteIdIndex = 0
    
    ' Get Position settings from the registry
    On Error Resume Next
    Me.Top = VBGetPrivateProfileString(RegAppname, "frmLookupList.Top", RegFileName)
    Me.Left = VBGetPrivateProfileString(RegAppname, "frmLookupList.Left", RegFileName)
    Me.width = VBGetPrivateProfileString(RegAppname, "frmLookupList.Width", RegFileName)
    Me.Height = VBGetPrivateProfileString(RegAppname, "frmLookupList.Height", RegFileName)
    
    '*** Check if Outside Screen Viewing Area
    Dim screenwidth, screenheight As Single

    screenwidth = Screen.width \ Screen.TwipsPerPixelX
    screenheight = Screen.Height \ Screen.TwipsPerPixelY
    
    If Screen.Height < Me.Top + Me.Height + 400 Then
        Me.Top = Screen.Height - Me.Height - 400
    End If
    
'    Me.Caption = VBGetPrivateProfileString(RegAppname, "frmLookupList.Caption", RegFileName)
    
    If Me.Caption = "" Then
        Me.Caption = "DB Lookup List"
    End If
    
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
    
    Me.chkHighlightLookpFieldAfterNextPage = VBGetPrivateProfileString(RegAppname, "frmLookupList.chkHighlightLookpFieldAfterNextPage", RegFileName)

    On Error GoTo ERROR_HANDLER
    
    'Make the DataGrid INVISIBLE while we get the I101TableLookupFields
    grdDataGrid.Visible = False
    
    
    '*** Check if running INDEXING or I101AIM Command
    Dim dblApplicationRECID As Double
    
    If bolAIM_Command_AddFile = True Then
        dblApplicationRECID = frmImaging101Search.txtApplicationRECID
    ElseIf bolBatchScanningModule = True Then
        dblApplicationRECID = Imaging101ScanMainPix.txtApplicationRECID
    Else
        dblApplicationRECID = frmIndex.txtApplicationRECID
    End If
    
    
    '**************************************************************************************
    '*** Get the Lookup Table DB Connection String for the current Application
    
        Dim conn As ADODB.Connection
        Dim rs As ADODB.Recordset
        Set conn = New ADODB.Connection
        Set rs = New ADODB.Recordset
        
        conn.ConnectionString = RegImaging101ConnectionString
        conn.ConnectionTimeout = 120
        conn.mode = adModeRead
        conn.Open
        
        ssql = "SELECT * FROM I101Applications WHERE ApplicationRECID = " & dblApplicationRECID
        
        With rs
            .ActiveConnection = conn
            .CursorLocation = adUseServer
            .CursorType = adOpenDynamic
            .LOCKTYPE = adLockReadOnly
            .Source = ssql
        End With

        rs.Open
        
        strApplicationCommitBatchTo = rs.Fields!ApplicationCommitBatchTo & ""
        

        'Hold values for "Site A"
        arrLookupDBConnectionString(0) = rs.Fields!LookupDBConnectionString & ""
        arrLookupDBTableName(0) = rs.Fields!LookupDBTableName & ""
        arrLookupDBTableIsOnSQLServer(0) = rs.Fields!LookupDBTableIsOnSQLServer & ""
        arrLookupDBWhereClause(0) = rs.Fields!LookupDBWhereClause & ""

        'Hold values for "Site B"
        arrLookupDBConnectionString(1) = rs.Fields!LookupDBConnectionString_B & ""
        arrLookupDBTableName(1) = rs.Fields!LookupDBTableName_B & ""
        arrLookupDBTableIsOnSQLServer(1) = rs.Fields!LookupDBTableIsOnSQLServer_B & ""
        arrLookupDBWhereClause(1) = rs.Fields!LookupDBWhereClause_B & ""
        
        dblCaseIdCutoff = CDbl(rs.Fields!CaseIdCutoff & "")
        
        RegApplicationBatchNameDelimiter = rs.Fields!ApplicationBatchNameDelimiter & ""
        
        strFieldToSelectAfterLookupClick = rs.Fields!FieldToSelectAfterLookupClick & ""
        strFieldToSelectAfterDocListClick = rs.Fields!FieldToSelectAfterDocListClick & ""
    
        bolAutoLookupOnBatchLoad = rs.Fields!AutoLookupOnBatchLoad & ""
        
        rs.Close
        conn.Close
        Set rs = Nothing
        Set conn = Nothing

    

    ' Load and Set up the SELECT Statement Variables for the TABLE LOOKUP Fields
    AdodcLookupList(0).ConnectionString = RegImaging101ConnectionString
    AdodcLookupList(0).RecordSource = "SELECT * FROM I101TableLookupFields WHERE ApplicationRECID = " & dblApplicationRECID & " ORDER BY DisplayOrder"
    AdodcLookupList(0).Refresh
    
    If Not bolAIM_Command_AddFile Then
    
        Dim intDelimiterLocation As Integer

        If bolBatchScanningModule = True Then
        
            'Set BatchName on Table Lookup field
            If InStr(Imaging101ScanMainPix.txtBatchName.Text, RegApplicationBatchNameDelimiter) Then
                intDelimiterLocation = InStr(Imaging101ScanMainPix.txtBatchName, RegApplicationBatchNameDelimiter)
                frmLookupList.txtTableLookupField = Right(Imaging101ScanMainPix.txtBatchName, Len(Imaging101ScanMainPix.txtBatchName) - intDelimiterLocation)
        '                frmLookupList.txtTableLookupField = Mid(frmIndex.txtBatchName, InStr(1, frmIndex.txtBatchName, RegApplicationBatchNameDelimiter) + 1, Len(frmIndex.txtBatchName) - InStr(1, frmIndex.txtBatchName, RegApplicationBatchNameDelimiter))
            Else
                frmLookupList.txtTableLookupField = Imaging101ScanMainPix.txtBatchName
            End If
        
        Else
    
            'Set BatchName on Table Lookup field
            If InStr(frmIndex.txtBatchName.Text, RegApplicationBatchNameDelimiter) Then
                intDelimiterLocation = InStr(frmIndex.txtBatchName, RegApplicationBatchNameDelimiter)
                frmLookupList.txtTableLookupField = Right(frmIndex.txtBatchName, Len(frmIndex.txtBatchName) - intDelimiterLocation)
        '                frmLookupList.txtTableLookupField = Mid(frmIndex.txtBatchName, InStr(1, frmIndex.txtBatchName, RegApplicationBatchNameDelimiter) + 1, Len(frmIndex.txtBatchName) - InStr(1, frmIndex.txtBatchName, RegApplicationBatchNameDelimiter))
            Else
                frmLookupList.txtTableLookupField = frmIndex.txtBatchName
            End If
        
        End If
        
    End If
    

    
    
    'If NO Lookup Fields or Table are Available
    ' OR if the bolAutoLookupOnBatchLoad is OFF
    ' SET a Global Flag to PREVENT doing a Table Lookup
    If AdodcLookupList(0).Recordset.EOF _
    Or arrLookupDBConnectionString(0) = "" Then
        gNoLookupFieldsAvailable = True
        Exit Sub
    Else
        gNoLookupFieldsAvailable = False
    End If
    
    'Initialize variables
    strFieldListForSelect = ""
    strFieldListForWhere = ""
    strFieldListForOrderBy = ""
    
    Dim AdodcLookupListRecordCount As Integer
    AdodcLookupListRecordCount = AdodcLookupList(0).Recordset.RecordCount
    
    For intIndex = 1 To AdodcLookupListRecordCount
        
         '*** 2021-05-21 - Jacob - Added arrColumnTitle()
        ReDim Preserve arrApplicationFieldNameForInput(intIndex)
        ReDim Preserve arrLookupTableFieldName(intIndex)
        ReDim Preserve arrColumnTitle(intIndex)
        ReDim Preserve arrShowColumn(intIndex)
        ReDim Preserve arrAllowLookup(intIndex)
        ReDim Preserve arrTrimFieldBeginChar(intIndex)
        ReDim Preserve arrTrimFieldLength(intIndex)
        ReDim Preserve arrTreatAsNumeric(intIndex)
        ReDim Preserve arrFieldSearchCondition(intIndex) As String

        'ADD The Fields to the Array for Populating the Index Fields later as user Enters search value
        arrLookupTableFieldName(intIndex) = AdodcLookupList(intSiteIdIndex).Recordset.Fields("LookupTableFieldName") & ""
        arrApplicationFieldNameForInput(intIndex) = AdodcLookupList(intSiteIdIndex).Recordset.Fields("ApplicationFieldNameForInput") & ""
        arrShowColumn(intIndex) = AdodcLookupList(intSiteIdIndex).Recordset.Fields("ShowColumn") & ""
        arrColumnTitle(intIndex) = AdodcLookupList(intSiteIdIndex).Recordset.Fields("ColumnTitle") & ""
        arrAllowLookup(intIndex) = AdodcLookupList(intSiteIdIndex).Recordset.Fields("AllowLookup") & ""
        arrTrimFieldBeginChar(intIndex) = AdodcLookupList(intSiteIdIndex).Recordset.Fields("TrimFieldBeginChar") & ""
        arrTrimFieldLength(intIndex) = AdodcLookupList(intSiteIdIndex).Recordset.Fields("TrimFieldLength") & ""
        arrTreatAsNumeric(intIndex) = AdodcLookupList(intSiteIdIndex).Recordset.Fields("TreatAsNumeric") & ""
        arrFieldSearchCondition(intIndex) = AdodcLookupList(intSiteIdIndex).Recordset.Fields("FieldSearchCondition") & ""
        
        If UCase(AdodcLookupList(intSiteIdIndex).Recordset.Fields("ShowColumn")) = vbChecked Then
            'Prepare DISPLAY FIELDS portion of SELECT Statement
            If strFieldListForSelect = "" Then
                'Set up First Field
                strFieldListForSelect = arrLookupTableFieldName(intIndex) & " AS " & AdodcLookupList(intSiteIdIndex).Recordset.Fields("ColumnTitle")
                'Simply use the SAME string as the Field List -- MAY Change Later
                strFieldListForOrderBy = arrLookupTableFieldName(intIndex)
            Else
                'Insert commas before additional fields
                strFieldListForSelect = strFieldListForSelect & ", " & arrLookupTableFieldName(intIndex) & " AS " & AdodcLookupList(intSiteIdIndex).Recordset.Fields("ColumnTitle")
                'Simply use the SAME string as the Field List -- MAY Change Later
                strFieldListForOrderBy = strFieldListForOrderBy & ", " & arrLookupTableFieldName(intIndex)
            End If
        End If
 
        
        'Get the next Field record
        AdodcLookupList(intSiteIdIndex).Recordset.MoveNext
    Next
    
  
    
    ' Set the Connection String to the Lookup Table
    
    For intIndex = 0 To 1
        
        intSiteIdIndex = intIndex
        
        If Trim(arrLookupDBConnectionString(intSiteIdIndex)) <> "" Then
            'For TTC, set the "Site B" connection string
            Set arrAdodcLookupList(intSiteIdIndex) = New ADODB.Connection
            arrAdodcLookupList(intSiteIdIndex).ConnectionString = arrLookupDBConnectionString(intSiteIdIndex)
            arrAdodcLookupList(intSiteIdIndex).Open
'            arrAdodcLookupList(intSiteIdIndex).Recordset.Close
'            arrAdodcLookupList(intSiteIdIndex).Recordset.Open
        End If
    
        '****************************************
        '*** Determine the Database Type
        If InStr(UCase(arrLookupDBConnectionString(intSiteIdIndex)), "SQLOLEDB.1") Then
            '*** SQL-Server/MSDE
            strDataBaseType(intSiteIdIndex) = "SQL-Server/MSDE"
        ElseIf InStr(UCase(arrLookupDBConnectionString(intSiteIdIndex)), "MICROSOFT.JET") Then
            '*** MS-Acess/Jet DB's
            strDataBaseType(intSiteIdIndex) = "MS-Acess/Jet"
        Else
            '*** ALL Other DB's
            '*** Do NOT make it Case-Insensitive... it KILLS Performance on Some DB's
            '    like Sentai's Progress DB
            strDataBaseType(intSiteIdIndex) = "OTHER"
        End If
        
    Next

    Me.Show
    DoEvents
    
    
    
    
    
    
    '************************************************************
    '*** If AutoLookupOnBatchLoad is ON, Click the FIND Button
    
    If bolAutoLookupOnBatchLoad Then
        cmdFind_Click
    End If

Exit Sub
    
ERROR_HANDLER:
    ' Err 91 = List is empty
    If Err.Number = 91 Then
        Resume Next
    End If
    MsgBox "Form Load - Error: " & Err.Number & " - " & Err.Description & " [" & Err.Source & "] - (" & RegLookupListConnectionString & ") - (" & AdodcLookupList(intSiteIdIndex).RecordSource & ")", vbExclamation
    

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    'Do NOT allow this form to be unloaded while it's still loading.
    If funcIsFormLoaded2("frmIndex") And (Not bolIndexFormLoadComplete) Then
        Cancel = True
        Exit Sub
    End If
    
        If UnloadMode = vbUser Then
            Cancel = True
            Exit Sub
        End If
        
        'Save Form settings to the registry, except if minimized (i.e.- Top or Left < 0)
        If Me.Top >= 0 And Me.Left >= 0 Then
            result = WritePrivateProfileString(RegAppname, "frmLookupList.Top", Me.Top, RegFileName)
            result = WritePrivateProfileString(RegAppname, "frmLookupList.Left", Me.Left, RegFileName)
            result = WritePrivateProfileString(RegAppname, "frmLookupList.Width", Me.width, RegFileName)
            result = WritePrivateProfileString(RegAppname, "frmLookupList.Height", Me.Height, RegFileName)
'            result = WritePrivateProfileString(RegAppname, "frmLookupList.Caption", Me.Caption, RegFileName)
        End If
    
    result = WritePrivateProfileString(RegAppname, "frmLookupList.chkHighlightLookpFieldAfterNextPage", frmLookupList.chkHighlightLookpFieldAfterNextPage, RegFileName)

    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'IGNORE Any Errors
    On Error Resume Next
    
    For intSiteIdIndex = 0 To 1
        arrAdodcLookupList(intSiteIdIndex).Close
        Set arrAdodcLookupList(intSiteIdIndex) = Nothing
    Next
    
    Screen.MousePointer = vbDefault
    
End Sub
Public Sub MakeTopMost()
    ' Make this window the top most window so that it appears in
    ' front of the ecIndex window
    Dim lReturn As Long
    
    ' The SetWindowPos API call will do this.
    lReturn = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
End Sub

Private Sub Form_Resize()
  On Error Resume Next
    If frmLookupList.WindowState <> vbMinimized Then
          'This will resize the grid when the form is resized
          grdDataGrid.width = Me.ScaleWidth
        '  grdDataGrid.Height = Me.ScaleHeight - grdDataGrid.Top - cmbArea.Height - Toolbar1.Height - 5
          grdDataGrid.Height = Me.ScaleHeight - grdDataGrid.Top - 20
          grdDataGrid.Splits(0).Columns(0).width = 2000
          grdDataGrid.Splits(0).Columns(1).width = 2000
    End If
End Sub




Private Sub grdDataGrid_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)

    'Never enters here

End Sub

Private Sub grdDataGrid_ButtonClick(ByVal ColIndex As Integer)

    'Never enters here
    
End Sub

Private Sub grdDataGrid_Click()

'        subGetDataGridCellValues
'        frmIndex.SetFocus

End Sub

Private Sub grdDataGrid_DblClick()

        subGetDataGridCellValues
        
        If bolAIM_Command_AddFile Then
            frmImaging101Search.SetFocus
        ElseIf bolBatchScanningModule = True Then
            'Scanning
            Imaging101ScanMainPix.SetFocus
        Else
            frmIndex.SetFocus
        End If
        
End Sub

Private Sub grdDataGrid_KeyPress(KeyAscii As Integer)
    'Catch Enter key
    If KeyAscii = 13 Then
        subGetDataGridCellValues
'        grdDataGrid.SetFocus
        'Setting KeyAscii = 0 will prevent the grdDataGrid from STEALING the Focus Back
        KeyAscii = 0
        
        If bolAIM_Command_AddFile Then
            frmImaging101Search.SetFocus
        ElseIf bolBatchScanningModule = True Then
            'Scanning
            Imaging101ScanMainPix.SetFocus
        Else
            frmIndex.SetFocus
        End If
    End If
    
    If KeyAscii = Asc("[") And frmIndex.Visible = True Then
        If bolAIM_Command_AddFile Then
            frmImaging101Search.SetFocus
        Else
            frmIndex.SetFocus
        End If

        'Cancel the Keypress by setting the KeyAscii to Zero (0)
        KeyAscii = 0
    End If
    
End Sub





Private Sub grdDataGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    ' *** This procedure will send all the index info to the proper Indexing Form
    ' *** based on the field names and values entered.
    ' *** NOTE:  The field names are case sensitive!
    
    
    'GET OUT of Here if the Form is not completelly loaded OR there was no previous row.
    '2011-08-29 - Jacob - Added "LastRow = Null" check
    '2012-04-20 - Jacob - Added "LastRow = Empty" check
    If LastRow = Null Or LastCol = 0 Or Not bolFormLoaded Or IsNull(LastRow) Or bolWalkingDownDataGrid Or LastRow = Empty Then
        Exit Sub
    End If

    subGetDataGridCellValues
    
        If bolAIM_Command_AddFile Then
            frmImaging101Search.SetFocus
        ElseIf bolBatchScanningModule = True Then
            Imaging101ScanMainPix.SetFocus
        Else
            frmIndex.SetFocus
        End If


    
End Sub




Private Sub txtTableLookupField_Change()

    
'    If Trim(txtTableLookupField) = "" Then
'        ' Get Empty Result Set
'        Select Case gsecSiteInformationClientShort
'            Case "TTC"
'                AdodcLookupList(intSiteIdIndex).RecordSource = "select " & strAccountNumberDbField & ", " & strAccountNameDbField & ", " & strAccountCaseDbField & ", " & strAccountCaseIdDbField & ", " & strAccountTicketNumberDbField & ", DateStamp from " & strAccountTableName & " Where 0=1  Order by " & strAccountNumberDbField
'            Case "JMH"
'                AdodcLookupList(intSiteIdIndex).RecordSource = "select " & strAccountNumberDbField & ", " & strAccountNameDbField & ", DateStamp from " & strAccountTableName & " Where 0=1  Order by " & strAccountNumberDbField
'            Case "HA"
'                AdodcLookupList(intSiteIdIndex).RecordSource = "select " & strAccountNumberDbField & ", " & strAccountNameDbField & ", " & strAccountCaseDbField & ", " & strAccountCaseIdDbField & ", " & strAccountTicketNumberDbField & " from " & strAccountTableName & " Where 0=1  Order by " & strAccountNumberDbField
'            Case "WASD"
'                AdodcLookupList(intSiteIdIndex).RecordSource = "select " & strAccountNumberDbField & ", " & strAccountNameDbField & " FROM " & strAccountTableName & " WHERE 0=1  ORDER BY " & strAccountNumberDbField
'            Case "CCCS"
'                AdodcLookupList(intSiteIdIndex).RecordSource = "select " & strAccountNumberDbField & ", " & strAccountNameDbField & ", " & strAccountCaseDbField & ", " & strAccountCaseIdDbField & ", " & strAccountTicketNumberDbField & " from " & strAccountTableName & " Where 0=1  Order by " & strAccountNumberDbField
'            Case Else
'                AdodcLookupList(intSiteIdIndex).RecordSource = "select * FROM " & strAccountTableName & " WHERE 0=1  ORDER BY " & strAccountNumberDbField
'        End Select
'        AdodcLookupList(intSiteIdIndex).Refresh
'    ElseIf Len(Trim(txtTableLookupField)) >= strAccountNumberSearchChars Then
'        ' Find Patient Records
'        Select Case gsecSiteInformationClientShort
'            Case "TTC"
'                AdodcLookupList(intSiteIdIndex).RecordSource = "select " & strAccountNumberDbField & ", " & strAccountNameDbField & ", " & strAccountCaseDbField & ", " & strAccountCaseIdDbField & ", " & strAccountTicketNumberDbField & ", DateStamp from " & strAccountTableName & " Where " & strAccountNumberDbField & " LIKE '" & Trim(txtTableLookupField) & "%'" & " OR " & strAccountNameDbField & " LIKE '" & Trim(txtTableLookupField) & "%'" & " OR " & strAccountCaseDbField & " LIKE '" & Trim(txtTableLookupField) & "%'" & " OR " & strAccountTicketNumberDbField & " LIKE '" & Trim(txtTableLookupField) & "%'" & " ORDER BY " & strAccountNumberDbField
'            Case "JMH"
'                AdodcLookupList(intSiteIdIndex).RecordSource = "select " & strAccountNumberDbField & ", " & strAccountNameDbField & ", DateStamp from " & strAccountTableName & " Where " & strAccountNumberDbField & " LIKE '" & Trim(txtTableLookupField) & "%'" & " OR " & strAccountNameDbField & " LIKE '" & Trim(txtTableLookupField) & "%'" & " ORDER BY " & strAccountNumberDbField
'            Case "HA"
'                AdodcLookupList(intSiteIdIndex).RecordSource = "select " & strAccountNumberDbField & ", " & strAccountNameDbField & ", " & strAccountCaseDbField & ", " & strAccountCaseIdDbField & ", " & strAccountTicketNumberDbField & " from " & strAccountTableName & " Where " & strAccountNumberDbField & " LIKE '" & Trim(txtTableLookupField) & "%'" & " OR " & strAccountNameDbField & " LIKE '" & Trim(txtTableLookupField) & "%'" & " OR " & strAccountCaseDbField & " LIKE ' " & Trim(txtTableLookupField) & "%' OR " & strAccountCaseIdDbField & " LIKE '" & Trim(txtTableLookupField) & "%' OR " & strAccountTicketNumberDbField & " LIKE '" & Trim(txtTableLookupField) & "%'" & " ORDER BY " & strAccountNumberDbField
'            Case "WASD"
'                AdodcLookupList(intSiteIdIndex).RecordSource = "select " & strAccountNumberDbField & ", " & strAccountNameDbField & " FROM " & strAccountTableName & " WHERE " & strAccountNumberDbField & " LIKE '" & Trim(txtTableLookupField) & "%'" & " ORDER BY " & strAccountNumberDbField
'            Case "CCCS"
'                AdodcLookupList(intSiteIdIndex).RecordSource = "select " & strAccountNumberDbField & ", " & strAccountNameDbField & ", " & strAccountCaseDbField & ", " & strAccountCaseIdDbField & ", " & strAccountTicketNumberDbField & " from " & strAccountTableName & " Where " & strAccountNumberDbField & " LIKE '" & Trim(txtTableLookupField) & "%'" & " OR " & strAccountNameDbField & " LIKE '" & Trim(txtTableLookupField) & "%'" & " OR " & strAccountCaseDbField & " LIKE '" & Trim(txtTableLookupField) & "%'" & " OR " & strAccountTicketNumberDbField & " LIKE '" & Trim(txtTableLookupField) & "%'" & " ORDER BY " & strAccountNumberDbField
'            Case Else
'                AdodcLookupList(intSiteIdIndex).RecordSource = "select * FROM " & strAccountTableName & " Where " & strAccountNumberDbField & " LIKE '" & Trim(txtTableLookupField) & "%'" & " OR " & strAccountNameDbField & " LIKE '" & Trim(txtTableLookupField) & "%'" & " ORDER BY " & strAccountNumberDbField
'        End Select
'        AdodcLookupList(intSiteIdIndex).Refresh
'    End If

    'Initialize the Where variable

    
End Sub

Private Sub cmdClearFields_Click()
    Me.SetFocus
    txtTableLookupField = ""
    txtRowDescription = ""
    txtTableLookupField.SetFocus
End Sub




Private Sub txtTableLookupField_GotFocus()
    
'    txtTableLookupField.SelStart = 0
'    txtTableLookupField.SelLength = Len(txtTableLookupField)

    If frmLookupList.Visible = True Then
        frmLookupList.txtTableLookupField.SetFocus
        frmLookupList.txtTableLookupField.SelStart = 0
        frmLookupList.txtTableLookupField.SelLength = Len(frmLookupList.txtTableLookupField)
    End If

End Sub


Private Sub subGetDataGridCellValues()

    If bolAIM_Command_AddFile = True Then
        subAssignDataGRidCellValues frmImaging101Search
        
    ElseIf bolBatchScanningModule = True Then
        subAssignDataGRidCellValues Imaging101ScanMainPix
        
    Else
        subAssignDataGRidCellValues frmIndex
    End If
    
    
    
End Sub

Private Sub subAssignDataGRidCellValues(WorkingForm As Form)
      
    On Error Resume Next
    Err.Clear
    
    txtRowDescription = grdDataGrid.Splits(0).Columns(1).Text
    
    If Err.Number > 0 Then
        Exit Sub
    End If
    
    On Error GoTo ERROR_HANDLER
    
    '8/15/2017 - Jacob - Added Ability to do Table Lookup from Batch Scanning Module
    If bolBatchScanningModule = True Then
    
        WorkingForm.SetFocus
        WorkingForm.txtBatchName.SetFocus
        'Assign the FIRST defined Field Column value
        WorkingForm.txtBatchName.Text = grdDataGrid.Splits(0).Columns(0).Text
        WorkingForm.txtBatchDesc.Text = grdDataGrid.Splits(0).Columns(1).Text
        
        'Get out of here now
        Exit Sub
        
    End If
    
    Dim intMaxFieldLength As Integer
    
    
''MsgBox AdodcLookupList(intSiteIdIndex).Recordset.Fields.Item(0).Name & ", " & AdodcLookupList(intSiteIdIndex).Recordset.Fields.Item(0).Value
    ' Loop through the Fields, looking for a match
'    For intIndex = 0 To WorkingForm.lblFieldDescription.Count - 1
'        ' Set field default value
'        Select Case UCase(Trim(WorkingForm.lblFieldDescription.item(intIndex).Caption))
'            Case UCase(Trim(strAccountNumberEcField))
'                WorkingForm.mebIndexValues(intIndex).Text = Trim(Mid(grdDataGrid.Splits(0).Columns(0).Text, CDbl(strAccountNumberEcFieldBegin), CDbl(strAccountNumberEcFieldLength)))
'            Case UCase(Trim(strAccountNameEcField))
'                WorkingForm.mebIndexValues(intIndex).Text = Left(Trim(grdDataGrid.Splits(0).Columns(1).Text), WorkingForm.txtFieldSize(intIndex))
'            Case UCase(Trim(strAccountCaseEcField))
'                WorkingForm.mebIndexValues(intIndex).Text = Trim(grdDataGrid.Splits(0).Columns(2).Text)
'            Case UCase(Trim(strAccountCaseIdEcField))
'                WorkingForm.mebIndexValues(intIndex).Text = Trim(grdDataGrid.Splits(0).Columns(3).Text)
'            Case UCase(Trim(strAccountTicketNumberEcField))
'                WorkingForm.mebIndexValues(intIndex).Text = Trim(grdDataGrid.Splits(0).Columns(4).Text)
'        End Select
'    Next

    Dim intRow As Integer
    
    'Loop through the DEFINED Table Lookup Fields
    'The ARRAY arrApplicationFieldNameForInput() starts at 1
    For intRow = 1 To UBound(arrApplicationFieldNameForInput())
        
        If UCase(arrShowColumn(intRow)) = vbChecked Then
            
            ';LOOP through the Application Fields
            For intIndex = 0 To WorkingForm.lblFieldDescription.Count - 1
            
                'Set the Maximum Field Size
                If Int(WorkingForm.txtFieldSize(intIndex)) > 0 Then
                    intMaxFieldLength = WorkingForm.txtFieldSize(intIndex)
                Else
                    'max size of a Masked Edit box.
                    intMaxFieldLength = 64
                End If
    
                ' *** Set field Value for the Fields Defined in I101TableLookupFields...
                '      2014-01-29 - Jacob - Added check for txtFieldDefaultValue to ONLY set field if NO DEFAULT Value is Set.
                '      2014-04-28 - Jacob - Changed
                If UCase(Trim(WorkingForm.lblFieldDescription.item(intIndex).Caption)) = _
                    UCase(Trim(arrApplicationFieldNameForInput(intRow))) _
                    Then
                    
                    If Trim(WorkingForm.txtFieldTableLookupOverridesDefault(intIndex).Text) = "1" Then
                        
                        '*** 2021-05-21 - Jacob -   Changed to use  "arrColumnTitle(intRow)", instead of "intRow -1"
                        '                                              and Remove any Delimiters like Brackets or Single-quotes from the ColumnTitle field
                        arrColumnTitle(intRow) = Replace(arrColumnTitle(intRow), "[", "")
                        arrColumnTitle(intRow) = Replace(arrColumnTitle(intRow), "]", "")
                        arrColumnTitle(intRow) = Replace(arrColumnTitle(intRow), "'", "")
 _
                        '*** Check if the txtIndexValues TextBox control is VISIBLE...
                        '    this will handle saving the value of the the TextBox control
                        '    instead of the mebIndexValues Masked Edit control as needed.
                        '    This is because the Masked Edit control has a MAX size of 64 Char.
                        If WorkingForm.txtIndexValues(intIndex).Visible = True Then
                            
                            'Check if the TRIM Field values were set
                            If (arrTrimFieldBeginChar(intRow) <> "") And (arrTrimFieldLength(intRow) <> "") Then
                                WorkingForm.txtIndexValues(intIndex).Text = Left(Trim(Mid(grdDataGrid.Splits(0).Columns(arrColumnTitle(intRow)).Text, CDbl(arrTrimFieldBeginChar(intRow)), CDbl(arrTrimFieldLength(intRow)))), intMaxFieldLength)
                            Else
                                WorkingForm.txtIndexValues(intIndex).Text = Left(grdDataGrid.Splits(0).Columns(arrColumnTitle(intRow)).Text, intMaxFieldLength)
                            End If
                        
                        Else
                        
                            'Check if the TRIM Field values were set
                            If (arrTrimFieldBeginChar(intRow) <> "") And (arrTrimFieldLength(intRow) <> "") Then
                                WorkingForm.mebIndexValues(intIndex).Text = Left(Trim(Mid(grdDataGrid.Splits(0).Columns(arrColumnTitle(intRow)).Text, CDbl(arrTrimFieldBeginChar(intRow)), CDbl(arrTrimFieldLength(intRow)))), intMaxFieldLength)
                            Else
                                WorkingForm.mebIndexValues(intIndex).Text = Left(grdDataGrid.Splits(0).Columns(arrColumnTitle(intRow)).Text, intMaxFieldLength)
                            End If
                        
                        End If
                        
                    End If
                    
                    '2021-05-21 - Jacob - Since we found the field we were looking for, get out now.
                    Exit For
                    
                End If
                
            Next
            
        End If
        
    Next

    
    ' ***** Find the Field Specified in the INI file as
    ' *****   FIELDAFTERCLICK and set the focus to it.
    ' ***** NOTE:  The reason for this second loop pass is to make sure we get
    '              ALL the field values populated first.
    For intIndex = 0 To WorkingForm.lblFieldDescription.Count - 1
        Select Case UCase(Trim(WorkingForm.lblFieldDescription.item(intIndex).Caption))
            Case UCase(Trim(strFieldToSelectAfterLookupClick))
                WorkingForm.SetFocus
                
                '*** Check if the txtIndexValues TextBox control is VISIBLE...
                '    this will handle saving the value of the the TextBox control
                '    instead of the mebIndexValues Masked Edit control as needed.
                '    This is because the Masked Edit control has a MAX size of 64 Char.
                If WorkingForm.txtIndexValues(intIndex).Visible = True Then
                    WorkingForm.txtIndexValues(intIndex).SetFocus
                Else
                    WorkingForm.mebIndexValues(intIndex).SetFocus
                End If
                Exit For
        End Select
    Next
   
Exit Sub
    
ERROR_HANDLER:
    ' Err 91 = List is empty
    If Err.Number = 91 Or Err.Number = 6160 Then
        txtTableLookupField.SelStart = 0
        txtTableLookupField.SelLength = Len(txtTableLookupField)
        txtTableLookupField.SetFocus
        Resume
        Exit Sub
    End If
    MsgBox "Table Lookup grdDataGrid_Click - Error: " & Err.Number & " - " & Err.Description & " [" & Err.Source & "]", vbExclamation
    Exit Sub
    
    ' We must trap for error in case the user enters an invalid field name
       


End Sub

Private Sub txtTableLookupField_KeyPress(KeyAscii As Integer)

'    If KeyAscii = Asc("[") And funcIsFormLoaded2("frmIndex") = True Then
'        If bolAIM_Command_AddFile Then
'            frmImaging101Search.SetFocus
'        Else
'            frmIndex.SetFocus
'        End If
'
'        'Cancel the Keypress by setting the KeyAscii to Zero (0)
'        KeyAscii = 0
'    End If
'
'    If KeyAscii = Asc("]") And funcIsFormLoaded2("frmLookupList") = True Then
'        frmLookupList.SetFocus
'        'Cancel the Keypress by setting the KeyAscii to Zero (0)
'        KeyAscii = 0
'    End If
'
'    If KeyAscii = Asc("\") And funcIsFormLoaded2("MainMDIForm") = True Then
'        MainMDIForm.SetFocus
'        'Cancel the Keypress by setting the KeyAscii to Zero (0)
'        KeyAscii = 0
'    End If
    

End Sub

