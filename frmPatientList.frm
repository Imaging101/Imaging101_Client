VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPatientList 
   Caption         =   "Number-Name List"
   ClientHeight    =   2190
   ClientLeft      =   2985
   ClientTop       =   6225
   ClientWidth     =   5970
   Icon            =   "frmPatientList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2190
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPatientAccountNumber 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Width           =   1935
   End
   Begin VB.CommandButton cmdClearFields 
      Caption         =   "&Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   0
      Width           =   615
   End
   Begin VB.TextBox txtPatientAccountName 
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   3255
   End
   Begin MSDataGridLib.DataGrid grdDataGrid 
      Bindings        =   "frmPatientList.frx":0442
      Height          =   1695
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   2990
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdodcPatientList 
      Height          =   330
      Left            =   0
      Top             =   2280
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
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
      Caption         =   "AdodcPatientList"
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
End
Attribute VB_Name = "frmPatientList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'JR - 7/12/2001:  Added Support for AccountNumberBegin/intAccountNumberLength INI file parameters!
'                   convert the Strings to Longs using "clng()"
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


Private Const HWND_TOPMOST = -1
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long



Private Sub Form_Load()
    Dim RegConnectString As String
    Dim RegPatientListConnectionType As String
    
    ' Get Position settings from the registry
    On Error Resume Next
    Me.Top = VBGetPrivateProfileString(RegAppname, "frmPatientList.Top", RegFileName)
    Me.Left = VBGetPrivateProfileString(RegAppname, "frmPatientList.Left", RegFileName)
    Me.Width = VBGetPrivateProfileString(RegAppname, "frmPatientList.Width", RegFileName)
    Me.Height = VBGetPrivateProfileString(RegAppname, "frmPatientList.Height", RegFileName)
    Me.Caption = VBGetPrivateProfileString(RegAppname, "frmPatientList.Caption", RegFileName)
    If Me.Caption = "" Then Me.Caption = "Number-Name List"
    
    strAccountTableName = VBGetPrivateProfileString(RegAppname, "frmPatientList.AccountTableName", RegFileName)
    
    strAccountNumberEcField = VBGetPrivateProfileString(RegAppname, "frmPatientList.AccountNumberEcField", RegFileName)
    strAccountNumberDbField = VBGetPrivateProfileString(RegAppname, "frmPatientList.AccountNumberDbField", RegFileName)
    
    strAccountNameEcField = VBGetPrivateProfileString(RegAppname, "frmPatientList.AccountNameEcField", RegFileName)
    strAccountNameDbField = VBGetPrivateProfileString(RegAppname, "frmPatientList.AccountNameDbField", RegFileName)
    
    strAccountCaseEcField = VBGetPrivateProfileString(RegAppname, "frmPatientList.AccountCaseEcField", RegFileName)
    strAccountCaseDbField = VBGetPrivateProfileString(RegAppname, "frmPatientList.AccountCaseDbField", RegFileName)
    
    strAccountCaseIdEcField = VBGetPrivateProfileString(RegAppname, "frmPatientList.AccountCaseIdEcField", RegFileName)
    strAccountCaseIdDbField = VBGetPrivateProfileString(RegAppname, "frmPatientList.AccountCaseIdDbField", RegFileName)
    
    strAccountTicketNumberEcField = VBGetPrivateProfileString(RegAppname, "frmPatientList.AccountTicketNumberEcField", RegFileName)
    strAccountTicketNumberDbField = VBGetPrivateProfileString(RegAppname, "frmPatientList.AccountTicketNumberDbField", RegFileName)
    
    strAccountNumberEcFieldBegin = VBGetPrivateProfileString(RegAppname, "frmPatientList.AccountNumberEcFieldBegin", RegFileName)
    strAccountNumberEcFieldLength = VBGetPrivateProfileString(RegAppname, "frmPatientList.AccountNumberEcFieldLength", RegFileName)
    
    strAccountNumberSearchChars = VBGetPrivateProfileString(RegAppname, "frmPatientList.AccountNumberSearchChars", RegFileName)
    
    strFIELDAFTERCLICK = VBGetPrivateProfileString(RegAppname, "frmDocumentList.FIELDAFTERCLICK", RegFileName)
    On Error GoTo 0

'''''''    ' Get Database Connections settings from the registry
'''''''    On Error Resume Next
'''''''    RegPatientListConnectionType = VBGetPrivateProfileString("DATABASE", "AdodcPatientList.ConnectionType", RegFileName)
'''''''    RegPatientListConnectionString = VBGetPrivateProfileString("DATABASE", "AdodcPatientList.ConnectionString." & RegPatientListConnectionType, RegFileName)
'''''''    On Error GoTo 0
'''''''
    '*** Connect to PatientList DB
    AdodcPatientList.ConnectionString = RegPatientListConnectionString
    Select Case gsecSiteInformationClientShort
        Case "TTC"
            AdodcPatientList.RecordSource = "select " & strAccountNumberDbField & ", " & strAccountNameDbField & ", " & strAccountCaseIdDbField & ", " & strAccountTicketNumberDbField & ", DateStamp from " & strAccountTableName & " Where 0=1  Order by " & strAccountNumberDbField
        Case "JMH"
            AdodcPatientList.RecordSource = "select " & strAccountNumberDbField & ", " & strAccountNameDbField & ", DateStamp from " & strAccountTableName & " Where 0=1  Order by " & strAccountNumberDbField
        Case "HA"
            AdodcPatientList.RecordSource = "select " & strAccountNumberDbField & ", " & strAccountNameDbField & ", " & strAccountCaseIdDbField & ", " & strAccountCaseDbField & ", " & strAccountTicketNumberDbField & " from " & strAccountTableName & " Where 0=1  Order by " & strAccountNumberDbField
        Case "WASD"
            AdodcPatientList.RecordSource = "select " & strAccountNumberDbField & ", " & strAccountNameDbField & " from " & strAccountTableName & " Where 0=1  Order by " & strAccountNumberDbField
        Case Else
            AdodcPatientList.RecordSource = "select * from " & strAccountTableName & " Where 0=1  Order by " & strAccountNumberDbField
        
    End Select
    
    AdodcPatientList.Refresh
    
    'Set Account # from BatchName
    frmPatientList.txtPatientAccountNumber = Mid(frmIndex.txtBatchName, InStr(1, frmIndex.txtBatchName, "-") + 1, Len(frmIndex.txtBatchName) - InStr(1, frmIndex.txtBatchName, "-"))
    
    'Click to Select the Account ONLY if the BatchName includes a Dash "-" character
    If InStr(1, frmIndex.txtBatchName, "-") > 0 Then
        ' LOAD The Client
        ' This is a bit dangerous to do since it might overwrite the client
        '   on an already indexed batch.
'        grdDataGrid_Click
    End If

    

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        'The following "Cancel" would prevent the window from closing.
        ''Cancel = True
    Else
        'Save Form settings to the registry, except if minimized (i.e.- Top or Left < 0)
        If Me.Top >= 0 And Me.Left >= 0 Then
            Result = WritePrivateProfileString(RegAppname, "frmPatientList.Top", Me.Top, RegFileName)
            Result = WritePrivateProfileString(RegAppname, "frmPatientList.Left", Me.Left, RegFileName)
            Result = WritePrivateProfileString(RegAppname, "frmPatientList.Width", Me.Width, RegFileName)
            Result = WritePrivateProfileString(RegAppname, "frmPatientList.Height", Me.Height, RegFileName)
            Result = WritePrivateProfileString(RegAppname, "frmPatientList.Caption", Me.Caption, RegFileName)
        End If
    
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    
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
  'This will resize the grid when the form is resized
  grdDataGrid.Width = Me.ScaleWidth
'  grdDataGrid.Height = Me.ScaleHeight - grdDataGrid.Top - cmbArea.Height - Toolbar1.Height - 5
  grdDataGrid.Height = Me.ScaleHeight - grdDataGrid.Top - 20
  grdDataGrid.Splits(0).Columns(0).Width = 2000
  grdDataGrid.Splits(0).Columns(1).Width = 2000

End Sub

Private Sub grdDataGrid_Click()
    ' *** This procedure will send all the index info to ecIndex
    ' *** based on the field names and values entered.
    ' *** NOTE:  The field names are case sensitive!
    
    On Error GoTo ListView1_DblClick_Error
    
    txtPatientAccountName = grdDataGrid.Splits(0).Columns(1).Text
''MsgBox AdodcPatientList.Recordset.Fields.Item(0).Name & ", " & AdodcPatientList.Recordset.Fields.Item(0).Value
    ' Loop through the Fields, looking for a match
    For intIndex = 0 To frmIndex.lblFieldDescription.Count - 1
        ' Set field default value
        Select Case UCase(Trim(frmIndex.lblFieldDescription.item(intIndex).Caption))
            Case UCase(Trim(strAccountNumberEcField))
                frmIndex.mebIndexValues(intIndex).Text = Trim(Mid(grdDataGrid.Splits(0).Columns(0).Text, CLng(strAccountNumberEcFieldBegin), CLng(strAccountNumberEcFieldLength)))
            Case UCase(Trim(strAccountNameEcField))
                frmIndex.mebIndexValues(intIndex).Text = Left(Trim(grdDataGrid.Splits(0).Columns(1).Text), frmIndex.txtFieldSize(intIndex))
            Case UCase(Trim(strAccountCaseEcField))
                frmIndex.mebIndexValues(intIndex).Text = Trim(grdDataGrid.Splits(0).Columns(2).Text)
            Case UCase(Trim(strAccountCaseIdEcField))
                frmIndex.mebIndexValues(intIndex).Text = Trim(grdDataGrid.Splits(0).Columns(3).Text)
            Case UCase(Trim(strAccountTicketNumberEcField))
                frmIndex.mebIndexValues(intIndex).Text = Trim(grdDataGrid.Splits(0).Columns(4).Text)
        End Select
    Next
    
    
    ' ***** Find the Field Specified in the INI file as
    ' *****   FIELDAFTERCLICK and set the focus to it.
    ' ***** NOTE:  The reason for this second loop pass is to make sure we get
    '              ALL the field values populated first.
    For intIndex = 0 To frmIndex.lblFieldDescription.Count - 1
        Select Case UCase(Trim(frmIndex.lblFieldDescription.item(intIndex).Caption))
            Case UCase(Trim(strFIELDAFTERCLICK))
                frmIndex.SetFocus
                frmIndex.mebIndexValues(intIndex).SetFocus
        End Select
    Next
    
    
    Exit Sub
    
ListView1_DblClick_Error:
    ' Err 91 = List is empty
    If Err.Number = 91 Then
        Resume Next
    End If
    MsgBox "Client grdDataGrid_Click - Error: " & Err.Number & " - " & Err.Description & " [" & Err.Source & "]", vbExclamation
    Exit Sub
    
    ' We must trap for error in case the user enters an invalid field name
    
    
    
End Sub


Private Sub grdDataGrid_KeyPress(KeyAscii As Integer)
    'Catch Enter key
    If KeyAscii = 13 Then
        grdDataGrid_Click
    End If
End Sub
Private Sub txtPatientAccountNumber_Change()
    If Trim(txtPatientAccountNumber) = "" Then
        ' Get Empty Result Set
        Select Case gsecSiteInformationClientShort
            Case "TTC"
                AdodcPatientList.RecordSource = "select " & strAccountNumberDbField & ", " & strAccountNameDbField & ", " & strAccountCaseDbField & ", " & strAccountCaseIdDbField & ", " & strAccountTicketNumberDbField & ", DateStamp from " & strAccountTableName & " Where 0=1  Order by " & strAccountNumberDbField
            Case "JMH"
                AdodcPatientList.RecordSource = "select " & strAccountNumberDbField & ", " & strAccountNameDbField & ", DateStamp from " & strAccountTableName & " Where 0=1  Order by " & strAccountNumberDbField
            Case "HA"
                AdodcPatientList.RecordSource = "select " & strAccountNumberDbField & ", " & strAccountNameDbField & ", " & strAccountCaseDbField & ", " & strAccountCaseIdDbField & ", " & strAccountTicketNumberDbField & " from " & strAccountTableName & " Where 0=1  Order by " & strAccountNumberDbField
            Case "WASD"
                AdodcPatientList.RecordSource = "select " & strAccountNumberDbField & ", " & strAccountNameDbField & " FROM " & strAccountTableName & " WHERE 0=1  ORDER BY " & strAccountNumberDbField
            Case Else
                AdodcPatientList.RecordSource = "select * FROM " & strAccountTableName & " WHERE 0=1  ORDER BY " & strAccountNumberDbField
        End Select
        AdodcPatientList.Refresh
    ElseIf Len(Trim(txtPatientAccountNumber)) >= strAccountNumberSearchChars Then
        ' Find Patient Records
        Select Case gsecSiteInformationClientShort
            Case "TTC"
                AdodcPatientList.RecordSource = "select " & strAccountNumberDbField & ", " & strAccountNameDbField & ", " & strAccountCaseDbField & ", " & strAccountCaseIdDbField & ", " & strAccountTicketNumberDbField & ", DateStamp from " & strAccountTableName & " Where " & strAccountNumberDbField & " LIKE '" & Trim(txtPatientAccountNumber) & "%'" & " OR " & strAccountNameDbField & " LIKE '" & Trim(txtPatientAccountNumber) & "%'" & " OR " & strAccountCaseDbField & " LIKE '" & Trim(txtPatientAccountNumber) & "%'" & " OR " & strAccountTicketNumberDbField & " LIKE '" & Trim(txtPatientAccountNumber) & "%'" & " ORDER BY " & strAccountNumberDbField
            Case "JMH"
                AdodcPatientList.RecordSource = "select " & strAccountNumberDbField & ", " & strAccountNameDbField & ", DateStamp from " & strAccountTableName & " Where " & strAccountNumberDbField & " LIKE '" & Trim(txtPatientAccountNumber) & "%'" & " OR " & strAccountNameDbField & " LIKE '" & Trim(txtPatientAccountNumber) & "%'" & " ORDER BY " & strAccountNumberDbField
            Case "HA"
                AdodcPatientList.RecordSource = "select " & strAccountNumberDbField & ", " & strAccountNameDbField & ", " & strAccountCaseDbField & ", " & strAccountCaseIdDbField & ", " & strAccountTicketNumberDbField & " from " & strAccountTableName & " Where " & strAccountNumberDbField & " LIKE '" & Trim(txtPatientAccountNumber) & "%'" & " OR " & strAccountNameDbField & " LIKE '" & Trim(txtPatientAccountNumber) & "%'" & " OR " & strAccountCaseDbField & " LIKE ' " & Trim(txtPatientAccountNumber) & "%' OR " & strAccountCaseIdDbField & " LIKE '" & Trim(txtPatientAccountNumber) & "%' OR " & strAccountTicketNumberDbField & " LIKE '" & Trim(txtPatientAccountNumber) & "%'" & " ORDER BY " & strAccountNumberDbField
            Case "WASD"
                AdodcPatientList.RecordSource = "select " & strAccountNumberDbField & ", " & strAccountNameDbField & " FROM " & strAccountTableName & " WHERE " & strAccountNumberDbField & " LIKE '" & Trim(txtPatientAccountNumber) & "%'" & " ORDER BY " & strAccountNumberDbField
            Case Else
                AdodcPatientList.RecordSource = "select * FROM " & strAccountTableName & " Where " & strAccountNumberDbField & " LIKE '" & Trim(txtPatientAccountNumber) & "%'" & " OR " & strAccountNameDbField & " LIKE '" & Trim(txtPatientAccountNumber) & "%'" & " ORDER BY " & strAccountNumberDbField
        End Select
        AdodcPatientList.Refresh
    End If
'    txtPatientAccountNumber.SetFocus
    
End Sub

Private Sub cmdClearFields_Click()
    txtPatientAccountNumber = ""
    txtPatientAccountName = ""
    txtPatientAccountNumber.SetFocus
End Sub
