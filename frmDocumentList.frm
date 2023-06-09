VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmDocumentList 
   Caption         =   "Document List"
   ClientHeight    =   7020
   ClientLeft      =   6465
   ClientTop       =   255
   ClientWidth     =   3195
   Icon            =   "frmDocumentList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   3195
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView ListView1 
      Height          =   5895
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   10398
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      Default         =   -1  'True
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox txtSearchString 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc AdodcAreaList 
      Height          =   330
      Left            =   120
      Top             =   7320
      Visible         =   0   'False
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
      Caption         =   "AdodcAreaList"
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
   Begin MSDataListLib.DataCombo DataComboAreaList 
      Bindings        =   "frmDocumentList.frx":0442
      DataSource      =   "AdodcAreaList"
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   ""
      Text            =   "*Select Area"
   End
   Begin MSAdodcLib.Adodc AdodcDocList 
      Height          =   330
      Left            =   120
      Top             =   6960
      Visible         =   0   'False
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
      Caption         =   "AdodcDocList"
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
Attribute VB_Name = "frmDocumentList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strDOCGROUP As String
Dim strDOCTYPE As String
Dim RegConnectionWildcard As String

Private Const HWND_TOPMOST = -1
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Sub cmdFind_Click()
        AdodcDocList.RecordSource = "select FORMDESC, DOCGROUP, DOCTYPE, AREA, APPLICATION  from DOCTYPES"
        If Trim(DataComboAreaList.Text) <> "*ALL*" Then
            AdodcDocList.RecordSource = AdodcDocList.RecordSource & " WHERE APPLICATION = '" & frmIndex.txtApplicationName & "' AND AREA = '" & DataComboAreaList.Text & "' AND FORMDESC Like '" & RegConnectionWildcard & txtSearchString & RegConnectionWildcard & "'"
        End If
        AdodcDocList.RecordSource = AdodcDocList.RecordSource & " order by FORMDESC"
        AdodcDocList.Refresh

        subListViewPopulate
End Sub

Private Sub Form_Activate()
''    MsgBox "Form Activate"
End Sub

Private Sub Form_GotFocus()
    txtCurrentModule = "frmDocumentList"
    txtSearchString.SetFocus

End Sub

Private Sub Form_Initialize()
  
    Dim RegConnectString As String
    Dim RegConnectionType As String
    
'''''''    ' Get Database Connections settings from the registry
'''''''    On Error Resume Next
'''''''    RegConnectionType = VBGetPrivateProfileString("DATABASE", "AdodcDocList.ConnectionType", RegFileName)
'''''''    RegConnectionString = VBGetPrivateProfileString("DATABASE", "AdodcDocList.ConnectionString." & RegConnectionType, RegFileName)
'''''''    On Error GoTo 0
    
    '*** Set SQL wildcard string
''    If RegConnectionType = "Access" Then
''        RegConnectionWildcard = "*"
''    Else
        RegConnectionWildcard = "%"
''    End If
        
    
    '*** Connect to DocList FIRST because when AreaList Combo populates
    '    it automatically triggers a "DataComboAreaList_Change()" event!
    
    AdodcDocList.ConnectionString = RegDocTypeListConnectionString
    AdodcDocList.RecordSource = ""
    
    AdodcAreaList.ConnectionString = RegDocTypeListConnectionString
    AdodcAreaList.RecordSource = "Select distinct [area] from DOCTYPES"
    DataComboAreaList.ListField = "Area"
    AdodcAreaList.Refresh

    ' Get Position settings from the registry
'    On Error Resume Next
    Me.Top = VBGetPrivateProfileString(RegAppname, "frmDocumentList.Top", RegFileName)
    Me.Left = VBGetPrivateProfileString(RegAppname, "frmDocumentList.Left", RegFileName)
    Me.Width = VBGetPrivateProfileString(RegAppname, "frmDocumentList.Width", RegFileName)
    Me.Height = VBGetPrivateProfileString(RegAppname, "frmDocumentList.Height", RegFileName)
    Me.Caption = VBGetPrivateProfileString(RegAppname, "frmDocumentList.Caption", RegFileName)
    If Me.Caption = "" Then Me.Caption = "Document List"
    
    strDOCGROUP = VBGetPrivateProfileString(RegAppname, "frmDocumentList.DOCGROUP", RegFileName)
    strDOCTYPE = VBGetPrivateProfileString(RegAppname, "frmDocumentList.DOCTYPE", RegFileName)
    strFIELDAFTERCLICK = VBGetPrivateProfileString(RegAppname, "frmDocumentList.FIELDAFTERCLICK", RegFileName)

    ' Populate the Area List Combo with an Asterisk (*) to select all records.
    DataComboAreaList.Text = "*"
    
    On Error GoTo 0



End Sub

Private Sub Form_Load()
''    MsgBox "Form Load"
    'Select Records
''    cmdFind_Click
    
    
End Sub

Private Sub Form_LostFocus()
    Debug.Print "Lost Focus"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Debug.Print "MouseDown"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Debug.Print "MouseMove"
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Debug.Print "MouseUp"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    Else
        'Save Form settings to the registry, except if minimized (i.e.- Top or Left < 0)
        If Me.Top > 0 And Me.Left > 0 Then
            Result = WritePrivateProfileString(RegAppname, "frmDocumentList.Top", Me.Top, RegFileName)
            Result = WritePrivateProfileString(RegAppname, "frmDocumentList.Left", Me.Left, RegFileName)
            Result = WritePrivateProfileString(RegAppname, "frmDocumentList.Width", Me.Width, RegFileName)
            Result = WritePrivateProfileString(RegAppname, "frmDocumentList.Height", Me.Height, RegFileName)
            Result = WritePrivateProfileString(RegAppname, "frmDocumentList.Caption", Me.Caption, RegFileName)
        End If
    
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    txtCurrentModule = ""

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
  
  DataComboAreaList.Width = Me.ScaleWidth

  ListView1.Width = Me.ScaleWidth
  ListView1.Height = Me.ScaleHeight - ListView1.Top - DataComboAreaList.Height - 5

End Sub

Private Sub subListViewPopulate()
    '*** Setup Up ListView properties - BEGIN
    ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Clear
    ListView1.View = lvwReport
    '    Column widths are in PIXELS!
    ListView1.MultiSelect = True
    ListView1.FullRowSelect = True
    ListView1.LabelEdit = lvwManual
    ListView1.AllowColumnReorder = True
    
        '***  SET COLUMN HEADINGS
        For intListIndex = 0 To AdodcDocList.Recordset.Fields.Count - 1
            ListView1.ColumnHeaders.Add , , AdodcDocList.Recordset.Fields.item(intListIndex).name, Len(AdodcDocList.Recordset.Fields.item(intListIndex).name) * 500, lvwColumnLeft
        Next
                
    On Error Resume Next
    AdodcDocList.Recordset.MoveFirst
    While Not AdodcDocList.Recordset.EOF
            For intListIndex = 0 To AdodcDocList.Recordset.Fields.Count - 1
                If intListIndex = 0 Then
                    If Not IsNull(AdodcDocList.Recordset.Fields.item(intListIndex).Value) Then
                        Set lstItem = ListView1.ListItems.Add(, , AdodcDocList.Recordset.Fields.item(intListIndex).Value)
''                        ' Set Width of Column to size of text
''                        funcListView_SetColumnWidth frmDocumentList.ListView1, 1, AdodcDocList.Recordset.Fields.Item(intListIndex).Value
                    End If
                Else
                    '* This null check is to make sure we don't Skip fields caused by an error.
                    If Not IsNull(AdodcDocList.Recordset.Fields.item(intListIndex).Value) Then
                        ' Not null... show value
                        Set lstSubItem = lstItem.ListSubItems.Add(, , AdodcDocList.Recordset.Fields.item(intListIndex).Value)
                    Else
                        ' Null... show empty string
                        Set lstSubItem = lstItem.ListSubItems.Add(, , "")
                    End If
                End If
            Next
            AdodcDocList.Recordset.MoveNext
    Wend
    On Error GoTo 0
    '*** Setup Up ListView properties - END

End Sub

Private Sub DataComboAreaList_Change()
        AdodcDocList.RecordSource = "select FORMDESC, DOCGROUP, DOCTYPE, AREA, APPLICATION, PAGES from DOCTYPES WHERE APPLICATION = '" & frmIndex.txtApplicationName & "'"
        If Trim(DataComboAreaList.Text) <> "*ALL*" Then
            ' Display Questionable, Do Not File and Separators Sheets always because AREA is set to NULL
            AdodcDocList.RecordSource = AdodcDocList.RecordSource & " AND AREA = '" & DataComboAreaList.Text & "'" & " OR AREA is null"
        End If
        AdodcDocList.RecordSource = AdodcDocList.RecordSource & " order by FORMDESC"
        AdodcDocList.Refresh
   
        subListViewPopulate
End Sub


Private Sub ListView1_Click()
    ' Allow Single-Click to select Items
    ListView1_DblClick
End Sub

Private Sub ListView1_DblClick()

    
    On Error GoTo ListView1_DblClick_Error
    
    ' *** This procedure will send the Index Values to the Index Fields
    ' *** based on the Field Description defined in the INI file.
    ' *** NOTE:  The Field Descriptions are case sensitive!
    For intIndex = 0 To frmIndex.lblFieldDescription.Count - 1
        Select Case Trim(frmIndex.lblFieldDescription.item(intIndex).Caption)
            Case Trim(strDOCGROUP)
                frmIndex.mebIndexValues(intIndex).Text = Trim(Me.ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(1).Text)
            Case Trim(strDOCTYPE)
                frmIndex.mebIndexValues(intIndex).Text = Trim(Me.ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(2).Text)
        End Select
    Next
    
    ' ***** Find the Field Specified in the INI file as
    ' *****   FIELDAFTERCLICK and set the focus to it.
    ' ***** NOTE:  The reason for this second loop pass is to make sure we get
    '              ALL the field values populated first.
    For intIndex = 0 To frmIndex.lblFieldDescription.Count - 1
        Select Case Trim(frmIndex.lblFieldDescription.item(intIndex).Caption)
            Case Trim(strFIELDAFTERCLICK)
                frmIndex.SetFocus
                frmIndex.mebIndexValues(intIndex).SetFocus
        End Select
    Next
    
    frmIndex.txtBatchDocDesc = Trim(Me.ListView1.ListItems(ListView1.SelectedItem.Index).Text)
    frmIndex.txtExpectedPages = Trim(Me.ListView1.ListItems.item(ListView1.SelectedItem.Index).ListSubItems.item(5).Text)
    
    ' Move to the next Image if AutoAdvance Global variable is checked,
    '  and the item clicked was a Separator or Questionable
    If gAutoAdvanceOnSeparator = "" Then gAutoAdvanceOnSeparator = vbUnchecked
    If gAutoAdvanceOnSeparator = vbChecked Then
        Select Case Trim(Me.ListView1.ListItems(ListView1.SelectedItem.Index).Text)
        Case txtSeparator
            frmIndex.cmdNextImage_Click
        Case txtQuestionable
            frmIndex.cmdNextImage_Click
        End Select
    End If
    
Exit Sub

ListView1_DblClick_Error:
    ' Err 91 = List is empty
    If Err.Number = 91 Then
        Resume Next
    End If
    ' Err 35600 = Index out of bounds
    If Err.Number = 35600 Then
        Resume Next
    End If
    MsgBox "ListView1_DblClick - Error: " & Err.Number & " - " & Err.Description & " [" & Err.Source & "]", vbExclamation
    Exit Sub
    
End Sub

Private Sub txtSearchString_Change()
        
    Select Case Len(txtSearchString)
        Case Is < 1
            DataComboAreaList_Change
        Case Is >= 3
            cmdFind_Click
    End Select
    
End Sub
