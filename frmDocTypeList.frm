VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDocTypeList 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Document List"
   ClientHeight    =   6990
   ClientLeft      =   6465
   ClientTop       =   255
   ClientWidth     =   3195
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDocTypeList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   3195
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      MaskColor       =   &H00C0C0FF&
      Picture         =   "frmDocTypeList.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton cmdApplicationEditDoctypes 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      Picture         =   "frmDocTypeList.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   600
      Width           =   480
   End
   Begin VB.CheckBox chkMoveToNextPageOnSingleClick 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Next Page on Single-Click"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   2295
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5895
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   10398
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.TextBox txtSearchString 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   1095
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
      Bindings        =   "frmDocTypeList.frx":0D56
      DataSource      =   "AdodcAreaList"
      Height          =   288
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3132
      _ExtentX        =   5530
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   ""
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdodcDocTypeList 
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
Attribute VB_Name = "frmDocTypeList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strDOCGROUP As String
Dim strDOCTYPE As String
Dim strDOCSUBTYPE As String
Dim strPAGES As String
Dim intPagesLoop As Integer
Dim bolProcessingDocTypeSelection As Boolean

Dim RegConnectionWildcard As String

Private Const HWND_TOPMOST = -1
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long



Private Sub cmdApplicationEditDoctypes_Click()

    frmDOCTYPES2.Show

End Sub

Private Sub cmdFind_Click()

        'JR 12/20/2004 - Corrected problem with DocTypeList where doctypes flagged as "*ALL*" would
        '                appear for ALL Applications and the FIND Criteria was not working.
        'JR 12-6-2013 - DocType Search during AIM AddFile would cause ERROR
        
        Dim strApplicationName As String
        
        If bolAIM_Command_AddFile = True Then
            strApplicationName = frmImaging101Search.txtApplicationName
        Else
            strApplicationName = frmIndex.txtApplicationName
        End If
        
        AdodcDocTypeList.RecordSource = "select FORMDESC, DOCGROUP, DOCTYPE, DOCSUBTYPE, AREA, APPLICATION, PAGES, RouteToQueue   " & _
                                        " FROM DOCTYPES " & _
                                        " WHERE APPLICATION = '" & strApplicationName & "' " & _
                                        " AND FORMDESC Like " & _
                                        " '" & RegConnectionWildcard & txtSearchString & RegConnectionWildcard & "'"
        
        If Trim(DataComboAreaList.Text) <> "*ALL*" Then
            AdodcDocTypeList.RecordSource = AdodcDocTypeList.RecordSource & _
                                            " AND AREA = '" & DataComboAreaList.Text & "' "

        End If
        
        AdodcDocTypeList.RecordSource = AdodcDocTypeList.RecordSource & " ORDER by FORMDESC"
        AdodcDocTypeList.Refresh

        subListViewPopulate
End Sub

Private Sub DataComboAreaList_GotFocus()
' Gotfocus

End Sub

Private Sub Initialize_DocTypeList()

   Dim RegConnectString As String
    Dim RegConnectionType As String
    

    '**********************************************************
    '*** See if we should display the "Edit DocTypes" Button
    If gsecRightsBatchAllowDocTypeEdit = vbChecked Then
        cmdApplicationEditDoctypes.Visible = True
    Else
        cmdApplicationEditDoctypes.Visible = False
    End If
     
    
'''''''    ' Get Database Connections settings from the registry
'''''''    On Error Resume Next
'''''''    RegConnectionType = VBGetPrivateProfileString("DATABASE", "AdodcDocTypeList.ConnectionType", RegFileName)
'''''''    RegConnectionString = VBGetPrivateProfileString("DATABASE", "AdodcDocTypeList.ConnectionString." & RegConnectionType, RegFileName)
'''''''    On Error GoTo 0
    
    '*** Set SQL wildcard string
''    If RegConnectionType = "Access" Then
''        RegConnectionWildcard = "*"
''    Else
        RegConnectionWildcard = "%"
''    End If
        
    
    '*** Connect to DocList FIRST because when AreaList Combo populates
    '    it automatically triggers a "DataComboAreaList_Change()" event!
    
    AdodcDocTypeList.ConnectionString = RegDocTypeListConnectionString
    AdodcDocTypeList.RecordSource = ""
    
    AdodcAreaList.ConnectionString = RegDocTypeListConnectionString
    AdodcAreaList.RecordSource = "Select distinct [area] from DOCTYPES"
    DataComboAreaList.ListField = "Area"
    AdodcAreaList.Refresh

    ' Get Position settings from the registry
'    On Error Resume Next
    Me.Top = VBGetPrivateProfileString(RegAppname, "frmDocTypeList.Top", RegFileName)
    Me.Left = VBGetPrivateProfileString(RegAppname, "frmDocTypeList.Left", RegFileName)
    Me.width = VBGetPrivateProfileString(RegAppname, "frmDocTypeList.Width", RegFileName)
    Me.Height = VBGetPrivateProfileString(RegAppname, "frmDocTypeList.Height", RegFileName)
'    Me.Caption = VBGetPrivateProfileString(RegAppname, "frmDocTypeList.Caption", RegFileName)
'    If Me.Caption = "" Then Me.Caption = "Document List"
    
    'Determines if we should move to the next page on a single click
    Dim chkValue As String
    
    If bolAIM_Command_AddFile = True Then
        strDOCGROUP = funcGetFieldFromDB(RegImaging101ConnectionString, "I101Applications", "ApplicationRECID = " & frmImaging101Search.txtApplicationRECID, "FieldToAssignDocumentGroup") & ""
        strDOCTYPE = funcGetFieldFromDB(RegImaging101ConnectionString, "I101Applications", "ApplicationRECID = " & frmImaging101Search.txtApplicationRECID, "FieldToAssignDocumentType") & ""
        strDOCSUBTYPE = funcGetFieldFromDB(RegImaging101ConnectionString, "I101Applications", "ApplicationRECID = " & frmImaging101Search.txtApplicationRECID, "FieldToAssignDocumentSubType") & ""
        strFIELDAFTERCLICK = funcGetFieldFromDB(RegImaging101ConnectionString, "I101Applications", "ApplicationRECID = " & frmImaging101Search.txtApplicationRECID, "FieldToSelectAfterDocListClick") & ""
        chkValue = ""
    Else
        strDOCGROUP = funcGetFieldFromDB(RegImaging101ConnectionString, "I101Applications", "ApplicationRECID = " & frmIndex.txtApplicationRECID, "FieldToAssignDocumentGroup") & ""
        strDOCTYPE = funcGetFieldFromDB(RegImaging101ConnectionString, "I101Applications", "ApplicationRECID = " & frmIndex.txtApplicationRECID, "FieldToAssignDocumentType") & ""
        strDOCSUBTYPE = funcGetFieldFromDB(RegImaging101ConnectionString, "I101Applications", "ApplicationRECID = " & frmIndex.txtApplicationRECID, "FieldToAssignDocumentSubType") & ""
        strFIELDAFTERCLICK = funcGetFieldFromDB(RegImaging101ConnectionString, "I101Applications", "ApplicationRECID = " & frmIndex.txtApplicationRECID, "FieldToSelectAfterDocListClick") & ""
        chkValue = funcGetSetUserSettings("GET", "MoveToNextPageOnSingleClick", chkMoveToNextPageOnSingleClick.Value) & ""
    End If
    
    If chkValue = "" Then
        chkMoveToNextPageOnSingleClick.Value = 0
    Else
        chkMoveToNextPageOnSingleClick.Value = chkValue
    End If
    
    ' Populate the Area List Combo with an Asterisk (*) to select all records.
'    DataComboAreaList.Text = "*"
    
    'GET the LAST selected Area for This user
    DataComboAreaList.Text = funcGetSetUserSettings("GET", "DocTypeListLastSelection", DataComboAreaList.Text)
    
    If DataComboAreaList.Text = "" Then
        'If NO Area set, set focus to allow selection
        DataComboAreaList.Text = "*Select Area"
    End If
    On Error GoTo 0


End Sub

Private Sub Form_Activate()
''    MsgBox "Form Activate"
End Sub

Private Sub Form_GotFocus()
    txtCurrentModule = "frmDocTypeList"
    txtSearchString.SetFocus

End Sub

Private Sub Form_Load()
''    MsgBox "Form Load"
    'Select Records
''    cmdFind_Click
    
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
    
    Call Initialize_DocTypeList
    
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


        If UnloadMode = vbUser Then
            Cancel = True
            Exit Sub
        End If
        
        'Save Form settings to the registry, except if minimized (i.e.- Top or Left < 0)
        If Me.Top > 0 And Me.Left > 0 Then
            result = WritePrivateProfileString(RegAppname, "frmDocTypeList.Top", Me.Top, RegFileName)
            result = WritePrivateProfileString(RegAppname, "frmDocTypeList.Left", Me.Left, RegFileName)
            result = WritePrivateProfileString(RegAppname, "frmDocTypeList.Width", Me.width, RegFileName)
            result = WritePrivateProfileString(RegAppname, "frmDocTypeList.Height", Me.Height, RegFileName)
'            result = WritePrivateProfileString(RegAppname, "frmDocTypeList.Caption", Me.Caption, RegFileName)
    
            funcGetSetUserSettings "SET", "MoveToNextPageOnSingleClick", chkMoveToNextPageOnSingleClick
        End If
    
'    End If
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
  
  DataComboAreaList.width = Me.ScaleWidth

    If frmDocTypeList.WindowState = vbNormal Then
        ListView1.width = Me.ScaleWidth
        ListView1.Height = Me.ScaleHeight - ListView1.Top - DataComboAreaList.Height - 5
    End If
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
        For intListIndex = 0 To AdodcDocTypeList.Recordset.Fields.count - 1
            ListView1.ColumnHeaders.Add , , AdodcDocTypeList.Recordset.Fields.item(intListIndex).name, Len(AdodcDocTypeList.Recordset.Fields.item(intListIndex).name) * 500, lvwColumnLeft
        Next
                
    On Error Resume Next
    AdodcDocTypeList.Recordset.MoveFirst
    While Not AdodcDocTypeList.Recordset.EOF
            For intListIndex = 0 To AdodcDocTypeList.Recordset.Fields.count - 1
                If intListIndex = 0 Then
                    If Not IsNull(AdodcDocTypeList.Recordset.Fields.item(intListIndex).Value) Then
                        Set lstItem = ListView1.ListItems.Add(, , AdodcDocTypeList.Recordset.Fields.item(intListIndex).Value)
''                        ' Set Width of Column to size of text
''                        funcListView_SetColumnWidth frmDocTypeList.ListView1, 1, AdodcDocTypeList.Recordset.Fields.Item(intListIndex).Value
                    End If
                Else
                    '* This null check is to make sure we don't Skip fields caused by an error.
                    If Not IsNull(AdodcDocTypeList.Recordset.Fields.item(intListIndex).Value) Then
                        ' Not null... show value
                        Set lstSubItem = lstItem.ListSubItems.Add(, , AdodcDocTypeList.Recordset.Fields.item(intListIndex).Value)
                    Else
                        ' Null... show empty string
                        Set lstSubItem = lstItem.ListSubItems.Add(, , "")
                    End If
                End If
            Next
            AdodcDocTypeList.Recordset.MoveNext
    Wend
    On Error GoTo 0
    '*** Setup Up ListView properties - END

End Sub

Private Sub DataComboAreaList_Change()

        If bolAIM_Command_AddFile = True Then
            AdodcDocTypeList.RecordSource = "select FORMDESC, DOCGROUP, DOCTYPE, DOCSUBTYPE, AREA, APPLICATION, PAGES, RouteToQueue,CommitViaFTP from DOCTYPES WHERE APPLICATION = '" & frmImaging101Search.txtApplicationName & "'"
        Else
            AdodcDocTypeList.RecordSource = "select FORMDESC, DOCGROUP, DOCTYPE, DOCSUBTYPE, AREA, APPLICATION, PAGES, RouteToQueue, CommitViaFTP from DOCTYPES WHERE APPLICATION = '" & frmIndex.txtApplicationName & "'"
        End If
        
        If Trim(DataComboAreaList.Text) <> "*ALL*" Then
            ' Display Questionable, Do Not File and Separators Sheets always because AREA is set to NULL
            AdodcDocTypeList.RecordSource = AdodcDocTypeList.RecordSource & " AND AREA = '" & DataComboAreaList.Text & "'"
        End If
        AdodcDocTypeList.RecordSource = AdodcDocTypeList.RecordSource & " OR AREA is null"
        AdodcDocTypeList.RecordSource = AdodcDocTypeList.RecordSource & " order by FORMDESC"
        AdodcDocTypeList.Refresh
   
        subListViewPopulate
        
        'Save the selected Area
        funcGetSetUserSettings "SET", "DocTypeListLastSelection", DataComboAreaList.Text
End Sub


Private Sub ListView1_Click()
    ' Allow Single-Click to select Items
    
    'Make sure user can't double click
    If bolProcessingDocTypeSelection Then
        Exit Sub
    Else
        bolProcessingDocTypeSelection = True
    End If
    
    
    If bolAIM_Command_AddFile = True Then
        ListView1_Click_AssignFieldValues frmImaging101Search
        
    Else
        ListView1_Click_AssignFieldValues frmIndex
    End If
    
    
    
End Sub

    
    
Private Sub ListView1_Click_AssignFieldValues(WorkingForm As Form)

    Dim intPageIndex As Integer
    Dim strRouteToQueue As String
    Dim strCommitViaFTP As String
    
    On Error GoTo ListView1_DblClick_Error
    
    ' *** This procedure will send the Index Values to the Index Fields
    ' *** based on the Field Description defined in the INI file.
    ' *** NOTE:  The Field Descriptions are case sensitive!
    For intIndex = 0 To WorkingForm.lblFieldDescription.count - 1
        Select Case Trim(WorkingForm.lblFieldDescription.item(intIndex).Caption)
            Case Trim(strDOCGROUP)
                '*** Check if the txtIndexValues TextBox control is VISIBLE...
                '    this will handle saving the value of the the TextBox control
                '    instead of the mebIndexValues Masked Edit control as needed.
                '    This is because the Masked Edit control has a MAX size of 64 Char.
                If WorkingForm.txtIndexValues(intIndex).Visible = True Then
                    WorkingForm.txtIndexValues(intIndex).Text = Trim(Me.ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(1).Text)
                Else
                    WorkingForm.mebIndexValues(intIndex).Text = Trim(Me.ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(1).Text)
                End If
                
            Case Trim(strDOCTYPE)
                '*** Check if the txtIndexValues TextBox control is VISIBLE...
                '    this will handle saving the value of the the TextBox control
                '    instead of the mebIndexValues Masked Edit control as needed.
                '    This is because the Masked Edit control has a MAX size of 64 Char.
                If WorkingForm.txtIndexValues(intIndex).Visible = True Then
                    WorkingForm.txtIndexValues(intIndex).Text = Trim(Me.ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(2).Text)
                Else
                    WorkingForm.mebIndexValues(intIndex).Text = Trim(Me.ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(2).Text)
                End If
                
            Case Trim(strDOCSUBTYPE)
                '*** Check if the txtIndexValues TextBox control is VISIBLE...
                '    this will handle saving the value of the the TextBox control
                '    instead of the mebIndexValues Masked Edit control as needed.
                '    This is because the Masked Edit control has a MAX size of 64 Char.
                If WorkingForm.txtIndexValues(intIndex).Visible = True Then
                    WorkingForm.txtIndexValues(intIndex).Text = Trim(Me.ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(3).Text)
                Else
                    WorkingForm.mebIndexValues(intIndex).Text = Trim(Me.ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(3).Text)
                End If
        End Select
        
        'If NOT Adding a File via Office Integration Module
        If bolAIM_Command_AddFile = False Then
            '*** Check if the RouteToQueue value is set for the selected DOCTYPE
            '***   and set it ONLY if the field is BLANK
            strRouteToQueue = Trim(Me.ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(7).Text)
            If WorkingForm.txtFieldRouteToBatchQueue(intIndex).Text = "1" _
            And strRouteToQueue <> "" Then
                '*** Check if the txtIndexValues TextBox control is VISIBLE...
                '    this will handle saving the value of the the TextBox control
                '    instead of the mebIndexValues Masked Edit control as needed.
                '    This is because the Masked Edit control has a MAX size of 64 Char.
                If WorkingForm.txtIndexValues(intIndex).Visible = True Then
                    WorkingForm.txtIndexValues(intIndex).Text = strRouteToQueue
                Else
                    WorkingForm.mebIndexValues(intIndex).Text = strRouteToQueue
                End If
            
                
            End If
            
             'Send the CommitViaFTP flag
             WorkingForm.txtCommitViaFTP.Text = Trim(Me.ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(8).Text)
           
        End If  'bolAIM_Command_AddFile = False
        
    Next
    
    If bolAIM_Command_AddFile = False Then
    
        '*** Get the number of Pages this Document Type is expected to have
        '     we will process this in the ListView1_DblClick() sub
        strPAGES = Trim(Me.ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(6).Text)
        'Make sure there is at least 1 Page set for each document
        If Trim(strPAGES) = "" Then
            strPAGES = "1"
        End If
        
        ' ***** Find the Field Specified in the INI file as
        ' *****   FIELDAFTERCLICK and set the focus to it.
        ' ***** NOTE:  The reason for this second loop pass is to make sure we get
        '              ALL the field values populated first.
        For intIndex = 0 To WorkingForm.lblFieldDescription.count - 1
            Select Case Trim(WorkingForm.lblFieldDescription.item(intIndex).Caption)
                Case Trim(strFIELDAFTERCLICK)
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
                    
            End Select
        Next
        
        WorkingForm.txtBatchDocDesc = Trim(Me.ListView1.ListItems(ListView1.SelectedItem.Index).Text)
'        WorkingForm.txtExpectedPages = Trim(Me.ListView1.ListItems.item(ListView1.SelectedItem.Index).ListSubItems.item(6).Text)
        
        ' Move to the next Image if AutoAdvance Global variable is checked,
        '  and the item clicked was a Separator or Questionable
        If gAutoAdvanceOnSeparator = "" Then
            gAutoAdvanceOnSeparator = vbUnchecked
        End If
        
        'Clear the DocType Fields if AutoAdvance selected UNLESS it is the LAST Image!
        If gAutoAdvanceOnSeparator = vbChecked _
        And WorkingForm.ListView1.SelectedItem.Index < (WorkingForm.ListView1.ListItems.count - 1) _
        Then
            Select Case Trim(Me.ListView1.ListItems(ListView1.SelectedItem.Index).Text)
            Case txtSeparator
                WorkingForm.cmdNextImage_Click
                subClearDocTypeFields
            Case txtQuestionable
                WorkingForm.cmdNextImage_Click
                subClearDocTypeFields
            Case txtDoNotFile
                WorkingForm.cmdNextImage_Click
                subClearDocTypeFields
            Case Else
            
                ' Check if we should Auto-Advance
                subCheckForAutoAdvance
                   
            End Select
        
        Else
        
                ' Check if we should Auto-Advance
                subCheckForAutoAdvance
    
        End If
    
    End If  'bolAIM_Command_AddFile = False
    
    'Done processing the DocType
    bolProcessingDocTypeSelection = False
    
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
    MsgBox "frmDocTypeList ListView1_DblClick - Error: " & Err.Number & " - " & Err.Description & " [" & Err.Source & "]", vbExclamation
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


Sub subClearDocTypeFields()

    ' *** This procedure will send the Index Values to the Index Fields
    ' *** based on the Field Description defined in the INI file.
    ' *** NOTE:  The Field Descriptions are case sensitive!
    For intIndex = 0 To frmIndex.lblFieldDescription.count - 1
        Select Case Trim(frmIndex.lblFieldDescription.item(intIndex).Caption)
            Case Trim(strDOCGROUP)
                '*** Check if the txtIndexValues TextBox control is VISIBLE...
                '    this will handle saving the value of the the TextBox control
                '    instead of the mebIndexValues Masked Edit control as needed.
                '    This is because the Masked Edit control has a MAX size of 64 Char.
                If frmIndex.txtIndexValues(intIndex).Visible = True Then
                    frmIndex.txtIndexValues(intIndex).Text = ""
                Else
                    frmIndex.mebIndexValues(intIndex).Text = ""
                End If
                
            Case Trim(strDOCTYPE)
                '*** Check if the txtIndexValues TextBox control is VISIBLE...
                '    this will handle saving the value of the the TextBox control
                '    instead of the mebIndexValues Masked Edit control as needed.
                '    This is because the Masked Edit control has a MAX size of 64 Char.
                If frmIndex.txtIndexValues(intIndex).Visible = True Then
                    frmIndex.txtIndexValues(intIndex).Text = ""
                Else
                    frmIndex.mebIndexValues(intIndex).Text = ""
                End If
                    
            Case Trim(strDOCSUBTYPE)
                '*** Check if the txtIndexValues TextBox control is VISIBLE...
                '    this will handle saving the value of the the TextBox control
                '    instead of the mebIndexValues Masked Edit control as needed.
                '    This is because the Masked Edit control has a MAX size of 64 Char.
                If frmIndex.txtIndexValues(intIndex).Visible = True Then
                    frmIndex.txtIndexValues(intIndex).Text = ""
                Else
                    frmIndex.mebIndexValues(intIndex).Text = ""
                End If
                    
        End Select
    Next

End Sub

Private Sub subCheckForAutoAdvance()

            '*** If the BOOKMARK is ALREADY SET -- GET OUT OF HERE NOW!
            '    We are already in the middle of a Copy operation
            '    Otherwise, the Bookmark/Copy will be in an inconsistent state
            If blnBookMark = True Then
                Exit Sub
            End If
            

            'Automatically move the the Next Page or the Set # of PAGES if chkMoveToNextPageOnSingleClick is checked
            If chkMoveToNextPageOnSingleClick = vbChecked Then
'                For intPagesLoop = 1 To CInt(strPAGES)
'                    frmIndex.cmdNextImage_Click
'                Next
                intPageIndex = frmIndex.ListView1.SelectedItem.Index
                If CInt(strPAGES) <= (frmIndex.ListView1.ListItems.count) - intPageIndex Then
                    'We have enough pages left in Batch for this DocType
                    'First SET the BookMark
                    frmIndex.cmdBookMark_Set blnNoPrompt
                    'Set the LAST Page we have to index
                    intPageIndex = frmIndex.ListView1.SelectedItem.Index + CInt(strPAGES) - 1
                    frmIndex.ListView1.ListItems.item(intPageIndex).Selected = True
                    ' Copy the Index values
                    frmIndex.cmdBookMark_Set blnNoPrompt
                    ' Now move to the Next Page
'                    intPageIndex = frmIndex.ListView1.SelectedItem.Index + 1
'                    frmIndex.ListView1.ListItems.item(intPageIndex).Selected = True
'                    frmIndex.ListView1_Click
                    
                    ' Finally, Move to the NEXT page
                    frmIndex.cmdNextImage_Click
                Else
                    'DocType wants more than pages left in Batch
                    frmIndex.cmdBookMark_Set blnNoPrompt
                    'Set the LAST Page we have to index equal to the LAST Page in the Batch
                    intPageIndex = frmIndex.ListView1.ListItems.count
                    frmIndex.ListView1.ListItems.item(intPageIndex).Selected = True
                    ' Copy the Index values
                    frmIndex.cmdBookMark_Set blnNoPrompt
                    ' Finally, Force the Commit Prompt because we indexed the LAST Page
                    frmIndex.cmdNextImage_Click
                End If
                
            End If
            
End Sub
