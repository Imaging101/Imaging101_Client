VERSION 5.00
Begin VB.Form frmDropDownList 
   Caption         =   "Select one"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3225
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
   ScaleHeight     =   4455
   ScaleWidth      =   3225
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstDropDownsList 
      DragIcon        =   "frmDropDownList.frx":0000
      Height          =   4155
      ItemData        =   "frmDropDownList.frx":0442
      Left            =   0
      List            =   "frmDropDownList.frx":0449
      TabIndex        =   0
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "frmDropDownList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objObjectForReturn As Object



Public Function funcPopulateDropDown(ApplicationName As String, FieldName As String, ObjectForReturn As Object) As String
    
    On Error Resume Next
    
    'Set the Module variable so we can use it in the lstDropDownsList_DblClick() sub
    Set objObjectForReturn = ObjectForReturn
    
    '***************************************
    '*** LOAD BATCH QUEUES LIST DROP-DOWN
        
    Dim con As Connection
    Dim rs As Recordset
    Dim txtSelectStatement As String
    
    Set con = New ADODB.Connection
    
    con.ConnectionTimeout = 120
    con.CommandTimeout = 600
    
    con.Open RegImaging101ConnectionString
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    
    rs.Open ApplicationName, con
    
    '2017-03-16 Jacob - Added "AND" to prevent showing items for Deleted. Purged or MovedOut documents
    txtSelectStatement = "SELECT DISTINCT " & FieldName & _
                "  FROM " & ApplicationName & _
                " WHERE " & FieldName & " IS NOT NULL " & _
                "      AND (DocumentLocked IS NULL or DocumentLocked = ''  or DocumentLocked = 'MI')"
    
    'If NOT Numeric, don't show blanks either
    If rs.Fields("" & FieldName & "").Type <> adNumeric Then
        txtSelectStatement = txtSelectStatement & " AND " & FieldName & " <> '' "
    End If
                
    txtSelectStatement = txtSelectStatement & " ORDER BY " & FieldName
                
    rs.Close
    
    rs.Source = txtSelectStatement
    
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LOCKTYPE = adLockReadOnly
    
    con.Errors.Clear
    
    rs.Open
'    rs.MoveFirst
    
    frmDropDownList.lstDropDownsList.Clear
    'Add a Blank value to allow clearing the BatchOwner
    frmDropDownList.lstDropDownsList.AddItem ""
    For intIndex = 0 To rs.RecordCount - 1
        frmDropDownList.lstDropDownsList.AddItem rs.Fields("" & FieldName & "")
        rs.MoveNext
    Next
        
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    con.Close
        Set con = Nothing
    
    

End Function





Private Sub Form_Load()

    ' Get Position settings from the registry
    On Error Resume Next
    Me.Top = VBGetPrivateProfileString(RegAppname, "frmDropDownList.Top", RegFileName)
    Me.Left = VBGetPrivateProfileString(RegAppname, "frmDropDownList.Left", RegFileName)
    Me.width = VBGetPrivateProfileString(RegAppname, "frmDropDownList.Width", RegFileName)
    Me.Height = VBGetPrivateProfileString(RegAppname, "frmDropDownList.Height", RegFileName)
    On Error GoTo 0
    
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
    


End Sub





Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'Save Form settings to the registry, except if minimized (i.e.- Top or Left < 0)
    If Me.Top >= 0 And Me.Left >= 0 Then
        result = WritePrivateProfileString(RegAppname, "frmDropDownList.Top", Me.Top, RegFileName)
        result = WritePrivateProfileString(RegAppname, "frmDropDownList.Left", Me.Left, RegFileName)
        result = WritePrivateProfileString(RegAppname, "frmDropDownList.Width", Me.width, RegFileName)
        result = WritePrivateProfileString(RegAppname, "frmDropDownList.Height", Me.Height, RegFileName)
'''        Result = WritePrivateProfileString(RegAppname, "frmImaging101BatchList.Caption", Me.Caption, RegFileName)
    End If

End Sub

Private Sub Form_Resize()

  lstDropDownsList.Height = Me.ScaleHeight - lstDropDownsList.Top
  lstDropDownsList.width = Me.width - lstDropDownsList.Left - 180

End Sub

Private Sub lstDropDownsList_DblClick()

    On Error GoTo ERROR_HANDLER
    
        Dim intCursorPos As Integer
        Dim intFieldLen As Integer
        Dim txtTextLeft As String
        Dim txtTextRight As String
        Dim txtTextToReturn As String
    
    If Me.Caption = "Field List" And funcIsFormLoaded2("frmImaging101SearchTemplate") = True Then
    
        
        intCursorPos = frmImaging101SearchTemplate.txtWhereFreehand.SelStart
        intFieldLen = Len(lstDropDownsList.Text) + 2 'to account for spaces before and after
        txtTextLeft = Left(frmImaging101SearchTemplate.txtWhereFreehand.Text, intCursorPos)
        txtTextRight = Right(frmImaging101SearchTemplate.txtWhereFreehand.Text, Len(frmImaging101SearchTemplate.txtWhereFreehand.Text) - intCursorPos)
        
        If InStr(UCase(lstDropDownsList.Text), "DATE") = 0 Then
            txtTextToReturn = lstDropDownsList.Text
        ElseIf InStr(UCase(lstDropDownsList.Text), "{") = 0 Then
            'NOT a Special Field
            txtTextToReturn = "CAST(" & lstDropDownsList.Text & " AS DATE ) = 'mm/dd/yyyy' "
        Else
            'Special Field
            txtTextToReturn = "CAST(" & lstDropDownsList.Text & " AS DATE ) "
        End If
        
        frmImaging101SearchTemplate.txtWhereFreehand.Text = txtTextLeft & " " & txtTextToReturn & " " & txtTextRight
        frmImaging101SearchTemplate.SetFocus
        frmImaging101SearchTemplate.txtWhereFreehand.SetFocus
        frmImaging101SearchTemplate.txtWhereFreehand.SelStart = intCursorPos + intFieldLen
    
    ElseIf Me.Caption = "Field List" And frmImaging101Search.txtWhereFreehand.Visible = True Then
    
        intCursorPos = frmImaging101Search.txtWhereFreehand.SelStart
        intFieldLen = Len(lstDropDownsList.Text) + 2 'to account for spaces before and after
        txtTextLeft = Left(frmImaging101Search.txtWhereFreehand.Text, intCursorPos)
        txtTextRight = Right(frmImaging101Search.txtWhereFreehand.Text, Len(frmImaging101Search.txtWhereFreehand.Text) - intCursorPos)
        
        If InStr(UCase(lstDropDownsList.Text), "DATE") = 0 Then
            txtTextToReturn = lstDropDownsList.Text
        Else
            txtTextToReturn = "CAST(" & lstDropDownsList.Text & " AS DATE ) = 'mm/dd/yyyy' "
        End If
        
        frmImaging101Search.txtWhereFreehand.Text = txtTextLeft & " " & txtTextToReturn & " " & txtTextRight
        frmImaging101Search.SetFocus
        frmImaging101Search.txtWhereFreehand.SetFocus
        frmImaging101Search.txtWhereFreehand.SelStart = intCursorPos + intFieldLen
    
    Else
        
        'Set the Objects value using that object's defined FORMAT
        On Error Resume Next
        If Trim(objObjectForReturn.Format) = "" Then
            objObjectForReturn = lstDropDownsList.Text
        Else
            objObjectForReturn = Format(lstDropDownsList.Text, objObjectForReturn.Format)
        End If
        
        'Set the value back to the object
        Me.Hide
        Me.Show
        objObjectForReturn.SetFocus
        
        Unload Me
    End If
    
    
Exit Sub

ERROR_HANDLER:
    
    'Just leave form open
    
    
End Sub


Private Sub lstDropDownsList_KeyPress(KeyAscii As Integer)

    'Catch Enter key
    If KeyAscii = 13 Then
        lstDropDownsList_DblClick
    End If

End Sub
