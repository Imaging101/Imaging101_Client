VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmDOCTYPES 
   Caption         =   "DOCTYPES"
   ClientHeight    =   7275
   ClientLeft      =   285
   ClientTop       =   735
   ClientWidth     =   11370
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7275
   ScaleWidth      =   11370
   Begin VB.PictureBox picFields 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1620
      Left            =   0
      ScaleHeight     =   1620
      ScaleWidth      =   11370
      TabIndex        =   20
      Top             =   5055
      Width           =   11370
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   7080
         TabIndex        =   30
         Top             =   840
         Width           =   4095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   7080
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   1200
         Width           =   1575
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   7080
         TabIndex        =   29
         Top             =   480
         Width           =   4095
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   7080
         TabIndex        =   28
         Top             =   120
         Width           =   4095
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1200
         TabIndex        =   27
         Top             =   480
         Width           =   3375
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1200
         TabIndex        =   26
         Top             =   120
         Width           =   3615
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000016&
         Caption         =   "Average # of Pages"
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
         Left            =   5160
         TabIndex        =   32
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000016&
         Caption         =   "Document Description"
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
         Left            =   5160
         TabIndex        =   25
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000016&
         Caption         =   "Document Type"
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
         Left            =   5160
         TabIndex        =   24
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000016&
         Caption         =   "Document Group"
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
         Left            =   5160
         TabIndex        =   23
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000016&
         Caption         =   "Area"
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
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000016&
         Caption         =   "Application"
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
         Left            =   120
         TabIndex        =   21
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.ComboBox cmbOrderBy 
      Height          =   315
      ItemData        =   "frmDOCTYPES.frx":0000
      Left            =   1800
      List            =   "frmDOCTYPES.frx":001F
      TabIndex        =   17
      Text            =   "APPLICATION,AREA,DOCGROUP,DOCTYPE,FORMDESC"
      Top             =   600
      Width           =   6735
   End
   Begin VB.ComboBox cmbApplicationList 
      Height          =   315
      Left            =   1800
      TabIndex        =   15
      Top             =   240
      Width           =   4935
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   11370
      TabIndex        =   7
      Top             =   6675
      Width           =   11370
      Begin VB.CommandButton cmdDuplicate 
         Caption         =   "D&uplicate"
         Height          =   300
         Left            =   6120
         TabIndex        =   19
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   300
         Left            =   1213
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   59
         TabIndex        =   13
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   4675
         TabIndex        =   12
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   3521
         TabIndex        =   11
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   2367
         TabIndex        =   10
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   300
         Left            =   1213
         TabIndex        =   9
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   300
         Left            =   59
         TabIndex        =   8
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox picStatBox 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   11370
      TabIndex        =   1
      Top             =   6975
      Width           =   11370
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Picture         =   "frmDOCTYPES.frx":0185
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Picture         =   "frmDOCTYPES.frx":04C7
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Picture         =   "frmDOCTYPES.frx":0809
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "frmDOCTYPES.frx":0B4B
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   6
         Top             =   0
         Width           =   3360
      End
   End
   Begin MSDataGridLib.DataGrid grdDataGrid 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   11370
      _ExtentX        =   20055
      _ExtentY        =   6165
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
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
   Begin VB.Label Label1 
      Caption         =   "Order By"
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
      TabIndex        =   18
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label lblSelectApplication 
      Caption         =   "Select &Application"
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
      TabIndex        =   16
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmDOCTYPES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents adoPrimaryRS As Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

Dim intRow
Dim intColumn

Dim db As Connection






Private Sub cmbApplicationList_Click()

    subPopulateDataGrid


End Sub



Private Sub cmbOrderBy_Click()

    subPopulateDataGrid
    

End Sub

Private Sub cmdDuplicate_Click()

  On Error GoTo AddErr
  
  'Bookmark the row the user is currently on so we can copy it's values
  grdDataGrid.RowBookmark (grdDataGrid.Row)
  
  'Set in Add Mode
  cmdAdd_Click
  
  For intLoop = 0 To grdDataGrid.Columns.Count - 1
         'Show the contents of the cell in a textbox
         If IsNull(grdDataGrid.Columns(intLoop).CellValue(mvBookMark)) Then
            grdDataGrid.Columns(intLoop) = ""
         Else
            grdDataGrid.Columns(intLoop) = grdDataGrid.Columns(intLoop).CellValue(mvBookMark)
         End If
  Next


  Exit Sub
AddErr:
  MsgBox Err.Description



End Sub

Private Sub Form_Load()

    subLoadApplicationDropDown
    cmbApplicationList.Text = frmConfig.txtApplicationName
    
    subPopulateDataGrid
    
      mbDataChanged = False
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  'This will resize the grid when the form is resized
  grdDataGrid.Height = Me.ScaleHeight - grdDataGrid.Top - 30 - picButtons.Height - picStatBox.Height - picFields.Height
  lblStatus.Width = Me.Width - 1500
  cmdNext.Left = lblStatus.Width + 700
  cmdLast.Left = cmdNext.Left + 340
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If mbEditFlag Or mbAddNewFlag Then Exit Sub

  Select Case KeyCode
    Case vbKeyEscape
      cmdClose_Click
    Case vbKeyEnd
      cmdLast_Click
    Case vbKeyHome
      cmdFirst_Click
    Case vbKeyUp, vbKeyPageUp
      If Shift = vbCtrlMask Then
        cmdFirst_Click
      Else
        cmdPrevious_Click
      End If
    Case vbKeyDown, vbKeyPageDown
      If Shift = vbCtrlMask Then
        cmdLast_Click
      Else
        CmdNext_Click
      End If
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  If adReason = adRsnUndoAddNew Then
  End If
  
  On Error GoTo ERROR_HANDLER
  'This will display the current record position for this recordset
  lblStatus.Caption = "Record: " & CStr(adoPrimaryRS.AbsolutePosition)

Exit Sub

ERROR_HANDLER:
'    adoPrimaryRS.MoveFirst
    Resume Next
    
End Sub

Private Sub adoPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  
    Debug.Print "Private Sub adoPrimaryRS_WillChangeRecord"
    
  'This is where you put validation code
  'This event gets called when the following actions occur
  Dim bCancel As Boolean
  
  On Error GoTo ERROR_HANDLER
  

  Select Case adReason
    Case adRsnAddNew
          Debug.Print "Case adRsnAddNew"
    Case adRsnClose
          Debug.Print "Case adRsnClose"
    Case adRsnDelete
          Debug.Print "Case adRsnDelete"
    Case adRsnFirstChange
          Debug.Print "Case adRsnFirstChange"
    Case adRsnMove
          Debug.Print "Case adRsnMove"
    Case adRsnRequery
          Debug.Print "Case adRsnRequery"
    Case adRsnResynch
          Debug.Print "Case adRsnResynch"
    Case adRsnUndoAddNew
          Debug.Print "Case adRsnUndoAddNew"
    Case adRsnUndoDelete
          Debug.Print "Case adRsnUndoDelete"
    Case adRsnUndoUpdate
          Debug.Print "Case adRsnUndoUpdate"
    Case adRsnUpdate
          Debug.Print "Case adRsnUpdate"
      
  End Select

  If bCancel Then adStatus = adStatusCancel
  
Exit Sub

ERROR_HANDLER:
    MsgBox Err.Description, vbOKOnly
'    Resume Next

  
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  adoPrimaryRS.MoveLast
'  adoPrimaryRS.AddNew
  'Instead of the AddNew... Position the Cursor at the "*" row...
  '  This is usually ready to add a row
  grdDataGrid.Row = grdDataGrid.Row + 1
  grdDataGrid.SetFocus
  SetButtons False

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
  With adoPrimaryRS
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  On Error GoTo RefreshErr
  Set grdDataGrid.DataSource = Nothing
  adoPrimaryRS.Requery
  Set grdDataGrid.DataSource = adoPrimaryRS

  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdEdit_Click()
  On Error GoTo EditErr

  lblStatus.Caption = "Edit record"
  mbEditFlag = True
  SetButtons False
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub
Private Sub cmdCancel_Click()

  On Error Resume Next

  SetButtons True
  mbEditFlag = False
  mbAddNewFlag = False
  
  adoPrimaryRS.CancelUpdate
  
  If Err.Number <> 0 Then
    Exit Sub
  End If
  
  If mvBookMark > 0 Then
    adoPrimaryRS.Bookmark = mvBookMark
  Else
    adoPrimaryRS.MoveFirst
  End If
  mbDataChanged = False

End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr

  adoPrimaryRS.UpdateBatch adAffectAll

  If mbAddNewFlag Then
    adoPrimaryRS.MoveLast              'move to the new record
    grdDataGrid.Row = grdDataGrid.Row + 1
    grdDataGrid.SetFocus
  
  End If

  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False

  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError

  adoPrimaryRS.MoveFirst
  mbDataChanged = False

  Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  On Error GoTo GoLastError

  adoPrimaryRS.MoveLast
  mbDataChanged = False

  Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub CmdNext_Click()
  On Error GoTo GoNextError

  If Not adoPrimaryRS.EOF Then adoPrimaryRS.MoveNext
  If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
     'moved off the end so go back
    adoPrimaryRS.MoveLast
  End If
  'show the current record
  mbDataChanged = False

  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
  On Error GoTo GoPrevError

  If Not adoPrimaryRS.BOF Then adoPrimaryRS.MovePrevious
  If adoPrimaryRS.BOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    adoPrimaryRS.MoveFirst
  End If
  'show the current record
  mbDataChanged = False

  Exit Sub

GoPrevError:
  MsgBox Err.Description
End Sub

Private Sub SetButtons(bVal As Boolean)
  cmdAdd.Visible = bVal
  cmdEdit.Visible = bVal
  cmdUpdate.Visible = Not bVal
  cmdCancel.Visible = Not bVal
  cmdDelete.Visible = bVal
  cmdClose.Visible = bVal
  cmdRefresh.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
End Sub


Private Sub subLoadApplicationDropDown()

    '***************************************
    '*** LOAD APPLICATION LIST DROP-DOWN
    
    Dim Con As ADODB.Connection
    Set Con = New ADODB.Connection
    Con.Open RegImaging101ConnectionString
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = Con
    
    rs.Source = "Select ApplicationName, ApplicationRECID from I101Applications ORDER BY ApplicationName"
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LockType = adLockReadOnly
    
    Con.Errors.Clear
    
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
    Con.Close
    Set Con = Nothing

    '****************************
End Sub

Private Sub subPopulateDataGrid()
        
      Dim strSQL As String
      Dim strSELECT As String
      Dim strWhere As String
      Dim strORDERBY As String
      
      Set db = New Connection
      db.CursorLocation = adUseClient
      
      db.ConnectionString = RegImaging101ConnectionString
      db.Open
    
      Set adoPrimaryRS = New Recordset
      
      strSELECT = "select APPLICATION,AREA,DOCGROUP,DOCTYPE,FORMDESC,PAGES from DOCTYPES "
      If Trim(cmbApplicationList.Text) <> "" Then
        strWhere = "WHERE APPLICATION = '" & cmbApplicationList.Text & "'"
      End If
      
      If cmbOrderBy.Text <> "" Then
        strORDERBY = "Order by " & cmbOrderBy.Text
      End If
      
      strSQL = strSELECT & " " & strWhere & " " & strORDERBY
      
      With adoPrimaryRS
        .ActiveConnection = db
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .Source = strSQL
      End With
      
      adoPrimaryRS.Open
    
      Set grdDataGrid.DataSource = adoPrimaryRS

End Sub

Private Sub grdDataGrid_AfterUpdate()
    Debug.Print "grdDataGrid_AfterUpdate"
    'Reset buttons
    SetButtons True
End Sub

Private Sub grdDataGrid_BeforeDelete(Cancel As Integer)

    Result = MsgBox("Are you sure you wish to Delete this Row?", vbYesNo, "Delete Row")
    If Result = vbNo Then
        Cancel = True
    End If
    
End Sub

Private Sub grdDataGrid_BeforeUpdate(Cancel As Integer)

    Debug.Print "grdDataGrid_BeforeUpdate"
    
End Sub

Private Sub grdDataGrid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    intColumn = grdDataGrid.ColContaining(X)   'Set the Cell Column number
    intRow = grdDataGrid.RowContaining(Y)      'Set the Cell Row number

    On Error Resume Next  ' Ignore errors after adding rows.
    mvBookMark = grdDataGrid.RowBookmark(intRow) 'Set Bookmark Value


End Sub

Private Sub grdDataGrid_OnAddNew()
    Debug.Print "grdDataGrid_OnAddNew"
    mbAddNewFlag = True

End Sub

Private Sub grdDataGrid_Validate(Cancel As Boolean)

    Debug.Print "grdDataGrid_Validate"

End Sub
