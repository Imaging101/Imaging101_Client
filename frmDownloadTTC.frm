VERSION 5.00
Begin VB.Form frmDownloadTTC 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Download TTC Tables"
   ClientHeight    =   5292
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   9648
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   7.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDownloadTTC.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5292
   ScaleWidth      =   9648
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4320
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   2
      Top             =   3120
      Width           =   2295
   End
   Begin VB.TextBox txtActionBeforeError 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   960
      Width           =   7095
   End
   Begin VB.CommandButton cmdDownloadTablesNow 
      Caption         =   "&Download Tables Now"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4320
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Processing Record #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   3360
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   384
      Left            =   240
      Picture         =   "frmDownloadTTC.frx":0442
      Top             =   120
      Width           =   384
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Download TTC Tables"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   840
      TabIndex        =   3
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "frmDownloadTTC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDownloadTablesNow_Click()
    
    cmdDownloadTablesNow.enabled = False
    cmdCancel.enabled = False
    
    ' Get Database Connections settings from the registry
    txtActionBeforeError = "Get fields from Registry"
    DoEvents
    On Error Resume Next
    RegLookupListConnectionType = VBGetPrivateProfileString("DATABASE", "AdodcLookupList.ConnectionType", RegFileName)
    RegLookupListConnectionString = VBGetPrivateProfileString("DATABASE", "AdodcLookupList.ConnectionString." & RegLookupListConnectionType, RegFileName)
    
    strAccountTableName = VBGetPrivateProfileString(RegAppname, "frmLookupList.AccountTableName", RegFileName)
    
    strAccountNumberDbField = VBGetPrivateProfileString(RegAppname, "frmLookupList.AccountNumberDbField", RegFileName)
    strAccountNameDbField = VBGetPrivateProfileString(RegAppname, "frmLookupList.AccountNameDbField", RegFileName)
    strAccountCaseDbField = VBGetPrivateProfileString(RegAppname, "frmLookupList.AccountCaseDbField", RegFileName)
    strAccountCaseIdDbField = VBGetPrivateProfileString(RegAppname, "frmLookupList.AccountCaseIdDbField", RegFileName)
    strAccountTicketNumberDbField = VBGetPrivateProfileString(RegAppname, "frmLookupList.AccountTicketNumberDbField", RegFileName)
    
    RegDocListConnectionType = VBGetPrivateProfileString("DATABASE", "AdodcDocTypeList.ConnectionType", RegFileName)
    RegDocListConnectionString = VBGetPrivateProfileString("DATABASE", "AdodcDocTypeList.ConnectionString." & RegDocListConnectionType, RegFileName)
    
    RegTTCConnectionType = VBGetPrivateProfileString("DATABASE", "AdodcTTC.ConnectionType", RegFileName)
    RegTTCConnectionString = VBGetPrivateProfileString("DATABASE", "AdodcTTC.ConnectionString." & RegTTCConnectionType, RegFileName)
    On Error GoTo 0
    
    
    On Error GoTo DOWNLOAD_TTC_DATA_FIELDS_ERROR
    
    Set connLookupList = New ADODB.Connection
    Set cmdLookupList = New ADODB.Command
    Set rsLookupList = New ADODB.Recordset
    
    Set connDocList = New ADODB.Connection
    Set cmdDocList = New ADODB.Command
    Set rsDocList = New ADODB.Recordset
    
    Set connTTC = New ADODB.Connection
    Set cmdTTC = New ADODB.Command
    Set rsTTC = New ADODB.Recordset
    

    
    '******************************************************************
    '   UPDATE CLIENTS / CASES
    '******************************************************************
    
    '*** Connect to LookupList DB
    connLookupList.ConnectionString = RegLookupListConnectionString
    connLookupList.ConnectionTimeout = 120
    connLookupList.mode = adModeReadWrite
    connLookupList.Open
    
        txtActionBeforeError = "ZAP ALL RECORDS FROM Client Table"
        Text1 = ""
        DoEvents
        Set rsLookupList = New ADODB.Recordset
        connLookupList.BeginTrans
            rsLookupList.Open "DELETE FROM " & strAccountTableName, connLookupList, adOpenDynamic, adLockOptimistic
        connLookupList.CommitTrans
        
        Set rsLookupList = Nothing
        
        txtActionBeforeError = "Open Client Table"
        Set rsLookupList = New ADODB.Recordset
        rsLookupList.Open strAccountTableName, connLookupList, adOpenDynamic, adLockOptimistic
     
    
    ' Set up for TTC DB connection
    txtActionBeforeError = "Prepare TTC CLIENTS/CASES Database Connections"
    DoEvents
    
    On Error GoTo DOWNLOAD_TTC_DATA_FIELDS_ERROR
    
    connTTC.ConnectionString = RegTTCConnectionString
    connTTC.ConnectionTimeout = 120
    connTTC.mode = adModeReadWrite
    connTTC.Open
    
    Set cmdTTC.ActiveConnection = connTTC
    
    txtActionBeforeError = "Requesting Client & Case Information from TTC Database"
    DoEvents
    Set rsTTC = New ADODB.Recordset
    txtSource = "select clients.id as client_id, clients.lname, clients.fname, clients.mi, clients.address1, clients.city, clients.state, cases.casenumber, cases.id as case_id, cases.ticketnumber  from cases, clients where clients.id=cases.clientid order by client_id"
    rsTTC.Open txtSource, connTTC
    
    txtActionBeforeError = "Update LOCAL Database with TTC CLIENTS/CASES Downloaded from TTC Database"
    DoEvents

    Me.SetFocus
    
    connLookupList.BeginTrans
    
    While Not rsTTC.EOF()
            rsLookupList.AddNew
            rsLookupList(strAccountNumberDbField) = rsTTC("client_id")
            rsLookupList(strAccountNameDbField) = rsTTC("lname") & ", " & rsTTC("fname") & " " & rsTTC("mi") & " / " & rsTTC("address1") & ", " & rsTTC("city") & ", " & rsTTC("state")
            rsLookupList(strAccountCaseDbField) = rsTTC("casenumber")
            rsLookupList(strAccountCaseIdDbField) = rsTTC("case_id")
            rsLookupList(strAccountTicketNumberDbField) = rsTTC("ticketnumber")
            rsLookupList.Update
            Text1 = rsTTC("client_id")
            DoEvents
        rsTTC.MoveNext
    Wend
    
    txtActionBeforeError = "Commit Updates to LOCAL Database"
    DoEvents
    connLookupList.CommitTrans
    txtActionBeforeError = "TTC CLIENTS/CASES UPDATE COMPLETE!"
    DoEvents
    
    rsTTC.Close
    Set rsTTC = Nothing
    
    rsLookupList.Close
    Set rsLookupList = Nothing
    
    
    
    '******************************************************************
    '   UPDATE DOCUMENT TYPES
    '******************************************************************
    
    
    '*** Connect to LookupList DB
    connDocList.ConnectionString = RegDocListConnectionString
    connDocList.ConnectionTimeout = 120
    connDocList.mode = adModeReadWrite
    connDocList.Open
        
    connDocList.BeginTrans
            
        txtActionBeforeError = "ZAP ALL RECORDS FROM Document Types Table"
        Text1 = ""
        DoEvents
        Set rsDocList = New ADODB.Recordset
        ' ZAP all entries EXCEPT  Separator and Questionable
        rsDocList.Open "DELETE FROM DOCTYPES WHERE (DOCTYPES.DOCTYPE<>'*??????????*') AND (DOCTYPES.DOCTYPE<>'*SEPARATOR SHEET*') AND (DOCTYPES.DOCTYPE<>'*DO NOT FILE*')", connDocList, adOpenDynamic, adLockOptimistic
        Set rsDocList = Nothing
           
        txtActionBeforeError = "Open DOCUMENT TYPES Connection"
        Set rsDocList = New ADODB.Recordset
        rsDocList.Open "DOCTYPES", connDocList, adOpenDynamic, adLockOptimistic
        
        Set rsTTC = New ADODB.Recordset
        txtSource = "select name  from files_names order by name"
        txtActionBeforeError = "Open DOCUMENT TYPES Table"
        rsTTC.Open txtSource, connTTC
         
        While Not rsTTC.EOF()
            rsDocList.AddNew
            rsDocList("APPLICATION") = "TTC"  'Hard Coded
            rsDocList("AREA") = "*"           'Hard Coded
            rsDocList("DOCGROUP") = Replace(rsTTC("name"), Chr(39), "")
            rsDocList("DOCTYPE") = Replace(rsTTC("name"), Chr(39), "")
            rsDocList("FORMDESC") = Replace(rsTTC("name"), Chr(39), "")
            rsDocList.Update
            Text1 = rsTTC("name")
            DoEvents
            rsTTC.MoveNext
        Wend
        
        txtActionBeforeError = "Commit Document Types Updates to LOCAL Database"
        DoEvents
    connDocList.CommitTrans
    txtActionBeforeError = "DOCUMENT TYPES UPDATE COMPLETE!"
    Text1 = ""
    DoEvents
      
    
    
    rsTTC.Close
    Set rsTTC = Nothing
    
    rsDocList.Close
    Set rsDocList = Nothing
    
    cmdDownloadTablesNow.enabled = True
    cmdCancel.enabled = True
  
    
Exit Sub

DOWNLOAD_TTC_DATA_FIELDS_ERROR:
        MsgBox "DOWNLOAD TTC DATA FIELDS: " & Err.Number & " - " & Err.Description & "  DURING ACTION: (" & txtActionBeforeError & ") - [Transaction Rolled Back - Data NOT Imported]", vbExclamation
        
        If connLookupList.BeginTrans = True Then
            connLookupList.RollbackTrans
        End If
        If connLookupList.BeginTrans = True Then
            connLookupList.RollbackTrans
        End If
        Screen.MousePointer = vbDefault
End Sub


Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMainMenu.Show
    
End Sub
