VERSION 5.00
Begin VB.Form frmDownloadJMH 
   BackColor       =   &H00008000&
   Caption         =   "Download JMH Tables"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   9645
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Imaging101"
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   9015
      Begin VB.TextBox txtImaging101Docs 
         Height          =   285
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   8655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "QuickIndex"
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   9015
      Begin VB.TextBox txtQuickIndexDocs 
         Height          =   285
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   8655
      End
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H0080FF80&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4560
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   9375
   End
   Begin VB.TextBox txtActionBeforeError 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   2760
      Width           =   9375
   End
   Begin VB.CommandButton cmdDownloadTablesNow 
      BackColor       =   &H0080FF80&
      Caption         =   "&Download Tables Now"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4440
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmDownloadJMH.frx":0000
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Processing Record"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Download JMH Tables"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   840
      TabIndex        =   3
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "frmDownloadJMH"
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
    
''    On Error GoTo DOWNLOAD_JMH_DATA_FIELDS_ERROR
    
    Set connImaging101DocList = New ADODB.Connection
    Set cmdImaging101DocList = New ADODB.Command
    Set rsImaging101DocList = New ADODB.Recordset
    
    Set connQuickIndexDocList = New ADODB.Connection
    Set cmdQuickIndexDocList = New ADODB.Command
    Set rsQuickIndexDocList = New ADODB.Recordset
    
    '******************************************************************
    '   UPDATE DOCUMENT TYPES
    '******************************************************************
    
    
    '*** Connect to Imaging101DocList DB
    connImaging101DocList.ConnectionString = txtImaging101Docs
    connImaging101DocList.ConnectionTimeout = 120
    connImaging101DocList.mode = adModeReadWrite
    connImaging101DocList.Open
    
    '*** Connect to Imaging101DocList DB
    connQuickIndexDocList.ConnectionString = txtQuickIndexDocs
    connQuickIndexDocList.ConnectionTimeout = 120
    connQuickIndexDocList.mode = adModeRead
    connQuickIndexDocList.Open
        
        txtActionBeforeError = "ZAP ALL RECORDS FROM Document Types Table"
        Text1 = ""
        DoEvents
        Set rsImaging101DocList = New ADODB.Recordset
        
        ' ZAP all entries EXCEPT  Separator and Questionable
        rsImaging101DocList.Open "DELETE FROM DOCTYPES WHERE (DOCTYPES.DOCTYPE<>'*??????????*') AND (DOCTYPES.DOCTYPE<>'*SEPARATOR SHEET*') AND (DOCTYPES.DOCTYPE<>'*DO NOT FILE*')", connImaging101DocList, adOpenDynamic, adLockOptimistic
        Set rsImaging101DocList = Nothing
           
           
        txtActionBeforeError = "Open Imaging101 DOCUMENT TYPES Connection"
        Set rsImaging101DocList = New ADODB.Recordset
        rsImaging101DocList.Open "DOCTYPES", connImaging101DocList, adOpenDynamic, adLockOptimistic
        
        
        Set rsQuickIndexDocList = New ADODB.Recordset
        txtSource = "select *  from JMHDOCS order by ID"
        txtActionBeforeError = "Open QuickIndex DOCUMENT TYPES Table"
        rsQuickIndexDocList.Open txtSource, connQuickIndexDocList
         
         
        While Not rsQuickIndexDocList.EOF()
            rsImaging101DocList.AddNew
            rsImaging101DocList("APPLICATION") = Replace(rsQuickIndexDocList("Application"), " ", "_")
            rsImaging101DocList("AREA") = rsQuickIndexDocList("Area")
            rsImaging101DocList("DOCGROUP") = rsQuickIndexDocList("Subfolder")
            rsImaging101DocList("DOCTYPE") = rsQuickIndexDocList("FormName")
            rsImaging101DocList("FORMDESC") = rsQuickIndexDocList("FormName")
            rsImaging101DocList.Update
            Text1 = rsQuickIndexDocList("ID") & " - " & rsQuickIndexDocList("Area") & " - " & rsQuickIndexDocList("FormName")
            DoEvents
            rsQuickIndexDocList.MoveNext
        Wend
        
        txtActionBeforeError = "Commit Document Types Updates to LOCAL Database"
        DoEvents
    txtActionBeforeError = "DOCUMENT TYPES UPDATE COMPLETE!"
    Text1 = ""
    DoEvents
      
    
    
    rsQuickIndexDocList.Close
    Set rsQuickIndexDocList = Nothing
    
    rsImaging101DocList.Close
    Set rsImaging101DocList = Nothing
    
    cmdDownloadTablesNow.enabled = True
    cmdCancel.enabled = True
  
    
Exit Sub

DOWNLOAD_JMH_DATA_FIELDS_ERROR:
        MsgBox "DOWNLOAD JMH DATA FIELDS: " & Err.Number & " - " & Err.Description & "  DURING ACTION: (" & txtActionBeforeError & ") - [Transaction Rolled Back - Data NOT Imported]", vbExclamation
        
        Screen.MousePointer = vbDefault
End Sub


Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    '*** Imaging101 Tables
    RegImaging101DocListConnectionType = VBGetPrivateProfileString("DATABASE", "AdodcDocTypeList.ConnectionType", "Imaging101.INI")
    RegImaging101DocListConnectionString = VBGetPrivateProfileString("DATABASE", "AdodcDocTypeList.ConnectionString." & RegImaging101DocListConnectionType, "Imaging101.INI")
    txtImaging101Docs = RegImaging101DocListConnectionString
    
    '*** QuickIndex Tables
    RegQuickIndexDocListConnectionType = VBGetPrivateProfileString("DATABASE", "AdodcDocTypeList.ConnectionType", "QuickIndex.INI")
    RegQuickIndexDocListConnectionString = VBGetPrivateProfileString("DATABASE", "AdodcDocTypeList.ConnectionString." & RegQuickIndexDocListConnectionType, "QuickIndex.INI")
    txtQuickIndexDocs = RegQuickIndexDocListConnectionString
    
    On Error GoTo 0
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMainMenu.Show
    
End Sub
