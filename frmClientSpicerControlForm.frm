VERSION 5.00
Object = "{71C182E1-878D-11D1-8108-020701190C00}#8.0#0"; "view.ocx"
Object = "{22B7B2BB-4EFA-11D2-81FC-0000D1108734}#8.0#0"; "Edit.ocx"
Object = "{895CDC7A-8837-11D1-8109-020701190C00}#8.0#0"; "docctrl.ocx"
Object = "{C8B15BE2-E8D8-11D1-818A-0000D1108734}#8.0#0"; "SpiConfg.ocx"
Begin VB.Form frmClientSpicerControlForm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "frmSpicerControlForm"
   ClientHeight    =   4470
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4035
   LinkTopic       =   "frmSpicerControlForm"
   ScaleHeight     =   4470
   ScaleWidth      =   4035
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin SPICERDOCUMENTLib.SpicerDoc SpicerDoc2 
      Left            =   615
      Top             =   3000
      _Version        =   524288
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
   Begin SPICERVIEWLib.SpicerView SpicerView1 
      Height          =   2760
      Left            =   180
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1995
      _Version        =   524288
      _ExtentX        =   3519
      _ExtentY        =   4868
      _StockProps     =   0
   End
   Begin SPICEREDITLib.SpicerEdit SpicerEdit1 
      Left            =   1185
      Top             =   3000
      _Version        =   524288
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
   Begin SPICERCONFIGURATIONLib.SpicerConfiguration SpicerConfiguration1 
      Left            =   1740
      Top             =   3000
      _Version        =   524288
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
   Begin SPICERDOCUMENTLib.SpicerDoc SpicerDoc1 
      Left            =   75
      Top             =   3015
      _Version        =   524288
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmClientSpicerControlForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
    
    '*** 2021-03-31 - Jacob - Enable BatchMessageMode and On Error to allow errors to flow through
    SpicerConfiguration1.BatchMessageMode = True
    
     '*** 2021-03-31 - Jacob - Added ERROR_HANDLER to log Unload Errors
   On Error GoTo ERROR_HANDLER
    
    'frmExportFromArchive.poSendMail_Status "+++ frmClientSpicerControlForm - Form_Unload()  | SpicerDoc1.CloseDocument (False)"
    SpicerDoc1.CloseDocument (False)
    DoEvents
    
Exit Sub

ERROR_HANDLER:
    
    On Error Resume Next

    bolErrorOccured = True
    
'    frmExportFromArchive.poSendMail_Status "+++ frmClientSpicerControlForm - Form_Unload()  ERROR:   Error #: " & Err.Number & " - " & Err.Description
    funcWriteToDebugLog Me.name, "+++ frmClientSpicerControlForm - Form_Unload()  ERROR:   Error #: " & Err.Number & " - " & Err.Description
'    funcWriteToSystemEventLog frmImaging101AutoExport.NTService1, svcMessageError, "+++ frmClientSpicerControlForm - Form_Unload()  ERROR:   Error #: " & Err.Number & " - " & Err.Description
    
    frmExportFromArchive.subErrorLoggingHandler
    
    DoEvents

End Sub
