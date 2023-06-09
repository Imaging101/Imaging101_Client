VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFileDialog 
   Caption         =   "File Dialog"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   8616
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   7.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3480
   ScaleWidth      =   8616
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCommonDialogueIndexRootDir 
      Caption         =   "..."
      Height          =   255
      Left            =   6120
      TabIndex        =   2
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox txtFileSpec 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Text            =   "*.*"
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtFilePath 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Text            =   "C:\SampleImages\Spicer"
      Top             =   360
      Width           =   4695
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmFileDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCommonDialogueIndexRootDir_Click()
   With CommonDialog1
      .MaxFileSize = 2048   ' Set as appropriate
      .FileName = ""
      .Filter = txtFileSpec    ' "All Files|*.*"
      .flags = cdlOFNAllowMultiselect + cdlOFNExplorer
      .ShowOpen
      Me.txtFilePath = .FileName & vbNullChar
   End With

End Sub

Private Sub Form_Load()
    Me.txtFilePath = txtBatchRootDir & "\" & txtecFileCabinetID & "\" & txtecID
    Me.txtFileSpec = txtecID & ".*"

End Sub
