VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmDirTreeList 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   4635
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   8385
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picImaging101Logo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6600
      Picture         =   "frmDirTreeList.frx":0000
      ScaleHeight     =   375
      ScaleWidth      =   1575
      TabIndex        =   4
      Top             =   120
      Width           =   1572
   End
   Begin ComctlLib.ListView lstFolders 
      Height          =   2892
      Left            =   1200
      TabIndex        =   2
      Top             =   1560
      Width           =   4212
      _ExtentX        =   7435
      _ExtentY        =   5106
      View            =   2
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin ComctlLib.ListView lstPointers 
      Height          =   2892
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   852
      _ExtentX        =   1508
      _ExtentY        =   5106
      View            =   2
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.DirListBox Dir1 
      Height          =   936
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3612
   End
   Begin ComctlLib.ListView lstFiles 
      Height          =   2892
      Left            =   5520
      TabIndex        =   3
      Top             =   1560
      Width           =   2652
      _ExtentX        =   4683
      _ExtentY        =   5106
      View            =   2
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C07000&
      Height          =   225
      Left            =   6600
      TabIndex        =   5
      Top             =   480
      Width           =   1245
   End
End
Attribute VB_Name = "frmDirTreeList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

        lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    
        Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
        

End Sub


Private Sub Dir1_Change()

    Dim arrayFolders() As String
    Dim arrayPointers() As Long
    Dim arrayFilenames() As String
    Dim i As Integer
    
    lstPointers.ListItems.Clear
    lstFiles.ListItems.Clear
    lstFolders.ListItems.Clear
    
    
    PathLogInit
    LogPath Dir1.Path, arrayFolders, arrayPointers, arrayFilenames
    
    For i = 0 To UBound(arrayFilenames) - 1
        Set lstItem = lstPointers.ListItems.Add(, , arrayPointers(i))
        Set lstSubItem = lstItem.ListSubItem.Add(, , arrayPointers(i))

        Set lstItem = lstFiles.ListItems.Add(, , arrayFolders(i))
        Set lstSubItem = lstItem.ListSubItems.Add(, , arrayFolders(i))
        
        Set lstItem = lstFolders.ListItems.Add(, , arrayFilenames(i))
        Set lstSubItem = lstItem.ListSubItems.Add(, , arrayFilenames(i))
        
        DoEvents
    Next
    
End Sub

