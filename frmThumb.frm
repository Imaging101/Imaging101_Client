VERSION 5.00
Object = "{1EBDE7A6-8547-11D2-869A-0000929B139D}#8.0#0"; "Thumb.ocx"
Begin VB.Form frmThumb 
   Caption         =   "Thumbnail Window"
   ClientHeight    =   2145
   ClientLeft      =   60
   ClientTop       =   8940
   ClientWidth     =   9330
   Icon            =   "frmThumb.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   2145
   ScaleWidth      =   9330
   Begin SPICERTHUMBNAILLib.SpicerThumbnail SpicerThumbnail1 
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   615
      _Version        =   524288
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   0
   End
   Begin VB.Timer Timer1 
      Left            =   960
      Top             =   1320
   End
   Begin VB.Label Label3 
      Caption         =   "the longer it takes!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   6855
   End
   Begin VB.Label Label2 
      Caption         =   "The larger the window size,"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6855
   End
   Begin VB.Label Label1 
      Caption         =   "Re-calculating thumbnail sizes..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5175
   End
End
Attribute VB_Name = "frmThumb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bolFormIsLoaded  As String

Private Sub Form_Activate()

    bolFormIsLoaded = True

End Sub

Private Sub Form_Click()

    result = MainMDIForm.funcZoomToSavedFactor

    If bolErrorOccured Then
            MsgBox "funcZoomToSavedFactor() ERROR:" & vbCrLf & " While attempting to set Saved ZOOM Factor." & vbCrLf & vbCrLf & result, vbCritical
    End If
    


End Sub

Private Sub Form_Load()

    ' This flag is to make sure we don't recalculate the thumbnails twice when the form resizes!
    bolFormIsLoaded = False
    
    ' Get saved settings from the registry
    On Error Resume Next
    Me.Top = VBGetPrivateProfileString(RegAppname, "frmThumb.Top", RegFileName)
    Me.Left = VBGetPrivateProfileString(RegAppname, "frmThumb.Left", RegFileName)
    Me.width = VBGetPrivateProfileString(RegAppname, "frmThumb.Width", RegFileName)
    Me.Height = VBGetPrivateProfileString(RegAppname, "frmThumb.Height", RegFileName)


    SpicerThumbnail1.Left = 0
    SpicerThumbnail1.Top = 0
    SpicerThumbnail1.width = ScaleWidth
    SpicerThumbnail1.Height = ScaleHeight


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Save Form settings to the registry
    If Me.Top > 0 And Me.Left > 0 Then
        result = WritePrivateProfileString(RegAppname, "frmThumb.Top", Me.Top, RegFileName)
        result = WritePrivateProfileString(RegAppname, "frmThumb.Left", Me.Left, RegFileName)
        result = WritePrivateProfileString(RegAppname, "frmThumb.Width", Me.width, RegFileName)
        result = WritePrivateProfileString(RegAppname, "frmThumb.Height", Me.Height, RegFileName)
    End If
    
End Sub

Private Sub Form_Resize()
    
'    If Me.Height > Me.Width Then
'        Me.Width = 3500
'    Else
'        Me.Height = 3500
'    End If
    
    If bolFormIsLoaded = True Then
    
        SpicerThumbnail1.Visible = False
        
        Timer1.enabled = False
        Timer1.Interval = 3000
        Timer1.enabled = True
    End If
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

''    Set frmThumb = Nothing
    frmThumb.SpicerThumbnail1.BindToViewControl Nothing
    
End Sub



Private Sub Timer1_Timer()

    Timer1.enabled = False
    
    frmThumb.SpicerThumbnail1.BindToViewControl Nothing
    SpicerThumbnail1.width = ScaleWidth
    SpicerThumbnail1.Height = ScaleHeight
     
    SpicerThumbnail1.Visible = True

    frmThumb.SpicerThumbnail1.BindToViewControl MainMDIForm.ActiveForm.SpicerView1.object

    
End Sub
