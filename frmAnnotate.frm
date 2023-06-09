VERSION 5.00
Begin VB.Form frmAnnotate 
   Caption         =   "Annotate"
   ClientHeight    =   2145
   ClientLeft      =   3960
   ClientTop       =   450
   ClientWidth     =   2175
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
   ScaleHeight     =   2145
   ScaleWidth      =   2175
   Begin VB.CommandButton cmdRedactWhite 
      Caption         =   "White"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      Picture         =   "frmAnnotate.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdRedactBlack 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Black"
      Height          =   735
      Left            =   720
      Picture         =   "frmAnnotate.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton cmdPaste 
      Caption         =   "Paste"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      Picture         =   "frmAnnotate.frx":1994
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdCut 
      Caption         =   "Cut"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      Picture         =   "frmAnnotate.frx":1FFE
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdSelect 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select"
      Height          =   735
      Left            =   0
      Picture         =   "frmAnnotate.frx":2668
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdDrawFreehand 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Free"
      Height          =   735
      Left            =   1440
      Picture         =   "frmAnnotate.frx":2AAA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton cmdCopy 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Copy"
      Height          =   735
      Left            =   1440
      Picture         =   "frmAnnotate.frx":3774
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdMoveResize 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Move"
      Height          =   735
      Left            =   720
      Picture         =   "frmAnnotate.frx":3DDE
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdAnnotate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Note"
      Height          =   735
      Left            =   720
      Picture         =   "frmAnnotate.frx":4220
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton cmdText 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Text"
      Height          =   735
      Left            =   0
      Picture         =   "frmAnnotate.frx":4662
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton cmdLine 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Line"
      Height          =   735
      Left            =   1440
      Picture         =   "frmAnnotate.frx":4AA4
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton cmdHighlightArea 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Highlight"
      Height          =   735
      Left            =   0
      Picture         =   "frmAnnotate.frx":576E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   735
   End
End
Attribute VB_Name = "frmAnnotate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAnnotate_Click()

    Call MainMDIForm.ActiveForm.subResetAnnotationButtons
    cmdAnnotate.BackColor = vbGreen
    
'    ' Get the page ID for current page of active document
'    Dim lPageID As Long
'    Dim lLayerID As Long
'    Dim iLayerNumber As Integer
'
'    iLayerNumber = 1
'    lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(lPageID, iLayerNumber)
'    MainMDIForm.ActiveForm.SpicerMarkup1.ActiveLayer = lLayerID
'    MainMDIForm.ActiveForm.SpicerMarkup1.ActiveTool = IN_TOOL_ANNOTATION

    Call MainMDIForm.ActiveForm.SelectTool(IN_TOOL_ANNOTATION)

    'Set FLAG to show that an Annotation or Modification may have occured.
    bolAnnotationAdded = True

End Sub

Private Sub cmdCopy_Click()

    Call MainMDIForm.ActiveForm.subResetAnnotationButtons
    cmdCopy.BackColor = vbGreen
    
'    ' Get the page ID for current page of active document
'    Dim lPageID As Long
'    Dim lLayerID As Long
'    Dim iLayerNumber As Integer
'
'    iLayerNumber = 1
'    lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(lPageID, iLayerNumber)
'    MainMDIForm.ActiveForm.SpicerMarkup1.ActiveLayer = lLayerID
'    MainMDIForm.ActiveForm.SpicerMarkup1.ActiveTool = IN_TOOL_COPY
    
    Call MainMDIForm.ActiveForm.SelectTool(IN_TOOL_COPY)
    
End Sub

Private Sub cmdCut_Click()


    Call MainMDIForm.ActiveForm.subResetAnnotationButtons
    cmdCut.BackColor = vbGreen
    
'    ' Get the page ID for current page of active document
'    Dim lPageID As Long
'    Dim lLayerID As Long
'    Dim iLayerNumber As Integer
'
'    iLayerNumber = 1
'    lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(lPageID, iLayerNumber)
'    MainMDIForm.ActiveForm.SpicerMarkup1.ActiveLayer = lLayerID
'    MainMDIForm.ActiveForm.SpicerMarkup1.ActiveTool = IN_TOOL_CUT

    Call MainMDIForm.ActiveForm.SelectTool(IN_TOOL_CUT)
    
    'Set FLAG to show that an Annotation or Modification may have occured.
    bolAnnotationAdded = True

End Sub

Private Sub cmdDrawFreehand_Click()

    Call MainMDIForm.ActiveForm.subResetAnnotationButtons
    cmdDrawFreehand.BackColor = vbGreen
    
    Call MainMDIForm.ActiveForm.SelectTool(IN_TOOL_HIGHLIGHT)
    
    'Set FLAG to show that an Annotation or Modification may have occured.
    bolAnnotationAdded = True


End Sub

Private Sub cmdPaste_Click()

    Call MainMDIForm.ActiveForm.subResetAnnotationButtons
    cmdPaste.BackColor = vbGreen
    
'    ' Get the page ID for current page of active document
'    Dim lPageID As Long
'    Dim lLayerID As Long
'    Dim iLayerNumber As Integer
'
'    iLayerNumber = 1
'    lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(lPageID, iLayerNumber)
'    MainMDIForm.ActiveForm.SpicerMarkup1.ActiveLayer = lLayerID
'    MainMDIForm.ActiveForm.SpicerMarkup1.ActiveTool = IN_TOOL_PASTE

    Call MainMDIForm.ActiveForm.SelectTool(IN_TOOL_PASTE)

End Sub
Private Sub cmdDrawBox_Click()

    Call MainMDIForm.ActiveForm.subResetAnnotationButtons
    cmdDrawBox.BackColor = vbGreen
    
   Dim BatchTools As IBatchTools
   Dim lLayerID As Long
   Dim iObjState As VECTOR_STATE
   Dim lVectObjectID As Long
   Dim lParentID As Long
   Dim dX1 As Double
   Dim dY1 As Double
   Dim dX2 As Double
   Dim dY2 As Double
   Dim dThickness As Double
   Dim iThicknessUnits As UNIT_TYPE
   Dim lFrameColor As Long
   Dim lFillColor As Long
   Dim iLineStyle As LINE_STYLE
   Dim iJoinStyle As LINE_JOIN_STYLE
   Dim iFillStyle As FILL_STYLE
   ' Set the object variable for the IBatchtools interface to the Markup Control object

   Set BatchTools = MainMDIForm.ActiveForm.SpicerMarkup1.object
   ' Assign values to the variables
   lLayerID = MainMDIForm.ActiveForm.SpicerMarkup1.ActiveLayer
   ' Find the identifier of the active edit layer
   iObjState = IN_OBJFLAG_VISIBLE + IN_OBJFLAG_NEW   ' A new, visible vector object
   lVectObjectID = 0
   lParentID = 0
   dX1 = 2.3
   dY1 = 3
   dX2 = 3.6
   dY2 = 3.5
   dThickness = 0.005
   iThicknessUnits = IN_UNITS_INCH
   lFrameColor = 8323072   ' Dark red

   lFillColor = 8323072   ' Dark red
   iLineStyle = IN_LINE_SOLID
   iJoinStyle = IN_JOIN_ROUND
   iFillStyle = IN_FILL_TRANSPARENT
   'Place a box on the active edit layer
   BatchTools.PlaceBox lLayerID, iObjState, lVectObjectID, lParentID, dX1, dY1, _
     dX2, dY2, dThickness, iThicknessUnits, lFrameColor, lFillColor, iLineStyle, _
     iJoinStyle, iFillStyle
     
   ' De-initialize the object variable
   Set BatchTools = Nothing
   
   MainMDIForm.ActiveForm.SpicerView1.Refresh
   
   
End Sub

Private Sub cmdHighlightArea_Click()

    Call MainMDIForm.ActiveForm.subResetAnnotationButtons
    
    cmdHighlightArea.BackColor = vbGreen
    
    Call MainMDIForm.ActiveForm.SelectTool(IN_TOOL_HIGHLIGHTAREA)

    'Set the fillcolor to Yellow
    MainMDIForm.ActiveForm.SpicerMarkup1.FillColor(IN_TOOL_HIGHLIGHTAREA, IN_PROPSCOPE_ACTVTOOL, 0, 0) = 16776960

    'Set FLAG to show that an Annotation or Modification may have occured.
    bolAnnotationAdded = True

End Sub

Private Sub cmdLine_Click()
    
    Call MainMDIForm.ActiveForm.subResetAnnotationButtons
    cmdLine.BackColor = vbGreen
    
'    ' Get the page ID for current page of active document
'    Dim lPageID As Long
'    Dim lLayerID As Long
'    Dim iLayerNumber As Integer
'
'    ' Get the page ID for current page of active document
'    lPageID = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
'    iLayerNumber = 2
'    lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(lPageID, iLayerNumber)
'
'    MainMDIForm.ActiveForm.SpicerMarkup1.ActiveLayer = lLayerID
'    MainMDIForm.ActiveForm.SpicerMarkup1.ActiveTool = IN_TOOL_LINE

    Call MainMDIForm.ActiveForm.SelectTool(IN_TOOL_ARROW)
    
    'Set FLAG to show that an Annotation or Modification may have occured.
    bolAnnotationAdded = True


End Sub

Private Sub cmdMoveResize_Click()
    
    Call MainMDIForm.ActiveForm.subResetAnnotationButtons
    cmdMoveResize.BackColor = vbGreen
    
'    ' Get the page ID for current page of active document
'    Dim lPageID As Long
'    Dim lLayerID As Long
'    Dim iLayerNumber As Integer
'
'    iLayerNumber = 1
'    lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(lPageID, iLayerNumber)
'    MainMDIForm.ActiveForm.SpicerMarkup1.ActiveLayer = lLayerID
'    MainMDIForm.ActiveForm.SpicerMarkup1.ActiveTool = IN_TOOL_MOVERESIZE

    Call MainMDIForm.ActiveForm.SelectTool(IN_TOOL_MOVERESIZE)
    
    'Set FLAG to show that an Annotation or Modification may have occured.
    bolAnnotationAdded = True

End Sub


Private Sub cmdRedactBlack_Click()

    Call MainMDIForm.ActiveForm.subResetAnnotationButtons
    cmdRedactBlack.BackColor = vbGreen
    
    Call MainMDIForm.ActiveForm.SelectTool(IN_TOOL_HIGHLIGHTAREA)

    'Set the fillcolor to Black
    MainMDIForm.ActiveForm.SpicerMarkup1.FillColor(IN_TOOL_HIGHLIGHTAREA, IN_PROPSCOPE_ACTVTOOL, 0, 0) = 16777215

    'Set FLAG to show that an Annotation or Modification may have occured.
    bolAnnotationAdded = True

End Sub


Private Sub cmdSelect_Click()

    ' Get the page ID for current page of active document
    Dim lPageID As Long
    Dim lLayerID As Long
    Dim iLayerNumber As Integer
    
'    iLayerNumber = 1
'    lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(lPageID, iLayerNumber)
'    MainMDIForm.ActiveForm.SpicerMarkup1.ActiveLayer = lLayerID
        
    If cmdSelect.BackColor = vbGreen Then
        Call MainMDIForm.ActiveForm.subResetAnnotationButtons
        cmdSelect.BackColor = vbButtonFace
'        MainMDIForm.ActiveForm.SpicerMarkup1.ActiveTool = IN_TOOL_NOTOOL
        Call MainMDIForm.ActiveForm.SelectTool(IN_TOOL_NOTOOL)
    Else
        Call MainMDIForm.ActiveForm.subResetAnnotationButtons
        cmdSelect.BackColor = vbGreen
'        MainMDIForm.ActiveForm.SpicerMarkup1.ActiveTool = IN_TOOL_SELECT
        Call MainMDIForm.ActiveForm.SelectTool(IN_TOOL_SELECT)

    End If
    
    'Set FLAG to show that an Annotation or Modification may have occured.
    bolAnnotationAdded = True
    
End Sub

Private Sub cmdText_Click()
    
    Call MainMDIForm.ActiveForm.subResetAnnotationButtons
    cmdText.BackColor = vbGreen
    
'    ' Get the page ID for current page of active document
'    Dim lPageID As Long
'    Dim lLayerID As Long
'    Dim iLayerNumber As Integer
'
'    iLayerNumber = 1
'    lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(lPageID, iLayerNumber)
'    MainMDIForm.ActiveForm.SpicerMarkup1.ActiveLayer = lLayerID
'    MainMDIForm.ActiveForm.SpicerMarkup1.ActiveTool = IN_TOOL_TEXT

    Call MainMDIForm.ActiveForm.SelectTool(IN_TOOL_TEXT)

    'Set FLAG to show that an Annotation or Modification may have occured.
    bolAnnotationAdded = True

End Sub

Private Sub Form_Load()

    ' Get saved settings from the registry
    On Error Resume Next
    Me.Top = VBGetPrivateProfileString(RegAppname, "frmAnnotate.Top", RegFileName)
    Me.Left = VBGetPrivateProfileString(RegAppname, "frmAnnotate.Left", RegFileName)
    Me.width = VBGetPrivateProfileString(RegAppname, "frmAnnotate.Width", RegFileName)
    Me.Height = VBGetPrivateProfileString(RegAppname, "frmAnnotate.Height", RegFileName)
    On Error GoTo 0

    funcMakeTopMost Me, True

'    If gsecUserID = "jacob" Then
'        cmdLayerLoad.Visible = True
'        cmdLayerSave.Visible = True
'    Else
'        cmdLayerLoad.Visible = False
'        cmdLayerSave.Visible = False
'    End If
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next

    'Save Form settings to the registry
    If Me.Top > 0 And Me.Left > 0 Then
        result = WritePrivateProfileString(RegAppname, "frmAnnotate.Top", Me.Top, RegFileName)
        result = WritePrivateProfileString(RegAppname, "frmAnnotate.Left", Me.Left, RegFileName)
        result = WritePrivateProfileString(RegAppname, "frmAnnotate.Width", Me.width, RegFileName)
        result = WritePrivateProfileString(RegAppname, "frmAnnotate.Height", Me.Height, RegFileName)
    End If

    'De-Select any active Tool
    Call MainMDIForm.ActiveForm.SelectTool(IN_TOOL_NOTOOL)

    'SAFE way of saying: Set Me = Nothing
    funcWriteToDebugLog Me.name, "BEGIN SAFE UNLOAD of frmAnnotate"
    
    Dim Form As Form
    For Each Form In Forms
            If Form Is Me Then
                    funcWriteToDebugLog Me.name, "frmAnnotate = Nothing"
                    Set Form = Nothing
                    funcWriteToDebugLog Me.name, "Exit For"
                    Exit For
            End If
    Next Form


End Sub
