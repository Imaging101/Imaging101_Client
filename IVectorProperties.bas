Attribute VB_Name = "mod_IVectorProperties"
' File:      IVectorProperties.bas
' Created:   1998July30 by Rob Young
' Purpose:   To test the Spicer Markup Control's IVectorProperties interface.
' Revisions: RobY Jan4/99  - Added functionality to open more than 1 file. ChildForm1 is no longer used access objects because it has been set to ChildForm() which is an array of ChildForm1's forms.
'                          - Anywhere ChildForm1 was used has been changed to ChildForm(iIndex) or MainMDIForm.ActiveForm.
'   $log$
'   $nokeywords$
'
' Copyright (C) 1998 SPICER Corporation, All rights reserved.
Option Explicit
Dim MenuName As Menu

Public Sub IVectorProperties_BindToDocumentControl()
   ' RobY Aug12/98
   Dim VectorProperties As IVectorProperties
   
   ' Set object variable for IVectorProperties interface to Markup ctrl object
   Set VectorProperties = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   'Bind the Spicer Markup Control to the Spicer Document Control
   VectorProperties.BindToDocumentControl MainMDIForm.ActiveForm.SpicerDoc1.object
   
   ' De-initialize the object variable
   Set VectorProperties = Nothing
End Sub

Public Sub IVectorProperties_BindToViewControl()
   ' RobY Aug12/98
   Dim VectorProperties As IVectorProperties
   
   ' Set object variable for IVectorProperties interface to Markup ctrl object
   Set VectorProperties = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   'Bind the Spicer Markup Control to the Spicer View Control
   VectorProperties.BindToViewControl MainMDIForm.ActiveForm.SpicerView1.object
   
   ' De-initialize the object variable
   Set VectorProperties = Nothing
End Sub

Public Sub IVectorProperties_IconCentered()
   ' RobY Oct28/98
   ' Modified by RobY Nov25/98 - Added code so user can specify scope, and toggle between true and false
   ' Modified by RobY Dec1/98 - Removed tool parameter
   ' Modified by RobY Dec3/98 - Renamed from Centered to IconCentered
   Dim VectorProperties As IVectorProperties
   Dim Scope As PROPERTY_SCOPE
   Dim lLayerID As Long
   Dim iLayerNum As Integer
   Dim lObjectID As Long
   
   ' Set object variable for IVectorProperties interface to Markup ctrl object
   Set VectorProperties = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Initialize variables
   lLayerID = 0
   lObjectID = 0
   
   ' Prompt for scope to use
   Scope = InputBox("Please enter the scope that you want to use. Check help file for values.", "Scope")
   If ((Scope >= 128) And (Scope < 256)) Or (Scope >= 384) Then
      ' Prompt for layer number
      iLayerNum = InputBox("Please enter the layer number of where the hotspot is to center.", "Layer Number")
      ' Prompt for object id
      lObjectID = InputBox("Please enter the object id of the hotspot.", "Object ID")
      ' Get the layer id of the specified layer
      lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   End If
   ' Set the value for the IconCentered property
   VectorProperties.IconCentered(Scope, lLayerID, lObjectID) = True
   MsgBox "The specified hotspot has been centered.", vbInformation, "IconCentered"
   
   ' De-initialize the object variable
   Set VectorProperties = Nothing
End Sub

Public Sub IVectorProperties_ChangeTextDialog()
   ' RobY Aug12/98
   ' Modified by RobY Dec1/98 - Setup so user can specify values
   Dim VectorProperties As IVectorProperties
   Dim Scope As PROPERTY_SCOPE
   Dim lLayerID As Long
   Dim iLayerNum As Integer
   Dim lObjectID As Long
   
   ' Set object variable for IVectorProperties interface to Markup ctrl object
   Set VectorProperties = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Initialize variables
   lLayerID = 0
   lObjectID = 0
   
   ' Prompt for scope to use
   Scope = InputBox("Please enter the scope that you want to use. Check help file for values.", "Scope")
   If ((Scope >= 128) And (Scope < 256)) Or (Scope >= 384) Then
      iLayerNum = InputBox("Please enter the layer number of where the the object to change.", "Layer Number")
      ' Prompt for object id
      lObjectID = InputBox("Please enter the object id of the object to change.", "Object ID")
      ' Get the layer id of the specified layer
      lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   End If
   '  Display the Change Text Dialog
   VectorProperties.ChangeTextDialog Scope, lLayerID, lObjectID
   
   ' De-initialize the object variable
   Set VectorProperties = Nothing
End Sub

Public Sub IVectorProperties_DimensionStyle()
   ' RobY Nov2/98
   ' Modified by RobY Dec1/98 - Setup so user can specify values
   ' Modified by RobY Dec3/98 - Removed tool parameter from command
   Dim VectorProperties As IVectorProperties
   Dim lLayerID As Long
   Dim iDimStyle As DIMENSION_STYLE
   Dim strDimStyle As String
   Dim Scope As PROPERTY_SCOPE
   Dim lObjectID As Long
   Dim iLayerNum As Integer
   
   ' Set object variable for IVectorProperties interface to Markup ctrl object
   Set VectorProperties = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Initialize variables
   lLayerID = 0
   lObjectID = 0

   ' Prompt for dimension style to use
   iDimStyle = InputBox("Please enter a dimension style value:" + Chr(13) + "External Linear = 0" + _
                        Chr(13) + "Internal Linear = 1" + Chr(13) + "Leader Text = 2", "Dimension Style")
   ' Prompt for scope to use
   Scope = InputBox("Please enter the scope that you want to use. Check help file for values.", "Scope")
   If ((Scope >= 128) And (Scope < 256)) Or (Scope >= 384) Then
      iLayerNum = InputBox("Please enter the layer number of where the the object to change.", "Layer Number")
      ' Prompt for object id
      lObjectID = InputBox("Please enter the object id of the object to change.", "Object ID")
      ' Get the layer id of the specified layer
      lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   End If
   ' Set the dimension style
   VectorProperties.DimensionStyle(Scope, lLayerID, lObjectID) = iDimStyle
   ' Convert value of dimension style to string
   Select Case iDimStyle
      Case 0
         strDimStyle = "External Linear"
      Case 1
         strDimStyle = "Internal Linear"
      Case 2
         strDimStyle = "Leader Text"
   End Select
   MsgBox "The dimension style has been set to " + strDimStyle + ".", vbInformation, "Dimension Style"

   ' De-initialize the object variable
   Set VectorProperties = Nothing
End Sub

Public Sub IVectorProperties_DualArrowHeads()
   ' RobY Nov2/98
   ' Modified by RobY Dec1/98 - Setup so user can specify values
   Dim VectorProperties As IVectorProperties
   Dim lLayerID As Long
   Dim bValue As Boolean
   Dim bOrigValue As Boolean
   Dim strValue As String
   Dim Scope As PROPERTY_SCOPE
   Dim lObjectID As Long
   Dim iLayerNum As Integer
   Dim iToolType As TOOL_TYPE
   
   ' Set object variable for IVectorProperties interface to Markup ctrl object
   Set VectorProperties = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Initialize variables
   lLayerID = 0
   lObjectID = 0
   
   ' Prompt for scope to use
   Scope = InputBox("Please enter the scope that you want to use. Check help file for values.", "Scope")
   If ((Scope >= 128) And (Scope < 256)) Or (Scope >= 384) Then
      iLayerNum = InputBox("Please enter the layer number of where the the object to change.", "Layer Number")
      ' Prompt for object id
      lObjectID = InputBox("Please enter the object id of the object to change.", "Object ID")
      ' Get the layer id of the specified layer
      lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   End If
   ' Prompt for tool
   iToolType = InputBox("Please enter the tool type. Setting is ignored depending on the scope.", "Tool Type")
   ' Get the original setting
   bOrigValue = VectorProperties.DualArrowHeads(iToolType, Scope, lLayerID, lObjectID)
   If bOrigValue = False Then
      ' Set dual arrow heads to TRUE
      VectorProperties.DualArrowHeads(iToolType, Scope, lLayerID, lObjectID) = True
   Else
      ' Set dual arrow heads to FALSE
      VectorProperties.DualArrowHeads(iToolType, Scope, lLayerID, lObjectID) = False
   End If
   bValue = VectorProperties.DualArrowHeads(iToolType, Scope, lLayerID, lObjectID)
   If bValue = bOrigValue Then
      MsgBox "Set failed to change value." + Chr(13) + "Original Value:" + Str(bOrigValue) + Chr(13) + _
               "New Value:" + Str(bValue), vbCritical, "Failure"
   Else
      Select Case bValue
         Case False
            strValue = "FALSE"
         Case Else
            strValue = "TRUE"
      End Select
      MsgBox "Dual arrow heads has been set to " + strValue, vbInformation, "Dual Arrow Heads"
   End If
   
   ' De-initialize the object variable
   Set VectorProperties = Nothing
End Sub

Public Sub IVectorProperties_EditPreferencesDialog()
   ' RobY Aug12/98
   ' Modified by RobY Dec1/98 - Setup so user can specify values
   ' Modified by RobY Dec 16/98 - Rewrote code to include availability and to allow user to specify scope(new parameter)
   Dim VectorProperties As IVectorProperties
   Dim Scope As PROPERTY_SCOPE
   Dim Available As COMMAND_AVAILABILITY
   
   ' Set object variable for IVectorProperties interface to Markup ctrl object
   Set VectorProperties = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Prompt for scope to use
   Scope = InputBox("Please enter the scope that you want to use. Check help file for values.", "Scope")
   ' Get the value of availability
   Available = VectorProperties.EditPreferencesDialogAvailability(Scope)
   ' Do the case statement that corresponds to the availability returned
   Select Case Available
      Case IN_UI_ENABLED
         MsgBox "Command is ENABLED. Click OK to display dialog.", vbInformation, "Availability"
         'Display the dialog
         VectorProperties.EditPreferencesDialog Scope
      Case IN_UI_GREYED
         MsgBox "Command is GREYED.", vbInformation, "Availability"
      Case IN_UI_REMOVED
         MsgBox "Command is REMOVED.", vbInformation, "Availability"
      Case IN_UI_CHECKED
         MsgBox "Command is CHECKED.", vbInformation, "Availability"
      Case IN_UI_ENABLED_CHECKED
         MsgBox "Command is ENABLED and CHECKED. Click OK to display dialog.", vbInformation, "Availability"
         'Display the dialog
         VectorProperties.EditPreferencesDialog Scope
      Case Else
         MsgBox "Cannot open dialog. Check to make that the scope you specified is valid.", vbExclamation, "Availability"
   End Select
   
   ' De-initialize the object variable
   Set VectorProperties = Nothing
End Sub

Public Sub IVectorProperties_FillColor()
   ' RobY Nov6/98
   ' Modified by RobY Dec1/98 - Setup so user can specify values
   Dim VectorProperties As IVectorProperties
   Dim lLayerID As Long
   Dim lFillColor As Long
   Dim strFillColor As String
   Dim Scope As PROPERTY_SCOPE
   Dim lObjectID As Long
   Dim iLayerNum As Integer
   Dim iToolType As TOOL_TYPE
   
   ' Set object variable for IVectorProperties interface to Markup ctrl object
   Set VectorProperties = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Initialize variables
   lLayerID = 0
   lObjectID = 0

   ' Prompt for scope to use
   Scope = InputBox("Please enter the scope that you want to use. Check help file for values.", "Scope")
   If ((Scope >= 128) And (Scope < 256)) Or (Scope >= 384) Then
      iLayerNum = InputBox("Please enter the layer number of where the the object to change.", "Layer Number")
      ' Prompt for object id
      lObjectID = InputBox("Please enter the object id of the object to change.", "Object ID")
      ' Get the layer id of the specified layer
      lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   End If
   ' Prompt for fill color to use
   lFillColor = InputBox("Please enter a fill color value(24-bit):", "Fill Color")
   ' Prompt for tool
   iToolType = InputBox("Please enter the tool type. Setting is ignored depending on the scope.", "Tool Type")
   ' Set the fill style
   VectorProperties.FillColor(iToolType, Scope, lLayerID, lObjectID) = lFillColor
   MsgBox "The fill color has been set to" + Str(lFillColor), vbInformation, "Fill Color"

   ' De-initialize the object variable
   Set VectorProperties = Nothing
End Sub

Public Sub IVectorProperties_FillStyle()
   ' RobY Nov3/98
   ' Modified by RobY Dec1/98 - Setup so user can specify values
   Dim VectorProperties As IVectorProperties
   Dim lLayerID As Long
   Dim iFillStyle As FILL_STYLE
   Dim strFillStyle As String
   Dim Scope As PROPERTY_SCOPE
   Dim lObjectID As Long
   Dim iLayerNum As Integer
   Dim iToolType As TOOL_TYPE
   
   ' Set object variable for IVectorProperties interface to Markup ctrl object
   Set VectorProperties = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Initialize variables
   lLayerID = 0
   lObjectID = 0

   ' Prompt for scope to use
   Scope = InputBox("Please enter the scope that you want to use. Check help file for values.", "Scope")
   If ((Scope >= 128) And (Scope < 256)) Or (Scope >= 384) Then
      iLayerNum = InputBox("Please enter the layer number of where the the object to change.", "Layer Number")
      ' Prompt for object id
      lObjectID = InputBox("Please enter the object id of the object to change.", "Object ID")
      ' Get the layer id of the specified layer
      lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   End If
   ' Prompt for fill style to use
   iFillStyle = InputBox("Please enter a fill style value:" + Chr(13) + "TRANSPARENT = 1" + _
                        Chr(13) + "OPAQUE = 2" + Chr(13) + "ERASE = 3" + Chr(13) + "TRANSLUCENT = 4" + _
                        Chr(13) + "HATCH = 5", "Fill Style")
   ' Prompt for tool
   iToolType = InputBox("Please enter the tool type. Setting is ignored depending on the scope.", "Tool Type")
   ' Set the fill style
   VectorProperties.FillStyle(iToolType, Scope, lLayerID, lObjectID) = iFillStyle
   ' Convert value of fill style to string
   Select Case iFillStyle
      Case 1
         strFillStyle = "TRANSPARENT"
      Case 2
         strFillStyle = "OPAQUE"
      Case 3
         strFillStyle = "ERASE"
      Case 4
         strFillStyle = "TRANSLUCENT"
      Case 5
         strFillStyle = "HATCH"
   End Select
   MsgBox "The fill style has been set to " + strFillStyle + ".", vbInformation, "Fill Style"

   ' De-initialize the object variable
   Set VectorProperties = Nothing
End Sub

Public Sub IVectorProperties_FillStyleAvailability(Index As Integer, MenuClicked As Boolean)
   ' RobY Dec16/98
   Dim VectorProperties As IVectorProperties
   Dim iCount As Integer
   Dim iFillStyle As FILL_STYLE
   Dim Available(4) As COMMAND_AVAILABILITY
   Dim MenuName As Menu

   ' Set object variable for IVectorProperties interface to Markup ctrl object
   Set VectorProperties = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Set menu status in FillStyleAvailability menu
   For iCount = 0 To 4
      ' Set the menu name to change
      Set MenuName = MainMDIForm.ActiveForm.mnu_Markup_IVectorProperties_FillStyleAvailability_Array(iCount)
      ' Get availability for each fill style
      Available(iCount) = VectorProperties.FillStyleAvailability(iCount + 1)
      ' Set the menu status through Availability procedure in mod_Global.bas module
      Availability Available(iCount), MenuName
   Next iCount
   
   ' Change fill style for active tool if clicked
   If MenuClicked = True Then
      If Index = 0 Then
         VectorProperties.FillStyle(0, IN_PROPSCOPE_ACTVTOOL, 0, 0) = IN_FILL_TRANSPARENT
      ElseIf Index = 1 Then
         VectorProperties.FillStyle(0, IN_PROPSCOPE_ACTVTOOL, 0, 0) = IN_FILL_OPAQUE
      ElseIf Index = 2 Then
         VectorProperties.FillStyle(0, IN_PROPSCOPE_ACTVTOOL, 0, 0) = IN_FILL_ERASE
      ElseIf Index = 3 Then
         VectorProperties.FillStyle(0, IN_PROPSCOPE_ACTVTOOL, 0, 0) = IN_FILL_TRANSLUCENT
      ElseIf Index = 4 Then
         VectorProperties.FillStyle(0, IN_PROPSCOPE_ACTVTOOL, 0, 0) = IN_FILL_HATCH
      End If
   End If
   
   ' De-initialize the object variable
   Set VectorProperties = Nothing
   Set MenuName = Nothing
End Sub

Public Sub IVectorProperties_Font()
   ' RobY Nov6/98
   ' Modified by RobY Dec1/98 - Setup so user can specify values
   Dim VectorProperties As IVectorProperties
   Dim lLayerID As Long
   Dim strGetFont As String
   Dim strSetFont As String
   Dim strOrigFont As String
   Dim Scope As PROPERTY_SCOPE
   Dim lObjectID As Long
   Dim iLayerNum As Integer
   Dim iToolType As TOOL_TYPE
   
   ' Set object variable for IVectorProperties interface to Markup ctrl object
   Set VectorProperties = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Initialize variables
   lLayerID = 0
   lObjectID = 0

   ' Prompt for scope to use
   Scope = InputBox("Please enter the scope that you want to use. Check help file for values.", "Scope")
   If ((Scope >= 128) And (Scope < 256)) Or (Scope >= 384) Then
      iLayerNum = InputBox("Please enter the layer number of where the the object to change.", "Layer Number")
      ' Prompt for object id
      lObjectID = InputBox("Please enter the object id of the object to change.", "Object ID")
      ' Get the layer id of the specified layer
      lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   End If
   ' Prompt for font to use
   strSetFont = InputBox("Please enter a font to use.", "Font")
   ' Prompt for tool
   iToolType = InputBox("Please enter the tool type. Setting is ignored depending on the scope.", "Tool Type")
   ' Get the original font
   strOrigFont = VectorProperties.Font(iToolType, Scope, lLayerID, lObjectID)
   ' Set the font
   VectorProperties.Font(iToolType, Scope, lLayerID, lObjectID) = strSetFont
   ' Get font that was set
   strGetFont = VectorProperties.Font(iToolType, Scope, lLayerID, lObjectID)
   If strGetFont = strSetFont Then
      MsgBox "Original Font: " + strOrigFont + Chr(13) + "Change To: " + strSetFont + _
            "Returned: " + strGetFont, vbInformation, "Success"
   Else
      MsgBox "Original Font: " + strOrigFont + Chr(13) + "Returned: " + strGetFont, vbCritical, "Failed"
   End If
   
   ' De-initialize the object variable
   Set VectorProperties = Nothing
End Sub

Public Sub IVectorProperties_GetArrowHeadLength()
   ' RobY Nov6/98
   ' Modified by RobY Dec1/98 - Setup so user can specify values
   Dim VectorProperties As IVectorProperties
   Dim iUnitType As UNIT_TYPE
   Dim dLength As Double
   Dim Scope As PROPERTY_SCOPE
   Dim lObjectID As Long
   Dim iLayerNum As Integer
   Dim iToolType As TOOL_TYPE
   Dim lLayerID As Long
   
   ' Set object variable for IVectorProperties interface to Markup ctrl object
   Set VectorProperties = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Initialize variables
   lLayerID = 0
   lObjectID = 0

   ' Prompt for scope to use
   Scope = InputBox("Please enter the scope that you want to use. Check help file for values.", "Scope")
   If ((Scope >= 128) And (Scope < 256)) Or (Scope >= 384) Then
      iLayerNum = InputBox("Please enter the layer number of where the the object to change.", "Layer Number")
      ' Prompt for object id
      lObjectID = InputBox("Please enter the object id of the object to change.", "Object ID")
      ' Get the layer id of the specified layer
      lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   End If
   ' Prompt for tool
   iToolType = InputBox("Please enter the tool type. Setting is ignored depending on the scope.", "Tool Type")
   ' Get the arrow head length
   VectorProperties.GetArrowHeadLength iToolType, Scope, lLayerID, lObjectID, dLength, iUnitType
   MsgBox "Length: " + Str(dLength) + Chr(13) + "Unit Type: " + Str(iUnitType), vbInformation, "GetArrowHeadLength"
   
   ' De-initialize the object variable
   Set VectorProperties = Nothing
End Sub

Public Sub IVectorProperties_GetLineThickness()
   ' RobY Nov6/98
   ' Modified by RobY Dec1/98 - Setup so user can specify values
   Dim VectorProperties As IVectorProperties
   Dim iUnitType As UNIT_TYPE
   Dim dThickness As Double
   Dim Scope As PROPERTY_SCOPE
   Dim lObjectID As Long
   Dim iLayerNum As Integer
   Dim iToolType As TOOL_TYPE
   Dim lLayerID As Long
   
   ' Set object variable for IVectorProperties interface to Markup ctrl object
   Set VectorProperties = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Initialize variables
   lLayerID = 0
   lObjectID = 0

   ' Prompt for scope to use
   Scope = InputBox("Please enter the scope that you want to use. Check help file for values.", "Scope")
   If ((Scope >= 128) And (Scope < 256)) Or (Scope >= 384) Then
      iLayerNum = InputBox("Please enter the layer number of where the the object to change.", "Layer Number")
      ' Prompt for object id
      lObjectID = InputBox("Please enter the object id of the object to change.", "Object ID")
      ' Get the layer id of the specified layer
      lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   End If
   ' Prompt for tool
   iToolType = InputBox("Please enter the tool type. Setting is ignored depending on the scope.", "Tool Type")
   ' Get the line thickness
   VectorProperties.GetLineThickness iToolType, Scope, lLayerID, lObjectID, dThickness, iUnitType
   MsgBox "Line Thickness: " + Str(dThickness) + Chr(13) + "Unit Type: " + Str(iUnitType), vbInformation, "GetLineThickness"
   
   ' De-initialize the object variable
   Set VectorProperties = Nothing
End Sub

Public Sub IVectorProperties_GetTextHeight()
   ' RobY Nov6/98
   ' Modified by RobY Dec1/98 - Setup so user can specify values
   Dim VectorProperties As IVectorProperties
   Dim iUnitType As UNIT_TYPE
   Dim dHeight As Double
   Dim Scope As PROPERTY_SCOPE
   Dim lObjectID As Long
   Dim iLayerNum As Integer
   Dim iToolType As TOOL_TYPE
   Dim lLayerID As Long
   
   ' Set object variable for IVectorProperties interface to Markup ctrl object
   Set VectorProperties = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Initialize variables
   lLayerID = 0
   lObjectID = 0

   ' Prompt for scope to use
   Scope = InputBox("Please enter the scope that you want to use. Check help file for values.", "Scope")
   If ((Scope >= 128) And (Scope < 256)) Or (Scope >= 384) Then
      iLayerNum = InputBox("Please enter the layer number of where the the object to change.", "Layer Number")
      ' Prompt for object id
      lObjectID = InputBox("Please enter the object id of the object to change.", "Object ID")
      ' Get the layer id of the specified layer
      lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   End If
   ' Prompt for tool
   iToolType = InputBox("Please enter the tool type. Setting is ignored depending on the scope.", "Tool Type")
   ' Get the text height
   VectorProperties.GetTextHeight iToolType, Scope, lLayerID, lObjectID, dHeight, iUnitType
   MsgBox "Text Height: " + Str(dHeight) + Chr(13) + "Unit Type: " + Str(iUnitType), vbInformation, "GetTextHeight"
   
   ' De-initialize the object variable
   Set VectorProperties = Nothing
End Sub

Public Sub IVectorProperties_GetTextWidth()
   ' RobY Nov6/98
   ' Modified by RobY Dec1/98 - Setup so user can specify values
   Dim VectorProperties As IVectorProperties
   Dim iUnitType As UNIT_TYPE
   Dim dWidth As Double
   Dim Scope As PROPERTY_SCOPE
   Dim lObjectID As Long
   Dim iLayerNum As Integer
   Dim iToolType As TOOL_TYPE
   Dim lLayerID As Long
   
   ' Set object variable for IVectorProperties interface to Markup ctrl object
   Set VectorProperties = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Initialize variables
   lLayerID = 0
   lObjectID = 0

   ' Prompt for scope to use
   Scope = InputBox("Please enter the scope that you want to use. Check help file for values.", "Scope")
   If ((Scope >= 128) And (Scope < 256)) Or (Scope >= 384) Then
      iLayerNum = InputBox("Please enter the layer number of where the the object to change.", "Layer Number")
      ' Prompt for object id
      lObjectID = InputBox("Please enter the object id of the object to change.", "Object ID")
      ' Get the layer id of the specified layer
      lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   End If
   ' Prompt for tool
   iToolType = InputBox("Please enter the tool type. Setting is ignored depending on the scope.", "Tool Type")
   ' Get the text width
   VectorProperties.GetTextWidth iToolType, Scope, lLayerID, lObjectID, dWidth, iUnitType
   MsgBox "Text Tool Width: " + Str(dWidth) + Chr(13) + "Unit Type: " + Str(iUnitType), vbInformation, "GetTextWidth"
   
   ' De-initialize the object variable
   Set VectorProperties = Nothing
End Sub

Public Sub IVectorProperties_Iconized()
   ' RobY Nov9/98
   ' Modified by RobY Dec1/98 - Setup so user can specify values
   ' Modified by RobY Dec3/98 - Removed tool parameter from command
   Dim VectorProperties As IVectorProperties
   Dim iUnitType As UNIT_TYPE
   Dim bCurIconized As Boolean
   Dim bNewIconized As Boolean
   Dim Scope As PROPERTY_SCOPE
   Dim lObjectID As Long
   Dim iLayerNum As Integer
   Dim lLayerID As Long
   
   ' Set object variable for IVectorProperties interface to Markup ctrl object
   Set VectorProperties = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Initialize variables
   lLayerID = 0
   lObjectID = 0

   ' Prompt for scope to use
   Scope = InputBox("Please enter the scope that you want to use. Check help file for values.", "Scope")
   If ((Scope >= 128) And (Scope < 256)) Or (Scope >= 384) Then
      iLayerNum = InputBox("Please enter the layer number of where the the object to change.", "Layer Number")
      ' Prompt for object id
      lObjectID = InputBox("Please enter the object id of the object to change.", "Object ID")
      ' Get the layer id of the specified layer
      lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   End If
   ' Is the annotation tool displayed as an icon (TRUE or FALSE)
   bCurIconized = VectorProperties.Iconized(Scope, lLayerID, lObjectID)
   ' Set the iconized property to the opposite of the current setting
   VectorProperties.Iconized(Scope, lLayerID, lObjectID) = Not bCurIconized
   ' Get the new setting of the iconized property
   bNewIconized = VectorProperties.Iconized(Scope, lLayerID, lObjectID)
   If bCurIconized = bNewIconized Then
      MsgBox "Failed to set to new value." + Chr(13) + "Original Setting: " + Str(bCurIconized) + _
            Chr(13) + "New Setting: " + Str(bNewIconized), vbCritical, "Failed"
   Else
      MsgBox "Original Setting: " + Str(bCurIconized) + Chr(13) + "New Setting: " + Str(bNewIconized), vbInformation, "Success"
   End If
   
   ' De-initialize the object variable
   Set VectorProperties = Nothing
End Sub

Public Sub IVectorProperties_IconNumber()
   ' RobY Nov9/98
   ' Modified by RobY Dec1/98 - Setup so user can specify values
   ' Modified by RobY Dec3/98 - Removed tool parameter from command
   Dim VectorProperties As IVectorProperties
   Dim iUnitType As UNIT_TYPE
   Dim iIconNum As Integer
   Dim Scope As PROPERTY_SCOPE
   Dim lObjectID As Long
   Dim iLayerNum As Integer
   Dim lLayerID As Long
   
   ' Set object variable for IVectorProperties interface to Markup ctrl object
   Set VectorProperties = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Initialize variables
   lLayerID = 0
   lObjectID = 0

   ' Prompt for scope to use
   Scope = InputBox("Please enter the scope that you want to use. Check help file for values.", "Scope")
   If ((Scope >= 128) And (Scope < 256)) Or (Scope >= 384) Then
      iLayerNum = InputBox("Please enter the layer number of where the the object to change.", "Layer Number")
      ' Prompt for object id
      lObjectID = InputBox("Please enter the object id of the object to change.", "Object ID")
      ' Get the layer id of the specified layer
      lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   End If
   iIconNum = VectorProperties.IconNumber(Scope, lLayerID, lObjectID)
      MsgBox "Icon Number: " + Str(iIconNum), vbInformation, "Icon Number"
   
   ' De-initialize the object variable
   Set VectorProperties = Nothing
End Sub

Public Sub IVectorProperties_LineCapStyle()
   ' RobY Nov9/98
   ' Modified by RobY Dec1/98 - Setup so user can specify values
   Dim VectorProperties As IVectorProperties
   Dim iCurCapStyle As LINE_CAP_STYLE
   Dim strCapStyle As String
   Dim iNewCapStyle As LINE_CAP_STYLE
   Dim iResponse As Integer
   Dim Scope As PROPERTY_SCOPE
   Dim lObjectID As Long
   Dim iLayerNum As Integer
   Dim lLayerID As Long
   
   ' Set object variable for IVectorProperties interface to Markup ctrl object
   Set VectorProperties = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Initialize variables
   lLayerID = 0
   lObjectID = 0

   ' Prompt for scope to use
   Scope = InputBox("Please enter the scope that you want to use. Check help file for values.", "Scope")
   If ((Scope >= 128) And (Scope < 256)) Or (Scope >= 384) Then
      iLayerNum = InputBox("Please enter the layer number of where the the object to change.", "Layer Number")
      ' Prompt for object id
      lObjectID = InputBox("Please enter the object id of the object to change.", "Object ID")
      ' Get the layer id of the specified layer
      lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   End If
   ' Get the cap style for the line tool
   iCurCapStyle = VectorProperties.LineCapStyle(IN_TOOL_LINE, Scope, lLayerID, lObjectID)
   Select Case iCurCapStyle
      Case IN_CAP_ROUND
         strCapStyle = "ROUND"
      Case IN_CAP_SQUARE
         strCapStyle = "SQUARE"
   End Select
   iResponse = MsgBox("The line cap style is currently set to " + strCapStyle + "." + Chr(13) + _
                        "Do you want to change the cap style?", vbYesNo, "Line Cap Style")
   If iResponse = vbYes Then
      If iCurCapStyle = IN_CAP_ROUND Then
         VectorProperties.LineCapStyle(IN_TOOL_LINE, Scope, lLayerID, lObjectID) = IN_CAP_SQUARE
      Else
         VectorProperties.LineCapStyle(IN_TOOL_LINE, Scope, lLayerID, lObjectID) = IN_CAP_ROUND
      End If
      ' Get the new cap style
      iNewCapStyle = VectorProperties.LineCapStyle(IN_TOOL_LINE, Scope, lLayerID, lObjectID)
      If iNewCapStyle = iCurCapStyle Then
         MsgBox "Failed to change cap style to new setting!", vbCritical, "Failed"
      Else
         MsgBox "Line cap style has been changed to new setting.", vbInformation, "Success"
      End If
   Else
      MsgBox "Line cap style was not changed.", vbInformation, "Line Cap Style"
   End If
   
   ' De-initialize the object variable
   Set VectorProperties = Nothing
End Sub

Public Sub IVectorProperties_LineColor()
   ' RobY Oct30/98
   ' Modified by RobY Dec1/98 - Setup so user can specify values
   Dim VectorProperties As IVectorProperties
   Dim Scope As PROPERTY_SCOPE
   Dim lObjectID As Long
   Dim iLayerNum As Integer
   Dim iToolType As TOOL_TYPE
   Dim lLayerID As Long
   
   ' Set object variable for IVectorProperties interface to Markup ctrl object
   Set VectorProperties = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Initialize variables
   lLayerID = 0
   lObjectID = 0

   ' Prompt for scope to use
   Scope = InputBox("Please enter the scope that you want to use. Check help file for values.", "Scope")
   If ((Scope >= 128) And (Scope < 256)) Or (Scope >= 384) Then
      iLayerNum = InputBox("Please enter the layer number of where the the object to change.", "Layer Number")
      ' Prompt for object id
      lObjectID = InputBox("Please enter the object id of the object to change.", "Object ID")
      ' Get the layer id of the specified layer
      lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   End If
   ' Prompt for tool
   iToolType = InputBox("Please enter the tool type. Setting is ignored depending on the scope.", "Tool Type")
   ' Set the line color to cyan(65535)
   VectorProperties.LineColor(iToolType, Scope, lLayerID, lObjectID) = 65535
   MsgBox "Set line colour to cyan.", vbInformation, "Line Color"
   
   ' De-initialize the object variable
   Set VectorProperties = Nothing
End Sub

Public Sub IVectorProperties_LineFillStyle()
   ' RobY Nov18/98
   ' Modified by RobY Dec1/98 - Setup so user can specify values
   Dim VectorProperties As IVectorProperties
   Dim iOrigFillStyle As LINE_FILL_OPERATION
   Dim iNewFillStyle As LINE_FILL_OPERATION
   Dim strFillStyle(2) As String
   Dim iResponse As Integer
   Dim Scope As PROPERTY_SCOPE
   Dim lObjectID As Long
   Dim iLayerNum As Integer
   Dim iToolType As TOOL_TYPE
   Dim lLayerID As Long
   
   ' Set object variable for IVectorProperties interface to Markup ctrl object
   Set VectorProperties = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Initialize variables
   lLayerID = 0
   lObjectID = 0

   ' Prompt for scope to use
   Scope = InputBox("Please enter the scope that you want to use. Check help file for values.", "Scope")
   If ((Scope >= 128) And (Scope < 256)) Or (Scope >= 384) Then
      iLayerNum = InputBox("Please enter the layer number of where the the object to change.", "Layer Number")
      ' Prompt for object id
      lObjectID = InputBox("Please enter the object id of the object to change.", "Object ID")
      ' Get the layer id of the specified layer
      lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   End If
   ' Prompt for tool
   iToolType = InputBox("Please enter the tool type. Setting is ignored depending on the scope.", "Tool Type")
   ' Get the fill style
   iOrigFillStyle = VectorProperties.LineFillStyle(IN_TOOL_LINE, IN_PROPSCOPE_SESSION, 0, 0)
   ' Convert value of fill style to string
   Select Case iOrigFillStyle
      Case 1
         strFillStyle(0) = "OPAQUE"
         strFillStyle(1) = "TRANSPARENT"
      Case 2
         strFillStyle(0) = "TRANSPARENT"
         strFillStyle(1) = "OPAQUE"
   End Select
   ' Prompt user
   iResponse = MsgBox("The fill style is currently set to " + strFillStyle(0) + "." + Chr(13) + _
                  "Do you want to change it to " + strFillStyle(1) + "?", vbYesNo + vbInformation, "LineFillStyle")
   If iResponse = vbYes Then
      If iOrigFillStyle = IN_ROP_OPAQUE Then
            VectorProperties.LineFillStyle(iToolType, Scope, lLayerID, lObjectID) = IN_ROP_TRANSLUCENT
      Else
            VectorProperties.LineFillStyle(iToolType, Scope, lLayerID, lObjectID) = IN_ROP_OPAQUE
      End If
      iNewFillStyle = VectorProperties.LineFillStyle(iToolType, Scope, lLayerID, lObjectID)
      If iOrigFillStyle = iNewFillStyle Then
         MsgBox "Failed to change to new setting!", vbCritical, "Failed"
      Else
         MsgBox "The fill style has been changed to new setting.", vbInformation, "Success"
      End If
   Else
      MsgBox "The fill style was not changed.", vbInformation, "LineFillStyle"
   End If
      
   ' De-initialize the object variable
   Set VectorProperties = Nothing
End Sub

Public Sub IVectorProperties_LineJoinStyle()
   ' RobY Nov9/98
   ' Modified by RobY Dec1/98 - Setup so user can specify values
   Dim VectorProperties As IVectorProperties
   Dim iCurJoinStyle As LINE_JOIN_STYLE
   Dim strJoinStyle As String
   Dim iNewJoinStyle As LINE_JOIN_STYLE
   Dim iResponse As Integer
   Dim Scope As PROPERTY_SCOPE
   Dim lObjectID As Long
   Dim iLayerNum As Integer
   Dim iToolType As TOOL_TYPE
   Dim lLayerID As Long
   
   ' Set object variable for IVectorProperties interface to Markup ctrl object
   Set VectorProperties = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Initialize variables
   lLayerID = 0
   lObjectID = 0

   ' Prompt for scope to use
   Scope = InputBox("Please enter the scope that you want to use. Check help file for values.", "Scope")
   If ((Scope >= 128) And (Scope < 256)) Or (Scope >= 384) Then
      iLayerNum = InputBox("Please enter the layer number of where the the object to change.", "Layer Number")
      ' Prompt for object id
      lObjectID = InputBox("Please enter the object id of the object to change.", "Object ID")
      ' Get the layer id of the specified layer
      lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   End If
   ' Prompt for tool
   iToolType = InputBox("Please enter the tool type. Setting is ignored depending on the scope.", "Tool Type")
   'Get the join style
   iCurJoinStyle = VectorProperties.LineJoinStyle(iToolType, Scope, lLayerID, lObjectID)
   Select Case iCurJoinStyle
      Case IN_CAP_ROUND
         strJoinStyle = "ROUND"
      Case IN_CAP_SQUARE
         strJoinStyle = "SQUARE"
   End Select
   iResponse = MsgBox("The line join style is currently set to " + strJoinStyle + "." + Chr(13) + _
                        "Do you want to change the join style?", vbYesNo, "Line Join Style")
   If iResponse = vbYes Then
      If iCurJoinStyle = IN_CAP_ROUND Then
         VectorProperties.LineJoinStyle(iToolType, Scope, lLayerID, lObjectID) = IN_JOIN_MITER
      Else
         VectorProperties.LineJoinStyle(iToolType, Scope, lLayerID, lObjectID) = IN_CAP_ROUND
      End If
     ' Get the new join style
      iNewJoinStyle = VectorProperties.LineJoinStyle(iToolType, Scope, lLayerID, lObjectID)
      If iNewJoinStyle = iCurJoinStyle Then
         MsgBox "Failed to change join style to new setting!", vbCritical, "Failed"
      Else
         MsgBox "Line join style has been changed to new setting.", vbInformation, "Success"
      End If
   Else
      MsgBox "Line join style was not changed.", vbInformation, "Line Join Style"
   End If
   
   ' De-initialize the object variable
   Set VectorProperties = Nothing
End Sub

Public Sub IVectorProperties_LineStyle()
   ' RobY Nov16/98
   ' Modified by RobY Dec2/98 - Setup so user can specify values
   Dim VectorProperties As IVectorProperties
   Dim iCurLineStyle As LINE_STYLE
   Dim strLineStyle As String
   Dim iNewLineStyle As LINE_STYLE
   Dim iGetNewLineStyle As LINE_STYLE
   Dim iResponse As Integer
   Dim Scope As PROPERTY_SCOPE
   Dim lObjectID As Long
   Dim iLayerNum As Integer
   Dim iToolType As TOOL_TYPE
   Dim lLayerID As Long
   
   ' Set object variable for IVectorProperties interface to Markup ctrl object
   Set VectorProperties = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Initialize variables
   lLayerID = 0
   lObjectID = 0

   ' Prompt for scope to use
   Scope = InputBox("Please enter the scope that you want to use. Check help file for values.", "Scope")
   If ((Scope >= 128) And (Scope < 256)) Or (Scope >= 384) Then
      iLayerNum = InputBox("Please enter the layer number of where the the object to change.", "Layer Number")
      ' Prompt for object id
      lObjectID = InputBox("Please enter the object id of the object to change.", "Object ID")
      ' Get the layer id of the specified layer
      lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   End If
   ' Prompt for tool
   iToolType = InputBox("Please enter the tool type. Setting is ignored depending on the scope.", "Tool Type")
   ' Get the line style
   iCurLineStyle = VectorProperties.LineStyle(iToolType, Scope, lLayerID, lObjectID)
   LineStyles iCurLineStyle, strLineStyle
   iResponse = MsgBox("The line style is currently set to " + strLineStyle + "." + Chr(13) + _
                        "Do you want to change the line style?", vbYesNo, "Line Style")
   If iResponse = vbYes Then
      iNewLineStyle = InputBox("Please enter new line style value." + Chr(13) + "SOLID = 1" + Chr(13) + _
                           "DASH = 2" + Chr(13) + "DOT = 3" + Chr(13) + "NULL = 4", "New Line Style")
      ' Set the new line style
      VectorProperties.LineStyle(iToolType, Scope, lLayerID, lObjectID) = iNewLineStyle
      ' Get the new line style
      iGetNewLineStyle = VectorProperties.LineStyle(iToolType, Scope, lLayerID, lObjectID)
      LineStyles iGetNewLineStyle, strLineStyle
      If iNewLineStyle = iGetNewLineStyle Then
         MsgBox "Line style has been changed to " + strLineStyle + ".", vbInformation, "Success"
      Else
         MsgBox "Failed to change line style to " + strLineStyle + "!", vbCritical, "Failed"
      End If
   Else
      MsgBox "Line style was not changed.", vbInformation, "Line Style"
   End If
   
   ' De-initialize the object variable
   Set VectorProperties = Nothing
End Sub

Public Sub IVectorProperties_LoadSymbolDialog()
   ' RobY Aug12/98
   ' Modified by RobY Nov20/98 Added scope and layerid parameters
   ' Modified by RobY Dec2/98 - Setup so user can specify values
   Dim VectorProperties As IVectorProperties
   Dim Scope As PROPERTY_SCOPE
   Dim iLayerNum As Integer
   Dim lLayerID As Long
   
   ' Set object variable for IVectorProperties interface to Markup ctrl object
   Set VectorProperties = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Initialize variables
   lLayerID = 0

   ' Prompt for scope to use
   Scope = InputBox("Please enter the scope that you want to use. Check help file for values.", "Scope")
   If ((Scope >= 128) And (Scope < 256)) Or (Scope >= 384) Then
      iLayerNum = InputBox("Please enter the layer number of where the the object to change.", "Layer Number")
      ' Get the layer id of the specified layer
      lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   End If
   ' Display the dialog
   VectorProperties.LoadSymbolDialog Scope, lLayerID
   
   ' De-initialize the object variable
   Set VectorProperties = Nothing
End Sub

Public Sub IVectorProperties_ProportionalArrowHead()
   ' RobY Nov18/98
   ' Modified by RobY Dec2/98 - Setup so user can specify values
   Dim VectorProperties As IVectorProperties
   Dim bOrigPropArrowHead As Boolean
   Dim bNewPropArrowHead As Boolean
   Dim iResponse As Integer
   Dim Scope As PROPERTY_SCOPE
   Dim lObjectID As Long
   Dim iLayerNum As Integer
   Dim iToolType As TOOL_TYPE
   Dim lLayerID As Long
   
   ' Set object variable for IVectorProperties interface to Markup ctrl object
   Set VectorProperties = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Initialize variables
   lLayerID = 0
   lObjectID = 0

   ' Prompt for scope to use
   Scope = InputBox("Please enter the scope that you want to use. Check help file for values.", "Scope")
   If ((Scope >= 128) And (Scope < 256)) Or (Scope >= 384) Then
      iLayerNum = InputBox("Please enter the layer number of where the the object to change.", "Layer Number")
      ' Prompt for object id
      lObjectID = InputBox("Please enter the object id of the object to change.", "Object ID")
      ' Get the layer id of the specified layer
      lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   End If
   ' Prompt for tool
   iToolType = InputBox("Please enter the tool type. Setting is ignored depending on the scope.", "Tool Type")
   'Get the original setting
   bOrigPropArrowHead = VectorProperties.ProportionalArrowHead(iToolType, Scope, lLayerID, lObjectID)
   ' Prompt user
   iResponse = MsgBox("The proportional value is currently set to " + Str(bOrigPropArrowHead) + "." + Chr(13) + _
                        "Do you want to change the current setting?", vbYesNo + vbInformation, "ProportionalArrowHead")
   If iResponse = vbYes Then
      If bOrigPropArrowHead = False Then
         ' Set to true if false
         VectorProperties.ProportionalArrowHead(iToolType, Scope, lLayerID, lObjectID) = True
      Else
         ' Set to false if true
         VectorProperties.ProportionalArrowHead(iToolType, Scope, lLayerID, lObjectID) = False
      End If
      ' Get the new setting
      bNewPropArrowHead = VectorProperties.ProportionalArrowHead(iToolType, Scope, lLayerID, lObjectID)
      ' Check to make sure that setting was changed
      If bOrigPropArrowHead = bNewPropArrowHead Then
         MsgBox "Failed to change to new setting!", vbCritical, "Failed"
      Else
         MsgBox "The proportional value has been changed to" + Str(bNewPropArrowHead) + ".", vbInformation, "Success"
      End If
   Else
      MsgBox "The proportional value not changed.", vbInformation, "ProportionalArrowHead"
   End If
   
   ' De-initialize the object variable
   Set VectorProperties = Nothing
End Sub

Public Sub IVectorProperties_SetArrowHeadLength()
   ' RobY Nov19/98
   ' Modified by RobY Dec2/98 - Setup so user can specify values
   Dim VectorProperties As IVectorProperties
   Dim iUnitType As UNIT_TYPE
   Dim dLength As Double
   Dim Scope As PROPERTY_SCOPE
   Dim lObjectID As Long
   Dim iLayerNum As Integer
   Dim iToolType As TOOL_TYPE
   Dim lLayerID As Long
   
   ' Set object variable for IVectorProperties interface to Markup ctrl object
   Set VectorProperties = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Initialize variables
   lLayerID = 0
   lObjectID = 0

   ' Prompt for scope to use
   Scope = InputBox("Please enter the scope that you want to use. Check help file for values.", "Scope")
   If ((Scope >= 128) And (Scope < 256)) Or (Scope >= 384) Then
      iLayerNum = InputBox("Please enter the layer number of where the the object to change.", "Layer Number")
      ' Prompt for object id
      lObjectID = InputBox("Please enter the object id of the object to change.", "Object ID")
      ' Get the layer id of the specified layer
      lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   End If
   ' Prompt for tool
   iToolType = InputBox("Please enter the tool type. Setting is ignored depending on the scope.", "Tool Type")
   dLength = InputBox("Please enter a value to set the arrow head length to.", "SetArrowHeadLength")
   iUnitType = InputBox("Please enter the unit type of the new value." + Chr(13) + _
                  "1 = inches, 2 = cm, 3 = pixels, 4 = feet, 5 = mm, 6 = m," + Chr(13) + _
                  "7 = points, 8 = twips, 9 = custom 1, 10 = custom 2," + Chr(13) + _
                  "11 = custom 3", "SetArrowHeadLength[Unit Type]")
   ' Set the arrow head length
   VectorProperties.SetArrowHeadLength iToolType, Scope, lLayerID, lObjectID, dLength, iUnitType
   MsgBox "Set Length to: " + Str(dLength) + Chr(13) + "Set Unit Type to: " + Str(iUnitType), vbInformation, "SetArrowHeadLength"
   
   ' De-initialize the object variable
   Set VectorProperties = Nothing
End Sub

Public Sub IVectorProperties_SetLineThickness()
   ' RobY Nov19/98
   ' Modified by RobY Dec2/98 - Setup so user can specify values
   Dim VectorProperties As IVectorProperties
   Dim iUnitType As UNIT_TYPE
   Dim dThickness As Double
   Dim Scope As PROPERTY_SCOPE
   Dim lObjectID As Long
   Dim iLayerNum As Integer
   Dim iToolType As TOOL_TYPE
   Dim lLayerID As Long
   
   ' Set object variable for IVectorProperties interface to Markup ctrl object
   Set VectorProperties = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Initialize variables
   lLayerID = 0
   lObjectID = 0

   ' Prompt for scope to use
   Scope = InputBox("Please enter the scope that you want to use. Check help file for values.", "Scope")
   If ((Scope >= 128) And (Scope < 256)) Or (Scope >= 384) Then
      iLayerNum = InputBox("Please enter the layer number of where the the object to change.", "Layer Number")
      ' Prompt for object id
      lObjectID = InputBox("Please enter the object id of the object to change.", "Object ID")
      ' Get the layer id of the specified layer
      lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   End If
   ' Prompt for tool
   iToolType = InputBox("Please enter the tool type. Setting is ignored depending on the scope.", "Tool Type")
   dThickness = InputBox("Please enter a value to set the line thickness to.", "SetLineThickness")
   iUnitType = InputBox("Please enter the unit type of the new value." + Chr(13) + _
                  "1 = inches, 2 = cm, 3 = pixels, 4 = feet, 5 = mm, 6 = m," + Chr(13) + _
                  "7 = points, 8 = twips, 9 = custom 1, 10 = custom 2," + Chr(13) + _
                  "11 = custom 3", "SetLineThickness[Unit Type]")
   ' Set the line thickness for the session
   VectorProperties.SetLineThickness iToolType, Scope, lLayerID, lObjectID, dThickness, iUnitType
   MsgBox "Set Thickness to: " + Str(dThickness) + Chr(13) + "Set Unit Type to: " + Str(iUnitType), vbInformation, "SetLineThickness"
   
   ' De-initialize the object variable
   Set VectorProperties = Nothing
End Sub

Public Sub IVectorProperties_SetTextHeight()
   ' RobY Nov19/98
   ' Modified by RobY Dec2/98 - Setup so user can specify values
   Dim VectorProperties As IVectorProperties
   Dim iUnitType As UNIT_TYPE
   Dim dHeight As Double
   Dim Scope As PROPERTY_SCOPE
   Dim lObjectID As Long
   Dim iLayerNum As Integer
   Dim iToolType As TOOL_TYPE
   Dim lLayerID As Long
   
   ' Set object variable for IVectorProperties interface to Markup ctrl object
   Set VectorProperties = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Initialize variables
   lLayerID = 0
   lObjectID = 0

   ' Prompt for scope to use
   Scope = InputBox("Please enter the scope that you want to use. Check help file for values.", "Scope")
   If ((Scope >= 128) And (Scope < 256)) Or (Scope >= 384) Then
      iLayerNum = InputBox("Please enter the layer number of where the the object to change.", "Layer Number")
      ' Prompt for object id
      lObjectID = InputBox("Please enter the object id of the object to change.", "Object ID")
      ' Get the layer id of the specified layer
      lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   End If
   ' Prompt for tool
   iToolType = InputBox("Please enter the tool type. Setting is ignored depending on the scope.", "Tool Type")
   dHeight = InputBox("Please enter a value to set the text height to.", "SetTextHeight")
   iUnitType = InputBox("Please enter the unit type of the new value." + Chr(13) + _
                  "1 = inches, 2 = cm, 3 = pixels, 4 = feet, 5 = mm, 6 = m," + Chr(13) + _
                  "7 = points, 8 = twips, 9 = custom 1, 10 = custom 2," + Chr(13) + _
                  "11 = custom 3", "SetTextHeight[Unit Type]")
   ' Set the text height
   VectorProperties.SetTextHeight iToolType, Scope, lLayerID, lObjectID, dHeight, iUnitType
   MsgBox "Set text height to: " + Str(dHeight) + Chr(13) + "Set Unit Type to: " + Str(iUnitType), vbInformation, "SetTextHeight"
   
   ' De-initialize the object variable
   Set VectorProperties = Nothing
End Sub

Public Sub IVectorProperties_SetTextWidth()
   ' RobY Nov19/98
   ' Modified by RobY Dec2/98 - Setup so user can specify values
   Dim VectorProperties As IVectorProperties
   Dim iUnitType As UNIT_TYPE
   Dim dWidth As Double
   Dim Scope As PROPERTY_SCOPE
   Dim lObjectID As Long
   Dim iLayerNum As Integer
   Dim iToolType As TOOL_TYPE
   Dim lLayerID As Long
   
   ' Set object variable for IVectorProperties interface to Markup ctrl object
   Set VectorProperties = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Initialize variables
   lLayerID = 0
   lObjectID = 0

   ' Prompt for scope to use
   Scope = InputBox("Please enter the scope that you want to use. Check help file for values.", "Scope")
   If ((Scope >= 128) And (Scope < 256)) Or (Scope >= 384) Then
      iLayerNum = InputBox("Please enter the layer number of where the the object to change.", "Layer Number")
      ' Prompt for object id
      lObjectID = InputBox("Please enter the object id of the object to change.", "Object ID")
      ' Get the layer id of the specified layer
      lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   End If
   ' Prompt for tool
   iToolType = InputBox("Please enter the tool type. Setting is ignored depending on the scope.", "Tool Type")
   dWidth = InputBox("Please enter a value to set the text width to.", "SetTextWidth")
   iUnitType = InputBox("Please enter the unit type of the new value." + Chr(13) + _
                  "1 = inches, 2 = cm, 3 = pixels, 4 = feet, 5 = mm, 6 = m," + Chr(13) + _
                  "7 = points, 8 = twips, 9 = custom 1, 10 = custom 2," + Chr(13) + _
                  "11 = custom 3", "SetTextWidth[Unit Type]")
   ' Set the text width
   VectorProperties.SetTextWidth iToolType, Scope, lLayerID, lObjectID, dWidth, iUnitType
   MsgBox "Set text width to: " + Str(dWidth) + Chr(13) + "Set Unit Type to: " + Str(iUnitType), vbInformation, "SetTextWidth"
   
   ' De-initialize the object variable
   Set VectorProperties = Nothing
End Sub

Public Sub IVectorProperties_SolidArrowHead()
   ' RobY Nov20/98
   ' Modified by RobY Dec2/98 - Setup so user can specify values
   Dim VectorProperties As IVectorProperties
   Dim bOrigSolidArrowHead As Boolean
   Dim bNewSolidArrowHead As Boolean
   Dim iResponse As Integer
   Dim Scope As PROPERTY_SCOPE
   Dim lObjectID As Long
   Dim iLayerNum As Integer
   Dim iToolType As TOOL_TYPE
   Dim lLayerID As Long
   
   ' Set object variable for IVectorProperties interface to Markup ctrl object
   Set VectorProperties = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Initialize variables
   lLayerID = 0
   lObjectID = 0

   ' Prompt for scope to use
   Scope = InputBox("Please enter the scope that you want to use. Check help file for values.", "Scope")
   If ((Scope >= 128) And (Scope < 256)) Or (Scope >= 384) Then
      iLayerNum = InputBox("Please enter the layer number of where the the object to change.", "Layer Number")
      ' Prompt for object id
      lObjectID = InputBox("Please enter the object id of the object to change.", "Object ID")
      ' Get the layer id of the specified layer
      lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   End If
   ' Prompt for tool
   iToolType = InputBox("Please enter the tool type. Setting is ignored depending on the scope.", "Tool Type")
   'Get the original setting
   bOrigSolidArrowHead = VectorProperties.SolidArrowHead(iToolType, Scope, lLayerID, lObjectID)
   ' Prompt user
   iResponse = MsgBox("The solid value for the arrow head tool is currently set to " + Str(bOrigSolidArrowHead) + "." + Chr(13) + _
                        "Do you want to change the current setting?", vbYesNo + vbInformation, "SolidArrowHead")
   If iResponse = vbYes Then
      If bOrigSolidArrowHead = False Then
         ' Set to true if false
         VectorProperties.SolidArrowHead(iToolType, Scope, lLayerID, lObjectID) = True
      Else
         ' Set to false if true
         VectorProperties.SolidArrowHead(iToolType, Scope, lLayerID, lObjectID) = False
      End If
      ' Get the new setting
      bNewSolidArrowHead = VectorProperties.SolidArrowHead(iToolType, Scope, lLayerID, lObjectID)
      ' Check to make sure that setting was changed
      If bOrigSolidArrowHead = bNewSolidArrowHead Then
         MsgBox "Failed to change to new setting!", vbCritical, "Failed"
      Else
         MsgBox "The solid value has been changed to" + Str(bNewSolidArrowHead) + ".", vbInformation, "Success"
      End If
   Else
      MsgBox "The solid value was not changed.", vbInformation, "SolidArrowHead"
   End If
   
   ' De-initialize the object variable
   Set VectorProperties = Nothing
End Sub

Public Sub IVectorProperties_TextJustification()
   ' RobY Nov20/98
   ' Modified by RobY Dec2/98 - Setup so user can specify values
   Dim VectorProperties As IVectorProperties
   Dim iNewTextJust As TEXT_JUSTIFICATION
   Dim iCurTextJust As TEXT_JUSTIFICATION
   Dim strTextJust As String
   Dim iResponse As Integer
   Dim iGetNewTextJust As TEXT_JUSTIFICATION
   Dim Scope As PROPERTY_SCOPE
   Dim lObjectID As Long
   Dim iLayerNum As Integer
   Dim iToolType As TOOL_TYPE
   Dim lLayerID As Long
   
   ' Set object variable for IVectorProperties interface to Markup ctrl object
   Set VectorProperties = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Initialize variables
   lLayerID = 0
   lObjectID = 0

   ' Prompt for scope to use
   Scope = InputBox("Please enter the scope that you want to use. Check help file for values.", "Scope")
   If ((Scope >= 128) And (Scope < 256)) Or (Scope >= 384) Then
      iLayerNum = InputBox("Please enter the layer number of where the the object to change.", "Layer Number")
      ' Prompt for object id
      lObjectID = InputBox("Please enter the object id of the object to change.", "Object ID")
      ' Get the layer id of the specified layer
      lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   End If
   ' Prompt for tool
   iToolType = InputBox("Please enter the tool type. Setting is ignored depending on the scope.", "Tool Type")
   iCurTextJust = VectorProperties.TextJustification(iToolType, Scope, lLayerID, lObjectID)
   strTextJust = BuildTextJustStr(iCurTextJust)
   iResponse = MsgBox("The text justification is currently set to: " + strTextJust + Chr(13) + _
                        "Do you want to change the current setting? ", vbInformation + vbYesNo, "TextJustification")
   If iResponse = vbYes Then
      iNewTextJust = InputBox("Please enter text justification value (can be ORed).", "Text Justification")
      VectorProperties.TextJustification(iToolType, Scope, lLayerID, lObjectID) = iNewTextJust
      ' Get the new setting
      iGetNewTextJust = VectorProperties.TextJustification(iToolType, Scope, lLayerID, lObjectID)
      ' Check to make sure that setting was changed
      If (iGetNewTextJust = iCurTextJust) And (iGetNewTextJust <> iNewTextJust) Then
         MsgBox "Failed to change to new setting!", vbCritical, "Failed"
      Else
         strTextJust = BuildTextJustStr(iGetNewTextJust)
         MsgBox "The text justification has been changed to:" + strTextJust, vbInformation, "Success"
      End If
   Else
      MsgBox "The text justification was not changed.", vbInformation, "TextJustification"
   End If
   
   ' De-initialize the object variable
   Set VectorProperties = Nothing
End Sub

Public Sub IVectorProperties_TextRotation()
   ' RobY Nov17/98
   ' Modified by RobY Dec2/98 - Setup so user can specify values
   Dim VectorProperties As IVectorProperties
   Dim Scope As PROPERTY_SCOPE
   Dim lObjectID As Long
   Dim iLayerNum As Integer
   Dim iToolType As TOOL_TYPE
   Dim lLayerID As Long
   
   ' Set object variable for IVectorProperties interface to Markup ctrl object
   Set VectorProperties = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Initialize variables
   lLayerID = 0
   lObjectID = 0

   ' Prompt for scope to use
   Scope = InputBox("Please enter the scope that you want to use. Check help file for values.", "Scope")
   If ((Scope >= 128) And (Scope < 256)) Or (Scope >= 384) Then
      iLayerNum = InputBox("Please enter the layer number of where the the object to change.", "Layer Number")
      ' Prompt for object id
      lObjectID = InputBox("Please enter the object id of the object to change.", "Object ID")
      ' Get the layer id of the specified layer
      lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   End If
   ' Prompt for tool
   iToolType = InputBox("Please enter the tool type. Setting is ignored depending on the scope.", "Tool Type")
   ' Set the rotation for the text tool to 45 degrees for the session
   VectorProperties.TextRotation(iToolType, Scope, lLayerID, lObjectID) = 45
   MsgBox "The rotation has been set to 45 degrees.", vbInformation, "Rotate Text"
   
   ' De-initialize the object variable
   Set VectorProperties = Nothing
End Sub

Public Sub IVectorProperties_TextTypeFace()
   ' RobY Nov20/98
   ' Modified by RobY Dec2/98 - Setup so user can specify values
   Dim VectorProperties As IVectorProperties
   Dim iNewTypeFace As TEXT_TYPEFACE
   Dim iCurTypeFace As TEXT_TYPEFACE
   Dim strTypeFace As String
   Dim iResponse As Integer
   Dim iGetNewTypeFace As TEXT_TYPEFACE
   Dim Scope As PROPERTY_SCOPE
   Dim lObjectID As Long
   Dim iLayerNum As Integer
   Dim iToolType As TOOL_TYPE
   Dim lLayerID As Long
   
   ' Set object variable for IVectorProperties interface to Markup ctrl object
   Set VectorProperties = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Initialize variables
   lLayerID = 0
   lObjectID = 0

   ' Prompt for scope to use
   Scope = InputBox("Please enter the scope that you want to use. Check help file for values.", "Scope")
   If ((Scope >= 128) And (Scope < 256)) Or (Scope >= 384) Then
      iLayerNum = InputBox("Please enter the layer number of where the the object to change.", "Layer Number")
      ' Prompt for object id
      lObjectID = InputBox("Please enter the object id of the object to change.", "Object ID")
      ' Get the layer id of the specified layer
      lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   End If
   ' Prompt for tool
   iToolType = InputBox("Please enter the tool type. Setting is ignored depending on the scope.", "Tool Type")
   iCurTypeFace = VectorProperties.TextTypeFace(iToolType, Scope, lLayerID, lObjectID)
   strTypeFace = BuildTypeFaceStr(iCurTypeFace)
   iResponse = MsgBox("The type face is currently set to: " + strTypeFace + Chr(13) + _
                        "Do you want to change the current setting? ", vbInformation + vbYesNo, "TextTypeFace")
   If iResponse = vbYes Then
      iNewTypeFace = InputBox("Please enter type face value (can be ORed)." + Chr(13) + "Normal = 1" + Chr(13) + _
                           "Bold = 2" + Chr(13) + "Italic = 16" + Chr(13) + "Underline = 32" + Chr(13) + "Strikeout = 64", "New Type Face")
      VectorProperties.TextTypeFace(iToolType, Scope, lLayerID, lObjectID) = iNewTypeFace
      ' Get the new setting
      iGetNewTypeFace = VectorProperties.TextTypeFace(iToolType, Scope, lLayerID, lObjectID)
      ' Check to make sure that setting was changed
      If (iGetNewTypeFace <> iNewTypeFace) Then
         MsgBox "Failed to change to new setting!", vbCritical, "Failed"
      Else
         strTypeFace = BuildTypeFaceStr(iGetNewTypeFace)
         MsgBox "The type face has been changed to:" + strTypeFace, vbInformation, "Success"
      End If
   Else
      MsgBox "The type face was not changed.", vbInformation, "TextTypeFace"
   End If
     
   ' De-initialize the object variable
   Set VectorProperties = Nothing
End Sub

Public Sub IVectorProperties_ToolPreferencesDialog()
   ' RobY Aug12/98
   ' Modified by RobY Dec1/98 - Setup so user can specify values
   ' Modified by RobY Dec 16/98 - Rewrote code to include availability and to allow user to specify scope(new parameter)
   Dim VectorProperties As IVectorProperties
   Dim Scope As PROPERTY_SCOPE
   Dim Available As COMMAND_AVAILABILITY
   
   ' Set object variable for IVectorProperties interface to Markup ctrl object
   Set VectorProperties = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Prompt for scope to use
   Scope = InputBox("Please enter the scope that you want to use. Check help file for values.", "Scope")
   ' Get the value of availability
   Available = VectorProperties.ToolPreferencesDialogAvailability(Scope)
   ' Do the case statement that corresponds to the availability returned
   Select Case Available
      Case IN_UI_ENABLED
         MsgBox "Command is ENABLED. Click OK to display dialog.", vbInformation, "Availability"
         ' Display the dialog
         VectorProperties.ToolPreferencesDialog Scope
      Case IN_UI_GREYED
         MsgBox "Command is GREYED.", vbInformation, "Availability"
      Case IN_UI_REMOVED
         MsgBox "Command is REMOVED.", vbInformation, "Availability"
      Case IN_UI_CHECKED
         MsgBox "Command is CHECKED.", vbInformation, "Availability"
      Case IN_UI_ENABLED_CHECKED
         MsgBox "Command is ENABLED and CHECKED. Click OK to display dialog.", vbInformation, "Availability"
         ' Display the dialog
         VectorProperties.ToolPreferencesDialog Scope
      Case Else
         MsgBox "Cannot open dialog. Check to make that the scope you specified is valid.", vbExclamation, "Availability"
   End Select
   
   ' De-initialize the object variable
   Set VectorProperties = Nothing
End Sub

Public Sub IVectorProperties_VectorObjectID()
   ' RobY Nov6/98
   ' Modified by RobY Dec2/98 - Setup so user can specify values
   ' Modified by RobY Dec10/98 - Commented out code that uses IN_PROPSCOPE_CURVECTS. For future use.
   Dim VectorProperties As IVectorProperties
   Dim lGetObjectID As Long
   Dim lNewObjectID As Long
   Dim iResponse As Integer
   Dim Scope As PROPERTY_SCOPE
   Dim lObjectID As Long
   Dim iLayerNum As Integer
   Dim lLayerID As Long
   
   ' Set object variable for IVectorProperties interface to Markup ctrl object
   Set VectorProperties = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Initialize variables
   lLayerID = 0
   lObjectID = 0

' Commented out by RobY - Used quotes down left side
   ' Get the object id
'   lGetObjectID = VectorProperties.VectorObjectID(IN_PROPSCOPE_CURVECTS, lLayerID, lObjectID)
'   iResponse = MsgBox("The Vector Object ID is " + Str(lGetObjectID) + Chr(13) + _
                        "Do you want to change it? ", vbInformation + vbYesNo, "VectorObjectID")
'   If iResponse = vbYes Then
      ' Prompt for scope to use
      Scope = InputBox("Please enter the scope that you want to use. Check help file for values.", "Scope")
      If ((Scope >= 128) And (Scope < 256)) Or (Scope >= 384) Then
         iLayerNum = InputBox("Please enter the layer number of where the the object to change.", "Layer Number")
         ' Get the layer id of the specified layer
         lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
' Remove next line when uncommenting code that was commented by RobY
         lGetObjectID = InputBox("Please enter object id of object to change.", "Object ID")
      End If
      lNewObjectID = InputBox("Please enter a new object id.", "New Object ID")
      VectorProperties.VectorObjectID(Scope, lLayerID, lGetObjectID) = lNewObjectID
' Remove next line when uncommenting code that was commented by RobY
      MsgBox "Vector object id had been changed.", vbInformation, "New Object ID"
      ' Get the new setting
'      lObjectID = VectorProperties.VectorObjectID(SCOPE, lLayerID, lNewObjectID)
      ' Check to make sure that setting was changed
'      If (lObjectID = lGetObjectID) Or (lObjectID <> lNewObjectID) Then
'         MsgBox "Failed to change to new setting!", vbCritical, "Failed"
'      Else
'         MsgBox "The vector object id has been changed to:" + Str(lObjectID), vbInformation, "Success"
'      End If
'   Else
'      MsgBox "The vector object id was not changed.", vbInformation, "VectorObjectID"
'   End If
   
   ' De-initialize the object variable
   Set VectorProperties = Nothing
End Sub

Public Sub IVectorProperties_VectorObjectState()
   ' RobY Nov18/98
   ' Modified by RobY Dec2/98 - Setup so user can specify values
   Dim VectorProperties As IVectorProperties
   Dim iVectorState As VECTOR_STATE
   Dim iNewVectorState As VECTOR_STATE
   Dim iGetNewVectorState As VECTOR_STATE
   Dim Scope As PROPERTY_SCOPE
   Dim lObjectID As Long
   Dim iLayerNum As Integer
   Dim lLayerID As Long
   Dim iResponse As Integer
   
   ' Set object variable for IVectorProperties interface to Markup ctrl object
   Set VectorProperties = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Initialize variables
   lLayerID = 0
   lObjectID = 0

   ' Prompt for scope to use
   Scope = InputBox("Please enter the scope that you want to use. Check help file for values.", "Scope")
   If ((Scope >= 128) And (Scope < 256)) Or (Scope >= 384) Then
      iLayerNum = InputBox("Please enter the layer number of where the the object to change.", "Layer Number")
      ' Prompt for object id
      lObjectID = InputBox("Please enter the object id of the object to change.", "Object ID")
      ' Get the layer id of the specified layer
      lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   End If
   iVectorState = VectorProperties.VectorObjectState(Scope, lLayerID, lObjectID)
   iResponse = MsgBox("The vector object state is " + Str(iVectorState) + Chr(13) + _
                        "Do you want to change it? ", vbInformation + vbYesNo, "VectorObjectID")
   If iResponse = vbYes Then
      iNewVectorState = InputBox("Please enter the new  vector object state.", "New Object State")
      VectorProperties.VectorObjectState(Scope, lLayerID, lObjectID) = iNewVectorState
      ' Get the new setting
      iGetNewVectorState = VectorProperties.VectorObjectState(Scope, lLayerID, lObjectID)
      ' Check to make sure that setting was changed
      If (iGetNewVectorState = iVectorState) Or (iGetNewVectorState <> iNewVectorState) Then
         MsgBox "Failed to change to new setting!", vbCritical, "Failed"
      Else
         MsgBox "The vector object state has been changed to:" + Str(iGetNewVectorState), vbInformation, "Success"
      End If
   Else
      MsgBox "The vector object state was not changed.", vbInformation, "VectorObjectState"
   End If
   
   ' De-initialize the object variable
   Set VectorProperties = Nothing
End Sub

Public Sub IVectorProperties_VectorObjectType()
   ' RobY Nov24/98
   ' Modified by RobY Dec2/98 - Setup so user can specify values
   Dim VectorProperties As IVectorProperties
   Dim iToolType As TOOL_TYPE
   Dim Scope As PROPERTY_SCOPE
   Dim lObjectID As Long
   Dim iLayerNum As Integer
   Dim lLayerID As Long
   
   ' Set object variable for IVectorProperties interface to Markup ctrl object
   Set VectorProperties = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Initialize variables
   lLayerID = 0
   lObjectID = 0

   ' Prompt for scope to use
   Scope = InputBox("Please enter the scope that you want to use. Check help file for values.", "Scope")
   If ((Scope >= 128) And (Scope < 256)) Or (Scope >= 384) Then
      iLayerNum = InputBox("Please enter the layer number of where the the object to change.", "Layer Number")
      ' Prompt for object id
      lObjectID = InputBox("Please enter the object id of the object to change.", "Object ID")
      ' Get the layer id of the specified layer
      lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   End If
   iToolType = VectorProperties.VectorObjectType(Scope, lLayerID, lObjectID)
   MsgBox "The vector object is a " + Str(iToolType), vbInformation, "VectorObjectType"
   
   ' De-initialize the object variable
   Set VectorProperties = Nothing
End Sub

Public Function LineStyles(iLineStyle As LINE_STYLE, strLineStyle As String)
   'RobY Nov16/98
   Select Case iLineStyle
      Case IN_LINE_SOLID
         strLineStyle = "SOLID"
      Case IN_LINE_DASH
         strLineStyle = "DASH"
      Case IN_LINE_DOT
         strLineStyle = "DOT"
      Case IN_LINE_NULL
         strLineStyle = "NULL"
   End Select
End Function

Public Function BuildTypeFaceStr(iTypeFace As TEXT_TYPEFACE) As String
   'RobY Nov20/98
   Dim strTypeFace As String
   
   ' Initialize string to new line
   strTypeFace = Chr(13)
   ' Build type face string
   If iTypeFace > 63 Then
      strTypeFace = strTypeFace + "Strike Out " + Chr(13)
      iTypeFace = iTypeFace - 64
   End If
   If iTypeFace > 31 Then
      strTypeFace = strTypeFace + "Underline " + Chr(13)
      iTypeFace = iTypeFace - 32
   End If
   If iTypeFace > 15 Then
      strTypeFace = strTypeFace + "Italic " + Chr(13)
      iTypeFace = iTypeFace - 16
   End If
   If iTypeFace > 1 Then
      strTypeFace = strTypeFace + "Bold " + Chr(13)
      iTypeFace = iTypeFace - 2
   End If
   If iTypeFace = 1 Then
      strTypeFace = strTypeFace + "Normal"
   End If
   BuildTypeFaceStr = strTypeFace
End Function

Public Function BuildTextJustStr(iTextJust As TEXT_JUSTIFICATION) As String
   'RobY Nov20/98
   Dim strTextJust As String
   
   ' Initialize string to new line
   strTextJust = Chr(13)
   ' Build text justification string
   If iTextJust > 47 Then
      strTextJust = strTextJust + "Middle Vertical " + Chr(13)
      iTextJust = iTextJust - 48
   ElseIf iTextJust > 31 Then
      strTextJust = strTextJust + "Top " + Chr(13)
      iTextJust = iTextJust - 32
   ElseIf iTextJust > 15 Then
      strTextJust = strTextJust + "Bottom " + Chr(13)
      iTextJust = iTextJust - 16
   End If
   If iTextJust > 5 Then
      strTextJust = strTextJust + "Fit " + Chr(13)
      iTextJust = iTextJust - 6
   ElseIf iTextJust > 4 Then
      strTextJust = strTextJust + "Middle Horizontal " + Chr(13)
      iTextJust = iTextJust - 5
   ElseIf iTextJust > 3 Then
      strTextJust = strTextJust + "Aligned " + Chr(13)
      iTextJust = iTextJust - 4
   ElseIf iTextJust > 2 Then
      strTextJust = strTextJust + "Right " + Chr(13)
      iTextJust = iTextJust - 3
   ElseIf iTextJust > 1 Then
      strTextJust = strTextJust + "Center " + Chr(13)
      iTextJust = iTextJust - 2
   ElseIf iTextJust > 0 Then
      strTextJust = strTextJust + "Left " + Chr(13)
      iTextJust = iTextJust - 1
   ElseIf iTextJust = 0 Then
      strTextJust = strTextJust + "Baseline " + Chr(13)
   End If
   BuildTextJustStr = strTextJust
End Function

