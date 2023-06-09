Attribute VB_Name = "mod_IHotspots"
' File:      IHotspots.bas
' Created:   1998July30 by Rob Young
' Purpose:   To test the Spicer Markup Control's IHotspots interface.
' Revisions: RobY Jan4/99  - Added functionality to open more than 1 file. ChildForm1 is no longer used access objects because it has been set to ChildForm() which is an array of ChildForm1's forms.
'                          - Anywhere ChildForm1 was used has been changed to ChildForm(iIndex) or MainMDIForm.ActiveForm.
'   $log$
'   $nokeywords$
'
' Copyright (C) 1998 SPICER Corporation, All rights reserved.
Option Explicit
Dim MenuName As Menu

Public Sub IHotspots_AttachHotspotDialog()
   ' RobY Aug13/98
   Dim HotSpots As IHotspots
   
   ' Set object variable for IHotSpots interface to Markup ctrl object
   Set HotSpots = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Display the dialog to attach hotspot data to the currently selected vector object.
   HotSpots.AttachHotspotDialog
   
   ' De-initialize the object variable
   Set HotSpots = Nothing
End Sub

Public Sub IHotspots_BindToDocumentControl()
   ' RobY Aug13/98
   Dim HotSpots As IHotspots
   
   ' Set object variable for IHotSpots interface to Markup ctrl object
   Set HotSpots = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   'Bind the Spicer Markup Control to the Spicer View Control
   HotSpots.BindToDocumentControl MainMDIForm.ActiveForm.SpicerDoc1.object
   
   ' De-initialize the object variable
   Set HotSpots = Nothing
End Sub

Public Sub IHotspots_BindToViewControl()
   ' RobY Aug13/98
   Dim HotSpots As IHotspots
   
   ' Set object variable for IHotSpots interface to Markup ctrl object
   Set HotSpots = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   'Bind the Spicer Markup Control to the Spicer View Control
   HotSpots.BindToViewControl MainMDIForm.ActiveForm.SpicerView1.object
   
   ' De-initialize the object variable
   Set HotSpots = Nothing
End Sub

Public Sub IHotspots_ChangeHotspotDialog()
   ' RobY Aug14/98
   Dim HotSpots As IHotspots
   
   ' Set object variable for IHotSpots interface to Markup ctrl object
   Set HotSpots = MainMDIForm.ActiveForm.SpicerMarkup1.object
    
   'Display the dialog to change the hotspot data for the currently selected vector object
   HotSpots.ChangeHotspotDialog
   
   ' De-initialize the object variable
   Set HotSpots = Nothing
End Sub

Public Sub IHotspots_ConvertHotspotIDToVectorObjectID()
   ' RobY Jan26/99
   Dim HotSpots As IHotspots
   Dim lLayerID As Long
   Dim iLayerNum As Integer
   Dim strHotspotID As String
   Dim lVectorObjectID As Long
   
   ' Set object variable for IHotSpots interface to Markup ctrl object
   Set HotSpots = MainMDIForm.ActiveForm.SpicerMarkup1.object
      
   iLayerNum = InputBox("Please enter the layer number of where the hotspot is located.", "Delete Hotspot")
   strHotspotID = InputBox("Please enter the id of the hotspot you would like to get vector object id for.", "Hotspot ID")
   lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   HotSpots.ConvertHotspotIDToVectorObjectID lLayerID, strHotspotID, lVectorObjectID
   MsgBox "HotspotID: " + strHotspotID + Chr(13) + "VectorObjectID: " + Str(lVectorObjectID), vbInformation, "VectorObjectID"
   
   ' De-initialize the object variable
   Set HotSpots = Nothing
End Sub

Public Sub IHotspots_ConvertVectorObjectIDToHotspotID()
   ' RobY Jan26/99
   Dim HotSpots As IHotspots
   Dim lLayerID As Long
   Dim iLayerNum As Integer
   Dim strHotspotID As String
   Dim lVectorObjectID As Long
   
   ' Set object variable for IHotSpots interface to Markup ctrl object
   Set HotSpots = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   iLayerNum = InputBox("Please enter the layer number of where the hotspot is located.", "Delete Hotspot")
   lVectorObjectID = InputBox("Please enter the vector object id of the hotspot to get the hotspot id of.", "VectorObjectID")
   lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   HotSpots.ConvertVectorObjectIDToHotspotID lLayerID, lVectorObjectID, strHotspotID
   MsgBox "VectorObjectID: " + Str(lVectorObjectID) + Chr(13) + "HotspotID: " + strHotspotID, vbInformation, "HotspotID"
   
   ' De-initialize the object variable
   Set HotSpots = Nothing
End Sub

Public Sub IHotspots_DeleteHotspot()
   ' RobY Aug14/98
   Dim HotSpots As IHotspots
   Dim lLayerID As Long
   Dim iLayerNum As Integer
   Dim strHotspotID As String
   
   ' Set object variable for IHotSpots interface to Markup ctrl object
   Set HotSpots = MainMDIForm.ActiveForm.SpicerMarkup1.object
    
   iLayerNum = InputBox("Please enter the layer number of where the hotspot is located.", "Delete Hotspot")
   strHotspotID = InputBox("Please enter the id of the hotspot you would like to delete.", "Delete Hotspot")
   lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   ' Delete the identified hotspot
   HotSpots.DeleteHotspot lLayerID, strHotspotID
   
   ' De-initialize the object variable
   Set HotSpots = Nothing
End Sub

Public Sub IHotspots_HotspotType()
   ' RobY Nov23/98
   Dim HotSpots As IHotspots
   Dim lLayerID As Long
   Dim strHotspotID As String
   Dim iObjectType As TOOL_TYPE
   Dim strObjectType As String
   
   ' Set object variable for IHotSpots interface to Markup ctrl object
   Set HotSpots = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   lLayerID = InputBox("Please specify the id of the layer where the hotspot is located.", "LayerID")
   strHotspotID = InputBox("Please specify the id of hotspot that you want to get the type for.", "HotspotID")
   ' Find out the type of hotspot
   iObjectType = HotSpots.HotspotType(lLayerID, strHotspotID)
   ' Convert the value of objectType to a string
   Select Case iObjectType
      Case IN_TOOL_BOX
         strObjectType = "a box."
      Case IN_TOOL_CIRCLE
         strObjectType = "a circle."
      Case IN_TOOL_ELLIPSE
         strObjectType = "an ellipse."
      Case IN_TOOL_ICON
         strObjectType = "an icon."
      Case IN_TOOL_LINE
         strObjectType = "a line."
   End Select
   ' Display a message stating what type of hotspot it is
   MsgBox "The specified hotspot is " + strObjectType, vbInformation
   
   ' De-initialize the object variable
   Set HotSpots = Nothing
End Sub

Public Sub IHotspots_PlaceBoxHotspot()
   ' RobY Nov23/98
   Dim HotSpots As IHotspots
   Dim lLayerID As Long
   Dim iLayerNum As Integer
   Dim lObjectID As Long
   Dim strHotspotID As String
   Dim dX1 As Double
   Dim dY1 As Double
   Dim dX2 As Double
   Dim dY2 As Double
   Dim dThickness As Double
   Dim iThicknessUnits As UNIT_TYPE
   Dim lLineColor As Long
   Dim lFillColor As Long
   Dim iLineStyle As LINE_STYLE
   Dim iJointStyle As LINE_JOIN_STYLE
   Dim iFillStyle As FILL_STYLE
   
   dX1 = 2.3
   dY1 = 1.2
   dX2 = 2.6
   dY2 = 1.3
   dThickness = 0.005
   iThicknessUnits = IN_UNITS_INCH
   lLineColor = 255
   lFillColor = 255
   iLineStyle = IN_LINE_SOLID
   iJointStyle = IN_JOIN_ROUND
   iFillStyle = IN_FILL_TRANSPARENT
   
   ' Set object variable for IHotSpots interface to Markup ctrl object
   Set HotSpots = MainMDIForm.ActiveForm.SpicerMarkup1.object
    
   iLayerNum = InputBox("Please enter the layer number of where to place the hotspot.", "Place Hotspot")
   lObjectID = InputBox("Please enter the vector object id of the hotspot you going to place.", "Place Hotspot")
   strHotspotID = InputBox("Please enter the id of the hotspot you going to place.", "Place Hotspot")
   lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   ' Create and place a hotspot
   HotSpots.PlaceBoxHotspot lLayerID, lObjectID, strHotspotID, "#PlaceBoxHotspot worked", dX1, dY1, dX2, dY2, dThickness, _
      iThicknessUnits, lLineColor, lFillColor, iLineStyle, iJointStyle, iFillStyle
   
   ' De-initialize the object variable
   Set HotSpots = Nothing
End Sub

Public Sub IHotspots_PlaceCircleHotspot()
   ' RobY Nov23/98
   Dim HotSpots As IHotspots
   Dim lLayerID As Long
   Dim iLayerNum As Integer
   Dim lObjectID As Long
   Dim strHotspotID As String
   Dim dX As Double
   Dim dY As Double
   Dim dRadius As Double
   Dim dThickness As Double
   Dim iThicknessUnits As UNIT_TYPE
   Dim lLineColor As Long
   Dim lFillColor As Long
   Dim iLineStyle As LINE_STYLE
   Dim iFillStyle As FILL_STYLE
   
   dX = 2.3
   dY = 1.2
   dRadius = 3
   dThickness = 0.005
   iThicknessUnits = IN_UNITS_INCH
   lLineColor = 255
   lFillColor = 255
   iLineStyle = IN_LINE_SOLID
   iFillStyle = IN_FILL_TRANSPARENT
   
   ' Set object variable for IHotSpots interface to Markup ctrl object
   Set HotSpots = MainMDIForm.ActiveForm.SpicerMarkup1.object
    
   iLayerNum = InputBox("Please enter the layer number of where to place the hotspot.", "Place Hotspot")
   lObjectID = InputBox("Please enter the vector object id of the hotspot you going to place.", "Place Hotspot")
   strHotspotID = InputBox("Please enter the id of the hotspot you going to place.", "Place Hotspot")
   lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   ' Create and place a hotspot
   HotSpots.PlaceCircleHotspot lLayerID, lObjectID, strHotspotID, "#PlaceCircleHotspot worked", dX, dY, dRadius, dThickness, _
      iThicknessUnits, lLineColor, lFillColor, iLineStyle, iFillStyle
   
   ' De-initialize the object variable
   Set HotSpots = Nothing
End Sub

Public Sub IHotspots_PlaceEllipseHotspot()
   ' RobY Nov23/98
   Dim HotSpots As IHotspots
   Dim lLayerID As Long
   Dim iLayerNum As Integer
   Dim lObjectID As Long
   Dim strHotspotID As String
   Dim dX As Double
   Dim dY As Double
   Dim dXRadius As Double
   Dim dYRadius As Double
   Dim dThickness As Double
   Dim iThicknessUnits As UNIT_TYPE
   Dim lLineColor As Long
   Dim lFillColor As Long
   Dim iLineStyle As LINE_STYLE
   Dim iFillStyle As FILL_STYLE
   Dim dRotation As Double
   
   dX = 2.3
   dY = 1.2
   dXRadius = 5
   dYRadius = 3
   dThickness = 0.005
   iThicknessUnits = IN_UNITS_INCH
   lLineColor = 255
   lFillColor = 255
   iLineStyle = IN_LINE_SOLID
   iFillStyle = IN_FILL_TRANSPARENT
   dRotation = 45
   
   ' Set object variable for IHotSpots interface to Markup ctrl object
   Set HotSpots = MainMDIForm.ActiveForm.SpicerMarkup1.object
    
   iLayerNum = InputBox("Please enter the layer number of where to place the hotspot.", "Place Hotspot")
   lObjectID = InputBox("Please enter the vector object id of the hotspot you going to place.", "Place Hotspot")
   strHotspotID = InputBox("Please enter the id of the hotspot you going to place.", "Place Hotspot")
   lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   ' Create and place a hotspot
   HotSpots.PlaceEllipseHotspot lLayerID, lObjectID, strHotspotID, "#PlaceEllipseHotspot worked", dX, dY, dXRadius, dYRadius, dThickness, _
      iThicknessUnits, lLineColor, lFillColor, iLineStyle, iFillStyle, dRotation
   
   ' De-initialize the object variable
   Set HotSpots = Nothing
End Sub

Public Sub IHotspots_PlaceIconHotspot()
   ' RobY Nov23/98
   Dim HotSpots As IHotspots
   Dim lLayerID As Long
   Dim iLayerNum As Integer
   Dim lObjectID As Long
   Dim strHotspotID As String
   Dim dX As Double
   Dim dY As Double
   
   dX = 2.3
   dY = 1.2

   ' Set object variable for IHotSpots interface to Markup ctrl object
   Set HotSpots = MainMDIForm.ActiveForm.SpicerMarkup1.object
    
   iLayerNum = InputBox("Please enter the layer number of where to place the hotspot.", "Place Hotspot")
   lObjectID = InputBox("Please enter the vector object id of the hotspot you going to place.", "Place Hotspot")
   strHotspotID = InputBox("Please enter the id of the hotspot you going to place.", "Place Hotspot")
   lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   ' Create and place a hotspot
   HotSpots.PlaceIconHotspot lLayerID, lObjectID, strHotspotID, "#PlaceIconHotspot worked!", dX, dY, 1
   
   ' De-initialize the object variable
   Set HotSpots = Nothing
End Sub

Public Sub IHotspots_PlaceLineHotspot()
   ' RobY Nov23/98
   Dim HotSpots As IHotspots
   Dim lLayerID As Long
   Dim iLayerNum As Integer
   Dim lObjectID As Long
   Dim strHotspotID As String
   Dim dX1 As Double
   Dim dY1 As Double
   Dim dX2 As Double
   Dim dY2 As Double
   Dim dThickness As Double
   Dim iThicknessUnits As UNIT_TYPE
   Dim lLineColor As Long
   Dim lFillColor As Long
   Dim iLineStyle As LINE_STYLE
   Dim iJointStyle As LINE_JOIN_STYLE
   Dim iLineFillStyle As LINE_FILL_OPERATION
   Dim iCapStyle As LINE_CAP_STYLE
   
   dX1 = 2.3
   dY1 = 1.2
   dX2 = 2.6
   dY2 = 1.3
   dThickness = 0.005
   iThicknessUnits = IN_UNITS_INCH
   lLineColor = 255
   lFillColor = 255
   iLineStyle = IN_LINE_SOLID
   iJointStyle = IN_JOIN_ROUND
   iLineFillStyle = IN_ROP_OPAQUE
   iCapStyle = IN_CAP_ROUND

   ' Set object variable for IHotSpots interface to Markup ctrl object
   Set HotSpots = MainMDIForm.ActiveForm.SpicerMarkup1.object
    
   iLayerNum = InputBox("Please enter the layer number of where to place the hotspot.", "Place Hotspot")
   lObjectID = InputBox("Please enter the vector object id of the hotspot you going to place.", "Place Hotspot")
   strHotspotID = InputBox("Please enter the id of the hotspot you going to place.", "Place Hotspot")
   lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   ' Create and place a hotspot
   HotSpots.PlaceLineHotspot lLayerID, lObjectID, strHotspotID, "#PlaceLineHotspot worked", dX1, dY1, dX2, dY2, dThickness, _
      iThicknessUnits, lLineColor, iLineStyle, iLineFillStyle, iCapStyle
   
   ' De-initialize the object variable
   Set HotSpots = Nothing
End Sub

Public Sub IHotspots_SetHotspotData()
   ' RobY Aug14/98
   Dim HotSpots As IHotspots
   Dim lLayerID As Long
   Dim iLayerNum As Integer
   Dim strHotspotID As String
   Dim strNewData As String
   
   ' Set object variable for IHotSpots interface to Markup ctrl object
   Set HotSpots = MainMDIForm.ActiveForm.SpicerMarkup1.object
    
   iLayerNum = InputBox("Please enter the layer number of the hotspot to change.", "Set Hotspot Data")
   strHotspotID = InputBox("Please enter the id of the hotspot to change.", "Set Hotspot Data")
   lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   strNewData = "SetHotspotData worked"
   ' Change the hotspot data
   HotSpots.SetHotspotData lLayerID, strHotspotID, strNewData
   
   ' De-initialize the object variable
   Set HotSpots = Nothing
End Sub

Public Sub IHotspots_SetHotspotIcon()
   ' RobY Aug14/98
   Dim HotSpots As IHotspots
   Dim lLayerID As Long
   Dim iLayerNum As Integer
   Dim strHotspotID As String
   Dim iNewIcon As Integer
   
   ' Set object variable for IHotSpots interface to Markup ctrl object
   Set HotSpots = MainMDIForm.ActiveForm.SpicerMarkup1.object
    
   iLayerNum = InputBox("Please enter the layer number of the hotspot to change.", "Set Hotspot Icon")
   strHotspotID = InputBox("Please enter the id of the hotspot you going to change.", "Set Hotspot Icon")
   lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   iNewIcon = 46
   ' Change the hotspot icon
   HotSpots.SetHotspotIcon lLayerID, strHotspotID, iNewIcon
   
   ' De-initialize the object variable
   Set HotSpots = Nothing
End Sub

Public Sub IHotspots_SetHotspotID()
   ' RobY Aug14/98
   Dim HotSpots As IHotspots
   Dim lLayerID As Long
   Dim iLayerNum As Integer
   Dim strHotspotID As String
   Dim strNewHotspotID As String
 
   ' Set object variable for IHotSpots interface to Markup ctrl object
   Set HotSpots = MainMDIForm.ActiveForm.SpicerMarkup1.object
    
   iLayerNum = InputBox("Please enter the layer number of the hotspot to change.", "Set Hotspot ID")
   strHotspotID = InputBox("Please enter the id of the hotspot to change.", "Set Hotspot ID")
   lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   strNewHotspotID = "SetHotspotID worked!"
   ' Change the hotspot ID
   HotSpots.SetHotspotID lLayerID, strHotspotID, strNewHotspotID
   
   ' De-initialize the object variable
   Set HotSpots = Nothing
End Sub

Public Sub IHotspots_SetHotspotResourceID()
   ' RobY Aug14/98
   Dim HotSpots As IHotspots
   Dim iHotspotNum As Integer
   Dim iNewHotspotNum As Integer
   
   ' Set object variable for IHotSpots interface to Markup ctrl object
   Set HotSpots = MainMDIForm.ActiveForm.SpicerMarkup1.object
    
   iHotspotNum = 1
   iNewHotspotNum = 46
   ' Change the hotspot resourceID
   HotSpots.SetHotspotResourceID iHotspotNum, iNewHotspotNum
   
   ' De-initialize the object variable
   Set HotSpots = Nothing
End Sub

Public Sub IHotspots_AttachHotspotDialogAvailability()
   'RobY Dec11/98
   Dim HotSpots As IHotspots
   Dim iAvailable As COMMAND_AVAILABILITY
   
   ' Set object variable for IHotSpots interface to Markup ctrl object
   Set HotSpots = MainMDIForm.ActiveForm.SpicerMarkup1.object
   ' Set the menu name to change
   Set MenuName = MainMDIForm.ActiveForm.mnu_Markup_IHotspots_Array(0)
   
   ' Get the value of availability
   iAvailable = HotSpots.AttachHotspotDialogAvailability
   ' Change the menu status through function
   Availability iAvailable, MenuName
   
   ' De-initialize object var
   Set HotSpots = Nothing
End Sub

Public Sub IHotspots_ChangeHotspotDialogAvailability()
   'RobY Dec11/98
   Dim HotSpots As IHotspots
   Dim iAvailable As COMMAND_AVAILABILITY
   
   ' Set object variable for IHotSpots interface to Markup ctrl object
   Set HotSpots = MainMDIForm.ActiveForm.SpicerMarkup1.object
   ' Set the menu name to change
   Set MenuName = MainMDIForm.ActiveForm.mnu_Markup_IHotspots_Array(3)
   
   ' Get the value of availability
   iAvailable = HotSpots.ChangeHotspotDialogAvailability
   ' Change the menu status through function
   Availability iAvailable, MenuName
   
   ' De-initialize object var
   Set HotSpots = Nothing
End Sub

Public Sub IHotspots_GetHotspotData()
   ' RobY Mar11/98
   Dim HotSpots As IHotspots
   Dim lLayerID As Long
   Dim iLayerNum As Integer
   Dim strHotspotID As String
   Dim strNewData As String
   
   ' Set object variable for IHotSpots interface to Markup ctrl object
   Set HotSpots = MainMDIForm.ActiveForm.SpicerMarkup1.object
    
   iLayerNum = InputBox("Please enter the layer number of the hotspot to get the data of.", "Get Hotspot Data")
   strHotspotID = InputBox("Please enter the id of the hotspot.", "Get Hotspot Data")
   lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   ' Get the hotspot data
   HotSpots.GetHotspotData lLayerID, strHotspotID, strNewData
   ' Display data of that hotspot
   MsgBox "The hotspot data of the hotspot specified is:" + Chr(13) + _
         strNewData, vbInformation, "Hotspot Data"
   
   ' De-initialize the object variable
   Set HotSpots = Nothing
End Sub





