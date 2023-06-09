Attribute VB_Name = "mod_IScaleScrollRotation"
' File:      IScaleScrollRotation.bas
' Created:   1998Apr28 by Rob Wood, Mark Simpson, and Rob Young
' Purpose:   To test the Spicer View Control's IScaleScrollRotation interface.
' Revisions: RobY Jan4/99  - Added functionality to open more than 1 file. ChildForm1 is no longer used access objects because it has been set to ChildForm() which is an array of ChildForm1's forms.
'                          - Anywhere ChildForm1 was used has been changed to ChildForm(iIndex) or MainMDIForm.ActiveForm.
'   $log$
'   $nokeywords$
'
' Copyright (C) 1998 SPICER Corporation, All rights reserved.

Option Explicit
Dim MenuName As Menu

Public Sub IScaleScrollRotation_ActivateZoomTool()
   ' RobY May5/98
   ' Define the instance for the interface
   Dim ScaleScrollRotation As IScaleScrollRotation
   
   Set ScaleScrollRotation = MainMDIForm.ActiveForm.SpicerView1.object
   ' Activate the zoom tool
   ScaleScrollRotation.ActivateZoomTool
   
   'De-initialize the object var
   Set ScaleScrollRotation = Nothing
End Sub

Public Sub IScaleScrollRotation_ActivateZoomToolAvailability()
   'RobY July1/98
   ' Modified by RobY Dec11/98 - changed global function to change menu status of ImageX according to what is returned from availablity
   Dim ScaleScrollRotation As IScaleScrollRotation
   Dim iAvailable As COMMAND_AVAILABILITY
   
   ' Set object variable for IScaleScrollRotation interface to doc ctrl object
   Set ScaleScrollRotation = MainMDIForm.ActiveForm.SpicerView1.object
   ' Set the menu name to change
   Set MenuName = MainMDIForm.ActiveForm.mnu_ViewCtrl_IScaleScrollRotation_Array(0)
   
   ' Get the value of availability
   iAvailable = ScaleScrollRotation.ActivateZoomToolAvailability
   ' Change the menu status through function
   Availability iAvailable, MenuName
   
   ' De-initialize object var
   Set ScaleScrollRotation = Nothing
End Sub

Public Sub IScaleScrollRotation_GetVisiblePageArea()
   ' RobY May7/98
   ' Define the instance for the interface
   Dim ScaleScrollRotation As IScaleScrollRotation
   Dim x1 As Long
   Dim x2 As Long
   Dim y1 As Long
   Dim y2 As Long
   Dim ActivePageId As Long
   
   Set ScaleScrollRotation = MainMDIForm.ActiveForm.SpicerView1.object
   ActivePageId = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
   ' Call the x1,y1, and x2,y2 coordinates
   ScaleScrollRotation.GetVisiblePageArea ActivePageId, IN_UNITS_PROPORTIONAL, x1, y1, x2, y2
   ' Display the x1,y1, and x2,y2 coordinates
   MsgBox ("Proportional Visible Page Area: " + "x1(" + Str(x1) + ") y1(" + Str(y1) + _
      ") x2(" + Str(x2) + ") y2(" + Str(y2) + ")")
   
   ' De-initialize the object var
   Set ScaleScrollRotation = Nothing
End Sub

Public Sub IScaleScrollRotation_GetZoomFactor()
   ' RobY May7/98
   ' Define the instance for the interface
   Dim ScaleScrollRotation As IScaleScrollRotation
   Dim pScale As Double
   Dim pX As Long
   Dim pY As Long
   
   Set ScaleScrollRotation = MainMDIForm.ActiveForm.SpicerView1.object
   ' Call the Scale, xCenter, and yCenter for the ZoomFactor
   ScaleScrollRotation.GetZoomFactor MainMDIForm.ActiveForm.SpicerView1.ActivePageId, pScale, pX, pY
   ' Display the Scale, xCenter, and yCenter
   MsgBox ("Zoom Factor: " + "Scale(" + Str(pScale) + "), xCenter(" + Str(pX) + _
      "), yCenter(" + Str(pY) + ")")
   
   ' De-initialize the object var
   Set ScaleScrollRotation = Nothing
End Sub

Public Sub IScaleScrollRotation_Rotation()
   ' RobY May7/98
   ' Define the instance for the interface
   Dim ScaleScrollRotation As IScaleScrollRotation
   Set ScaleScrollRotation = MainMDIForm.ActiveForm.SpicerView1.object
   
   MsgBox "Rotate counter clockwise 90 degrees.", vbInformation, "Rotation"
   ScaleScrollRotation.Rotation(MainMDIForm.ActiveForm.SpicerView1.ActivePageId) = IN_ROTATION_90
   
   MsgBox "Rotate clockwise 90 degrees from current position.", vbInformation, "Rotation"
   ScaleScrollRotation.Rotation(MainMDIForm.ActiveForm.SpicerView1.ActivePageId) = IN_ROTATION_90_CW
   
   MsgBox "Rotate counter clockwise 270 degrees from original position.", vbInformation, "Rotation"
   ScaleScrollRotation.Rotation(MainMDIForm.ActiveForm.SpicerView1.ActivePageId) = IN_ROTATION_270
   
   MsgBox "Rotate 180 degrees from original position.", vbInformation, "Rotation"
   ScaleScrollRotation.Rotation(MainMDIForm.ActiveForm.SpicerView1.ActivePageId) = IN_ROTATION_180
   
   MsgBox "Rotate 180 degrees from current position.", vbInformation, "Rotation"
   ScaleScrollRotation.Rotation(MainMDIForm.ActiveForm.SpicerView1.ActivePageId) = IN_ROTATION_180_REL
   
   MsgBox "Rotate counter clockwise 90 degrees from original.", vbInformation, "Rotation"
   ScaleScrollRotation.Rotation(MainMDIForm.ActiveForm.SpicerView1.ActivePageId) = IN_ROTATION_90_CCW
   
   MsgBox "Rotate 0 degrees which is original position.", vbInformation, "Rotation"
   ScaleScrollRotation.Rotation(MainMDIForm.ActiveForm.SpicerView1.ActivePageId) = IN_ROTATION_0
   
   ' De-initialize the object var
   Set ScaleScrollRotation = Nothing

End Sub

Public Sub IScaleScrollRotation_ScrollStepSize()
   ' RobY May6/98
   ' Define the instance for the interface
   Dim ScaleScrollRotation As IScaleScrollRotation
   
   Set ScaleScrollRotation = MainMDIForm.ActiveForm.SpicerView1.object
   
   ' Set step size for the scroll bars between 0 and 1
   ScaleScrollRotation.ScrollStepSize(MainMDIForm.ActiveForm.SpicerView1.ActivePageId) = 0.5
   MsgBox "Scroll step size has been set to 0.5.", vbInformation, "Scroll Step Size"
   
   ' De-initialize the object var
   Set ScaleScrollRotation = Nothing
End Sub

Public Sub IScaleScrollRotation_ScrollView(iIndex As Integer)
   ' RobY May6/98
   ' Modified by Roby Dec4/98 - Command selected by user through menu
   Dim ScaleScrollRotation As IScaleScrollRotation
   
   ' Set object variable for IScaleScrollRotation interface to view ctrl object
   Set ScaleScrollRotation = MainMDIForm.ActiveForm.SpicerView1.object
   
   ' Menu index number is used to determine which one is executed
   If iIndex = 0 Then
      ScaleScrollRotation.ScrollView(MainMDIForm.ActiveForm.SpicerView1.ActivePageId) = IN_SCROLL_CENTER
   ElseIf iIndex = 1 Then
      ScaleScrollRotation.ScrollView(MainMDIForm.ActiveForm.SpicerView1.ActivePageId) = IN_SCROLL_LEFTCENTER
   ElseIf iIndex = 2 Then
      ScaleScrollRotation.ScrollView(MainMDIForm.ActiveForm.SpicerView1.ActivePageId) = IN_SCROLL_RIGHTCENTER
   ElseIf iIndex = 3 Then
      ScaleScrollRotation.ScrollView(MainMDIForm.ActiveForm.SpicerView1.ActivePageId) = IN_SCROLL_TOPCENTER
   ElseIf iIndex = 4 Then
      ScaleScrollRotation.ScrollView(MainMDIForm.ActiveForm.SpicerView1.ActivePageId) = IN_SCROLL_TOPLEFT
   ElseIf iIndex = 5 Then
      ScaleScrollRotation.ScrollView(MainMDIForm.ActiveForm.SpicerView1.ActivePageId) = IN_SCROLL_TOPRIGHT
   ElseIf iIndex = 6 Then
      ScaleScrollRotation.ScrollView(MainMDIForm.ActiveForm.SpicerView1.ActivePageId) = IN_SCROLL_BOTTOMCENTER
   ElseIf iIndex = 7 Then
      ScaleScrollRotation.ScrollView(MainMDIForm.ActiveForm.SpicerView1.ActivePageId) = IN_SCROLL_BOTTOMLEFT
   ElseIf iIndex = 8 Then
      ScaleScrollRotation.ScrollView(MainMDIForm.ActiveForm.SpicerView1.ActivePageId) = IN_SCROLL_BOTTOMRIGHT
   ElseIf iIndex = 10 Then
      ScaleScrollRotation.ScrollView(MainMDIForm.ActiveForm.SpicerView1.ActivePageId) = IN_SCROLL_LEFTSTEP
   ElseIf iIndex = 11 Then
      ScaleScrollRotation.ScrollView(MainMDIForm.ActiveForm.SpicerView1.ActivePageId) = IN_SCROLL_RIGHTSTEP
   ElseIf iIndex = 12 Then
      ScaleScrollRotation.ScrollView(MainMDIForm.ActiveForm.SpicerView1.ActivePageId) = IN_SCROLL_UPSTEP
   ElseIf iIndex = 13 Then
      ScaleScrollRotation.ScrollView(MainMDIForm.ActiveForm.SpicerView1.ActivePageId) = IN_SCROLL_LEFTUPSTEP
   ElseIf iIndex = 14 Then
      ScaleScrollRotation.ScrollView(MainMDIForm.ActiveForm.SpicerView1.ActivePageId) = IN_SCROLL_RIGHTUPSTEP
   ElseIf iIndex = 15 Then
      ScaleScrollRotation.ScrollView(MainMDIForm.ActiveForm.SpicerView1.ActivePageId) = IN_SCROLL_DOWNSTEP
   ElseIf iIndex = 16 Then
      ScaleScrollRotation.ScrollView(MainMDIForm.ActiveForm.SpicerView1.ActivePageId) = IN_SCROLL_LEFTDOWNSTEP
   ElseIf iIndex = 17 Then
      ScaleScrollRotation.ScrollView(MainMDIForm.ActiveForm.SpicerView1.ActivePageId) = IN_SCROLL_RIGHTDOWNSTEP
   ElseIf iIndex = 19 Then
      ScaleScrollRotation.ScrollView(MainMDIForm.ActiveForm.SpicerView1.ActivePageId) = IN_SCROLL_LEFTPAGE
   ElseIf iIndex = 20 Then
      ScaleScrollRotation.ScrollView(MainMDIForm.ActiveForm.SpicerView1.ActivePageId) = IN_SCROLL_RIGHTPAGE
   ElseIf iIndex = 21 Then
      ScaleScrollRotation.ScrollView(MainMDIForm.ActiveForm.SpicerView1.ActivePageId) = IN_SCROLL_UPPAGE
   ElseIf iIndex = 22 Then
      ScaleScrollRotation.ScrollView(MainMDIForm.ActiveForm.SpicerView1.ActivePageId) = IN_SCROLL_LEFTUPPAGE
   ElseIf iIndex = 23 Then
      ScaleScrollRotation.ScrollView(MainMDIForm.ActiveForm.SpicerView1.ActivePageId) = IN_SCROLL_RIGHTUPPAGE
   ElseIf iIndex = 24 Then
      ScaleScrollRotation.ScrollView(MainMDIForm.ActiveForm.SpicerView1.ActivePageId) = IN_SCROLL_DOWNPAGE
   ElseIf iIndex = 25 Then
      ScaleScrollRotation.ScrollView(MainMDIForm.ActiveForm.SpicerView1.ActivePageId) = IN_SCROLL_LEFTDOWNPAGE
   ElseIf iIndex = 26 Then
      ScaleScrollRotation.ScrollView(MainMDIForm.ActiveForm.SpicerView1.ActivePageId) = IN_SCROLL_RIGHTDOWNPAGE
   End If
      
   ' De-initialize the object var
   Set ScaleScrollRotation = Nothing
End Sub

Public Sub IScaleScrollRotation_SetVisiblePageArea()
   ' RobY May7/98
   ' Modified by RobY July8/98
   Dim ScaleScrollRotation As IScaleScrollRotation
   Dim iUnits As Integer
   Dim lX1, lY1, lX2, lY2 As Long
   
   ' Set object variable for IScaleScrollRotation interface to view ctrl object
   Set ScaleScrollRotation = MainMDIForm.ActiveForm.SpicerView1.object
   iUnits = InputBox("Please enter the units you want to use." + Chr(13) + "PROPORTIONAL = 0" + _
               Chr(13) + "INCH = 1" + Chr(13) + "CM = 2" + Chr(13) + "FT = 4" + Chr(13) + "MM = 5" + _
               Chr(13) + "M = 6", "Unit Type")
   lX1 = InputBox("Please enter the value for the x1 coordinate.", "X1 Coordinate")
   lY1 = InputBox("Please enter the value for the y1 coordinate.", "Y1 Coordinate")
   lX2 = InputBox("Please enter the value for the x2 coordinate.", "X2 Coordinate")
   lY2 = InputBox("Please enter the value for the y2 coordinate.", "Y2 Coordinate")
   ScaleScrollRotation.SetVisiblePageArea MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iUnits, lX1, lY1, lX2, lY2
   
   ' De-initialize the object var
   Set ScaleScrollRotation = Nothing
End Sub

Public Sub IScaleScrollRotation_SetZoomFactor()
   ' RobY May7/98
   ' Modified by RobY July8/98
   Dim ScaleScrollRotation As IScaleScrollRotation
   Dim iZoomLevel As ZOOM_LEVEL
   Dim dScale As Double
   Dim lXValue As Long
   Dim lYValue As Long
   
   ' Set object variable for IScaleScrollRotation interface to view ctrl object
   Set ScaleScrollRotation = MainMDIForm.ActiveForm.SpicerView1.object
   iZoomLevel = InputBox("Please enter the zoom level." + Chr(13) + "CUSTOM = 9" + Chr(13) + _
                        "CUSTOM CENTER = 10", "Zoom Level")
   dScale = InputBox("Please enter the new scale value.", "Scale")
   Select Case iZoomLevel
      Case IN_ZOOM_CUSTOM
         lXValue = 0
         lYValue = 0
      Case IN_ZOOM_CUSTOM_CENTER
         lXValue = InputBox("Please enter new center point for x", "X Center")
         lYValue = InputBox("Please enter new center point for y", "Y Center")
   End Select
   ScaleScrollRotation.SetZoomFactor MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iZoomLevel, dScale, lXValue, lYValue
   
   ' De-initialize the object var
   Set ScaleScrollRotation = Nothing
End Sub

Public Sub IScaleScrollRotation_ZoomLevel()
   ' RobY May5/98
   ' Not all constants are checked and will be done at a later time
   ' Define the instance for the interface
   Dim ScaleScrollRotation As IScaleScrollRotation
   Set ScaleScrollRotation = MainMDIForm.ActiveForm.SpicerView1.object
   
   MsgBox "Click OK to zoom in 1 to 1.", vbExclamation, "ZOOM LEVEL"
   ScaleScrollRotation.ZoomLevel(MainMDIForm.ActiveForm.SpicerView1.ActivePageId) = IN_ZOOM_1TO1
   
   MsgBox "Click OK to scale to fit.", vbExclamation, "ZOOM LEVEL"
   ScaleScrollRotation.ZoomLevel(MainMDIForm.ActiveForm.SpicerView1.ActivePageId) = IN_ZOOM_SCALETOFIT
   
   MsgBox "Click OK to zoom in.", vbExclamation, "ZOOM LEVEL"
   ScaleScrollRotation.ZoomLevel(MainMDIForm.ActiveForm.SpicerView1.ActivePageId) = IN_ZOOM_IN
   
   MsgBox "Click OK to zoom out.", vbExclamation, "ZOOM LEVEL"
   ScaleScrollRotation.ZoomLevel(MainMDIForm.ActiveForm.SpicerView1.ActivePageId) = IN_ZOOM_OUT
      
   MsgBox "Click OK to scale to fit vertically.", vbExclamation, "ZOOM LEVEL"
   ScaleScrollRotation.ZoomLevel(MainMDIForm.ActiveForm.SpicerView1.ActivePageId) = IN_ZOOM_VERTFIT
   
   MsgBox "Click OK to scale to fit horizontally.", vbExclamation, "ZOOM LEVEL"
   ScaleScrollRotation.ZoomLevel(MainMDIForm.ActiveForm.SpicerView1.ActivePageId) = IN_ZOOM_HORIZFIT
   
   MsgBox "Click OK to scale to zoom to actual size.", vbExclamation, "ZOOM LEVEL"
   ScaleScrollRotation.ZoomLevel(MainMDIForm.ActiveForm.SpicerView1.ActivePageId) = IN_ZOOM_ACTUALSIZE

   ' De-initialize the object var
   Set ScaleScrollRotation = Nothing
End Sub

Public Sub IScaleScrollRotation_ZoomStepSize()
   ' RobY May5/98
   ' Modified by RobY July9/98
   Dim ScaleScrollRotation As IScaleScrollRotation
   Dim sFactor As Single
   
   ' Set object variable for IScaleScrollRotation interface to view ctrl object
   Set ScaleScrollRotation = MainMDIForm.ActiveForm.SpicerView1.object
   sFactor = InputBox("Please enter a multiplication factor between 1.0 and 3.0.", "Multiplication Factor")
   'Set ZoomStepSize to a multipication factor between 1.0 to 3.0
   ScaleScrollRotation.ZoomStepSize(MainMDIForm.ActiveForm.SpicerView1.ActivePageId) = sFactor
   MsgBox "Zoom step size has been set to" + Str(sFactor), vbInformation, "Zoom Step Size"
   
   ' De-initialize the object var
   Set ScaleScrollRotation = Nothing
End Sub

' <EOF IScaleScrollRotation.bas>

