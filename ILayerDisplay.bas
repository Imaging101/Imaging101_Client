Attribute VB_Name = "mod_ILayerDisplay"
' File:      ILayerDisplay.bas
' Created:   1998Apr28 by Rob Wood, Mark Simpson, and Rob Young
' Purpose:   To test the Spicer View Control's ILayerDisplay interface.
' Revisions: RobY Jan4/99  - Added functionality to open more than 1 file. ChildForm1 is no longer used access objects because it has been set to ChildForm() which is an array of ChildForm1's forms.
'                          - Anywhere ChildForm1 was used has been changed to ChildForm(iIndex) or MainMDIForm.ActiveForm.
'   $log$
'   $nokeywords$
'
' Copyright (C) 1998 SPICER Corporation, All rights reserved.

Option Explicit
Dim MenuName As Menu

Public Sub ILayerDisplay_Color()
   ' RobY Sept18/98
   ' Modified by RobY Dec30/98 - Allow user to specify which layer and to what colour to change to
   Dim LayerDisplay As ILayerDisplay
   Dim LayerID As Long
   Dim lRGB As Long
   Dim lPageID As Long
   Dim iLayerNum As Integer
   Dim iNumLayers As Integer
   Dim iTemp As Integer
   Dim sTemp As String
   Dim iResponse As Integer
   
   Set LayerDisplay = MainMDIForm.ActiveForm.SpicerView1.object
   lPageID = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
   iNumLayers = MainMDIForm.ActiveForm.SpicerDoc1.NumberOfLayers(lPageID)
   sTemp = "The current page has" + Str(iNumLayers) + " layer(s)." + Chr(13)
   ' Get the RGB value for all layers on the active page
   For iTemp = 1 To iNumLayers
      ' Get the layer id
      LayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(lPageID, iTemp)
      sTemp = sTemp + "Layer" + Str(iTemp) + " RGB Colour: " + Str(LayerDisplay.Color(LayerID)) + Chr(13)
   Next iTemp
   sTemp = sTemp + Chr(13) + "Do you want to change the colour of one of these layers?"
   ' Display number of layers and for each layer display its colour
   iResponse = MsgBox(sTemp, vbInformation + vbYesNo, "24-Bit Colour Value")
   If iResponse = vbYes Then
      ' Get the layer number
      iLayerNum = InputBox("Please enter the layer number of the layer that you want to change.", "Layer Number")
      ' Get the layer id
      LayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(lPageID, iLayerNum)
      ' Prompt for new layer colour
      lRGB = InputBox("Please enter new 24-bit RGB colour value", "24-Bit Colour Value")
      ' Change layer to new colour
      LayerDisplay.Color(LayerID) = lRGB
      MsgBox "Layer" + Str(iLayerNum) + " colour has been changed to" + Str(lRGB) + ".", vbInformation, "New 24-Bit Colour Value"
   End If
   
   'De-initialize the object var
   Set LayerDisplay = Nothing
End Sub

Public Sub ILayerDisplay_LayerColorDialog()
   ' RobY May8/98
   ' Define the instance for the interface
   Dim LayerDisplay As ILayerDisplay
   
   Set LayerDisplay = MainMDIForm.ActiveForm.SpicerView1.object
   ' Display the dialog box
   LayerDisplay.LayerColorDialog
   
   'De-initialize the object var
   Set LayerDisplay = Nothing
End Sub

Public Sub ILayerDisplay_LayerColorDialogAvailability()
   'RobY July1/98
   ' Modified by RobY Dec11/98 - changed global function to change menu status of ImageX according to what is returned from availablity
   Dim LayerDisplay As ILayerDisplay
   Dim iAvailable As COMMAND_AVAILABILITY
   
   ' Set object variable for ILayerDisplay interface to doc ctrl object
   Set LayerDisplay = MainMDIForm.ActiveForm.SpicerView1.object
   ' Set the menu name to change
   Set MenuName = MainMDIForm.ActiveForm.mnu_ViewCtrl_ILayerDisplay_Array(1)
   
   ' Get the value of availability
   iAvailable = LayerDisplay.LayerColorDialogAvailability
   ' Change the menu status through function
   Availability iAvailable, MenuName
   
   ' De-initialize object var
   Set LayerDisplay = Nothing
End Sub

Public Sub ILayerDisplay_LayerCompareDialog()
   ' RobY May8/98
   ' Define the instance for the interface
   Dim LayerDisplay As ILayerDisplay
   
   Set LayerDisplay = MainMDIForm.ActiveForm.SpicerView1.object
   ' Display the dialog box
   LayerDisplay.LayerCompareDialog
   
   'De-initialize the object var
   Set LayerDisplay = Nothing
End Sub

Public Sub ILayerDisplay_LayerCompareDialogAvailability()
   'RobY July1/98
   ' Modified by RobY Dec11/98 - changed global function to change menu status of ImageX according to what is returned from availablity
   Dim LayerDisplay As ILayerDisplay
   Dim iAvailable As COMMAND_AVAILABILITY
   
   ' Set object variable for ILayerDisplay interface to doc ctrl object
   Set LayerDisplay = MainMDIForm.ActiveForm.SpicerView1.object
   ' Set the menu name to change
   Set MenuName = MainMDIForm.ActiveForm.mnu_ViewCtrl_ILayerDisplay_Array(2)
   
   ' Get the value of availability
   iAvailable = LayerDisplay.LayerCompareDialogAvailability
   ' Change the menu status through function
   Availability iAvailable, MenuName
   
   ' De-initialize object var
   Set LayerDisplay = Nothing
End Sub

Public Sub ILayerDisplay_LayerDisplayDialog()
   ' RobY May8/98
   ' Define the instance for the interface
   Dim LayerDisplay As ILayerDisplay
   
   Set LayerDisplay = MainMDIForm.ActiveForm.SpicerView1.object
   ' Display the dialog box
   LayerDisplay.LayerDisplayDialog
   
   'De-initialize the object var
   Set LayerDisplay = Nothing

End Sub

Public Sub ILayerDisplay_LayerDisplayDialogAvailability()
   'RobY July1/98
   ' Modified by RobY Dec11/98 - changed global function to change menu status of ImageX according to what is returned from availablity
   Dim LayerDisplay As ILayerDisplay
   Dim iAvailable As COMMAND_AVAILABILITY
   
   ' Set object variable for ILayerDisplay interface to doc ctrl object
   Set LayerDisplay = MainMDIForm.ActiveForm.SpicerView1.object
   ' Set the menu name to change
   Set MenuName = MainMDIForm.ActiveForm.mnu_ViewCtrl_ILayerDisplay_Array(3)
   
   ' Get the value of availability
   iAvailable = LayerDisplay.LayerDisplayDialogAvailability
   ' Change the menu status through function
   Availability iAvailable, MenuName
   
   ' De-initialize object var
   Set LayerDisplay = Nothing
End Sub

Public Sub ILayerDisplay_Visible()
   ' RobY May8/98
   ' Modified by Roby July10/98
   Dim LayerDisplay As ILayerDisplay
   Dim LayerID As Long
   Dim lPageID As Long
   Dim Visible As Boolean
   
   ' Set object variable for ILayerDisplay interface to the view control
   Set LayerDisplay = MainMDIForm.ActiveForm.SpicerView1.object
   ' Get the active page id
   lPageID = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
   ' Get the layer id for layer 1 of the active page
   LayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(lPageID, 1)
   Visible = LayerDisplay.Visible(LayerID)
   ' Display whether the layer is visible or not
   MsgBox ("Layer Visibility(page 1, layer 1): " + Str(Visible))
   
   'De-initialize the object var
   Set LayerDisplay = Nothing
End Sub

' <EOF ILayerDisplay.bas>


