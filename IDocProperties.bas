Attribute VB_Name = "mod_IDocProperties"
' File:      IDocProperties.bas
' Created:   1998Apr28 by Rob Wood, Mark Simpson, and Rob Young
' Purpose:   To test the Spicer Doc Control's IDocProperties interface.
' Revisions: RobY Dec17/98 - Commented out code for FitLayerToWindow. Removed from controls. Possibility of it being added back later.
'            RobY Jan4/99  - Added functionality to open more than 1 file. ChildForm1 is no longer used access objects because it has been set to ChildForm() which is an array of ChildForm1's forms.
'                          - Anywhere ChildForm1 was used has been changed to ChildForm(iIndex) or MainMDIForm.ActiveForm.
'   $log$
'   $nokeywords$
'
' Copyright (C) 1998 SPICER Corporation, All rights reserved.

Option Explicit
Dim MenuName As Menu

Public Sub IDocProperties_DeleteUserAttribute()
   'Mark Simpson May 12/98
   Dim lSetUserID As String
   Dim lNewAttributeString As String
   Dim lGetAttributeString As String
   Dim docProperties As IDocProperties
    
   'Set object variable for IDocProperties interface to doc ctrl object
   Set docProperties = MainMDIForm.ActiveForm.SpicerDoc1.object
   
   'Set UserID attribute to something
   lSetUserID = InputBox("Please enter what will be placed into the UserID string:")
   docProperties.UserAttribute(MainMDIForm.ActiveForm.SpicerDoc1.object, "UserID") = lSetUserID
   
   'Delete the UserID attribute
   docProperties.DeleteUserAttribute MainMDIForm.ActiveForm.SpicerDoc1.object, "UserID"
   
   'Attempt to obtain UserID value
   lGetAttributeString = docProperties.UserAttribute(MainMDIForm.ActiveForm.SpicerDoc1.object, "UserID")
   If lGetAttributeString = "" Then
      'Success
      MsgBox "Attempt to obtain UserID value after its deletion resulted in blank, therefore attribute was deleted.", vbInformation
   Else
      'Failure
       MsgBox "Attempt to obtain UserID value after its deletion returned:" + Chr(13) + lGetAttributeString + Chr(13) + "Therefore attribute was not deleted.", vbCritical
   End If
    
   'De-initialize the object variable
   Set docProperties = Nothing
End Sub

Public Sub IDocProperties_DocumentPropertiesDialog()
   'Mark Simpson May 8/98
   Dim docProperties As IDocProperties
   
   'Set object variable for IDocProperties interface to doc ctrl object
   Set docProperties = MainMDIForm.ActiveForm.SpicerDoc1.object
   
   'Open Document Properties Dialog for document
   docProperties.DocumentPropertiesDialog
   
   'De-initialize the object variable
   Set docProperties = Nothing
End Sub

Public Sub IDocProperties_DocumentPropertiesDialogAvailability()
   'RobY July1/98
   'Modified by RobY Dec4/98 - changed global function to change menu status of ImageX according to what is returned from availablity
   Dim docProperties As IDocProperties
   Dim iAvailable As COMMAND_AVAILABILITY
   
   ' Set object variable for IDocProperties interface to doc ctrl object
   Set docProperties = MainMDIForm.ActiveForm.SpicerDoc1.object
   ' Set the menu name to change
   Set MenuName = MainMDIForm.ActiveForm.mnu_DocCtrl_IDocProperties_Array(1)
   
   ' Get the value of availability
   iAvailable = docProperties.DocumentPropertiesDialogAvailability
   ' Change the menu status through function
   Availability iAvailable, MenuName
   
   ' De-initialize object var
   Set docProperties = Nothing
End Sub

Public Sub IDocProperties_Filename()
   'Mark Simpson May8/98
   Dim lFilename As String
   Dim docProperties As IDocProperties
   
   'Set object variable for IDocProperties interface to doc ctrl object
   Set docProperties = MainMDIForm.ActiveForm.SpicerDoc1.object
   
   'Get the Filename of the active document
   lFilename = docProperties.FileName(0)
   
   'Display Filename in message box
   MsgBox "This file is located in:" + Chr(13) + lFilename, vbInformation
   
   'De-initialize the object variable
   Set docProperties = Nothing
End Sub

'Public Sub IDocProperties_FitLayerToWindow()
   ' Mark Simpson June 8/98
   ' Modified by RobY July8/98
   ' Modified by MarkS Nov 17, 1998 Now asks which layer is going to be adjusted, as well as asking for the scale mode and center mode
'   Dim docProperties As IDocProperties
'   Dim lWhichLayer As Integer
'   Dim lNumLayers As Integer
'   Dim LayerID As Long
'   Dim lPageID As Long
'   Dim lScaleMode As Long
'   Dim lCenterMode As Long
         
   ' if there is only one layer don't bother performing this
   ' it is meant to have at least two layers, one some text or a logo that stays fixed in an area as the other image is zoomed in on
'   Dim docContents As IDocContents
   ' Get the page ID for current page of active document
'   lPageID = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
   ' Set object variable for IDocContents interface to doc ctrl object
'   Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
   ' Get the number of layers for the current page
'   lNumLayers = docContents.NumberOfLayers(lPageID)
'   If lNumLayers > 1 Then
      ' De-initialize the object variable so we can reset it to the IDocProperties interface
'      Set docContents = Nothing
   
'      lPageID = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
      
      'Ask which layer they want to adjust
'      lWhichLayer = InputBox("Please enter which layer you want to adjust:")
      
      'Ask what scale mode is to be used
'      lScaleMode = InputBox("Please enter the scale mode:" + Chr(13) + "IN_ZOOM_SCALETOFIT = 3" + Chr(13) + "IN_ZOOM_HORIZFIT = 4" + Chr(13) + "IN_ZOOM_VERTFIT = 5" + Chr(13) + "IN_ZOOM_ACTUALSIZE = 6" + Chr(13) + "IN_ZOOM_1TO1 = 8")
      
      'Ask what center mode is to be used
'      lCenterMode = InputBox("Please enter the center mode:" + Chr(13) + "IN_SCROLL_CENTER = 1" + Chr(13) + "IN_SCROLL_BOTTOMLEFT = 2" + Chr(13) + "IN_SCROLL_BOTTOMRIGHT = 3" + Chr(13) + "IN_SCROLL_TOPLEFT = 4" + Chr(13) + "IN_SCROLL_TOPRIGHT = 5" + Chr(13) + "IN_SCROLL_TOPCENTER = 22" + Chr(13) + "IN_SCROLL_RIGHTCENTER = 23" + Chr(13) + "IN_SCROLL_LEFTCENTER = 24" + Chr(13) + "IN_SCROLL_BOTTOMCENTER = 25")
      
      'Get Layer ID
'      LayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(lPageID, lWhichLayer)
      
      'Set object variable for IDocProperties interface to doc ctrl object
'      Set docProperties = MainMDIForm.ActiveForm.SpicerDoc1.object
      
      ' Fit page 1 layer 1 to the window
      'docProperties.FitLayerToWindow LayerID, IN_ZOOM_SCALETOFIT, IN_SCROLL_CENTER
'      docProperties.FitLayerToWindow LayerID, lScaleMode, lCenterMode
'   Else
'      MsgBox "Need more than one layer to properly test this." + Chr(13) + "One layer will be frozen to a size in the window while the other layer(s) can be zoomed in on." + Chr(13) + "Add another layer first."
'   End If
      
'   ' De-initialize the object var
'   Set docProperties = Nothing
'End Sub

Public Sub IDocProperties_Format()
   'Mark Simpson May8/98
   Dim lFormat As String
   Dim docProperties As IDocProperties
   
   'Set object variable for IDocProperties interface to doc ctrl object
   Set docProperties = MainMDIForm.ActiveForm.SpicerDoc1.object
   
   'Get the Format of the active document
   lFormat = docProperties.Format(0)
   
   'Display Format in message box
   MsgBox "The format ID of this file is:" + Chr(13) + lFormat, vbInformation
   ' Note that lFormat is currently only a format id number -- need to see if I can get the text for it
   
   'De-initialize the object variable
   Set docProperties = Nothing
End Sub
Public Sub IDocProperties_GetPageCalibration()
   'Mark Simpson May 19/98
   'Modified by RobY June22/98
   Dim iPageNum As Integer
   Dim scaleFactor As Double
   Dim units As UNIT_TYPE
   Dim bSet As Boolean
   Dim docProperties As IDocProperties
   Dim lPageID As Long
   
   'Set object variable for IDocProperties interface to doc ctrl object
   Set docProperties = MainMDIForm.ActiveForm.SpicerDoc1.object
      
   'Get the first page's ID
   iPageNum = InputBox("Obtain the calibrations for which page?")
   lPageID = MainMDIForm.ActiveForm.SpicerDoc1.pageID(iPageNum)
   'Get the calibration settings for the requested page
   docProperties.GetPageCalibration lPageID, scaleFactor, units, bSet
   
   'Display the result
   If bSet = False Then
      MsgBox "There are no calibrations set for this page."
   Else
      MsgBox "The calibrations for this page are:" + Chr(13) + "Scale Factor =" + Str(scaleFactor) + _
         " Unit Type =" + Str(units)
   End If
   
   'De-initialize the object variable
   Set docProperties = Nothing
End Sub

Public Sub IDocProperties_GetObjectExtents()
   'Roby Nov27/98
   Dim docProperties As IDocProperties
   Dim lObjectID As Long
   Dim dReturnX1 As Double
   Dim dReturnY1 As Double
   Dim dReturnX2 As Double
   Dim dReturnY2 As Double
   Dim iUnitType As Integer
   
   'Set object variable for IDocProperties interface to doc ctrl object
   Set docProperties = MainMDIForm.ActiveForm.SpicerDoc1.object
      
   ' Prompt for object id
   lObjectID = InputBox("Please enter the page or layer id to get the extents of.", "Object ID")
   ' Prompt for unit type
   iUnitType = InputBox("Please enter the unit type you want the extents returned in." + Chr(13) + _
                        Chr(13) + "INCH = 1" + Chr(13) + "CM = 2" + Chr(13) + "FT = 4" + Chr(13) + "MM = 5" + _
                        Chr(13) + "M = 6", "Unit Type")
   ' Get the extents
   docProperties.GetObjectExtents lObjectID, iUnitType, dReturnX1, dReturnY1, dReturnX2, dReturnY2
   MsgBox "Object Extents:" + Chr(13) + "x1=" + Str(dReturnX1) + Chr(13) + "y1=" + Str(dReturnY1) + Chr(13) + _
            "x2=" + Str(dReturnX2) + Chr(13) + "y2=" + Str(dReturnY2), vbInformation, "GetObjectExtents"
   
   'De-initialize the object variable
   Set docProperties = Nothing
End Sub

Public Sub IDocProperties_HeaderRotation()
   'RobY Nov26/98
   Dim docProperties As IDocProperties
   Dim iCurHeaderRot As ROTATION_ANGLE
   Dim iNewHeaderRot As ROTATION_ANGLE
   Dim iGetNewHeaderRot As ROTATION_ANGLE
   Dim strHeaderRot As String
   Dim lLayerID As String
   Dim iResponse As Integer
   Dim iLayerNum As Integer
   
   'Set object variable for IDocProperties interface to doc ctrl object
   Set docProperties = MainMDIForm.ActiveForm.SpicerDoc1.object
   
   ' Prompt for the raster layer number
   iLayerNum = InputBox("Please enter the layer number of the raster.", "Layer Number")
   ' Convert layer number to layer id
   lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   ' Get the current header rotation
   iCurHeaderRot = docProperties.HeaderRotation(lLayerID)
   ' Use function to return rotation string
   strHeaderRot = Rotation(iCurHeaderRot)
   ' Display result and prompt user for next action
   iResponse = MsgBox("The current Header Rotation for the specified raster layer is " + strHeaderRot + "." + _
                     Chr(13) + "Do you want to change the current value?", vbYesNo, "Header Rotation")
   If iResponse = vbYes Then
      ' Prompt for new rotation
      iNewHeaderRot = InputBox("Please enter a new header rotation." + Chr(13) + "IN_ROTATION_0 = 0" + Chr(13) + _
                                 "IN_ROTATION_90 = 1" + Chr(13) + "IN_ROTATION_180 = 2" + Chr(13) + "IN_ROTATION_270 = 3", "New Header Rotation")
      ' Set to new header rotation
      docProperties.HeaderRotation(lLayerID) = iNewHeaderRot
      ' Get new header rotation
      iGetNewHeaderRot = docProperties.HeaderRotation(lLayerID)
      If iNewHeaderRot <> iGetNewHeaderRot Then
         MsgBox "Failed to change header rotation to new setting!", vbCritical, "Failed"
      Else
         MsgBox "Header rotation has been changed to new setting.", vbInformation, "Success"
      End If
   Else
      MsgBox "The header rotation was not changed.", vbInformation, "Header Rotation"
   End If
    
   'De-initialize the object variable
   Set docProperties = Nothing
End Sub

Public Sub IDocProperties_Label()
   'Mark Simpson May8/98
   'modified by MarkS Nov18/98 result MsgBox incorrectly said active window, should have been active document
   'Modified by RobY Nov30/98    Changed objectid parameter from 0 to root id
   Dim lLabel As String
   Dim docProperties As IDocProperties
   
   'Set object variable for IDocProperties interface to doc ctrl object
   Set docProperties = MainMDIForm.ActiveForm.SpicerDoc1.object
   
   'Get the label of the active document
   lLabel = docProperties.Label(MainMDIForm.ActiveForm.SpicerDoc1.RootID)
   
   'Display label in message box
   MsgBox "The label of the active document is:" + Chr(13) + lLabel, vbInformation
   
   'De-initialize the object variable
   Set docProperties = Nothing
End Sub

Public Sub IDocProperties_LayerType()
   ' RobY Sept21/98
   Dim docProperties As IDocProperties
   Dim docContents As IDocContents
   Dim iLayerType As LAYER_TYPES
   Dim iNumLayers As Integer
   Dim lLayerID As Long
   Dim iCounter As Integer
   Dim sTemp As String
   Dim lPageID As Long
   Dim strLayerType As String
   
   ' Set object variable for IDocProperties interface to doc ctrl object
   Set docProperties = MainMDIForm.ActiveForm.SpicerDoc1.object
   ' Set object variable for IDocContents interface to doc ctrl object
   Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
   
   ' Get the page id of the active page
   lPageID = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
   ' Get the number of layers
   iNumLayers = docContents.NumberOfLayers(lPageID)
   ' Get the layer type for all the layers on the current page
   For iCounter = 1 To iNumLayers
      lLayerID = docContents.LayerID(lPageID, iCounter)
      iLayerType = docProperties.LayerType(lLayerID)
      Select Case iLayerType
         Case 0
            strLayerType = "Unknown"
         Case 2
            strLayerType = "Edit"
         Case 6
            strLayerType = "Raster"
      End Select
      sTemp = sTemp + "Layer" + Str(iCounter) + " Type is " + strLayerType + "." + Chr(13)
   Next iCounter
   ' Display the layer types for all layers on the active page
   MsgBox sTemp, vbInformation, "Layer Type"
   
   'De-initialize the object variable
   Set docProperties = Nothing
End Sub

Public Sub IDocProperties_Permissions()
   'Mark Simpson May 12/98
   Dim lPageID As Long
   Dim lOrigPerms As Long
   Dim lSetPerms As Long
   Dim lGetPerms As Long
   Dim lTempPerms As Long
   Dim lOrigPermsString As String
   Dim lSetPermsString As String
   Dim lGetPermsString As String
   Dim docProperties As IDocProperties
   
   'Get the pageID of page 1
   lPageID = MainMDIForm.ActiveForm.SpicerDoc1.pageID(1)
   
   'Set object variable for IDocProperties interface to doc ctrl object
   Set docProperties = MainMDIForm.ActiveForm.SpicerDoc1.object
   
   'Get original permissions of the document
   lOrigPerms = docProperties.PERMISSIONS(lPageID) 'MainMDIForm.ActiveForm.SpicerDoc1.object)
   
   'Call function to build original perms string
   lTempPerms = lOrigPerms
   lOrigPermsString = BuildPermsString(lTempPerms, lOrigPermsString)
       
   'Set new permissions
   lSetPerms = InputBox("Enter permissions to be set to:" + Chr(13) + "Read  1" + Chr(13) + "Write 2" + Chr(13) + "Create  4" + Chr(13) + "Modify 8" + Chr(13) + "Delete 16" + Chr(13) + "Print 32" + Chr(13) + "Hide 64" + Chr(13) + "Static 128" + Chr(13) + "User Defined 1 256" + Chr(13) + "User Defined 2 512" + Chr(13) + "Note permissions can be OR'ED by adding them together.")
   docProperties.PERMISSIONS(lPageID) = lSetPerms
   
   'Call function to build Set perms string
   lTempPerms = lSetPerms
   lSetPermsString = BuildPermsString(lTempPerms, lSetPermsString)
       
   'Get the new permissions
   lGetPerms = docProperties.PERMISSIONS(lPageID)
   
   'Call function to build get perms string
   lTempPerms = lGetPerms
   lGetPermsString = BuildPermsString(lTempPerms, lGetPermsString)
       
   'Confirm that get and set are correct
   If lSetPerms = lGetPerms Then
      'Success
      MsgBox "Original permissions for page 1 are:" + Chr(13) + lOrigPermsString + Chr(13) + Chr(13) + "New Permissions for page 1 are:" + Chr(13) + lGetPermsString, vbInformation
   Else
      'Failure
      MsgBox "Get/Set permissions not matching." + Chr(13) + "Orig: " + lOrigPermsString + Chr(13) + "Set: " + lSetPermsString + Chr(13) + "Get:" + lGetPermsString, vbCritical
   End If
   
   'Return to original permissions
   docProperties.PERMISSIONS(lPageID) = lOrigPerms
End Sub

Public Sub IDocProperties_RasterInfoString()
   'Mark Simpson May8/98
   Dim lLayerID As Integer
   Dim lFormatID As Integer
   Dim lOrigRasterNotes As String
   Dim lNewRasterNotes As String
   Dim lGetRasterNotes As String
   Dim docProperties As IDocProperties
   
   'Get the layer ID of the current active layer
   lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, 1)
   
   'Set object variable for IDocProperties interface to doc ctrl object
   Set docProperties = MainMDIForm.ActiveForm.SpicerDoc1.object
   
   'Make sure user has opened a file that can have raster notes
   lFormatID = docProperties.Format(MainMDIForm.ActiveForm.SpicerDoc1.object)
   Select Case lFormatID
      Case 1, 3, 8, 11 - 15, 22 - 26, 32, 33, 36 - 39, 49, 60 - 62, 64 - 66
         'Ask user for what raster notes will be set to
         lNewRasterNotes = InputBox("Enter what the raster notes will be:")
         
         'Get the Original Raster Notes
         lOrigRasterNotes = docProperties.RasterInfoString(lLayerID)
                
         'Set the RasterNotes
         docProperties.RasterInfoString(lLayerID) = lNewRasterNotes
         
         'Get the RasterNotes of the current active layer
         lGetRasterNotes = docProperties.RasterInfoString(lLayerID)
                  
         'Display result
         If lNewRasterNotes = lGetRasterNotes Then
            'success
            MsgBox "The new RasterNotes for this layer are:" + Chr(13) + lGetRasterNotes, vbInformation
         Else
            'failure
            MsgBox "Get/Set RasterNotes not matching." + Chr(13) + "Orig: " + lOrigRasterNotes + Chr(13) + "Set: " + lNewRasterNotes + Chr(13) + "Get:" + lGetRasterNotes, vbCritical
         End If
         
         'Return raster notes to original setting
         docProperties.RasterInfoString(lLayerID) = lOrigRasterNotes
         
      Case Else
         MsgBox "RasterNotes are not available for this format."
   End Select
    
   'De-initialize the object variable
   Set docProperties = Nothing
End Sub

Public Sub IDocProperties_ReadOnlyModsFlag()
   'Mark Simpson May11/98
   'Modified by RobY July7/98
   Dim bSetReadOnlyFlag As String
   Dim bGetReadOnlyFlag As String
   Dim docProperties As IDocProperties
   
   'Set object variable for IDocProperties interface to doc ctrl object
   Set docProperties = MainMDIForm.ActiveForm.SpicerDoc1.object
   
   'Get the original ReadOnlyModsFlag of the current active layer
   bGetReadOnlyFlag = docProperties.ReadOnlyModsFlag
   
   If (bGetReadOnlyFlag = False) Then
      docProperties.ReadOnlyModsFlag = True
      bSetReadOnlyFlag = docProperties.ReadOnlyModsFlag
      MsgBox "Originally set to: " + bGetReadOnlyFlag + Chr(13) + "Now set to:" + bSetReadOnlyFlag, vbInformation
   ElseIf (bGetReadOnlyFlag = True) Then
      docProperties.ReadOnlyModsFlag = False
      bSetReadOnlyFlag = docProperties.ReadOnlyModsFlag
      MsgBox "Originally set to: " + bGetReadOnlyFlag + Chr(13) + "Now set to:" + bSetReadOnlyFlag, vbInformation, "ReadOnlyModsFlag"
   End If
      
   'De-initialize the object variable
   Set docProperties = Nothing
End Sub

Public Sub IDocProperties_SetPageCalibration()
   'Mark Simpson May 19/98
   'RobY June22/98   Write code for this method
   Dim docProperties As IDocProperties
   Dim iPageNum As Integer
   Dim lPageID As Long
   Dim dScaleFactor As Double
   Dim iUnitType As UNIT_TYPE
   Dim strNewLine As String
   
   'Set object variable for IDocProperties interface to doc ctrl object
   Set docProperties = MainMDIForm.ActiveForm.SpicerDoc1.object
   strNewLine = Chr(13) & Chr(10) 'For advancing a line in the input box
   iPageNum = InputBox("Please enter page number to set calibration for.", "SetPageCalibration[Page Number]")
   lPageID = MainMDIForm.ActiveForm.SpicerDoc1.pageID(iPageNum)
   dScaleFactor = InputBox("Please enter scale factor to be used.", "SetPageCalibration[Scale Factor]")
   iUnitType = InputBox("Please enter one of the following unit types." + strNewLine + _
                  "1 = inches, 2 = cm, 3 = pixels, 4 = feet, 5 = mm, 6 = m," + strNewLine + _
                  "7 = points, 8 = twips, 9 = custom 1, 10 = custom 2," + strNewLine + _
                  "11 = custom 3", "SetPageCalibration[Unit Type]")
   docProperties.SetPageCalibration lPageID, dScaleFactor, iUnitType
   MsgBox "Page calibration has been set.", vbInformation
   
   'De-initialize the object variable
   Set docProperties = Nothing
End Sub

Public Sub IDocProperties_UserAttribute()
   'Mark Simpson May11/98
   Dim lOrigUserID As String
   Dim lSetUserID As String
   Dim lGetUserID As String
   Dim docProperties As IDocProperties
   
   'Set what the UserID will be set to
   lSetUserID = InputBox("Enter what the UserID attribute will be set to:")
   
   'Set object variable for IDocProperties interface to doc ctrl object
   Set docProperties = MainMDIForm.ActiveForm.SpicerDoc1.object
   
   'Get the original setting of the UserID attribute, if any
   lOrigUserID = docProperties.UserAttribute(MainMDIForm.ActiveForm.SpicerDoc1.object, "UserID")
   
   'Set the UserID attribute
   docProperties.UserAttribute(MainMDIForm.ActiveForm.SpicerDoc1.object, "UserID") = lSetUserID
   
   'Get the new value
   lGetUserID = docProperties.UserAttribute(MainMDIForm.ActiveForm.SpicerDoc1.object, "UserID")
   
   'Compare the get/set values
   If lSetUserID = lGetUserID Then
      'UserAttribute property is working properly
      MsgBox "UserID attribute was correctly set to:" + Chr(13) + lGetUserID, vbInformation
   Else
      'Get/Set UserID attribute did not work properly
      MsgBox "Error." + Chr(13) + "Set: " + lSetUserID + Chr(13) + "Get: " + lGetUserID, vbCritical
   End If
   
   'Return userID attribute to original setting
   docProperties.UserAttribute(MainMDIForm.ActiveForm.SpicerDoc1.object, "UserID") = lOrigUserID
   
   'De-initialize the object variable
   Set docProperties = Nothing
End Sub

Public Function BuildPermsString(lTempPerms As Long, lPermsString As String) As String
   'Build the string outlining the permissions
   If lTempPerms > 511 Then
      lPermsString = lPermsString + "User Defined 2  "
      lTempPerms = lTempPerms - 512
   End If
   If lTempPerms > 255 Then
      lPermsString = lPermsString + "User Defined 1  "
      lTempPerms = lTempPerms - 256
   End If
   If lTempPerms > 127 Then
      lPermsString = lPermsString + "Static  "
      lTempPerms = lTempPerms - 128
   End If
   If lTempPerms > 63 Then
      lPermsString = lPermsString + "Hide  "
      lTempPerms = lTempPerms - 64
   End If
   If lTempPerms > 31 Then
      lPermsString = lPermsString + "Print  "
      lTempPerms = lTempPerms - 32
   End If
   If lTempPerms > 15 Then
      lPermsString = lPermsString + "Delete  "
      lTempPerms = lTempPerms - 16
   End If
   If lTempPerms > 7 Then
      lPermsString = lPermsString + "Modify  "
      lTempPerms = lTempPerms - 8
   End If
   If lTempPerms > 3 Then
      lPermsString = lPermsString + "Create  "
      lTempPerms = lTempPerms - 4
   End If
   If lTempPerms > 1 Then
      lPermsString = lPermsString + "Write  "
      lTempPerms = lTempPerms - 2
   End If
   If lTempPerms = 1 Then
      lPermsString = lPermsString + "Read"
   End If
   
   BuildPermsString = lPermsString
End Function
Public Function Rotation(iRotation As ROTATION_ANGLE) As String
   ' RobY Nov26/98
   Dim strRotation As String
   
   Select Case iRotation
      Case IN_ROTATION_0
         strRotation = "0 degrees"
      Case IN_ROTATION_90
         strRotation = "90 degrees"
      Case IN_ROTATION_90_CCW
         strRotation = "90 degrees CCW"
      Case IN_ROTATION_90_CW
         strRotation = "90 degrees CW"
      Case IN_ROTATION_270
         strRotation = "270 degrees"
      Case IN_ROTATION_180_REL
         strRotation = "180 degrees relative"
      Case IN_ROTATION_180
         strRotation = "180 degrees"
   End Select
   Rotation = strRotation
End Function
' <EOF IDocProperties.bas>

