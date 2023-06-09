Attribute VB_Name = "mod_IDocSave"
' File:      IDocSave.bas
' Created:   1998Apr28 by Rob Wood, Mark Simpson, and Rob Young
' Purpose:   To test the Spicer Doc Control's IDocSave interface.
' Revisions: RobY Jan4/99  - Added functionality to open more than 1 file. ChildForm1 is no longer used access objects because it has been set to ChildForm() which is an array of ChildForm1's forms.
'                          - Anywhere ChildForm1 was used has been changed to ChildForm(iIndex) or MainMDIForm.ActiveForm.
'   $log$
'   $nokeywords$
'
' Copyright (C) 1998 SPICER Corporation, All rights reserved.
   
Option Explicit
Dim MenuName As Menu

Public Sub IDocSave_CreateThumbnailFile()
   ' RobY & RobW Apr30/98
   ' Modified by RobY May13/98
   ' modified by MarkS Nov18/98, now asks for colour option
   ' modified by MarkS Dec 1, 1998 - changed the CreateThumbnailFile call, had 1 (MIL) as format ID but this format doesn't support colour.  Changed to 7 (BMP).
   Dim docSave As IDocSave
   Dim lPageNum As Long
   Dim strFilename As String
   Dim strColour As String
   Dim bColour As Boolean
   
   ' Set object variable for IDocSave interface to doc ctrl object
   Set docSave = MainMDIForm.ActiveForm.SpicerDoc1.object
   
   ' Prompt user for page number to create preview file of
   lPageNum = InputBox("What page do you want to create a thumbnail of?", "Create Thumbnail File")
   
   ' Ask user if thumbnail will be bilevel or colour
   Do
      strColour = InputBox("Will the thumbnail be bilevel or colour?")
      If strColour = "bilevel" Then
         bColour = False
         Exit Do
      ElseIf strColour = "colour" Then
         bColour = True
         Exit Do
      Else
         MsgBox "Invalid value.  Must be 'bilevel' or 'colour'.  Try again."
      End If
   Loop
   
   ' Prompt user for where they want preview output  to be saved to
   strFilename = InputBox("Specify a location and name for the thumbnail file to go.", "Specify Output Location")
   ' Create preview file
   docSave.CreateThumbnailFile MainMDIForm.ActiveForm.SpicerDoc1.pageID(lPageNum), 320, 320, bColour, 7, strFilename
   
   'De-initialize the object var
   Set docSave = Nothing
End Sub

Public Sub IDocSave_DocumentDirtyStatus()
   'RobY May1/98
   'modified by MarkS Nov19/98, now uses BuildDirtyStatusString function to display what the value means
   Dim docSave As IDocSave
   Dim lStatus As Long
   Dim ltempStatus As Long
   Dim strStatusString As String
   
   ' Set object variable for IDocSave interface to doc ctrl object
   Set docSave = MainMDIForm.ActiveForm.SpicerDoc1.object
   
   ' Get the document's dirty status
   lStatus = docSave.DocumentDirtyStatus
   
   'Call function to build dirty status string
   ltempStatus = lStatus
   strStatusString = BuildDirtyStatusString(ltempStatus, strStatusString)
   
   MsgBox "Document Dirty Status of" + Str(docSave.DocumentDirtyStatus) + " which is:" + Chr(13) + strStatusString
   
   'De-initialize the object var
   Set docSave = Nothing
End Sub

Public Sub IDocSave_Export()
   'RobY May1/98
   'Modified by RobY July14/98
   Dim docSave As IDocSave
   Dim strFilename As String
   Dim lPageID As Long
   Dim iCheckExist As Integer
   Dim bCheckExist As Boolean
   Dim iFormatID As Integer
   Dim iLayerNum As Integer
   Dim iExportType As Integer
   
   ' Set object variable for IDocSave interface to doc ctrl object
   Set docSave = MainMDIForm.ActiveForm.SpicerDoc1.object
   
   iExportType = MsgBox("Do you want to export a page(if you choose No then it will export a layer)?", vbYesNo + vbQuestion, "Export Page or Layer")
   If iExportType = vbNo Then
      iLayerNum = InputBox("Please enter the layer number of the layer you want export.", "Layer Number")
   End If
   ' Prompt user for where they want the export to
   strFilename = InputBox("Specify a location and name for exported file.", "Export Page Location")
   ' Prompt to check existence of file
   iCheckExist = MsgBox("Do you want to check if file already exists?", vbQuestion + vbYesNo, "Check Existence")
   If iCheckExist = vbYes Then
      bCheckExist = True
   Else
      bCheckExist = False
   End If
   'Prompt for format id
   iFormatID = InputBox("Please enter the format type id you want to save the exported file as", "Format Type")
   lPageID = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
   ' Export to the specified location
   If iExportType = vbYes Then
      docSave.Export lPageID, bCheckExist, iFormatID, strFilename, "Exported Page"
   Else
      docSave.Export MainMDIForm.ActiveForm.SpicerDoc1.LayerID(lPageID, iLayerNum), bCheckExist, iFormatID, strFilename, "Exported Page"
   End If
   MsgBox "Export Complete. Check to see if file was created.", vbInformation, "Export Complete"
   
   'De-initialize the object var
   Set docSave = Nothing
End Sub

Public Sub IDocSave_ExportLayerDialog()
   'RobY May1/98
   Dim docSave As IDocSave
   
   ' Set object variable for IDocSave interface to doc ctrl object
   Set docSave = MainMDIForm.ActiveForm.SpicerDoc1.object
   
   'Call Export Layer Dialog for Page 1
   docSave.ExportLayerDialog (MainMDIForm.ActiveForm.SpicerDoc1.pageID(1))
   
   'De-initialize the object var
   Set docSave = Nothing
End Sub

Public Sub IDocSave_ExportLayerDialogAvailability()
   'RobY July1/98
   ' Modified by RobY Dec11/98 - changed global function to change menu status of ImageX according to what is returned from availablity
   Dim docSave As IDocSave
   Dim iAvailable As COMMAND_AVAILABILITY
   Dim lPageID As Long
   
   ' Set object variable for IDocSave interface to doc ctrl object
   Set docSave = MainMDIForm.ActiveForm.SpicerDoc1.object
   ' Set the menu name to change
   Set MenuName = MainMDIForm.ActiveForm.mnu_DocCtrl_IDocSave_Array(3)
   
   lPageID = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
   ' Get the value of availability
   iAvailable = docSave.ExportLayerDialogAvailability(lPageID)
   ' Change the menu status through function
   Availability iAvailable, MenuName
   
   ' De-initialize object var
   Set docSave = Nothing
End Sub

Public Sub IDocSave_ExportPageDialog()
   'RobY May1/98
   Dim docSave As IDocSave
   
   ' Set object variable for IDocSave interface to doc ctrl object
   Set docSave = MainMDIForm.ActiveForm.SpicerDoc1.object
   
   ' Call Export Page Dialog for Page 1
   docSave.ExportPageDialog MainMDIForm.ActiveForm.SpicerDoc1.pageID(1)
   
   ' De-initialize the object var
   Set docSave = Nothing
End Sub

Public Sub IDocSave_ExportPageDialogAvailability()
   'RobY July1/98
   ' Modified by RobY Dec11/98 - changed global function to change menu status of ImageX according to what is returned from availablity
   Dim docSave As IDocSave
   Dim iAvailable As COMMAND_AVAILABILITY
   Dim lPageID As Long
   
   ' Set object variable for IDocSave interface to doc ctrl object
   Set docSave = MainMDIForm.ActiveForm.SpicerDoc1.object
   ' Set the menu name to change
   Set MenuName = MainMDIForm.ActiveForm.mnu_DocCtrl_IDocSave_Array(4)

   lPageID = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
   ' Get the value of availability
   iAvailable = docSave.ExportPageDialogAvailability(lPageID)
   ' Change the menu status through function
   Availability iAvailable, MenuName
   
   ' De-initialize object var
   Set docSave = Nothing
End Sub

Public Sub IDocSave_Filename()
   'RobY May1/98
   Dim docSave As IDocSave
   
   ' Set object variable for IDocSave interface to doc ctrl object
   Set docSave = MainMDIForm.ActiveForm.SpicerDoc1.object
   
   ' Display the filename for the document
   MsgBox ("Filename: " + docSave.FileName(MainMDIForm.ActiveForm.SpicerDoc1.pageID(1)))
   
   ' De-initialize the object var
   Set docSave = Nothing
End Sub

Public Sub IDocSave_Format()
   ' RobY May1/98
   Dim docSave As IDocSave
   
   ' Set object variable for IDocSave interface to doc ctrl object
   Set docSave = MainMDIForm.ActiveForm.SpicerDoc1.object
   
   ' Display the format for the document
   MsgBox ("Format: " + Str(docSave.Format(MainMDIForm.ActiveForm.SpicerDoc1.pageID(1))))
   
   ' De-initialize the object var
   Set docSave = Nothing
End Sub

Public Sub IDocSave_HeaderRotation()
   'RobY Nov26/98
   Dim docSave As IDocSave
   Dim iCurHeaderRot As ROTATION_ANGLE
   Dim iNewHeaderRot As ROTATION_ANGLE
   Dim iGetNewHeaderRot As ROTATION_ANGLE
   Dim strHeaderRot As String
   Dim lLayerID As String
   Dim iResponse As Integer
   Dim iLayerNum As Integer
   
   'Set object variable for IDocSave interface to doc ctrl object
   Set docSave = MainMDIForm.ActiveForm.SpicerDoc1.object
   
   ' Prompt for the raster layer number
   iLayerNum = InputBox("Please enter the layer number of the raster.", "Layer Number")
   ' Convert layer number to layer id
   lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   ' Get the current header rotation
   iCurHeaderRot = docSave.HeaderRotation(lLayerID)
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
      docSave.HeaderRotation(lLayerID) = iNewHeaderRot
      ' Get new header rotation
      iGetNewHeaderRot = docSave.HeaderRotation(lLayerID)
      If iNewHeaderRot <> iGetNewHeaderRot Then
         MsgBox "Failed to change header rotation to new setting!", vbCritical, "Failed"
      Else
         MsgBox "Header rotation has been changed to new setting.", vbInformation, "Success"
      End If
   Else
      MsgBox "The header rotation was not changed.", vbInformation, "Header Rotation"
   End If
    
   'De-initialize the object variable
   Set docSave = Nothing
End Sub

Public Sub IDocSave_Label()
   ' RobY May1/98
   ' modified by MarkS Nov20/98, added functionality to get label of the document, page or layer
   ' Modified by RobY Dec4/98 - added functionality so user can cancel dialog
   Dim docSave As IDocSave
   Dim lPage As Long
   Dim strLayer As String
   Dim strObject As String
         
   ' Set object variable for IDocSave interface to doc ctrl object
   Set docSave = MainMDIForm.ActiveForm.SpicerDoc1.object
   
   ' Ask user which object they want information on
   Do
      strObject = InputBox("Do you want the label for a document, page or layer?")
      If strObject = "document" Then
         Exit Do
      ElseIf strObject = "page" Then
         Exit Do
      ElseIf strObject = "layer" Then
         Exit Do
      ElseIf strObject = "" Then
         Exit Sub
      Else
         MsgBox "Invalid value.  Must be 'document' or 'page' or 'layer'.  Try again."
      End If
   Loop
     
   Select Case strObject
      Case "document":
         MsgBox "Document's label is: " + docSave.Label(MainMDIForm.ActiveForm.SpicerDoc1.RootID)
      Case "page":
         'Get active page's ID
         lPage = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
         MsgBox "Currently viewed page's label is: " + docSave.Label(lPage)
      Case "layer":
         'ask which layer they want the label info for
         strLayer = InputBox("Which layer of the currently viewed page do you want the label for?")
         MsgBox "Layer " + strLayer + " of currently viewed page has label: " + Chr(13) + docSave.Label(MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, Val(strLayer)))
   End Select

End Sub

Public Sub IDocSave_LayerAttributesDialog()
   'RobY May1/98
   Dim docSave As IDocSave
   
   ' Set object variable for IDocSave interface to doc ctrl object
   Set docSave = MainMDIForm.ActiveForm.SpicerDoc1.object
   
   ' Call Layer Attribute Dialog for Page 1
   docSave.LayerAttributesDialog (MainMDIForm.ActiveForm.SpicerDoc1.pageID(1))
   
   ' De-initialize the object var
   Set docSave = Nothing
End Sub

Public Sub IDocSave_LayerAttribDialogAvailability()
   'RobY July1/98
   ' Modified by RobY Dec11/98 - changed global function to change menu status of ImageX according to what is returned from availablity
   Dim docSave As IDocSave
   Dim iAvailable As COMMAND_AVAILABILITY
   Dim lPageID As Long
   
   ' Set object variable for IDocSave interface to doc ctrl object
   Set docSave = MainMDIForm.ActiveForm.SpicerDoc1.object
   ' Set the menu name to change
   Set MenuName = MainMDIForm.ActiveForm.mnu_DocCtrl_IDocSave_Array(9)
   
   lPageID = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
   ' Get the value of availability
   iAvailable = docSave.LayerAttribDialogAvailability(lPageID)
   ' Change the menu status through function
   Availability iAvailable, MenuName
   
   ' De-initialize object var
   Set docSave = Nothing
End Sub

Public Sub IDocSave_ObjectDirtyStatus()
   'RobY May1/98
   ' modified by MarkS Nov20/98, now uses BuildDirtyStatusString function to display what the value means
   Dim docSave As IDocSave
   Dim lActivePage As Long
   Dim lStatus As Long
   Dim ltempStatus As Long
   Dim strStatusString As String
   
   ' Set object variable for IDocSave interface to doc ctrl object
   Set docSave = MainMDIForm.ActiveForm.SpicerDoc1.object
   
   ' Get the active page
   lActivePage = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
   
   ' Get the object's dirty status
   lStatus = docSave.ObjectDirtyStatus(lActivePage)
   
   'Call function to build dirty status string
   ltempStatus = lStatus
   strStatusString = BuildDirtyStatusString(ltempStatus, strStatusString)
   
   MsgBox "Object has Dirty Status of" + Str(docSave.DocumentDirtyStatus) + " which is:" + Chr(13) + strStatusString
   
   'De-initialize the object var
   Set docSave = Nothing
   
End Sub

Public Sub IDocSave_PageAttributesDialog()
   'RobY May1/98
   Dim docSave As IDocSave
   
   ' Set object variable for IDocSave interface to doc ctrl object
   Set docSave = MainMDIForm.ActiveForm.SpicerDoc1.object
   
   ' Call Page Attribute Dialog for Page 1
   docSave.PageAttributesDialog (MainMDIForm.ActiveForm.SpicerDoc1.pageID(1))
   
   ' De-initialize the object var
   Set docSave = Nothing
End Sub

Public Sub IDocSave_PageAttribDialogAvailability()
   'RobY July1/98
   ' Modified by RobY Dec11/98 - changed global function to change menu status of ImageX according to what is returned from availablity
   Dim docSave As IDocSave
   Dim iAvailable As COMMAND_AVAILABILITY
   Dim lPageID As Long
   
   ' Set object variable for IDocSave interface to doc ctrl object
   Set docSave = MainMDIForm.ActiveForm.SpicerDoc1.object
   ' Set the menu name to change
   Set MenuName = MainMDIForm.ActiveForm.mnu_DocCtrl_IDocSave_Array(11)
   
   lPageID = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
   ' Get the value of availability
   iAvailable = docSave.PageAttribDialogAvailability(lPageID)
   ' Change the menu status through function
   Availability iAvailable, MenuName
   
   ' De-initialize object var
   Set docSave = Nothing
End Sub

Public Sub IDocSave_RasterInfoString()
'   'RobY May1/98
'   Dim docSave As IDocSave
'
'   ' Set object variable for IDocSave interface to doc ctrl object
'   Set docSave = MainMDIForm.ActiveForm.SpicerDoc1.object
'
'   ' Display the notes for the raster
'   MsgBox ("Notes: " + docSave.RasterInfoString(MainMDIForm.ActiveForm.SpicerDoc1.LayerID(1, 0)))
'
'   ' De-initialize the object var
'   Set docSave = Nothing
   
   
   ' RobY May1/98
   ' modified by MarkS Nov20/98, used code from IDocProperties_RasterInfoString
   Dim lLayerID As Integer
   Dim lFormatID As Integer
   Dim lOrigRasterNotes As String
   Dim lNewRasterNotes As String
   Dim lGetRasterNotes As String
   Dim docSave As IDocSave
   Dim docProperties As IDocProperties
   
   'Get the layer ID of the current active layer
   lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, 1)
   
   'Set object variable for IDocProperties interface to doc ctrl object
   Set docProperties = MainMDIForm.ActiveForm.SpicerDoc1.object
   
   'Make sure user has opened a file that can have raster notes
   lFormatID = docProperties.Format(MainMDIForm.ActiveForm.SpicerDoc1.object)
   
   'Clear the docProperties and set object variable for IDocSave interface to doc ctrl object
   Set docProperties = Nothing
   Set docSave = MainMDIForm.ActiveForm.SpicerDoc1.object

   Select Case lFormatID
      Case 1, 3, 8, 11 - 15, 22 - 26, 32, 33, 36 - 39, 49, 60 - 62, 64 - 66
         'Ask user for what raster notes will be set to
         lNewRasterNotes = InputBox("Enter what the raster notes will be:")
         
         'Get the Original Raster Notes
         lOrigRasterNotes = docSave.RasterInfoString(lLayerID)
         MsgBox "Original RasterNotes for this layer were:" + Chr(13) + lOrigRasterNotes
                
         'Set the RasterNotes
         docSave.RasterInfoString(lLayerID) = lNewRasterNotes
         
         'Get the RasterNotes of the current active layer
         lGetRasterNotes = docSave.RasterInfoString(lLayerID)
                  
         'Display result
         If lNewRasterNotes = lGetRasterNotes Then
            'success
            MsgBox "The new RasterNotes for this layer are:" + Chr(13) + lGetRasterNotes, vbInformation
         Else
            'failure
            MsgBox "Get/Set RasterNotes not matching." + Chr(13) + "Orig: " + lOrigRasterNotes + Chr(13) + "Set: " + lNewRasterNotes + Chr(13) + "Get:" + lGetRasterNotes, vbCritical
         End If
         
         'Return raster notes to original setting
         docSave.RasterInfoString(lLayerID) = lOrigRasterNotes
         
      Case Else
         MsgBox "RasterNotes are not available for this format."
   End Select
    
   'De-initialize the object variable
   Set docSave = Nothing
End Sub

Public Sub IDocSave_Save()
   'RobY May1/98
   Dim docSave As IDocSave
   Dim strFilename As String
   Dim iFormatID As Integer
   
   ' Set object variable for IDocSave interface to doc ctrl object
   Set docSave = MainMDIForm.ActiveForm.SpicerDoc1.object
   
   'Prompt the user for format to be saved as
   iFormatID = InputBox("Please specify the format ID that you want to save as." + Chr(13) + "Refer to help file for available formats.", "File Format")
  'Prompt the user for filename to be saved as
   strFilename = InputBox("Specify a filename to save as.", "Save File As")
   'Save the file
   docSave.Save 0, True, iFormatID, strFilename, "Save Test Worked"
   
   'De-initialize the object var
   Set docSave = Nothing
End Sub

Public Sub IDocSave_SaveAsDialog()
   'RobY May1/98
   Dim docSave As IDocSave
   
   ' Set object variable for IDocSave interface to doc ctrl object
   Set docSave = MainMDIForm.ActiveForm.SpicerDoc1.object
   
   'Display the save as dialog box
   docSave.SaveAsDialog (False)
   
   ' De-initialize the object var
   Set docSave = Nothing
End Sub


Public Function BuildDirtyStatusString(ltempStatus As Long, strStatusString As String) As String
   'Build the string outlining the dirty status
   If ltempStatus > 8 Then
      strStatusString = strStatusString + "Modified since load  "
      ltempStatus = ltempStatus - 16
   End If
   If ltempStatus > 4 Then
      strStatusString = strStatusString + "New  "
      ltempStatus = ltempStatus - 8
   End If
   If ltempStatus > 2 Then
      strStatusString = strStatusString + "Deleted  "
      ltempStatus = ltempStatus - 4
   End If
   If ltempStatus > 1 Then
      strStatusString = strStatusString + "Modified  "
      ltempStatus = ltempStatus - 2
   End If
   If ltempStatus > 0 Then
      strStatusString = strStatusString + "Clean"
      ltempStatus = ltempStatus - 1
   End If
   
   BuildDirtyStatusString = strStatusString
End Function
' <EOF IDocSave.bas>

