Attribute VB_Name = "mod_IDocContents"
' File:      IDocContents.bas
' Created:   1998Apr28 by Rob Wood, Mark Simpson, and Rob Young
' Purpose:   To test the Spicer Doc Control's IDocContents interface.
' Revisions: RobY Jan4/99  - Added functionality to open more than 1 file. ChildForm1 is no longer used access objects because it has been set to ChildForm() which is an array of ChildForm1's forms.
'                          - Anywhere ChildForm1 was used has been changed to ChildForm(iIndex) or MainMDIForm.ActiveForm.
'   $log$
'   $nokeywords$
'
' Copyright (C) 1998 SPICER Corporation, All rights reserved.

Option Explicit
Dim MenuName As Menu

Public Sub IDocContents_CloseDocument()
   ' RobY Oct21/98
   ' Modified by RobY Dec15 - New boolean parameter added to command for saving modified documents.
   Dim docContents As IDocContents
   Dim iResponse As Integer
   
   ' Set object variable for IDocContents interface to doc ctrl object
   Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
   
   ' Prompt user if to save before closing
   iResponse = MsgBox("Do you want to save the document before closing?", vbYesNo + vbQuestion, "Save Document")
   If iResponse = vbYes Then
      ' Save before closing document if any changes made
      docContents.CloseDocument True
   Else
      ' Close the current document without saving
      docContents.CloseDocument False
   End If
   
   ' De-initialize the object variable
   Set docContents = Nothing
End Sub

Public Sub IDocContents_ConsolidateLayersDialog()
   ' RobW May 1/98
   Dim lPageID As Long
   Dim docContents As IDocContents
   
   ' Get the page ID for current page of active document
   lPageID = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
   ' Set object variable for IDocContents interface to doc ctrl object
   Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
   ' Open the Consolidate Layers Dialog
   docContents.ConsolidateLayersDialog lPageID
   
   ' De-initialize the object variable
   Set docContents = Nothing
End Sub

Public Sub IDocContents_ConsolidateLayersDlgAvailability()
   'RobY June24/98
   ' modified by MarkS Nov30, 1998 - now uses global function for string determination
   ' Modified by RobY Dec3/98 - changed global function to change menu status of ImageX according to what is returned from availablity
   Dim docContents As IDocContents
   Dim iAvailable As COMMAND_AVAILABILITY
   Dim lPageID As Long
   
   ' Set object variable for IDocContents interface to doc ctrl object
   Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
   ' Set the menu name to change
   Set MenuName = MainMDIForm.ActiveForm.mnu_DocCtrl_IDocContents_Array(1)
   
   ' Get the pageID for the active page
   lPageID = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
   
   ' Get the value of availability
   'iAvailable = docContents.ConsolidateLayersDlgAvailability(lPageID)
   ' Change the menu status through function
   'Availability iAvailable, MenuName
   
   ' De-initialize object var
   Set docContents = Nothing
End Sub

Public Sub IDocContents_DeleteObject()
   ' RobW May 8/98
   ' RobW Modified May 11/98
   Dim docContents As IDocContents
   Dim lPageID As Long
   Dim lLayerID As Long
   Dim strPageorLayer As String
   Dim strLayerID As String
   
   ' Set object variable for IDocContents interface to doc ctrl object
   Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
   
   ' Get page ID for active page
   lPageID = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
   
   ' Ask user if they want to delete a page or a layer
   Do
      strPageorLayer = InputBox("Would you like to delete the current page or the current layer?", "Delete a page or a layer?")
      If strPageorLayer = "page" Then
         ' Delete the active page
         docContents.DeleteObject lPageID
         Exit Do
      ElseIf strPageorLayer = "layer" Then
         ' Ask the user which layer they would like to delete
         strLayerID = InputBox("Which layer of the current page would you like to remove?", "Remove which layer?")
         ' Get layer ID for layer 1
         lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(lPageID, Val(strLayerID))
         
         ' Delete layer 1 from page 1
         docContents.DeleteObject lLayerID
         Exit Do
      Else
         MsgBox "Invalid value.  Must be 'page' or 'layer'.  Try again."
      End If
   Loop

   ' De-initialize object var
   Set docContents = Nothing
End Sub

Public Sub IDocContents_FlushDownloadDirectory()
   ' RobY Oct21/98
   ' Modified by RobY Dec8/98 - Added message to display that command has been executed
   Dim docContents As IDocContents
   
   ' Set object variable for IDocContents interface to doc ctrl object
   Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
   
   ' Flush download directory
   docContents.FlushDownloadDirectory
   ' Display message that the command has been executed
   MsgBox "The download directory has been deleted.", vbInformation, "FlushDownloadDirectory"
   
   ' De-initialize the object variable
   Set docContents = Nothing
End Sub

Public Sub IDocContents_FirstPageID()
   ' RobW May 5/98
   Dim docContents As IDocContents
   Dim lFirstPageID
   
   ' Set object variable for IDocContents interface to doc ctrl object
   Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
   ' Get and display the FirstPageID
   lFirstPageID = docContents.FirstPageID
   If lFirstPageID <> 0 Then
      MsgBox "First Page ID:" + Str(lFirstPageID), vbInformation
   Else
      MsgBox "First Page ID:" + Str(lFirstPageID), vbCritical
   End If
   
   ' De-initialize the object variable
   Set docContents = Nothing
End Sub

Public Sub IDocContents_ImportLayer()
   ' RobW May 5/98
   ' Modified by RobY Sept18/98  Removed last parameter(NewLayerID)
   Dim docContents As IDocContents
   Dim lPageID As Long
   Dim lLayerID As Long
   Dim iNumLayers As Integer
   Dim lNewObjectID As Long
   
   ' Get the page ID for current page of active document
   lPageID = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
   ' Set object variable for IDocContents interface to doc ctrl object
   Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
   ' Use the common dialog to allow user to select file to import
   MainMDIForm.ActiveForm.CommonDialog1.Flags = cdlOFNHideReadOnly
   MainMDIForm.ActiveForm.CommonDialog1.DialogTitle = "ImageX - Select a File for the ImportLayer Method"
   MainMDIForm.ActiveForm.CommonDialog1.ShowOpen
   ' Import the selected file
   docContents.ImportLayer lPageID, 0, MainMDIForm.ActiveForm.CommonDialog1.FileName
   ' Get the id of the last layer
   iNumLayers = docContents.NumberOfLayers(lPageID)
   lLayerID = docContents.LayerID(lPageID, iNumLayers)
   lNewObjectID = MainMDIForm.ActiveForm.SpicerDoc1.NewestObjectID
   ' Display the layerID of the newly imported layer
   If lLayerID = lNewObjectID Then
      MsgBox "ID of the imported layer: " + Str(lLayerID), vbInformation, "Success"
   Else
      MsgBox "ID of the imported layer: " + Str(lLayerID), vbCritical, "Error"
   End If
   ' Initialize filename to null
   MainMDIForm.ActiveForm.CommonDialog1.FileName = ""

   ' De-initialize object var
   Set docContents = Nothing
End Sub

Public Sub IDocContents_ImportLayerDialog()
   ' RobW May 1/98
   Dim lPageID As Long
   Dim docContents As IDocContents
   
   ' Get the page ID for current page of active document
   lPageID = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
   ' Set object variable for IDocContents interface to doc ctrl object
   Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
   ' Open the Import Layer Dialog for page one
   docContents.ImportLayerDialog lPageID
   
   ' De-initialize the object variable
   Set docContents = Nothing
End Sub

Public Sub IDocContents_ImportLayerDialogAvailability()
   'RobY June25/98
   ' modified by MarkS Nov30, 1998 - now uses global function for string determination
   ' Modified by RobY Dec3/98 - changed global function to change menu status of ImageX according to what is returned from availablity
   Dim docContents As IDocContents
   Dim iAvailable As COMMAND_AVAILABILITY
   Dim lPageID As Long
   Dim strResult As String
   
   ' Set object variable for IDocContents interface to doc ctrl object
   Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
   ' Set the menu name to change
   Set MenuName = MainMDIForm.ActiveForm.mnu_DocCtrl_IDocContents_Array(6)
   
   lPageID = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
   ' Get the value of availability
   iAvailable = docContents.ImportLayerDialogAvailability(lPageID)
   ' Change the menu status through function
   Availability iAvailable, MenuName
     
   ' De-initialize object var
   Set docContents = Nothing
End Sub

Public Sub IDocContents_ImportLayertoRegion()
   'MarkS Nov24/98
   Dim docContents As IDocContents
   Dim lPageID As Long
   Dim dX1 As Double
   Dim dY1 As Double
   Dim dX2 As Double
   Dim dY2 As Double
   Dim strRegionMode As String
   Dim regionMode As LAYER_REGION
   Dim lLayerID As Long
   Dim iNumLayers As Integer
   Dim lNewObjectID As Long
   
   'Set object variable for IDocContents interface to doc ctrl object
   Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
   
   ' Get the page ID for current page of active document
   lPageID = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
   
   ' Use the common dialog to allow user to select file to import
   MainMDIForm.ActiveForm.CommonDialog1.Flags = cdlOFNHideReadOnly
   MainMDIForm.ActiveForm.CommonDialog1.DialogTitle = "ImageX - Select a File for the ImportLayerToRegion Method"
   MainMDIForm.ActiveForm.CommonDialog1.ShowOpen
   
   ' Ask the user for the top left and bottom right coordinates for region
   MsgBox "You will now be prompted for the top left and bottom right coordinates of the region to import into." + Chr(13) + "Note that the y axis is positive moving in the up direction, and the x axis is positive moving to the right."
   dX1 = InputBox("Please enter the top left X coordinate:", "X1")
   dY1 = InputBox("Please enter the top left Y coordinate:", "Y1")
   dX2 = InputBox("Please enter the bottom right X coordinate:", "X2")
   dY2 = InputBox("Please enter the bottom right Y coordinate:", "Y2")
   
   ' Ask the user for the region mode
   strRegionMode = InputBox("Please enter the region mode:" + Chr(13) + "1 = IN_LYRRGN_POSITION" + Chr(13) + "2 = IN_LYRRGN_FIT" + Chr(13) + "3 = IN_LYRRGN_PRESERVE_ASPECT", "regionMode")
   Select Case strRegionMode
      Case 1
         regionMode = IN_LYRRGN_POSITION
      Case 2
         regionMode = IN_LYRRGN_FIT
      Case 3
         regionMode = IN_LYRRGN_PRESERVE_ASPECT
   End Select
      
   ' Import the selected file
   docContents.ImportLayerToRegion lPageID, MainMDIForm.ActiveForm.CommonDialog1.FileName, "ImportLayerToRegion test", 0, Val(dX1), Val(dY1), Val(dX2), Val(dY2), IN_UNITS_INCH, regionMode
   
   ' Get the id of the last layer
   iNumLayers = docContents.NumberOfLayers(lPageID)
   lLayerID = docContents.LayerID(lPageID, iNumLayers)
   lNewObjectID = MainMDIForm.ActiveForm.SpicerDoc1.NewestObjectID
   
   ' Display the layerID of the newly imported layer
   If lLayerID = lNewObjectID Then
      MsgBox "ID of the imported layer: " + Str(lLayerID), vbInformation, "Success"
   Else
      MsgBox "ID of the imported layer: " + Str(lLayerID), vbCritical, "Error"
   End If
   ' Initialize filename to null
   MainMDIForm.ActiveForm.CommonDialog1.FileName = ""

   'De-initialize object var
   Set docContents = Nothing
End Sub
Public Sub IDocContents_ImportPage()
   ' RobW May 5/98
   ' Modified By RobY June 22/98
   ' Modified by RobY Sept18/98  Removed last parameter(NewPageID)
   Dim docContents As IDocContents
   Dim lNewPageID As Long
   Dim lPageID As Long
   Dim lParentID As Long
   Dim iPageNum As Integer
   Dim iPosition As Integer
   Dim strPageLabel As String
   Dim ActivePage As IActivePage
   Dim iAnswer As Integer
   
   ' Set object variable for IDocContents interface to doc ctrl object
   Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
   ' Set object variable for IActivePage interface to the Spicer Doc ctrl object
   Set ActivePage = MainMDIForm.ActiveForm.SpicerView1.object
   ' Use the common dialog to allow user to select file to import
   MainMDIForm.ActiveForm.CommonDialog1.Flags = cdlOFNHideReadOnly
   MainMDIForm.ActiveForm.CommonDialog1.DialogTitle = "ImageX - Select a File for the ImportPage Method"
   MainMDIForm.ActiveForm.CommonDialog1.ShowOpen
   iPageNum = InputBox("Please enter page number relative to where the new page is imported.", "Relative Page Number")
   lPageID = MainMDIForm.ActiveForm.SpicerDoc1.pageID(iPageNum)
   lParentID = InputBox("Please enter parent ID." + Chr(10) + Chr(13) + _
                  "Same parent as page number = 0" + Chr(10) + Chr(13) + _
                  "Subordinate to specific page = objectID", "Parent ID")
   iPosition = InputBox("Please enter one of the following for position." + Chr(13) + Chr(10) + _
                  "0 = Before Target Page Number" + Chr(10) + Chr(13) + _
                  "1 = After Target Page Number" + Chr(10) + Chr(13) + _
                  "2 = Beginning of Document" + Chr(13) + Chr(10) + _
                  "3 = End of Document", "Position")
   strPageLabel = InputBox("Please enter a page label", "Page Label")
   ' Import the selected file into the document
   docContents.ImportPage 0, lPageID, iPosition, strPageLabel, MainMDIForm.ActiveForm.CommonDialog1.FileName
   ' Get the newest object id
   lNewPageID = MainMDIForm.ActiveForm.SpicerDoc1.NewestObjectID
   ' Display the PageID of the newly imported page
   If lNewPageID <> 0 Then
      MsgBox "Imported the selected doc." + Chr(13) + "ID of the imported page:" + Str(lNewPageID), vbInformation
      ActivePage.GotoPage lNewPageID
      iAnswer = MsgBox("Is this the page that you imported?", vbYesNo + vbQuestion, "Verify Page")
      If iAnswer = vbYes Then
         MsgBox "Import was successful", vbInformation, "Successful"
      Else
         MsgBox "Import failed", vbCritical, "Failed"
      End If
   Else
      MsgBox "Unsuccessfully imported the selected doc." + Chr(13) + "ID of the imported page:" + Str(lNewPageID), vbCritical
   End If
   ' Initialize filename to null
   MainMDIForm.ActiveForm.CommonDialog1.FileName = ""

   
   ' De-initialize object var
   Set docContents = Nothing
End Sub

Public Sub IDocContents_ImportPageDialog()
   ' RobW May 4/98
   Dim docContents As IDocContents
   Dim lPageID As Long
   
   ' Get the page ID for current page of active document
   lPageID = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
   ' Set object variable for IDocContents interface to doc ctrl object
   Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
   ' Open the Import Page Dialog
   docContents.ImportPageDialog lPageID, IN_NEWPAGE_AFTER
   
   ' De-initialize the object variable
   Set docContents = Nothing
End Sub

Public Sub IDocContents_ImportPageDialogAvailability()
   'RobY July17/98
   ' modified by MarkS Nov30, 1998 - now uses global function for string determination
   ' Modified by RobY Dec3/98 - changed global function to change menu status of ImageX according to what is returned from availablity
   Dim docContents As IDocContents
   Dim iAvailable As COMMAND_AVAILABILITY
   Dim strResult As String
   
   ' Set object variable for IDocContents interface to doc ctrl object
   Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
   ' Set the menu name to change
   Set MenuName = MainMDIForm.ActiveForm.mnu_DocCtrl_IDocContents_Array(9)
   ' Get the value of availability
   iAvailable = docContents.ImportPageDialogAvailability
   ' Change the menu status through function
   Availability iAvailable, MenuName
   
   ' De-initialize object var
   Set docContents = Nothing
End Sub

Public Sub IDocContents_InsertDocument()
   ' MarkS Dec 1, 1998 -
   Dim docContents As IDocContents
   Dim rasterYet As VbMsgBoxResult
   Dim strPageID As String
      
   ' Set object variable for IDocContents interface to doc ctrl object
   Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
   
   ' Ask user if they have performed a raster operation first. Stop if they have not.
   rasterYet = MsgBox("Have you performed a raster operation yet?", vbYesNo, "Must have completed a raster operation first")

   If rasterYet = vbNo Then
      MsgBox "The InsertDocument method is used when the 'OverWrite Raster' keyname in the [System] section of the Imagenation INI file is set to 0 (default). When 'OverWrite Raster=0', the Edit Control creates new raster images that do not overwrite the active document. InsertDocument is used to place the new raster into a document control.", , "Need To Perform Raster Operation First To Test This"
   Else
      ' Assuming raster operation has been performed with RasterOverwrite set to off.
      ' ask user for the page ID of the newly created raster (came from new page event?)
      strPageID = InputBox("Enter the pageID returned from the NewestPageID event", "PageID?")

      docContents.InsertDocument Val(strPageID)
   End If

   ' De-initialize the object variable
   Set docContents = Nothing
End Sub

Public Sub IDocContents_InternetFileDownload()
   ' RobY Oct21/98
   ' Modified by RobY Nov20/98   Removed LocalFileName parameter
   ' Modified by RobY Dec8/98 - Added prompts for user info
   Dim docContents As IDocContents
   Dim strServer As String
   Dim strUserID As String
   Dim strPassword As String
   Dim iPort As Integer
   Dim strRemoteFilePath As String

   ' Set object variable for IDocContents interface to doc ctrl object
   Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
   
   ' Prompt for server
   strServer = InputBox("Please enter the Internet server or URL name. (ex. www.spicer.com)", "Name")
   ' Prompt for user id
   strUserID = InputBox("Please enter a user id. (If not required, leave blank.)", "User ID")
   ' Prompt for password
   strPassword = InputBox("Please enter a password. (If not required, leave blank.)", "Password")
   ' Prompt for port #
   iPort = InputBox("Please enter the port number to use. (Use 0 for default port.)", "Port")
   ' Prompt for path and filename
   strRemoteFilePath = InputBox("Please enter path and filename to download. (ex. support/sherry/testing/sample129.smf)", "Name")
   ' Download the specified file
   docContents.InternetFileDownload strServer, strUserID, strPassword, iPort, strRemoteFilePath
   
   ' De-initialize the object variable
   Set docContents = Nothing
End Sub

Public Sub IDocContents_InternetFileUpload()
   ' RobY Oct21/98
   ' Modified by RobY Dec8/98 - Added prompts for user to specify info for upload
   Dim docContents As IDocContents
   Dim strServer As String
   Dim strUserID As String
   Dim strPassword As String
   Dim iPort As Integer
   Dim strRemoteLocation As String
   Dim strUploadFilename As String
   
   ' Set object variable for IDocContents interface to doc ctrl object
   Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
   
   'Prompt for file to upload
   MainMDIForm.ActiveForm.CommonDialog1.CancelError = True
   MainMDIForm.ActiveForm.CommonDialog1.Flags = cdlOFNHideReadOnly
   MainMDIForm.ActiveForm.CommonDialog1.DialogTitle = "Select File to Upload"
   MainMDIForm.ActiveForm.CommonDialog1.ShowOpen
   strUploadFilename = MainMDIForm.ActiveForm.CommonDialog1.FileName
   MainMDIForm.ActiveForm.CommonDialog1.FileName = ""
   ' Prompt for server
   strServer = InputBox("Please enter the Internet server or URL name. (ex. www.spicer.com)", "Name")
   ' Prompt for path and filename
   strRemoteLocation = InputBox("Please enter specific location (optional) to upload to.(ex. \future\support\sherry\testing\)", "Location")
   ' Prompt for user id
   strUserID = InputBox("Please enter a user id. (If not be required, leave blank.)", "User ID")
   ' Prompt for password
   strPassword = InputBox("Please enter a password. (If not be required, leave blank.)", "Password")
   ' Prompt for port #
   iPort = InputBox("Please enter the port number to use. (Use 0 for default port.)", "Port")
   ' Upload the specified file
   docContents.InternetFileUpload strServer, strUserID, strPassword, iPort, strUploadFilename, strRemoteLocation
   ' Initialize filename to null
   MainMDIForm.ActiveForm.CommonDialog1.FileName = ""

   ' De-initialize the object variable
   Set docContents = Nothing
End Sub

Public Sub IDocContents_LayerID()
   'RobW May 8/98
   Dim lPageID As Long
   Dim lLayerCount As Long
   Dim iTemp As Integer
   Dim sTemp As String
   Dim docContents As IDocContents
   
   ' Get the page ID for current page of active document
   lPageID = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
   ' Set object variable for IDocContents interface to doc ctrl object
   Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
   ' Get number of layers for current page of current doc
   lLayerCount = docContents.NumberOfLayers(lPageID)
   sTemp = "The current page has" + Str(lLayerCount) + " layer(s)." + Chr(13)
   ' Get Layer ID for all layers on current page of current doc
   For iTemp = 1 To lLayerCount
      sTemp = sTemp + "Layer" + Str(iTemp) + " ID: " + Str(docContents.LayerID(lPageID, iTemp)) + Chr(13)
   Next iTemp
   ' Display number of layers and ID's for each
   MsgBox sTemp, vbInformation
   
   Set docContents = Nothing
End Sub

Public Sub IDocContents_NewDocument()
   ' RobW May5/98
   ' Modified by RobW May 11/98 - Don't do bind here; user can use ImageX menu item
   Dim docContents As IDocContents
   Dim lRootID As Long
   
   ' Set object variable for IDocContents interface to doc ctrl object
   Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
   ' Create new document
   docContents.NewDocument
   ' Set child window's title
   MainMDIForm.ActiveForm.Caption = "New Document"
   ' Display new document's root ID / document ID
      ' Display new document's root ID / document ID
   lRootID = MainMDIForm.ActiveForm.SpicerDoc1.RootID
   If lRootID <> 0 Then
      MsgBox "New document Root ID:" + Str(lRootID), vbInformation
   Else
      MsgBox "New document Root ID:" + Str(lRootID), vbCritical
   End If
   
   ' De-initialize object var
   Set docContents = Nothing
End Sub

Public Sub IDocContents_NewestObjectID()
   ' RobY Oct21/98
   Dim docContents As IDocContents
   Dim lNewObjectID As Long
   
   ' Set object variable for IDocContents interface to doc ctrl object
   Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
   ' Get the ID of the most recently created page or layer.
   lNewObjectID = docContents.NewestObjectID
   MsgBox "Newest Object ID: " + Str(lNewObjectID), vbInformation, "Newest Object ID"
   
   ' De-initialize the object variable
   Set docContents = Nothing
End Sub

Public Sub IDocContents_NewLayer()
   ' RobW May 7/98
   ' Modified by RobY Sept 21/98 Removed last parameter(NewLayerID) from NewLayer command
   Dim docContents As IDocContents
   Dim lPageID As Long
   Dim lNewLayerID As Long
   Dim iNumLayers As Integer
   Dim lNewObjectID As Long
   
   ' Get the page ID for current page of active document
   lPageID = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
   ' Set object variable for IDocContents interface to doc ctrl object
   Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
   ' Add a new layer
   docContents.NewLayer lPageID, IN_LAYER_FULLEDIT
   ' Get the id of the last layer
   iNumLayers = docContents.NumberOfLayers(lPageID)
   lNewLayerID = docContents.LayerID(lPageID, iNumLayers)
   lNewObjectID = MainMDIForm.ActiveForm.SpicerDoc1.NewestObjectID
   ' Display the layerID of the newly imported layer
   If lNewLayerID = lNewObjectID Then
      MsgBox "ID of the new layer: " + Str(lNewLayerID), vbInformation
   Else
      MsgBox "ID of the new layer: " + Str(lNewLayerID), vbCritical
   End If
   
   ' De-initialize object var
   Set docContents = Nothing
End Sub

Public Sub IDocContents_NewPage()
   ' RobW May 8/98
   ' Modified by RobY Aug11/98      Changed first parameter to 0
   ' Modified by RobY Sept21/98     Removed last parameter(lNewLayerID) from NewPage command
   Dim docContents As IDocContents
   Dim lBeforePageCount As Integer
   Dim lAfterPageCount As Integer
   
   ' Set object variable for IDocContents interface to doc ctrl object
   Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
   ' Get current number of pages for current doc
   lBeforePageCount = docContents.NumberOfPages
   ' Add a new page
   docContents.NewPage 0, 0, IN_NEWPAGE_END, "New Title"
   ' Get number of pages after added new page
   lAfterPageCount = docContents.NumberOfPages
   ' Make sure one new page was added
   If lAfterPageCount = lBeforePageCount Then
      MsgBox "The NewPage method failed." + Chr(13) + "Original page count:" + Str(lBeforePageCount) + Chr(13) + "New page count:" + Str(lAfterPageCount), vbCritical, "Failed"
   Else
      MsgBox " A new page was successfully added at the end." + Chr(13) + "Original page count:" + Str(lBeforePageCount) + Chr(13) + "New page count:" + Str(lAfterPageCount) + Chr(13) + "Page ID of the new page:", vbInformation, "Success"
   End If
   
   ' De-initialize object var
   Set docContents = Nothing
End Sub

Public Sub IDocContents_NewSection()
   ' RobY Oct21/98
   Dim docContents As IDocContents
   Dim lNewObjectID As Long
   
   ' Set object variable for IDocContents interface to doc ctrl object
   Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
   ' Add a new section
   docContents.NewSection 0, 0, IN_NEWPAGE_END, "New Section"
   
   ' De-initialize the object variable
   Set docContents = Nothing
End Sub

Public Sub IDocContents_NumberOfLayers()
   ' RobW May 4/98
   Dim docContents As IDocContents
   Dim lPageID As Long
   
   ' Get the page ID for current page of active document
   lPageID = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
   ' Set object variable for IDocContents interface to doc ctrl object
   Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
   ' Get the number of layers for the current page
   MsgBox "The current page has" + Str(docContents.NumberOfLayers(lPageID)) + " layer(s).", vbInformation
   
   ' De-initialize the object variable
   Set docContents = Nothing
End Sub

Public Sub IDocContents_NumberOfPages()
   ' RobW May 5/98
   Dim docContents As IDocContents
   Dim lPageCount As Integer
   
   ' Set object variable for IDocContents interface to doc ctrl object
   Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
   ' Get number of pages for current doc
   lPageCount = docContents.NumberOfPages
   ' Display number of pages in current document
   MsgBox "Current document has" + Str(lPageCount) + " pages.", vbInformation
   
   ' De-initialize object var
   Set docContents = Nothing
End Sub

Public Sub IDocContents_OpenFile()
   ' RobW May5/98
   ' Modified by RobW May 11/98 - Don't do bind here; user can use ImageX menu item
   Dim docContents As IDocContents
   
   ' Set object variable for IDocContents interface to doc ctrl object
   Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
   ' Show the common dialog so can select document
   MainMDIForm.ActiveForm.CommonDialog1.Flags = cdlOFNHideReadOnly
   MainMDIForm.ActiveForm.CommonDialog1.CancelError = True
   MainMDIForm.ActiveForm.CommonDialog1.DialogTitle = "ImageX - Select a File for the OpenFile Method"
   MainMDIForm.ActiveForm.CommonDialog1.ShowOpen
   ' Open selected file in the doc control
   docContents.OpenFile MainMDIForm.ActiveForm.CommonDialog1.FileName
   ' Change the child window's title to the filename that was opened
   MainMDIForm.ActiveForm.Caption = MainMDIForm.ActiveForm.CommonDialog1.FileName
   MainMDIForm.ActiveForm.CommonDialog1.InitDir = CurDir
   ' Initialize filename to null
   MainMDIForm.ActiveForm.CommonDialog1.FileName = ""
   
   ' De-initialize object var
   Set docContents = Nothing
End Sub

Public Sub IDocContents_PageID()
   ' RobW May5/98
   Dim docContents As IDocContents
   Dim lPageOneID As Long
   
   ' Set object variable for IDocContents interface to doc ctrl object
   Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
   ' Display Page ID for page one of current document
   lPageOneID = docContents.pageID(1)
   If lPageOneID <> 0 Then
      MsgBox "Page 1 ID:" + Str(lPageOneID), vbInformation
   Else
      MsgBox "Page 1 ID:" + Str(lPageOneID), vbCritical
   End If
   
   ' De-initialize object var
   Set docContents = Nothing
End Sub

Public Sub IDocContents_ParentID()
   ' RobW May 11/98
   ' Modified by RobY July8/98
   Dim docContents As IDocContents
   Dim lParentID As Long
   Dim lPageID As Long
   
   ' Get the page ID for current page of active document
   lPageID = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
   ' Set object variable for IDocContents interface to doc ctrl object
   Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
   ' Get Parent ID for acitve page of current doc
   lParentID = docContents.parentID(lPageID)
   ' Display parent ID
   If lParentID <> 0 Then
      MsgBox "Parent ID for active page:" + Str(lParentID), vbInformation
   Else
      MsgBox "Parent ID for active page:" + Str(lParentID), vbCritical
   End If
   
   ' De-initialize object var
   Set docContents = Nothing
End Sub

Public Sub IDocContents_RemoveLayerDialog()
   ' RobW May 1/98
   Dim lPageID As Long
   Dim docContents As IDocContents
   
   ' Get the page ID for current page of active document
   lPageID = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
   ' Set object variable for IDocContents interface to doc ctrl object
   Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
   ' Open the Remove Layer Dialog for page one
   docContents.RemoveLayerDialog lPageID
   
   ' De-initialize the object variable
   Set docContents = Nothing
End Sub

Public Sub IDocContents_RemoveLayerDialogAvailability()
   'RobY June25/98
   ' modified by MarkS Nov30, 1998 - now uses global function for string determination
   ' Modified by RobY Dec3/98 - changed global function to change menu status of ImageX according to what is returned from availablity
   Dim docContents As IDocContents
   Dim iAvailable As COMMAND_AVAILABILITY
   Dim lPageID As Long
   Dim strResult As String
   
   ' Set object variable for IDocContents interface to doc ctrl object
   Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
   ' Set the menu name to change
   Set MenuName = MainMDIForm.ActiveForm.mnu_DocCtrl_IDocContents_Array(24)
   
   lPageID = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
   ' Get the value of availability
   iAvailable = docContents.RemoveLayerDialogAvailability(lPageID)
   ' Change the menu status through function
   Availability iAvailable, MenuName
     
   ' De-initialize object var
   Set docContents = Nothing
End Sub

Public Sub IDocContents_RemovePageDialog()
   ' RobW May 1/98
   Dim docContents As IDocContents
   
   ' Set object variable for IDocContents interface to doc ctrl object
   Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
   ' Open the Remove Page Dialog
   docContents.RemovePageDialog
   
   ' De-initialize the object variable
   Set docContents = Nothing
End Sub

Public Sub IDocContents_RemovePageDialogAvailability()
   'RobY June25/98
   ' modified by MarkS Nov30, 1998 - now uses global function for string determination
   ' Modified by RobY Dec3/98 - changed global function to change menu status of ImageX according to what is returned from availablity
   Dim docContents As IDocContents
   Dim iAvailable As COMMAND_AVAILABILITY
   Dim strResult As String
   
   ' Set object variable for IDocContents interface to doc ctrl object
   Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
   ' Set the menu name to change
   Set MenuName = MainMDIForm.ActiveForm.mnu_DocCtrl_IDocContents_Array(25)
   
   ' Get the value of availability
   iAvailable = docContents.RemovePageDialogAvailability
   ' Change the menu status through function
   Availability iAvailable, MenuName
     
   ' De-initialize object var
   Set docContents = Nothing
End Sub

Public Sub IDocContents_ReorderLayers()
   ' RobW May 11/98
   ' Modified by RobY July8/98
   Dim lPageID As Long
   Dim iNumberOfLayers As Integer
   Dim lBottomLayerID As Long
   Dim docContents As IDocContents
   
   ' Set object variable for IDocContents interface to doc ctrl object
   Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
   ' Get the page ID for page 1 of the active document
   lPageID = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
   ' Get the number of layers on the first page
   iNumberOfLayers = docContents.NumberOfLayers(lPageID)
   ' Get the layerID of the bottom layer on the first page
   lBottomLayerID = docContents.LayerID(lPageID, iNumberOfLayers)
   ' Reorder the layers - move the bottom layer to the top (on page 1)
    docContents.ReorderLayers lBottomLayerID, 1
    MsgBox "Moved the bottom layer to the top (on active page)."
   
   ' De-initialize the object variable
   Set docContents = Nothing
End Sub

Public Sub IDocContents_ReorderLayersDialog()
   ' RobW May 1/98
   Dim lPageID As Long
   Dim docContents As IDocContents
   
   ' Get the page ID for current page of active document
   lPageID = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
   ' Set object variable for IDocContents interface to doc ctrl object
   Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
   ' Open the Reorder Layers Dialog for page one
   docContents.ReorderLayersDialog lPageID
   
   ' De-initialize the object variable
   Set docContents = Nothing
End Sub

Public Sub IDocContents_ReorderLayersDialogAvailability()
   'RobY June25/98
   ' modified by MarkS Nov30, 1998 - now uses global function for string determination
   ' Modified by RobY Dec3/98 - changed global function to change menu status of ImageX according to what is returned from availablity
   Dim docContents As IDocContents
   Dim iAvailable As COMMAND_AVAILABILITY
   Dim lPageID As Long
   Dim strResult As String
   
   ' Set object variable for IDocContents interface to doc ctrl object
   Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
   ' Set the menu name to change
   Set MenuName = MainMDIForm.ActiveForm.mnu_DocCtrl_IDocContents_Array(27)
   
   lPageID = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
   ' Get the value of availability
   iAvailable = docContents.ReorderLayersDialogAvailability(lPageID)
   ' Change the menu status through function
   Availability iAvailable, MenuName
      
   ' De-initialize object var
   Set docContents = Nothing
End Sub

Public Sub IDocContents_ReorderPages()
   ' RobW May 11/98
   ' Modified By RobY June 19/98
   Dim docContents As IDocContents
   Dim lPageIDToMove As Long
   Dim lPageIDTarget As Long
   Dim iPageNumTarget As Integer
   Dim iPageNumToMove As Integer
   Dim iPosition As Integer
   
   iPageNumToMove = InputBox("Please enter page number of page to moved.", "Page Number To Move")
   iPageNumTarget = InputBox("Please enter the target page number", "Target Page Number")
   iPosition = InputBox("Please enter 0 or 1 for position." + Chr(13) + Chr(10) + _
                  "0 = Before Target Page Number" + Chr(10) + Chr(13) + _
                  "1 = After Target Page Number", "Position")
   ' Get the page ID for page number to move
   lPageIDToMove = MainMDIForm.ActiveForm.SpicerDoc1.pageID(iPageNumToMove)
   ' Get the page ID for the target page number
   lPageIDTarget = MainMDIForm.ActiveForm.SpicerDoc1.pageID(iPageNumTarget)
   ' Set object variable for IDocContents interface to doc ctrl object
   Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
   If iPosition = 0 Then
      docContents.ReorderPages lPageIDToMove, lPageIDTarget, IN_NEWPAGE_BEFORE
      MsgBox "Moved page" + Str(iPageNumToMove) + " to before page" + Str(iPageNumTarget) + _
         " of the document."
   ElseIf iPosition = 1 Then
      docContents.ReorderPages lPageIDToMove, lPageIDTarget, IN_NEWPAGE_AFTER
      MsgBox "Moved page " + Str(iPageNumToMove) + " to after page " + Str(iPageNumTarget) + _
         " of the document."
   End If
   
   ' De-initialize object var
   Set docContents = Nothing
End Sub

Public Sub IDocContents_ReorderPagesDialog()
   ' RobW May 1/98
   Dim docContents As IDocContents
   
   ' Set object variable for IDocContents interface to doc ctrl object
   Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
   
   ' Open the Reorder Pages Dialog
   docContents.ReorderPagesDialog
   
   ' De-initialize the object variable
   Set docContents = Nothing
End Sub

Public Sub IDocContents_ReorderPagesDialogAvailability()
   'RobY June25/98
   ' modified by MarkS Nov30, 1998 - now uses global function for string determination
   ' Modified by RobY Dec3/98 - changed global function to change menu status of ImageX according to what is returned from availablity
   Dim docContents As IDocContents
   Dim iAvailable As COMMAND_AVAILABILITY
   Dim strResult As String
   
   ' Set object variable for IDocContents interface to doc ctrl object
   Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
   ' Set the menu name to change
   Set MenuName = MainMDIForm.ActiveForm.mnu_DocCtrl_IDocContents_Array(29)
   
   ' Get the value of availability
   iAvailable = docContents.ReorderPagesDialogAvailability
   ' Change the menu status through function
   Availability iAvailable, MenuName
   
   ' De-initialize object var
   Set docContents = Nothing
End Sub

Public Sub IDocContents_ReplacePageDialog()
   ' RobW May 1/98
   Dim lPageID As Long
   Dim docContents As IDocContents
   
   ' Get the page ID for current page of active document
   lPageID = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
   ' Set object variable for IDocContents interface to doc ctrl object
   Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
   ' Open the Replace Page Dialog, replace page one
   docContents.ReplacePageDialog lPageID
   
   ' De-initialize the object variable
   Set docContents = Nothing
End Sub

Public Sub IDocContents_ReplacePageDialogAvailability()
   'RobY June25/98
   ' modified by MarkS Nov30, 1998 - now uses global function for string determination
   ' Modified by RobY Dec3/98 - changed global function to change menu status of ImageX according to what is returned from availablity
   Dim docContents As IDocContents
   Dim iAvailable As COMMAND_AVAILABILITY
   Dim lPageID As Long
   Dim strResult As String
   
   ' Set object variable for IDocContents interface to doc ctrl object
   Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
   ' Set the menu name to change
   Set MenuName = MainMDIForm.ActiveForm.mnu_DocCtrl_IDocContents_Array(30)
   
   lPageID = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
   ' Get the value of availability
   iAvailable = docContents.ReplacePageDialogAvailability(lPageID)
   ' Change the menu status through function
   Availability iAvailable, MenuName
   
   ' De-initialize object var
   Set docContents = Nothing
End Sub

Public Sub IDocContents_RootID()
   ' RobW May 5/98
   Dim docContents As IDocContents
   Dim lRootID As Long
   
   ' Set object variable for IDocContents interface to doc ctrl object
   Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
   ' Display root ID for current document
   lRootID = docContents.RootID
   If lRootID <> 0 Then
      MsgBox "Root ID:" + Str(lRootID), vbInformation
   Else
      MsgBox "Root ID:" + Str(lRootID), vbCritical
   End If
   
   ' De-initialize object var
   Set docContents = Nothing
End Sub

Public Sub IDocContents_Unload()
   ' RobW May 11/98
   Dim docContents As IDocContents
   
   ' Set object variable for IDocContents interface to doc ctrl object
   Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
   ' Unload current document
   docContents.Unload docContents.RootID
   MsgBox "Unloaded the document."
   
   ' De-initialize object var
   Set docContents = Nothing
End Sub

' <EOF IDocContents.bas>





