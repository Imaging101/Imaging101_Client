Attribute VB_Name = "mod_ISpicerDoc"
' File:      ISpicerDoc.bas
' Created:   1998Apr28 by Rob Wood, Mark Simpson, and Rob Young
' Purpose:   To test the Spicer Doc Control's ISpicerDoc interface.
' Revisions: RobY Jan4/99  - Added functionality to open more than 1 file. ChildForm1 is no longer used access objects because it has been set to ChildForm() which is an array of ChildForm1's forms.
'                          - Anywhere ChildForm1 was used has been changed to ChildForm(iIndex) or MainMDIForm.ActiveForm.
'   $log$
'   $nokeywords$
'
' Copyright (C) 1998 SPICER Corporation, All rights reserved.

Option Explicit

Public Sub ISpicerDoc_AboutBox()
   ' MarkS & RobW Apr27/98
   ' Display About Box
   MainMDIForm.ActiveForm.SpicerDoc1.AboutBox
End Sub

Public Sub ISpicerDoc_CloseDocument()
   ' RobY Sept18/98
   ' MarkS Nov9,1998 Added changing of window title after closing of a document
   ' Modified by RobY Dec15 - New boolean parameter added to command for saving modified documents.
   Dim iResponse As Integer
   
   ' Prompt user if to save before closing
   iResponse = MsgBox("Do you want to save the document before closing?", vbYesNo + vbQuestion, "Save Document")
   If iResponse = vbYes Then
      ' Save before closing document if any changes made
      MainMDIForm.ActiveForm.SpicerDoc1.CloseDocument True
   Else
      ' Close the current document without saving
      MainMDIForm.ActiveForm.SpicerDoc1.CloseDocument False
   End If
   
   ' Change window's caption as doc is no longer present
   MainMDIForm.ActiveForm.Caption = "Document Window"
End Sub
Public Sub ISpicerDoc_FirstPageID()
   ' MarkS & RobW Apr27/98
   ' Display the FirstPageID
   Dim lFirstPageID As Long
   
   lFirstPageID = MainMDIForm.ActiveForm.SpicerDoc1.FirstPageID
   If lFirstPageID <> 0 Then
      MsgBox "First Page ID:" + Str(lFirstPageID), vbInformation
   Else
      MsgBox "First Page ID:" + Str(lFirstPageID), vbCritical
   End If
End Sub

Public Sub ISpicerDoc_NewDocument()
   ' MarkS & RobW Apr27/98
   ' Modified by RobW May 8/98
   ' Modified by RobW May 11/98 - Don't do bind here; user can use ImageX menu item
   Dim lRootID As Long
      
   ' Create new document
   MainMDIForm.ActiveForm.SpicerDoc1.NewDocument
   ' Set child window's title
   MainMDIForm.ActiveForm.Caption = "New Document"
   ' Display new document's root ID / document ID
   lRootID = MainMDIForm.ActiveForm.SpicerDoc1.RootID
   If lRootID <> 0 Then
      MsgBox "New document Root ID:" + Str(lRootID), vbInformation
   Else
      MsgBox "New document Root ID:" + Str(lRootID), vbCritical
   End If
End Sub

Public Sub ISpicerDoc_NewestObjectID()
   ' RobY Sept18/98
   Dim lNewObjectID As Long
   
   ' Get the ID of the most recently created page or layer.
   lNewObjectID = MainMDIForm.ActiveForm.SpicerDoc1.NewestObjectID
   MsgBox "Newest Object ID: " + Str(lNewObjectID), vbInformation, "Newest Object ID"
End Sub

Public Sub ISpicerDoc_OpenFile()
   ' MarkS & RobW Apr27/98
   ' Show the common dialog so can select document
   ' Modified by RobW May 11/98 - Don't do bind here; user can use ImageX menu item
   ' Modified by MarkS Nov9,1998 - Only change window title if file loads properly
   ' Modified by RobY Dec1/98 - Removed return value for OpenFile and if statement associated with it. No longer valid.
   
   MainMDIForm.ActiveForm.CommonDialog1.CancelError = True
   MainMDIForm.ActiveForm.CommonDialog1.Flags = cdlOFNHideReadOnly
   MainMDIForm.ActiveForm.CommonDialog1.DialogTitle = "ImageX - Select a File for the OpenFile Method"
   MainMDIForm.ActiveForm.CommonDialog1.ShowOpen
   ' Open selected file in the doc control
   MainMDIForm.ActiveForm.SpicerDoc1.OpenFile (MainMDIForm.ActiveForm.CommonDialog1.FileName)
   ' Change the child window's title to the filename if it was opened successfully
   MainMDIForm.ActiveForm.Caption = MainMDIForm.ActiveForm.CommonDialog1.FileName
   MainMDIForm.ActiveForm.CommonDialog1.InitDir = CurDir
   ' Initialize filename to null
   MainMDIForm.ActiveForm.CommonDialog1.FileName = ""
End Sub

Public Sub ISpicerDoc_LayerID()
   'RobW May 8/98
   'Modified by RobY July6/98
   Dim lPageID As Long
   Dim iPageNum As Integer
   Dim lLayerCount As Long
   Dim iTemp As Integer
   Dim sTemp As String
   Dim docContents As IDocContents
   
   ' Get the page number of the page that has the layer ids you want
   iPageNum = InputBox("Please enter the page number that you want to get the layer ids for.", "Page Number")
   ' Get the page id of page number entered
   lPageID = MainMDIForm.ActiveForm.SpicerDoc1.pageID(iPageNum)
   ' Set object variable for IDocContents interface to doc ctrl object
   Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
   ' Get number of layers for current page of current doc
   lLayerCount = docContents.NumberOfLayers(lPageID)
   sTemp = "The page has" + Str(lLayerCount) + " layer(s)." + Chr(13)
   ' Get Layer ID for all layers on current page of current doc
   For iTemp = 1 To lLayerCount
      sTemp = sTemp + "Layer" + Str(iTemp) + " ID: " + Str(MainMDIForm.ActiveForm.SpicerDoc1.LayerID(lPageID, iTemp)) + Chr(13)
   Next iTemp
   ' Display number of layers and ID's for each
   MsgBox sTemp, vbInformation
   
   ' De-initialize object var
   Set docContents = Nothing
End Sub

Public Sub ISpicerDoc_PageID()
   ' RobW Apr30/98
   ' Modified by RobW May 8/98
   ' Modified by RobY July6/98
   Dim lPageOneID As Long
   Dim iPageNum As Integer
   
   ' Prompt for page number
   iPageNum = InputBox("Please enter the page number that you want the page id for.", "Page Number")
   ' Get the page id for the page number entered
   lPageOneID = MainMDIForm.ActiveForm.SpicerDoc1.pageID(iPageNum)
   ' Display Page ID for the current document
   If lPageOneID <> 0 Then
      MsgBox "Page" + Str(iPageNum) + " ID:" + Str(lPageOneID), vbInformation
   Else
      MsgBox "Page" + Str(iPageNum) + " ID:" + Str(lPageOneID), vbCritical
   End If
End Sub

Public Sub ISpicerDoc_RootID()
   ' RobW Apr30/98
   Dim lRootID As Long

   ' Display new document's root ID / document ID
   lRootID = MainMDIForm.ActiveForm.SpicerDoc1.RootID
   If lRootID <> 0 Then
      MsgBox "Root ID:" + Str(lRootID), vbInformation
   Else
      MsgBox "Root ID:" + Str(lRootID), vbCritical
   End If
End Sub
      
' <EOF ISpicerDoc.bas>

