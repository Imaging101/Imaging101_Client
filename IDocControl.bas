Attribute VB_Name = "mod_IDocControl"
' File:      IDocControl.bas
' Created:   1998Apr28 by Rob Wood, Mark Simpson, and Rob Young
' Purpose:   To test the Spicer Doc Control's IDocControl interface.
' Revisions:
'   $log$
'   $nokeywords$
'
' Copyright (C) 1998 SPICER Corporation, All rights reserved.

Option Explicit

Public Sub IDocControl_AboutBox()
   ' MarkS & RobW Apr27/98
   ' Display About Box
   ChildForm1.SpicerDoc1.AboutBox
End Sub

Public Sub IDocControl_FirstPageID()
   ' MarkS & RobW Apr27/98
   ' Display the FirstPageID
   Dim lFirstPageID As Long
   lFirstPageID = ChildForm1.SpicerDoc1.FirstPageID
   If lFirstPageID <> 0 Then
      MsgBox "First Page ID:" + Str(lFirstPageID), vbInformation
   Else
      MsgBox "First Page ID:" + Str(lFirstPageID), vbCritical
   End If
End Sub

Public Sub IDocControl_NewDocument()
   ' MarkS & RobW Apr27/98
   ' Modified by RobW May 8/98
   ' Modified by RobW May 11/98 - Don't do bind here; user can use ImageX menu item
   Dim lRootID As Long
   ' Create new document
   ChildForm1.SpicerDoc1.NewDocument
   ' Set child window's title
   ChildForm1.Caption = "New Document"
   ' Display new document's root ID / document ID
   lRootID = ChildForm1.SpicerDoc1.RootID
   If lRootID <> 0 Then
      MsgBox "New document Root ID:" + Str(lRootID), vbInformation
   Else
      MsgBox "New document Root ID:" + Str(lRootID), vbCritical
   End If
End Sub

Public Sub IDocControl_OpenFile()
   ' MarkS & RobW Apr27/98
   ' Show the common dialog so can select document
   ' Modified by RobW May 11/98 - Don't do bind here; user can use ImageX menu item
   ChildForm1.CommonDialog1.CancelError = True
   ChildForm1.CommonDialog1.Flags = cdlOFNHideReadOnly
   ChildForm1.CommonDialog1.DialogTitle = "ImageX - Select a File for the OpenFile Method"
   ChildForm1.CommonDialog1.ShowOpen
   ' Open selected file in the doc control
   ChildForm1.SpicerDoc1.OpenFile ChildForm1.CommonDialog1.FileName
   ' Change the child window's title to the filename that was opened
   ChildForm1.Caption = ChildForm1.CommonDialog1.FileName
End Sub

Public Sub IDocControl_LayerID()
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
   lPageID = ChildForm1.SpicerDoc1.pageID(iPageNum)
   ' Set object variable for IDocContents interface to doc ctrl object
   Set docContents = ChildForm1.SpicerDoc1.object
   ' Get number of layers for current page of current doc
   lLayerCount = docContents.NumberOfLayers(lPageID)
   sTemp = "The current page has" + Str(lLayerCount) + " layer(s)." + Chr(13)
   ' Get Layer ID for all layers on current page of current doc
   For iTemp = 1 To lLayerCount
      sTemp = sTemp + "Layer" + Str(iTemp) + " ID: " + Str(ChildForm1.SpicerDoc1.LayerID(lPageID, iTemp)) + Chr(13)
   Next iTemp
   ' Display number of layers and ID's for each
   MsgBox sTemp, vbInformation
   ' De-initialize object var
   Set docContents = Nothing
End Sub

Public Sub IDocControl_PageID()
   ' RobW Apr30/98
   ' Modified by RobW May 8/98
   ' Modified by RobY July6/98
   Dim lPageOneID As Long
   Dim iPageNum As Integer
   ' Prompt for page number
   iPageNum = InputBox("Please enter the page number that you want the page id for.", "Page Number")
   ' Get the page id for the page number entered
   lPageOneID = ChildForm1.SpicerDoc1.pageID(iPageNum)
   ' Display Page ID for the current document
   If lPageOneID <> 0 Then
      MsgBox "Page" + Str(iPageNum) + " ID:" + Str(lPageOneID), vbInformation
   Else
      MsgBox "Page" + Str(iPageNum) + " ID:" + Str(lPageOneID), vbCritical
   End If
End Sub

Public Sub IDocControl_RootID()
   ' RobW Apr30/98
   Dim lRootID As Long
   ' Display new document's root ID / document ID
   lRootID = ChildForm1.SpicerDoc1.RootID
   If lRootID <> 0 Then
      MsgBox "Root ID:" + Str(lRootID), vbInformation
   Else
      MsgBox "Root ID:" + Str(lRootID), vbCritical
   End If
End Sub
      
' <EOF IDocControl.bas>
