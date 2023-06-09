Attribute VB_Name = "mod_IVectorEdit"
' File:      IVectorEdit.bas
' Created:   1998July30 by Rob Young
' Purpose:   To test the Spicer Markup Control's IVectorEdit interface.
' Revisions: RobY Jan4/99  - Added functionality to open more than 1 file. ChildForm1 is no longer used access objects because it has been set to ChildForm() which is an array of ChildForm1's forms.
'                          - Anywhere ChildForm1 was used has been changed to ChildForm(iIndex) or MainMDIForm.ActiveForm.
'   $log$
'   $nokeywords$
'
' Copyright (C) 1998 SPICER Corporation, All rights reserved.
Option Explicit
Dim MenuName As Menu

Public Sub IVectorEdit_BindToViewControl()
   ' RobY Aug12/98
   Dim VectorEdit As IVectorEdit
   
   ' Set object variable for IVectorEdit interface to Markup ctrl object
   Set VectorEdit = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   'Bind the Spicer Markup Control to the Spicer View Control
   VectorEdit.BindToViewControl MainMDIForm.ActiveForm.SpicerView1.object
   
   ' De-initialize the object variable
   Set VectorEdit = Nothing
End Sub

Public Sub IVectorEdit_Copy()
   ' RobY Aug12/98
   ' Modified by RobY Oct22/98 - Name of command was changed to Copy
   Dim VectorEdit As IVectorEdit
   
   ' Set object variable for IVectorEdit interface to Markup ctrl object
   Set VectorEdit = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Copy the objects
   VectorEdit.Copy
   
   ' De-initialize the object variable
   Set VectorEdit = Nothing
End Sub

Public Sub IVectorEdit_Cut()
   ' RobY Aug12/98
   ' Modified by RobY Oct22/98 - Name of command was changed to Cut
   Dim VectorEdit As IVectorEdit
   
   ' Set object variable for IVectorEdit interface to Markup ctrl object
   Set VectorEdit = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Cut the objects
   VectorEdit.Cut
   
   ' De-initialize the object variable
   Set VectorEdit = Nothing
End Sub

Public Sub IVectorEdit_Paste()
   ' RobY Oct22/98
   Dim VectorEdit As IVectorEdit
   
   ' Set object variable for IVectorEdit interface to Markup ctrl object
   Set VectorEdit = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Paste the objects
   VectorEdit.Paste
   
   ' De-initialize the object variable
   Set VectorEdit = Nothing
End Sub

Public Sub IVectorEdit_Redo()
   ' RobY Aug12/98
   Dim VectorEdit As IVectorEdit
   
   ' Set object variable for IVectorEdit interface to Markup ctrl object
   Set VectorEdit = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Redo what was undone
   VectorEdit.Redo
   
   ' De-initialize the object variable
   Set VectorEdit = Nothing
End Sub

Public Sub IVectorEdit_Undo()
   ' RobY Aug12/98
   Dim VectorEdit As IVectorEdit
   
   ' Set object variable for IVectorEdit interface to Markup ctrl object
   Set VectorEdit = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Redo what was undone
   VectorEdit.Undo
   
   ' De-initialize the object variable
   Set VectorEdit = Nothing
End Sub

Public Sub IVectorEdit_TextSearchDialog()
   
   
   ' RobY Mar11/99
   Dim VectorEdit As IVectorEdit
   
    On Error GoTo ErrorOccurred
    
   ' Set object variable for IVectorEdit interface to Markup ctrl object
   Set VectorEdit = MainMDIForm.ActiveForm.SpicerDoc1.object
   
   If VectorEdit.TextSearchDialogAvailability Then
    ' Display the Find dialog
      VectorEdit.TextSearchDialog
    End If
    
   ' De-initialize the object variable
   Set VectorEdit = Nothing
   
Exit Sub

ErrorOccurred:
    
    funcQuickMessage "SHOW", "ERROR - IVectorEdit_TextSearchDialog():  Error " & Err.Number & "  - " & Err.Description
    
End Sub

Public Sub IVectorEdit_CopyAvailability()
   ' RobY Dec11/98
   Dim VectorEdit As IVectorEdit
   Dim iAvailable As COMMAND_AVAILABILITY
   
   ' Set object variable for IHotSpots interface to Markup ctrl object
   Set VectorEdit = MainMDIForm.ActiveForm.SpicerMarkup1.object
   ' Set the menu name to change
   Set MenuName = MainMDIForm.ActiveForm.mnu_Markup_IVectorEdit_Array(1)
   
   ' Get the value of availability
   iAvailable = VectorEdit.CopyAvailability
   ' Change the menu status through function
   Availability iAvailable, MenuName
   
   ' De-initialize object var
   Set VectorEdit = Nothing
End Sub
Public Sub IVectorEdit_CutAvailability()
   ' RobY Dec11/98
   Dim VectorEdit As IVectorEdit
   Dim iAvailable As COMMAND_AVAILABILITY
   
   ' Set object variable for IHotSpots interface to Markup ctrl object
   Set VectorEdit = MainMDIForm.ActiveForm.SpicerMarkup1.object
   ' Set the menu name to change
   Set MenuName = MainMDIForm.ActiveForm.mnu_Markup_IVectorEdit_Array(2)
   
   ' Get the value of availability
   iAvailable = VectorEdit.CutAvailability
   ' Change the menu status through function
   Availability iAvailable, MenuName
   
   ' De-initialize object var
   Set VectorEdit = Nothing
End Sub

Public Sub IVectorEdit_PasteAvailability()
   ' RobY Dec11/98
   Dim VectorEdit As IVectorEdit
   Dim iAvailable As COMMAND_AVAILABILITY
   
   ' Set object variable for IHotSpots interface to Markup ctrl object
   Set VectorEdit = MainMDIForm.ActiveForm.SpicerMarkup1.object
   ' Set the menu name to change
   Set MenuName = MainMDIForm.ActiveForm.mnu_Markup_IVectorEdit_Array(3)
   
   ' Get the value of availability
   iAvailable = VectorEdit.PasteAvailability
   ' Change the menu status through function
   Availability iAvailable, MenuName
   
   ' De-initialize object var
   Set VectorEdit = Nothing
End Sub

Public Sub IVectorEdit_RedoAvailability()
   ' RobY Dec11/98
   Dim VectorEdit As IVectorEdit
   Dim iAvailable As COMMAND_AVAILABILITY
   
   ' Set object variable for IHotSpots interface to Markup ctrl object
   Set VectorEdit = MainMDIForm.ActiveForm.SpicerMarkup1.object
   ' Set the menu name to change
   Set MenuName = MainMDIForm.ActiveForm.mnu_Markup_IVectorEdit_Array(4)
   
   ' Get the value of availability
   iAvailable = VectorEdit.RedoAvailability
   ' Change the menu status through function
   Availability iAvailable, MenuName
   
   ' De-initialize object var
   Set VectorEdit = Nothing
End Sub

Public Sub IVectorEdit_UndoAvailability()
   ' RobY Dec11/98
   Dim VectorEdit As IVectorEdit
   Dim iAvailable As COMMAND_AVAILABILITY
   
   ' Set object variable for IHotSpots interface to Markup ctrl object
   Set VectorEdit = MainMDIForm.ActiveForm.SpicerMarkup1.object
   ' Set the menu name to change
   Set MenuName = MainMDIForm.ActiveForm.mnu_Markup_IVectorEdit_Array(5)
   
   ' Get the value of availability
   iAvailable = VectorEdit.UndoAvailability
   ' Change the menu status through function
   Availability iAvailable, MenuName
   
   ' De-initialize object var
   Set VectorEdit = Nothing
End Sub

