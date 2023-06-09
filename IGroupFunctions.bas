Attribute VB_Name = "mod_IGroupFunctions"
' File:      IGroupFunctions.bas
' Created:   1998July30 by Rob Young
' Purpose:   To test the Spicer Markup Control's IGroupFunctions interface.
' Revisions: RobY Jan4/99  - Added functionality to open more than 1 file. ChildForm1 is no longer used access objects because it has been set to ChildForm() which is an array of ChildForm1's forms.
'                          - Anywhere ChildForm1 was used has been changed to ChildForm(iIndex) or MainMDIForm.ActiveForm.
'   $log$
'   $nokeywords$
'
' Copyright (C) 1998 SPICER Corporation, All rights reserved.
Option Explicit
Dim MenuName As Menu

Public Sub IGroupFunctions_BindSelectedObjects()
   'RobY Aug12/98
   Dim GroupFunctions As IGroupFunctions
   
   ' Set object variable for IGroupFunctions interface to Markup ctrl object
   Set GroupFunctions = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   'Bind the selected objects
   GroupFunctions.BindSelectedObjects
   
   ' De-initialize the object variable
   Set GroupFunctions = Nothing
End Sub

Public Sub IGroupFunctions_BindToViewControl()
   ' RobY Aug12/98
   Dim GroupFunctions As IGroupFunctions
   
   ' Set object variable for IGroupFunctions interface to Markup ctrl object
   Set GroupFunctions = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   'Bind the Spicer Markup Control to the Spicer View Control
   GroupFunctions.BindToViewControl MainMDIForm.ActiveForm.SpicerView1.object
   
   ' De-initialize the object variable
   Set GroupFunctions = Nothing
End Sub

Public Sub IGroupFunctions_DeleteSelectedObjects()
   ' RobY Aug12/98
   Dim GroupFunctions As IGroupFunctions
   
   ' Set object variable for IGroupFunctions interface to Markup ctrl object
   Set GroupFunctions = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Deselect the selected objects
   GroupFunctions.DeleteSelectedObjects
   
   ' De-initialize the object variable
   Set GroupFunctions = Nothing
End Sub

Public Sub IGroupFunctions_DeselectAll()
   ' RobY Aug12/98
   Dim GroupFunctions As IGroupFunctions
   
   ' Set object variable for IGroupFunctions interface to Markup ctrl object
   Set GroupFunctions = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Deselect all objects
   GroupFunctions.DeselectAll
   
   ' De-initialize the object variable
   Set GroupFunctions = Nothing
End Sub

Public Sub IGroupFunctions_SaveSymbol()
   ' RobY Nov16/98  Method moved from IDocSave interface. Moved and modified code to reflect changes. No longer needs docwinid.
   Dim GroupFunctions As IGroupFunctions
   Dim strFilename As String
   
   ' Set object variable for IGroupFunctions interface to Markup ctrl object
   Set GroupFunctions = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   'Prompt the user for filename to save symbol as
   strFilename = InputBox("Specify a name to save the file as.", "Save Symbol")
   'Save as a symbol
   GroupFunctions.SaveSymbol strFilename, "Symbol Test"
   
   'De-initialize the object var
   Set GroupFunctions = Nothing
End Sub

Public Sub IGroupFunctions_SelectAllObjectsOnLayer()
   ' RobY Aug12/98
   Dim GroupFunctions As IGroupFunctions
   
   ' Set object variable for IGroupFunctions interface to Markup ctrl object
   Set GroupFunctions = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Selected all objects on active layer
   GroupFunctions.SelectAllObjectsOnLayer
   
   ' De-initialize the object variable
   Set GroupFunctions = Nothing
End Sub

Public Sub IGroupFunctions_UnbindSelectedGroupObject()
   ' RobY Aug12/98
   Dim GroupFunctions As IGroupFunctions
   
   ' Set object variable for IGroupFunctions interface to Markup ctrl object
   Set GroupFunctions = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Unbind selected group objects
   GroupFunctions.UnbindSelectedGroupObject
   
   ' De-initialize the object variable
   Set GroupFunctions = Nothing
End Sub

Public Sub IGroupFunctions_BindSelectedObjectsAvailability()
   'RobY Dec11/98
   Dim GroupFunctions As IGroupFunctions
   Dim iAvailable As COMMAND_AVAILABILITY
   
   ' Set object variable for IGroupFunctions interface to Markup ctrl object
   Set GroupFunctions = MainMDIForm.ActiveForm.SpicerMarkup1.object
   ' Set the menu name to change
   Set MenuName = MainMDIForm.ActiveForm.mnu_Markup_IGroupFunctions_Array(0)
   
   ' Get the value of availability
   iAvailable = GroupFunctions.BindSelectedObjectsAvailability
   ' Change the menu status through function
   Availability iAvailable, MenuName
   
   ' De-initialize object var
   Set GroupFunctions = Nothing
End Sub

Public Sub IGroupFunctions_DeleteSelectedObjectsAvailability()
   'RobY Dec11/98
   Dim GroupFunctions As IGroupFunctions
   Dim iAvailable As COMMAND_AVAILABILITY
   
   ' Set object variable for IGroupFunctions interface to Markup ctrl object
   Set GroupFunctions = MainMDIForm.ActiveForm.SpicerMarkup1.object
   ' Set the menu name to change
   Set MenuName = MainMDIForm.ActiveForm.mnu_Markup_IGroupFunctions_Array(2)
   
   ' Get the value of availability
   iAvailable = GroupFunctions.DeleteSelectedObjectsAvailability
   ' Change the menu status through function
   Availability iAvailable, MenuName
   
   ' De-initialize object var
   Set GroupFunctions = Nothing
End Sub

Public Sub IGroupFunctions_DeselectAllAvailability()
   'RobY Dec11/98
   Dim GroupFunctions As IGroupFunctions
   Dim iAvailable As COMMAND_AVAILABILITY
   
   ' Set object variable for IGroupFunctions interface to Markup ctrl object
   Set GroupFunctions = MainMDIForm.ActiveForm.SpicerMarkup1.object
   ' Set the menu name to change
   Set MenuName = MainMDIForm.ActiveForm.mnu_Markup_IGroupFunctions_Array(3)
   
   ' Get the value of availability
   iAvailable = GroupFunctions.DeselectAllAvailability
   ' Change the menu status through function
   Availability iAvailable, MenuName
   
   ' De-initialize object var
   Set GroupFunctions = Nothing
End Sub

Public Sub IGroupFunctions_SelectAllObjectsOnLayerAvailability()
   'RobY Dec11/98
   Dim GroupFunctions As IGroupFunctions
   Dim iAvailable As COMMAND_AVAILABILITY
   
   ' Set object variable for IGroupFunctions interface to Markup ctrl object
   Set GroupFunctions = MainMDIForm.ActiveForm.SpicerMarkup1.object
   ' Set the menu name to change
   Set MenuName = MainMDIForm.ActiveForm.mnu_Markup_IGroupFunctions_Array(5)
   
   ' Get the value of availability
   iAvailable = GroupFunctions.SelectAllObjectsOnLayerAvailability
   ' Change the menu status through function
   Availability iAvailable, MenuName
   
   ' De-initialize object var
   Set GroupFunctions = Nothing
End Sub

Public Sub IGroupFunctions_UnbindSelectedGroupObjectAvailability()
   'RobY Dec11/98
   Dim GroupFunctions As IGroupFunctions
   Dim iAvailable As COMMAND_AVAILABILITY
   
   ' Set object variable for IGroupFunctions interface to Markup ctrl object
   Set GroupFunctions = MainMDIForm.ActiveForm.SpicerMarkup1.object
   ' Set the menu name to change
   Set MenuName = MainMDIForm.ActiveForm.mnu_Markup_IGroupFunctions_Array(6)
   
   ' Get the value of availability
   iAvailable = GroupFunctions.UnbindSelectedGroupObjectAvailability
   ' Change the menu status through function
   Availability iAvailable, MenuName
   
   ' De-initialize object var
   Set GroupFunctions = Nothing
End Sub



