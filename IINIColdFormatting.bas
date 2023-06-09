Attribute VB_Name = "mod_IINIColdFormatting"
' File:      IINIColdFormatting.bas
' Created:   1998Nov30 by Rob Young
' Purpose:   To test the Spicer Configuration Control's IINIColdFormatting interface.
' Revisions: RobY Jan4/99  - Added functionality to open more than 1 file. ChildForm1 is no longer used access objects because it has been set to ChildForm() which is an array of ChildForm1's forms.
'                          - Anywhere ChildForm1 was used has been changed to ChildForm(iIndex) or MainMDIForm.ActiveForm.
'   $log$
'   $nokeywords$
'
' Copyright (C) 1998 SPICER Corporation, All rights reserved.
Option Explicit

Public Sub IINIColdFormatting_CharactersPerInch()
   'RobY Dec2/98
   Dim ColdFormat As IINIColdFormatting
   Dim iGetCurValue As Integer
   Dim iNewValue As Integer
   Dim iResponse As Integer
   Dim iGetNewValue As Integer
   
   ' Set object variable for IINIColdFormatting interface to Configuration ctrl object
   Set ColdFormat = MainMDIForm.ActiveForm.SpicerConfiguration1.object
   
   iGetCurValue = ColdFormat.CharactersPerInch
   iResponse = MsgBox("Currently set to " + Str(iGetCurValue) + "." + Chr(13) + "Do you want to change it?", vbYesNo + vbInformation, "Return Value")
   If iResponse = vbYes Then
      iNewValue = InputBox("Please enter new value.", "New Value")
      ColdFormat.CharactersPerInch = iNewValue
      iGetNewValue = ColdFormat.CharactersPerInch
      MsgBox "Old Value = " + Str(iGetCurValue) + Chr(13) + "New Value = " + Str(iGetNewValue), vbInformation, "Changed"
   Else
      MsgBox "Value was not changed.", vbInformation, "No Change"
   End If
   
   ' De-initialize the object variable
   Set ColdFormat = Nothing
End Sub

Public Sub IINIColdFormatting_ColdOrientation()
   'RobY Dec2/98
   Dim ColdFormat As IINIColdFormatting
   Dim iGetCurOrientation As Orientation
   Dim iNewOrientation As Orientation
   Dim iGetNewOrientation As Orientation
   Dim iResponse As Integer
   
   ' Set object variable for IINIColdFormatting interface to Configuration ctrl object
   Set ColdFormat = MainMDIForm.ActiveForm.SpicerConfiguration1.object
   
   iGetCurOrientation = ColdFormat.ColdOrientation
   iResponse = MsgBox("Orientation currently set to " + Str(iGetCurOrientation) + "." + Chr(13) + "Do you want to change it?", vbYesNo + vbInformation, "Return Value")
   If iResponse = vbYes Then
      iNewOrientation = InputBox("Please enter new value.", "New Value")
      ColdFormat.ColdOrientation = iNewOrientation
      iGetNewOrientation = ColdFormat.ColdOrientation
      MsgBox "Old Value = " + Str(iGetCurOrientation) + Chr(13) + "New Value = " + Str(iGetNewOrientation), vbInformation, "Changed"
   Else
      MsgBox "Value was not changed.", vbInformation, "No Change"
   End If
      
   ' De-initialize the object variable
   Set ColdFormat = Nothing
End Sub

Public Sub IINIColdFormatting_LeftOffset()
   'RobY Dec2/98
   Dim ColdFormat As IINIColdFormatting
   Dim dGetCurValue As Double
   Dim dNewValue As Double
   Dim iResponse As Integer
   Dim dGetNewValue As Double
   
   ' Set object variable for IINIColdFormatting interface to Configuration ctrl object
   Set ColdFormat = MainMDIForm.ActiveForm.SpicerConfiguration1.object
   
   dGetCurValue = ColdFormat.LeftOffset
   iResponse = MsgBox("Currently set to " + Str(dGetCurValue) + "." + Chr(13) + "Do you want to change it?", vbYesNo + vbInformation, "Return Value")
   If iResponse = vbYes Then
      dNewValue = InputBox("Please enter new value.", "New Value")
      ColdFormat.LeftOffset = dNewValue
      dGetNewValue = ColdFormat.LeftOffset
      MsgBox "Old Value = " + Str(dGetCurValue) + Chr(13) + "New Value = " + Str(dGetNewValue), vbInformation, "Changed"
   Else
      MsgBox "Value was not changed.", vbInformation, "No Change"
   End If
   
   ' De-initialize the object variable
   Set ColdFormat = Nothing
End Sub

Public Sub IINIColdFormatting_LinesPerInch()
   'RobY Dec2/98
   Dim ColdFormat As IINIColdFormatting
   Dim iGetCurValue As Integer
   Dim iNewValue As Integer
   Dim iResponse As Integer
   Dim iGetNewValue As Integer
   
   ' Set object variable for IINIColdFormatting interface to Configuration ctrl object
   Set ColdFormat = MainMDIForm.ActiveForm.SpicerConfiguration1.object
   
   iGetCurValue = ColdFormat.LinesPerInch
   iResponse = MsgBox("Currently set to " + Str(iGetCurValue) + "." + Chr(13) + "Do you want to change it?", vbYesNo + vbInformation, "Return Value")
   If iResponse = vbYes Then
      iNewValue = InputBox("Please enter new value.", "New Value")
      ColdFormat.LinesPerInch = iNewValue
      iGetNewValue = ColdFormat.LinesPerInch
      MsgBox "Old Value = " + Str(iGetCurValue) + Chr(13) + "New Value = " + Str(iGetNewValue), vbInformation, "Changed"
   Else
      MsgBox "Value was not changed.", vbInformation, "No Change"
   End If
   
   ' De-initialize the object variable
   Set ColdFormat = Nothing
End Sub

Public Sub IINIColdFormatting_OverlayFilename()
   'RobY Dec2/98
   Dim ColdFormat As IINIColdFormatting
   Dim strFilename As String
   Dim strNewFilename As String
   Dim iResponse As Integer
   Dim strGetNewFilename As String
   
   ' Set object variable for IINIColdFormatting interface to Configuration ctrl object
   Set ColdFormat = MainMDIForm.ActiveForm.SpicerConfiguration1.object
   
   strFilename = ColdFormat.OverlayFilename
   iResponse = MsgBox("Currently set to " + strFilename + "." + Chr(13) + "Do you want to change it?", vbYesNo + vbInformation, "Return Value")
   If iResponse = vbYes Then
      strFilename = InputBox("Please enter new path and filename.", "New Value")
      ColdFormat.OverlayFilename = strNewFilename
      strGetNewFilename = ColdFormat.OverlayFilename
      MsgBox "Old Value = " + strFilename + Chr(13) + "New Value = " + strGetNewFilename, vbInformation, "Changed"
   Else
      MsgBox "Value was not changed.", vbInformation, "No Change"
   End If
   
   ' De-initialize the object variable
   Set ColdFormat = Nothing
End Sub

Public Sub IINIColdFormatting_OverlayType()
   'RobY Dec2/98
   Dim ColdFormat As IINIColdFormatting
   Dim iGetOverLayType As OVERLAY_TYPE
   Dim iNewOverLayType As OVERLAY_TYPE
   Dim iResponse As Integer
   Dim iGetNewOverLayType As OVERLAY_TYPE
   
   ' Set object variable for IINIColdFormatting interface to Configuration ctrl object
   Set ColdFormat = MainMDIForm.ActiveForm.SpicerConfiguration1.object
   
   iGetOverLayType = ColdFormat.OverlayType
   iResponse = MsgBox("Currently set to " + Str(iGetOverLayType) + "." + Chr(13) + "Do you want to change it?", vbYesNo + vbInformation, "Return Value")
   If iResponse = vbYes Then
      iNewOverLayType = InputBox("Please enter new value.", "New Value")
      ColdFormat.OverlayType = iNewOverLayType
      iGetNewOverLayType = ColdFormat.OverlayType
      MsgBox "Old Value = " + Str(iGetOverLayType) + Chr(13) + "New Value = " + Str(iGetNewOverLayType), vbInformation, "Changed"
   Else
      MsgBox "Value was not changed.", vbInformation, "No Change"
   End If
   
   ' De-initialize the object variable
   Set ColdFormat = Nothing
End Sub

Public Sub IINIColdFormatting_TopOffset()
   'RobY Dec2/98
   Dim ColdFormat As IINIColdFormatting
   Dim dGetCurValue As Double
   Dim dNewValue As Double
   Dim iResponse As Integer
   Dim dGetNewValue As Double
   
   ' Set object variable for IINIColdFormatting interface to Configuration ctrl object
   Set ColdFormat = MainMDIForm.ActiveForm.SpicerConfiguration1.object
   
   dGetCurValue = ColdFormat.TopOffset
   iResponse = MsgBox("Currently set to " + Str(dGetCurValue) + "." + Chr(13) + "Do you want to change it?", vbYesNo + vbInformation, "Return Value")
   If iResponse = vbYes Then
      dNewValue = InputBox("Please enter new value.", "New Value")
      ColdFormat.TopOffset = dNewValue
      dGetNewValue = ColdFormat.TopOffset
      MsgBox "Old Value = " + Str(dGetCurValue) + Chr(13) + "New Value = " + Str(dGetNewValue), vbInformation, "Changed"
   Else
      MsgBox "Value was not changed.", vbInformation, "No Change"
   End If
   
   ' De-initialize the object variable
   Set ColdFormat = Nothing
End Sub

Public Sub IINIColdFormatting_Units()
   'RobY Dec2/98
   Dim ColdFormat As IINIColdFormatting
   Dim iGetCurValue As UNIT_TYPE
   Dim iNewValue As UNIT_TYPE
   Dim iResponse As Integer
   Dim iGetNewValue As UNIT_TYPE
   
   ' Set object variable for IINIColdFormatting interface to Configuration ctrl object
   Set ColdFormat = MainMDIForm.ActiveForm.SpicerConfiguration1.object
   
   iGetCurValue = ColdFormat.units
   iResponse = MsgBox("Currently set to " + Str(iGetCurValue) + "." + Chr(13) + "Do you want to change it?", vbYesNo + vbInformation, "Return Value")
   If iResponse = vbYes Then
      iNewValue = InputBox("Please enter new value.", "New Value")
      ColdFormat.units = iNewValue
      iGetNewValue = ColdFormat.units
      MsgBox "Old Value = " + Str(iGetCurValue) + Chr(13) + "New Value = " + Str(iGetNewValue), vbInformation, "Changed"
   Else
      MsgBox "Value was not changed.", vbInformation, "No Change"
   End If
   
   ' De-initialize the object variable
   Set ColdFormat = Nothing
End Sub

