Attribute VB_Name = "mod_ISpicerConfiguration"
' File:      ISpicerConfiguation.bas
' Created:   1998July30 by Rob Young
' Purpose:   To test the Spicer Cofiguration Control's ISpicerConfiguration interface.
' Revisions: RobY Jan4/99  - Added functionality to open more than 1 file. ChildForm1 is no longer used access objects because it has been set to ChildForm() which is an array of ChildForm1's forms.
'                          - Anywhere ChildForm1 was used has been changed to ChildForm(iIndex) or MainMDIForm.ActiveForm.
'   $log$
'   $nokeywords$
'
' Copyright (C) 1998 SPICER Corporation, All rights reserved.
Option Explicit

Public Sub ISpicerConfiguration_AboutBox()
   'RobY July30/98
   'Display the About Box for the Spicer Configuration Control
   MainMDIForm.ActiveForm.SpicerConfiguration1.AboutBox
End Sub

Public Sub ISpicerConfiguration_BatchMessageMode()
   'RobY Jan15/99
   Dim bGetBatchMessageMode As Boolean
   Dim bCheckSet As Boolean
   
   ' Get value of BatchMessageMode
   bGetBatchMessageMode = MainMDIForm.ActiveForm.SpicerConfiguration1.BatchMessageMode
   If bGetBatchMessageMode = False Then
      MainMDIForm.ActiveForm.SpicerConfiguration1.BatchMessageMode = True
   Else
      MainMDIForm.ActiveForm.SpicerConfiguration1.BatchMessageMode = False
   End If
   ' Get value of BatchMessageMode after setting it
   bCheckSet = MainMDIForm.ActiveForm.SpicerConfiguration1.BatchMessageMode
   If bCheckSet = bGetBatchMessageMode Then
      MsgBox "Failed to set BatchMessageMode to " + Str(Not bGetBatchMessageMode), vbCritical, "Failure"
   Else
      MsgBox "BatchMessage has been set to " + Str(bCheckSet), vbInformation, "BatchMessageMode"
   End If
End Sub

Public Sub ISpicerConfiguration_GetPrinterName()
   'RobY Jan7/99
   Dim strPrinterName As String
   
   ' Get the printer name
   MainMDIForm.ActiveForm.SpicerConfiguration1.GetPrinterName strPrinterName
   MsgBox "The current printer is " + strPrinterName, vbInformation, "Print Name"
End Sub

Public Sub ISpicerConfiguration_GetPrinterNames()
   'RobY Jan7/99
   Dim strPrinterNames As String
   
   ' Get the printer names that are available
   MainMDIForm.ActiveForm.SpicerConfiguration1.GetPrinterNames strPrinterNames
   MsgBox "The printers that are available are " + strPrinterNames, vbInformation, "Print Name"
End Sub

Public Sub ISpicerConfiguration_SetPrinterName()
   'RobY Jan7/98
   Dim strPrinterName As String
   Dim strDriverName As String
   Dim strDeviceName As String
   
   ' Get the printer name to set
   strPrinterName = InputBox("Please enter the name of the printer to set to.", "Printer Name")
   ' Get the driver name to set
   strDriverName = InputBox("Please enter the name of the driver to set to.", "Driver Name")
    ' Get the device name to set
   strDeviceName = InputBox("Please enter the name of the device to set to.", "Device Name")
   ' Set the printer to use
   MainMDIForm.ActiveForm.SpicerConfiguration1.SetPrinterName strPrinterName, strDriverName, strDeviceName
   MsgBox "The printer has been set to " + strPrinterName + ", " + strDriverName + ", " + strDeviceName + "." + Chr(13) + _
            "Use GetPrinterName to verify that set worked.", vbInformation, "Print Name"
End Sub

