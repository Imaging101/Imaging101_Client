Attribute VB_Name = "mod_IINIMaskTable"
' File:      IINIMaskTable.bas
' Created:   1999Jan18 by Rob Young
' Purpose:   To test the Spicer Configuration Control's IINIMaskTable interface.
' Revisions:
'   $log$
'   $nokeywords$
'
' Copyright (C) 1998 SPICER Corporation, All rights reserved.
Option Explicit

Public Sub IINIMaskTable_DefaultMaskTable()
   ' RobY Jan18/99
   Dim INIMaskTable As IINIMaskTable
   Dim lMaskTableID As Long
   
   ' Set object variable for INIMaskTable interface to Configuration ctrl object
   Set INIMaskTable = MainMDIForm.ActiveForm.SpicerConfiguration1.object
   
   lMaskTableID = InputBox("Please enter the mask table id to set as the default.", "Mask Table ID")
   INIMaskTable.DefaultMaskTable(IN_MASKTABLE_TYPE_RASTERIZE) = lMaskTableID
   
   ' De-initialize the object variable
   Set INIMaskTable = Nothing
End Sub

Public Sub IINIMaskTable_LoadMaskTableFile()
   ' RobY Jan20/99
   Dim INIMaskTable As IINIMaskTable
   
   ' Set object variable for INIMaskTable interface to Configuration ctrl object
   Set INIMaskTable = MainMDIForm.ActiveForm.SpicerConfiguration1.object
   
   MainMDIForm.ActiveForm.CommonDialog1.CancelError = True
   MainMDIForm.ActiveForm.CommonDialog1.Flags = cdlOFNHideReadOnly
   MainMDIForm.ActiveForm.CommonDialog1.DialogTitle = "Select Mask File"
   MainMDIForm.ActiveForm.CommonDialog1.Filter = "Spicer Mask (*.spm)|*.spm|All Files (*.*)|*.*"
   MainMDIForm.ActiveForm.CommonDialog1.ShowOpen
   INIMaskTable.LoadMaskTableFile MainMDIForm.ActiveForm.CommonDialog1.filename
   ' Initialize filename to null
   MainMDIForm.ActiveForm.CommonDialog1.filename = ""
   MainMDIForm.ActiveForm.CommonDialog1.Filter = "All Files (*.*)|*.*"
   
   ' De-initialize the object variable
   Set INIMaskTable = Nothing
End Sub
