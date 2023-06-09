Attribute VB_Name = "mod_IINIVectorPens"
' File:      IINIVectorPens.bas
' Created:   1998Nov30 by Rob Young
' Purpose:   To test the Spicer Configuration Control's IINIVectorPens interface.
' Revisions: RobY Jan4/99  - Added functionality to open more than 1 file. ChildForm1 is no longer used access objects because it has been set to ChildForm() which is an array of ChildForm1's forms.
'                          - Anywhere ChildForm1 was used has been changed to ChildForm(iIndex) or MainMDIForm.ActiveForm.
'   $log$
'   $nokeywords$
'
' Copyright (C) 1998 SPICER Corporation, All rights reserved.
Option Explicit

Public Sub IINIVectorPens_GetVectorPen()
   'RobY Nov30/98
   Dim INIVectorPens As IINIVectorPens
   Dim iPenNum As Integer
   Dim lColor As Long
   Dim dWidth As Double
   Dim iUnitType As UNIT_TYPE
   Dim iDashCount As Integer
   Dim dDash1 As Double
   Dim dDash2 As Double
   Dim dDash3 As Double
   Dim dDash4 As Double
   
   ' Set object variable for IINIVectorPens interface to Configuration ctrl object
   Set INIVectorPens = MainMDIForm.ActiveForm.SpicerConfiguration1.object
      
   iPenNum = InputBox("Please enter a pen number to find the information for it.", "Pen Number")
   INIVectorPens.GetVectorPen iPenNum, lColor, dWidth, iUnitType, iDashCount, dDash1, dDash2, dDash3, dDash4
   ' Display settings
   MsgBox "Pen Number: " + Str(iPenNum) + Chr(13) + "Color: " + Str(lColor) + Chr(13) + "Width: " + Str(dWidth) + Chr(13) + "Unit Type: " + Str(iUnitType) + Chr(13) + _
            "Dash Count: " + Str(iDashCount) + Chr(13) + "Dash 1: " + Str(dDash1) + Chr(13) + "Dash 2: " + _
            Str(dDash2) + Chr(13) + "Dash 3: " + Str(dDash3) + Chr(13) + "Dash 4: " + Str(dDash4), vbInformation, "Get Vector Pen"
            
   ' De-initialize the object variable
   Set INIVectorPens = Nothing
End Sub

Public Sub IINIVectorPens_LoadPenSettings()
   'RobY Nov30/98
   Dim INIVectorPens As IINIVectorPens
   Dim strFilename As String
   
   ' Set object variable for IINIVectorPens interface to Configuration ctrl object
   Set INIVectorPens = MainMDIForm.ActiveForm.SpicerConfiguration1.object
   
   ' Prompt for filename
   strFilename = InputBox("Please enter the filename of the file that has the pen setting to load.", "Load Pen Settings")
   ' Load pen settings
   INIVectorPens.LoadPenSettings strFilename
   
   ' De-initialize the object variable
   Set INIVectorPens = Nothing
End Sub

Public Sub IINIVectorPens_SetVectorPen()
   'RobY Nov30/98
   Dim INIVectorPens As IINIVectorPens
   Dim iPenNum As Integer
   Dim lColor As Long
   Dim dWidth As Double
   Dim iUnitType As UNIT_TYPE
   Dim iDashCount As Integer
   Dim dDash1 As Double
   Dim dDash2 As Double
   Dim dDash3 As Double
   Dim dDash4 As Double

   ' Set object variable for IINIVectorPens interface to Configuration ctrl object
   Set INIVectorPens = MainMDIForm.ActiveForm.SpicerConfiguration1.object
   
   ' Prompt user for values
   iPenNum = InputBox("Please enter the pen number to set the width for.", "Pen Number")
   dWidth = InputBox("Please enter the width of the pen.", "Width")
   iUnitType = InputBox("Please enter the unit type to use as a number.", "Unit Type")
   lColor = InputBox("Please enter the color of the pen as a 24-bit color.", "Pen Color")
   iDashCount = InputBox("Please enter the number of dashes.", "Dash Count")
   dDash1 = InputBox("Please enter value for dash 1.", "Dash 1")
   dDash2 = InputBox("Please enter value for dash 2.", "Dash 2")
   dDash3 = InputBox("Please enter value for dash 3.", "Dash 3")
   dDash4 = InputBox("Please enter value for dash 4.", "Dash 4")
   ' Execute command
   INIVectorPens.SetVectorPen iPenNum, lColor, dWidth, iUnitType, iDashCount, dDash1, dDash2, dDash3, dDash4
   
   ' De-initialize the object variable
   Set INIVectorPens = Nothing
End Sub

Public Sub IINIVectorPens_SetVectorPenPattern()
   'RobY Nov30/98
   Dim INIVectorPens As IINIVectorPens
   Dim iPenNum As Integer
   Dim iDashCount As Integer
   Dim dDash1 As Double
   Dim dDash2 As Double
   Dim dDash3 As Double
   Dim dDash4 As Double

   ' Set object variable for IINIVectorPens interface to Configuration ctrl object
   Set INIVectorPens = MainMDIForm.ActiveForm.SpicerConfiguration1.object
   
   ' Prompt for the pen values
   iPenNum = InputBox("Please enter the pen number to set the width for.", "Pen Number")
   iDashCount = InputBox("Please enter the number of dashes.", "Dash Count")
   dDash1 = InputBox("Please enter value for dash 1.", "Dash 1")
   dDash2 = InputBox("Please enter value for dash 2.", "Dash 2")
   dDash3 = InputBox("Please enter value for dash 3.", "Dash 3")
   dDash4 = InputBox("Please enter value for dash 4.", "Dash 4")
   ' Execute command
   INIVectorPens.SetVectorPenPattern iPenNum, iDashCount, dDash1, dDash2, dDash3, dDash4
   
   ' De-initialize the object variable
   Set INIVectorPens = Nothing
End Sub

Public Sub IINIVectorPens_SetVectorPenWidth()
   'RobY Nov30/98
   Dim INIVectorPens As IINIVectorPens
   Dim iPenNum As Integer
   Dim dWidth As Double
   Dim iUnitType As UNIT_TYPE
   
   ' Set object variable for IINIVectorPens interface to Configuration ctrl object
   Set INIVectorPens = MainMDIForm.ActiveForm.SpicerConfiguration1.object
   
   ' Prompt for the pen number
   iPenNum = InputBox("Please enter the pen number to set the width for.", "Pen Number")
   ' Prompt for pen width
   dWidth = InputBox("Please enter the width of the pen.", "Width")
   ' Prompt for unit type
   iUnitType = InputBox("Please enter the unit type to use as a number.", "Unit Type")
   ' Execute command
   INIVectorPens.SetVectorPenWidth iPenNum, dWidth, iUnitType
   
   ' De-initialize the object variable
   Set INIVectorPens = Nothing
End Sub

Public Sub IINIVectorPens_VectorPenColor()
   'RobY Nov30/98
   Dim INIVectorPens As IINIVectorPens
   Dim lPenColor As Long
   
   ' Set object variable for IINIVectorPens interface to Configuration ctrl object
   Set INIVectorPens = MainMDIForm.ActiveForm.SpicerConfiguration1.object
   
   lPenColor = INIVectorPens.VectorPenColor(1)
   ' Display the pen color for pen 1
   MsgBox "Pen color for Pen 1 is " + Str(lPenColor), vbInformation, "Pen Color"
   
   ' De-initialize the object variable
   Set INIVectorPens = Nothing
End Sub

