Attribute VB_Name = "mod_IPrintView"
' File:      IPrintView.bas
' Created:   1998Apr28 by Rob Wood, Mark Simpson, and Rob Young
' Purpose:   To test the Spicer View Control's IPrintView interface.
' Revisions: RobY Jan4/99  - Added functionality to open more than 1 file. ChildForm1 is no longer used access objects because it has been set to ChildForm() which is an array of ChildForm1's forms.
'                          - Anywhere ChildForm1 was used has been changed to ChildForm(iIndex) or MainMDIForm.ActiveForm.
'   $log$
'   $nokeywords$
'
' Copyright (C) 1998 SPICER Corporation, All rights reserved.

Option Explicit
Dim MenuName As Menu

Public Sub IPrintView_FaxDialog()
   ' RobY May13/98
   ' Define the instance for the interface
   Dim PrintView As IPrintView
   Set PrintView = MainMDIForm.ActiveForm.SpicerView1.object
   
   ' Display the fax dialog box
   PrintView.FaxDialog
   
   ' De-initialize the object var
   Set PrintView = Nothing

End Sub

Public Sub IPrintView_FaxDialogAvailability()
   'RobY July1/98
   ' Modified by RobY Dec11/98 - changed global function to change menu status of ImageX according to what is returned from availablity
   Dim PrintView As IPrintView
   Dim iAvailable As COMMAND_AVAILABILITY
   
   ' Set object variable for IPrintView interface to doc ctrl object
   Set PrintView = MainMDIForm.ActiveForm.SpicerView1.object
   ' Set the menu name to change
   Set MenuName = MainMDIForm.ActiveForm.mnu_ViewCtrl_IPrintView_Array(0)
   
   ' Get the value of availability
   iAvailable = PrintView.FaxDialogAvailability
   ' Change the menu status through function
   Availability iAvailable, MenuName
   
   ' De-initialize object var
   Set PrintView = Nothing
End Sub

Public Sub IPrintView_PrintDocument()
   ' RobY May12/98
   ' Modified by RobY June26/98
   ' Modified by RobY Oct27/98 - Removed PrintType parameter from method and code. Not used.
   Dim PrintView As IPrintView
   Dim iPageRange As Integer
   Dim lFirstPage As Long
   Dim lLastPage As Long
   Dim lNumCopies As Long
   Dim iPrintMode As Integer
   Dim iTile As Integer
   Dim bTile As Boolean
   Dim iZoomMode As Integer
   Dim iPrintOrient As Integer
   Dim iBanner As Integer
   Dim bBanner As Boolean
   Dim iStamp As Integer
   Dim bStamp As Boolean
   Dim strReturn As String
   
   ' Set default values
   iPageRange = 0
   lFirstPage = 0
   lLastPage = 0
   lNumCopies = 0
   iPrintMode = 0
   bTile = False
   iZoomMode = 0
   iPrintOrient = 0
   bBanner = False
   bStamp = False
   strReturn = Chr(13) + Chr(10)
   ' Define the instance for the interface
   Set PrintView = MainMDIForm.ActiveForm.SpicerView1.object
   
   iPageRange = InputBox("Enter page range." + strReturn + "Current Page = 1" + _
                        strReturn + "Page Range = 2" + strReturn + _
                        "All Pages = 3", "Page Range")
   If iPageRange = 2 Then
      lFirstPage = InputBox("What page do you want to start printing from?", "First Page")
      lLastPage = InputBox("What page do you want to finish at?", "Last Page")
   End If
   lNumCopies = InputBox("How many copies would you like?", "Number of Copies")
   iPrintMode = InputBox("Please enter a print mode." + strReturn + "Document = 1," + strReturn + _
                     "Layers = 2," + strReturn + "As Displayed = 4" + strReturn + "Active Raster = 8" + strReturn + _
                     "Active Edit = 16" + strReturn + "Rasters Displayed = 32" + strReturn + "Edits Displayed = 64", _
                     "Print Mode")
   iTile = MsgBox("Do you want to print tiled?", vbQuestion + vbYesNo, "Tile")
   If iTile = vbYes Then
      bTile = True
   End If
   iZoomMode = InputBox("Please enter zoom mode from list." + strReturn + "SCALE TO FIT = 3" + _
                     strReturn + "ACTUAL SIZE = 6" + strReturn + "ZOOM HALF PAGE = 11" + _
                     strReturn + "NO SCALE = 12" + strReturn + "ACTUAL SIZE OR FIT = 13", "Zoom Mode")
   iPrintOrient = InputBox("Please enter print orientation from list." + strReturn + _
                        "BEST FIT = 1" + strReturn + "PORTRAIT = 2" + strReturn + _
                        "LANDSCAPE = 3" + strReturn + "MIN LENGTH = 4", "Print Orientation")
   iBanner = MsgBox("Do you want to print a banner?", vbQuestion + vbYesNo, "Banner")
   If iBanner = vbYes Then
      bBanner = True
   End If
   iStamp = MsgBox("Do you want to print a stamp?", vbQuestion + vbYesNo, "Stamp")
   If iStamp = vbYes Then
      bStamp = True
   End If
   PrintView.PrintDocument iPageRange, lFirstPage, lLastPage, lNumCopies, iPrintMode, bTile, iZoomMode, iPrintOrient, bBanner, bStamp

   ' De-initialize the object var
   Set PrintView = Nothing
End Sub

Public Sub IPrintView_PrintArea()
   ' RobY May13/98
   ' Modified by RobY June26/98
   Dim PrintView As IPrintView
   Dim x1 As Long
   Dim x2 As Long
   Dim y1 As Long
   Dim y2 As Long
   Dim iBanner As Integer
   Dim iStamp As Integer
   
   ' Define the instance for the interface
   Set PrintView = MainMDIForm.ActiveForm.SpicerView1.object
   
   ' Get x and y values for area to print
   x1 = InputBox("Enter the value for X1.", "X1 Value")
   y1 = InputBox("Enter the value for Y1.", "Y1 Value")
   x2 = InputBox("Enter the value for X2.", "X2 Value")
   y2 = InputBox("Enter the value for Y2.", "Y2 Value")
   ' Prompt to display banner
   iBanner = MsgBox("Do you want to display a banner?", vbQuestion + vbYesNo, "Banner")
   ' Prompt to display stamp
   iStamp = MsgBox("Do you want to display a stamp?", vbQuestion + vbYesNo, "Stamp")
   ' Print the specified area
   If iBanner = vbYes And iStamp = vbYes Then
      PrintView.PrintArea IN_UNITS_PROPORTIONAL, x1, y1, x2, y2, True, True
   ElseIf iBanner = vbNo And iStamp = vbNo Then
      PrintView.PrintArea IN_UNITS_PROPORTIONAL, x1, y1, x2, y2, False, False
   ElseIf iBanner = vbYes And iStamp = vbNo Then
      PrintView.PrintArea IN_UNITS_PROPORTIONAL, x1, y1, x2, y2, True, False
   ElseIf iBanner = vbNo And iStamp = vbYes Then
      PrintView.PrintArea IN_UNITS_PROPORTIONAL, x1, y1, x2, y2, False, True
   End If
   
   ' De-initialize the object var
   Set PrintView = Nothing
End Sub

Public Sub IPrintView_PrintDialog()
   ' RobY May13/98
   ' Define the instance for the interface
   Dim PrintView As IPrintView
   Set PrintView = MainMDIForm.ActiveForm.SpicerView1.object
   
   ' Display the print dialog box
   PrintView.PrintDialog
   
   ' De-initialize the object var
   Set PrintView = Nothing
End Sub

Public Sub IPrintView_PrintDialogAvailability()
   'RobY June26/98
   ' Modified by RobY Dec11/98 - changed global function to change menu status of ImageX according to what is returned from availablity
   Dim PrintView As IPrintView
   Dim iAvailable As COMMAND_AVAILABILITY
   
   ' Set object variable for IprintView interface to doc ctrl object
   Set PrintView = MainMDIForm.ActiveForm.SpicerView1.object
   ' Set the menu name to change
   Set MenuName = MainMDIForm.ActiveForm.mnu_ViewCtrl_IPrintView_Array(2)
   
   ' Get the value of availability
   iAvailable = PrintView.PrintDialogAvailability
   ' Change the menu status through function
   Availability iAvailable, MenuName
   
   ' De-initialize object var
   Set PrintView = Nothing
End Sub

' <EOF IPrintView.bas>

