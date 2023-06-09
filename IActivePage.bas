Attribute VB_Name = "mod_IActivePage"
' File:      IActivePage.bas
' Created:   1998Apr28 by Rob Wood, Mark Simpson, and Rob Young
' Purpose:   To test the Spicer View Control's IActivePage interface.
' Revisions: RobY Jan4/99  - Added functionality to open more than 1 file. ChildForm1 is no longer used access objects because it has been set to ChildForm() which is an array of ChildForm1's forms.
'                          - Anywhere ChildForm1 was used has been changed to ChildForm(iIndex) or MainMDIForm.ActiveForm.
'   $log$
'   $nokeywords$
'
' Copyright (C) 1998 SPICER Corporation, All rights reserved.

Option Explicit
Dim MenuName As Menu

Public Sub IActivePage_ActivePageID()
   ' RobY May11/98
   ' Define the instance for the interface
   Dim ActivePage As IActivePage
   Set ActivePage = MainMDIForm.ActiveForm.SpicerView1.object
   
   ' Display the Active Page ID
   MsgBox ("Active Page ID: " + Str(ActivePage.ActivePageId) + " (Should NOT be zero)")
   
   ' De-initialize the object var
   Set ActivePage = Nothing
End Sub

Public Sub IActivePage_BindToDocumentControl()
   ' RobY May11/98
   ' Define the instance for the interface
   Dim ActivePage As IActivePage
   Set ActivePage = MainMDIForm.ActiveForm.SpicerView1.object

   ' Bind doc control to the view control
   ActivePage.BindToDocumentControl (MainMDIForm.ActiveForm.SpicerDoc1.object)
   
   'De-initialize the object var
   Set ActivePage = Nothing

End Sub

Public Sub IActivePage_FindTextMatch()
   'RobY June18/98
   ' Define the instance for the interface
   Dim ActivePage As IActivePage
   Dim strText As String
   Dim iFlag As Integer
   Dim iDirection As Integer
   Set ActivePage = MainMDIForm.ActiveForm.SpicerView1.object
   
   strText = InputBox("Enter text to find.", "FindTextMatch")
   iFlag = InputBox("Case Insensitive = 1" + Chr(10) + Chr(13) + _
      "Find Whole Word = 2" + Chr(10) + Chr(13) + "Match Case = 4" + _
      Chr(13) + Chr(10) + "**NOTE: Can be ORed!**", "FindTextMatch")
   iDirection = InputBox("Please enter one of the following directions:" + Chr(10) + Chr(13) + _
      "First = 0" + Chr(10) + Chr(13) + "Last = 1" + Chr(10) + Chr(13) + _
      "Previous = 2" + Chr(10) + Chr(13) + "Next = 3", "FindTextMatch")
   'Find the text
   ActivePage.FindTextMatch strText, iFlag, iDirection
   
   'De-initialize the object var
   Set ActivePage = Nothing
End Sub

Public Sub IActivePage_GotoFirstPage()
   ' RobY May11/98
   ' Define the instance for the interface
   Dim ActivePage As IActivePage
   Set ActivePage = MainMDIForm.ActiveForm.SpicerView1.object
   
   ' Displays the first page
   ActivePage.GotoFirstPage
   
   'De-initialize the object var
   Set ActivePage = Nothing
End Sub

Public Sub IActivePage_GotoLastPage()
   ' RobY May11/98
   ' Define the instance for the interface
   Dim ActivePage As IActivePage
   Set ActivePage = MainMDIForm.ActiveForm.SpicerView1.object
   
   ' Displays the last page
   ActivePage.GotoLastPage
   
   'De-initialize the object var
   Set ActivePage = Nothing
End Sub

Public Sub IActivePage_GotoPage()
   ' RobY May11/98
   ' Define the instance for the interface
   Dim ActivePage As IActivePage
   Dim iPageNum As Integer
   Dim lPageID As Long
   Set ActivePage = MainMDIForm.ActiveForm.SpicerView1.object
   
   ' Prompt user for page number
   iPageNum = InputBox("Enter a page number to display a specific page of the document.", "Goto Page")
   ' Get the pageID for the page number the user entered
   lPageID = MainMDIForm.ActiveForm.SpicerDoc1.pageID(iPageNum)
   ' Goes to page that the user specified
   ActivePage.GotoPage (lPageID)
   
   'De-initialize the object var
   Set ActivePage = Nothing

End Sub

Public Sub IActivePage_GotoPageRelative()
   ' RobY May11/98
   ' Define the instance for the interface
   Dim ActivePage As IActivePage
   Dim iPageRelative As Integer
   Set ActivePage = MainMDIForm.ActiveForm.SpicerView1.object
   
   ' Prompt user for a number to use for the GotoPageRelative command
   iPageRelative = InputBox("Enter a number to use to make the page with the specified relative position the current page.", "Goto Page")

   ' Displays the page with the specified relative position the current page
   ActivePage.GotoPageRelative (iPageRelative)
     
   'De-initialize the object var
   Set ActivePage = Nothing
End Sub

Public Sub IActivePage_hWnd()
   ' RobY May11/98
   ' Define the instance for the interface
   Dim ActivePage As IActivePage
   Set ActivePage = MainMDIForm.ActiveForm.SpicerView1.object

   ' Displays the window handle(hWnd)
   MsgBox ("Window Handle(hWnd): " + Str(ActivePage.hwnd))
   
   'De-initialize the object var
   Set ActivePage = Nothing
End Sub

Public Sub IActivePage_PageContentsDialog()
   ' RobY May12/98
   ' Define the instance for the interface
   Dim ActivePage As IActivePage
   Set ActivePage = MainMDIForm.ActiveForm.SpicerView1.object
   
   'Display the page contents dialog
   ActivePage.PageContentsDialog
   
   'De-initialize the object var
   Set ActivePage = Nothing
End Sub

Public Sub IActivePage_PageContentsDialogAvailability()
   'RobY July1/98
   ' Modified by RobY Dec11/98 - changed global function to change menu status of ImageX according to what is returned from availablity
   Dim ActivePage As IActivePage
   Dim iAvailable As COMMAND_AVAILABILITY
   
   ' Set object variable for IActivePage interface to doc ctrl object
   Set ActivePage = MainMDIForm.ActiveForm.SpicerView1.object
   ' Set the menu name to change
   Set MenuName = MainMDIForm.ActiveForm.mnu_ViewCtrl_IActivePage_Array(8)
   
   ' Get the value of availability
   iAvailable = ActivePage.PageContentsDialogAvailability
   ' Change the menu status through function
   Availability iAvailable, MenuName
   
   ' De-initialize object var
   Set ActivePage = Nothing
End Sub

Public Sub IActivePage_PageGotoDialog()
   ' RobY May12/98
   ' Define the instance for the interface
   Dim ActivePage As IActivePage
   Set ActivePage = MainMDIForm.ActiveForm.SpicerView1.object

   ' Display the page go to dialog
   ActivePage.PageGotoDialog
   
   'De-initialize the object var
   Set ActivePage = Nothing
End Sub

Public Sub IActivePage_PageGotoDialogAvailability()
   'RobY July1/98
   ' Modified by RobY Dec11/98 - changed global function to change menu status of ImageX according to what is returned from availablity
   Dim ActivePage As IActivePage
   Dim iAvailable As COMMAND_AVAILABILITY
   
   ' Set object variable for IActivePage interface to doc ctrl object
   Set ActivePage = MainMDIForm.ActiveForm.SpicerView1.object
   ' Set the menu name to change
   Set MenuName = MainMDIForm.ActiveForm.mnu_ViewCtrl_IActivePage_Array(9)
   
   ' Get the value of availability
   iAvailable = ActivePage.PageGotoDialogAvailability
   ' Change the menu status through function
   Availability iAvailable, MenuName
   
   ' De-initialize object var
   Set ActivePage = Nothing
End Sub

Public Sub IActivePage_Refresh()
   ' RobY May11/98
   ' Define the instance for the interface
   Dim ActivePage As IActivePage
   Set ActivePage = MainMDIForm.ActiveForm.SpicerView1.object

   ActivePage.Refresh
   'De-initialize the object var
   Set ActivePage = Nothing
End Sub

Public Sub IActivePage_RefreshAvailability()
   'RobY July1/98
   ' Modified by RobY Dec11/98 - changed global function to change menu status of ImageX according to what is returned from availablity
   Dim ActivePage As IActivePage
   Dim iAvailable As COMMAND_AVAILABILITY
   
   ' Set object variable for IDocContents interface to doc ctrl object
   Set ActivePage = MainMDIForm.ActiveForm.SpicerView1.object
   ' Set the menu name to change
   Set MenuName = MainMDIForm.ActiveForm.mnu_ViewCtrl_IActivePage_Array(11)
   
   ' Get the value of availability
   iAvailable = ActivePage.RefreshAvailability
   ' Change the menu status through function
   Availability iAvailable, MenuName
   
   ' De-initialize object var
   Set ActivePage = Nothing
End Sub

Public Sub IActivePage_RefreshMode()
   ' RobY May11/98
   ' Define the instance for the interface
   Dim ActivePage As IActivePage
   Dim bRefreshModeValue As Boolean
   Set ActivePage = MainMDIForm.ActiveForm.SpicerView1.object
   
   If ActivePage.RefreshMode = False Then
      ActivePage.RefreshMode = True
   Else
      ActivePage.RefreshMode = False
   End If
   bRefreshModeValue = ActivePage.RefreshMode
   MsgBox "Refresh Mode = " + Str(bRefreshModeValue), vbExclamation, "Refresh Mode"
   
   'De-initialize the object var
   Set ActivePage = Nothing
End Sub

Public Sub IActivePage_ShowTextMatch()
   ' RobY May12/98
   ' Modified By RobYJune22/98
   Dim ActivePage As IActivePage
   Dim iState As String
   Dim bState As Boolean
   
   ' Set object variable for IActivePage interface to view ctrl object
   Set ActivePage = MainMDIForm.ActiveForm.SpicerView1.object
   
   iState = MsgBox("Do you want to show text match?", vbQuestion + vbYesNo, "Show Text Match")
   If iState = vbYes Then
      bState = True
   Else
      bState = False
   End If
   ActivePage.ShowTextMatch (bState)
   MsgBox ("ShowTextMatch = " + Str(bState))
   
   'De-initialize the object var
   Set ActivePage = Nothing
End Sub

' <EOF IActivePage.bas>


