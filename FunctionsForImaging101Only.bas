Attribute VB_Name = "FunctionsForImaging101Only"


Function funcDisableFormsWhileLoadingImages()

'    If funcIsFormLoaded2("frmMainMenu") Then
'        frmMainMenu.Enabled = False
'    End If
'
'    If funcIsFormLoaded2("frmImaging101Search") Then
'        frmImaging101Search.Enabled = False
'    End If
'
'    If funcIsFormLoaded2("frmImaging101Retrieve") Then
'        frmImaging101Retrieve.Enabled = False
'    End If
'
'    If funcIsFormLoaded2("MainMDIForm") Then
'        MainMDIForm.PictureButtonBar.Enabled = False
'    End If
'
'    If funcIsFormLoaded2("frmImaging101BatchList") Then
'        frmImaging101BatchList.Enabled = False
'    End If
'
'    If funcIsFormLoaded2("frmIndex") Then
'        frmIndex.Enabled = False
'    End If
'
'    If funcIsFormLoaded2("frmDocTypeList") Then
'        frmDocTypeList.Enabled = False
'    End If
'
'    If funcIsFormLoaded2("frmLookupList") Then
'        frmLookupList.Enabled = False
'    End If

End Function

Function funcEnableFormsAfterLoadingImages()

'    If funcIsFormLoaded2("frmMainMenu") Then
'        frmMainMenu.Enabled = True
'    End If
'
'    If funcIsFormLoaded2("frmImaging101Search") Then
'        frmImaging101Search.Enabled = True
'    End If
'
'    If funcIsFormLoaded2("frmImaging101Retrieve") Then
'        frmImaging101Retrieve.Enabled = True
'    End If
'
'    If funcIsFormLoaded2("MainMDIForm") Then
'        MainMDIForm.PictureButtonBar.Enabled = True
'    End If
'
'    If funcIsFormLoaded2("frmImaging101BatchList") Then
'        frmImaging101BatchList.Enabled = True
'    End If
'
'    If funcIsFormLoaded2("frmIndex") Then
'        frmIndex.Enabled = True
'    End If
'
'    If funcIsFormLoaded2("frmDocTypeList") Then
'        frmDocTypeList.Enabled = True
'    End If
'
'    If funcIsFormLoaded2("frmLookupList") Then
'        frmLookupList.Enabled = True
'    End If

End Function

Function funcValidateBarCodeLicense() As Boolean
    
'    Dim cCRC32 As New cCRC32
'    Dim strfullstring As String
'    Dim txtLicenseKey As String
'    Dim strBarcodeLicenseKey As String
'
'    strBarcodeLicenseKey = VBGetPrivateProfileString(RegAppname, "frmConfig.txtBarcodeLicenseKey", RegFileName)
'
'    strfullstring = gsecSiteInformationLicenseCode & gProcessorID
'
'    txtLicenseKey = cCRC32.GetStringCRC32(strfullstring)
'
'    If txtLicenseKey = strBarcodeLicenseKey Then
'        funcValidateBarCodeLicense = True
'    Else
'        funcValidateBarCodeLicense = False
'    End If
    
End Function


       Public Function funcAskUserFor_Imaging101_RemoteHost() As String
            strDomain = InputBox("Please enter the Imaging101 Server name.", "Imaging101 Server Name")
            result = WritePrivateProfileString(RegAppname, "Imaging101_RemoteHost", strDomain, RegFileName)
            result = WritePrivateProfileString(RegAppname, "Imaging101_RemoteHost", strDomain, "C:\WINDOWS\Imaging101Client.INI")
            'RETURN SERVER NAME
            funcAskUserFor_Imaging101_RemoteHost = strDomain
       End Function

