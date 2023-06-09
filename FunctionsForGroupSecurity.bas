Attribute VB_Name = "FunctionsForGroupSecurity"
Public Function funcGetDomainGroupsAndUsers()


    ' *****************************
    ' * List All Groups in the Domain and
    ' * List All Members of each Group
    ' *
    ' * Output to a text file on the user's desktop in the format:
    ' * group name <tab> type <tab> member name <tab> type
    ' * Prompt for text file name.
    ' * Written by James Anderson, July 2009
    ' *****************************
    ' Variables
    Const MY_DOMAIN = "dc=fabricam,dc=com"
    ' *****************************
    ' Start Main
    On Error Resume Next
    Const ADS_SCOPE_SUBTREE = 2
    Const ADS_GROUP_TYPE_GLOBAL_GROUP = &H2
    Const ADS_GROUP_TYPE_LOCAL_GROUP = &H4
    Const ADS_GROUP_TYPE_UNIVERSAL_GROUP = &H8
    Const ADS_GROUP_TYPE_SECURITY_ENABLED = &H80000000
    Const E_ADS_PROPERTY_NOT_FOUND = &H8000500D
    Const MYPROMPT = "Enter the Output filename (i.e. Groups.txt) that will be saved on your desktop:"
    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    Set objFSO = CreateObject("Scripting.FileSystemObject")
     
    ' Setup the output file
    If UCase(Right(Wscript.fullname, 12)) = "\CSCRIPT.EXE" Then
      Wscript.StdOut.Write MYPROMPT & " "
      strMyFileName = Wscript.StdIn.ReadLine
    Else
      strMyFileName = InputBox(MYPROMPT)
    End If
    If strMyFileName = "" Then
      Wscript.Quit
    End If
    Set WshShell = CreateObject("WScript.Shell")
    Set WshSysEnv = WshShell.Environment("PROCESS")
    strMyFileName = WshSysEnv("USERPROFILE") & "\Desktop\" & strMyFileName
    Set WshSysEnv = Nothing
    Set WshShell = Nothing
    If objFSO.FileExists(strMyFileName) Then
      'objFSO.DeleteFile(strMyFileName)
      Wscript.Echo "That filename already exists"
      Wscript.Quit
    End If
     
    ' Get a recordset of groups in AD
    Set objMyOutput = objFSO.OpenTextFile(strMyFileName, ForWriting, True)
    Set objConnection = CreateObject("ADODB.Connection")
    Set objCommand = CreateObject("ADODB.Command")
    objConnection.Provider = "ADsDSOObject"
    objConnection.Open "Active Directory Provider"
    Set objCommand.ActiveConnection = objConnection
    objCommand.Properties("Page Size") = 1000
    objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE
    objCommand.CommandText = _
        "SELECT ADsPath, Name FROM 'LDAP://" & MY_DOMAIN & "' WHERE objectCategory='group'"
    Set objRecordSet = objCommand.Execute
    objRecordSet.MoveFirst
     
    ' For each Group, Get group properties
    Do Until objRecordSet.EOF
      Set objGroup = GetObject(objRecordSet.Fields("ADsPath").Value)
      strGroupName = objRecordSet.Fields("Name").Value
      If objGroup.GroupType And ADS_GROUP_TYPE_LOCAL_GROUP Then
        strGroupDesc = "Domain local "
      ElseIf objGroup.GroupType And ADS_GROUP_TYPE_GLOBAL_GROUP Then
        strGroupDesc = "Global "
      ElseIf objGroup.GroupType And ADS_GROUP_TYPE_UNIVERSAL_GROUP Then
        strGroupDesc = "Universal "
      Else
        strGroupDesc = "Unknown "
      End If
      If objGroup.GroupType And ADS_GROUP_TYPE_SECURITY_ENABLED Then
        strGroupDesc = strGroupDesc & "Security group"
      Else
        strGroupDesc = strGroupDesc & "Distribution group"
      End If
     
      ' Check if there are members
      Err.Clear
      arrMemberOf = objGroup.GetEx("Member")
      If Err.Number = E_ADS_PROPERTY_NOT_FOUND Then
        ' Write a line to the outputfile with group properties and no members
        objMyOutput.WriteLine (strGroupName & vbTab & strGroupDesc & vbTab & "<null>" & vbTab & "<null>")
      Else
        ' For each group member, get member properties
        For Each strMemberOf In arrMemberOf
          Set objMember = GetObject("LDAP://" & strMemberOf)
          strMemberName = Right(objMember.name, Len(objMember.name) - 3)
          ' Write a line to the outputfile with group and member properties
          objMyOutput.WriteLine (strGroupName & vbTab & strGroupDesc & vbTab & strMemberName & vbTab & objMember.Class)
          Set objMember = Nothing
        Next
      End If
      objRecordSet.MoveNext
      Set objGroup = Nothing
    Loop
    objMyOutput.Close
    Wscript.Echo "Done!"

End Function

Public Function funcGetDomainGroupsAndUsers2()

'Option Explicit
Dim adoCommand, adoConnection, strBase, strFilter, strAttributes
Dim objRootDSE, strDNSDomain, strQuery, adoRecordset, strName, strCN

' Setup ADO objects.
Set adoCommand = CreateObject("ADODB.Command")
Set adoConnection = CreateObject("ADODB.Connection")
adoConnection.Provider = "ADsDSOObject"
adoConnection.Open "Active Directory Provider"
Set adoCommand.ActiveConnection = adoConnection

' Search entire Active Directory domain.
Set objRootDSE = GetObject("LDAP://RootDSE")

strDNSDomain = objRootDSE.Get("defaultNamingContext")
strBase = "<LDAP://" & strDNSDomain & ">"

' Filter on user objects.
'strFilter = "(&(objectCategory=person)(objectClass=user))"
strFilter = "(&(objectCategory=group))"

' Comma delimited list of attribute values to retrieve.
strAttributes = "sAMAccountName,cn"

' Construct the LDAP syntax query.
strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
adoCommand.CommandText = strQuery
adoCommand.Properties("Page Size") = 100
adoCommand.Properties("Timeout") = 30
adoCommand.Properties("Cache Results") = False

' Run the query.
Set adoRecordset = adoCommand.Execute

' Enumerate the resulting recordset.
Do Until adoRecordset.EOF
    ' Retrieve values and display.
    strName = adoRecordset.Fields("sAMAccountName").Value
    strCN = adoRecordset.Fields("cn").Value
    Debug.Print "NT Name: " & strName & ", Common Name: " & strCN
    ' Move to the next record in the recordset.
    adoRecordset.MoveNext
Loop

' Clean up.
adoRecordset.Close
adoConnection.Close
End Function

Public Function AllGroups() As String()

        'PURPOSE:  Gets all groups for the current domain
        'and returns them in a string array, using LDAP
        
        'Requires: ADSI, LDAP provider
        'This function tested on Windows 2000 RC2
        
        'RETURNS: String array containing all
        'Groups for the current domain
        
        'Requires VB6 because in lower versions
        'array cannot be return type for a
        'function
        
        'EXAMPLE
        'Dim sArray() As String
        'Dim iCtr As Integer
        
        'sArray = AllGroups
        'For iCtr = 0 To UBound(sArray)
        '    Debug.Print sArray(iCtr)
        'Next
        
        Dim conn As New ADODB.Connection
        Dim rs As ADODB.Recordset
        Dim oRoot As IADs
        Dim oDomain As IADs
        Dim sBase As String
        Dim sFilter As String
        Dim sDomain As String
        
        Dim sAttribs As String
        Dim sDepth As String
        Dim sQuery As String
        Dim sAns() As String
        Dim iElement As Integer
        
        On Error GoTo errhandler:
        
        Set oRoot = GetObject("LDAP://rootDSE")
        sDomain = oRoot.Get("defaultNamingContext")
        Set oDomain = GetObject("LDAP://" & sDomain)
        sBase = "<" & oDomain.ADsPath & ">"
        sFilter = "(&(objectCategory=group))"
        sAttribs = "name"
        sDepth = "subTree"
        
        sQuery = sBase & ";" & sFilter & ";" & sAttribs & ";" & sDepth
                           
        conn.Open _
          "Data Source=Active Directory Provider;Provider=ADsDSOObject"
          
        Set rs = conn.Execute(sQuery)
        ReDim sAns(0) As String
        
        With rs
            Do While Not .EOF
                iElement = IIf(sAns(0) = "", 0, iElement + 1)
                ReDim Preserve sAns(iElement) As String
                sAns(iElement) = rs("name")
               .MoveNext
            Loop
        End With
        AllGroups = sAns
        
errhandler:
        
        On Error Resume Next
        If rs.State <> 0 Then rs.Close
        If conn.State <> 0 Then conn.Close
        Set rs = Nothing
        Set conn = Nothing
        Set oRoot = Nothing
        Set oDomain = Nothing

End Function
