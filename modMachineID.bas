Attribute VB_Name = "modMachineID"
'******************************************************************************
'*** Get the CPU and BIOS unique ID
'***  To get them use the following code:
'
'  Dim CPU As String
'  CPU = GetWmiDeviceSingleValue("Win32_Processor", "ProcessorID")
'
'  Dim BIOS As String
'  BIOS = GetWmiDeviceSingleValue("Win32_BIOS", "SerialNumber")
'
Option Explicit

Private m_mainWmi As Object
Private m_deviceLists As Collection

Private Function GetMainWMIObject() As Object
  On Error GoTo eh
  If m_mainWmi Is Nothing Then
    Set m_mainWmi = GetObject("WinMgmts:")
  End If
  Set GetMainWMIObject = m_mainWmi
  Exit Function
eh:
  Set GetMainWMIObject = Nothing
End Function

Public Function WmiIsAvailable() As Boolean
  WmiIsAvailable = CBool(Not GetMainWMIObject Is Nothing)
End Function

Public Function GetWmiDeviceSingleValue(ByVal WmiClass As String, ByVal WmiProperty As String) As String
  On Error GoTo done
  Dim result As String
  
  Dim wmiclassObjList As Object
  Set wmiclassObjList = GetWmiDeviceList(WmiClass)
  Dim wmiclassObj As Object
  For Each wmiclassObj In wmiclassObjList
    result = CallByName(wmiclassObj, WmiProperty, VbGet)
    Exit For
  Next

done:
  GetWmiDeviceSingleValue = Trim(result)
End Function

Public Function GetWmiDeviceList(ByVal WmiClass As String) As Object
  If m_deviceLists Is Nothing Then
    Set m_deviceLists = New Collection
  End If
  
  On Error GoTo fetchNew
  
  Set GetWmiDeviceList = m_deviceLists.item(WmiClass)
  Exit Function
  
fetchNew:
  Dim devList As Object
  Set devList = GetWmiDeviceListInternal(WmiClass)
  If Not devList Is Nothing Then
    Call m_deviceLists.Add(devList, WmiClass)
  End If
  Set GetWmiDeviceList = devList
End Function

Private Function GetWmiDeviceListInternal(ByVal WmiClass As String) As Object
  On Error GoTo eh
  Set GetWmiDeviceListInternal = GetMainWMIObject.Instancesof(WmiClass)
  Exit Function
eh:
  Set GetWmiDeviceListInternal = Nothing
End Function



