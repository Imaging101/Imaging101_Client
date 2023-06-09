VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSimpleCDBurner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const CLSID_CDBURN As String = "fbeb8a05-beee-4442-804e-409d6c4515e9"
Private Const IID_CDBURN As String = "3d73a659-e5d0-4d42-afc0-5121ba425c8d"

Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" _
   (ByVal hWndOwner As Long, _
    ByVal nFolder As Long, _
    pidl As Long) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetMalloc Lib "shell32.dll" (ppMalloc As IVBMalloc) As Long
Private Const CSIDL_CDBURN_AREA  As Long = &H3B&        '// USERPROFILE\Local Settings\Application Data\Microsoft\CD Burning
Private Const MAX_PATH As Long = 260&

Private Const FAIL_BIT As Long = &H80000000

Private m_alloc As IVBMalloc
Private m_cdBurn As ICDBurn
Private m_hWndOwner As Long

Private Const ERR_BASE As Long = 47600

Public Property Get BurnStagingAreaFolder() As String
Dim pidl As Long
On Error Resume Next
   ' Get pidl of special folder:
   SHGetSpecialFolderLocation m_hWndOwner, CSIDL_CDBURN_AREA, pidl
   If Err = 0 Then
      ' Convert it to a path:
      BurnStagingAreaFolder = PathFromPidl(pidl)
      ' Free the pidl:
      Allocator.Free ByVal pidl
      pidl = 0
   End If
End Property

Property Get Allocator() As IVBMalloc
    If m_alloc Is Nothing Then SHGetMalloc m_alloc
    Set Allocator = m_alloc
End Property

Private Function PathFromPidl(ByVal pidl As Long) As String
Dim sPath As String
Dim lR As Long
   sPath = String$(MAX_PATH, 0)
   lR = SHGetPathFromIDList(pidl, sPath)
   If lR <> 0 Then
      PathFromPidl = Left$(sPath, lstrlen(sPath))
   End If
End Function

Public Property Get HasRecordableDrive() As Boolean
Dim lHasRecorder As Long
   m_cdBurn.HasRecordableDrive lHasRecorder
   HasRecordableDrive = Not (lHasRecorder = 0)
End Property

Public Property Get RecorderDriveLetter() As String
   Dim sDrive As String
   sDrive = String(10, 0)
   m_cdBurn.GetRecorderDriveLetter sDrive, Len(sDrive)
   Dim iPos As Long
   iPos = InStr(sDrive, vbNullChar)
   If (iPos > 1) Then
      sDrive = Left(sDrive, iPos - 1)
   End If
   RecorderDriveLetter = sDrive
End Property

Public Sub Burn()
   m_cdBurn.Burn m_hWndOwner
End Sub

Public Sub Initialise(ByVal hWndOwner As Long)
    Dim clsidCDBurn As UUID
    
   With clsidCDBurn
      .Data1 = &HFBEB8A05
      .Data2 = &HBEEE
      .Data3 = &H4442
      .Data4(0) = &H80
      .Data4(1) = &H4E
      .Data4(2) = &H40
      .Data4(3) = &H9D
      .Data4(4) = &H6C
      .Data4(5) = &H45
      .Data4(6) = &H15
      .Data4(7) = &HE9
   End With
Dim iidCDBurn As UUID
   With iidCDBurn
      .Data1 = &H3D73A659
      .Data2 = &HE5D0
      .Data3 = &H4D42
      .Data4(0) = &HAF
      .Data4(1) = &HC0
      .Data4(2) = &H51
      .Data4(3) = &H21
      .Data4(4) = &HBA
      .Data4(5) = &H42
      .Data4(6) = &H5C
      .Data4(7) = &H8D
   End With
   
   Dim hr As Long
   hr = CoCreateInstance(clsidCDBurn, Nothing, CLSCTX_INPROC_SERVER, iidCDBurn, m_cdBurn)
   If (FAILED(hr)) Then
      Err.Raise ERR_BASE + 1, App.EXEName & ".cSimpleCDBurner", "Failed to instantiate CDBurn implementation"
   End If
   m_hWndOwner = hWndOwner

End Sub

Private Function FAILED(ByVal hResult As Long) As Boolean
   FAILED = ((hResult And FAIL_BIT) = FAIL_BIT)
End Function

