Attribute VB_Name = "HardDriveSpace"
Option Explicit

Private Type ULong ' Unsigned Long
  Byte1 As Byte
  Byte2 As Byte
  Byte3 As Byte
  Byte4 As Byte
End Type

Public Type LargeInt ' Large Integer
  LoDWord As ULong
  HiDWord As ULong
  LoDWord2 As ULong
  HiDWord2 As ULong
End Type

Public Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" _
   (ByVal lpRootPathName As String, FreeBytesAvailableToCaller As LargeInt, _
   TotalNumberOfBytes As LargeInt, TotalNumberOfFreeBytes As LargeInt) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" _
   (ByVal nDrive As String) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias _
   "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) _
   As Long
Public Declare Function GetVolumeInformation Lib "kernel32" Alias _
  "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer _
  As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, _
  lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal _
  lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" _
  (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal _
  uReturnLength As Long, ByVal hwndCallback As Long) As Long

Public HardDriveList, FloppyDriveList, NetworkDriveList, CDRomDriveList, RamDiskList As String
   
    
Public Function CULong(Byte1 As Byte, Byte2 As Byte, Byte3 As Byte, Byte4 As Byte) As Double
  CULong = Byte4 * 2 ^ 24 + Byte3 * 2 ^ 16 + Byte2 * 2 ^ 8 + Byte1
End Function

Public Sub GetDriveList()
  Dim AllDrives As String
  AllDrives = Space(200)
  If GetLogicalDriveStrings(Len(AllDrives), AllDrives) Then
    AllDrives = UCase(Trim(AllDrives))
  End If
  AllDrives = Left(AllDrives, Len(AllDrives) - 1)
  Do While Len(AllDrives) > 0
    Select Case GetDriveType(Mid(AllDrives, 1, 3))
      Case 2
        FloppyDriveList = FloppyDriveList & Mid(AllDrives, 1, 3)
      Case 3
        HardDriveList = HardDriveList & Mid(AllDrives, 1, 3)
      Case 4
        NetworkDriveList = NetworkDriveList & Mid(AllDrives, 1, 3)
      Case 5
        CDRomDriveList = CDRomDriveList & Mid(AllDrives, 1, 3)
      Case 6
        RamDiskList = RamDiskList & Mid(AllDrives, 1, 3)
    End Select
    AllDrives = Mid(AllDrives, 5)
  Loop
End Sub

Public Sub OpenCdTray(curDrive As String)
  mciSendString "Open " & curDrive & " Type CDAudio Alias CD", 0&, 0, 0
  mciSendString "Set CD Door Open", 0&, 0, 0
  mciSendString "Close CD", 0&, 0, 0
End Sub
