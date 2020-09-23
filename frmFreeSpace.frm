VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmFreeSpace 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Drive Space"
   ClientHeight    =   6555
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   6765
   StartUpPosition =   2  'CenterScreen
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   4215
      Left            =   375
      OleObjectBlob   =   "frmFreeSpace.frx":0000
      TabIndex        =   0
      Top             =   480
      Width           =   6015
   End
   Begin VB.Frame Frame1 
      Caption         =   "Drive Stats"
      Height          =   1560
      Left            =   195
      TabIndex        =   3
      Top             =   4920
      Width           =   6375
      Begin VB.Label lblUsedDiskSpace 
         AutoSize        =   -1  'True
         Caption         =   "545 MB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4800
         TabIndex        =   13
         Top             =   360
         Width           =   660
      End
      Begin VB.Label lblFreeDiskSpace 
         AutoSize        =   -1  'True
         Caption         =   "345 MB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4800
         TabIndex        =   12
         Top             =   720
         Width           =   660
      End
      Begin VB.Label lblUsedSpace 
         AutoSize        =   -1  'True
         Caption         =   "Used Space"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3480
         TabIndex        =   11
         Top             =   360
         Width           =   1140
      End
      Begin VB.Label lblFreeSpace 
         AutoSize        =   -1  'True
         Caption         =   "Free Space"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3480
         TabIndex        =   10
         Top             =   720
         Width           =   1065
      End
      Begin VB.Label lblDiskLabel 
         AutoSize        =   -1  'True
         Caption         =   "Ryan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1440
         TabIndex        =   9
         Top             =   720
         Width           =   480
      End
      Begin VB.Label lblDiskName 
         AutoSize        =   -1  'True
         Caption         =   "Disk Label"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   960
      End
      Begin VB.Label lblFileSystem 
         AutoSize        =   -1  'True
         Caption         =   "Fat 32"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1440
         TabIndex        =   7
         Top             =   1080
         Width           =   540
      End
      Begin VB.Label lblFileType 
         AutoSize        =   -1  'True
         Caption         =   "File System"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   1050
      End
      Begin VB.Label lblDiskSize 
         AutoSize        =   -1  'True
         Caption         =   "32 GB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1440
         TabIndex        =   5
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Disk Size"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   840
      End
   End
   Begin VB.Label lblDrive 
      AutoSize        =   -1  'True
      Caption         =   "C:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3772
      TabIndex        =   2
      Top             =   0
      Width           =   420
   End
   Begin VB.Label lblDisk 
      AutoSize        =   -1  'True
      Caption         =   "Drive "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2572
      TabIndex        =   1
      Top             =   0
      Width           =   1065
   End
   Begin VB.Menu mnuFloppyDrive 
      Caption         =   "Floppy Drives"
      Begin VB.Menu mnuFloppyDrives 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu mnuHardDrives 
      Caption         =   "Hard Drives"
      Begin VB.Menu mnuDiskDrive 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu mnuCDDrive 
      Caption         =   "CD Drives"
      Begin VB.Menu mnuCDDrives 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu mnuNetDrive 
      Caption         =   "Network Drive"
      Begin VB.Menu mnuNetworkDrives 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu mnuRam 
      Caption         =   "Ram Disk"
      Begin VB.Menu mnuRamDisk 
         Caption         =   "-"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmFreeSpace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FreeSpace, TotalSpace, UsedSpace As Double
Dim UsedPercent As Integer
Dim FileSystem, DiskLabel As String

Private Sub Form_Load()
  With MSChart1
    'set background color
    .Backdrop.Fill.Style = VtFillStyleBrush
    .Backdrop.Fill.Brush.FillColor.Set 220, 0, 0
    With .Plot
      'set pie background color
      .Backdrop.Fill.Style = VtFillStyleBrush
      .Backdrop.Fill.Brush.FillColor.Set 192, 192, 192
      'set pie color
      With .SeriesCollection.Item(1)
        .SeriesMarker.Auto = False
        .DataPoints.Item(-1).Brush.Style = VtBrushStyleSolid
        .DataPoints.Item(-1).Brush.FillColor.Set 250, 250, 0
      End With
      'set pie2 color
      With .SeriesCollection.Item(2)
        .SeriesMarker.Auto = False
        .DataPoints.Item(-1).Brush.Style = VtBrushStyleSolid
        .DataPoints.Item(-1).Brush.FillColor.Set 0, 0, 250
      End With
    End With
  End With
  Call MakeMenu
  Call SystemInfo("C:\")
  Call GetDiskSpaces("C:\")
  Call DiskStats("C:\")
End Sub

Private Sub mnuNetworkDrives_Click(index As Integer)
  Dim retVal As Long
  retVal = SystemInfo(mnuNetworkDrives(index).Caption)
  If retVal = 0 Then
    MsgBox "Error reading this drive." & vbCrLf & "Try again.", _
      vbOKOnly + vbCritical, "Drive Error"
  Else
    Call GetDiskSpaces(mnuNetworkDrives(index).Caption)
    Call DiskStats(mnuNetworkDrives(index).Caption)
  End If
End Sub

Private Sub mnuRamDisk_Click(index As Integer)
  Dim retVal As Long
  retVal = SystemInfo(mnuRamDisk(index).Caption)
  If retVal = 0 Then
    MsgBox "Error reading this drive.", vbOKOnly + vbCritical, "Drive Error"
  Else
    Call GetDiskSpaces(mnuRamDisk(index).Caption)
    Call DiskStats(mnuRamDisk(index).Caption)
  End If
End Sub

Private Sub mnuDiskDrive_Click(index As Integer)
  Dim retVal As Long
  retVal = SystemInfo(mnuDiskDrive(index).Caption)
  If retVal = 0 Then
    MsgBox "Error reading this drive." & vbCrLf & "Try again.", _
      vbOKOnly + vbCritical, "Drive Error"
  Else
    Call GetDiskSpaces(mnuDiskDrive(index).Caption)
    Call DiskStats(mnuDiskDrive(index).Caption)
  End If
End Sub

Private Sub mnuFloppyDrives_Click(index As Integer)
  Dim retVal As Long
  retVal = SystemInfo(mnuFloppyDrives(index).Caption)
  If retVal = 0 Then
    MsgBox "There is no disk in this drive." & vbCrLf & _
      "Insert a disk and try again", vbOKOnly + vbCritical, "Drive Error"
  Else
    Call GetDiskSpaces(mnuFloppyDrives(index).Caption, True)
    Call DiskStats(mnuFloppyDrives(index).Caption, True)
  End If
End Sub

Private Sub mnucdDrives_Click(index As Integer)
  Dim retVal As Long
  retVal = SystemInfo(mnuCDDrives(index).Caption)
  If retVal = 0 Then
    OpenCdTray mnuCDDrives(index).Caption
    MsgBox "There is no disk in this drive." & vbCrLf & _
      "Insert a disk, wait for the drive" & vbCrLf & "light to go out and try again.", _
      vbOKOnly + vbCritical, "Drive Error"
  Else
    Call GetDiskSpaces(mnuCDDrives(index).Caption, True)
    Call DiskStats(mnuCDDrives(index).Caption, True)
    If Len(lblFreeDiskSpace.Caption) = 4 And lblFileSystem.Caption = "CDFS" Then
      lblFreeDiskSpace.Caption = "Read only disk" & vbCrLf & "or drive"
    End If
  End If
End Sub

Private Sub DiskStats(curDrive As String, Optional smallDisk As Boolean = False)
  lblDrive.Caption = Mid(curDrive, 1, 2)
  If smallDisk = True Then
    lblDiskSize.Caption = Format(TotalSpace, "###,###.##") & " MB"
  Else
    lblDiskSize.Caption = Format(TotalSpace, "###,###.##") & " GB"
  End If
  lblFreeDiskSpace.Caption = Format(FreeSpace, "###,###.##") & " MB"
  lblUsedDiskSpace.Caption = Format(UsedSpace, "###,###.##") & " MB"
  With MSChart1
    .Column = 1
    Select Case UsedPercent
      Case Is = 0
        If (FreeSpace / 1024) <> TotalSpace Then
          .RowLabel = "< 1% Full"
          UsedPercent = 1
        Else
          .RowLabel = UsedPercent & "% Full"
        End If
      Case Else
        .RowLabel = UsedPercent & "% Full"
    End Select
    .Data = UsedPercent
    .Column = 2
    .Data = 100 - UsedPercent
  End With
  If Len(DiskLabel) = 0 Then
    lblDiskLabel.Caption = "(No Label)"
  Else
    lblDiskLabel.Caption = DiskLabel
  End If
  lblFileSystem.Caption = FileSystem
End Sub

Private Sub GetDiskSpaces(strPath As String, Optional smallDisk As Boolean = False)
  Dim nFreeBytesToCaller As LargeInt
  Dim nTotalBytes As LargeInt
  Dim nTotalFreeBytes As LargeInt
  Call GetDiskFreeSpaceEx(strPath, nFreeBytesToCaller, nTotalBytes, nTotalFreeBytes)
  FreeSpace = CULong(nFreeBytesToCaller.HiDWord.Byte1, _
    nFreeBytesToCaller.HiDWord.Byte2, nFreeBytesToCaller.HiDWord.Byte3, _
    nFreeBytesToCaller.HiDWord.Byte4) * 2 ^ 32 + _
    CULong(nFreeBytesToCaller.LoDWord.Byte1, nFreeBytesToCaller.LoDWord.Byte2, _
    nFreeBytesToCaller.LoDWord.Byte3, nFreeBytesToCaller.LoDWord.Byte4)
  TotalSpace = CULong(nTotalBytes.HiDWord.Byte1, nTotalBytes.HiDWord.Byte2, _
    nTotalBytes.HiDWord.Byte3, nTotalBytes.HiDWord.Byte4) * 2 ^ 32 + _
    CULong(nTotalBytes.LoDWord.Byte1, nTotalBytes.LoDWord.Byte2, _
    nTotalBytes.LoDWord.Byte3, nTotalBytes.LoDWord.Byte4)
  UsedSpace = (TotalSpace - FreeSpace)
  UsedPercent = ((UsedSpace / TotalSpace) * 100)
  UsedSpace = ((UsedSpace / 1024) / 1024)
  FreeSpace = ((FreeSpace / 1024) / 1024)
  If smallDisk = True Then
    TotalSpace = ((TotalSpace / 1024) / 1024)
  Else
    TotalSpace = (((TotalSpace / 1024) / 1024) / 1024)
  End If
End Sub

Private Function SystemInfo(drive As String) As Long
  Dim sSysName As String * 255
  Dim sVolBuf As String * 255
  Dim lSerialNum, lSysFlags, lComponentLength As Long
  SystemInfo = GetVolumeInformation(drive, sVolBuf, 255, lSerialNum, lComponentLength, _
    lSysFlags, sSysName, 255)
  DiskLabel = Left(sVolBuf, InStr(sVolBuf, Chr(0)) - 1)
  FileSystem = Left(sSysName, InStr(sSysName, Chr(0)) - 1)
End Function

Private Sub MakeMenu()
  Dim strFloppyDrive, strHardDrive, strCDDrive, strNetDrive, strRamDrive As String
  Dim i As Integer
  Call GetDriveList

  i = 1
  If Len(FloppyDriveList) = 0 Then
    mnuFloppyDrive.Visible = False
  Else
    Do While Len(FloppyDriveList) > 0
      Load mnuFloppyDrives(i)
      mnuFloppyDrives(i).Caption = Mid(FloppyDriveList, 1, 3)
      FloppyDriveList = Mid(FloppyDriveList, 4)
      i = i + 1
    Loop
  End If
  
  i = 1
  If Len(HardDriveList) = 0 Then
    mnuHardDrives.Visible = False
  Else
    Do While Len(HardDriveList) > 0
      Load mnuDiskDrive(i)
      mnuDiskDrive(i).Caption = Mid(HardDriveList, 1, 3)
      HardDriveList = Mid(HardDriveList, 4)
      i = i + 1
    Loop
  End If
  
  i = 1
  If Len(CDRomDriveList) = 0 Then
    mnuCDDrive.Visible = False
  Else
    Do While Len(CDRomDriveList) > 0
      Load mnuCDDrives(i)
      mnuCDDrives(i).Caption = Mid(CDRomDriveList, 1, 3)
      CDRomDriveList = Mid(CDRomDriveList, 4)
      i = i + 1
    Loop
  End If
  
  i = 1
  If Len(NetworkDriveList) = 0 Then
    mnuNetDrive.Visible = False
  Else
    Do While Len(NetworkDriveList) > 0
      Load mnuNetworkDrives(i)
      mnuNetworkDrives(i).Caption = Mid(NetworkDriveList, 1, 3)
      NetworkDriveList = Mid(NetworkDriveList, 4)
      i = i + 1
    Loop
  End If
  
  i = 1
  If Len(RamDiskList) = 0 Then
    mnuRam.Visible = False
  Else
    Do While Len(RamDiskList) > 0
      Load mnuRamDisk(i)
      mnuRamDisk(i).Caption = Mid(RamDiskList, 1, 3)
      RamDiskList = Mid(RamDiskList, 4)
      i = i + 1
    Loop
  End If
End Sub
