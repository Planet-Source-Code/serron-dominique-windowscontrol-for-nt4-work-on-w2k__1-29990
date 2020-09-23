VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_SysInfos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'======================================================================
'=========================  INFOS SYSTEM ==============================
'======================================================================

Private Type SYSTEM_INFO       ' Pour API
    dwOemID As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    dwReserved As Long
End Type
 
Private Type OSVERSIONINFO     ' Pour API
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Type MEMORYSTATUS      ' Pour API
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type
 
Const HKY_CPUSPEED = HKEY_LOCAL_MACHINE
Const KEY_CPUSPEED1 = "Hardware\Description\System\CentralProcessor\0"
Const KEY_CPUSPEED2 = "Hardware\Description\System\CentralProcessor\1"
Const VAL_CPUSPEED = "~MHz"

Const HKY_SYSINFO1 = HKEY_LOCAL_MACHINE
Const KEY_SYSINFO1 = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const VAL_SYSINFO1 = "Path"
Const HKY_SYSINFO2 = HKEY_LOCAL_MACHINE
Const KEY_SYSINFO2 = "SOFTWARE\Microsoft\Shared Tools Location"
Const VAL_SYSINFO2 = "MSINFO"
 
Enum ProcessorTypes
   PROCESSOR_INTEL_386 = 386
   PROCESSOR_INTEL_486 = 486
   PROCESSOR_INTEL_PENTIUM = 586
   PROCESSOR_MIPS_R4000 = 4000
   PROCESSOR_ALPHA_21064 = 21064
End Enum

Enum OsVersion
   Windows32S = 0
   Windows95 = 1
   WindowsNT = 2
End Enum

Enum StartUpModes
   START_MODE_NORMAL = 0
   START_MODE_SAFE = 1
   START_MODE_SAFENET = 2
End Enum

Const SM_CLEANBOOT = 67  ' pour fonction getsystemmetrics

'Private Type hostent
'       hName As Long
'       hAliases As Long
'       hAddrType As Integer
'       hLength As Integer
'       hAddrList As Long
'End Type

'Const WSADescription_Len = 256
'Const WSASYS_Status_Len = 128
'Public Const MIN_SOCKETS_REQD = 1
'Public Const WS_VERSION_REQD = &H101
'Public Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
'Public Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&


'Private Type WSADATA
'       wVersion As Integer
'       wHighVersion As Integer
'       szDescription(0 To WSADescription_Len) As Byte
'       szSystemStatus(0 To WSASYS_Status_Len) As Byte
'       iMaxSockets As Integer
'       iMaxUdpDg As Integer
'       lpszVendorInfo As Long
'End Type

Enum ListeTypesLecteurs
    DRIVE_UNKNOW = 0
    DRIVE_NONE = 1
    DRIVE_REMOVABLE = 2
    DRIVE_FIXED = 3
    DRIVE_REMOTE = 4
    DRIVE_CDROM = 5
    DRIVE_RAMDISK = 6
End Enum

'  Device Parameters for GetDeviceCaps()
Const DP_HORZRES = 8            '  Horizontal width in pixels
Const DP_VERTRES = 10           '  Vertical width in pixels
Const DP_BITSPIXEL = 12         '  Number of bits per pixel
Const DP_PLANES = 14            '  Number of planes


Private Declare Function API_GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" ( _
       ByVal lpRootPathName As String, _
       ByVal lpVolumeNameBuffer As String, _
       ByVal nVolumeNameSize As Long, _
       lpVolumeSerialNumber As Long, _
       lpMaximumComponentLength As Long, _
       lpFileSystemFlags As Long, _
       ByVal lpFileSystemNameBuffer As String, _
       ByVal nFileSystemNameSize As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (LpVersionInformation As OSVERSIONINFO) As Long
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetTickCount& Lib "kernel32" ()
 
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, _
                                                 ByVal hdc As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, _
                                                    ByVal nIndex As Long) As Long
'================================================================================
'================================================================================

Dim VerInfo As OSVERSIONINFO
Dim SysInfo As SYSTEM_INFO
Dim MemSts As MEMORYSTATUS

Dim FlagLectureOS As Integer
Dim FlagLectureSystem As Integer

'==============================================================
'==============================================================

'---------- Infos OS ----------

Private Sub LectureInfosOS()
   Dim Ret As Long
   
   
   
   VerInfo.dwOSVersionInfoSize = Len(VerInfo)
   Ret = GetVersionEx(VerInfo)
   FlagLectureOS = 1
End Sub

Public Property Get OS_PlatformID() As Long
   If FlagLectureOS = 0 Then Call LectureInfosOS
   OS_PlatformID = VerInfo.dwPlatformId
End Property

Public Property Get OS_CSDVersion() As String
   If FlagLectureOS = 0 Then Call LectureInfosOS
   OS_CSDVersion = VerInfo.szCSDVersion
End Property

Public Property Get OS_BuildNumber() As Long
   If FlagLectureOS = 0 Then Call LectureInfosOS
   OS_BuildNumber = VerInfo.dwBuildNumber
End Property

Public Property Get OS_MajorVersion() As Long
   If FlagLectureOS = 0 Then Call LectureInfosOS
   OS_MajorVersion = VerInfo.dwMajorVersion
End Property

Public Property Get OS_MinorVersion() As Long
   If FlagLectureOS = 0 Then Call LectureInfosOS
   OS_MinorVersion = VerInfo.dwMinorVersion
End Property

Public Property Get OS_StartUpMode() As Long
   OS_StartUpMode = GetSystemMetrics(SM_CLEANBOOT)
End Property

Public Property Get OS_LogginDuration() As Long
   OS_LogginDuration = Int(GetTickCount&() / 1000)
End Property

'---------- Infos System ----------

Private Sub LectureInfosSystem()
   GetSystemInfo SysInfo
   FlagLectureSystem = 1
End Sub

Public Property Get Sys_ProcessorType() As Long
   If FlagLectureSystem = 0 Then Call LectureInfosSystem
   Sys_ProcessorType = SysInfo.dwProcessorType
End Property

Public Property Get Sys_NumberOfProcessor() As Long
   If FlagLectureSystem = 0 Then Call LectureInfosSystem
   Sys_NumberOfProcessor = SysInfo.dwNumberOfProcessors
End Property

Public Property Get Sys_ActiveProcessorMask() As Long
   If FlagLectureSystem = 0 Then Call LectureInfosSystem
   Sys_ActiveProcessorMask = SysInfo.dwActiveProcessorMask
End Property

Public Property Get Sys_Processor1Speed() As Integer
   Sys_Processor1Speed = LectureRegistre(HKY_CPUSPEED, KEY_CPUSPEED1, VAL_CPUSPEED, 0, 0)
End Property

Public Property Get Sys_Processor2Speed() As Integer
   Sys_Processor2Speed = LectureRegistre(HKY_CPUSPEED, KEY_CPUSPEED2, VAL_CPUSPEED, 0, 0)
End Property

'---------- Mémoire ----------

Public Property Get Mem_PhysTot() As Long
   Call GlobalMemoryStatus(MemSts)
   Mem_PhysTot = MemSts.dwTotalPhys
End Property

Public Property Get Mem_PhysAvl() As Long
   Call GlobalMemoryStatus(MemSts)
   Mem_PhysAvl = MemSts.dwAvailPhys
End Property

Public Property Get Mem_VirtTot() As Long
   Call GlobalMemoryStatus(MemSts)
   Mem_VirtTot = MemSts.dwTotalVirtual
End Property

Public Property Get Mem_VirtAvl() As Long
   Call GlobalMemoryStatus(MemSts)
   Mem_VirtAvl = MemSts.dwAvailVirtual
End Property

'---------- Réseau ----------

Public Property Get Net_UserName() As String
   Dim Txt As String
   Txt = String$(200, 0)
   Call GetUserName(Txt, 200)       ' API
   Net_UserName = SuppSZ(Txt)
End Property

Public Property Get Net_ComputerName() As String
   Dim Txt As String
   Txt = String$(200, 0)
   Call GetComputerName(Txt, 200)   ' API
   Net_ComputerName = SuppSZ(Txt)
End Property

'---------- Affichage ----------

Public Property Get Scr_Width() As Long
   Dim hdccaps As Long
   Dim hdesktopwnd As Long
   Dim Ret As Long
  
   hdccaps = GetDC(hdesktopwnd)
      Scr_Width = GetDeviceCaps(hdccaps, DP_HORZRES)
   Ret = ReleaseDC(hdesktopwnd, hdccaps)
End Property

Public Property Get Scr_Height() As Long
   Dim hdccaps As Long
   Dim hdesktopwnd As Long
   Dim Ret As Long
  
   hdccaps = GetDC(hdesktopwnd)
      Scr_Height = GetDeviceCaps(hdccaps, DP_VERTRES)
   Ret = ReleaseDC(hdesktopwnd, hdccaps)
End Property

Public Property Get Scr_Bits() As Integer
   Dim hdccaps As Long
   Dim hdesktopwnd As Long
   Dim Ret As Long
   Dim BitsPerPixels As Long
  
   hdccaps = GetDC(hdesktopwnd)
   
   ' Si 1 bit, nb cols dépend nb plans
   BitsPerPixels = GetDeviceCaps(hdccaps, DP_BITSPIXEL)
   Select Case BitsPerPixels
      Case 1
          Scr_Bits = GetDeviceCaps(hdccaps, DP_PLANES)
      Case Else
          Scr_Bits = BitsPerPixels
   End Select
   
   Ret = ReleaseDC(hdesktopwnd, hdccaps)
End Property

Public Property Get Scr_NbColors() As String
  Scr_NbColors = 2 ^ Scr_Bits
End Property

'---------- Disque ----------

Public Property Get GetDriveType(Drive As String) As ListeTypesLecteurs
   GetDriveType = API_GetDriveType(Drive)
End Property

Public Sub GetVolumeInfos(Drive As String, VolumeName As String, _
                               SerialNumber As Long, SystemName As String)
   Dim VolName As String
   Dim SysName As String
   Dim SystemFlag As Long
   Dim MCL As Long
   Dim Ret As Long
   
   VolName = String$(100, 0): SysName = String$(100, 0): SerialNumber = 0
   Ret = GetVolumeInformation(Drive, VolName, 100, SerialNumber, MCL, _
                              SystemFlag, SysName, 100)
   VolumeName = SuppSZ(VolName)
   SystemName = SuppSZ(SysName)
End Sub



'==========================================

Public Sub StartMSSysInfo()
    Dim SysInfoPath As String
    On Error GoTo StartMSSysInfo_Err
        
    ' Tentative lecture clef contenant chemin complet
    SysInfoPath = LectureRegistre(HKY_SYSINFO1, KEY_SYSINFO1, VAL_SYSINFO1, "--", "--")
    If (SysInfoPath = "--") Then
      ' Tentative lecture clef contenant chemin partiel
      SysInfoPath = LectureRegistre(HKY_SYSINFO2, KEY_SYSINFO2, VAL_SYSINFO2, "--", "--")
      If (SysInfoPath = "--") Then
        GoTo StartMSSysInfo_Err
        Else
            If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
              SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
             Else
                 GoTo StartMSSysInfo_Err
            End If
      End If
    End If
         
    Call Shell(SysInfoPath, vbNormalFocus)
    Exit Sub

StartMSSysInfo_Err:
    MsgBox "Microsoft Infos Système non trouvé !!", vbOKOnly
End Sub


'Public Function GetHdiskSpace(DrivePath As String, RoundtoMB As Boolean) As Double
'  Dim Drive As String
'  Dim SectorsPerCluster As Long, BytesPerSector As Long
'  Dim NumberOfFreeClusters As Long, TotalClusters As Long, Sts As Long
'  Dim DS
'  Drive = Left(Trim(DrivePath), 1) & ":\"     ' Ensure path is at the root.
'  Sts = GetDiskFreeSpace(Drive, SectorsPerCluster, BytesPerSector, NumberOfFreeClusters, TotalClusters)
'  If Sts <> 0 Then
'    GetHdiskSpace = SectorsPerCluster * BytesPerSector * NumberOfFreeClusters
'    DS = Format$(GetHdiskSpace, "###,###")
'    If RoundtoMB = True Then
'    GetHdiskSpace = DS / 1024 / 1024
'    End If
'  Else
'    GetHdiskSpace = -1
'  End If
'End Function

Private Function SuppSZ(Txt As String) As String
   Dim Pos As Integer
   
   Pos = InStr(Txt, Chr$(0))
   If (Pos > 0) Then
     SuppSZ = Left$(Txt, Pos - 1)
     Else
          SuppSZ = Trim$(Txt)
   End If
End Function


