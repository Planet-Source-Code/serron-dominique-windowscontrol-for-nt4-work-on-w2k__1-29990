Attribute VB_Name = "Mod_ShutDownNT"
Option Explicit

'======================================================================
'=========================  SHUT DOWN NT ==============================
'======================================================================

Private Const ANYSIZE_ARRAY = 1
Private Const SE_PRIVILEGE_ENABLED = &H2
Private Const SE_PRIVILEGE_ENABLED_BY_DEFAULT = &H1
Private Const SE_PRIVILEGE_USED_FOR_ACCESS = &H80000000
Private Const PRIVILEGE_SET_ALL_NECESSARY = (1)
Private Const TOKEN_ADJUST_PRIVILEGES = &H20
Private Const LB_ITEMFROMPOINT = &H1A9
Private Const nomno = &H20
Private Const SE_SHUTDOWN_NAME = "SeShutdownPrivilege"
Public Const EWX_LogOff As Long = 0
Public Const EWX_SHUTDOWN As Long = 1
Public Const EWX_REBOOT As Long = 2
Public Const EWX_FORCE As Long = 4
Public Const EWX_POWEROFF As Long = 8

Private Type luid
   lowpart As Long
   highpart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
   pLuid As luid
   Attributes As Long
End Type

Private Type TOKEN_PRIVILEGES
   PrivilegeCount As Long
   Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type

Private Declare Function ExitWindowsEx Lib "user32" _
       (ByVal uFlags As Long, _
        ByVal dwReserved As Long) As Long

Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" _
       (ByVal Tokenhandle As Long, _
        ByVal DisableAllPrivileges As Long, _
        NewState As TOKEN_PRIVILEGES, _
        ByVal BufferLength As Long, ByVal _
        PreviousState As String, _
        ReturnLength As Long) As Long

Private Declare Function OpenProcessToken Lib "advapi32.dll" _
       (ByVal ProcessHandle As Long, _
        ByVal DesiredAccess As Long, _
        Tokenhandle As Long) As Long

Private Declare Function GetCurrentProcess Lib "KERNEL32" () As Long

Private Declare Function GetLastError Lib "KERNEL32" () As Long

Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" _
        Alias "LookupPrivilegeValueA" _
       (ByVal lpSystemName As String, _
        ByVal lpName As String, _
        lpLuid As luid) As Long
        
Private Declare Function InitiateSystemShutdown Lib "advapi32.dll" _
        Alias "InitiateSystemShutdownA" _
        (ByVal lpMachineName As String, _
         ByVal lpMessage As String, _
         ByVal dwTimeout As Integer, _
         ByVal bForceAppsClosed As Long, _
         ByVal bRebootAfterShutdown As Long) As Long

Private Declare Function AbortSystemShutdown Lib "advapi32.dll" _
        Alias "AbortSystemShutdownA" (ByVal lpMachineName As String) As Long


'======================================================================

Sub ShutDownNT(Methode As Integer, FlagForce As Boolean)
   Dim lRet As Long
   Dim Tokenhandle As Long
   Dim TP As TOKEN_PRIVILEGES
   Dim lui As luid
   Dim Masque As Long

   If App.PrevInstance = True Then End
   ' Privilegien setzten um Winnt herunterfahren zu können

   lRet = OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES, Tokenhandle)
   lRet = LookupPrivilegeValue(vbNullString, SE_SHUTDOWN_NAME, lui)
   TP.PrivilegeCount = 1
   TP.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED Or _
                                 SE_PRIVILEGE_ENABLED_BY_DEFAULT Or _
                                 SE_PRIVILEGE_USED_FOR_ACCESS
   TP.Privileges(0).pLuid = lui
   lRet = AdjustTokenPrivileges(Tokenhandle, 0, TP, 0, vbNullString, 0)

   Select Case Methode
      Case 1
          Masque = EWX_POWEROFF
      Case 2
          Masque = EWX_SHUTDOWN
      Case 3
          Masque = EWX_REBOOT
      Case 4
          Masque = EWX_LogOff
      Case Else
   End Select
   If FlagForce = True Then Masque = Masque Or EWX_FORCE
   lRet = ExitWindowsEx(Masque, &HFFFF)
End Sub

Public Function InitiateShutDownNetComputer(CompName As String, MessageToUser As String, SecondsUntilShutdown As Long, ForceAppsClosed As Integer, RebootAfter As Integer) As Long
   InitiateShutDownNetComputer = InitiateSystemShutdown(CompName, MessageToUser, SecondsUntilShutdown, ForceAppsClosed, RebootAfter)
End Function


Public Function AbortShutDownNetComputer(CompName As String) As Long
   AbortShutDownNetComputer = AbortSystemShutdown(CompName)
End Function

