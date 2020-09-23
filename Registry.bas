Attribute VB_Name = "Mod_Registry"
Option Explicit

'Special Area
Public Enum RegTypeNames
      REG_NONE = 0                       ' No value type
      REG_SZ = 1                         ' Unicode nul terminated string
      REG_EXPAND_SZ = 2                  ' Unicode nul terminated string
      REG_BINARY = 3                     ' Free form binary
      REG_DWORD = 4                      ' 32-bit number
      REG_DWORD_LITTLE_ENDIAN = 4        ' 32-bit number (same as REG_DWORD)
      REG_DWORD_BIG_ENDIAN = 5           ' 32-bit number
      REG_LINK = 6                       ' Symbolic Link (unicode)
      REG_MULTI_SZ = 7                   ' Multiple Unicode strings
      REG_RESOURCE_LIST = 8              ' Resource list in the resource map
      REG_FULL_RESOURCE_DESCRIPTOR = 9   ' Resource list in the hardware description
      REG_RESOURCE_REQUIREMENTS_LIST = 10
End Enum

Public Enum hKeyNames
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_CURRENT_CONFIG = &H80000005
End Enum

Type TypeNetResource       ' Pour connexion distante
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As String
    lpRemoteName As String
    lpComment As String
    lpProvider As String
End Type

Type TypeNetRegHandles
    Users As Long
    CurrentUser As Long
    LocalMachine As Long
End Type
Dim NetRegHandles As TypeNetRegHandles

Public Const ERROR_NONE = 0
Public Const ERROR_BADDB = 1
Public Const ERROR_BADKEY = 2
Public Const ERROR_CANTOPEN = 3
Public Const ERROR_CANTREAD = 4
Public Const ERROR_CANTWRITE = 5
Public Const ERROR_OUTOFMEMORY = 6
Public Const ERROR_ARENA_TRASHED = 7
Public Const ERROR_ACCESS_DENIED = 8
Public Const ERROR_INVALID_PARAMETERS = 87
Public Const ERROR_NO_MORE_ITEMS = 259

Public Const KEY_ALL_ACCESS = &H3F

Public Const REG_OPTION_NON_VOLATILE = 0

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal HKey As Long) As Long
Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal HKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal HKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal HKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal HKey As Long, ByVal lpValueName As String) As Long

'===========
Declare Function RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" (ByVal lpMachineName As String, ByVal HKey As Long, phkResult As Long) As Long
Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (lpNetResource As TypeNetResource, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
Declare Function WNetCancelConnection2 Lib "mpr.dll" Alias "WNetCancelConnection2A" (ByVal lpName As String, ByVal dwFlags As Long, ByVal fForce As Long) As Long

Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long

Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long



   Private Function EcritureValeurReg(ByVal HKey As Long, sValueName As String, lType As Long, vValue As Variant) As Long
       Dim lValue As Long
       Dim sValue As String
       
       Select Case lType
           Case REG_SZ
               sValue = vValue & Chr$(0)
               EcritureValeurReg = RegSetValueExString(HKey, sValueName, 0&, lType, sValue, Len(sValue))
           Case REG_DWORD
               lValue = vValue
               EcritureValeurReg = RegSetValueExLong(HKey, sValueName, 0&, lType, lValue, 4)
           End Select
   End Function

   Private Function LectureValeurReg(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
       Dim i As Integer
       Dim cch As Long
       Dim lrc As Long
       Dim lType As Long
       Dim lValue As Long
       Dim szValue As String
       Dim sNewVal As String

       On Error GoTo LectureValeurRegError

       ' Determine taille et type de la valeur
       lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
       If lrc <> ERROR_NONE Then Error 5

       vValue = Empty
       Select Case lType
           Case REG_SZ, REG_EXPAND_SZ
               szValue = String(cch, 0)
               lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, szValue, cch)
               If lrc = ERROR_NONE Then vValue = Left$(szValue, cch - 1)
           
           Case REG_DWORD, REG_DWORD_BIG_ENDIAN
               lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
               If lrc = ERROR_NONE Then vValue = lValue
           
           Case REG_BINARY
               szValue = String(cch, 0)
               lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, szValue, cch)
               If lrc = ERROR_NONE Then
                 szValue = Left$(szValue, cch - 1)
                 vValue = ""
                 For i = 1 To Len(szValue)
                    sNewVal = Format(Hex(Asc(Mid(szValue, i, 1))), "00")
                    If Len(sNewVal) = 1 Then
                        sNewVal = "0" & sNewVal
                    End If
                    vValue = vValue & sNewVal & " "
                 Next
                 vValue = Trim(vValue)
               End If
           
           Case Else
               lrc = -1
               MsgBox ("Type de clef inconnu !! (" & lType & ")")
       End Select

LectureValeurRegExit:
       LectureValeurReg = lrc
       Exit Function
LectureValeurRegError:
       Resume LectureValeurRegExit
   End Function

Function HandleReel(HKey As hKeyNames) As Long
    ' Entree : HKey local
    ' Sortie : HKey réel : local ou distant si existe
    HandleReel = HKey
    Select Case HKey
       Case HKEY_USERS
           If (NetRegHandles.Users <> 0) Then HandleReel = NetRegHandles.Users
       Case HKEY_CURRENT_USER
           If (NetRegHandles.CurrentUser <> 0) Then HandleReel = NetRegHandles.CurrentUser
       Case HKEY_LOCAL_MACHINE
           If (NetRegHandles.LocalMachine <> 0) Then HandleReel = NetRegHandles.LocalMachine
       Case Else
           MsgBox ("HandleReel : clef non prévue : " & HKey)
    End Select
End Function

Function LectureRegistre(HKeyParam As hKeyNames, Chemin As String, Key As String, DftNoValue As Variant, DftNoKey As Variant) As Variant
        Dim Ret1, Ret2, Ret3 As Long   'result of the API functions
        Dim HKeyOpened As Long         'handle of opened key
        Dim HKey As Long
        Dim vValeur As Variant         'setting of queried value
      
        HKey = HandleReel(HKeyParam)
           
        Ret1 = RegOpenKeyEx(HKey, Chemin, 0, KEY_ALL_ACCESS, HKeyOpened)
        Ret2 = LectureValeurReg(HKeyOpened, Key, vValeur)
        Ret3 = RegCloseKey(HKeyOpened)
        
        If IsEmpty(vValeur) Then vValeur = DftNoValue
        If (Ret1 <> 0) Then vValeur = DftNoKey
        LectureRegistre = vValeur
        
End Function

Public Sub EcritureRegistre(HKeyParam As hKeyNames, Path As String, Key As String, TypeSetting As RegTypeNames, Setting As Variant)
    Dim Tmp, Ret1, Ret2, Ret3 As Long       'result of the SetValueEx function
    Dim HKeyOpened As Long          'handle of open key
    Dim HKey As Long
    Dim Msg As String

    HKey = HandleReel(HKeyParam)

    Ret1 = RegCreateKeyEx(HKey, Path, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, HKeyOpened, Tmp)
    Ret2 = EcritureValeurReg(HKeyOpened, Key, TypeSetting, Setting)
    Ret3 = RegCloseKey(HKeyOpened)
    
    If (Ret1 <> ERROR_NONE) Then
      Msg = "ERREUR DE CREATION DE CLEF" & vbCrLf & vbCrLf & "Clef : " & Path & "\" & Key & _
             vbCrLf & vbCrLf & "Code d'erreur : " & Ret1
      Call MsgBox(Msg, vbCritical + vbOKOnly + vbApplicationModal, "Erreur d'écriture de la Base des registres")
      Exit Sub
    End If
    If (Ret2 <> ERROR_NONE) Then
      Msg = "ERREUR D'ECRITURE DE CLEF" & vbCrLf & vbCrLf & "Clef : " & Path & "\" & Key & _
            vbCrLf & vbCrLf & "Code d'erreur : " & Ret1
      Call MsgBox(Msg, vbCritical + vbOKOnly + vbApplicationModal, "Erreur d'écriture de la Base des registres")
    End If
End Sub

Public Function EffaceRegistre(HKeyParam As hKeyNames, Path As String, Key As String) As Boolean
  Dim lRetVal As Long       'result of the SetValueEx function
  Dim HKeyOpened As Long          'handle of open key
  Dim HKey As Long
  Dim Msg As String

  HKey = HandleReel(HKeyParam)

  lRetVal = RegCreateKeyEx(HKey, Path, 0&, vbNullString, REG_OPTION_NON_VOLATILE, _
                           KEY_ALL_ACCESS, 0&, HKeyOpened, lRetVal)
  lRetVal = RegDeleteValue(HKeyOpened, Key)
  RegCloseKey (HKeyOpened)

  EffaceRegistre = IIf(lRetVal = ERROR_NONE, True, False)
End Function

'====================
'====================

Public Sub CloseAllRegKey()
   RegCloseKey (NetRegHandles.Users)
   RegCloseKey (NetRegHandles.CurrentUser)
   RegCloseKey (NetRegHandles.LocalMachine)
   NetRegHandles.Users = 0
   NetRegHandles.CurrentUser = 0
   NetRegHandles.LocalMachine = 0
End Sub

Public Function GetIPCConnection(RemoteComputer As String, UserName As String, Password As String) As Long

   Dim Ret As Long
   Dim MyNetStruct As TypeNetResource
   Dim RemoteHKey As Long

   NetRegHandles.Users = 0
   NetRegHandles.CurrentUser = 0
   NetRegHandles.LocalMachine = 0

   MyNetStruct.dwType = 0
   MyNetStruct.lpLocalName = "" & Chr$(0)
   MyNetStruct.lpRemoteName = "\\" + Trim$(RemoteComputer) + "\ipc$" & Chr$(0)
   MyNetStruct.lpProvider = "" & Chr$(0)

   Ret = WNetAddConnection2(MyNetStruct, Password & Chr$(0), UserName & Chr$(0), 0)
   GetIPCConnection = Ret
   If Ret <> 0 Then Exit Function
   
   Ret = RegConnectRegistry(Trim$(RemoteComputer), HKEY_USERS, RemoteHKey)
   NetRegHandles.Users = RemoteHKey
   GetIPCConnection = Ret
   If Ret <> 0 Then Exit Function
   
   Ret = RegConnectRegistry(Trim$(RemoteComputer), HKEY_CURRENT_USER, RemoteHKey)
   NetRegHandles.CurrentUser = RemoteHKey
   GetIPCConnection = Ret
   If Ret <> 0 Then Exit Function
   
   Ret = RegConnectRegistry(Trim$(RemoteComputer), HKEY_LOCAL_MACHINE, RemoteHKey)
   NetRegHandles.LocalMachine = RemoteHKey
   GetIPCConnection = Ret
   If Ret <> 0 Then Exit Function
   

End Function

Public Function DecodeSystemError(CodeErr As Long) As String

  Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000

  Dim Ret As Long
  Dim Buffer As String * 256
  
  Ret = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, CodeErr, 0, Buffer, Len(Buffer), 0)
  If Ret > 0 Then
    DecodeSystemError = Left$(Buffer, Ret - 2) & " (" & Format(CodeErr) & ")"
   Else
       DecodeSystemError = "Erreur inconnue " & Format(CodeErr)
  End If

End Function

