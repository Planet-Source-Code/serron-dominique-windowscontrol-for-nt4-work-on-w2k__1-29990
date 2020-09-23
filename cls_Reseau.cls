VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Reseau"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type ip_option_information
    TTL             As Byte     'Time To Live
    Tos             As Byte     'Type Of Service
    Flags           As Byte     'IP header flags
    OptionsSize     As Byte     'Size in bytes of options data
    OptionsData     As Long     'Pointer to options data
End Type

'   structure that is returned from the ping to give status and error information
Private Type icmp_echo_reply
    Address         As Long             'Replying address
    Status          As Long             'Reply IP_STATUS, values as defined above
    RoundTripTime   As Long             'RTT in milliseconds
    DataSize        As Integer          'Reply data size in bytes
    Reserved        As Integer          'Reserved for system use
    DataPointer     As Long             'Pointer to the reply data
    Options         As ip_option_information    'Reply options
    Data            As String * 250     'Reply data which should be a copy of the string sent, NULL terminated
                                        ' this field length should be large enough to contain the string sent
End Type

Private Type WSADATA
  wversion As Integer
  wHighVersion As Integer
  szDescription(0 To 256) As Byte
  szSystemStatus(0 To 128) As Byte
  iMaxSockets As Integer
  iMaxUdpDg As Integer
  lpszVendorInfo As Long
End Type

'Private Type WSAData
'        wversion            As Integer
'        wHighVersion        As Integer
'         szDescription       As String * 257
'         szSystemStatus      As String * 129
'        iMaxSockets         As Integer
'        iMaxUdpDg           As Integer
''        lpVendorInfo        As String * 200
'        lpVendorInfo        As Long
'End Type
                        
Private Type hostent
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type
                        
Const IP_STATUS_BASE = 11000
Const IP_SUCCESS = 0
Const IP_WINSOCK_ASW = -1
Const IP_WINSOCK_VER = -2
Const IP_WINSOCK_MIN = -3
Const IP_BUF_TOO_SMALL = (IP_STATUS_BASE + 1)
Const IP_DEST_NET_UNREACHABLE = (IP_STATUS_BASE + 2)
Const IP_DEST_HOST_UNREACHABLE = (IP_STATUS_BASE + 3)
Const IP_DEST_PROT_UNREACHABLE = (IP_STATUS_BASE + 4)
Const IP_DEST_PORT_UNREACHABLE = (IP_STATUS_BASE + 5)
Const IP_NO_RESOURCES = (IP_STATUS_BASE + 6)
Const IP_BAD_OPTION = (IP_STATUS_BASE + 7)
Const IP_HW_ERROR = (IP_STATUS_BASE + 8)
Const IP_PACKET_TOO_BIG = (IP_STATUS_BASE + 9)
Const IP_REQ_TIMED_OUT = (IP_STATUS_BASE + 10)
Const IP_BAD_REQ = (IP_STATUS_BASE + 11)
Const IP_BAD_ROUTE = (IP_STATUS_BASE + 12)
Const IP_TTL_EXPIRED_TRANSIT = (IP_STATUS_BASE + 13)
Const IP_TTL_EXPIRED_REASSEM = (IP_STATUS_BASE + 14)
Const IP_PARAM_PROBLEM = (IP_STATUS_BASE + 15)
Const IP_SOURCE_QUENCH = (IP_STATUS_BASE + 16)
Const IP_OPTION_TOO_BIG = (IP_STATUS_BASE + 17)
Const IP_BAD_DESTINATION = (IP_STATUS_BASE + 18)
Const IP_ADDR_DELETED = (IP_STATUS_BASE + 19)
Const IP_SPEC_MTU_CHANGE = (IP_STATUS_BASE + 20)
Const IP_MTU_CHANGE = (IP_STATUS_BASE + 21)
Const IP_UNLOAD = (IP_STATUS_BASE + 22)
Const IP_ADDR_ADDED = (IP_STATUS_BASE + 23)
Const IP_GENERAL_FAILURE = (IP_STATUS_BASE + 50)
Const MAX_IP_STATUS = IP_STATUS_BASE + 50
Const IP_PENDING = (IP_STATUS_BASE + 255)
Const PING_TIMEOUT = 200

Const AF_INET = 4&  ' Pour gethostbyaddr

Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long

Private Declare Function IcmpSendEcho Lib "icmp.dll" _
  (ByVal IcmpHandle As Long, ByVal DestinationAddress As Long, _
  ByVal RequestData As String, ByVal RequestSize As Integer, _
  RequestOptions As ip_option_information, ReplyBuffer As icmp_echo_reply, _
  ByVal ReplySize As Long, ByVal Timeout As Long) As Long

'===

Private Declare Function WSAStartup Lib "wsock32" (ByVal wVersionRequested As Integer, lpWSAData As WSADATA) As Integer
Private Declare Function WSACleanup Lib "wsock32" () As Integer

Private Declare Function GetHostByName Lib "wsock32.dll" Alias "gethostbyname" _
                                              (ByVal HostName As String) As Long

Private Declare Function GetHostByAddr Lib "wsock32.dll" Alias "gethostbyaddr" _
                                (Addr As Long, _
                                 addrLen As Long, _
                                 addrType As Long) As Long


Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
                         (Dest As Any, Src As Any, ByVal cb&)

Private Declare Sub CopyMemory2 Lib "kernel32" Alias "RtlMoveMemory" _
                         (Dest As Any, ByVal Src As Long, ByVal cb As Long)

 Const WS_VERSION_REQD = &H101
 Const WS_VERSION_MAJOR = (WS_VERSION_REQD / &H100) And &HFF&
 Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
 Const MIN_SOCKETS_REQD = 1
 
Enum ListTypePingAdd
   ADD_IP = 1
   ADD_NETBIOS_DNS = 2
End Enum

Private Enum ListTypeIPAdd
   ADDTYPE_INTERN = 1        ' La valeur correspond à l'emplacement de l'adr
   ADDTYPE_EXTERN = 2        ' dans la structure de réponse.
End Enum
 
'================= pour mac adresses =============
 
Private Type NTWRKCNTRLBLCK
    ncb_command As Byte
    ncb_retcode As Byte
    ncb_lsn As Byte
    ncb_num As Byte
    ncb_buffer As Long
    ncb_length As Integer
    ncb_callname(0 To 15) As Byte
    ncb_name(0 To 15) As Byte
    ncb_rto As Byte
    ncb_sto As Byte
    lpFunc As Long
    ncb_lana_num As Byte
    ncb_cmd_cplt As Byte
    ncb_reserve(0 To 9) As Byte
    ncb_event As Long
End Type


Private Type LANA_ENUM
    length As Byte
    lana(0 To 256) As Byte
 End Type


Private Type ADAPTER_STATUS
    adapter_address(0 To 5) As Byte
    rev_major As Byte
    reserved0 As Byte
    adapter_type As Byte
    rev_minor As Byte
    duration As Integer
    frmr_recv As Integer
    frmr_xmit As Integer
    iframe_recv_err As Integer
    xmit_aborts As Integer
    xmit_success As Long
    recv_success As Long
    iframe_xmit_err As Integer
    recv_buff_unavail As Integer
    t1_timeouts As Integer
    ti_timeouts As Integer
    reserved1 As Long
    free_ncbs As Integer
    max_cfg_ncbs As Integer
    max_ncbs As Integer
    xmit_buf_unavail As Integer
    max_dgram_size As Integer
    pending_sess As Integer
    max_cfg_sess As Integer
    max_sess As Integer
    max_sess_pkt_size As Integer
    name_count As Integer
 End Type


Private Type NAME_BUFFER
   name_(0 To 15) As Byte
   name_num As Byte
   name_flags As Byte
End Type

Private Type NET_STATUS
   Adapter As ADAPTER_STATUS
   NameBuffer(30) As NAME_BUFFER
End Type
 
Const NCBENUM = &H37
Const NCBRESET = &H32
Const NCBASTAT = &H33

Private Declare Function NetBios Lib "Netapi32.dll" Alias "Netbios" _
(ByRef pncb As NTWRKCNTRLBLCK) As Byte

'================= pour net send =============

Private Declare Function NetMessageBufferSend Lib "Netapi32.dll" ( _
    ByVal sServerName$, ByVal sMsgName$, ByVal sFromName$, _
    ByVal sMessageText$, ByVal lBufferLength&) As Long



Dim StatusInitSocket As Integer   ' 0=Non init 1=OK <0=NOK (voir code err)
Dim PingAddresseIN As String      ' Mémo dernière adresse demandée
Dim PingIPF As String             ' IP Found d'apres NetBios ou DNS
Dim PingStatus_S As String        ' Message de status
Dim PingRTT As Long               ' Round Trip Time
Dim ConvStatus_S As String        ' Message de status conv IP <-> NetBios-DNS
Dim MAC_Add(0 To 19) As String    ' MAC Adresses trouvées


 

'====================================================================
'====================================================================
Private Sub Class_Initialize()
   Dim lret As Long
   
   lret = InitSocket()      ' Fonction locale
   If (lret <> 0) Then
     StatusInitSocket = -1
     PingStatus_S = IPStatusMessage(CLng(StatusInitSocket))
    Else
          StatusInitSocket = 1
   End If
End Sub

Private Sub Class_Terminate()
   If StatusInitSocket = 1 Then Call WSACleanup
End Sub

Public Function DoPing(Address As String, TypeAddress As ListTypePingAdd, _
               MessageToSend As String, PingTimeOut As Integer, pTTL As Long) As Integer

    Dim hFile       As Long             ' handle for the icmp port opened
    Dim lret        As Long             ' hold return values as required
    Dim lIPAddress  As Long
    Dim pOptions    As ip_option_information
    Dim pReturn     As icmp_echo_reply
    Dim iRet        As Integer
    Dim szTemp      As String
    
    PingAddresseIN = Trim$(Address)
    PingRTT = -1
    DoPing = 0   ' Par défaut : retour fonction = echec
    
    PingIPF = PingAddresseIN
    
   If (StatusInitSocket <> 1) Then
     ConvStatus_S = "Socket non initialisé"
     Exit Function
   End If
          
    If TypeAddress = ADD_NETBIOS_DNS Then
      szTemp = NameToAddr(Address, ADDTYPE_INTERN)
      If szTemp = "" Then 'Host Not Found
        PingIPF = PingAddresseIN
        PingStatus_S = "Hôte non trouvé"
        Exit Function
      End If
      Address = szTemp
     End If
        
    lIPAddress = ConvertIPToLong(Address)
    If (lIPAddress = 0) Then
      PingStatus_S = "Adresse incorrecte"
      Exit Function
    End If
    
    hFile = IcmpCreateFile()
    
    pOptions.TTL = pTTL
    
    lret = IcmpSendEcho(hFile, _
                        lIPAddress, _
                        MessageToSend, _
                        Len(MessageToSend), _
                        pOptions, _
                        pReturn, _
                        Len(pReturn), _
                        PingTimeOut)
    
    PingIPF = ConvertLongToIP(pReturn.Address)
    PingRTT = pReturn.RoundTripTime
    
    PingStatus_S = IPStatusMessage(pReturn.Status)
    If (pReturn.Status = IP_SUCCESS) Then DoPing = 1   ' retour fonction : réussite
                       
    lret = IcmpCloseHandle(hFile)
    If (lret <> 1) Then PingStatus_S = PingStatus_S & " / ERR FERMETURE"
    
End Function

Private Function InitSocket() As Integer
    Dim pWsaData As WSADATA
    Dim Ret As Integer
    Dim LVer As Integer
    Dim HVer As Integer
     
    Ret = WSAStartup(WS_VERSION_REQD, pWsaData)

    If Ret <> 0 Then
      InitSocket = IP_WINSOCK_ASW
      Exit Function
    End If

    LVer = (pWsaData.wversion And &HFF)
    HVer = ((pWsaData.wversion / &H100) And &HFF)
    If ((LVer < WS_VERSION_MAJOR) Or ((LVer = WS_VERSION_MAJOR) And _
                                      (HVer < WS_VERSION_MINOR))) Then
           InitSocket = IP_WINSOCK_VER
           Exit Function
    End If

    If pWsaData.iMaxSockets < MIN_SOCKETS_REQD Then
      InitSocket = IP_WINSOCK_MIN
      Exit Function
    End If

    InitSocket = 0
End Function

Private Function ConvertIPToLong(szIP As String) As Long
   '
   '   Conversion "a.b.c.d" en Val(&Hddccbbaa).
   '
    Dim szTemp              As String
    Dim iValCount           As Integer
    Dim DotValues(1 To 4)   As String
    
    szTemp = szIP
    iValCount = 0
    
    While (InStr(szTemp, ".") > 0) And (iValCount < 3)
        iValCount = iValCount + 1
        DotValues(iValCount) = Mid(szTemp, 1, InStr(szTemp, ".") - 1)
        szTemp = Mid(szTemp, InStr(szTemp, ".") + 1)
    Wend
        
    iValCount = iValCount + 1
    DotValues(iValCount) = szTemp
    
    ' Si 4 valeurs n'ont pas été trouvées => mauvaise adresse
    If iValCount <> 4 Then
        ConvertIPToLong = 0
        Exit Function
    End If
    
    ' Verif intervals
    For iValCount = 1 To 4
       If (IsNumeric(DotValues(iValCount)) = False Or Val(DotValues(iValCount)) < 0 Or Val(DotValues(iValCount)) > 255) Then DotValues(1) = "ERR"
    Next
    If DotValues(1) = "ERR" Then
      ConvertIPToLong = 0
      Exit Function
    End If
        
    ConvertIPToLong = Val("&H" & Right("00" & Hex(Val(DotValues(4))), 2) & _
                                 Right("00" & Hex(Val(DotValues(3))), 2) & _
                                 Right("00" & Hex(Val(DotValues(2))), 2) & _
                                 Right("00" & Hex(Val(DotValues(1))), 2))
End Function

Private Function ConvertLongToIP(Address As Long) As String
  Dim aa(1 To 4) As Byte
  CopyMemory aa(1), Address, 4
  ConvertLongToIP = Trim$(aa(1)) & "." & Trim$(aa(2)) & "." & _
                    Trim$(aa(3)) & "." & Trim$(aa(4))
End Function

Private Function NameToAddr(ByVal HostName As String, NumAdd As ListTypeIPAdd) As String
    Dim PointerToPointer As Long  ' Adresse structure reception
    Dim host As hostent           ' Structure de reception
    Dim pAddr As Long             ' Pointeur vers add IP
    Dim lAddr As Long             ' Add IP en long
    Dim i As Integer
    
    PointerToPointer = GetHostByName(HostName) ' Ret = pointeur vers structure

    If ((PointerToPointer = -1) Or (PointerToPointer = 0)) Then
      Exit Function
    End If
        
    i = 0
    CopyMemory host, ByVal PointerToPointer, LenB(host) ' Recup infos dans structure
    Do
       i = i + 1
       ' Récup pointeur
       CopyMemory pAddr, ByVal (host.h_addr_list + ((i - 1) * LenB(host.h_addr_list))), 4
    Loop While (pAddr <> 0) And (i <> NumAdd)
    
    If ((i = NumAdd) And (pAddr <> 0)) Then
      CopyMemory lAddr, ByVal pAddr, 4 ' Recup add IP
      NameToAddr = ConvertLongToIP(lAddr)
     Else
          NameToAddr = ""
   End If
End Function

Private Function AddrToName(Addr As String) As String
  Dim heEntry As hostent
  Dim strHost As String * 255
  Dim strTemp As String
  Dim lngRet As Long
  Dim lngIP As Long
  
  On Error GoTo AddrToNameError
  AddrToName = ""
  
  lngIP = ConvertIPToLong(Addr)
  lngRet = GetHostByAddr(lngIP, 4, AF_INET)
    
  If lngRet = 0 Then
    Exit Function
  End If
  CopyMemory2 heEntry, lngRet, Len(heEntry)
  CopyMemory2 ByVal strHost, heEntry.h_name, 255
  AddrToName = SuppSZ(strHost)
  
AddrToNameError:
End Function

Private Function IPStatusMessage(Status As Long) As String
    Dim Msg As String
    Select Case Status
        Case IP_SUCCESS: Msg = "OK"
        Case IP_BUF_TOO_SMALL: Msg = "Buffer trop petit"
        Case IP_DEST_NET_UNREACHABLE: Msg = "Réseau de destination intouchable"
        Case IP_DEST_HOST_UNREACHABLE: Msg = "Hôte de destination intouchable"
        Case IP_DEST_PROT_UNREACHABLE: Msg = "Protocol de destination intouchable"
        Case IP_DEST_PORT_UNREACHABLE: Msg = "Port de destination intouchable"
        Case IP_NO_RESOURCES: Msg = "Pas de ressources disponibles"
        Case IP_BAD_OPTION: Msg = "Err dans les options"
        Case IP_HW_ERROR: Msg = "Hôte inconnu"
        Case IP_PACKET_TOO_BIG: Msg = "Paquet trop grand"
        Case IP_REQ_TIMED_OUT: Msg = "Dépassement de délai (TimeOut)"
        Case IP_BAD_REQ: Msg = "ip bad req"
        Case IP_BAD_ROUTE: Msg = "Mauvaise route"
        Case IP_TTL_EXPIRED_TRANSIT: Msg = "ip ttl expired transit"
        Case IP_TTL_EXPIRED_REASSEM: Msg = "ip ttl expired reassem"
        Case IP_PARAM_PROBLEM: Msg = "Erreur dans les paramètres"
        Case IP_SOURCE_QUENCH: Msg = "ip source quench"
        Case IP_OPTION_TOO_BIG: Msg = "Options trop grandes"
        Case IP_BAD_DESTINATION: Msg = "Mauvaise destination"
        Case IP_ADDR_DELETED: Msg = "Adresse effacée"
        Case IP_SPEC_MTU_CHANGE: Msg = "ip spec mtu change"
        Case IP_MTU_CHANGE: Msg = "ip mtu_change"
        Case IP_UNLOAD: Msg = "ip déchargé"
        Case IP_ADDR_ADDED: Msg = "Adresse ajoutée"
        Case IP_GENERAL_FAILURE: Msg = "Erreur générale"
        Case IP_PENDING: Msg = "ip pending"
        Case PING_TIMEOUT: Msg = "Ping TimeOut"
        Case IP_WINSOCK_ASW: Msg = "Pas de réponse de Windows Sockets 32 bits !!"
        Case IP_WINSOCK_VER: Msg = "Version de Windows Socket 32 bits non supportée"
        Case IP_WINSOCK_MIN: Msg = "Minimum de sockets requis non alloué"
        Case Else: Msg = "Code de retour inconnu (" & Status & ")"
    End Select

    IPStatusMessage = Msg
End Function

Private Function GetMAC() As Boolean
   Dim NCB As NTWRKCNTRLBLCK
   Dim LanEnum As LANA_ENUM
   Dim Status As NET_STATUS
   Dim sMacAddress As String
   Dim sHex As String
   Dim bRet As Byte
   Dim sBuff As String
   Dim l As Integer
   Dim i As Integer
   Dim k As Integer
   
    NCB.ncb_command = NCBENUM
    NCB.ncb_buffer = VarPtr(LanEnum)
    NCB.ncb_length = LenB(LanEnum)
    bRet = NetBios(NCB)
    sBuff = ""
    l = LanEnum.length

'Debug.Print "==========="

    If l > 0 Then


        For k = 0 To l
            NCB.ncb_command = NCBRESET
            NCB.ncb_lana_num = LanEnum.lana(k)
            bRet = NetBios(NCB)
            
            NCB.ncb_command = NCBASTAT
            NCB.ncb_lana_num = LanEnum.lana(k)
            NCB.ncb_callname(0) = 42 'Max number of sessions


            For i = 1 To 14
                NCB.ncb_callname(i) = 32 'Max number of names 32, Use NAME_NUMBER_1
            Next i


            NCB.ncb_callname(15) = 0
            NCB.ncb_buffer = VarPtr(Status)
            NCB.ncb_length = LenB(Status)
            bRet = NetBios(NCB)
            sBuff = sBuff & "Network #" & Hex$(k) & vbCrLf & Chr$(9)
            sMacAddress = ""
            
            For i = 0 To 5
                sHex = Hex(Status.Adapter.adapter_address(i))
                If Len(sHex) = 1 Then sHex = "0" & sHex
                sMacAddress = sMacAddress & sHex
                If i <> 5 Then sMacAddress = sMacAddress + "-"
            Next i
            'Debug.Print sMacAddress
            MAC_Add(k) = sMacAddress
         Next
      End If
      
        
      

End Function

'============================================================================
'============================ Fonctions publiques ===========================
'============================================================================

Public Function DNS_To_IP(AddrIN As String, AddrOUT As String) As Integer
   DNS_To_IP = -1
   If (StatusInitSocket <> 1) Then
     ConvStatus_S = "Socket non initialisé"
     Exit Function
   End If
   
   AddrOUT = NameToAddr(AddrIN, ADDTYPE_INTERN)
   
   If AddrOUT <> "" Then
     ConvStatus_S = "Trouvé : " & AddrOUT
     DNS_To_IP = 0
    Else
         ConvStatus_S = "<Non trouvé>"
   End If
End Function

Public Function IP_To_DNS(AddrIN As String, AddrOUT As String) As Integer
   IP_To_DNS = -1
   If (StatusInitSocket <> 1) Then
     ConvStatus_S = "Socket non initialisé"
     Exit Function
   End If
   
   If (ChekIP(AddrIN) = False) Then
     ConvStatus_S = "Adresse incorrecte"
     Exit Function
   End If
   
   AddrOUT = AddrToName(AddrIN)
   
   If AddrOUT <> "" Then
     ConvStatus_S = "Trouvé : " & AddrOUT
     IP_To_DNS = 0
    Else
         ConvStatus_S = "<Non trouvé>"
   End If
End Function

Public Function GetLocalIPInt() As String
   GetLocalIPInt = NameToAddr("", ADDTYPE_INTERN)
End Function

Public Function GetLocalIPExt() As String
   GetLocalIPExt = NameToAddr("", ADDTYPE_EXTERN)
End Function

Public Function ChekIP(IP As String) As Boolean
   ChekIP = IIf(ConvertIPToLong(IP) <> 0, True, False)
End Function

Public Sub MAC_DoSearch()
   Call GetMAC
End Sub

Public Function MAC_Address(n As Integer) As String
   MAC_Address = ""
   If (n >= 0 And n <= 20) Then MAC_Address = MAC_Add(n)
End Function
'============================================================================

Public Property Get PingStatusMsg() As String
   PingStatusMsg = PingStatus_S
End Property

Public Property Get PingIPReply() As String
   PingIPReply = PingIPF
End Property

Public Property Get PingRoundTripTime() As Long
   PingRoundTripTime = PingRTT
End Property

Public Property Get ConvStatusMsg() As String
   ConvStatusMsg = ConvStatus_S
End Property
'============================================================================

Public Function EnvoiMessage(Sender As String, Dest As String, Msg As String) As Long
 
   EnvoiMessage = NetMessageBufferSend("", StrConv(Dest, vbUnicode), _
                                           StrConv(Sender, vbUnicode), _
                                           StrConv(Msg, vbUnicode), _
                                           Len(StrConv(Msg, vbUnicode)))
End Function
