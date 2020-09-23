Attribute VB_Name = "Mod_LectureEcriture"

'==================================================================================
'======================== PROCEDURES LECTURE-ECRITURE =============================
'==================================================================================
Option Explicit

Const KEY_LOG = "Software\Microsoft\Windows NT\CurrentVersion\WinLogon"
Const KEY_CVE = "Software\Microsoft\Windows NT\CurrentVersion"
Const KEY_SYS = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
Const KEY_EXP = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
Const KEY_NET = "Software\Microsoft\Windows\CurrentVersion\Policies\Network"
Const KEY_DSK = "Control Panel\Desktop"
Const KEY_WIM = "Control Panel\Desktop\WindowMetrics"
Const KEY_KBU = "Control Panel\Keyboard"
Const KEY_KBD = ".Default\Control Panel\Keyboard"

Public Const HKY_SERNUM = HKEY_LOCAL_MACHINE
Public Const KEY_SERNUM = KEY_CVE
Public Const VAL_SERNUM = "ProductId"
Public Const VAL_REGOWN = "RegisteredOwner"
Public Const VAL_REGORG = "RegisteredOrganization"

Public Const HKY_ALOGIN = HKEY_LOCAL_MACHINE
Public Const KEY_ALOGIN = KEY_LOG
Public Const VAL_ALENBD = "AutoAdminLogon"
Public Const VAL_ALUSER = "DefaultUserName"
Public Const VAL_ALDOMN = "DefaultDomainName"
Public Const VAL_ALPSWD = "DefaultPassword"

Public Const HKY_MSGLOG = HKEY_LOCAL_MACHINE
Public Const KEY_MSGLOG = KEY_LOG
Public Const VAL_MSLCAP = "LegalNoticeCaption"
Public Const VAL_MSLTXT = "LegalNoticeText"

Public Const HKY_POPLOG = HKEY_LOCAL_MACHINE
Public Const KEY_POPLOG = KEY_LOG
Public Const VAL_POPLOG = "LogonPrompt"

Public Const HKY_IKIUSR = HKEY_CURRENT_USER
Public Const KEY_IKIUSR = KEY_KBU
Public Const VAL_IKIUSR = "InitialKeyboardIndicators"

Public Const HKY_IKIDFT = HKEY_USERS
Public Const KEY_IKIDFT = KEY_KBD
Public Const VAL_IKIDFT = "InitialKeyboardIndicators"

Public Const HKY_ISDLOG = HKEY_LOCAL_MACHINE
Public Const KEY_ISDLOG = KEY_LOG
Public Const VAL_ISDLOG = "ShutdownWithoutLogon"

Public Const HKY_CPLAFF = HKEY_CURRENT_USER
Public Const KEY_CPLAFF = KEY_SYS
Public Const VAL_CPAGEN = "NoDispCPL"
Public Const VAL_CPAAPP = "NoDispAppearancePage"
Public Const VAL_CPABKG = "NoDispBackGroundPage"
Public Const VAL_CPASCR = "NoDispScrSavPage"
Public Const VAL_CPASET = "NoDispSettingsPage"

Public Const HKY_INHDSK = HKEY_CURRENT_USER
Public Const KEY_INHDSK = KEY_EXP
Public Const VAL_INHDSK = "NoDeskTop"

Public Const HKY_ICONET = HKEY_CURRENT_USER
Public Const KEY_ICONET = KEY_EXP
Public Const VAL_ICONET = "NoNetHood"

Public Const HKY_ENTNET = HKEY_CURRENT_USER
Public Const KEY_ENTNET = KEY_NET
Public Const VAL_ENTNET = "NoEntireNetwork"

Public Const HKY_NETWKC = HKEY_CURRENT_USER
Public Const KEY_NETWKC = KEY_NET
Public Const VAL_NETWKC = "NoWorkgroupContents"

Public Const HKY_ICOWEB = HKEY_CURRENT_USER
Public Const KEY_ICOWEB = KEY_EXP
Public Const VAL_ICOWEB = "NoInternetIcon"

Public Const HKY_CLKDRT = HKEY_CURRENT_USER
Public Const KEY_CLKDRT = KEY_EXP
Public Const VAL_CLKDRT = "NoViewContextMenu"

Public Const HKY_IMPAJT = HKEY_CURRENT_USER
Public Const KEY_IMPAJT = KEY_EXP
Public Const VAL_IMPAJT = "NoAddPrinter"

Public Const HKY_IMPSUP = HKEY_CURRENT_USER
Public Const KEY_IMPSUP = KEY_EXP
Public Const VAL_IMPSUP = "NoDeletePrinter"

Public Const HKY_STPWIN = HKEY_CURRENT_USER
Public Const KEY_STPWIN = KEY_EXP
Public Const VAL_STPWIN = "NoClose"

Public Const HKY_LOGOFF = HKEY_CURRENT_USER
Public Const KEY_LOGOFF = KEY_EXP
Public Const VAL_LOGOFF = "NoLogOff"

Public Const HKY_PRGCMN = HKEY_CURRENT_USER
Public Const KEY_PRGCMN = KEY_EXP
Public Const VAL_PRGCMN = "NoCommonGroups"

Public Const HKY_CMDRUN = HKEY_CURRENT_USER
Public Const KEY_CMDRUN = KEY_EXP
Public Const VAL_CMDRUN = "NoRun"

Public Const HKY_CMDFND = HKEY_CURRENT_USER
Public Const KEY_CMDFND = KEY_EXP
Public Const VAL_CMDFND = "NoFind"

Public Const HKY_CNFGEN = HKEY_CURRENT_USER
Public Const KEY_CNFGEN = KEY_EXP
Public Const VAL_CNFGEN = "NoSetFolders"

Public Const HKY_CNFTKB = HKEY_CURRENT_USER
Public Const KEY_CNFTKB = KEY_EXP
Public Const VAL_CNFTKB = "NoSetTaskbar"

Public Const HKY_LOCKST = HKEY_CURRENT_USER
Public Const KEY_LOCKST = KEY_SYS
Public Const VAL_LOCKST = "DisableLockWorkstation"

Public Const HKY_TSKMGR = HKEY_CURRENT_USER
Public Const KEY_TSKMGR = KEY_SYS
Public Const VAL_TSKMGR = "DisableTaskMgr"

Public Const HKY_CHGPWD = HKEY_CURRENT_USER
Public Const KEY_CHGPWD = KEY_SYS
Public Const VAL_CHGPWD = "DisableChangePassword"

Public Const HKY_HIDDRV = HKEY_CURRENT_USER
Public Const KEY_HIDDRV = KEY_EXP
Public Const VAL_HIDDRV = "NoDrives"

Public Const HKY_SIZICD = HKEY_CURRENT_USER
Public Const KEY_SIZICD = KEY_WIM
Public Const VAL_SIZICD = "Shell Icon Size"

Public Const HKY_SIZICS = HKEY_CURRENT_USER
Public Const KEY_SIZICS = KEY_WIM
Public Const VAL_SIZICS = "Shell Small Icon Size"

Public Const HKY_TMRMNU = HKEY_CURRENT_USER
Public Const KEY_TMRMNU = KEY_DSK
Public Const VAL_TMRMNU = "MenuShowDelay"

' Constantes internes : Valeurs arbitraires
Public Const REGSTS_INOKEY = -9999
Public Const REGSTS_INOVALUE = -9998
Public Const REGSTS_IINVVAL = -9997
Public Const REGSTS_SZNOKEY = "<Cle non trouvée!!>"
Public Const REGSTS_SZNOVALUE = "<Valeur non trouvée!!>"

' ======== NUMERO DE SERIE ========
Public Property Get WindowsSerialNumber() As String
   Dim Ret As String
   Ret = LectureRegistre(HKY_SERNUM, KEY_SERNUM, VAL_SERNUM, REGSTS_SZNOVALUE, REGSTS_SZNOKEY)
   If (InStr(Ret, "OEM") > 0) Then
     WindowsSerialNumber = Format(Ret, "@@@@@-@@@-@@@@@@@-@@@@@")
     Else
          WindowsSerialNumber = Ret
   End If
End Property

' ======== PROPRIETAIRE WINDOWS - UTILISATEUR ========
Public Property Get RegisteredOwner() As String
   RegisteredOwner = LectureRegistre(HKY_SERNUM, KEY_SERNUM, VAL_REGOWN, REGSTS_SZNOVALUE, REGSTS_SZNOKEY)
End Property

' ======== PROPRIETAIRE WINDOWS - SOCIETE ========
Public Property Get RegisteredOrganization() As String
   RegisteredOrganization = LectureRegistre(HKY_SERNUM, KEY_SERNUM, VAL_REGORG, REGSTS_SZNOVALUE, REGSTS_SZNOKEY)
End Property

' ======== AUTO LOGIN ========
Public Property Get RegALActif() As Integer
   Dim Ret As String
   Ret = LectureRegistre(HKY_ALOGIN, KEY_ALOGIN, VAL_ALENBD, "0", REGSTS_SZNOKEY)
   Select Case Ret
      Case "0": RegALActif = 0
      Case "1": RegALActif = 1
      Case REGSTS_SZNOVALUE: RegALActif = REGSTS_INOVALUE
      Case REGSTS_SZNOKEY: RegALActif = REGSTS_INOKEY
      Case Else: RegALActif = REGSTS_IINVVAL
   End Select
End Property
Public Property Let RegALActif(Etat As Integer)
   If (Etat = 1) Then
     Call EcritureRegistre(HKY_ALOGIN, KEY_ALOGIN, VAL_ALENBD, REG_SZ, "1")
    Else
       Call EffaceRegistre(HKY_ALOGIN, KEY_ALOGIN, VAL_ALENBD)
   End If
End Property

Public Property Get RegALUserName() As String
   RegALUserName = LectureRegistre(HKY_ALOGIN, KEY_ALOGIN, VAL_ALUSER, REGSTS_SZNOVALUE, REGSTS_SZNOKEY)
End Property
Public Property Let RegALUserName(Txt As String)
     Call EcritureRegistre(HKY_ALOGIN, KEY_ALOGIN, VAL_ALUSER, REG_SZ, Txt)
End Property

Public Property Get RegALDomain() As String
   RegALDomain = LectureRegistre(HKY_ALOGIN, KEY_ALOGIN, VAL_ALDOMN, REGSTS_SZNOVALUE, REGSTS_SZNOKEY)
End Property
Public Property Let RegALDomain(Txt As String)
   If (Txt <> "") Then
     Call EcritureRegistre(HKY_ALOGIN, KEY_ALOGIN, VAL_ALDOMN, REG_SZ, Txt)
   End If
End Property

Public Property Get RegALPassword() As String
   RegALPassword = LectureRegistre(HKY_ALOGIN, KEY_ALOGIN, VAL_ALPSWD, "", REGSTS_SZNOKEY)
End Property
Public Property Let RegALPassword(Txt As String)
   If (Txt <> FLG_SZ_AEFFACER) Then
     Call EcritureRegistre(HKY_ALOGIN, KEY_ALOGIN, VAL_ALPSWD, REG_SZ, Txt)
     Else
       Call EffaceRegistre(HKY_ALOGIN, KEY_ALOGIN, VAL_ALPSWD)
   End If
End Property

' ======== POPUP START UP ========
Public Property Get RegPopStartUpTitre() As String
   RegPopStartUpTitre = LectureRegistre(HKY_MSGLOG, KEY_MSGLOG, VAL_MSLCAP, REGSTS_SZNOVALUE, REGSTS_SZNOKEY)
End Property
Public Property Let RegPopStartUpTitre(Txt As String)
     Call EcritureRegistre(HKY_MSGLOG, KEY_MSGLOG, VAL_MSLCAP, REG_SZ, Txt)
End Property

Public Property Get RegPopStartUpText() As String
   RegPopStartUpText = LectureRegistre(HKY_MSGLOG, KEY_MSGLOG, VAL_MSLTXT, REGSTS_SZNOVALUE, REGSTS_SZNOKEY)
End Property
Public Property Let RegPopStartUpText(Txt As String)
   Call EcritureRegistre(HKY_MSGLOG, KEY_MSGLOG, VAL_MSLTXT, REG_SZ, Txt)
End Property

' ======== POPUP LOGON ========
Public Property Get RegLogonPrompt() As String
   RegLogonPrompt = LectureRegistre(HKY_POPLOG, KEY_POPLOG, VAL_POPLOG, REGSTS_SZNOVALUE, REGSTS_SZNOKEY)
End Property
Public Property Let RegLogonPrompt(Txt As String)
   Call EcritureRegistre(HKY_POPLOG, KEY_POPLOG, VAL_POPLOG, REG_SZ, Txt)
End Property


' ======== INITIAL KEYBOARD INDICATORS ========
Public Property Get RegInitialKeyboardIndicatorsDefault() As Integer
   Dim Ret As String
   Ret = LectureRegistre(HKY_IKIDFT, KEY_IKIDFT, VAL_IKIDFT, REGSTS_SZNOVALUE, REGSTS_SZNOKEY)
   Select Case Ret
      Case 0 To 7: RegInitialKeyboardIndicatorsDefault = Val(Ret)
      Case REGSTS_SZNOVALUE: RegInitialKeyboardIndicatorsDefault = REGSTS_INOVALUE
      Case REGSTS_SZNOKEY: RegInitialKeyboardIndicatorsDefault = REGSTS_INOKEY
      Case Else: RegInitialKeyboardIndicatorsDefault = REGSTS_IINVVAL
   End Select
End Property
Public Property Let RegInitialKeyboardIndicatorsDefault(V As Integer)
   Call EcritureRegistre(HKY_IKIDFT, KEY_IKIDFT, VAL_IKIDFT, REG_SZ, V)
End Property

Public Property Get RegInitialKeyboardIndicatorsCurrentUser() As Integer
   Dim Ret As String
   Ret = LectureRegistre(HKY_IKIUSR, KEY_IKIUSR, VAL_IKIUSR, REGSTS_SZNOVALUE, REGSTS_SZNOKEY)
   Select Case Ret
      Case 0 To 7: RegInitialKeyboardIndicatorsCurrentUser = Val(Ret)
      Case REGSTS_SZNOVALUE: RegInitialKeyboardIndicatorsCurrentUser = REGSTS_INOVALUE
      Case REGSTS_SZNOKEY: RegInitialKeyboardIndicatorsCurrentUser = REGSTS_INOKEY
      Case Else: RegInitialKeyboardIndicatorsCurrentUser = REGSTS_IINVVAL
   End Select
End Property
Public Property Let RegInitialKeyboardIndicatorsCurrentUser(V As Integer)
   Call EcritureRegistre(HKY_IKIUSR, KEY_IKIUSR, VAL_IKIUSR, REG_SZ, V)
End Property

' ======== INH SHUTDOWN DANS LOGIN ========
Public Property Get RegInhShutdownDansLogin() As Integer
   Dim Ret As String
   Ret = LectureRegistre(HKY_ISDLOG, KEY_ISDLOG, VAL_ISDLOG, REGSTS_SZNOVALUE, REGSTS_SZNOKEY)
   Select Case Ret
      Case 0: RegInhShutdownDansLogin = 1
      Case 1: RegInhShutdownDansLogin = 0
      Case REGSTS_SZNOVALUE: RegInhShutdownDansLogin = REGSTS_INOVALUE
      Case REGSTS_SZNOKEY: RegInhShutdownDansLogin = REGSTS_INOKEY
      Case Else: RegInhShutdownDansLogin = REGSTS_IINVVAL
   End Select
End Property
Public Property Let RegInhShutdownDansLogin(Val As Integer)
   Call EcritureRegistre(HKY_MSGLOG, KEY_MSGLOG, VAL_MSLTXT, REG_SZ, 1 - Val)
End Property

' ======== CPL AFFICHAGE ========
Public Property Get RegInhCPLAffAccesCPL() As Integer
   Dim Ret As Long
   Ret = LectureRegistre(HKY_CPLAFF, KEY_CPLAFF, VAL_CPAGEN, 0, 0)
   Select Case Ret
      Case 0, 1: RegInhCPLAffAccesCPL = Ret
      Case Else: RegInhCPLAffAccesCPL = REGSTS_IINVVAL
   End Select
End Property
Public Property Let RegInhCPLAffAccesCPL(V As Integer)
   If V = 1 Then
     Call EcritureRegistre(HKY_CPLAFF, KEY_CPLAFF, VAL_CPAGEN, REG_DWORD, 1)
    Else
       Call EffaceRegistre(HKY_CPLAFF, KEY_CPLAFF, VAL_CPAGEN)
   End If
End Property

Public Property Get RegInhCPLAffOgtApp() As Integer
   Dim Ret As Long
   Ret = LectureRegistre(HKY_CPLAFF, KEY_CPLAFF, VAL_CPAAPP, 0, 0)
   Select Case Ret
      Case 0, 1: RegInhCPLAffOgtApp = Ret
      Case Else: RegInhCPLAffOgtApp = REGSTS_IINVVAL
   End Select
End Property
Public Property Let RegInhCPLAffOgtApp(Val As Integer)
   If Val = 1 Then
     Call EcritureRegistre(HKY_CPLAFF, KEY_CPLAFF, VAL_CPAAPP, REG_DWORD, 1)
    Else
       Call EffaceRegistre(HKY_CPLAFF, KEY_CPLAFF, VAL_CPAAPP)
   End If
End Property

Public Property Get RegInhCPLAffOgtAP() As Integer
   Dim Ret As Long
   Ret = LectureRegistre(HKY_CPLAFF, KEY_CPLAFF, VAL_CPABKG, 0, 0)
   Select Case Ret
      Case 0, 1: RegInhCPLAffOgtAP = Ret
      Case Else: RegInhCPLAffOgtAP = REGSTS_IINVVAL
   End Select
End Property
Public Property Let RegInhCPLAffOgtAP(Val As Integer)
   If Val = 1 Then
     Call EcritureRegistre(HKY_CPLAFF, KEY_CPLAFF, VAL_CPABKG, REG_DWORD, 1)
    Else
       Call EffaceRegistre(HKY_CPLAFF, KEY_CPLAFF, VAL_CPABKG)
   End If
End Property

Public Property Get RegInhCPLAffOgtEco() As Integer
   Dim Ret As Long
   Ret = LectureRegistre(HKY_CPLAFF, KEY_CPLAFF, VAL_CPASCR, 0, 0)
   Select Case Ret
      Case 0, 1: RegInhCPLAffOgtEco = Ret
      Case Else: RegInhCPLAffOgtEco = REGSTS_IINVVAL
   End Select
End Property
Public Property Let RegInhCPLAffOgtEco(Val As Integer)
   If Val = 1 Then
     Call EcritureRegistre(HKY_CPLAFF, KEY_CPLAFF, VAL_CPASCR, REG_DWORD, 1)
    Else
       Call EffaceRegistre(HKY_CPLAFF, KEY_CPLAFF, VAL_CPASCR)
   End If
End Property

Public Property Get RegInhCPLAffOgtCnf() As Integer
   Dim Ret As Long
   Ret = LectureRegistre(HKY_CPLAFF, KEY_CPLAFF, VAL_CPASET, 0, 0)
   Select Case Ret
      Case 0, 1: RegInhCPLAffOgtCnf = Ret
      Case Else: RegInhCPLAffOgtCnf = REGSTS_IINVVAL
   End Select
End Property
Public Property Let RegInhCPLAffOgtCnf(Val As Integer)
   If Val = 1 Then
     Call EcritureRegistre(HKY_CPLAFF, KEY_CPLAFF, VAL_CPASET, REG_DWORD, 1)
    Else
       Call EffaceRegistre(HKY_CPLAFF, KEY_CPLAFF, VAL_CPASET)
   End If
End Property

' ======== INH ICONE VOISINAGE RESEAU ========
Public Property Get RegInhIconeReseau() As Integer
   Dim Ret As Long
   Ret = LectureRegistre(HKY_ICONET, KEY_ICONET, VAL_ICONET, 0, REGSTS_INOKEY)
   Select Case Ret
      Case 0, 1, REGSTS_INOKEY: RegInhIconeReseau = Ret
      Case Else:                RegInhIconeReseau = REGSTS_IINVVAL
   End Select
End Property
Public Property Let RegInhIconeReseau(Val As Integer)
   If Val = 1 Then
     Call EcritureRegistre(HKY_ICONET, KEY_ICONET, VAL_ICONET, REG_DWORD, 1)
    Else
       Call EffaceRegistre(HKY_ICONET, KEY_ICONET, VAL_ICONET)
   End If
End Property

' ======== INH RESEAU GLOBAL ========
Public Property Get RegInhReseauGlobal() As Integer
   Dim Ret As Long
   Ret = LectureRegistre(HKY_ENTNET, KEY_ENTNET, VAL_ENTNET, 0, 0)
   Select Case Ret
      Case 0, 1, REGSTS_INOKEY: RegInhReseauGlobal = Ret
      Case Else:                RegInhReseauGlobal = REGSTS_IINVVAL
   End Select
End Property
Public Property Let RegInhReseauGlobal(Val As Integer)
   If Val = 1 Then
     Call EcritureRegistre(HKY_ENTNET, KEY_ENTNET, VAL_ENTNET, REG_DWORD, 1)
    Else
       Call EffaceRegistre(HKY_ENTNET, KEY_ENTNET, VAL_ENTNET)
   End If
End Property

' ======== INH CONTENU GROUPE TRAVAIL ========
Public Property Get RegInhContenuGroupesTravail() As Integer
   Dim Ret As Long
   Ret = LectureRegistre(HKY_NETWKC, KEY_NETWKC, VAL_NETWKC, 0, 0)
   Select Case Ret
      Case 0, 1, REGSTS_INOKEY: RegInhContenuGroupesTravail = Ret
      Case Else:                RegInhContenuGroupesTravail = REGSTS_IINVVAL
   End Select
End Property
Public Property Let RegInhContenuGroupesTravail(Val As Integer)
   If Val = 1 Then
     Call EcritureRegistre(HKY_NETWKC, KEY_NETWKC, VAL_NETWKC, REG_DWORD, 1)
    Else
       Call EffaceRegistre(HKY_NETWKC, KEY_NETWKC, VAL_NETWKC)
   End If
End Property

' ======== INH ICONE INTERNET EXPLORER ========
Public Property Get RegInhIconeIE() As Integer
   Dim Ret As Long
   Ret = LectureRegistre(HKY_ICOWEB, KEY_ICOWEB, VAL_ICOWEB, 0, REGSTS_INOKEY)
   Select Case Ret
      Case 0, 1, REGSTS_INOKEY: RegInhIconeIE = Ret
      Case Else:                RegInhIconeIE = REGSTS_IINVVAL
   End Select
End Property
Public Property Let RegInhIconeIE(Val As Integer)
   If Val = 1 Then
     Call EcritureRegistre(HKY_ICOWEB, KEY_ICOWEB, VAL_ICOWEB, REG_DWORD, 1)
    Else
       Call EffaceRegistre(HKY_ICOWEB, KEY_ICOWEB, VAL_ICOWEB)
   End If
End Property

' ======== INH CLIC DROIT SOURIS ========
Public Property Get RegInhClicDroit() As Integer
   Dim Ret As Long
   Ret = LectureRegistre(HKY_CLKDRT, KEY_CLKDRT, VAL_CLKDRT, 0, REGSTS_INOKEY)
   Select Case Ret
      Case 0, 1, REGSTS_INOKEY: RegInhClicDroit = Ret
      Case Else:                RegInhClicDroit = REGSTS_IINVVAL
   End Select
End Property
Public Property Let RegInhClicDroit(Val As Integer)
   If Val = 1 Then
     Call EcritureRegistre(HKY_CLKDRT, KEY_CLKDRT, VAL_CLKDRT, REG_DWORD, 1)
    Else
       Call EffaceRegistre(HKY_CLKDRT, KEY_CLKDRT, VAL_CLKDRT)
   End If
End Property

' ======== INH IMPRIMANTE AJOUT ========
Public Property Get RegInhImpAjout() As Integer
   Dim Ret As Long
   Ret = LectureRegistre(HKY_IMPAJT, KEY_IMPAJT, VAL_IMPAJT, 0, REGSTS_INOKEY)
   Select Case Ret
      Case 0, 1, REGSTS_INOKEY: RegInhImpAjout = Ret
      Case Else:                RegInhImpAjout = REGSTS_IINVVAL
   End Select
End Property
Public Property Let RegInhImpAjout(Val As Integer)
   If Val = 1 Then
     Call EcritureRegistre(HKY_IMPAJT, KEY_IMPAJT, VAL_IMPAJT, REG_DWORD, 1)
    Else
       Call EffaceRegistre(HKY_IMPAJT, KEY_IMPAJT, VAL_IMPAJT)
   End If
End Property

' ======== INH IMPRIMANTE SUPP ========
Public Property Get RegInhImpSupp() As Integer
   Dim Ret As Long
   Ret = LectureRegistre(HKY_IMPSUP, KEY_IMPSUP, VAL_IMPSUP, 0, REGSTS_INOKEY)
   Select Case Ret
      Case 0, 1, REGSTS_INOKEY: RegInhImpSupp = Ret
      Case Else:                RegInhImpSupp = REGSTS_IINVVAL
   End Select
End Property
Public Property Let RegInhImpSupp(Val As Integer)
   If Val = 1 Then
     Call EcritureRegistre(HKY_IMPSUP, KEY_IMPSUP, VAL_IMPSUP, REG_DWORD, 1)
    Else
       Call EffaceRegistre(HKY_IMPSUP, KEY_IMPSUP, VAL_IMPSUP)
   End If
End Property

' ======== INH BUREAU ========
Public Property Get RegInhDesktop() As Integer
   Dim Ret As Long
   Ret = LectureRegistre(HKY_INHDSK, KEY_INHDSK, VAL_INHDSK, 0, REGSTS_INOKEY)
   Select Case Ret
      Case 0, 1, REGSTS_INOKEY: RegInhDesktop = Ret
      Case Else:                RegInhDesktop = REGSTS_IINVVAL
   End Select
End Property
Public Property Let RegInhDesktop(Val As Integer)
   If Val = 1 Then
     Call EcritureRegistre(HKY_INHDSK, KEY_INHDSK, VAL_INHDSK, REG_DWORD, 1)
    Else
       Call EffaceRegistre(HKY_INHDSK, KEY_INHDSK, VAL_INHDSK)
   End If
End Property

' ======== INH ARRET WINDOWS ========
Public Property Get RegInhArretWindows() As Integer
   Dim Ret As Long
   Ret = LectureRegistre(HKY_STPWIN, KEY_STPWIN, VAL_STPWIN, 0, REGSTS_INOKEY)
   Select Case Ret
      Case 0, 1, REGSTS_INOKEY: RegInhArretWindows = Ret
      Case Else:                RegInhArretWindows = REGSTS_IINVVAL
   End Select
End Property
Public Property Let RegInhArretWindows(Val As Integer)
   If Val = 1 Then
     Call EcritureRegistre(HKY_STPWIN, KEY_STPWIN, VAL_STPWIN, REG_DWORD, 1)
    Else
       Call EffaceRegistre(HKY_STPWIN, KEY_STPWIN, VAL_STPWIN)
   End If
End Property

' ======== INH LOG OFF ========
Public Property Get RegInhLogOff() As Integer
   Dim Ret As Long
   Ret = LectureRegistre(HKY_LOGOFF, KEY_LOGOFF, VAL_LOGOFF, 0, REGSTS_INOKEY)
   Select Case Ret
      Case 0, 1, REGSTS_INOKEY: RegInhLogOff = Ret
      Case Else:                RegInhLogOff = REGSTS_IINVVAL
   End Select
End Property
Public Property Let RegInhLogOff(Val As Integer)
   If Val = 1 Then
     Call EcritureRegistre(HKY_LOGOFF, KEY_LOGOFF, VAL_LOGOFF, REG_DWORD, 1)
    Else
       Call EffaceRegistre(HKY_LOGOFF, KEY_LOGOFF, VAL_LOGOFF)
   End If
End Property

' ======== INH PROGRAMMES COMMUNS ========
Public Property Get RegInhPrgsCommuns() As Integer
   Dim Ret As Long
   Ret = LectureRegistre(HKY_PRGCMN, KEY_PRGCMN, VAL_PRGCMN, 0, REGSTS_INOKEY)
   Select Case Ret
      Case 0, 1, REGSTS_INOKEY: RegInhPrgsCommuns = Ret
      Case Else:                RegInhPrgsCommuns = REGSTS_IINVVAL
   End Select
End Property
Public Property Let RegInhPrgsCommuns(Val As Integer)
   If Val = 1 Then
     Call EcritureRegistre(HKY_PRGCMN, KEY_PRGCMN, VAL_PRGCMN, REG_DWORD, 1)
    Else
       Call EffaceRegistre(HKY_PRGCMN, KEY_PRGCMN, VAL_PRGCMN)
   End If
End Property

' ======== INH COMMANDE RUN ========
Public Property Get RegInhCmdRun() As Integer
   Dim Ret As Long
   Ret = LectureRegistre(HKY_CMDRUN, KEY_CMDRUN, VAL_CMDRUN, 0, REGSTS_INOKEY)
   Select Case Ret
      Case 0, 1, REGSTS_INOKEY: RegInhCmdRun = Ret
      Case Else:                RegInhCmdRun = REGSTS_IINVVAL
   End Select
End Property
Public Property Let RegInhCmdRun(Val As Integer)
   If Val = 1 Then
     Call EcritureRegistre(HKY_CMDRUN, KEY_CMDRUN, VAL_CMDRUN, REG_DWORD, 1)
    Else
       Call EffaceRegistre(HKY_CMDRUN, KEY_CMDRUN, VAL_CMDRUN)
   End If
End Property

' ======== INH COMMANDE FIND ========
Public Property Get RegInhCmdFind() As Integer
   Dim Ret As Long
   Ret = LectureRegistre(HKY_CMDFND, KEY_CMDFND, VAL_CMDFND, 0, REGSTS_INOKEY)
   Select Case Ret
      Case 0, 1, REGSTS_INOKEY: RegInhCmdFind = Ret
      Case Else:                RegInhCmdFind = REGSTS_IINVVAL
   End Select
End Property
Public Property Let RegInhCmdFind(Val As Integer)
   If Val = 1 Then
     Call EcritureRegistre(HKY_CMDFND, KEY_CMDFND, VAL_CMDFND, REG_DWORD, 1)
    Else
       Call EffaceRegistre(HKY_CMDFND, KEY_CMDFND, VAL_CMDFND)
   End If
End Property
' ======== INH CONFIG GENERALE ========
Public Property Get RegInhCnfGen() As Integer
   Dim Ret As Long
   Ret = LectureRegistre(HKY_CNFGEN, KEY_CNFGEN, VAL_CNFGEN, 0, REGSTS_INOKEY)
   Select Case Ret
      Case 0, 1, REGSTS_INOKEY: RegInhCnfGen = Ret
      Case Else:                RegInhCnfGen = REGSTS_IINVVAL
   End Select
End Property
Public Property Let RegInhCnfGen(Val As Integer)
   If Val = 1 Then
     Call EcritureRegistre(HKY_CNFGEN, KEY_CNFGEN, VAL_CNFGEN, REG_DWORD, 1)
    Else
       Call EffaceRegistre(HKY_CNFGEN, KEY_CNFGEN, VAL_CNFGEN)
   End If
End Property

' ======== INH CONFIG TASK BAR ========
Public Property Get RegInhCnfTaskBar() As Integer
   Dim Ret As Long
   Ret = LectureRegistre(HKY_CNFTKB, KEY_CNFTKB, VAL_CNFTKB, 0, REGSTS_INOKEY)
   Select Case Ret
      Case 0, 1, REGSTS_INOKEY: RegInhCnfTaskBar = Ret
      Case Else:                RegInhCnfTaskBar = REGSTS_IINVVAL
   End Select
End Property
Public Property Let RegInhCnfTaskBar(Val As Integer)
   If Val = 1 Then
     Call EcritureRegistre(HKY_CNFTKB, KEY_CNFTKB, VAL_CNFTKB, REG_DWORD, 1)
    Else
       Call EffaceRegistre(HKY_CNFTKB, KEY_CNFTKB, VAL_CNFTKB)
   End If
End Property

' ======== INH VEROUILLER STATION ========
Public Property Get RegInhLockStation() As Integer
   Dim Ret As Long
   Ret = LectureRegistre(HKY_LOCKST, KEY_LOCKST, VAL_LOCKST, 0, 0)
   Select Case Ret
      Case 0, 1: RegInhLockStation = Ret
      Case Else: RegInhLockStation = REGSTS_IINVVAL
   End Select
End Property
Public Property Let RegInhLockStation(Val As Integer)
   If Val = 1 Then
     Call EcritureRegistre(HKY_LOCKST, KEY_LOCKST, VAL_LOCKST, REG_DWORD, 1)
    Else
       Call EffaceRegistre(HKY_LOCKST, KEY_LOCKST, VAL_LOCKST)
   End If
End Property

' ======== INH TASK MANAGER ========
Public Property Get RegInhTaskManager() As Integer
   Dim Ret As Long
   Ret = LectureRegistre(HKY_TSKMGR, KEY_TSKMGR, VAL_TSKMGR, 0, 0)
   Select Case Ret
      Case 0, 1: RegInhTaskManager = Ret
      Case Else: RegInhTaskManager = REGSTS_IINVVAL
   End Select
End Property
Public Property Let RegInhTaskManager(Val As Integer)
   If Val = 1 Then
     Call EcritureRegistre(HKY_TSKMGR, KEY_TSKMGR, VAL_TSKMGR, REG_DWORD, 1)
    Else
       Call EffaceRegistre(HKY_TSKMGR, KEY_TSKMGR, VAL_TSKMGR)
   End If
End Property

' ======== INH CHANGE PASSWORD ========
Public Property Get RegInhChangePassword() As Integer
   Dim Ret As Long
   Ret = LectureRegistre(HKY_CHGPWD, KEY_CHGPWD, VAL_CHGPWD, 0, 0)
   Select Case Ret
      Case 0, 1: RegInhChangePassword = Ret
      Case Else: RegInhChangePassword = REGSTS_IINVVAL
   End Select
End Property
Public Property Let RegInhChangePassword(Val As Integer)
   If Val = 1 Then
     Call EcritureRegistre(HKY_CHGPWD, KEY_CHGPWD, VAL_CHGPWD, REG_DWORD, 1)
    Else
       Call EffaceRegistre(HKY_CHGPWD, KEY_CHGPWD, VAL_CHGPWD)
   End If
End Property

' ======== CACHE LECTEURS ========
Public Property Get RegCacheLecteurs() As String
   Dim Ret As Long
   Dim i   As Integer
   Dim Config As String
   Ret = LectureRegistre(HKY_HIDDRV, KEY_HIDDRV, VAL_HIDDRV, 0, REGSTS_INOKEY)
   Select Case Ret
      Case REGSTS_INOKEY: RegCacheLecteurs = REGSTS_SZNOKEY
      Case Else
          Config = ""
          For i = 0 To 25
             Config = Config & IIf(Ret And (2 ^ i), "1", "0")
          Next
          RegCacheLecteurs = Config
   End Select
End Property
Public Property Let RegCacheLecteurs(Config As String)
   Dim i As Integer
   Dim V As Long
   If (Val(Config) <> 0) Then
     For i = 0 To 25
        V = V + IIf(Mid$(Config, i + 1, 1) = "1", (2 ^ i), 0)
     Next
     Call EcritureRegistre(HKY_HIDDRV, KEY_HIDDRV, VAL_HIDDRV, REG_DWORD, V)
    Else
       Call EffaceRegistre(HKY_HIDDRV, KEY_HIDDRV, VAL_HIDDRV)
   End If
End Property


' ======== TAILLE ICONES BUREAU ========
Public Property Get RegTailleIconesBureau() As Integer
   Dim Ret As String
   Ret = LectureRegistre(HKY_SIZICD, KEY_SIZICD, VAL_SIZICD, REGSTS_SZNOVALUE, REGSTS_SZNOKEY)
   Select Case Trim$(Ret)
      Case REGSTS_SZNOVALUE: RegTailleIconesBureau = REGSTS_INOVALUE
      Case REGSTS_SZNOKEY: RegTailleIconesBureau = REGSTS_INOKEY
      Case Else:
          If Val(Ret) > 0 And Val(Ret) < 200 Then
            RegTailleIconesBureau = Val(Ret)
           Else
                  RegTailleIconesBureau = REGSTS_IINVVAL
          End If
   End Select
End Property
Public Property Let RegTailleIconesBureau(V As Integer)
   Call EcritureRegistre(HKY_SIZICD, KEY_SIZICD, VAL_SIZICD, REG_SZ, V)
End Property

Public Property Get RegTailleIconesMenuStart() As Integer
   Dim Ret As String
   Ret = LectureRegistre(HKY_SIZICS, KEY_SIZICS, VAL_SIZICS, REGSTS_SZNOVALUE, REGSTS_SZNOKEY)
   Select Case Trim$(Ret)
      Case REGSTS_SZNOVALUE: RegTailleIconesMenuStart = REGSTS_INOVALUE
      Case REGSTS_SZNOKEY: RegTailleIconesMenuStart = REGSTS_INOKEY
      Case Else:
          If Val(Ret) > 0 And Val(Ret) < 200 Then
            RegTailleIconesMenuStart = Val(Ret)
           Else
                  RegTailleIconesMenuStart = REGSTS_IINVVAL
          End If
   End Select
End Property
Public Property Let RegTailleIconesMenuStart(V As Integer)
   Call EcritureRegistre(HKY_SIZICS, KEY_SIZICS, VAL_SIZICS, REG_SZ, V)
End Property

' ======== DELAY ========
Public Property Get RegDelaiMenus() As Integer
   Dim Ret As String
   Ret = LectureRegistre(HKY_TMRMNU, KEY_TMRMNU, VAL_TMRMNU, REGSTS_SZNOVALUE, REGSTS_SZNOKEY)
   Select Case Trim$(Ret)
      Case REGSTS_SZNOVALUE: RegDelaiMenus = REGSTS_INOVALUE
      Case REGSTS_SZNOKEY: RegDelaiMenus = REGSTS_INOKEY
      Case Else:
          If Val(Ret) > 0 And Val(Ret) < 20000 Then
            RegDelaiMenus = Val(Ret)
           Else
                  RegDelaiMenus = REGSTS_IINVVAL
          End If
   End Select
End Property
Public Property Let RegDelaiMenus(V As Integer)
   Call EcritureRegistre(HKY_TMRMNU, KEY_TMRMNU, VAL_TMRMNU, REG_SZ, V)
End Property


Public Property Get SerialNumber() As String
   SerialNumber = LectureRegistre(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", _
                        "ProductId", -99, -98)
End Property



' ======== INH TOUCHE WINDOWS ========
'Public Property Get RegToucheWin() As Integer
'   RegToucheWin = LectureRegistre(HKY_TCHWIN, KEY_TCHWIN, VAL_TCHWIN, REGSTS_INOVALUE, REGSTS_INOKEY)
'End Property

' ======== INH ALT TAB ========
'Public Property Get RegAltTab() As Integer
'   RegAltTab = LectureRegistre(HKY_ALTTAB, KEY_ALTTAB, VAL_ALTTAB, REGSTS_INOVALUE, REGSTS_INOKEY)
'End Property

' ======== INH TOUCHE SHIFT AU DEMARRAGE ========
'Public Property Get RegToucheShift() As String
'   RegToucheShift = LectureRegistre(HKY_TSHIFT, KEY_TSHIFT, VAL_TSHIFT, REGSTS_SZNOVALUE, REGSTS_SZNOKEY)
'End Property

' ======== INH BOUTON DROIT ========
'Public Property Get RegBoutonDroit() As Integer
'   RegBoutonDroit = LectureRegistre(HKY_BTNDRT, KEY_BTNDRT, VAL_BTNDRT, REGSTS_INOVALUE, REGSTS_INOKEY)
'End Property

' ======== VEROUILLAGE PAVE NUMERIQUE AU DEMARRAGE ========
'Public Property Get RegPaveNum() As Integer
'   RegPaveNum = LectureRegistre(HKY_PAVNUM, KEY_PAVNUM, VAL_PAVNUM, REGSTS_SZNOVALUE, REGSTS_SZNOKEY)
'End Property

' ======== LANCEMENT PROGRAMME PAR CTRL-ESC ========
'Public Property Get RegCtrlEsc() As String
'   RegCtrlEsc = LectureRegistre(HKY_CTLESC, KEY_CTLESC, VAL_CTLESC, REGSTS_SZNOVALUE, REGSTS_SZNOKEY)
'End Property
