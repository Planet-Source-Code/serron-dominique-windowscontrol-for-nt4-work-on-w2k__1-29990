Attribute VB_Name = "Mod_Général"
Option Explicit

Global COL_ROUGE As Long     ' Constantes définies dans Init
Global COL_ORANGE As Long

Public Const FLG_SZ_AEFFACER = "<A EFFACER>"  ' Flag pour valeurs à effacer

Public Const INETADD_HEUREATOMIQUE = "http://132.163.135.130:14"

Enum EnumIndexBoutons   ' Index des groupes de boutons d'aide et de sauvegarde
    IDX_ALOGIN = 100      ' AutoLogin
    IDX_MSGLOG = 200      ' Message avant login
    IDX_POPLOG = 210      ' Message de la popup de login
    IDX_KEYB11 = 250      ' Indicateurs claviers
    IDX_KEYB12 = 251
    IDX_KEYB13 = 252
    IDX_KEYB21 = 260
    IDX_KEYB22 = 261
    IDX_KEYB23 = 262
    IDX_ISDLOG = 280      ' Inhibition arret systeme dans popup login
       
    IDX_CPAGEN = 300    ' CPL Affichage : acces complet
    IDX_CPAAPP = 301    '                 Apparance
    IDX_CPABKG = 302    '                 BackGround
    IDX_CPASCR = 303    '                 Screen Saver
    IDX_CPACNF = 304    '                 Settings
    IDX_INHDSK = 310    ' Inh Bureau
    IDX_ICONET = 320    ' Inh Icone Voisinage réseau
    IDX_ENTNET = 321    ' Inh Entire network
    IDX_NETWKC = 322    ' Inh Workgroup content
    IDX_ICOWEB = 330    ' Inh Icone Internet Explorer
    IDX_CLKDRT = 340    ' Inh Clic Droit
    IDX_IMPAJT = 350    ' Inh Imprimante Ajout
    IDX_IMPSUP = 351    ' Inh Imprimante Supprimer
    IDX_SIZICD = 360    ' Taille Icones Bureau
    IDX_SIZICS = 361    ' Taille Icones Menu Start
    IDX_TMRMNU = 380    ' Delai Menus
    
    IDX_STPWIN = 400    ' Inh Arrêt windows
    IDX_CMDRUN = 401    ' Inh Commande RUN menu Start
    IDX_CMDFND = 402    ' Inh Commande FIND menu Start
    IDX_CNFGEN = 403    ' Inh Configuration principale
    IDX_CNFTKB = 404    ' Inh Configuration TaskBar
    IDX_LOGOFF = 405    ' Inh LogOff
    IDX_PRGCMN = 406    ' Inh Programmes communs
    
    IDX_LOCKST = 420    ' Inh Vérouiller station
    IDX_TSKMGR = 421    ' Inh Task Manager
    IDX_CHGPWD = 422    ' Inh Changer Mot de passe
    
    IDX_HIDDRV = 500    ' Cacher lecteurs
    
    IDX_CONNEC = 600    ' Connexion
    
    IDX_ATMCLK = 10601  ' Heure atomique
    
    IDX_SHTLOC = 10001  ' ShutDown local
    IDX_SHTDST = 10002  ' ShutDown distant
    
End Enum

Const HKY_PARAMS = HKEY_CURRENT_USER
Const KEY_PARAMS = "SOFTWARE\WindowsControl"
Const VAL_NMDEST = "NetMsg_Dest"
Const VAL_NMTDST = "NetMsg_TypeDestination"
Const VAL_NMSIGN = "NetMsg_Signature"
Const VAL_NMSGAC = "NetMsg_SignatureActive"

Type TypeParamPrg
    RepProgActuel As String
    RepProgAll As String
    RepStartActuel As String
    RepStartAll As String
    NomRaccourci As String
    RaccourciActuel As String
    RaccourciAll As String
End Type
Global ParamPrg As TypeParamPrg

Type TypeParamNetMsgDest
   DNS As String
   Comment As String
End Type

Type TypeParamNetMsg
   Dest() As TypeParamNetMsgDest
   TypeDest As Integer
   Signature As String
   SignatureActive As Integer
End Type
Global ParamNetMsg As TypeParamNetMsg

Public SysInfos As New cls_SysInfos
Public Elmts As New cls_ElementsBureau
Public Tray As New cls_IconeTray
Public Net As New cls_Reseau
Public Inet As New cls_Internet

Public Declare Function Sound Lib "kernel32" Alias "Beep" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Global RefAide As Integer       ' Ref de l'aide à afficher dans form 'Aide'

Sub Main()
   '===== INIT =====
   COL_ROUGE = RGB(255, 0, 0)      ' Variables utilisées comme constantes
   COL_ORANGE = RGB(255, 127, 0)
   
   Call DetectionCheminsPrg
'   Call GetSystemInformation
'   Call DetectionHandleElements
   
   Frm_Principal.Show
   
   ' Affiche l'icone dans la barre des taches
   Tray.AppHandle = Frm_Principal.hWnd
   Tray.IconeHandle = Frm_Principal.imgIcone.Picture
   Tray.ToolTips = App.Title
   Tray.AfficheIcone
   
End Sub

Sub QuitterProg()
   ' Supprime l'icone
   Tray.AppHandle = Frm_Principal.hWnd   ' Maj du Handle
   Tray.SupprimeIcone
   Set SysInfos = Nothing  ' Libère mémoires des classes
   Set Elmts = Nothing
   Set Tray = Nothing
   Set Net = Nothing
   
End Sub

'============================ FONCTIONS COMMUNES ============================

Function SuppSZ(Txt As String) As String
   Dim Pos As Integer
   
   Pos = InStr(Txt, Chr$(0))
   If (Pos > 0) Then
     SuppSZ = Left$(Txt, Pos - 1)
     Else
          SuppSZ = Trim$(Txt)
   End If
End Function

Function ConvSecToDHMS(Sec_in As Long) As String
   Static j As Long
   Static H As Long
   Static m As Long
   Static S As Long
   
   j = Int(Sec_in / 86400)
   Sec_in = Sec_in - (j * 86400)
   
   H = Int(Sec_in / 3600)
   Sec_in = Sec_in - (H * 3600)
   
   m = Int(Sec_in / 60)
   S = Sec_in - (m * 60)
   
   ConvSecToDHMS = j & " jours, " & Right$("00" & H, 2) & ":" & _
                                    Right$("00" & m, 2) & ":" & _
                                    Right$("00" & S, 2)
End Function

Public Sub Pause(ByVal seconds As Single)
   Call Sleep(Int(seconds * 1000#))
End Sub

'=====================================================================
'=========================== Paramètres ==============================
'=====================================================================

Sub LectureParamNetMsg()
   Dim i As Integer
   Dim Ret As String
   
   ReDim ParamNetMsg.Dest(0)
   
   i = 1
   Do
     Ret = LectureRegistre(HKY_PARAMS, KEY_PARAMS, VAL_NMDEST & Trim$(Val(i)), "", "")
     If (Ret <> "") Then
       ReDim Preserve ParamNetMsg.Dest(UBound(ParamNetMsg.Dest) + 1)
       ParamNetMsg.Dest(UBound(ParamNetMsg.Dest)).DNS = Trim$(Left$(Ret, 30))
       ParamNetMsg.Dest(UBound(ParamNetMsg.Dest)).Comment = Mid$(Ret, 31)
     End If
     i = i + 1
   Loop While Ret <> ""
   
   ParamNetMsg.Signature = LectureRegistre(HKY_PARAMS, KEY_PARAMS, VAL_NMSIGN, "", "")
   ParamNetMsg.TypeDest = LectureRegistre(HKY_PARAMS, KEY_PARAMS, VAL_NMTDST, 0, 0)
   ParamNetMsg.SignatureActive = LectureRegistre(HKY_PARAMS, KEY_PARAMS, VAL_NMSGAC, 0, 0)
End Sub

Sub EcritureParamNetMsg()
   Dim i As Integer
   Dim Ret As String
   
   i = 0
   Do
     i = i + 1
   Loop While (EffaceRegistre(HKY_PARAMS, KEY_PARAMS, VAL_NMDEST & Trim$(Val(i))) = True)
   
   For i = 1 To UBound(ParamNetMsg.Dest)
         Call EcritureRegistre(HKY_PARAMS, KEY_PARAMS, VAL_NMDEST & Trim$(Val(i)), REG_SZ, _
                               Left$(ParamNetMsg.Dest(i).DNS & String$(50, 32), 30) & _
                               ParamNetMsg.Dest(i).Comment)
   Next
   
   Call EcritureRegistre(HKY_PARAMS, KEY_PARAMS, VAL_NMSIGN, REG_SZ, ParamNetMsg.Signature)
   Call EcritureRegistre(HKY_PARAMS, KEY_PARAMS, VAL_NMTDST, REG_DWORD, ParamNetMsg.TypeDest)
   Call EcritureRegistre(HKY_PARAMS, KEY_PARAMS, VAL_NMSGAC, REG_DWORD, ParamNetMsg.SignatureActive)
End Sub


