VERSION 5.00
Begin VB.Form Frm_AideReg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aide"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6900
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdCloseHelp 
      Cancel          =   -1  'True
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblPriseEnCompte 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   195
      Left            =   5400
      TabIndex        =   9
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label5 
      Caption         =   "Prise en compte :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   8
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "du système"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblModifications 
      AutoSize        =   -1  'True
      Caption         =   "Label4"
      Height          =   195
      Left            =   1440
      TabIndex        =   5
      Top             =   1320
      Width           =   480
   End
   Begin VB.Label Label3 
      Caption         =   "Modifications"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   195
      Left            =   1440
      TabIndex        =   3
      Top             =   600
      Width           =   480
   End
   Begin VB.Label lblFonction 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   195
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label2 
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Fonction"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Frm_AideReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCloseHelp_Click()
   Unload Me
End Sub

Function HKeyName(HKey As hKeyNames) As String
   Select Case HKey
      Case HKEY_CLASSES_ROOT:      HKeyName = "\\HKey_Classes_Root\"
      Case HKEY_CURRENT_USER:      HKeyName = "\\HKey_Current_User\"
      Case HKEY_LOCAL_MACHINE:     HKeyName = "\\HKey_Local_Machine\"
      Case HKEY_USERS:             HKeyName = "\\HKey_Users\"
      Case Else:                   HKeyName = "\\<CLEF INCONNUE!!>\"
   End Select

End Function

Private Sub Form_Load()
   Dim Txt1, Txt2, Txt3, Txt4 As String
   Dim l, L1, L2, L3, L4 As Long
   
   Me.Icon = Frm_Principal.imgIcone.Picture
   
   Select Case RefAide
      Case IDX_ALOGIN
          Txt1 = "Auto Login"
          Txt2 = "Permet de se logger automatiquement au démarrage." & vbCrLf & _
                 "Le système ne s'arrête plus sur la popup de logon."
          Txt3 = HKeyName(HKY_ALOGIN) & KEY_ALOGIN & vbCrLf & _
                 "AutoAdminLogon = 1  (0=Inactif)" & vbCrLf & _
                 "DefaultUserName = xxxx" & vbCrLf & _
                 "DefaultPassword = xxxx  (ATTENTION : le mot de passe est affiché en clair !!!)" & vbCrLf & _
                 "DefaultDomainName = xxxx"
          Txt4 = "Reconnexion"
      Case IDX_MSGLOG
          Txt1 = "Popup avertissement"
          Txt2 = "Cette popup s'affiche au logon, après avoir appuyé sur Ctrl-Alt-Supp." & vbCrLf & _
                 "Une fois validée, la popup de logon apparait." & vbCrLf & _
                 "Il est possible de paramétrer le titre et le contenu de cette popup."
          Txt3 = HKeyName(HKY_MSGLOG) & KEY_MSGLOG & vbCrLf & _
                 "LegalNoticeCaption pour le titre (string)" & vbCrLf & _
                 "LegalNoticeText pour le contenu (string)."
          Txt4 = "Reconnexion"
      Case IDX_POPLOG
          Txt1 = "Popup Login"
          Txt2 = "Ce texte est affiché dans la popup de logon demandant le nom d'utilisateur et son mot de passe." & vbCrLf & _
                 "Utile pour indiquer à l'opérateur comment se logger (nom et mot de passe)" & vbCrLf & _
                 "sans utiliser l'autologon."
          Txt3 = HKeyName(HKY_POPLOG) & KEY_POPLOG & vbCrLf & _
                 VAL_POPLOG & "(string)"
          Txt4 = "Reconnexion"
      Case IDX_KEYB11
          Txt1 = "Configuration clavier"
          Txt2 = "Configuration initiale des vérouillages clavier avant et après le login"
          Txt3 = "Avant : " & HKeyName(HKY_IKIDFT) & KEY_IKIDFT & "\" & VAL_IKIDFT & " (string)" & vbCrLf & _
                 "Après : " & HKeyName(HKY_IKIUSR) & KEY_IKIUSR & "\" & VAL_IKIUSR & " (string)" & vbCrLf & _
                 "Codage binaire : 1=Caps  2=Num  4=Scroll"
          Txt4 = "Reconnexion/Redémarrage"
      Case IDX_ISDLOG
          Txt1 = "Inhibition arrêt depuis login"
          Txt2 = "Supprime la possibilité d'arrêter le système puis la popup de login."
          Txt3 = HKeyName(HKY_ISDLOG) & KEY_ISDLOG & vbCrLf & _
                 VAL_ISDLOG & " (string)" & "  -  0 = Inhibition active"
          Txt4 = "Redémarrage"
      Case IDX_CPAGEN
          Txt1 = "Panneau d'Affichage"
          Txt2 = "Ces options permettent de masquer certains onglets du panneau de configuration" & vbCrLf & _
                 "de l'affichage, ou de bloquer l'accès complet au panneau."
          Txt3 = HKeyName(HKY_CPLAFF) & KEY_CPLAFF & vbCrLf & _
                 "NoDispCPL                           Acces complet au panneau" & vbCrLf & _
                 "NoDispAppearancePage      Onglet Apparance" & vbCrLf & _
                 "NoDispBackGroundPage     Onglet Arrière-Plan" & vbCrLf & _
                 "NoDispScrSavPage             Onglet Economiseur d'écran" & vbCrLf & _
                 "NoDispSettingsPage             Onglet Configuration"
          Txt4 = "Immédiat"
      Case IDX_ICONET
          Txt1 = "Icone Voisinage réseau"
          Txt2 = "L'icone 'Voisinage réseau' peut être retirée du bureau, de l'exploreur et" & vbCrLf & _
                 "des popups Ouvrir/Sauver/Sauver sous grâce à cette option."
          Txt3 = HKeyName(HKY_CPLAFF) & KEY_CPLAFF & vbCrLf & _
                 VAL_ICONET & " (dword)"
          Txt4 = "Reconnexion"
      
      Case IDX_INHDSK
          Txt1 = "Bureau"
          Txt2 = "Désactive l'affichage des icones du bureau et le menu contextuel (bouton droit souris)"
          Txt3 = HKeyName(HKY_INHDSK) & KEY_INHDSK & vbCrLf & _
                 VAL_INHDSK & " (dword)"
          Txt4 = "Reconnexion"
      
      Case IDX_ENTNET
          Txt1 = "Réseau : Réseau global"
          Txt2 = "Supprime l'option 'Réseau global' dans le voisinage réseau." & vbCrLf & _
                 "Evite de voir tous les domaines disponibles."
          Txt3 = HKeyName(HKY_ENTNET) & KEY_ENTNET & vbCrLf & _
                 VAL_ENTNET & " (dword)"
          Txt4 = "Reconnexion"
      
      Case IDX_NETWKC
          Txt1 = "Contenu Groupe de travail"
          Txt2 = "Supprime l'affichage des autres PCs du même domaine dans le voisinage réseau."
          Txt3 = HKeyName(HKY_NETWKC) & KEY_NETWKC & vbCrLf & _
                 VAL_NETWKC & " (dword)"
          Txt4 = "Reconnexion"

      Case IDX_ICOWEB
          Txt1 = "Icone Internet Explorer"
          Txt2 = "Cache l'icone Internet Explorer du bureau."
          Txt3 = HKeyName(HKY_ICOWEB) & KEY_ICOWEB & vbCrLf & _
                 VAL_ICOWEB & " (dword)"
          Txt4 = "Reconnexion"

      Case IDX_CLKDRT
          Txt1 = "Click Droit"
          Txt2 = "Inhibition des menus contextuels obtenus par le bouton droit" & vbCrLf & _
                 "de la souris sur le bureau et dans l'exploreur."
          Txt3 = HKeyName(HKY_CLKDRT) & KEY_CLKDRT & vbCrLf & _
                 VAL_CLKDRT & " (dword)"
          Txt4 = "Reconnexion"

      Case IDX_IMPAJT
          Txt1 = "Imprimante"
          Txt2 = "Inhibition de la possibilité d'ajouter ou de supprimer des imprimantes"
          Txt3 = HKeyName(HKY_IMPAJT) & KEY_IMPAJT & vbCrLf & _
                 "Ajout            : " & VAL_IMPAJT & " (dword)" & vbCrLf & _
                 "Suppression : " & VAL_IMPSUP & " (dword)"
          Txt4 = "On=Immédiat Off=Reconnexion"

      Case IDX_STPWIN
          Txt1 = "Inhib. éléments bouton START"
          Txt2 = "Inhibition de commandes accessibles depuis le bouton START: " & _
                 vbCrLf & "Arrêt de windows - Logoff / Commandes Exécuter et Trouver / " & _
                 "Panneau de" & vbCrLf & "configuration et config. barre des tâches" & _
                 " / Section programmes communs"
          Txt3 = HKeyName(HKY_STPWIN) & KEY_STPWIN & vbCrLf & _
                 "Arrêt de Windows : " & VAL_STPWIN & " (dword)            " & _
                 "Commande Exécuter : " & VAL_CMDRUN & " (dword)" & vbCrLf & _
                 "Commande Trouver : " & VAL_CMDFND & " (dword)            " & _
                 "Config Principale : " & VAL_CNFGEN & " (dword)" & vbCrLf & _
                 "Config barre tâches : " & VAL_CNFTKB & " (dword)         " & _
                 "Logoff : " & VAL_LOGOFF & " (dword)" & vbCrLf & _
                 "Prg communs : " & VAL_PRGCMN & " (dword)" & vbCrLf
          Txt4 = "On=Immédiat Off=Reconnexion"

      Case IDX_LOCKST
          Txt1 = "Inhib. fonctions 'Sécurité NT'"
          Txt2 = "Inhibition de commandes de la sécurité NT (Ctrl-Alt-Supp) : " & _
                 vbCrLf & "Vérouillage de la station / " & _
                          "Appel du gestionnaire des tâches / " & _
                          "Changement du mot de passe"
          Txt3 = HKeyName(HKY_LOCKST) & KEY_LOCKST & vbCrLf & _
                 "Vérouiller Station : " & VAL_LOCKST & " (dword)" & vbCrLf & _
                 "Gestion. tâches    : " & VAL_TSKMGR & " (dword)" & vbCrLf & _
                 "Mot de passe       : " & VAL_CHGPWD & " (dword)"
          Txt4 = "Immédiat"

      Case IDX_HIDDRV
          Txt1 = "Cacher lecteurs"
          Txt2 = "Permet de cacher des lecteurs dans l'exploreur et dans les" & vbCrLf & _
                 "popups Ouvrir / Sauver / Sauver sous" & vbCrLf & _
                 "Attention : le disque sur lequel se trouve NT est parfois visible !"
          Txt3 = HKeyName(HKY_HIDDRV) & KEY_HIDDRV & vbCrLf & _
                 VAL_HIDDRV & " (dword)" & vbCrLf & _
                 "Chaque bit représente un lecteur A=bit 0  B=bit 1..."
          Txt4 = "Reconnexion"

      Case IDX_SIZICD
          Txt1 = "Taille des icônes du Bureau"
          Txt2 = "Spécifie la taille en pixels des icônes dans les fenêtres de navigation" & vbCrLf & _
                 "(Mon ordinateur, réseau...). Bien que cela fonctionne avec toutes" & vbCrLf & _
                 "les tailles, il est recommendé d'utiliser des multiples de 16"
          Txt3 = HKeyName(HKY_SIZICD) & KEY_SIZICD & "\" & vbCrLf & VAL_SIZICD & " (string)"
          Txt4 = "Immédiat"

      Case IDX_SIZICS
          Txt1 = "Taille des icônes du menu Start"
          Txt2 = "Spécifie la taille en pixels des icônes du menu start (Programmes," & vbCrLf & _
                 "Documents, Paramètres...). Bien que cela fonctionne avec toutes" & vbCrLf & _
                 "les tailles, il est recommendé d'utiliser des multiples de 16"
          Txt3 = HKeyName(HKY_SIZICS) & KEY_SIZICS & "\" & vbCrLf & VAL_SIZICS & " (string)"
          Txt4 = "Immédiat"

      Case IDX_TMRMNU
          Txt1 = "Délai d'affichage des menus"
          Txt2 = "Dans les menus déroulants, délai en ms avant d'afficher un sous menu."
          Txt3 = HKeyName(HKY_TMRMNU) & KEY_TMRMNU & "\" & VAL_TMRMNU & " (string)"
          Txt4 = "Reconnexion"


      Case Else
          Txt1 = "ERR AIDE"
          Txt2 = "FICHE INCONNUE CODE " & RefAide
          Txt3 = ""
   End Select
   
   lblFonction.Caption = Txt1
   lblDescription.Caption = Txt2
   lblModifications.Caption = Txt3
   lblPriseEnCompte = Txt4
   
   L1 = lblFonction.Left + lblFonction.Width
   L2 = lblDescription.Left + lblDescription.Width
   L3 = lblModifications.Left + lblModifications.Width
   L4 = lblPriseEnCompte.Left + lblPriseEnCompte.Width
   l = L1
   If L2 > l Then l = L2
   If L3 > l Then l = L3
   If L4 > l Then l = L4
   Me.Width = l + 300
   
   cmdCloseHelp.Left = (Me.Width / 2) - (cmdCloseHelp.Width / 2)
   
End Sub

