VERSION 5.00
Begin VB.Form Frm_Aide 
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
      TabIndex        =   2
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "xxxxxxxxxx"
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   720
      UseMnemonic     =   0   'False
      Width           =   750
   End
   Begin VB.Label lblTitre 
      AutoSize        =   -1  'True
      Caption         =   "xxxxxxxxxx"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1440
      TabIndex        =   0
      Top             =   180
      Width           =   1065
   End
End
Attribute VB_Name = "Frm_Aide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCloseHelp_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Dim Txt1, Txt2 As String
   
   Me.Icon = Frm_Principal.imgIcone.Picture
   
   Select Case RefAide
      Case IDX_ATMCLK
          Txt1 = "Heure Atomique"
          Txt2 = "L'heure atomique est recherch�e sur le site du " & vbCrLf & _
                 "National Institute of Standards & Technology (NIST)." & vbCrLf & _
                 "www.bldrdoc.gov/doc-tour/atomic_clock.html" & vbCrLf & _
                 "Adresse Donn�es : http://132.163.135.130:14" & vbCrLf & vbCrLf & _
                 "La connexion internet est effectu�e d'apr�s les param�tres " & _
                 "par d�faut de Windows" & vbCrLf & vbCrLf & _
                 "L'affichage se compose de la date julienne, la date et l'heure."
      Case IDX_SHTLOC
          Txt1 = "Shutdown Local"
          Txt2 = "Arr�te l'ordinateur local." & vbCrLf & _
                 "L'option 'Force' demande � Windows de tuer tous les programmes et t�ches" & vbCrLf & _
                 "plutot que de les fermer par la proc�dure normale." & vbCrLf & _
                 "Toutes les donn�es non enregistr�es seront perdues."
      Case IDX_SHTDST
          Txt1 = "Shutdown distant"
          Txt2 = "Idem shutdown local mais pour un PC distant." & vbCrLf & _
                 "Il faut �tre d�j� connect� avec le droit d'arr�t (administrateur ou utilisateur avec pouvoir)." & vbCrLf & _
                 "Fournir le nom du PC ou son IP et un �ventuel message." & vbCrLf & _
                 "Pour tester en local, utiliser l'IP et non le nom." & vbCrLf & _
                 "Sous WK2, il faut �tre connect� physiquement m�me pour un test local."
      
      Case IDX_CONNEC
          Txt1 = "Connexion PC distant"
          Txt2 = "Permet de se connecter � un autre PC pour modifier sa base des registres � distance." & vbCrLf & _
                 "Il faut connaitre un compte administrateur." & vbCrLf & _
                 "Si la connexion est r�ussie, certaines infos du syst�me distant sont affich�es." & vbCrLf & _
                 "Equivalent � la commande 'Registre/Connexion au registre r�seau' de Regedit."

      Case Else
          Txt1 = "ERR AIDE"
          Txt2 = "FICHE INCONNUE CODE " & RefAide
   End Select
   
   lblTitre.Caption = Txt1
   lblDescription.Caption = Txt2
   
'   L1 = lblTitre.Left + lblFonction.Width
'   L2 = lblDescription.Left + lblDescription.Width
'   l = L1
'   If L2 > l Then l = L2
'   Me.Width = l + 300
   
   lblTitre.Left = (Me.Width / 2) - (lblTitre.Width / 2)
   cmdCloseHelp.Left = (Me.Width / 2) - (cmdCloseHelp.Width / 2)
   
End Sub

