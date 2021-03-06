VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_IconeTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'======================================================================
'======================  ICONE SYSTEM TRAY ============================
'======================================================================

Private Type TypeIconeTray
   cbSize As Long           'Taille de l'ic�ne (en octets)
   hWnd As Long             'Handle de la fen�tre charg�e de reception messages
   uID As Long              'Identificateur de l'ic�ne
   uFlags As Long
   uCallbackMessage As Long 'Messages � renvoyer
   hIcon As Long            'Handle de l'ic�ne
   szTip As String * 64     'Texte � mettre dans la bulle d'aide
End Type
Dim IconeT As TypeIconeTray

'Constantes n�cessaires
Const AJOUT = &H0
Const MODIF = &H1
Const SUPPRIME = &H2
Const MOUSEMOVE = &H200
Const MESSAGE = &H1
Const ICONE = &H2
Const TIP = &H4

Enum ListeBtSouris
    DOUBLE_CLICK_GAUCHE = &H203
    BOUTON_GAUCHE_POUSSE = &H201
    BOUTON_GAUCHE_LEVE = &H202
    DOUBLE_CLICK_DROIT = &H206
    BOUTON_DROIT_POUSSE = &H204
    BOUTON_DROIT_LEVE = &H205
End Enum

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As TypeIconeTray) As Boolean

'===========================================================================
'===========================================================================

Public Property Let AppHandle(hApp As Long)
   IconeT.hWnd = hApp
End Property

Public Property Get AppHandle() As Long
   AppHandle = IconeT.hWnd
End Property

Public Property Let IconeHandle(hIcone As Long)
   IconeT.hIcon = hIcone
End Property

Public Property Get IconeHandle() As Long
   IconeHandle = IconeT.hIcon
End Property

Public Property Let ToolTips(Msg As String)
   IconeT.szTip = Msg & Chr$(0)
End Property

Public Property Get ToolTips() As String
   ToolTips = IconeT.szTip
End Property

Public Sub AfficheIcone()
   IconeT.cbSize = Len(IconeT) 'Taille de l'ic�ne en octet
   IconeT.uID = 1&             'Identificateur de l'ic�ne
   IconeT.uFlags = ICONE Or TIP Or MESSAGE
   IconeT.uCallbackMessage = MOUSEMOVE     'Renvoyer les messages concernant l'action de la souris
   Shell_NotifyIcon AJOUT, IconeT
End Sub

Public Sub SupprimeIcone()
   IconeT.cbSize = Len(IconeT)
'   IconeT.hWnd = Frm_Principal.hWnd
   IconeT.uID = 1&
   Shell_NotifyIcon SUPPRIME, IconeT
End Sub

