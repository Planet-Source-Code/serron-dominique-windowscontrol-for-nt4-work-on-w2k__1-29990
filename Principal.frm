VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Frm_Principal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10800
   Icon            =   "Principal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   10800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmRegistres 
      Height          =   2655
      Index           =   7
      Left            =   2280
      TabIndex        =   252
      Top             =   3960
      Width           =   5295
      Begin VB.CommandButton cmdRegHelp 
         BackColor       =   &H00C000C0&
         Caption         =   "?"
         Height          =   255
         Index           =   280
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   281
         Top             =   1440
         Width           =   255
      End
      Begin VB.CheckBox chkOptReg 
         Caption         =   "Inhibition arrêt système"
         Height          =   255
         Index           =   280
         Left            =   240
         TabIndex        =   280
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Frame Frame19 
         Caption         =   "Configuration clavier "
         ForeColor       =   &H8000000D&
         Height          =   1050
         Left            =   240
         TabIndex        =   253
         Top             =   240
         Width           =   3255
         Begin VB.CheckBox chkOptReg 
            Height          =   255
            Index           =   262
            Left            =   2400
            TabIndex        =   265
            Top             =   720
            Width           =   255
         End
         Begin VB.CheckBox chkOptReg 
            Height          =   255
            Index           =   261
            Left            =   1920
            TabIndex        =   264
            Top             =   720
            Width           =   255
         End
         Begin VB.CheckBox chkOptReg 
            Height          =   255
            Index           =   260
            Left            =   1440
            TabIndex        =   263
            Top             =   720
            Width           =   255
         End
         Begin VB.CheckBox chkOptReg 
            Height          =   255
            Index           =   252
            Left            =   2400
            TabIndex        =   262
            Top             =   480
            Width           =   255
         End
         Begin VB.CheckBox chkOptReg 
            Height          =   255
            Index           =   251
            Left            =   1920
            TabIndex        =   261
            Top             =   480
            Width           =   255
         End
         Begin VB.CommandButton cmdRegHelp 
            BackColor       =   &H00C000C0&
            Caption         =   "?"
            Height          =   255
            Index           =   250
            Left            =   2880
            Style           =   1  'Graphical
            TabIndex        =   255
            Top             =   200
            Width           =   255
         End
         Begin VB.CheckBox chkOptReg 
            Height          =   255
            Index           =   250
            Left            =   1440
            TabIndex        =   254
            Top             =   480
            Width           =   255
         End
         Begin VB.Label Label45 
            Caption         =   "Scroll"
            Height          =   255
            Left            =   2280
            TabIndex        =   260
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label44 
            Caption         =   "Num"
            Height          =   255
            Left            =   1800
            TabIndex        =   259
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label43 
            Caption         =   "Caps"
            Height          =   255
            Left            =   1320
            TabIndex        =   258
            Top             =   240
            Width           =   375
         End
         Begin VB.Label lblIKIUsr 
            Caption         =   "Utilisateur actuel"
            Height          =   255
            Left            =   120
            TabIndex        =   257
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label lblIKIDft 
            Caption         =   "Avant login"
            Height          =   255
            Left            =   120
            TabIndex        =   256
            Top             =   480
            Width           =   855
         End
      End
   End
   Begin VB.Frame frmRegistres 
      Height          =   2655
      Index           =   2
      Left            =   2160
      TabIndex        =   50
      Top             =   4440
      Width           =   5295
      Begin VB.Frame Frame21 
         Caption         =   "Délais (ms) "
         ForeColor       =   &H8000000D&
         Height          =   930
         Left            =   2720
         TabIndex        =   275
         Top             =   1440
         Width           =   2415
         Begin VB.TextBox txtDelaiMenus 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   930
            MaxLength       =   5
            TabIndex        =   278
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton cmdRegSave 
            Caption         =   "S"
            Height          =   255
            Index           =   380
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   277
            Top             =   240
            Width           =   255
         End
         Begin VB.CommandButton cmdRegHelp 
            BackColor       =   &H00808000&
            Caption         =   "?"
            Height          =   255
            Index           =   380
            Left            =   2040
            Style           =   1  'Graphical
            TabIndex        =   276
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label47 
            Caption         =   "Menu"
            Height          =   255
            Left            =   120
            TabIndex        =   279
            Top             =   280
            Width           =   975
         End
      End
      Begin VB.Frame Frame20 
         Caption         =   "Tailles Icones (Pixels)"
         ForeColor       =   &H8000000D&
         Height          =   930
         Left            =   150
         TabIndex        =   266
         Top             =   1440
         Width           =   2415
         Begin VB.CommandButton cmdRegHelp 
            BackColor       =   &H0000C0C0&
            Caption         =   "?"
            Height          =   255
            Index           =   361
            Left            =   2040
            Style           =   1  'Graphical
            TabIndex        =   274
            Top             =   570
            Width           =   255
         End
         Begin VB.CommandButton cmdRegSave 
            Caption         =   "S"
            Height          =   255
            Index           =   361
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   273
            Top             =   570
            Width           =   255
         End
         Begin VB.CommandButton cmdRegHelp 
            BackColor       =   &H0000C0C0&
            Caption         =   "?"
            Height          =   255
            Index           =   360
            Left            =   2040
            Style           =   1  'Graphical
            TabIndex        =   272
            Top             =   255
            Width           =   255
         End
         Begin VB.CommandButton cmdRegSave 
            Caption         =   "S"
            Height          =   255
            Index           =   360
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   271
            Top             =   255
            Width           =   255
         End
         Begin VB.TextBox txtTailleIconesMenuStart 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1050
            MaxLength       =   3
            TabIndex        =   268
            Top             =   555
            Width           =   615
         End
         Begin VB.TextBox txtTailleIconesBureau 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1050
            MaxLength       =   3
            TabIndex        =   267
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label42 
            Caption         =   "Menu Start"
            Height          =   255
            Left            =   120
            TabIndex        =   270
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label41 
            Caption         =   "Bureau"
            Height          =   255
            Left            =   120
            TabIndex        =   269
            Top             =   285
            Width           =   975
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Inhibitions Réseau "
         ForeColor       =   &H8000000D&
         Height          =   1050
         Left            =   150
         TabIndex        =   165
         Top             =   240
         Width           =   2415
         Begin VB.CommandButton cmdRegHelp 
            BackColor       =   &H00808000&
            Caption         =   "?"
            Height          =   255
            Index           =   322
            Left            =   2040
            Style           =   1  'Graphical
            TabIndex        =   171
            Top             =   720
            Width           =   255
         End
         Begin VB.CommandButton cmdRegHelp 
            BackColor       =   &H00808000&
            Caption         =   "?"
            Height          =   255
            Index           =   321
            Left            =   2040
            Style           =   1  'Graphical
            TabIndex        =   170
            Top             =   480
            Width           =   255
         End
         Begin VB.CheckBox chkOptReg 
            Caption         =   "Groupes de travail"
            Height          =   255
            Index           =   322
            Left            =   120
            TabIndex        =   169
            Top             =   720
            Width           =   1695
         End
         Begin VB.CheckBox chkOptReg 
            Caption         =   "Réseau global"
            Height          =   255
            Index           =   321
            Left            =   120
            TabIndex        =   168
            Top             =   480
            Width           =   1695
         End
         Begin VB.CheckBox chkOptReg 
            Caption         =   "Voisinage réseau"
            Height          =   255
            Index           =   320
            Left            =   120
            TabIndex        =   167
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton cmdRegHelp 
            BackColor       =   &H00808000&
            Caption         =   "?"
            Height          =   255
            Index           =   320
            Left            =   2040
            Style           =   1  'Graphical
            TabIndex        =   166
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.CommandButton cmdRegHelp 
         BackColor       =   &H00808000&
         Caption         =   "?"
         Height          =   255
         Index           =   340
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   135
         Top             =   720
         Width           =   255
      End
      Begin VB.CheckBox chkOptReg 
         Caption         =   "Click Droit"
         Height          =   255
         Index           =   340
         Left            =   3120
         TabIndex        =   134
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdRegHelp 
         BackColor       =   &H00808000&
         Caption         =   "?"
         Height          =   255
         Index           =   310
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   103
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton cmdRegHelp 
         BackColor       =   &H00808000&
         Caption         =   "?"
         Height          =   255
         Index           =   330
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   131
         Top             =   960
         Width           =   255
      End
      Begin VB.CheckBox chkOptReg 
         Caption         =   "Icone Internet Exp."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   -1  'True
         EndProperty
         Height          =   255
         Index           =   330
         Left            =   3120
         TabIndex        =   130
         Top             =   960
         Width           =   1695
      End
      Begin VB.CheckBox chkOptReg 
         Caption         =   "Bureau"
         Height          =   255
         Index           =   310
         Left            =   3120
         TabIndex        =   102
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame frmRegistres 
      Height          =   2655
      Index           =   4
      Left            =   8160
      TabIndex        =   146
      Top             =   3840
      Width           =   5295
      Begin VB.CommandButton cmdRegHelp 
         BackColor       =   &H00808000&
         Caption         =   "?"
         Height          =   255
         Index           =   500
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   150
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton cmdRegSave 
         Caption         =   "S"
         Height          =   255
         Index           =   500
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   149
         Top             =   960
         Width           =   255
      End
      Begin VB.ListBox lstLecteurs 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1980
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   147
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label Label26 
         Caption         =   "Cacher lecteurs dans Explorer et Popups Windows"
         Height          =   255
         Left            =   120
         TabIndex        =   148
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Frame frmRegistres 
      Height          =   2655
      Index           =   5
      Left            =   6840
      TabIndex        =   158
      Top             =   1800
      Width           =   5295
      Begin VB.Frame Frame13 
         Caption         =   "Inhibitions Imprimantes "
         ForeColor       =   &H8000000D&
         Height          =   810
         Left            =   120
         TabIndex        =   173
         Top             =   1440
         Width           =   2775
         Begin VB.CheckBox chkOptReg 
            Caption         =   "Supprimer"
            Height          =   255
            Index           =   351
            Left            =   120
            TabIndex        =   176
            Top             =   480
            Width           =   2175
         End
         Begin VB.CheckBox chkOptReg 
            Caption         =   "Ajouter"
            Height          =   255
            Index           =   350
            Left            =   120
            TabIndex        =   175
            Top             =   240
            Width           =   2175
         End
         Begin VB.CommandButton cmdRegHelp 
            BackColor       =   &H0000C0C0&
            Caption         =   "?"
            Height          =   255
            Index           =   350
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   174
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Inhibitions Panneau de configuration de l'affichage "
         ForeColor       =   &H8000000D&
         Height          =   1050
         Left            =   120
         TabIndex        =   159
         Top             =   240
         Width           =   5055
         Begin VB.CommandButton cmdRegHelp 
            BackColor       =   &H0000C0C0&
            Caption         =   "?"
            Height          =   255
            Index           =   300
            Left            =   4680
            Style           =   1  'Graphical
            TabIndex        =   172
            Top             =   240
            Width           =   255
         End
         Begin VB.CheckBox chkOptReg 
            Caption         =   "Access complet au panneau"
            Height          =   255
            Index           =   300
            Left            =   240
            TabIndex        =   164
            Top             =   240
            Width           =   2415
         End
         Begin VB.CheckBox chkOptReg 
            Caption         =   "Onglet Apparance"
            Height          =   255
            Index           =   301
            Left            =   240
            TabIndex        =   163
            Top             =   480
            Width           =   1695
         End
         Begin VB.CheckBox chkOptReg 
            Caption         =   "Onglet Arrière plan"
            Height          =   255
            Index           =   302
            Left            =   240
            TabIndex        =   162
            Top             =   720
            Width           =   1695
         End
         Begin VB.CheckBox chkOptReg 
            Caption         =   "Onglet Eco écran"
            Height          =   255
            Index           =   303
            Left            =   2880
            TabIndex        =   161
            Top             =   240
            Width           =   1815
         End
         Begin VB.CheckBox chkOptReg 
            Caption         =   "Onglet Configuration"
            Height          =   255
            Index           =   304
            Left            =   2880
            TabIndex        =   160
            Top             =   480
            Width           =   1815
         End
      End
   End
   Begin VB.Frame frmRegistres 
      Height          =   2655
      Index           =   3
      Left            =   7560
      TabIndex        =   89
      Top             =   1200
      Width           =   5295
      Begin VB.CommandButton cmdRegHelp 
         BackColor       =   &H0000C0C0&
         Caption         =   "?"
         Height          =   255
         Index           =   420
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   97
         Top             =   1800
         Width           =   255
      End
      Begin VB.Frame Frame8 
         Caption         =   "Inhibitions Fonctions 'Sécurités Win NT' "
         ForeColor       =   &H8000000D&
         Height          =   855
         Left            =   120
         TabIndex        =   98
         Top             =   1560
         Width           =   5055
         Begin VB.CheckBox chkOptReg 
            Caption         =   "Changer Mot de passe"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   -1  'True
            EndProperty
            Height          =   255
            Index           =   422
            Left            =   2280
            TabIndex        =   101
            Top             =   240
            Width           =   2055
         End
         Begin VB.CheckBox chkOptReg 
            Caption         =   "Gestionnaire tâches"
            Height          =   255
            Index           =   421
            Left            =   240
            TabIndex        =   100
            Top             =   480
            Width           =   1815
         End
         Begin VB.CheckBox chkOptReg 
            Caption         =   "Vérouiller Station"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   -1  'True
            EndProperty
            Height          =   255
            Index           =   420
            Left            =   240
            TabIndex        =   99
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.CommandButton cmdRegHelp 
         BackColor       =   &H0000C0C0&
         Caption         =   "?"
         Height          =   255
         Index           =   400
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   96
         Top             =   900
         Width           =   255
      End
      Begin VB.Frame Frame7 
         Caption         =   "Inhibitions Fonctions du bouton 'Démarrer' "
         ForeColor       =   &H8000000D&
         Height          =   1095
         Left            =   120
         TabIndex        =   90
         Top             =   180
         Width           =   5055
         Begin VB.CheckBox chkOptReg 
            Caption         =   "Prgs communs"
            Height          =   255
            Index           =   406
            Left            =   3600
            TabIndex        =   133
            Top             =   240
            Width           =   1335
         End
         Begin VB.CheckBox chkOptReg 
            Caption         =   "Log Off"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   -1  'True
            EndProperty
            Height          =   255
            Index           =   405
            Left            =   1800
            TabIndex        =   132
            Top             =   720
            Width           =   1455
         End
         Begin VB.CheckBox chkOptReg 
            Caption         =   "Arrêt Windows"
            Height          =   255
            Index           =   400
            Left            =   240
            TabIndex        =   95
            Top             =   240
            Width           =   1455
         End
         Begin VB.CheckBox chkOptReg 
            Caption         =   "Exécuter"
            Height          =   255
            Index           =   401
            Left            =   240
            TabIndex        =   94
            Top             =   480
            Width           =   1335
         End
         Begin VB.CheckBox chkOptReg 
            Caption         =   "Rechercher"
            Height          =   255
            Index           =   402
            Left            =   240
            TabIndex        =   93
            Top             =   720
            Width           =   1695
         End
         Begin VB.CheckBox chkOptReg 
            Caption         =   "Config. : Controls"
            Height          =   255
            Index           =   403
            Left            =   1800
            TabIndex        =   92
            Top             =   240
            Width           =   1815
         End
         Begin VB.CheckBox chkOptReg 
            Caption         =   "Config. : Barre tâche"
            Height          =   255
            Index           =   404
            Left            =   1800
            TabIndex        =   91
            Top             =   480
            Width           =   1815
         End
      End
   End
   Begin VB.Frame frmRegistres 
      Height          =   2655
      Index           =   1
      Left            =   7080
      TabIndex        =   20
      Top             =   960
      Width           =   5295
      Begin VB.CommandButton cmdRegSave 
         Caption         =   "S"
         Height          =   255
         Index           =   210
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   1680
         Width           =   255
      End
      Begin VB.CommandButton cmdRegSave 
         Caption         =   "S"
         Height          =   255
         Index           =   200
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton cmdRegHelp 
         BackColor       =   &H00808000&
         Caption         =   "?"
         Height          =   255
         Index           =   210
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   1680
         Width           =   255
      End
      Begin VB.CommandButton cmdRegHelp 
         BackColor       =   &H00808000&
         Caption         =   "?"
         Height          =   255
         Index           =   200
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox txtPopStartUpTitre 
         Height          =   285
         Left            =   840
         TabIndex        =   28
         Text            =   "txtPopStartUpTitre"
         Top             =   480
         Width           =   3735
      End
      Begin VB.TextBox txtPopStartUpContenu 
         Height          =   525
         Left            =   840
         MultiLine       =   -1  'True
         TabIndex        =   27
         Text            =   "Principal.frx":030A
         Top             =   840
         Width           =   3735
      End
      Begin VB.TextBox txtMsgLogon 
         Height          =   735
         Left            =   840
         TabIndex        =   26
         Text            =   "txtMsgLogon"
         Top             =   1680
         Width           =   3735
      End
      Begin VB.Label Label10 
         Caption         =   "login :"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Popup avertissement :"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label7 
         Caption         =   "Titre :"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Contenu :"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Popup"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1680
         Width           =   495
      End
   End
   Begin VB.Frame frmRegistres 
      Height          =   2655
      Index           =   0
      Left            =   6720
      TabIndex        =   19
      Top             =   120
      Width           =   5295
      Begin VB.CommandButton cmdRegHelp 
         Caption         =   "?"
         Height          =   255
         Index           =   100
         Left            =   4800
         TabIndex        =   48
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtALVerifPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2760
         PasswordChar    =   "*"
         TabIndex        =   38
         Top             =   1560
         Width           =   2295
      End
      Begin VB.CommandButton cmdRegSave 
         Caption         =   "Sauver"
         Height          =   375
         Index           =   100
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox txtALUsername 
         Height          =   285
         Left            =   240
         TabIndex        =   36
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox txtALPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   35
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox txtALDomain 
         Height          =   285
         Left            =   2760
         TabIndex        =   34
         Top             =   960
         Width           =   2295
      End
      Begin VB.CheckBox chkAutoLogin 
         Caption         =   "Activer l' AutoLogon"
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblALVerifPassword 
         AutoSize        =   -1  'True
         Caption         =   "Vérification mot de passe"
         Height          =   195
         Left            =   2760
         TabIndex        =   43
         Top             =   1320
         Width           =   1770
      End
      Begin VB.Label lblALUsername 
         AutoSize        =   -1  'True
         Caption         =   "Nom d'utilisateur"
         Height          =   195
         Left            =   240
         TabIndex        =   42
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label lblALPassword 
         AutoSize        =   -1  'True
         Caption         =   "Mot de passe"
         Height          =   195
         Left            =   240
         TabIndex        =   41
         Top             =   1320
         Width           =   960
      End
      Begin VB.Label lblALDomain 
         AutoSize        =   -1  'True
         Caption         =   "Domaine"
         Height          =   195
         Left            =   2760
         TabIndex        =   40
         Top             =   720
         Width           =   630
      End
      Begin VB.Label lblALNonDisponible 
         Caption         =   "Fonction disponible que sous NT 4.0 !!"
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
         Left            =   480
         TabIndex        =   39
         Top             =   2160
         Width           =   3495
      End
   End
   Begin VB.Frame frmOutils 
      Height          =   2655
      Index           =   0
      Left            =   120
      TabIndex        =   111
      Top             =   3240
      Width           =   5295
      Begin VB.Frame Frame9 
         Caption         =   " Shut Down Poste Distant "
         ForeColor       =   &H8000000D&
         Height          =   2415
         Left            =   1680
         TabIndex        =   119
         Top             =   120
         Width           =   3495
         Begin VB.CommandButton cmdHelp 
            Caption         =   "?"
            Height          =   255
            Index           =   10002
            Left            =   3120
            TabIndex        =   251
            Top             =   1800
            Width           =   255
         End
         Begin VB.CommandButton cmdNetShutDown 
            Caption         =   "Stop"
            Height          =   255
            Index           =   2
            Left            =   2240
            TabIndex        =   128
            Top             =   1810
            Width           =   615
         End
         Begin VB.CommandButton cmdNetShutDown 
            Caption         =   "Start"
            Height          =   255
            Index           =   1
            Left            =   1400
            TabIndex        =   127
            Top             =   1810
            Width           =   615
         End
         Begin VB.TextBox txtNetShutDownDelai 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2400
            TabIndex        =   126
            Text            =   "60"
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox txtNetShutDownMessage 
            Height          =   915
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   124
            Top             =   820
            Width           =   3255
         End
         Begin VB.CheckBox chkNetShutDownReboot 
            Caption         =   "Reboot"
            Height          =   255
            Left            =   100
            TabIndex        =   123
            Top             =   2050
            Width           =   975
         End
         Begin VB.CheckBox chkNetShutDownForce 
            Caption         =   "Force"
            Height          =   255
            Left            =   100
            TabIndex        =   122
            Top             =   1850
            Width           =   735
         End
         Begin VB.TextBox txtNetShutDownComputerName 
            Height          =   285
            Left            =   120
            TabIndex        =   120
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label lblNetShutDownMsg 
            Alignment       =   2  'Center
            Caption         =   "xxxxxxxxxxxxxxxxxxxxxxx"
            Height          =   255
            Left            =   1040
            TabIndex        =   129
            Top             =   2100
            Width           =   2175
         End
         Begin VB.Label Label25 
            Caption         =   "Délai (sec)"
            Height          =   255
            Left            =   2520
            TabIndex        =   125
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label11 
            Caption         =   "Nom / IP Poste distant"
            Height          =   255
            Left            =   120
            TabIndex        =   121
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame frmShutDown 
         Caption         =   " Shut Down NT"
         ForeColor       =   &H8000000D&
         Height          =   2415
         Left            =   120
         TabIndex        =   112
         Top             =   120
         Width           =   1455
         Begin VB.CommandButton cmdHelp 
            Caption         =   "?"
            Height          =   255
            Index           =   10001
            Left            =   1140
            TabIndex        =   250
            Top             =   2100
            Width           =   255
         End
         Begin VB.OptionButton optShutDownWin 
            Caption         =   "&Power Off"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   118
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton optShutDownWin 
            Caption         =   "&Shut Down"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   117
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton optShutDownWin 
            Caption         =   "&ReBoot"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   116
            Top             =   840
            Width           =   855
         End
         Begin VB.OptionButton optShutDownWin 
            Caption         =   "&LogOff"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   115
            Top             =   1080
            Width           =   855
         End
         Begin VB.CheckBox chkShutDownForce 
            Caption         =   "&Force"
            Height          =   255
            Left            =   120
            TabIndex        =   114
            Top             =   1440
            Width           =   975
         End
         Begin VB.CommandButton cmdShutDownWin 
            Caption         =   "&Go !"
            Height          =   375
            Left            =   480
            TabIndex        =   113
            Top             =   1920
            Width           =   495
         End
      End
   End
   Begin VB.Frame frmRegistres 
      Height          =   2655
      Index           =   6
      Left            =   120
      TabIndex        =   231
      Top             =   5640
      Width           =   5295
      Begin VB.Frame Frame18 
         Height          =   135
         Left            =   360
         TabIndex        =   243
         Top             =   1650
         Width           =   4575
      End
      Begin VB.CommandButton cmdNetRegConnect 
         Caption         =   "Connexion"
         Height          =   375
         Left            =   3600
         TabIndex        =   237
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdNetRegLocal 
         Caption         =   "Local"
         Height          =   375
         Left            =   3840
         TabIndex        =   236
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtNetRegPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   235
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtNetRegUserName 
         Height          =   285
         Left            =   1440
         TabIndex        =   234
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtNetRegComputerName 
         Height          =   285
         Left            =   1440
         TabIndex        =   233
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "?"
         Height          =   255
         Index           =   600
         Left            =   4920
         TabIndex        =   232
         Top             =   2280
         Width           =   255
      End
      Begin VB.Label Label40 
         Caption         =   "Num. série :"
         Height          =   255
         Left            =   360
         TabIndex        =   246
         Top             =   2320
         Width           =   1335
      End
      Begin VB.Label Label39 
         Caption         =   "Société :"
         Height          =   255
         Left            =   360
         TabIndex        =   245
         Top             =   2120
         Width           =   1335
      End
      Begin VB.Label Label38 
         Caption         =   "Propriétaire :"
         Height          =   255
         Left            =   360
         TabIndex        =   244
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label lblNetRegSerialNumber 
         Caption         =   "xxxxxxxxxxxxxxxxxxxxxxxx"
         Height          =   615
         Left            =   1920
         TabIndex        =   242
         Top             =   1920
         Width           =   3015
      End
      Begin VB.Label lblNetRegMessage 
         Alignment       =   2  'Center
         Caption         =   "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   120
         TabIndex        =   241
         Top             =   1300
         Width           =   5055
      End
      Begin VB.Label Label37 
         Caption         =   "Mot de passe"
         Height          =   255
         Left            =   360
         TabIndex        =   240
         Top             =   990
         Width           =   975
      End
      Begin VB.Label Label36 
         Caption         =   "Utilisateur"
         Height          =   255
         Left            =   360
         TabIndex        =   239
         Top             =   650
         Width           =   855
      End
      Begin VB.Label Label35 
         Caption         =   "PC distant"
         Height          =   255
         Left            =   360
         TabIndex        =   238
         Top             =   270
         Width           =   855
      End
   End
   Begin VB.Timer tmrBouclagePing 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7920
      Top             =   6840
   End
   Begin VB.Frame frmOutils 
      Height          =   2655
      Index           =   3
      Left            =   1800
      TabIndex        =   178
      Top             =   5160
      Width           =   5295
      Begin VB.CheckBox chkPingRepeatPing 
         Caption         =   "Check1"
         Height          =   200
         Left            =   750
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   249
         ToolTipText     =   "Ping répétitif"
         Top             =   1040
         UseMaskColor    =   -1  'True
         Width           =   200
      End
      Begin VB.CommandButton cmdPingLocalIP 
         Caption         =   "IP"
         Height          =   255
         Left            =   4800
         TabIndex        =   248
         ToolTipText     =   "IP = IP local"
         Top             =   750
         Width           =   375
      End
      Begin VB.CommandButton cmdPingAutoPing 
         Caption         =   " Auto Ping"
         Height          =   510
         Index           =   1
         Left            =   2880
         TabIndex        =   247
         Top             =   750
         Width           =   855
      End
      Begin VB.TextBox txtPingAddDNS 
         Height          =   285
         Left            =   2040
         TabIndex        =   188
         Top             =   360
         Width           =   3135
      End
      Begin VB.Frame Frame14 
         Height          =   700
         Left            =   2640
         TabIndex        =   197
         Top             =   600
         Width           =   40
      End
      Begin VB.TextBox txtPingTTL 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1875
         TabIndex        =   195
         Text            =   "128"
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtPingTimeOut 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   193
         Text            =   "1000"
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton cmdPingStartPing 
         Caption         =   "Ping"
         Height          =   510
         Left            =   120
         TabIndex        =   182
         Top             =   750
         Width           =   855
      End
      Begin VB.CommandButton cmdPingAutoPing 
         Caption         =   " Auto IP->DNS"
         Height          =   510
         Index           =   2
         Left            =   3840
         TabIndex        =   192
         Top             =   750
         Width           =   855
      End
      Begin VB.CommandButton cmdPingDNSToIP 
         Caption         =   "<"
         Height          =   220
         Left            =   1750
         TabIndex        =   191
         Top             =   480
         Width           =   220
      End
      Begin VB.CommandButton cmdPingIPToDNS 
         Caption         =   ">"
         Height          =   220
         Left            =   1750
         TabIndex        =   190
         Top             =   240
         Width           =   220
      End
      Begin VB.CommandButton cmdPingClear 
         Caption         =   "Eff."
         Height          =   255
         Left            =   4800
         TabIndex        =   183
         ToolTipText     =   "Efface liste résultats"
         Top             =   1000
         Width           =   375
      End
      Begin VB.ListBox lstPingResultats 
         Height          =   1230
         Left            =   120
         TabIndex        =   181
         Top             =   1320
         Width           =   5055
      End
      Begin VB.TextBox txtPingAddIP 
         Height          =   285
         Left            =   120
         TabIndex        =   179
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label28 
         Caption         =   "TTL"
         Height          =   255
         Left            =   2025
         TabIndex        =   196
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         Caption         =   "TimeOut"
         Height          =   255
         Left            =   1035
         TabIndex        =   194
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         Caption         =   "NetBios - DNS"
         Height          =   255
         Left            =   2040
         TabIndex        =   189
         Top             =   120
         Width           =   3135
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         Caption         =   "IP"
         Height          =   255
         Left            =   120
         TabIndex        =   180
         Top             =   150
         Width           =   1575
      End
   End
   Begin VB.Frame frmOutils 
      Height          =   2655
      Index           =   6
      Left            =   2400
      TabIndex        =   222
      Top             =   5880
      Width           =   5295
      Begin VB.CommandButton cmdHelp 
         BackColor       =   &H00808000&
         Caption         =   "?"
         Height          =   300
         Index           =   10601
         Left            =   4800
         TabIndex        =   225
         Top             =   960
         Width           =   300
      End
      Begin VB.Frame Frame17 
         Caption         =   "Heure atomique  (UTC de Greenwich) "
         ForeColor       =   &H8000000D&
         Height          =   1095
         Left            =   120
         TabIndex        =   223
         Top             =   240
         Width           =   5055
         Begin VB.CommandButton cmdHeureAtomiqueONOFF 
            Caption         =   "OFF"
            Height          =   255
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   227
            Top             =   720
            Width           =   495
         End
         Begin VB.Label lblHeureAtomiqueEnCours 
            Alignment       =   2  'Center
            Caption         =   "Lecture en cours..."
            Height          =   255
            Left            =   120
            TabIndex        =   226
            Top             =   720
            Visible         =   0   'False
            Width           =   4455
         End
         Begin VB.Label lblHeureAtomique 
            Alignment       =   2  'Center
            Caption         =   "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   224
            Top             =   370
            Width           =   4815
         End
      End
   End
   Begin VB.Frame frmOutils 
      Height          =   2655
      Index           =   5
      Left            =   2160
      TabIndex        =   208
      Top             =   5520
      Width           =   5295
      Begin VB.CheckBox chkNetMsgSignatureActive 
         Caption         =   "Activer"
         Height          =   255
         Left            =   2640
         TabIndex        =   217
         Top             =   400
         Width           =   975
      End
      Begin VB.Frame Frame16 
         Caption         =   " Signature "
         Height          =   1230
         Left            =   2520
         TabIndex        =   216
         Top             =   150
         Width           =   2655
         Begin VB.CommandButton cmdNetMsgDefSignature 
            Caption         =   "Définir"
            Height          =   255
            Left            =   1800
            TabIndex        =   221
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton optNetMsgTypeSignature 
            Caption         =   "Fin message"
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   219
            Top             =   850
            Width           =   1215
         End
         Begin VB.OptionButton optNetMsgTypeSignature 
            Caption         =   "Entête"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   218
            Top             =   850
            Width           =   975
         End
         Begin VB.Label lblNetMsgSignature 
            Alignment       =   2  'Center
            Caption         =   "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
            Height          =   255
            Left            =   120
            TabIndex        =   220
            Top             =   580
            Width           =   2415
         End
      End
      Begin VB.CommandButton cmdNetMsgEnvoi 
         Caption         =   "Apperçu"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   215
         Top             =   1000
         Width           =   855
      End
      Begin VB.CommandButton cmdNetMsgAjoutSupp 
         Caption         =   "-"
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   214
         ToolTipText     =   "Supprimer destinataire"
         Top             =   1120
         Width           =   255
      End
      Begin VB.CommandButton cmdNetMsgAjoutSupp 
         Caption         =   "+"
         Height          =   255
         Index           =   0
         Left            =   2160
         TabIndex        =   212
         ToolTipText     =   "Nouveau destinataire"
         Top             =   900
         Width           =   255
      End
      Begin VB.CommandButton cmdNetMsgEnvoi 
         Caption         =   "Envoyer"
         Height          =   375
         Index           =   0
         Left            =   1080
         TabIndex        =   211
         Top             =   1000
         Width           =   975
      End
      Begin VB.TextBox txtNetMsgMsg 
         Height          =   1095
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   210
         Top             =   1440
         Width           =   5055
      End
      Begin VB.ComboBox lstNetMsgDestinataire 
         Height          =   315
         ItemData        =   "Principal.frx":0321
         Left            =   120
         List            =   "Principal.frx":0323
         Style           =   2  'Dropdown List
         TabIndex        =   209
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label lblNetMsgStatus 
         Alignment       =   2  'Center
         Caption         =   "xxxxxxxxxxxxxxxxx"
         Height          =   255
         Left            =   120
         TabIndex        =   230
         Top             =   750
         Width           =   2055
      End
      Begin VB.Label lblNetMsgDestName 
         Alignment       =   2  'Center
         Caption         =   "xxxxxxxxxxxxxxxxx"
         Height          =   255
         Left            =   120
         TabIndex        =   213
         Top             =   550
         Width           =   2055
      End
   End
   Begin VB.Timer tmrHeureAtomique 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   7920
      Top             =   7320
   End
   Begin VB.Frame frmOutils 
      Height          =   2655
      Index           =   1
      Left            =   1320
      TabIndex        =   104
      Top             =   4800
      Width           =   5295
      Begin VB.Frame Frame4 
         Caption         =   "Lancer Windows Control au démarrage"
         ForeColor       =   &H8000000D&
         Height          =   1095
         Index           =   0
         Left            =   960
         TabIndex        =   105
         Top             =   600
         Width           =   3255
         Begin VB.OptionButton optAutoDemarrage 
            Caption         =   "Désactivé"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   109
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optAutoDemarrage 
            Caption         =   "Pour l'utilisateur actuel"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   108
            Top             =   480
            Width           =   1935
         End
         Begin VB.OptionButton optAutoDemarrage 
            Caption         =   "Pour tous les utilisateurs"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   107
            Top             =   720
            Width           =   2055
         End
         Begin VB.CommandButton cmdSauveAutoDemarrage 
            Caption         =   "Valider"
            Height          =   255
            Left            =   2280
            TabIndex        =   106
            Top             =   720
            Width           =   855
         End
      End
   End
   Begin VB.Frame frmOutils 
      Height          =   2655
      Index           =   2
      Left            =   840
      TabIndex        =   136
      Top             =   4320
      Width           =   5295
      Begin VB.Frame Frame10 
         Caption         =   "Afficher panneaux de configuration "
         ForeColor       =   &H8000000D&
         Height          =   1335
         Left            =   120
         TabIndex        =   137
         Top             =   360
         Width           =   5055
         Begin VB.CommandButton cmdAffCPL 
            Caption         =   "Region."
            Height          =   375
            Index           =   6
            Left            =   1320
            TabIndex        =   145
            Top             =   840
            Width           =   735
         End
         Begin VB.CommandButton cmdAffCPL 
            Caption         =   "Serveur"
            Height          =   375
            Index           =   5
            Left            =   480
            TabIndex        =   144
            Top             =   840
            Width           =   735
         End
         Begin VB.CommandButton cmdAffCPL 
            Caption         =   "System"
            Height          =   375
            Index           =   4
            Left            =   3840
            TabIndex        =   143
            Top             =   360
            Width           =   735
         End
         Begin VB.CommandButton cmdAffCPL 
            Caption         =   "Aff."
            Height          =   375
            Index           =   0
            Left            =   480
            TabIndex        =   142
            Top             =   360
            Width           =   735
         End
         Begin VB.CommandButton cmdAffCPL 
            Caption         =   "Clavier"
            Height          =   375
            Index           =   3
            Left            =   3000
            TabIndex        =   141
            Top             =   360
            Width           =   735
         End
         Begin VB.CommandButton cmdAffCPL 
            Caption         =   "Souris"
            Height          =   375
            Index           =   2
            Left            =   2160
            TabIndex        =   140
            Top             =   360
            Width           =   735
         End
         Begin VB.CommandButton cmdAffCPL 
            Caption         =   "Réseau"
            Height          =   375
            Index           =   1
            Left            =   1320
            TabIndex        =   139
            Top             =   360
            Width           =   735
         End
      End
   End
   Begin VB.Frame frmOutils 
      Height          =   2655
      Index           =   4
      Left            =   480
      TabIndex        =   198
      Top             =   3720
      Width           =   5295
      Begin VB.ListBox lstInfosReseau 
         Height          =   1620
         Left            =   120
         TabIndex        =   202
         Top             =   960
         Width           =   5055
      End
      Begin VB.CommandButton cmdGetMACAddress 
         Caption         =   "MAC Adresses"
         Height          =   375
         Left            =   1560
         TabIndex        =   200
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label33 
         Caption         =   "Infos Réseaux :"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   240
         TabIndex        =   201
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Timer tmrDureeLogin 
      Interval        =   1000
      Left            =   7920
      Top             =   7800
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   5530
      _Version        =   393216
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      ForeColor       =   -2147483641
      MouseIcon       =   "Principal.frx":0325
      TabCaption(0)   =   "DB Registres"
      TabPicture(0)   =   "Principal.frx":0777
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "RepereCadresRegistre"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(2)=   "Frame11"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Barre tâches"
      TabPicture(1)   =   "Principal.frx":0793
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdCache(0)"
      Tab(1).Control(1)=   "cmdCache(1)"
      Tab(1).Control(2)=   "cmdCache(2)"
      Tab(1).Control(3)=   "cmdCache(3)"
      Tab(1).Control(4)=   "cmdAffiche(0)"
      Tab(1).Control(5)=   "cmdAffiche(1)"
      Tab(1).Control(6)=   "cmdAffiche(2)"
      Tab(1).Control(7)=   "cmdAffiche(3)"
      Tab(1).Control(8)=   "Label19"
      Tab(1).Control(9)=   "Line2"
      Tab(1).Control(10)=   "lbl5"
      Tab(1).Control(11)=   "Label5"
      Tab(1).Control(12)=   "lblNoms(0)"
      Tab(1).Control(13)=   "lblNoms(1)"
      Tab(1).Control(14)=   "lblNoms(2)"
      Tab(1).Control(15)=   "lblNoms(3)"
      Tab(1).ControlCount=   16
      TabCaption(2)   =   "Outils"
      TabPicture(2)   =   "Principal.frx":07AF
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame15"
      Tab(2).Control(1)=   "optChoixOutils(2)"
      Tab(2).Control(2)=   "optChoixOutils(1)"
      Tab(2).Control(3)=   "optChoixOutils(0)"
      Tab(2).Control(4)=   "optChoixOutils(3)"
      Tab(2).Control(5)=   "RepereCadresOutils"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Système Infos"
      TabPicture(3)   =   "Principal.frx":07CB
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblProcesseur"
      Tab(3).Control(1)=   "lbl2"
      Tab(3).Control(2)=   "lblOS"
      Tab(3).Control(3)=   "lbl1"
      Tab(3).Control(4)=   "Label23"
      Tab(3).Control(5)=   "lblStartUpMode"
      Tab(3).Control(6)=   "lblDureeLogin"
      Tab(3).Control(7)=   "Label27"
      Tab(3).Control(8)=   "Label31"
      Tab(3).Control(9)=   "lblInfosAff"
      Tab(3).Control(10)=   "Frame5"
      Tab(3).Control(11)=   "Frame6"
      Tab(3).Control(12)=   "cmdMSSysInfos"
      Tab(3).ControlCount=   13
      TabCaption(4)   =   "A Propos de..."
      TabPicture(4)   =   "Principal.frx":07E7
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "lblTitreApp2"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "lblTitreApp1"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Line3"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "lblVersion"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "lblDescApp"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "Label12"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "Line4"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "Label16"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "Label17"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "Label18"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).Control(10)=   "Label20"
      Tab(4).Control(10).Enabled=   0   'False
      Tab(4).Control(11)=   "Label15"
      Tab(4).Control(11).Enabled=   0   'False
      Tab(4).Control(12)=   "Label14"
      Tab(4).Control(12).Enabled=   0   'False
      Tab(4).Control(13)=   "Label13"
      Tab(4).Control(13).Enabled=   0   'False
      Tab(4).Control(14)=   "Frame3"
      Tab(4).Control(14).Enabled=   0   'False
      Tab(4).Control(15)=   "cmdAide4"
      Tab(4).Control(15).Enabled=   0   'False
      Tab(4).Control(16)=   "chkAide2"
      Tab(4).Control(16).Enabled=   0   'False
      Tab(4).Control(17)=   "chkAide3"
      Tab(4).Control(17).Enabled=   0   'False
      Tab(4).Control(18)=   "chkAide1"
      Tab(4).Control(18).Enabled=   0   'False
      Tab(4).Control(19)=   "cmdAide1"
      Tab(4).Control(19).Enabled=   0   'False
      Tab(4).Control(20)=   "cmdAide2"
      Tab(4).Control(20).Enabled=   0   'False
      Tab(4).Control(21)=   "cmdAide3"
      Tab(4).Control(21).Enabled=   0   'False
      Tab(4).ControlCount=   22
      Begin VB.Frame Frame15 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   -69480
         TabIndex        =   203
         Top             =   360
         Width           =   855
         Begin VB.OptionButton optMenuChoixOutils 
            Caption         =   "O"
            Height          =   255
            Index           =   3
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   207
            ToolTipText     =   "Options"
            Top             =   400
            Width           =   375
         End
         Begin VB.OptionButton optMenuChoixOutils 
            Caption         =   "W"
            Height          =   255
            Index           =   1
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   206
            ToolTipText     =   "Windows"
            Top             =   50
            Width           =   375
         End
         Begin VB.OptionButton optMenuChoixOutils 
            Caption         =   "I"
            Height          =   255
            Index           =   2
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   205
            ToolTipText     =   "Internet"
            Top             =   400
            Width           =   375
         End
         Begin VB.OptionButton optMenuChoixOutils 
            Caption         =   "R"
            Height          =   255
            Index           =   0
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   204
            ToolTipText     =   "Réseau"
            Top             =   50
            Width           =   375
         End
      End
      Begin VB.OptionButton optChoixOutils 
         Caption         =   "xxxx"
         Height          =   375
         Index           =   2
         Left            =   -69480
         Style           =   1  'Graphical
         TabIndex        =   199
         Top             =   2160
         Width           =   855
      End
      Begin VB.OptionButton optChoixOutils 
         Caption         =   "xxxx"
         Height          =   375
         Index           =   1
         Left            =   -69480
         Style           =   1  'Graphical
         TabIndex        =   177
         Top             =   1680
         Width           =   855
      End
      Begin VB.Frame Frame11 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   -69480
         TabIndex        =   153
         Top             =   360
         Width           =   855
         Begin VB.OptionButton optMenuChoixRegistre 
            Caption         =   "L"
            Height          =   255
            Index           =   0
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   157
            ToolTipText     =   "Logon"
            Top             =   50
            Width           =   375
         End
         Begin VB.OptionButton optMenuChoixRegistre 
            Caption         =   "B"
            Height          =   255
            Index           =   2
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   156
            ToolTipText     =   "Boutons - Commandes"
            Top             =   400
            Width           =   375
         End
         Begin VB.OptionButton optMenuChoixRegistre 
            Caption         =   "A"
            Height          =   255
            Index           =   1
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   155
            ToolTipText     =   "Affichage"
            Top             =   50
            Width           =   375
         End
         Begin VB.OptionButton optMenuChoixRegistre 
            Caption         =   "C"
            Height          =   255
            Index           =   3
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   154
            ToolTipText     =   "Connexion PC distant"
            Top             =   400
            Width           =   375
         End
      End
      Begin VB.CommandButton cmdMSSysInfos 
         Caption         =   "MS Sys Infos"
         Height          =   255
         Left            =   -69840
         TabIndex        =   152
         Top             =   1560
         Width           =   1095
      End
      Begin VB.OptionButton optChoixOutils 
         Caption         =   "xxxx"
         Height          =   375
         Index           =   0
         Left            =   -69480
         Style           =   1  'Graphical
         TabIndex        =   138
         Top             =   1200
         Width           =   855
      End
      Begin VB.OptionButton optChoixOutils 
         Caption         =   "xxxx"
         Height          =   375
         Index           =   3
         Left            =   -69480
         Style           =   1  'Graphical
         TabIndex        =   110
         Top             =   2640
         Width           =   855
      End
      Begin VB.Frame Frame6 
         Caption         =   "Réseau"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1215
         Left            =   -71640
         TabIndex        =   81
         Top             =   1800
         Width           =   2895
         Begin VB.Label lblAdresseIPExt 
            Caption         =   "xxxxxxxxxxxxxxxxxx"
            Height          =   255
            Left            =   1320
            TabIndex        =   229
            Top             =   930
            Width           =   1455
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "          Ext :"
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
            TabIndex        =   228
            Top             =   930
            Width           =   1095
         End
         Begin VB.Label lblAdresseIPInt 
            Caption         =   "xxxxxxxxxxxxxxxxxx"
            Height          =   255
            Left            =   1320
            TabIndex        =   185
            Top             =   690
            Width           =   1455
         End
         Begin VB.Label Label30 
            Caption         =   "Adr. IP Int :"
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
            TabIndex        =   184
            Top             =   690
            Width           =   1095
         End
         Begin VB.Label lblComputerName 
            Caption         =   "xxxxxxxxxxxxxxxxxx"
            Height          =   255
            Left            =   1320
            TabIndex        =   85
            Top             =   450
            Width           =   1455
         End
         Begin VB.Label Label22 
            Caption         =   "Ordinateur :"
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
            TabIndex        =   84
            Top             =   450
            Width           =   1095
         End
         Begin VB.Label lblUserName 
            Caption         =   "xxxxxxxxxxxxxxxxxx"
            Height          =   255
            Left            =   1320
            TabIndex        =   83
            Top             =   230
            Width           =   1455
         End
         Begin VB.Label Label21 
            Caption         =   "Utilisateur :"
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
            TabIndex        =   82
            Top             =   230
            Width           =   1095
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Mémoire"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1215
         Left            =   -74880
         TabIndex        =   72
         Top             =   1800
         Width           =   3015
         Begin VB.Line Line1 
            X1              =   720
            X2              =   2880
            Y1              =   500
            Y2              =   500
         End
         Begin VB.Label lblMemoire 
            Alignment       =   1  'Right Justify
            Caption         =   "xxxxxxxxxxxx"
            Height          =   255
            Index           =   3
            Left            =   1800
            TabIndex        =   80
            Top             =   800
            Width           =   975
         End
         Begin VB.Label lblMemoire 
            Alignment       =   1  'Right Justify
            Caption         =   "xxxxxxxxxxxx"
            Height          =   255
            Index           =   2
            Left            =   720
            TabIndex        =   79
            Top             =   800
            Width           =   975
         End
         Begin VB.Label lblMemoire 
            Alignment       =   1  'Right Justify
            Caption         =   "xxxxxxxxxxxx"
            Height          =   255
            Index           =   1
            Left            =   1800
            TabIndex        =   78
            Top             =   550
            Width           =   975
         End
         Begin VB.Label lblMemoire 
            Alignment       =   1  'Right Justify
            Caption         =   "xxxxxxxxxxxx"
            Height          =   255
            Index           =   0
            Left            =   735
            TabIndex        =   77
            Top             =   550
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Virt. :"
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
            TabIndex        =   76
            Top             =   800
            Width           =   495
         End
         Begin VB.Label Label3 
            Caption         =   "Phys. :"
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
            TabIndex        =   75
            Top             =   550
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Total"
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
            Left            =   2160
            TabIndex        =   74
            Top             =   250
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Disponible"
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
            Left            =   840
            TabIndex        =   73
            Top             =   250
            Width           =   960
         End
      End
      Begin VB.CommandButton cmdAide3 
         BackColor       =   &H00C000C0&
         Caption         =   "S"
         Height          =   255
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   2640
         Width           =   255
      End
      Begin VB.CommandButton cmdAide2 
         BackColor       =   &H00808000&
         Caption         =   "S"
         Height          =   255
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   2340
         Width           =   255
      End
      Begin VB.CommandButton cmdAide1 
         BackColor       =   &H0000C0C0&
         Caption         =   "S"
         Height          =   255
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   2040
         Width           =   255
      End
      Begin VB.CheckBox chkAide1 
         Caption         =   "Ongl..."
         Enabled         =   0   'False
         Height          =   255
         Left            =   3000
         TabIndex        =   64
         Top             =   2085
         Value           =   2  'Grayed
         Width           =   855
      End
      Begin VB.CheckBox chkAide3 
         BackColor       =   &H000000FF&
         Caption         =   "Ongl..."
         Enabled         =   0   'False
         Height          =   255
         Left            =   3000
         TabIndex        =   60
         Top             =   2640
         Width           =   855
      End
      Begin VB.CheckBox chkAide2 
         BackColor       =   &H000080FF&
         Caption         =   "Ongl..."
         Enabled         =   0   'False
         Height          =   255
         Left            =   3000
         TabIndex        =   59
         Top             =   2355
         Width           =   855
      End
      Begin VB.CommandButton cmdAide4 
         BackColor       =   &H000000FF&
         Caption         =   "S"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   1800
         Width           =   255
      End
      Begin VB.Frame Frame3 
         Height          =   135
         Left            =   600
         TabIndex        =   55
         Top             =   1560
         Width           =   5175
      End
      Begin VB.CommandButton cmdCache 
         Height          =   195
         Index           =   0
         Left            =   -72360
         TabIndex        =   12
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton cmdCache 
         Height          =   195
         Index           =   1
         Left            =   -72360
         TabIndex        =   11
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton cmdCache 
         Height          =   195
         Index           =   2
         Left            =   -72360
         TabIndex        =   10
         Top             =   1560
         Width           =   375
      End
      Begin VB.CommandButton cmdCache 
         Height          =   195
         Index           =   3
         Left            =   -72360
         TabIndex        =   9
         Top             =   1800
         Width           =   375
      End
      Begin VB.CommandButton cmdAffiche 
         Height          =   195
         Index           =   0
         Left            =   -71640
         TabIndex        =   8
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton cmdAffiche 
         Height          =   195
         Index           =   1
         Left            =   -71640
         TabIndex        =   7
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton cmdAffiche 
         Height          =   195
         Index           =   2
         Left            =   -71640
         TabIndex        =   6
         Top             =   1560
         Width           =   375
      End
      Begin VB.CommandButton cmdAffiche 
         Height          =   195
         Index           =   3
         Left            =   -71640
         TabIndex        =   5
         Top             =   1800
         Width           =   375
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   1815
         Left            =   -69480
         TabIndex        =   21
         Top             =   1200
         Width           =   855
         Begin VB.OptionButton optChoixRegistre 
            Caption         =   "xxxxxx"
            Height          =   375
            Index           =   3
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   1440
            Width           =   855
         End
         Begin VB.OptionButton optChoixRegistre 
            Caption         =   "xxxxxx"
            Height          =   375
            Index           =   2
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   960
            Width           =   855
         End
         Begin VB.OptionButton optChoixRegistre 
            Caption         =   "xxxxxx"
            Height          =   375
            Index           =   1
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   480
            Width           =   855
         End
         Begin VB.OptionButton optChoixRegistre 
            Caption         =   "xxxxxx"
            Height          =   375
            Index           =   0
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.Label lblInfosAff 
         Caption         =   "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
         Height          =   255
         Left            =   -73440
         TabIndex        =   187
         Top             =   1440
         Width           =   3615
      End
      Begin VB.Label Label31 
         Caption         =   "Affichage :"
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
         Left            =   -74760
         TabIndex        =   186
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label27 
         Caption         =   "Durée login :"
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
         Left            =   -74760
         TabIndex        =   151
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Shape RepereCadresOutils 
         Height          =   2655
         Left            =   -74880
         Top             =   360
         Width           =   5295
      End
      Begin VB.Label lblDureeLogin 
         Caption         =   "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
         Height          =   255
         Left            =   -73440
         TabIndex        =   88
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label lblStartUpMode 
         AutoSize        =   -1  'True
         Caption         =   "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
         Height          =   195
         Left            =   -73440
         TabIndex        =   87
         Top             =   960
         Width           =   2250
      End
      Begin VB.Label Label23 
         Caption         =   "Mode  :"
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
         Left            =   -74760
         TabIndex        =   86
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label13 
         Caption         =   "Immédiat"
         Height          =   255
         Left            =   600
         TabIndex        =   71
         Top             =   2070
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "Reconnexion (Relog)"
         Height          =   255
         Left            =   600
         TabIndex        =   70
         Top             =   2370
         Width           =   1695
      End
      Begin VB.Label Label15 
         Caption         =   "Redémarrage PC"
         Height          =   255
         Left            =   600
         TabIndex        =   69
         Top             =   2670
         Width           =   1215
      End
      Begin VB.Label Label20 
         Caption         =   "Valeur inexistante"
         Height          =   255
         Left            =   3960
         TabIndex        =   65
         Top             =   2115
         Width           =   1275
      End
      Begin VB.Label Label19 
         Caption         =   "Effets immédiats mais volatiles. Tous les éléments réapparaitront au prochain login !"
         Height          =   255
         Left            =   -74760
         TabIndex        =   63
         Top             =   2520
         Width           =   5895
      End
      Begin VB.Label Label18 
         Caption         =   "Clef non trouvée."
         Height          =   255
         Left            =   3960
         TabIndex        =   62
         Top             =   2670
         Width           =   1335
      End
      Begin VB.Label Label17 
         Caption         =   "Donnée inconnue => Pas touche !"
         Height          =   255
         Left            =   3960
         TabIndex        =   61
         Top             =   2385
         Width           =   2475
      End
      Begin VB.Label Label16 
         Caption         =   "Err lecture valeur. Modifications refusées."
         Height          =   255
         Left            =   3360
         TabIndex        =   58
         Top             =   1830
         Width           =   3015
      End
      Begin VB.Line Line4 
         X1              =   2880
         X2              =   2880
         Y1              =   1800
         Y2              =   3000
      End
      Begin VB.Label Label12 
         Caption         =   "Prise en compte de la modification :"
         Height          =   255
         Left            =   240
         TabIndex        =   56
         Top             =   1800
         Width           =   2535
      End
      Begin VB.Label lblDescApp 
         AutoSize        =   -1  'True
         Caption         =   "lblDescApp"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   54
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblVersion 
         Caption         =   "lblVersion"
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
         Left            =   5280
         TabIndex        =   53
         Top             =   750
         Width           =   1095
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   1080
         X2              =   5400
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Shape RepereCadresRegistre 
         Height          =   2655
         Left            =   -74880
         Top             =   360
         Width           =   5295
      End
      Begin VB.Line Line2 
         X1              =   -72480
         X2              =   -71040
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label lbl5 
         Caption         =   "Cache"
         Height          =   255
         Left            =   -72360
         TabIndex        =   18
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Affiche"
         Height          =   255
         Left            =   -71640
         TabIndex        =   17
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblNoms 
         Caption         =   "Barre tâches"
         Height          =   255
         Index           =   0
         Left            =   -73800
         TabIndex        =   16
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblNoms 
         Caption         =   "Prgs barre tâches"
         Height          =   255
         Index           =   1
         Left            =   -73800
         TabIndex        =   15
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblNoms 
         Caption         =   "Bouton Start"
         Height          =   255
         Index           =   2
         Left            =   -73800
         TabIndex        =   14
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label lblNoms 
         Caption         =   "Horloge"
         Height          =   255
         Index           =   3
         Left            =   -73800
         TabIndex        =   13
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label lbl1 
         Caption         =   "OS :"
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
         Left            =   -74760
         TabIndex        =   4
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblOS 
         Caption         =   "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
         Height          =   255
         Left            =   -73440
         TabIndex        =   3
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label lbl2 
         Caption         =   "Processeur(s) :"
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
         Left            =   -74760
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblProcesseur 
         Caption         =   "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
         Height          =   255
         Left            =   -73440
         TabIndex        =   1
         Top             =   720
         Width           =   4695
      End
      Begin VB.Label lblTitreApp1 
         BackStyle       =   0  'Transparent
         Caption         =   "Windows Control pour NT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Left            =   1170
         TabIndex        =   51
         Top             =   330
         Width           =   4335
      End
      Begin VB.Label lblTitreApp2 
         Caption         =   "Windows Control pour NT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   495
         Left            =   1200
         TabIndex        =   52
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.Image imgIcone 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   7920
      Picture         =   "Principal.frx":0803
      Top             =   8280
      Width           =   480
   End
End
Attribute VB_Name = "Frm_Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FlagRefreshEnCours As Integer         ' Flag pour ne pas sauver au démarrage
Dim DernierChoixMenuReg(3) As Integer     ' Dernier choix menu registre
Dim DernierChoixMenuOutils(3) As Integer  ' Dernier choix menu outils
Dim FlagDemandeArret As Boolean           ' Flag appui touche ESC

Private Type TypeGetMACAddress
   NbReturn As Integer
   MACAddress(1 To 20) As String
End Type




'Dim HKey As Long

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then FlagDemandeArret = True
End Sub

'============================================================================
'=============================== FORMULAIRE =================================
'============================================================================

Private Sub Form_Load()
   Me.Icon = Me.imgIcone.Picture
   Me.Caption = App.Title
  
   RepereCadresRegistre.Visible = False
   RepereCadresOutils.Visible = False
   Frm_Principal.Width = SSTab1.Width + 90
   Frm_Principal.Height = SSTab1.Height + 380
   
   optChoixRegistre(0).Value = True
   optMenuChoixRegistre(0).Value = True
   optChoixOutils(0).Value = True
   optMenuChoixOutils(0).Value = True

 '  optPingTypeAdresse(0).Value = True
   optShutDownWin(1).Value = True
   lblNetShutDownMsg.Caption = ""
   lblDureeLogin.Caption = ""
   lblHeureAtomique.Caption = "<Arrêté>"
   
   Call cmdNetRegLocal_Click     ' Connexion en local

'   txtNetShutDownMessage.Text = "Activé par \\" & _
                                NetInfo.ComputerName & "\" & NetInfo.UserName & vbCrLf

   lblVersion.Caption = "Ver. " & App.Major & "." & App.Minor & "." & App.Revision
   lblDescApp.Caption = "Par SERRON Dominique     (Jan 2000 - Juin 2001)" & vbCrLf & _
                        "Ecrit en Visual Basic 5.0 et 6.0"
      
   SSTab1.Tab = 3 ' Onglet Infos
   
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Static rec As Boolean
   Static Msg As Long

   Msg = X / Screen.TwipsPerPixelX
   If rec = False Then
     rec = True
     Select Case Msg
        Case DOUBLE_CLICK_GAUCHE:
            Me.WindowState = vbNormal
        Case BOUTON_GAUCHE_POUSSE:
        Case BOUTON_GAUCHE_LEVE:
        Case DOUBLE_CLICK_DROIT:
        Case BOUTON_DROIT_POUSSE:
        Case BOUTON_DROIT_LEVE:
     End Select
     rec = False
   End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Dim Msg As String
   Dim Ret As Integer

   Select Case UnloadMode
      Case vbFormControlMenu:    Msg = "Menu Système"
      Case vbFormCode:           Msg = "Code interne"
      Case vbAppWindows:         Msg = "Fermeture de Windows"
      Case vbAppTaskManager:     Msg = "Gestionnaire des tâches"
      Case vbFormMDIForm:        Msg = "Fermeture MDI mère"
      Case Else:                 Msg = "<Inconnu>"
   End Select

   Ret = MsgBox("Quitter ? (" & Msg & ")", vbOKCancel + vbQuestion)
   If Ret = vbOK Then
     Call QuitterProg
    Else
         Cancel = 1
   End If
 End Sub


'============================================================================
'======================== GESTION ONGLETS / MENUS ===========================
'============================================================================

Private Sub SSTab1_Click(PreviousTab As Integer)
   ' Animations onglets
   Select Case SSTab1.Tab
      Case 0
          optChoixOutils_Click (-1)        ' Efface tous les frames
          optChoixRegistre_Click (99)      ' Affiche frame sélectionné
      Case 2
          optChoixRegistre_Click (-1)      ' Efface tous les frames
          optChoixOutils_Click (0)         ' Affiche frame sélectionné
      Case Else
          optChoixOutils_Click (-1)        ' Efface tous les frames
          optChoixRegistre_Click (-1)      ' Efface tous les frames
   End Select
   
   ' Maj infos dans l'onglet
   Select Case SSTab1.Tab
      Case 3  ' System Infos
          AfficheInfosSystème
   End Select
End Sub

Private Sub optChoixRegistre_Click(Index As Integer)
   Dim i As Integer
   Dim Menu As Integer
   Dim Choix As Integer
   Dim NewIndex As Integer
   
   For i = 0 To 3
      If (optMenuChoixRegistre(i).Value = True) Then Menu = i
   Next
   For i = 0 To 3
      If (optChoixRegistre(i).Value = True) Then Choix = i
   Next
   
   Select Case (Menu * 10) + Choix
      Case 0: NewIndex = 0
      Case 1: NewIndex = 1
      Case 2: NewIndex = 7
      Case 10: NewIndex = 2
      Case 11: NewIndex = 4
      Case 20: NewIndex = 3
      Case 21: NewIndex = 5
      Case 30: NewIndex = 6
      Case Else
           NewIndex = 0
           MsgBox ("ERR INDEX (CODE " & ((Menu * 10) + Index) & ")")
   End Select
   
   For i = 0 To 7
      If ((i = NewIndex) And (Index <> -1)) Then
        frmRegistres(i).Left = RepereCadresRegistre.Left
        frmRegistres(i).Top = RepereCadresRegistre.Top
        DernierChoixMenuReg(Menu) = Choix
        AffInfosRegistres (i)
       Else
            frmRegistres(i).Left = -10000
            frmRegistres(i).Top = -10000
      End If
   Next
End Sub

Private Sub optChoixOutils_Click(Index As Integer)
   Dim i As Integer
   Dim Menu As Integer
   Dim Choix As Integer
   Dim NewIndex As Integer
   
   For i = 0 To 3
      If (optMenuChoixOutils(i).Value = True) Then Menu = i
   Next
   For i = 0 To 3
      If (optChoixOutils(i).Value = True) Then Choix = i
   Next
   
   Select Case (Menu * 10) + Choix
      Case 0: NewIndex = 3
      Case 1: NewIndex = 5
      Case 2: NewIndex = 4
      Case 10: NewIndex = 2
      Case 13: NewIndex = 0
      Case 20: NewIndex = 6
      Case 30: NewIndex = 1
      Case Else
           NewIndex = 0
           MsgBox ("ERR INDEX (CODE " & ((Menu * 10) + Index) & ")")
   End Select
   
   For i = 0 To 6
      If ((i = NewIndex) And (Index <> -1)) Then
        frmOutils(i).Left = RepereCadresOutils.Left
        frmOutils(i).Top = RepereCadresOutils.Top
        DernierChoixMenuOutils(Menu) = Choix
        AffInfosOutils (i)
       Else
            frmOutils(i).Left = -10000
            frmOutils(i).Top = -10000
      End If
   Next

End Sub

Private Sub optMenuChoixRegistre_Click(Index As Integer)
   Select Case Index
      Case 0
          optChoixRegistre(0).Caption = "AutoLog"
          optChoixRegistre(1).Caption = "Messages"
          optChoixRegistre(2).Caption = "Options"
          optChoixRegistre(3).Caption = ""
      Case 1
          optChoixRegistre(0).Caption = "Bureau"
          optChoixRegistre(1).Caption = "Lecteurs"
          optChoixRegistre(2).Caption = ""
          optChoixRegistre(3).Caption = ""
      Case 2
          optChoixRegistre(0).Caption = "Cmds"
          optChoixRegistre(1).Caption = "Config."
          optChoixRegistre(2).Caption = ""
          optChoixRegistre(3).Caption = ""
      Case 3
          optChoixRegistre(0).Caption = ""
          optChoixRegistre(1).Caption = ""
          optChoixRegistre(2).Caption = ""
          optChoixRegistre(3).Caption = ""
   End Select
   optChoixRegistre(0).Visible = (optChoixRegistre(0).Caption <> "")
   optChoixRegistre(1).Visible = (optChoixRegistre(1).Caption <> "")
   optChoixRegistre(2).Visible = (optChoixRegistre(2).Caption <> "")
   optChoixRegistre(3).Visible = (optChoixRegistre(3).Caption <> "")
   
   If (optChoixRegistre(DernierChoixMenuReg(Index)).Value = False) Then
     optChoixRegistre(DernierChoixMenuReg(Index)).Value = True
    Else
         optChoixRegistre_Click DernierChoixMenuReg(Index)
   End If
End Sub

Private Sub optMenuChoixOutils_Click(Index As Integer)
   Select Case Index
      Case 0
          optChoixOutils(0).Caption = "Ping"        ' Reseau
          optChoixOutils(1).Caption = "Msg"
          optChoixOutils(2).Caption = "MAC @"
          optChoixOutils(3).Caption = ""
      Case 1
          optChoixOutils(0).Caption = "Config"      ' Windows
          optChoixOutils(1).Caption = ""
          optChoixOutils(2).Caption = ""
          optChoixOutils(3).Caption = "ShutDown"
      Case 2
          optChoixOutils(0).Caption = "Divers"      ' Internet
          optChoixOutils(1).Caption = ""
          optChoixOutils(2).Caption = ""
          optChoixOutils(3).Caption = ""
      Case 3
          optChoixOutils(0).Caption = "StartUp"     ' Options
          optChoixOutils(1).Caption = ""
          optChoixOutils(2).Caption = ""
          optChoixOutils(3).Caption = ""
   End Select
   optChoixOutils(0).Visible = (optChoixOutils(0).Caption <> "")
   optChoixOutils(1).Visible = (optChoixOutils(1).Caption <> "")
   optChoixOutils(2).Visible = (optChoixOutils(2).Caption <> "")
   optChoixOutils(3).Visible = (optChoixOutils(3).Caption <> "")
   
   If (optChoixOutils(DernierChoixMenuOutils(Index)).Value = False) Then
     optChoixOutils(DernierChoixMenuOutils(Index)).Value = True
    Else
         optChoixOutils_Click DernierChoixMenuOutils(Index)
   End If
End Sub


'============================================================================
'============================= ONGLET REGISTRE ==============================
'============================================================================

Private Sub cmdRegHelp_Click(Index As Integer)
   RefAide = Index
   Frm_AideReg.Show , Me
End Sub

Private Sub cmdRegSave_Click(Index As Integer)
   SauverReg Index
End Sub

Private Sub chkOptReg_Click(Index As Integer)
   SauverReg Index
End Sub

Sub SauverReg(Index As Integer)
   If FlagRefreshEnCours = 1 Then Exit Sub
   Select Case Index
         Case IDX_ALOGIN
          If txtALPassword.Text = txtALVerifPassword.Text Then
            RegALActif = IIf(chkAutoLogin.Value = 1, 1, 0)
            RegALUserName = txtALUsername.Text
            RegALDomain = txtALDomain.Text
            RegALPassword = IIf(chkAutoLogin.Value = 1, txtALPassword.Text, FLG_SZ_AEFFACER)
            cmdRegSave(Index).Enabled = False
           Else
               Call MsgBox("Mots de passe différents !!", vbCritical + vbOKOnly, "Erreur")
               txtALPassword.SetFocus
               Exit Sub
          End If
      Case IDX_MSGLOG
          RegPopStartUpTitre = txtPopStartUpTitre.Text
          RegPopStartUpText = txtPopStartUpContenu.Text
          cmdRegSave(Index).Enabled = False
      Case IDX_POPLOG
          RegLogonPrompt = txtMsgLogon.Text
          cmdRegSave(Index).Enabled = False
      
      Case IDX_KEYB11, IDX_KEYB12, IDX_KEYB13
          RegInitialKeyboardIndicatorsDefault = IIf(chkOptReg(IDX_KEYB11).Value = 1, 1, 0) + _
                                                IIf(chkOptReg(IDX_KEYB12).Value = 1, 2, 0) + _
                                                IIf(chkOptReg(IDX_KEYB13).Value = 1, 4, 0)
                                         
      Case IDX_KEYB21, IDX_KEYB22, IDX_KEYB23
          RegInitialKeyboardIndicatorsCurrentUser = IIf(chkOptReg(IDX_KEYB21).Value = 1, 1, 0) + _
                                                    IIf(chkOptReg(IDX_KEYB22).Value = 1, 2, 0) + _
                                                    IIf(chkOptReg(IDX_KEYB23).Value = 1, 4, 0)
                                         
      Case IDX_ISDLOG: RegInhShutdownDansLogin = chkOptReg(Index).Value
                                         
      Case IDX_CPAGEN: RegInhCPLAffAccesCPL = chkOptReg(Index).Value
      Case IDX_CPAAPP: RegInhCPLAffOgtApp = chkOptReg(Index).Value
      Case IDX_CPABKG: RegInhCPLAffOgtAP = chkOptReg(Index).Value
      Case IDX_CPASCR: RegInhCPLAffOgtEco = chkOptReg(Index).Value
      Case IDX_CPACNF: RegInhCPLAffOgtCnf = chkOptReg(Index).Value
    
      Case IDX_INHDSK: RegInhDesktop = chkOptReg(Index).Value
      Case IDX_ICONET: RegInhIconeReseau = chkOptReg(Index).Value
      Case IDX_ENTNET: RegInhReseauGlobal = chkOptReg(Index).Value
      Case IDX_NETWKC: RegInhContenuGroupesTravail = chkOptReg(Index).Value
      Case IDX_ICOWEB: RegInhIconeIE = chkOptReg(Index).Value
      Case IDX_CLKDRT: RegInhClicDroit = chkOptReg(Index).Value
      Case IDX_IMPAJT: RegInhImpAjout = chkOptReg(Index).Value
      Case IDX_IMPSUP: RegInhImpSupp = chkOptReg(Index).Value

      Case IDX_STPWIN: RegInhArretWindows = chkOptReg(Index).Value
      Case IDX_LOGOFF: RegInhLogOff = chkOptReg(Index).Value
      Case IDX_CMDRUN: RegInhCmdRun = chkOptReg(Index).Value
      Case IDX_CMDFND: RegInhCmdFind = chkOptReg(Index).Value
      Case IDX_CNFGEN: RegInhCnfGen = chkOptReg(Index).Value
      Case IDX_CNFTKB: RegInhCnfTaskBar = chkOptReg(Index).Value
      Case IDX_PRGCMN: RegInhPrgsCommuns = chkOptReg(Index).Value
      
      Case IDX_LOCKST: RegInhLockStation = chkOptReg(Index).Value
      Case IDX_TSKMGR: RegInhTaskManager = chkOptReg(Index).Value
      Case IDX_CHGPWD: RegInhChangePassword = chkOptReg(Index).Value
      
      Case IDX_HIDDRV
          Dim Config As String
          Dim i As Integer
          
          Config = ""
          For i = 0 To 25
             Config = Config & IIf(lstLecteurs.Selected(i) = True, "1", "0")
          Next
          RegCacheLecteurs = Config
          cmdRegSave(Index).Enabled = False
      
      Case IDX_SIZICD
          RegTailleIconesBureau = Val(txtTailleIconesBureau.Text)
          cmdRegSave(IDX_SIZICD).Enabled = False
          
      Case IDX_SIZICS
          RegTailleIconesMenuStart = Val(txtTailleIconesMenuStart.Text)
          cmdRegSave(IDX_SIZICS).Enabled = False
          
      Case IDX_TMRMNU
          RegDelaiMenus = Val(txtDelaiMenus.Text)
          cmdRegSave(IDX_TMRMNU).Enabled = False
          
      Case Else
          MsgBox ("SAUVER : CODE INCONNU (" & Index & ")")
   End Select
End Sub

'============================= MENU AUTO LOGIN ==============================

Private Sub chkAutoLogin_Click()
    cmdRegSave(IDX_ALOGIN).Enabled = True
    If chkAutoLogin.Value = vbUnchecked Then
        lblALUsername.Enabled = False
        lblALPassword.Enabled = False
        lblALVerifPassword.Enabled = False
        lblALDomain.Enabled = False
        txtALUsername.Enabled = False
        txtALPassword.Enabled = False
        txtALVerifPassword.Enabled = False
        txtALDomain.Enabled = False
    Else
        lblALUsername.Enabled = True
        lblALPassword.Enabled = True
        lblALVerifPassword.Enabled = True
        lblALDomain.Enabled = True
        txtALUsername.Enabled = True
        txtALPassword.Enabled = True
        txtALVerifPassword.Enabled = True
        txtALDomain.Enabled = True
    End If
End Sub

Private Sub txtALDomain_Change()
   cmdRegSave(IDX_ALOGIN).Enabled = True
End Sub
Private Sub txtALDomain_GotFocus()
    txtALDomain.SelStart = 0:   txtALDomain.SelLength = Len(txtALDomain.Text)
End Sub

Private Sub txtALPassword_Change()
   cmdRegSave(IDX_ALOGIN).Enabled = True
End Sub
Private Sub txtALPassword_GotFocus()
   txtALPassword.SelStart = 0:   txtALPassword.SelLength = Len(txtALPassword.Text)
End Sub

Private Sub txtALUsername_Change()
   cmdRegSave(IDX_ALOGIN).Enabled = True
End Sub
Private Sub txtALUsername_GotFocus()
    txtALUsername.SelStart = 0:   txtALUsername.SelLength = Len(txtALUsername.Text)
End Sub

Private Sub txtALVerifPassword_Change()
   cmdRegSave(IDX_ALOGIN).Enabled = True
End Sub
Private Sub txtALVerifPassword_GotFocus()
   txtALVerifPassword.SelStart = 0:   txtALVerifPassword.SelLength = Len(txtALVerifPassword.Text)
End Sub

'================================= MENU LOGIN ===============================

Private Sub txtPopStartUpTitre_Change()
   cmdRegSave(IDX_MSGLOG).Enabled = True
End Sub
Private Sub txtPopStartUpTitre_GotFocus()
    txtPopStartUpTitre.SelStart = 0:    txtPopStartUpTitre.SelLength = Len(txtPopStartUpTitre.Text)
End Sub

Private Sub txtPopStartUpContenu_Change()
   cmdRegSave(IDX_MSGLOG).Enabled = True
End Sub
Private Sub txtPopStartUpContenu_GotFocus()
   txtPopStartUpContenu.SelStart = 0:    txtPopStartUpContenu.SelLength = Len(txtPopStartUpContenu.Text)
End Sub

Private Sub txtMsgLogon_Change()
   cmdRegSave(IDX_POPLOG).Enabled = True
End Sub
Private Sub txtMsgLogon_GotFocus()
   txtMsgLogon.SelStart = 0:    txtMsgLogon.SelLength = Len(txtMsgLogon.Text)
End Sub

'=============================== MENU BUREAU ================================

Private Sub txtTailleIconesBureau_Change()
   Dim iTemp As Long
   Dim NewState As Boolean
   
   NewState = False
   If IsNumeric(txtTailleIconesBureau.Text) = True Then
     iTemp = Val(txtTailleIconesBureau.Text)
     If ((iTemp > 0) And (iTemp <= 200)) Then NewState = True
   End If
   cmdRegSave(IDX_SIZICD).Enabled = NewState
End Sub

Private Sub txtTailleIconesMenuStart_Change()
   Dim iTemp As Long
   Dim NewState As Boolean
   
   NewState = False
   If IsNumeric(txtTailleIconesMenuStart.Text) = True Then
     iTemp = Val(txtTailleIconesMenuStart.Text)
     If ((iTemp > 0) And (iTemp <= 200)) Then NewState = True
   End If
   cmdRegSave(IDX_SIZICS).Enabled = NewState
End Sub

Private Sub txtDelaiMenus_Change()
   Dim iTemp As Long
   Dim NewState As Boolean
   
   NewState = False
   If IsNumeric(txtDelaiMenus.Text) = True Then
     iTemp = Val(txtDelaiMenus.Text)
     If ((iTemp > 0) And (iTemp <= 20000)) Then NewState = True
   End If
   cmdRegSave(IDX_TMRMNU).Enabled = NewState
End Sub

'=========================== MENU CACHE LECTEURS ============================

Private Sub lstLecteurs_ItemCheck(Item As Integer)
   cmdRegSave(IDX_HIDDRV).Enabled = True
End Sub

'======================== MENU CONNEXION DISTANTE ===========================

Private Sub cmdNetRegConnect_Click()
   Dim Ret As Long
   Dim szRet As String
   Dim OpenKeyVal As Long
   Dim OpenHiveVal As Long
   
   lblNetRegMessage = "Demande en cours..."
   Screen.MousePointer = vbHourglass
   DoEvents
   
   Ret = GetIPCConnection(Trim$(txtNetRegComputerName.Text), _
                          txtNetRegUserName.Text, _
                          txtNetRegPassword.Text)

   Screen.MousePointer = vbDefault
   
   If Ret = 0 Then
     lblNetRegMessage = "Connecté à " & Trim$(txtNetRegComputerName.Text)
     Call AfficheSerialNumEtOwnerName
     optMenuChoixRegistre(3).BackColor = vbGreen
     cmdNetRegConnect.Enabled = False
     cmdNetRegLocal.Enabled = True
    Else
          lblNetRegMessage = DecodeSystemError(Ret)
          optMenuChoixRegistre(3).BackColor = vbButtonFace
          Call AfficheSerialNumEtOwnerName
   End If
End Sub
Private Sub cmdNetRegLocal_Click()
   Call CloseAllRegKey
   lblNetRegMessage = "Mode local"
   Call AfficheSerialNumEtOwnerName
   optMenuChoixRegistre(3).BackColor = vbButtonFace
   cmdNetRegConnect.Enabled = True
   cmdNetRegLocal.Enabled = False
End Sub

Sub AfficheSerialNumEtOwnerName()
   lblNetRegSerialNumber.Caption = RegisteredOwner & vbCrLf & _
                                   RegisteredOrganization & vbCrLf & _
                                   WindowsSerialNumber
End Sub

'============================================================================
'========================== ONGLET BARRE TACHES =============================
'============================================================================

Private Sub cmdAffiche_Click(Index As Integer)
   Dim Ret As Long

   Select Case Index
      Case 0
          Ret = Elmts.AfficheElement(ELMT_BARRE_TACHE)
      Case 1
          Ret = Elmts.AfficheElement(ELMT_PRGS_BARRE_TACHE)
      Case 2
          Ret = Elmts.AfficheElement(ELMT_BOUTON_START)
      Case 3
          Ret = Elmts.AfficheElement(ELMT_HORLOGE)
      Case Else
          Ret = 0
   End Select
   If Ret <> 1 Then Call MsgBox("Une erreur est survenue !! (code " & Ret & ")", _
                                 vbExclamation + vbOKOnly, "ERREUR")
End Sub

Private Sub cmdCache_Click(Index As Integer)
   Dim Ret As Long

   Select Case Index
      Case 0
          Ret = Elmts.CacheElement(ELMT_BARRE_TACHE)
      Case 1
          Ret = Elmts.CacheElement(ELMT_PRGS_BARRE_TACHE)
      Case 2
          Ret = Elmts.CacheElement(ELMT_BOUTON_START)
      Case 3
          Ret = Elmts.CacheElement(ELMT_HORLOGE)
      Case Else
          Ret = 0
   End Select
   If Ret <> 1 Then Call MsgBox("Une erreur est survenue !! (code " & Ret & ")", _
                                 vbExclamation + vbOKOnly, "ERREUR")
End Sub

'============================================================================
'============================= ONGLET OUTILS ================================
'============================================================================

Private Sub cmdHelp_Click(Index As Integer)
   RefAide = Index
   Frm_Aide.Show , Me
End Sub

Private Sub tmrHeureAtomique_Timer()
'   If (frmOutils(6).Left > 0) Then
     Static Ret As String
     lblHeureAtomiqueEnCours.Visible = True
     DoEvents
     Ret = Inet.ReadFile(INETADD_HEUREATOMIQUE)
     lblHeureAtomiqueEnCours.Visible = False
     DoEvents
     If (Ret <> "") Then
       lblHeureAtomique.Caption = Mid$(Ret, 2, 23)
      Else
           lblHeureAtomique.Caption = "< Echec connexion >"
     End If
'   End If
End Sub


'=============== MENU SHUT DOWN / SHUT DOWN POSTE DISTANT ===================

Private Sub cmdShutDownWin_Click()
   Dim Msg As String
   Dim Opt As Long
   Dim Ret As Integer
   Dim Chx As Integer

   Msg = "Voulez-vous vraiment arrêter Windows ?"
   Opt = vbQuestion + vbOKCancel + vbDefaultButton2 + vbSystemModal

   Ret = MsgBox(Msg, Opt, "Confirmation de ShutDown")
   If optShutDownWin(1).Value = True Then Chx = 1
   If optShutDownWin(2).Value = True Then Chx = 2
   If optShutDownWin(3).Value = True Then Chx = 3
   If optShutDownWin(4).Value = True Then Chx = 4
   If (Ret = vbOK) Then Call ShutDownNT(Chx, chkShutDownForce.Value)
End Sub

Private Sub cmdNetShutDown_Click(Index As Integer)
   Dim Ret As Long
   Select Case Index
      Case 1
          If (IsNumeric(txtNetShutDownDelai) = False) Then
            lblNetShutDownMsg.Caption = "DELAI INVALIDE"
            Exit Sub
          End If
          Screen.MousePointer = vbHourglass
          lblNetShutDownMsg.Caption = "Demande START en cours..."
          DoEvents
          Ret = InitiateShutDownNetComputer(txtNetShutDownComputerName & vbNullString, _
                                            txtNetShutDownMessage, _
                                            Val(txtNetShutDownDelai), _
                                            chkNetShutDownForce, _
                                            chkNetShutDownReboot)
          Screen.MousePointer = vbDefault
          lblNetShutDownMsg.Caption = IIf(Ret = 1, "OK", "== ECHEC ==")
      Case 2
          Screen.MousePointer = vbHourglass
          lblNetShutDownMsg.Caption = "Demande STOP en cours..."
          DoEvents
          Ret = AbortShutDownNetComputer(txtNetShutDownComputerName)
          Screen.MousePointer = vbDefault
          lblNetShutDownMsg.Caption = IIf(Ret = 1, "OK", "ECHEC !!")
   End Select
End Sub

Private Sub chkNetShutDownForce_Click()
   lblNetShutDownMsg = ""
End Sub
Private Sub chkNetShutDownReboot_Click()
   lblNetShutDownMsg = ""
End Sub
Private Sub txtNetShutDownComputerName_Change()
   lblNetShutDownMsg = ""
End Sub
Private Sub txtNetShutDownDelai_Change()
   lblNetShutDownMsg = ""
End Sub
Private Sub txtNetShutDownMessage_Change()
   lblNetShutDownMsg = ""
End Sub

'=============================== MENU CONFIG ================================

Private Sub cmdAffCPL_Click(Index As Integer)
   Me.MousePointer = vbHourglass
   Select Case Index
      Case 0
          Shell "rundll32 shell32,Control_RunDLL desk.cpl"    ' Affichage
      Case 1
          Shell "rundll32 shell32,Control_RunDLL ncpa.cpl"    ' Réseau
      Case 2
          Shell "rundll32 shell32,Control_RunDLL main.cpl @0" ' Souris
      Case 3
          Shell "rundll32 shell32,Control_RunDLL main.cpl @1" ' Clavier
      Case 4
          Shell "rundll32 shell32,Control_RunDLL sysdm.cpl"   ' System
      Case 5
          Shell "rundll32 shell32,Control_RunDLL srvmgr.cpl"  ' Serveur
      Case 6
          Shell "rundll32 shell32,Control_RunDLL intl.cpl"    ' Regional setting
          
   End Select
   Me.MousePointer = vbDefault
End Sub

'================================= MENU PING ================================

Private Sub cmdPingStartPing_Click()
   Dim Ret As Integer
   Dim MsgToSend As String
   Dim Msg As String
      
   tmrBouclagePing.Enabled = False  ' Timer utilisé pour Ping en boucle

   Screen.MousePointer = vbHourglass
   cmdPingStartPing.Enabled = False
   
   MsgToSend = "Ping string from WindowsControl for NT4.0"
   
   If txtPingAddDNS.Text <> "" Then
     PingAffNouvMsg ("[" & txtPingAddDNS.Text & "] => Ping en cours...")
     Ret = Net.DoPing(txtPingAddDNS.Text, ADD_NETBIOS_DNS, MsgToSend, _
                      Val(txtPingTimeOut.Text), Val(txtPingTTL.Text))  ' NetBios / DNS
     Beep
    Else
         PingAffNouvMsg ("[" & txtPingAddIP.Text & "] => Ping en cours...")
         Ret = Net.DoPing(txtPingAddIP.Text, ADD_IP, MsgToSend, _
                      Val(txtPingTimeOut.Text), Val(txtPingTTL.Text))  ' IP
   End If
   
   If Ret = 1 Then
     PingRemplaceMsg ("[" & Net.PingIPReply & "] - OK - Atteind en " & Net.PingRoundTripTime & " ms.")
     Call Sound(1500, 50)
    Else
         PingRemplaceMsg ("[" & Net.PingIPReply & "] - ERR - " & Net.PingStatusMsg)
         Call Sound(300, 50)
   End If
      
   cmdPingStartPing.Enabled = True
   Screen.MousePointer = vbDefault
   
   If chkPingRepeatPing.Value = 1 Then tmrBouclagePing.Enabled = True
End Sub

Private Sub txtPingAddIP_Change()
   If txtPingAddIP.Text <> "" Then txtPingAddDNS.Text = ""
   AffInfosOutils (3) ' Refresh état boutons
End Sub

Private Sub txtPingAddDNS_Change()
   If txtPingAddDNS.Text <> "" Then txtPingAddIP.Text = ""
   AffInfosOutils (3) ' Refresh état boutons
End Sub

Private Sub txtPingAddIP_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then Call cmdPingStartPing_Click
End Sub

Private Sub cmdPingIPToDNS_Click()
   Dim Msg As String
   Dim RetSts As Integer  ' Status de retour
   Dim RetAdd As String   ' Adresse de retour
   
   Screen.MousePointer = vbHourglass
   PingAffNouvMsg ("[" & txtPingAddIP.Text & "] => Recherche en cours...")
   RetSts = Net.IP_To_DNS(txtPingAddIP.Text, RetAdd)
      If (RetSts <> 0) Then
     Msg = "ERR : " & Net.ConvStatusMsg
    Else
          Msg = RetAdd
   End If
   PingRemplaceMsg ("[" & txtPingAddIP.Text & "] => " & Msg)
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdPingDNSToIP_Click()
   Dim Msg As String
   Dim RetSts As Integer  ' Status de retour
   Dim RetAdd As String   ' Adresse de retour

   Screen.MousePointer = vbHourglass
   PingAffNouvMsg ("[" & txtPingAddDNS.Text & "] => Recherche en cours...")
   RetSts = Net.DNS_To_IP(txtPingAddDNS.Text, RetAdd)
   If (RetSts <> 0) Then
     Msg = "ERR : " & Net.ConvStatusMsg
    Else
          Msg = RetAdd
   End If
   PingRemplaceMsg ("[" & txtPingAddDNS.Text & "] => " & Msg)
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdPingAutoPing_Click(Index As Integer)
   Dim Addr As String                   ' Champs Adresse
   Dim IP As String                     ' Adresse IP à convertir
   Dim EnTete As String                 ' 3 premières parties de l'adresse
   Dim Deb As Integer, Fin As Integer   ' Intervals de recherche
   Dim i As Integer                     ' Boucle
   Dim Pos1 As Integer, Pos2 As Integer ' Positions . et -
   Dim Msg As String                    ' Message final a afficher
   Dim RetSts1 As Integer               ' Status de retour
   Dim RetSts2 As Integer               ' Status de retour
   Dim RetAdd As String                 ' Adresse de retour

   Addr = txtPingAddIP
   Pos1 = 0: Pos2 = 0
   For i = 1 To 3
      Pos1 = InStr(Pos1 + 1, Addr, ".")
   Next
   If (Pos1 <> 0) Then
     EnTete = Left$(Addr, Pos1)
     Pos2 = InStr(Addr, "-")
     If (Pos2 <> 0) Then
       Deb = Val(Mid$(Addr, Pos1 + 1, Pos2 - Pos1 - 1))
       Fin = Val(Mid$(Addr, Pos2 + 1))
     End If
   End If
   
   If (EnTete = "" Or (Deb + Fin = 0)) Then
     If Index = 1 Then
       MsgBox ("Ping sur un interval." & vbCrLf & _
             "Entrez dans le champs IP une structure du type a.b.c.d-e" & vbCrLf & _
             "pour effectuer un ping des adresses de a.b.c.d jusque a.b.c.e")
      Else
           MsgBox ("Conversion IP -> NetBios-DNS sur un interval" & vbCrLf & _
                 "Entrez dans le champs IP une structure du type a.b.c.d-e" & vbCrLf & _
                 "pour effectuer une conversion des adresses de a.b.c.d jusque a.b.c.e")
     End If
     Exit Sub
   End If
   
   If (Net.ChekIP(EnTete & Trim$(Str$(Deb))) = False) Or _
      (Net.ChekIP(EnTete & Trim$(Str$(Fin))) = False) Then
     MsgBox ("Adresses ou intervals incorrects")
     Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   Select Case Index
    Case 1
        FlagDemandeArret = False
        For i = Deb To Fin
           IP = EnTete & Trim$(Str$(i))
           PingAffNouvMsg ("[" & IP & "] => Ping en cours...")
           RetSts1 = Net.DoPing(IP, ADD_IP, "ABC", Val(txtPingTimeOut.Text), _
                                             Val(txtPingTTL.Text))  ' IP
           If RetSts1 = 1 Then
             Msg = "Ping OK"
            Else
                  Msg = "<<< Ping NOK >>>"
           End If
           PingRemplaceMsg ("[" & IP & "] => " & Msg)
           If FlagDemandeArret = True Then FlagDemandeArret = False: Exit For
        Next
    Case 2
        PingAffNouvMsg ("====== Début conversion IP -> DNS ======")
        FlagDemandeArret = False
        For i = Deb To Fin
           IP = EnTete & Trim$(Str$(i))
           PingAffNouvMsg ("[" & IP & "] => Recherche en cours...")
           RetSts1 = Net.IP_To_DNS(IP, RetAdd)
           If (RetSts1 <> 0) Then
             Msg = "---"
            Else
                 Msg = RetAdd
           End If
      
           PingRemplaceMsg ("[" & IP & "] => " & Msg & "  -  Ping en cours...")
      
           RetSts2 = Net.DoPing(IP, ADD_IP, "ABC", Val(txtPingTimeOut.Text), _
                                             Val(txtPingTTL.Text))  ' IP
      
           Select Case ((RetSts1 = 0) * 10 + (RetSts2 = 1))
              Case 0
                  Msg = Msg & "      Ping NOK"
              Case -1
                  Msg = Msg & "      <<< Ping OK >>>"
              Case -10
                  Msg = Msg & "      <<< Ping NOK >>>"
              Case -11
                  Msg = Msg & "      Ping OK"
           End Select
      
           PingRemplaceMsg ("[" & IP & "] => " & Msg)
           If FlagDemandeArret = True Then FlagDemandeArret = False: Exit For
        Next
        PingAffNouvMsg ("====== Fin conversion IP -> DNS ======")
   End Select

   Screen.MousePointer = vbDefault
End Sub

Private Sub tmrBouclagePing_Timer()
   ' Ping en boucle. Désactiver timer si demande arrêt ou si plus dans l'onglet 'Ping'.
   If ((SSTab1.Tab <> 2) Or _
       (optMenuChoixOutils(0).Value = False) Or _
       (optChoixOutils(0).Value = False) Or _
       (FlagDemandeArret = True)) Then
     FlagDemandeArret = False
     chkPingRepeatPing.Value = 0
     tmrBouclagePing.Enabled = False
     Exit Sub
   End If
   
   ' Lancer que si l'option est toujours activée
   If chkPingRepeatPing.Value = 1 Then
     Call cmdPingStartPing_Click
    Else
         tmrBouclagePing.Enabled = False
   End If
End Sub

Private Sub cmdPingLocalIP_Click()
   Dim Pos1 As Integer
   Dim Txt As String
   Dim i As Integer
   
   Txt = Net.GetLocalIPInt
   Pos1 = 0
   For i = 1 To 3
      Pos1 = InStr(Pos1 + 1, Txt, ".")
   Next
   
   With txtPingAddIP
       .SetFocus
       .Text = Txt
       .SelStart = IIf(Pos1 <> 0, Pos1, 0)
       .SelLength = 100
   End With
End Sub

Private Sub PingAffNouvMsg(Msg As String)
   'Screen.MousePointer = vbHourglass
   lstPingResultats.AddItem Msg
   lstPingResultats.ListIndex = lstPingResultats.ListCount - 1
   DoEvents
End Sub

Private Sub PingRemplaceMsg(Msg As String)
   lstPingResultats.RemoveItem lstPingResultats.ListCount - 1
   lstPingResultats.AddItem Msg
   lstPingResultats.ListIndex = lstPingResultats.ListCount - 1
   'Screen.MousePointer = vbDefault
End Sub

Private Sub cmdPingClear_Click()
   lstPingResultats.Clear
End Sub

Private Sub lstPingResultats_dblClick()
   Dim T As String
   Dim P1 As Integer
   Dim P2 As Integer
   With lstPingResultats
       T = .List(.ListIndex)
   End With
   P1 = InStr(T, "[")
   P2 = InStr(T, "]")
   If ((P1 + P2) <> 0) Then T = Mid$(T, P1 + 1, P2 - P1 - 1)
   txtPingAddIP.Text = T
End Sub

'================================ MENU MAC @ ===============================

Private Sub cmdGetMACAddress_Click()
   Dim MACAdd(0 To 19) As String
   Dim i As Integer
   
   Net.MAC_DoSearch
   
   lstInfosReseau.Clear
   For i = 0 To 19
      If Net.MAC_Address(i) <> "" Then
        lstInfosReseau.AddItem "Réseau " & i & " : " & Net.MAC_Address(i)
      End If
   Next
End Sub

'============================ MENU Internet ==============================

Private Sub cmdHeureAtomiqueONOFF_Click()
   With cmdHeureAtomiqueONOFF
       Select Case .Caption
          Case "OFF"
              .Caption = "ON"
              .BackColor = vbGreen
              tmrHeureAtomique.Enabled = True
          Case Else
              .Caption = "OFF"
              .BackColor = vbButtonFace
              tmrHeureAtomique.Enabled = False
       End Select
   End With
End Sub


'============================ MENU Net Message ==============================

Private Sub cmdNetMsgEnvoi_Click(Index As Integer)
  Dim Msg As String
  Dim Ret As Long
  
  Select Case (chkNetMsgSignatureActive.Value = 1) * 10 + (optNetMsgTypeSignature(0).Value = True)
     Case -10
         Msg = txtNetMsgMsg.Text & vbCrLf & vbCrLf & ParamNetMsg.Signature
     Case -11
         Msg = "Message de " & ParamNetMsg.Signature & " : " & _
                vbCrLf & vbCrLf & txtNetMsgMsg.Text
     Case Else
         Msg = txtNetMsgMsg.Text
  End Select
   
   
  lblNetMsgStatus.Caption = "Envoi..."
  Screen.MousePointer = vbHourglass
  DoEvents
  
    Select Case Index
       Case 0
           Ret = Net.EnvoiMessage(SysInfos.Net_ComputerName, lblNetMsgDestName.Caption, Msg)
       Case 1
           Ret = Net.EnvoiMessage(SysInfos.Net_ComputerName, SysInfos.Net_ComputerName, Msg)
    End Select
    
  lblNetMsgStatus.Caption = IIf(Ret = 0, "-- OK --", "===== Echec =====")
  Screen.MousePointer = vbDefault

End Sub

Private Sub cmdNetMsgAjoutSupp_Click(Index As Integer)
   Select Case Index
      Case 0                    ' Ajouter
          Dim NewDNS As String
          Dim NewComment As String
          NewDNS = InputBox("Entrez nom machine (DNS) :")
          If (NewDNS <> "") Then
            NewComment = InputBox("Entrez commentaires :")
            If (NewComment <> "") Then
              ReDim Preserve ParamNetMsg.Dest(UBound(ParamNetMsg.Dest) + 1)
              ParamNetMsg.Dest(UBound(ParamNetMsg.Dest)).DNS = NewDNS
              ParamNetMsg.Dest(UBound(ParamNetMsg.Dest)).Comment = NewComment
              Call EcritureParamNetMsg
              AffInfosOutils (5) ' Force refresh liste
            End If
          End If
      Case 1                     ' Supprimer
          Dim i As Integer
          For i = 1 To UBound(ParamNetMsg.Dest)
             If (ParamNetMsg.Dest(i).Comment = lstNetMsgDestinataire.Text) Then
               i = i + 1000
             End If
          Next
          If (i > 1000) Then   ' Trouvé
            For i = i - 1001 To UBound(ParamNetMsg.Dest) - 1
               ParamNetMsg.Dest(i).DNS = ParamNetMsg.Dest(i + 1).DNS
               ParamNetMsg.Dest(i).Comment = ParamNetMsg.Dest(i + 1).Comment
            Next
            ReDim Preserve ParamNetMsg.Dest(UBound(ParamNetMsg.Dest) - 1)
            Call EcritureParamNetMsg: AffInfosOutils (5)  ' Force refresh liste
          End If
   End Select
End Sub

Private Sub cmdNetMsgDefSignature_Click()
   Dim Ret As String
   Ret = InputBox("Entrez la signature :")
   If (Ret <> "") Then
     ParamNetMsg.Signature = Ret
     Call EcritureParamNetMsg: AffInfosOutils (5)  ' Force refresh liste
   End If
End Sub

Private Sub optNetMsgTypeSignature_Click(Index As Integer)
   ParamNetMsg.TypeDest = Index
   Call EcritureParamNetMsg: AffInfosOutils (5)  ' Force refresh liste
End Sub

Private Sub chkNetMsgSignatureActive_Click()
   ParamNetMsg.SignatureActive = chkNetMsgSignatureActive.Value
   Call EcritureParamNetMsg: AffInfosOutils (5)  ' Force refresh liste
End Sub

Private Sub lstNetMsgDestinataire_Click()
   lblNetMsgDestName.Caption = NetMsg_FindDNS(lstNetMsgDestinataire.Text)
   lblNetMsgStatus = ""
End Sub

Private Sub txtNetMsgMsg_Change()
   lblNetMsgStatus = ""
End Sub

'=============================== MENU OPTIONS ===============================

Private Sub cmdSauveAutoDemarrage_Click()
  ' 0 = Supprimer les 2 raccourcis
  ' 1 = Utilisateur actuel seulement
  ' 2 = Tous les utilisateurs seulement
  If optAutoDemarrage(0).Value = True Then CreerRaccourci (0)
  If optAutoDemarrage(1).Value = True Then CreerRaccourci (1)
  If optAutoDemarrage(2).Value = True Then CreerRaccourci (2)
  cmdSauveAutoDemarrage.Enabled = False
End Sub

Private Sub optAutoDemarrage_Click(Index As Integer)
   cmdSauveAutoDemarrage.Enabled = True
End Sub

'============================================================================
'=========================== ONGLET SYSTEM INFOS ============================
'============================================================================

Private Sub tmrDureeLogin_Timer()
   If (SSTab1.Tab = 3) Then    ' Onglet System Infos
     lblDureeLogin.Caption = ConvSecToDHMS(SysInfos.OS_LogginDuration)
     lblMemoire(0).Caption = Format$(SysInfos.Mem_PhysAvl \ 1024, "###,###,###,### K")
   End If
End Sub

Private Sub cmdMSSysInfos_Click()
   Call SysInfos.StartMSSysInfo
End Sub

'============================================================================
'=========================== PROCEDURES DIVERSES ============================
'============================================================================

Sub AfficheInfosSystème()
   Dim Msg As String
   Dim lTemp As Long
   Dim szTemp As String
      
   ' ----- Message OS
   Select Case SysInfos.OS_PlatformID
      Case Windows32S
          Msg = "Windows 32S"
      Case Windows95
          Msg = "Windows 95"
      Case WindowsNT
          Msg = "Windows NT"
      Case Else
          Msg = "<OS inconnu>"
   End Select
   
   Msg = Msg & "  " & SysInfos.OS_MajorVersion & "." & SysInfos.OS_MinorVersion & _
               "   (Build " & SysInfos.OS_BuildNumber & ")   " & _
               SysInfos.OS_CSDVersion
   
   lblOS.Caption = Msg
   
   ' ----- Message Processeur
   Select Case SysInfos.Sys_ProcessorType
      Case PROCESSOR_INTEL_386
          Msg = "Intel 386"
      Case PROCESSOR_INTEL_486
          Msg = "Intel 486"
      Case PROCESSOR_INTEL_PENTIUM
          Msg = "Intel Pentium"
      Case PROCESSOR_MIPS_R4000
          Msg = "MIPS R4000"
      Case PROCESSOR_ALPHA_21064
          Msg = "DEC Alpha 21064"
      Case Else
          Msg = "<Inconnu>"
   End Select
      
   lTemp = SysInfos.Sys_Processor1Speed
   Msg = Msg & "   " & IIf(lTemp <> 0, lTemp, "??") & "  MHz"
   lTemp = SysInfos.Sys_Processor2Speed
   Msg = Msg & IIf(lTemp <> 0, "  -  " & lTemp & "  MHz", "")
      
   Msg = Msg & "    /    Actif(s) : "
   Msg = Msg & IIf(SysInfos.Sys_ActiveProcessorMask And 1, " 1 -", "")
   Msg = Msg & IIf(SysInfos.Sys_ActiveProcessorMask And 2, " 2 -", "")
   Msg = Msg & IIf(SysInfos.Sys_ActiveProcessorMask And 4, " 3 -", "")
   Msg = Msg & IIf(SysInfos.Sys_ActiveProcessorMask And 8, " 4 -", "")
   Msg = Left$(Msg, Len(Msg) - 1) & " sur " & SysInfos.Sys_NumberOfProcessor
   
   lblProcesseur.Caption = Msg

   ' ----- Message Mode StartUp

   Select Case SysInfos.OS_StartUpMode
      Case START_MODE_SAFE
          Msg = "Protégé"
      Case START_MODE_SAFENET
          Msg = "Protégé avec support réseau"
      Case Else
          Msg = "Normal"
   End Select
   
   lblStartUpMode.Caption = Msg

   ' ----- Message Affichage

   lblInfosAff.Caption = SysInfos.Scr_Width & " x " & SysInfos.Scr_Height & _
                         "    " & SysInfos.Scr_Bits & " bits  (" & _
                         SysInfos.Scr_NbColors & " couleurs)"


   ' Mémoire
   
   lblMemoire(0).Caption = Format$(SysInfos.Mem_PhysAvl \ 1024, "###,###,###,### K")
   lblMemoire(1).Caption = Format$(SysInfos.Mem_PhysTot \ 1024, "###,###,###,### K")
   lblMemoire(2).Caption = Format$(SysInfos.Mem_VirtAvl \ 1024, "###,###,###,### K")
   lblMemoire(3).Caption = Format$(SysInfos.Mem_VirtTot \ 1024, "###,###,###,### K")

   ' Réseau
   
   lblUserName.Caption = SysInfos.Net_UserName
   lblComputerName.Caption = SysInfos.Net_ComputerName
   lblAdresseIPInt = Net.GetLocalIPInt
   szTemp = Net.GetLocalIPExt
     lblAdresseIPExt = IIf(szTemp <> "", szTemp, "--")
End Sub

Sub ListeLecteurs()
   Dim i As Integer
   Dim Disque As String
   Dim DriveType As ListeTypesLecteurs
   Dim VolumeName As String
   Dim SerialNumber As Long
   Dim SystemName As String
   Dim SN As String
  
   lstLecteurs.Clear
  
   For i = 1 To 26
      Disque = Chr$(64 + i) & ":"
   
      DriveType = SysInfos.GetDriveType(Disque)
   
      VolumeName = String$(100, 0)
      SystemName = String$(100, 0)
      SN = ""
      If ((i <> 1) And (DriveType <> 99)) Then
        Call SysInfos.GetVolumeInfos(Disque & "\", VolumeName, SerialNumber, SystemName)
        If SerialNumber <> 0 Then
          SN = Hex$(SerialNumber)
          SN = Left$(SN, 4) & "-" & Mid$(SN, 5)
        End If
      End If
      lstLecteurs.AddItem Disque & " " & _
                          Left$(VolumeName & String$(20, " "), 11) & " " & _
                          Choose(DriveType + 1, "??????", "  --  ", "Mobile", " Fixe ", "Réseau", "CD-Rom", "RAM") & "  " & _
                          Format$(SystemName, " !@@@@@ ") & SN
   Next
End Sub

Sub AffInfosRegistres(Index As Integer)
   Dim i As Integer
   Dim iTemp As Integer
   Dim ErrFound As Integer
   Dim szTemp As String
   
   FlagRefreshEnCours = 1
   Select Case Index
   
   Case 0    '======== AUTO LOGIN ========
       chkAutoLogin.BackColor = vbButtonFace
         chkAutoLogin.Enabled = True
       cmdRegSave(IDX_ALOGIN).BackColor = vbButtonFace
         chkAutoLogin.Enabled = True
       
       ErrFound = 0
       Select Case RegALActif
          Case 0:                chkAutoLogin.Value = 0
          Case 1:                chkAutoLogin.Value = 1
          Case REGSTS_INOKEY:    chkAutoLogin.BackColor = COL_ROUGE: chkAutoLogin.Enabled = False: ErrFound = 1
          Case Else:             chkAutoLogin.BackColor = COL_ORANGE: chkAutoLogin.Enabled = False: ErrFound = 1
       End Select

       txtALUsername.Text = RegALUserName
       txtALDomain.Text = RegALDomain
       txtALPassword.Text = RegALPassword
       txtALVerifPassword.Text = txtALPassword.Text
   
       If ((txtALUsername.Text = REGSTS_SZNOKEY) Or (txtALDomain.Text = REGSTS_SZNOKEY) Or _
           (txtALPassword.Text = REGSTS_SZNOKEY)) Then ErrFound = 1
   
       If (SysInfos.OS_PlatformID = WindowsNT) Then
         lblALNonDisponible.Visible = False
        Else
          chkAutoLogin.Enabled = False
       End If

       Call chkAutoLogin_Click
   
       If (ErrFound <> 0) Then
         cmdRegSave(IDX_ALOGIN).BackColor = COL_ROUGE
         cmdRegSave(IDX_ALOGIN).Enabled = False
         chkAutoLogin.Enabled = False
       End If
       cmdRegSave(IDX_ALOGIN).Enabled = False
   
   Case 1      '======== LOGIN ========
       ErrFound = 0
       txtPopStartUpTitre.Text = RegPopStartUpTitre
       txtPopStartUpContenu.Text = RegPopStartUpText
       cmdRegSave(IDX_MSGLOG).BackColor = vbButtonFace
         cmdRegSave(IDX_MSGLOG).Enabled = True
       If ((txtPopStartUpTitre.Text = REGSTS_SZNOKEY) Or (txtPopStartUpContenu.Text = REGSTS_SZNOKEY)) Then
         cmdRegSave(IDX_MSGLOG).BackColor = COL_ROUGE
         cmdRegSave(IDX_MSGLOG).Enabled = False
         txtPopStartUpTitre.Enabled = False
         txtPopStartUpContenu.Enabled = False
       End If
       cmdRegSave(IDX_MSGLOG).Enabled = False
            
       txtMsgLogon.Text = RegLogonPrompt
       cmdRegSave(IDX_POPLOG).BackColor = vbButtonFace
         cmdRegSave(IDX_POPLOG).Enabled = True
       If (txtMsgLogon.Text = REGSTS_SZNOKEY) Then
         cmdRegSave(IDX_POPLOG).BackColor = COL_ROUGE
         cmdRegSave(IDX_POPLOG).Enabled = False
         txtMsgLogon.Enabled = False
       End If
       cmdRegSave(IDX_POPLOG).Enabled = False

   Case 2
       '======== INH BUREAU ========
       chkOptReg(IDX_INHDSK).BackColor = vbButtonFace
       chkOptReg(IDX_INHDSK).Enabled = True
       Select Case RegInhDesktop
          Case 0:                chkOptReg(IDX_INHDSK).Value = 0
          Case 1:                chkOptReg(IDX_INHDSK).Value = 1
          Case REGSTS_INOKEY:    chkOptReg(IDX_INHDSK).Enabled = False
                                 chkOptReg(IDX_INHDSK).BackColor = COL_ROUGE
          Case Else:             chkOptReg(IDX_INHDSK).Enabled = False
                                 chkOptReg(IDX_INHDSK).BackColor = COL_ORANGE
       End Select

       '======== INH RESEAU ========
       chkOptReg(IDX_ICONET).BackColor = vbButtonFace
       chkOptReg(IDX_ICONET).Enabled = True
       Select Case RegInhIconeReseau
          Case 0:                chkOptReg(IDX_ICONET).Value = 0
          Case 1:                chkOptReg(IDX_ICONET).Value = 1
          Case REGSTS_INOKEY:    chkOptReg(IDX_ICONET).Enabled = False
                                 chkOptReg(IDX_ICONET).BackColor = COL_ROUGE
          Case Else:             chkOptReg(IDX_ICONET).Enabled = False
                                 chkOptReg(IDX_ICONET).BackColor = COL_ORANGE
       End Select
         
       '======== INH RESEAU GLOBAL ========
       chkOptReg(IDX_ENTNET).BackColor = vbButtonFace
       chkOptReg(IDX_ENTNET).Enabled = True
       Select Case RegInhReseauGlobal
          Case 0:                chkOptReg(IDX_ENTNET).Value = 0
          Case 1:                chkOptReg(IDX_ENTNET).Value = 1
          Case REGSTS_INOKEY:    chkOptReg(IDX_ENTNET).Enabled = False
                                 chkOptReg(IDX_ENTNET).BackColor = COL_ROUGE
          Case Else:             chkOptReg(IDX_ENTNET).Enabled = False
                                 chkOptReg(IDX_ENTNET).BackColor = COL_ORANGE
       End Select
   
       '======== INH RESEAU : CONTENU GROUPE TRAVAIL ========
       chkOptReg(IDX_NETWKC).BackColor = vbButtonFace
       chkOptReg(IDX_NETWKC).Enabled = True
       Select Case RegInhContenuGroupesTravail
          Case 0:                chkOptReg(IDX_NETWKC).Value = 0
          Case 1:                chkOptReg(IDX_NETWKC).Value = 1
          Case REGSTS_INOKEY:    chkOptReg(IDX_NETWKC).Enabled = False
                                 chkOptReg(IDX_NETWKC).BackColor = COL_ROUGE
          Case Else:             chkOptReg(IDX_NETWKC).Enabled = False
                                 chkOptReg(IDX_NETWKC).BackColor = COL_ORANGE
       End Select
      
       '======== INH IE ========
       chkOptReg(IDX_ICOWEB).BackColor = vbButtonFace
       chkOptReg(IDX_ICOWEB).Enabled = True
       Select Case RegInhIconeIE
          Case 0:                chkOptReg(IDX_ICOWEB).Value = 0
          Case 1:                chkOptReg(IDX_ICOWEB).Value = 1
          Case REGSTS_INOKEY:    chkOptReg(IDX_ICOWEB).Enabled = False
                                 chkOptReg(IDX_ICOWEB).BackColor = COL_ROUGE
          Case Else:             chkOptReg(IDX_ICOWEB).Enabled = False
                                 chkOptReg(IDX_ICOWEB).BackColor = COL_ORANGE
       End Select
      
       '======== INH CLIC DROIT SOURIS ========
       chkOptReg(IDX_CLKDRT).BackColor = vbButtonFace
       chkOptReg(IDX_CLKDRT).Enabled = True
       Select Case RegInhClicDroit
          Case 0:                chkOptReg(IDX_CLKDRT).Value = 0
          Case 1:                chkOptReg(IDX_CLKDRT).Value = 1
          Case REGSTS_INOKEY:    chkOptReg(IDX_CLKDRT).Enabled = False
                                 chkOptReg(IDX_CLKDRT).BackColor = COL_ROUGE
          Case Else:             chkOptReg(IDX_CLKDRT).Enabled = False
                                 chkOptReg(IDX_CLKDRT).BackColor = COL_ORANGE
       End Select
       
       '======== TAILLE ICONES BUREAU ========
       txtTailleIconesBureau.BackColor = vbWindowBackground
       txtTailleIconesBureau.Enabled = True
       iTemp = RegTailleIconesBureau
       Select Case iTemp
          Case REGSTS_INOKEY:    txtTailleIconesBureau.Enabled = False
                                 txtTailleIconesBureau.BackColor = COL_ROUGE
          Case REGSTS_IINVVAL:   txtTailleIconesBureau.Enabled = False
                                 txtTailleIconesBureau.BackColor = COL_ORANGE
          Case REGSTS_INOVALUE:  txtTailleIconesBureau.Enabled = False
          Case Else:             txtTailleIconesBureau.Text = iTemp
       End Select
       cmdRegSave(IDX_SIZICD).Enabled = False
       
       '======== TAILLE ICONES MENU START ========
       txtTailleIconesMenuStart.BackColor = vbWindowBackground
       txtTailleIconesMenuStart.Enabled = True
       iTemp = RegTailleIconesMenuStart
       Select Case iTemp
          Case REGSTS_INOKEY:    txtTailleIconesMenuStart.Enabled = False
                                 txtTailleIconesMenuStart.BackColor = COL_ROUGE
          Case REGSTS_IINVVAL:   txtTailleIconesMenuStart.Enabled = False
                                 txtTailleIconesMenuStart.BackColor = COL_ORANGE
          Case REGSTS_INOVALUE:  txtTailleIconesMenuStart.Enabled = False
          Case Else:             txtTailleIconesMenuStart.Text = iTemp
       End Select
       cmdRegSave(IDX_SIZICS).Enabled = False
       
       '======== DELAIS ========
       txtDelaiMenus.BackColor = vbWindowBackground
       txtDelaiMenus.Enabled = True
       iTemp = RegDelaiMenus
       Select Case iTemp
          Case REGSTS_INOKEY:    txtDelaiMenus.Enabled = False
                                 txtDelaiMenus.BackColor = COL_ROUGE
          Case REGSTS_IINVVAL:   txtDelaiMenus.Enabled = False
                                 txtDelaiMenus.BackColor = COL_ORANGE
          Case REGSTS_INOVALUE:  txtDelaiMenus.Enabled = False
          Case Else:             txtDelaiMenus.Text = iTemp
       End Select
       cmdRegSave(IDX_TMRMNU).Enabled = False
       
   Case 3
       '======== INH ARRET WINDOWS ========
       chkOptReg(IDX_STPWIN).BackColor = vbButtonFace
       chkOptReg(IDX_STPWIN).Enabled = True
       Select Case RegInhArretWindows
          Case 0:                chkOptReg(IDX_STPWIN).Value = 0
          Case 1:                chkOptReg(IDX_STPWIN).Value = 1
          Case REGSTS_INOKEY:    chkOptReg(IDX_STPWIN).Enabled = False
                                 chkOptReg(IDX_STPWIN).BackColor = COL_ROUGE
          Case Else:             chkOptReg(IDX_STPWIN).Enabled = False
                                 chkOptReg(IDX_STPWIN).BackColor = COL_ORANGE
       End Select
      
       '======== INH LOG OFF ========
       chkOptReg(IDX_LOGOFF).BackColor = vbButtonFace
       chkOptReg(IDX_LOGOFF).Enabled = True
       Select Case RegInhLogOff
          Case 0:                chkOptReg(IDX_LOGOFF).Value = 0
          Case 1:                chkOptReg(IDX_LOGOFF).Value = 1
          Case REGSTS_INOKEY:    chkOptReg(IDX_LOGOFF).Enabled = False
                                 chkOptReg(IDX_LOGOFF).BackColor = COL_ROUGE
          Case Else:             chkOptReg(IDX_LOGOFF).Enabled = False
                                 chkOptReg(IDX_LOGOFF).BackColor = COL_ORANGE
       End Select
      
       '======== INH SECTION PROGS COMMUNS ========
       chkOptReg(IDX_PRGCMN).BackColor = vbButtonFace
       chkOptReg(IDX_PRGCMN).Enabled = True
       Select Case RegInhPrgsCommuns
          Case 0:                chkOptReg(IDX_PRGCMN).Value = 0
          Case 1:                chkOptReg(IDX_PRGCMN).Value = 1
          Case REGSTS_INOKEY:    chkOptReg(IDX_PRGCMN).Enabled = False
                                 chkOptReg(IDX_PRGCMN).BackColor = COL_ROUGE
          Case Else:             chkOptReg(IDX_PRGCMN).Enabled = False
                                 chkOptReg(IDX_PRGCMN).BackColor = COL_ORANGE
       End Select
      
       '======== INH COMMANDE RUN ========
       chkOptReg(IDX_CMDRUN).BackColor = vbButtonFace
       chkOptReg(IDX_CMDRUN).Enabled = True
       Select Case RegInhCmdRun
          Case 0:                chkOptReg(IDX_CMDRUN).Value = 0
          Case 1:                chkOptReg(IDX_CMDRUN).Value = 1
          Case REGSTS_INOKEY:    chkOptReg(IDX_CMDRUN).Enabled = False
                                 chkOptReg(IDX_CMDRUN).BackColor = COL_ROUGE
          Case Else:             chkOptReg(IDX_CMDRUN).Enabled = False
                                 chkOptReg(IDX_CMDRUN).BackColor = COL_ORANGE
       End Select
      
       '======== INH COMMANDE FIND ========
       chkOptReg(IDX_CMDFND).BackColor = vbButtonFace
       chkOptReg(IDX_CMDFND).Enabled = True
       Select Case RegInhCmdFind
          Case 0:                chkOptReg(IDX_CMDFND).Value = 0
          Case 1:                chkOptReg(IDX_CMDFND).Value = 1
          Case REGSTS_INOKEY:    chkOptReg(IDX_CMDFND).Enabled = False
                                 chkOptReg(IDX_CMDFND).BackColor = COL_ROUGE
          Case Else:             chkOptReg(IDX_CMDFND).Enabled = False
                                 chkOptReg(IDX_CMDFND).BackColor = COL_ORANGE
       End Select
       
       '======== INH CONFIG GENERALE ========
       chkOptReg(IDX_CNFGEN).BackColor = vbButtonFace
       chkOptReg(IDX_CNFGEN).Enabled = True
       Select Case RegInhCnfGen
          Case 0:                chkOptReg(IDX_CNFGEN).Value = 0
          Case 1:                chkOptReg(IDX_CNFGEN).Value = 1
          Case REGSTS_INOKEY:    chkOptReg(IDX_CNFGEN).Enabled = False
                                 chkOptReg(IDX_CNFGEN).BackColor = COL_ROUGE
          Case Else:             chkOptReg(IDX_CNFGEN).Enabled = False
                                 chkOptReg(IDX_CNFGEN).BackColor = COL_ORANGE
       End Select
      
       '======== INH CONFIG TASK BAR ========
       chkOptReg(IDX_CNFTKB).BackColor = vbButtonFace
       chkOptReg(IDX_CNFTKB).Enabled = True
       Select Case RegInhCnfTaskBar
          Case 0:                chkOptReg(IDX_CNFTKB).Value = 0
          Case 1:                chkOptReg(IDX_CNFTKB).Value = 1
          Case REGSTS_INOKEY:    chkOptReg(IDX_CNFTKB).Enabled = False
                                 chkOptReg(IDX_CNFTKB).BackColor = COL_ROUGE
          Case Else:             chkOptReg(IDX_CNFTKB).Enabled = False
                                 chkOptReg(IDX_CNFTKB).BackColor = COL_ORANGE
       End Select
       
       '======== INH LOCK WORKSTATION ========
       chkOptReg(IDX_LOCKST).BackColor = vbButtonFace
       chkOptReg(IDX_LOCKST).Enabled = True
       Select Case RegInhLockStation
          Case 0:                chkOptReg(IDX_LOCKST).Value = 0
          Case 1:                chkOptReg(IDX_LOCKST).Value = 1
          Case Else:             chkOptReg(IDX_LOCKST).Enabled = False
                                 chkOptReg(IDX_LOCKST).BackColor = COL_ORANGE
       End Select
    
       '======== INH TASK MANAGER ========
       chkOptReg(IDX_TSKMGR).BackColor = vbButtonFace
       chkOptReg(IDX_TSKMGR).Enabled = True
       Select Case RegInhTaskManager
          Case 0:                chkOptReg(IDX_TSKMGR).Value = 0
          Case 1:                chkOptReg(IDX_TSKMGR).Value = 1
          Case Else:             chkOptReg(IDX_TSKMGR).Enabled = False
                                 chkOptReg(IDX_TSKMGR).BackColor = COL_ORANGE
       End Select
   
       '======== INH CHANGE PASSWORD ========
       chkOptReg(IDX_CHGPWD).BackColor = vbButtonFace
       chkOptReg(IDX_CHGPWD).Enabled = True
       Select Case RegInhChangePassword
          Case 0:                chkOptReg(IDX_CHGPWD).Value = 0
          Case 1:                chkOptReg(IDX_CHGPWD).Value = 1
          Case Else:             chkOptReg(IDX_CHGPWD).Enabled = False
                                 chkOptReg(IDX_CHGPWD).BackColor = COL_ORANGE
       End Select
   Case 4
       '======== CACHE LECTEURS ========
       Call ListeLecteurs
   
       szTemp = RegCacheLecteurs
       cmdRegSave(IDX_HIDDRV).BackColor = vbButtonFace
       cmdRegSave(IDX_HIDDRV).Enabled = True
       Select Case szTemp
          Case REGSTS_SZNOKEY:   cmdRegSave(IDX_HIDDRV).Enabled = False
                                 'lstLecteurs.Enabled = False
                                 cmdRegSave(IDX_HIDDRV).BackColor = COL_ROUGE
          Case Else
              For i = 0 To 25
                 If (Mid$(szTemp, i + 1, 1) = "1") Then lstLecteurs.Selected(i) = True
              Next
       End Select
       lstLecteurs.ListIndex = 0
       cmdRegSave(IDX_HIDDRV).Enabled = False
   Case 5
       '======== AFFICHAGE ========
       'ErrFound = 0
       chkOptReg(IDX_CPAGEN).BackColor = vbButtonFace
       chkOptReg(IDX_CPAGEN).Enabled = True
       Select Case RegInhCPLAffAccesCPL
          Case 0:                chkOptReg(IDX_CPAGEN).Value = 0
          Case 1:                chkOptReg(IDX_CPAGEN).Value = 1
          Case Else:             chkOptReg(IDX_CPAGEN).BackColor = COL_ORANGE
                                 chkOptReg(IDX_CPAGEN).Enabled = False
       End Select
       
       chkOptReg(IDX_CPAAPP).BackColor = vbButtonFace
       chkOptReg(IDX_CPAAPP).Enabled = True
       Select Case RegInhCPLAffOgtApp
          Case 0:                chkOptReg(IDX_CPAAPP).Value = 0
          Case 1:                chkOptReg(IDX_CPAAPP).Value = 1
          Case Else:             chkOptReg(IDX_CPAAPP).BackColor = COL_ORANGE
                                 chkOptReg(IDX_CPAAPP).Enabled = False
       End Select
       
      chkOptReg(IDX_CPABKG).BackColor = vbButtonFace
      chkOptReg(IDX_CPABKG).Enabled = True
      Select Case RegInhCPLAffOgtAP
          Case 0:                chkOptReg(IDX_CPABKG).Value = 0
          Case 1:                chkOptReg(IDX_CPABKG).Value = 1
          Case Else:             chkOptReg(IDX_CPABKG).BackColor = COL_ORANGE
                                 chkOptReg(IDX_CPABKG).Enabled = False
       End Select
       
       chkOptReg(IDX_CPASCR).BackColor = vbButtonFace
       chkOptReg(IDX_CPASCR).Enabled = True
       Select Case RegInhCPLAffOgtEco
          Case 0:                chkOptReg(IDX_CPASCR).Value = 0
          Case 1:                chkOptReg(IDX_CPASCR).Value = 1
          Case Else:             chkOptReg(IDX_CPASCR).BackColor = COL_ORANGE
                                 chkOptReg(IDX_CPASCR).Enabled = False
       End Select
       
       chkOptReg(IDX_CPACNF).BackColor = vbButtonFace
       chkOptReg(IDX_CPACNF).Enabled = True
       Select Case RegInhCPLAffOgtCnf
          Case 0:                chkOptReg(IDX_CPACNF).Value = 0
          Case 1:                chkOptReg(IDX_CPACNF).Value = 1
          Case Else:             chkOptReg(IDX_CPACNF).BackColor = COL_ORANGE
                                 chkOptReg(IDX_CPACNF).Enabled = False
       End Select
       
       '======== INH IMPRIMANTE AJOUT ========
       chkOptReg(IDX_IMPAJT).BackColor = vbButtonFace
       chkOptReg(IDX_IMPAJT).Enabled = True
       Select Case RegInhImpAjout
          Case 0:                chkOptReg(IDX_IMPAJT).Value = 0
          Case 1:                chkOptReg(IDX_IMPAJT).Value = 1
          Case REGSTS_INOKEY:    chkOptReg(IDX_IMPAJT).Enabled = False
                                 chkOptReg(IDX_IMPAJT).BackColor = COL_ROUGE
          Case Else:             chkOptReg(IDX_IMPAJT).Enabled = False
                                 chkOptReg(IDX_IMPAJT).BackColor = COL_ORANGE
       End Select
      
       '======== INH IMPRIMANTE SUPPRIMER ========
       chkOptReg(IDX_IMPSUP).BackColor = vbButtonFace
       chkOptReg(IDX_IMPSUP).Enabled = True
       Select Case RegInhImpSupp
          Case 0:                chkOptReg(IDX_IMPSUP).Value = 0
          Case 1:                chkOptReg(IDX_IMPSUP).Value = 1
          Case REGSTS_INOKEY:    chkOptReg(IDX_IMPSUP).Enabled = False
                                 chkOptReg(IDX_IMPSUP).BackColor = COL_ROUGE
          Case Else:             chkOptReg(IDX_IMPSUP).Enabled = False
                                 chkOptReg(IDX_IMPSUP).BackColor = COL_ORANGE
       End Select
      
   Case 7
      
       '======== Initial Keyboard Indicators ========
       chkOptReg(IDX_KEYB11).Enabled = True
       chkOptReg(IDX_KEYB12).Enabled = True
       chkOptReg(IDX_KEYB13).Enabled = True
       lblIKIDft.BackColor = vbButtonFace
       iTemp = RegInitialKeyboardIndicatorsDefault
       chkOptReg(IDX_KEYB11).Value = IIf((iTemp And 1) > 0, 1, 0)
       chkOptReg(IDX_KEYB12).Value = IIf((iTemp And 2) > 0, 1, 0)
       chkOptReg(IDX_KEYB13).Value = IIf((iTemp And 4) > 0, 1, 0)
       Select Case iTemp
          Case 0 To 7
          Case REGSTS_INOKEY:    lblIKIDft.BackColor = COL_ROUGE
                                 chkOptReg(IDX_KEYB11).Enabled = False
                                 chkOptReg(IDX_KEYB12).Enabled = False
                                 chkOptReg(IDX_KEYB13).Enabled = False
          Case REGSTS_IINVVAL:   lblIKIDft.BackColor = COL_ORANGE
                                 chkOptReg(IDX_KEYB11).Enabled = False
                                 chkOptReg(IDX_KEYB12).Enabled = False
                                 chkOptReg(IDX_KEYB13).Enabled = False
          Case Else:             chkOptReg(IDX_KEYB11).Enabled = False
                                 chkOptReg(IDX_KEYB12).Enabled = False
                                 chkOptReg(IDX_KEYB13).Enabled = False
       End Select
      
       chkOptReg(IDX_KEYB21).Enabled = True
       chkOptReg(IDX_KEYB22).Enabled = True
       chkOptReg(IDX_KEYB23).Enabled = True
       lblIKIUsr.BackColor = vbButtonFace
       iTemp = RegInitialKeyboardIndicatorsCurrentUser
       chkOptReg(IDX_KEYB21).Value = IIf((iTemp And 1) > 0, 1, 0)
       chkOptReg(IDX_KEYB22).Value = IIf((iTemp And 2) > 0, 1, 0)
       chkOptReg(IDX_KEYB23).Value = IIf((iTemp And 4) > 0, 1, 0)
       Select Case iTemp
          Case 0 To 7
          Case REGSTS_INOKEY:    lblIKIUsr.BackColor = COL_ROUGE
                                 chkOptReg(IDX_KEYB21).Enabled = False
                                 chkOptReg(IDX_KEYB22).Enabled = False
                                 chkOptReg(IDX_KEYB23).Enabled = False
          Case REGSTS_IINVVAL:   lblIKIUsr.BackColor = COL_ORANGE
                                 chkOptReg(IDX_KEYB21).Enabled = False
                                 chkOptReg(IDX_KEYB22).Enabled = False
                                 chkOptReg(IDX_KEYB23).Enabled = False
          Case Else:             chkOptReg(IDX_KEYB21).Enabled = False
                                 chkOptReg(IDX_KEYB22).Enabled = False
                                 chkOptReg(IDX_KEYB23).Enabled = False
       End Select
      
       '======== INH SHUTDOWN DANS LOGIN ========
       chkOptReg(IDX_ISDLOG).BackColor = vbButtonFace
       chkOptReg(IDX_ISDLOG).Enabled = True
       Select Case RegInhShutdownDansLogin
          Case 0:                chkOptReg(IDX_ISDLOG).Value = 0
          Case 1:                chkOptReg(IDX_ISDLOG).Value = 1
          Case REGSTS_INOKEY:    chkOptReg(IDX_ISDLOG).Enabled = False
                                 chkOptReg(IDX_ISDLOG).BackColor = COL_ROUGE
          Case Else:             chkOptReg(IDX_ISDLOG).Enabled = False
                                 chkOptReg(IDX_ISDLOG).BackColor = COL_ORANGE
       End Select
            
   End Select
   FlagRefreshEnCours = 0
End Sub

Sub AffInfosOutils(Index As Integer)
   Select Case Index
      Case 1
          optAutoDemarrage(0).Value = True
          If (Dir(ParamPrg.RaccourciActuel) <> "") Then optAutoDemarrage(1).Value = True
          If (Dir(ParamPrg.RaccourciAll) <> "") Then optAutoDemarrage(2).Value = True
          cmdSauveAutoDemarrage.Enabled = False
          
      Case 3  ' Ping
          cmdPingStartPing.Enabled = ((txtPingAddIP <> "") Or (txtPingAddDNS <> ""))
          cmdPingAutoPing(1).Enabled = (txtPingAddIP <> "")
          cmdPingAutoPing(2).Enabled = (txtPingAddIP <> "")
          cmdPingIPToDNS.Enabled = ((txtPingAddIP <> "") And (txtPingAddDNS = ""))
          cmdPingDNSToIP.Enabled = ((txtPingAddDNS <> "") And (txtPingAddIP = ""))
      Case 5  ' Net Message
         Call LectureParamNetMsg
         Dim i As Integer
         If UBound(ParamNetMsg.Dest) <> lstNetMsgDestinataire.ListCount Then
           lstNetMsgDestinataire.Clear
           For i = 1 To UBound(ParamNetMsg.Dest)
              lstNetMsgDestinataire.AddItem ParamNetMsg.Dest(i).Comment
           Next
         End If
         lblNetMsgSignature.Caption = ParamNetMsg.Signature
         Select Case ParamNetMsg.TypeDest
            Case 0, 1
                optNetMsgTypeSignature(ParamNetMsg.TypeDest).Value = True
            Case Else
         End Select
         Select Case ParamNetMsg.SignatureActive
            Case 0
                chkNetMsgSignatureActive.Value = 0
                optNetMsgTypeSignature(0).Enabled = False
                optNetMsgTypeSignature(1).Enabled = False
                lblNetMsgSignature.Enabled = False
                cmdNetMsgDefSignature.Enabled = False
            Case 1
                chkNetMsgSignatureActive.Value = 1
                optNetMsgTypeSignature(0).Enabled = True
                optNetMsgTypeSignature(1).Enabled = True
                lblNetMsgSignature.Enabled = True
                cmdNetMsgDefSignature.Enabled = True
            Case Else
         End Select
         lblNetMsgDestName.Caption = NetMsg_FindDNS(lstNetMsgDestinataire.Text)
         lblNetMsgStatus = ""
      Case Else
   End Select
End Sub

Function NetMsg_FindDNS(Comment As String)
   Dim i As Integer
   
   i = 0
   Do
        If ParamNetMsg.Dest(i).Comment = Comment Then
          NetMsg_FindDNS = ParamNetMsg.Dest(i).DNS
          i = 9999
        End If
        i = i + 1
   Loop While i <= UBound(ParamNetMsg.Dest)
End Function

