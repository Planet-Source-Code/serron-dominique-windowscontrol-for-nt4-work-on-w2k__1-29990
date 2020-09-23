Attribute VB_Name = "Mod_StartUpLink"
Option Explicit

'======================================================================
'=================== PLACE RACCOURCI DANS STARTUP =====================
'======================================================================

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" _
    (ByVal pidl As Long, ByVal pszPath As String) As Long

Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" _
    (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long

Private Declare Function fCreateShellLink Lib "VB5stkit.dll" _
    (ByVal lpstrFolderName As String, _
     ByVal lpstrLinkName As String, _
     ByVal lpstrLinkPath As String, _
     ByVal lpstrLinkArgs As String) As Long

'================== PLACE RACCOURCI DANS STARTUP ======================

Sub DetectionCheminsPrg()
   Dim Dossier As String
   
   Dossier = DossierSpecial(2)
   If Right$(Dossier, 1) = "\" Then Dossier = Left$(Dossier, Len(Dossier) - 1)
   ParamPrg.RepProgActuel = Dossier
   
   Dossier = DossierSpecial(23)
   If Right$(Dossier, 1) = "\" Then Dossier = Left$(Dossier, Len(Dossier) - 1)
   ParamPrg.RepProgAll = Dossier
   
   Dossier = DossierSpecial(7)
   If Right$(Dossier, 1) = "\" Then Dossier = Left$(Dossier, Len(Dossier) - 1)
   ParamPrg.RepStartActuel = Dossier
   
   Dossier = DossierSpecial(24)
   If Right$(Dossier, 1) = "\" Then Dossier = Left$(Dossier, Len(Dossier) - 1)
   ParamPrg.RepStartAll = Dossier
   
   ParamPrg.NomRaccourci = App.Title & ".lnk"
   
   ParamPrg.RaccourciActuel = ParamPrg.RepStartActuel & "\" & ParamPrg.NomRaccourci
   ParamPrg.RaccourciAll = ParamPrg.RepStartAll & "\" & ParamPrg.NomRaccourci
End Sub


'récupère un dossier spécial style c:\windows, c:\windows\recent...
Function DossierSpecial(ByVal CSIDL As Long) As String
   Dim Ret As Long
   Dim Path As String
   Dim IDL As ITEMIDLIST

'   Ret = SHGetSpecialFolderLocation(frmPrincipal.hWnd, CSIDL, IDL)
   Ret = SHGetSpecialFolderLocation(0, CSIDL, IDL)
   If Ret = 0 Then
     Path = Space$(260)
     Ret = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal Path)
     If Ret Then DossierSpecial = Left$(Path, InStr(Path, Chr$(0)) - 1)
   End If
End Function

Function CHEMIN_RELATIF(Chemin1 As String, CheminRef As String) As String
   'Renvoie l'adresse de Chemin1 (qui est un dossier) relativement
   'à CheminRef
   'Attention : Chemin1 et CheminRef doivent être donnés sans le \ final
   'Ex : Si Chemin1 = "c:\program files" et CheminRef = "c:\windows"
   'alors CHEMIN_RELATIF = "..\program files"

   Dim Morceau1() As String, Morceau2() As String
   Dim TailleMorceau1 As Integer, TailleMorceau2 As Integer
   Dim i As Integer, j As Integer
   Dim Prov As String
   
   ReDim Morceau1(1 To 1)
   ReDim Morceau2(1 To 1)
   
   '*** Marque chaque élément du chemin dans un tableau
   '* pour chemin1
   For i = 1 To Len(Chemin1)
       If Mid(Chemin1, i, 1) = "\" Then
           TailleMorceau1 = TailleMorceau1 + 1
           ReDim Preserve Morceau1(1 To TailleMorceau1)
           Morceau1(TailleMorceau1) = Prov
           Prov = ""
       Else
           Prov = Prov & Mid(Chemin1, i, 1)
       End If
   Next i
   
   'rajoute le dernier élément (non précédé d'un slash)
   TailleMorceau1 = TailleMorceau1 + 1
   ReDim Preserve Morceau1(1 To TailleMorceau1)
   Morceau1(TailleMorceau1) = Prov
   
   '* pour CheminRef
   For i = 1 To Len(CheminRef)
       If Mid(CheminRef, i, 1) = "\" Then
           TailleMorceau2 = TailleMorceau2 + 1
           ReDim Preserve Morceau2(1 To TailleMorceau2)
           Morceau2(TailleMorceau2) = Prov
           Prov = ""
       Else
           Prov = Prov & Mid(CheminRef, i, 1)
       End If
   Next i
   
   TailleMorceau2 = TailleMorceau2 + 1
   ReDim Preserve Morceau2(1 To TailleMorceau2)
   Morceau2(TailleMorceau2) = Prov
   
   Prov = ""
   For i = 1 To TailleMorceau2 - 1
       Prov = Prov & "..\"
   Next i
   
   For i = 2 To TailleMorceau1
       Prov = Prov & Morceau1(i) & "\"
   Next i
   
   CHEMIN_RELATIF = Left(Prov, Len(Prov) - 1)  'retire le "\" final
End Function

Sub CreerRaccourci(IDDestination As Integer)
   ' Paramètre : 0 = Effacer les 2 raccourcis
   '             1 = Créer pour utilisateur actuel
   '             2 = Créer pour tous les utilisateurs
   Dim Ret As Long
   Dim CheminApp As String
   Dim NomFichier As String
   Dim DossierRelatif As String
   
   CheminApp = App.Path
   If Right$(App.Path, 1) <> "\" Then CheminApp = CheminApp & "\"

   NomFichier = CheminApp & App.EXEName & ".exe"   'Nom du fichier exe
     
   'Supprime les deux raccourcis
   If (Dir(ParamPrg.RaccourciActuel) <> "") Then Kill ParamPrg.RaccourciActuel
   If (Dir(ParamPrg.RaccourciAll) <> "") Then Kill ParamPrg.RaccourciAll
      
   ' Création des raccourcis
   Select Case IDDestination
      Case 1:  DossierRelatif = CHEMIN_RELATIF(ParamPrg.RepStartActuel, ParamPrg.RepProgActuel)
               Ret = fCreateShellLink(DossierRelatif, App.Title, NomFichier, "")
      Case 2:  DossierRelatif = CHEMIN_RELATIF(ParamPrg.RepStartAll, ParamPrg.RepProgAll)
               Ret = fCreateShellLink(DossierRelatif, App.Title, NomFichier, "")
      Case Else
   End Select
End Sub

