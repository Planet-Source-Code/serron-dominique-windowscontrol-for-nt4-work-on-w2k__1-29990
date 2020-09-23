VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_ElementsBureau"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Class1"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'======================================================================
'===========================  ELEMENTS ================================
'======================================================================

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
        Private Const SWP_SHOWWINDOW = &H40
        Private Const SWP_HIDEWINDOW = &H80
        Private Const SWP_FRAMECHANGED = &H20
        
Enum ListeElements
    ELMT_BARRE_TACHE = 1
    ELMT_BOUTON_START = 2
    ELMT_PRGS_BARRE_TACHE = 3
    ELMT_TRAY = 4
    ELMT_HORLOGE = 5
End Enum

Private Type TypeHandleElements
    BarreTaches As Long
    BoutonStart As Long
    PrgsBarreTaches As Long
    Tray As Long
    Horloge As Long
End Type
Dim HandleElements As TypeHandleElements

Dim FlagHandlesLus As Integer
'=========================================================================
'=========================================================================

Private Sub DetectionHandleElements()
   HandleElements.BarreTaches = FindWindow("Shell_TrayWnd", "")
   HandleElements.BoutonStart = FindWindowEx(HandleElements.BarreTaches, 0, "Button" & vbNullString, vbNullString)
   HandleElements.PrgsBarreTaches = FindWindowEx(HandleElements.BarreTaches, 0, "MSTaskSwWClass" & vbNullString, vbNullString)
   HandleElements.Tray = FindWindowEx(HandleElements.BarreTaches, 0, "traynotifywnd" & vbNullString, vbNullString)
   HandleElements.Horloge = FindWindowEx(HandleElements.Tray, 0, "TrayClockWClass" & vbNullString, vbNullString)
   FlagHandlesLus = 1
End Sub

Public Function CacheElement(Element As ListeElements) As Long
   Dim Ret As Long

   If FlagHandlesLus = 0 Then DetectionHandleElements

   Select Case Element
      Case ELMT_BARRE_TACHE
          Ret = SetWindowPos(HandleElements.BarreTaches, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
      Case ELMT_BOUTON_START
          Ret = SetWindowPos(HandleElements.BoutonStart, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
      Case ELMT_PRGS_BARRE_TACHE
          Ret = SetWindowPos(HandleElements.PrgsBarreTaches, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
      Case ELMT_TRAY
          Ret = SetWindowPos(HandleElements.Tray, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
      Case ELMT_HORLOGE
          Ret = SetWindowPos(HandleElements.Horloge, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
      Case Else
   End Select
   CacheElement = Ret
End Function

Public Function AfficheElement(Element As ListeElements) As Long
   Dim Ret As Long
   
   If FlagHandlesLus = 0 Then DetectionHandleElements
   
   Select Case Element
      Case ELMT_BARRE_TACHE
          Ret = SetWindowPos(HandleElements.BarreTaches, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
      Case ELMT_BOUTON_START
          Ret = SetWindowPos(HandleElements.BoutonStart, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
      Case ELMT_PRGS_BARRE_TACHE
          Ret = SetWindowPos(HandleElements.PrgsBarreTaches, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
      Case ELMT_TRAY
          Ret = SetWindowPos(HandleElements.Tray, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
      Case ELMT_HORLOGE
          Ret = SetWindowPos(HandleElements.Horloge, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
      Case Else
   End Select
   AfficheElement = Ret
End Function

