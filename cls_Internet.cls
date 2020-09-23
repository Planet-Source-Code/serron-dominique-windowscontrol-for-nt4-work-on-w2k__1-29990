VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Internet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
        (ByVal lpszAgent As String, ByVal dwAccessType As Long, ByVal lpszProxyName As String, _
         ByVal lpszProxyBypass As String, ByVal dwFlags As Long) As Long

Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer

Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" _
        (ByVal hInternetSession As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, _
         ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long

Private Declare Function InternetReadFile Lib "wininet.dll" _
        (ByVal hFile As Long, ByVal lpBuffer As String, ByVal dwNumberOfBytesToRead As Long, _
         lNumberOfBytesRead As Long) As Integer

Const INTERNET_OPEN_TYPE_PRECONFIG = 0   ' utiliser info de config de la base de registre
Const INTERNET_FLAG_EXISITING_CONNECT = &H20000000
Const INTERNET_FLAG_RELOAD = &H80000000 ' read from wire even if locally cached

Public Function ReadFile(InetAdd As String) As String
   Dim NbOctets As Long
   Dim hSession As Long
   Dim hUrlFile As Long
   Dim Ret As Integer
   Dim Buffer As String * 1024
   
   hSession = InternetOpen("My App", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, _
                           vbNullString, 0)

   hUrlFile = InternetOpenUrl(hSession, InetAdd, vbNullString, _
                              0, INTERNET_FLAG_RELOAD, 0)

   Ret = InternetReadFile(hUrlFile, Buffer, Len(Buffer), NbOctets)

   InternetCloseHandle (hUrlFile)
   InternetCloseHandle (hSession)

   ReadFile = Left$(Buffer, NbOctets)
 '  Info = ElimineChar(Info)
End Function
