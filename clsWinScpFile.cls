VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWinScpFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================
' Clasa clsWinScpFile
'=========================================================
'
' Aceasta clasa reprezinta un fisier sau folder de pe serverul FTP/SFTP/S3.
' Este folosita ca un simplu container de date pentru fiecare element dintr-un director.
'
' Proprietati:
'   Name        - numele fisierului sau folderului
'   LastWrite   - data ultimei modificari
'   Size        - dimensiunea fisierului in bytes (0 pentru foldere)
'   IsFolder    - True daca este folder, False daca este fisier
'   fullPath    - calea completa pe server, ex: "/folder1/file.txt"
'   Permissions - permisiuni fisier/folder, ex: "rwxr-xr-x"
'
' Observatii:
'   - Nu are metode proprii; toate operatiunile (download, upload, verificare, listare)
'     sunt gestionate de clasa principala clsWinSCPManager.
'   - Poate fi folosita in colectii (Collection) pentru a retine continutul unui folder.
'   - Proprietatea IsFolder permite diferentierea intre fisiere si foldere
'     fara a crea o clasa separata pentru foldere.
'
' Exemplu de utilizare:
'   Dim f As clsWinScpFile
'   Set f = New clsWinScpFile
'   f.Name = "test.txt"
'   f.IsFolder = False
'   f.fullPath = "/folder1/test.txt"
'   f.Size = 1024
'   f.LastWrite = Now
'   f.Permissions = "rw-r--r--"
'=========================================================
Option Compare Database
Option Explicit

Public Name As String
Public LastWrite As Date
Public Size As Long
Public IsFolder As Boolean
Public fullPath As String
Public Permissions As String
