VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWinSCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================================================
' Referinta rapida clsWinSCPManager + clsWinScpFile
'================================================================================
'
' Clasa: clsWinSCPManager
'-----------------------
' Proprietati publice:
'   FTP_USER             - username FTP/SFTP/S3
'   FTP_PASS             - parola
'   FTP_HOST             - host server
'   LogPath              - cale fisier log XML (activat daca nu e gol)
'   Protocol             - protocol FTP/SFTP/S3
'   Security             - tip securitate FTP
'   ConnectionTimeout    - timeout conexiune in milisecunde
'   SetCurrentFolder     - seteaza folderul curent
'   CurrentFolderPath    - returneaza calea folderului curent
'   IsOpen               - indica daca sesiunea este deschisa
'   SetCurrentFile       - seteaza fisierul curent
'   FileExists           - verifica daca fisierul exista pe server
'   CurrentFile          - returneaza obiect clsWinScpFile pentru fisierul curent
'   CurrentFolder        - returneaza colectia cu fisiere/foldere din folderul curent (recursiv)
'
' Functii publice:
'   OpenSession()                     - deschide sesiunea WinSCP
'   CloseSession()                    - inchide sesiunea si elibereaza resurse
'   DownloadFile(localPath, remotePath)       - descarca fisier de pe server
'   DownloadFileToFolder(localFolder, remotePath) - descarca fisier intr-un folder local
'   UploadFile(localPath, remotePath)         - incarca fisier pe server
'
' Evenimente:
'   FtpError(E)                        - eroare sesiune sau transfer
'   Progress(Percent, Size)            - progres transfer fisier
'   DownloadDone(File)                 - final descarcare fisier
'   UploadDone(RemoteFile, LastWrite)  - final upload fisier
'
' Observatii:
'   - GetDirectory() apeleaza GetDirectoryRecursive pentru a include subfoldere
'   - Logul XML este activ doar la deschiderea sesiunii
'   - WinSCPnet.dll trebuie inregistrat cu regasm
'
' Clasa: clsWinScpFile
'--------------------
' Proprietati publice:
'   Name        - numele fisierului sau folderului
'   LastWrite   - data ultimei modificari
'   Size        - dimensiunea fisierului (0 pentru foldere)
'   IsFolder    - True daca este folder, False daca este fisier
'   fullPath    - calea completa pe server
'   Permissions - permisiuni fisier/folder
'
' Observatii:
'   - Folosita ca container de date pentru colectii (Collection)
'   - Nu are metode proprii; toate operatiunile sunt gestionate de clsWinSCPManager
'   - Proprietatea IsFolder permite diferentierea intre fisiere si foldere
'
' Flux simplificat:
'
'   clsWinSCPManager
'       ¦
'       +--> OpenSession / CloseSession
'       ¦
'       +--> CurrentFolder (colCurrentDir) -> colectie clsWinScpFile
'       ¦       +--> Name, fullPath, Size, IsFolder, LastWrite, Permissions
'       ¦
'       +--> CurrentFile (clsWinScpFile)
'       ¦
'       +--> DownloadFile / UploadFile
'       ¦       +--> ridica evenimente: Progress, DownloadDone, UploadDone
'       ¦
'       +--> FtpError pentru erori sesiune sau transfer
'
'================================================================================
Option Compare Database
Option Explicit

Public Event FtpError(E As String)
Public Event Progress(Percent As Double, Size As Long)
Public Event DownloadDone(File As String)
Public Event UploadDone(RemoteFile As String, LastWrite As Date)

Public FTP_USER As String
Public FTP_PASS As String
Public FTP_HOST As String

Private WithEvents Ses As WinSCPnet.session
Attribute Ses.VB_VarHelpID = -1
Private mySessionOptions As WinSCPnet.sessionOptions

Private strCurrentDir As String
Private strCurrentFile As String
Private colCurrentDir As Collection
Private wscpFile As clsWinScpFile
Private blnIsOpen As Boolean
Private wscProtocol As WinSCPnet.Protocol
Private wscSecure As WinSCPnet.FtpSecure
Private lngTimeout As Long
Private strLogPath As String
Private blnUseLog As Boolean

Public Property Let LogPath(value As String)
strLogPath = value
blnUseLog = (value <> "")

End Property
Public Property Let Protocol(value As WinSCPnet.Protocol)
wscProtocol = value
End Property

Public Property Let Security(value As WinSCPnet.FtpSecure)
wscSecure = value
End Property

Public Property Let ConnectionTimeout(value As Long)
lngTimeout = value
End Property
'============================
' Property Let SetCurrentFolder
'============================
Public Property Let SetCurrentFolder(value As String)
On Error GoTo ErrHandler

strCurrentDir = value
Set colCurrentDir = Nothing

Exit Property

ErrHandler:
MsgBox "Error in SetCurrentFolder: " & Err.Description, vbCritical
End Property

'============================
' Property Get CurrentFolderPath
'============================
Public Property Get CurrentFolderPath() As String
On Error GoTo ErrHandler
 
CurrentFolderPath = strCurrentDir

Exit Property

ErrHandler:
MsgBox "Error in CurrentFolderPath: " & Err.Description, vbCritical
End Property

Public Property Get IsOpen() As Boolean
IsOpen = blnIsOpen
End Property

'============================
' Property Let SetCurrentFile
'============================
Public Property Let SetCurrentFile(value As String)
On Error GoTo ErrHandler

strCurrentFile = Replace(value, "\", "/")
Set wscpFile = Nothing

Exit Property

ErrHandler:
MsgBox "Error in SetCurrentFile: " & Err.Description, vbCritical
End Property

'============================
' Property Get FileExists
'============================
Public Property Get FileExists(value As String) As Boolean
On Error GoTo ErrHandler

Dim bolFileExists As Boolean
Dim wFile As Object

If Not blnIsOpen Then
    MsgBox "Conexiunea FTP nu este deschisa!", vbOKOnly + vbCritical
    Exit Property
End If

strCurrentFile = value

Set wscpFile = Nothing

If Not Ses Is Nothing Then
    bolFileExists = Ses.FileExists(value)
End If

If bolFileExists Then
    Set wFile = Ses.GetFileInfo(value)
    If Not wFile Is Nothing Then
        Set wscpFile = New clsWinScpFile
        With wscpFile
            .fullPath = wFile.FullName
            .Name = wFile.Name
            .Permissions = wFile.FilePermissions
            .LastWrite = wFile.LastWriteTime
            .IsFolder = wFile.IsDirectory
            .Size = wFile.Length32
        End With
    End If
End If

FileExists = bolFileExists
Exit Property

ErrHandler:
MsgBox "Error in FileExists property: " & Err.Description, vbCritical
FileExists = False
End Property

'============================
' Property Get CurrentFile
'============================
Public Property Get CurrentFile(Optional newFile As String = "") As clsWinScpFile
On Error GoTo ErrHandler
Dim fe As Boolean

If Not blnIsOpen Then
    MsgBox "Conexiunea FTP nu este deschisa!", vbOKOnly + vbCritical
    Exit Property
End If

If newFile <> "" Then
    If FileExists(newFile) Then
        Set CurrentFile = wscpFile
    Else
        Set CurrentFile = New clsWinScpFile  ' Avoids error if nothing found
    End If
        
ElseIf wscpFile Is Nothing Then
    fe = FileExists(strCurrentFile)
    Set CurrentFile = wscpFile

Else
    Set CurrentFile = wscpFile
End If

Exit Property

ErrHandler:
MsgBox "Error in CurrentFile property: " & Err.Description, vbCritical
Set CurrentFile = Nothing
End Property

'============================
' Property Get CurrentFolder
'============================
Public Property Get CurrentFolder(Optional newPath As String = "") As Collection
On Error GoTo ErrHandler

If Not blnIsOpen Then
    OpenSession
End If

If newPath <> "" Then
    strCurrentDir = newPath
Else
    If strCurrentDir = "" Then strCurrentDir = "/"
End If

If colCurrentDir Is Nothing Then
    Set colCurrentDir = New Collection
    GetDirectory
End If

Set CurrentFolder = colCurrentDir

Exit Property

ErrHandler:
MsgBox "Error in CurrentFolder property: " & Err.Description, vbCritical
Set CurrentFolder = Nothing
End Property

'============================
' Function DownloadFile with wildcards
'============================
Public Function DownloadFile(localFilePath As String, Optional remoteFilePath As String = "") As Boolean
Dim transferResult As Object
On Error GoTo ErrHandler

If Not blnIsOpen Then
    MsgBox "Conexiunea FTP nu este deschisa!", vbOKOnly + vbCritical
    Exit Function
End If

If remoteFilePath <> "" Then strCurrentFile = remoteFilePath

' Make sure session is open
If Ses Is Nothing Then
    MsgBox "Session is not open!", vbExclamation
    DownloadFile = False
    Exit Function
End If

' Download the remote file to the specified local file.
' The third parameter (False) means we do not remove the remote file.
Set transferResult = Ses.GetFiles(strCurrentFile, localFilePath, False)

' Check the transfer result.
If transferResult.IsSuccess Then
    DownloadFile = True
Else
    Dim failure As Variant
    For Each failure In transferResult.Failures
        MsgBox "Error downloading file: " & failure.Message, vbExclamation
    Next failure
    DownloadFile = False
End If

Exit Function

ErrHandler:
MsgBox "Error: " & Err.Description, vbCritical
DownloadFile = False
End Function

'============================
' Function DownloadFile specific file to dest folder. destfolder HAS to be a folder. the filename is kept from ftp
'============================
Public Function DownloadFileToFolder(localFolderPath As String, Optional remoteFilePath As String = "") As String
Dim transferResult As Object

On Error GoTo ErrHandler

If Not blnIsOpen Then
    Ses.Open mySessionOptions
End If

If remoteFilePath <> "" Then strCurrentFile = remoteFilePath

' Make sure session is created
If Ses Is Nothing Then
    MsgBox "Session is not created!", vbExclamation
    DownloadFileToFolder = ""
    Exit Function
End If

' Make sure session is open
If Not Ses.Opened Then
    MsgBox "Session is not opened!", vbExclamation
    DownloadFileToFolder = ""
    Exit Function
End If
' Download the remote file to the specified local file.
' The third parameter (False) means we do not remove the remote file.
Set transferResult = Ses.GetFileToDirectory(strCurrentFile, localFolderPath, False)

' Check the transfer result.
If transferResult.ERROR Is Nothing Then
    DownloadFileToFolder = transferResult.Destination
Else
    DownloadFileToFolder = ""
End If

Exit Function

ErrHandler:
MsgBox "Error: " & Err.Description, vbCritical
DownloadFileToFolder = False
End Function

Public Function UploadFile(localFilePath As String, remoteFilePath As String) As Boolean
On Error GoTo ErrHandler
Dim transferResult As Object

' Upload the file using PutFiles.
' The third parameter (False) indicates not to remove the local file after upload.
Set transferResult = Ses.PutFiles(localFilePath, remoteFilePath, False)

' Check if the upload succeeded.
If transferResult.IsSuccess Then
    UploadFile = True
Else
    Dim failure As Variant
    For Each failure In transferResult.Failures
        MsgBox "Error uploading file: " & failure.Message, vbExclamation
    Next failure
    UploadFile = False
End If

Exit Function

ErrHandler:
MsgBox "Error in UploadFile: " & Err.Description, vbCritical
UploadFile = False
End Function

'============================
' Method CloseSession
'============================
Public Sub CloseSession()
If Not Ses Is Nothing Then
    If Ses.Opened Then Ses.Close
    blnIsOpen = False
End If

End Sub

Public Sub OpenSession()
On Error GoTo ErrHandler

If Len(FTP_HOST) = 0 Or Len(FTP_PASS) = 0 Or Len(FTP_USER) = 0 Then
    MsgBox "Connection info is missing!", vbOKOnly + vbCritical
    Exit Sub
End If

blnIsOpen = False

Init_Options

If mySessionOptions Is Nothing Then
    MsgBox "FTP options not initialized!", vbOKOnly + vbCritical
    Exit Sub
End If

Set Ses = New WinSCPnet.session
If blnUseLog And Len(strLogPath) <> 0 Then
    Ses.XmlLogPath = strLogPath
    Ses.XmlLogPreserve = False
End If

If Not Ses Is Nothing Then
    Ses.Open mySessionOptions
    blnIsOpen = Ses.Opened
    strCurrentDir = "/"
Else
    MsgBox "FTP connection not opened!", vbOKOnly + vbCritical
    Exit Sub
End If

Exit Sub

ErrHandler:
MsgBox "Error in UploadFile: " & Err.Description, vbCritical
End Sub

'============================
' Class_Initialize
'============================
Private Sub Class_Initialize()
On Error GoTo ErrHandler

wscProtocol = Protocol_S3
wscSecure = FtpSecure_Implicit
strLogPath = ""
blnUseLog = False
lngTimeout = 5000

Exit Sub

ErrHandler:
MsgBox "Error in Class_Initialize: " & Err.Description, vbCritical
End Sub

Private Sub Init_Options()
Set mySessionOptions = New WinSCPnet.sessionOptions

With mySessionOptions
    .Protocol = wscProtocol
    .FtpSecure = wscSecure
    .HostName = FTP_HOST
    .UserName = FTP_USER
    .Password = FTP_PASS
    .TimeoutInMilliseconds = lngTimeout
End With
End Sub
'============================
' GetDirectoryRecursive
'============================
Private Sub GetDirectoryRecursive(ByVal path As String, ByRef colAll As Collection)
On Error GoTo ErrHandler

Dim directoryInfo As Object
Dim fileInfo As Object
Dim cFile As clsWinScpFile

Set directoryInfo = Ses.ListDirectory(path)

For Each fileInfo In directoryInfo.Files
    If fileInfo.Name <> ".." Then
        Set cFile = New clsWinScpFile
        With cFile
            .fullPath = fileInfo.FullName
            .Name = fileInfo.Name
            .Permissions = fileInfo.FilePermissions
            .LastWrite = fileInfo.LastWriteTime
            .IsFolder = fileInfo.IsDirectory
            .Size = fileInfo.Length32
        End With
        colAll.Add cFile

        ' Daca este folder, intra recursiv
        If fileInfo.IsDirectory Then
            GetDirectoryRecursive fileInfo.FullName, colAll
        End If
    End If
Next fileInfo

Exit Sub
ErrHandler:
MsgBox "Error in GetDirectoryRecursive: " & Err.Description, vbCritical
End Sub

'============================
' GetDirectory
'============================
Private Sub GetDirectory()
On Error GoTo ErrHandler

If Not blnIsOpen Then
    MsgBox "Conexiunea FTP nu este deschisa!", vbOKOnly + vbCritical
    Exit Sub
End If

Set colCurrentDir = New Collection
GetDirectoryRecursive strCurrentDir, colCurrentDir

' Set colCurrentDir = SortCollection(colCurrentDir)

Exit Sub
ErrHandler:
MsgBox "Error in GetDirectory: " & Err.Description, vbCritical
End Sub

'============================
' BubbleSort Function
'============================
Private Function BubbleSort(arr As Variant) As Variant
On Error GoTo ErrHandler
Dim I As Long, K As Long
Dim temp As Variant
Dim lb As Long, ub As Long, N As Long

' Get the lower and upper bounds of the array
lb = LBound(arr)
ub = UBound(arr)
N = ub - lb + 1

' Perform bubble sort: repeat (n-1) times
For K = 1 To N - 1
    For I = lb To ub - K
        If IsEmpty(arr(I + 1)) Then Exit For
        If arr(I).Name > arr(I + 1).Name Then
            Set temp = arr(I)
            Set arr(I) = arr(I + 1)
            Set arr(I + 1) = temp
        End If
    Next I
Next K

BubbleSort = arr
Exit Function
ErrHandler:
MsgBox "Error in BubbleSort: " & Err.Description, vbCritical
BubbleSort = arr
End Function

'============================
' Class_Terminate
'============================
Private Sub Class_Terminate()
On Error Resume Next
    If Not Ses Is Nothing Then Ses.dispose
    Set mySessionOptions = Nothing
    Set colCurrentDir = Nothing
End Sub

'============================
' Ses_Failed Event Handler
'============================
Private Sub Ses_Failed(ByVal sender As Variant, ByVal E As WinSCPnet.FailedEventArgs)
RaiseEvent FtpError(CStr(E))
End Sub

Private Sub Ses_FileTransferProgress(ByVal sender As Variant, ByVal E As WinSCPnet.FileTransferProgressEventArgs)
RaiseEvent Progress(E.FileProgress, E.cps)
End Sub

Private Sub Ses_FileTransferred(ByVal sender As Variant, ByVal E As WinSCPnet.TransferEventArgs)
'extra test if the remote size is equal to the downloaded size
If E.side = WinSCPnet.ProgressSide_Remote Then
    If wscpFile.Size = E.Length Then
        RaiseEvent DownloadDone(E.fileName)
    End If
Else
    Me.SetCurrentFile = E.touch.fileName
    RaiseEvent UploadDone(E.touch.fileName, E.touch.LastWriteTime)
End If
End Sub
