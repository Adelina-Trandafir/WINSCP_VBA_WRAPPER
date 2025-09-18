# VBA WinSCP FTP/SFTP Manager

## English

This project provides a VBA class (`clsWinSCPManager`) to manage FTP, SFTP, or S3 file transfers using **WinSCPnet.dll**. It allows opening sessions, uploading/downloading files, and listing directories recursively. Files and folders are represented by the `clsWinScpFile` class.

### Features

- Open and close FTP/SFTP/S3 sessions
- Upload and download files
- List directories recursively
- Track transfer progress via events
- Capture errors via events
- Optional XML logging
- Works with Windows 32-bit and 64-bit VBA (Access, Excel, etc.)

### Classes

#### clsWinSCPManager

- **Properties**
  - `FTP_USER`, `FTP_PASS`, `FTP_HOST` — connection credentials
  - `LogPath` — optional XML log path
  - `Protocol`, `Security` — FTP/SFTP/S3 options
  - `ConnectionTimeout` — timeout in milliseconds
  - `SetCurrentFolder` — set current folder on server
  - `CurrentFolder` — collection of `clsWinScpFile` objects
  - `SetCurrentFile`, `CurrentFile`, `FileExists` — file operations
  - `IsOpen` — indicates if the session is active

- **Methods**
  - `OpenSession()` — open session
  - `CloseSession()` — close session
  - `DownloadFile(localPath, remotePath)` — download file
  - `DownloadFileToFolder(localFolder, remotePath)` — download to folder
  - `UploadFile(localPath, remotePath)` — upload file

- **Events**
  - `FtpError(E)` — session or transfer error
  - `Progress(Percent, Size)` — file transfer progress
  - `DownloadDone(File)` — file download completed
  - `UploadDone(RemoteFile, LastWrite)` — file upload completed

#### clsWinScpFile

- Represents a file or folder on the server
- **Properties:** `Name`, `LastWrite`, `Size`, `IsFolder`, `fullPath`, `Permissions`
- Used in collections (`CurrentFolder`) to store directory contents
- Differentiates files and folders with `IsFolder` flag

### Requirements

- Windows
- VBA (Access, Excel, etc.)
- **WinSCPnet.dll** registered with `regasm`

---

## Română

Acest proiect oferă o clasă VBA (`clsWinSCPManager`) pentru gestionarea transferurilor FTP, SFTP sau S3 folosind **WinSCPnet.dll**. Permite deschiderea sesiunilor, încărcarea și descărcarea fișierelor și listarea recursivă a directoarelor. Fișierele și folderele sunt reprezentate de clasa `clsWinScpFile`.

### Funcționalități

- Deschide și închide sesiuni FTP/SFTP/S3
- Încarcă și descarcă fișiere
- Listează directoare recursiv
- Monitorizează progresul transferului prin evenimente
- Capturează erorile prin evenimente
- Log XML opțional
- Funcționează pe Windows 32-bit și 64-bit (Access, Excel etc.)

### Clase

#### clsWinSCPManager

- **Proprietăți**
  - `FTP_USER`, `FTP_PASS`, `FTP_HOST` — credențiale conexiune
  - `LogPath` — cale optională pentru log XML
  - `Protocol`, `Security` — opțiuni FTP/SFTP/S3
  - `ConnectionTimeout` — timeout în milisecunde
  - `SetCurrentFolder` — setează folderul curent pe server
  - `CurrentFolder` — colecție de obiecte `clsWinScpFile`
  - `SetCurrentFile`, `CurrentFile`, `FileExists` — operațiuni pe fișiere
  - `IsOpen` — indică dacă sesiunea este activă

- **Metode**
  - `OpenSession()` — deschide sesiunea
  - `CloseSession()` — închide sesiunea
  - `DownloadFile(localPath, remotePath)` — descarcă fișier
  - `DownloadFileToFolder(localFolder, remotePath)` — descarcă într-un folder
  - `UploadFile(localPath, remotePath)` — încarcă fișier

- **Evenimente**
  - `FtpError(E)` — eroare sesiune sau transfer
  - `Progress(Percent, Size)` — progres transfer fișier
  - `DownloadDone(File)` — fișier descărcat complet
  - `UploadDone(RemoteFile, LastWrite)` — fișier încărcat complet

#### clsWinScpFile

- Reprezintă un fișier sau folder de pe server
- **Proprietăți:** `Name`, `LastWrite`, `Size`, `IsFolder`, `fullPath`, `Permissions`
- Folosit în colecții (`CurrentFolder`) pentru a stoca conținutul unui folder
- Distinge fișierele de foldere prin flag-ul `IsFolder`

### Cerințe

- Windows
- VBA (Access, Excel etc.)
- **WinSCPnet.dll** înregistrat cu `regasm`
