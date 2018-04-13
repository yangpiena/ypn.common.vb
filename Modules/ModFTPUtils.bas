Attribute VB_Name = "ModFTPUtils"
Option Explicit

' Constants - InternetOpen.lAccessType
Public Const INTERNET_OPEN_TYPE_PRECONFIG As Long = 0&
Public Const INTERNET_OPEN_TYPE_DIRECT As Long = 1&
Public Const INTERNET_OPEN_TYPE_PROXY As Long = 3&
Public Const INTERNET_OPEN_TYPE_PRECONFIG_WITH_NO_AUTOPROXY As Long = 4&

' Constants - InternetOpen.dwFlags
Public Const INTERNET_FLAG_ASYNC As Long = &H10000000
Public Const INTERNET_FLAG_FROM_CACHE As Long = &H1000000
Public Const INTERNET_FLAG_OFFLINE As Long = INTERNET_FLAG_FROM_CACHE

' Constants - InternetConnect.nServerPort
Public Const INTERNET_INVALID_PORT_NUMBER As Long = 0&
Public Const INTERNET_DEFAULT_FTP_PORT As Long = 21&
Public Const INTERNET_DEFAULT_GOPHER_PORT As Long = 70&
Public Const INTERNET_DEFAULT_HTTP_PORT As Long = 80&
Public Const INTERNET_DEFAULT_HTTPS_PORT As Long = 443&
Public Const INTERNET_DEFAULT_SOCKS_PORT As Long = 1080&

' Constants - InternetConnect.dwService
Public Const INTERNET_SERVICE_FTP As Long = 1&
Public Const INTERNET_SERVICE_GOPHER As Long = 2&
Public Const INTERNET_SERVICE_HTTP As Long = 3&

' Constants - InternetConnect.dwFlags
Public Const INTERNET_FLAG_PASSIVE As Long = &H8000000

' Constants - FtpGetFile.dwFlags (FTP TransferType)
' Constants - FtpPutFile.dwFlags (FTP TransferType)
Public Const FTP_TRANSFER_TYPE_UNKNOWN As Long = &H0&
Public Const FTP_TRANSFER_TYPE_ASCII As Long = &H1&
Public Const FTP_TRANSFER_TYPE_BINARY As Long = &H2&
Public Const INTERNET_FLAG_TRANSFER_ASCII As Long = FTP_TRANSFER_TYPE_ASCII
Public Const INTERNET_FLAG_TRANSFER_BINARY As Long = FTP_TRANSFER_TYPE_BINARY

' Constants - FtpGetFile.dwFlags (Cache Flags)
' Constants - FtpPutFile.dwFlags (Cache Flags)
Public Const INTERNET_FLAG_RELOAD As Long = &H80000000
Public Const INTERNET_FLAG_RESYNCHRONIZE As Long = &H800
Public Const INTERNET_FLAG_NEED_FILE As Long = &H10
Public Const INTERNET_FLAG_HYPERLINK As Long = &H400

' Constants - FtpGetFile.dwFlagsAndAttributes
Public Const FILE_ATTRIBUTE_ARCHIVE As Long = &H20
Public Const FILE_ATTRIBUTE_ENCRYPTED As Long = &H4000
Public Const FILE_ATTRIBUTE_HIDDEN As Long = &H2
Public Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Public Const FILE_ATTRIBUTE_OFFLINE As Long = &H1000
Public Const FILE_ATTRIBUTE_READONLY As Long = &H1
Public Const FILE_ATTRIBUTE_SYSTEM As Long = &H4
Public Const FILE_ATTRIBUTE_TEMPORARY As Long = &H100
Public Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10
Public Const FILE_ATTRIBUTE_COMPRESSED As Long = &H800

'=================
' FILETIME 结构体
'=================
Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

'========================
' WIN32_FIND_DATA 结构体
'========================
Public Const MAX_PATH = 260
Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type

'=====================================
' 取得InternetHandle
'=====================================
Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
(ByVal lpszAgent As String, ByVal dwAccessType As Long, ByVal lpszProxyName As String, _
ByVal lpszProxyBypass As String, ByVal dwFlags As Long) As Long

'=====================
' 连接FTP
'=====================
Public Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" _
(ByVal HINTERNET As Long, ByVal lpszServerName As String, ByVal nServerPort As Integer, _
ByVal lpszUsername As String, ByVal lpszPassword As String, ByVal dwService As Long, _
ByVal dwFlags As Long, ByVal dwContext As Long) As Long

'===================================
' 关闭InternetHandle
'===================================
Public Declare Function InternetCloseHandle Lib "wininet.dll" _
(ByVal HINTERNET As Long) As Integer

'===========================================
' 取得Server的CurrentDirectory
'===========================================
Public Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" _
(ByVal hConnect As Long, ByVal lpszCurrentDirectory As String, _
ByRef lpdwCurrentDirectory As Long) As Boolean

'===========================================
' 设定Server的CurrentDirectory
'===========================================
Public Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" _
(ByVal hConnect As Long, ByVal lpszDirectory As String) As Long

'=================================
' 从Server取得文件
'=================================
Public Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" _
(ByVal hConnect As Long, ByVal lpszRemoteFile As String, _
ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, ByVal dwFlagsAndAttributes As Long, _
ByVal dwFlags As Long, ByVal dwContext As Long) As Long

'===============================
' 向Server传送文件
'===============================
Public Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" _
(ByVal hConnect As Long, ByVal lpszLocalFile As String, _
ByVal lpszNewRemoteFile As String, _
ByVal dwFlags As Long, ByVal dwContext As Long) As Long

'===============================
' 删除Server上面的文件
'===============================
Public Declare Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" _
(ByVal hConnect As Long, ByVal lpszFileName As String) As Long

'=================================
' 变更Server上的文件名
'=================================
Public Declare Function FtpRenameFile Lib "wininet.dll" Alias "FtpRenameFileA" _
(ByVal hConnect As Long, ByVal lpszExisting As String, ByVal lpszNew As String) As Long

'===================================
' 删除Server上的目录
'===================================
Public Declare Function FtpRemoveDirectory Lib "wininet.dll" Alias "FtpRemoveDirectoryA" _
(ByVal hConnect As Long, ByVal lpszDirectory As String) As Long

'======================================
' 检索指定的路径
'======================================
Public Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" _
(ByVal hFtpSession As Long, ByVal lpszSearchFile As String, _
lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContent As Long) As Long

'======================================
' 继续检索下一个路径
'======================================
Public Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" _
(ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long

'===============================================
' 从FTP中取得指定的目录下的内容
'===============================================
Public Sub Sample()
  Dim hOpen As Long       'InternetServer的Handle
  Dim hConnection As Long 'InternetSession的Handle
  Dim result As Long
  hOpen = 0
  hConnection = 0

  Dim hFind As Long
  Dim w32FindData As WIN32_FIND_DATA
  Dim strFile As String

  Dim FileList() As String '文件名一栏
  Dim cnt As Long
  cnt = -1

  '取得InternetServer的Handle - hOpen
  hOpen = InternetOpen("FTPSample", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
  If (hOpen <> 0) Then 'Handle取得成功

    '取得InternetSession的Handle（连接FTPServer） - hConnection
    hConnection = InternetConnect(hOpen, "192.168.45.12", INTERNET_INVALID_PORT_NUMBER, _
        "UserName", "Password", INTERNET_SERVICE_FTP, 0, 0)
    If (hConnection <> 0) Then '连接成功

      '改变FTPServer的CurrentDirectory
      result = FtpSetCurrentDirectory(hConnection, "home/cadsvr/plan/plan14735")
      If (result <> 0) Then '变更成功

        '取得文件一览
        hFind = FtpFindFirstFile(hConnection, "*.*", w32FindData, INTERNET_FLAG_RELOAD, 0)
        If (hFind = 0) Then
          MsgBox "文件名取得失败" & Err.LastDllError
        Else
          Do
            strFile = Left(w32FindData.cFileName, InStr(w32FindData.cFileName, vbNullChar) - 1)
            strFile = Mid(strFile, InStrRev(strFile, " ") + 1) '删除文件名中无用的字符
            If ((w32FindData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = &H10) Then
              strFile = strFile & "/" '如果取得的是目录，在目录后面加上/
            End If
            cnt = cnt + 1
            ReDim Preserve FileList(cnt)
            FileList(cnt) = strFile '将取得的文件名或者目录名，追加到文件名一览中
          Loop Until InternetFindNextFile(hFind, w32FindData) = 0 '取得下一个文件名
        End If

      Else
        MsgBox "目录移动失败" & Err.LastDllError
      End If
    Else
      MsgBox "FTPServer连接失败" & Err.LastDllError
    End If
  Else
    MsgBox "FTPServer连接失败" & Err.LastDllError
  End If

  '关闭InternetSession
  If (hConnection <> 0) Then InternetCloseHandle hConnection

  '关闭InternetServer
  If (hOpen <> 0) Then InternetCloseHandle hOpen

End Sub

'---------------------------------------------------------------------------------
'---------------------------------------------------------------------------------
Function FTPFileDownload(strSvrIP As String, strFtpUser As String, strFtpPass As String, strDownPath As String, strNewFilePath As String, DelFlg As Boolean) As Boolean
'---------------------------------------------------------------------------------
'FTP文件删除
'参数：
' strSvrIP=FTPServer的IP
' strFtpUser=用户名
' strFtpPass=密码
' strDownPath=下载的文件的全名
' DelFlg=删除Flag True:删除 False:不删除
'---------------------------------------------------------------------------------
Dim lnghInet                            As Long
Dim lnghConnect                         As Long
Dim lnghFile                            As Long
Dim lngReturn                           As Long
Dim udtFindData                         As WIN32_FIND_DATA
Dim booReturn                           As Boolean
'---------------------------------------------------------------------------------
Dim ret
'---------------------------------------------------------------------------------

    FTPFileDownload = False
    
    '初始化
    lnghInet = InternetOpen(App.EXEName, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)

    '连接FTP
    lnghConnect = InternetConnect(lnghInet, strSvrIP, INTERNET_INVALID_PORT_NUMBER, strFtpUser, strFtpPass, INTERNET_SERVICE_FTP, 0, 0)
    
    '文件存在确认
    lnghFile = FtpFindFirstFile(lnghConnect, strDownPath, udtFindData, INTERNET_FLAG_RELOAD, 0)
    ret = Left(udtFindData.cFileName, InStr(udtFindData.cFileName, vbNullChar) - 1)
    If ret = "" Then Exit Function
    
    '文件取得
    ret = FtpGetFile(lnghConnect, strDownPath, strNewFilePath, False, FILE_ATTRIBUTE_NORMAL, INTERNET_FLAG_RELOAD, 0&)
    If ret = 0 Then Exit Function
    
    '文件删除
    If DelFlg = True Then
        booReturn = FtpDeleteFile(lnghConnect, strDownPath)
        Sleep (500)
        If booReturn = False Then Exit Function
    End If
    
    lngReturn = InternetCloseHandle(lnghInet)
    lngReturn = InternetCloseHandle(lnghConnect)
    lngReturn = InternetCloseHandle(lnghFile)
    
    FTPFileDownload = True
    
End Function

'---------------------------------------------------------------------------------
Function FTPFileUpload(strSvrIP As String, strFtpUser As String, strFtpPass As String, strUpPath As String, strNewFilePath As String, DelFlg As Boolean) As Boolean
'---------------------------------------------------------------------------------
'FTP文件上传
'参数：
' strSvrIP=FTPServer的IP
' strFtpUser=用户名
' strFtpPass=密码
' strUpPath=上传文件在FTP上的保存路径
' DelFlg=删除Flag True:删除 False:不删除
'---------------------------------------------------------------------------------
Dim lnghInet                            As Long
Dim lnghConnect                         As Long
Dim lnghFile                            As Long
Dim lngReturn                           As Long
Dim udtFindData                         As WIN32_FIND_DATA
Dim booReturn                           As Boolean
'---------------------------------------------------------------------------------
Dim ret
'---------------------------------------------------------------------------------

    FTPFileUpload = False
    
    '初始化
    lnghInet = InternetOpen(App.EXEName, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)

    '连接FTP
    lnghConnect = InternetConnect(lnghInet, strSvrIP, INTERNET_INVALID_PORT_NUMBER, strFtpUser, strFtpPass, INTERNET_SERVICE_FTP, 0, 0)
    
'    '文件存在确认
'    lnghFile = FtpFindFirstFile(lnghConnect, strUpPath, udtFindData, INTERNET_FLAG_RELOAD, 0)
'    ret = Left(udtFindData.cFileName, InStr(udtFindData.cFileName, vbNullChar) - 1)
'    If ret = "" Then Exit Function
    
    '上传
    ret = FtpPutFile(lnghConnect, strUpPath, strNewFilePath, INTERNET_FLAG_RELOAD, 0&)
    If ret = 0 Then Exit Function
    
'    '删除
'    If DelFlg = True Then
'        booReturn = FtpDeletefile(lnghConnect, strUpPath)
'        Sleep (500)
'        If booReturn = False Then Exit Function
'    End If
    
    lngReturn = InternetCloseHandle(lnghInet)
    lngReturn = InternetCloseHandle(lnghConnect)
'    lngReturn = InternetCloseHandle(lnghFile)
    
    FTPFileUpload = True
    
End Function
