VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.UserControl SSLSocket 
   ClientHeight    =   525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1545
   InvisibleAtRuntime=   -1  'True
   Picture         =   "sslControl.ctx":0000
   ScaleHeight     =   525
   ScaleWidth      =   1545
   ToolboxBitmap   =   "sslControl.ctx":07AE
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   960
      Top             =   0
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   480
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "SSLSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event SSLConnect()
Public Event SSLClose()
Public Event SSLData(sData As String)
Public Event SSLError(iNumber As Integer, sDescription As String)
Public Event SSLProxyError()
Public Event SSLConnecting()
Public Event SSLSendingData(sData As String)

Private iLayer As Integer
Private InBuffer As String
Private iSeekLen As Integer
Private sBuffer As String

Private MASTER_KEY As String
Private CLIENT_READ_KEY As String
Private CLIENT_WRITE_KEY As String

Private Private_KEY As String
Private ENCODED_CERT As String
Private CONNECTION_ID As String

Private SEND_SEQUENCE_NUMBER As Double
Private RECV_SEQUENCE_NUMBER As Double

Private CLIENT_HELLO As String
Private CHALLENGE_DATA As String

Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hSessionKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, ByVal pbData As String, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32.dll" (ByVal hHash As Long, ByVal dwParam As Long, ByVal pbData As String, ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Private Declare Function CryptDeriveKey Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hBaseData As Long, ByVal dwFlags As Long, ByRef hSessionKey As Long) As Long
Private Declare Function CryptEncrypt Lib "advapi32.dll" (ByVal hSessionKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, ByRef pdwDataLen As Long, ByVal dwBufLen As Long) As Long
Private Declare Function CryptDecrypt Lib "advapi32.dll" (ByVal hSessionKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, ByRef pdwDataLen As Long) As Long
Private Declare Function CryptDestroyKey Lib "advapi32.dll" (ByVal hSessionKey As Long) As Long
Private Declare Function CryptImportKey Lib "advapi32.dll" (ByVal hProv As Long, ByVal pbData As String, ByVal dwDataLen As Long, ByVal hPubKey As Long, ByVal dwFlags As Long, ByRef phKey As Long) As Long
Private Declare Function CryptExportKey Lib "advapi32.dll" (ByVal hSessionKey As Long, ByVal hExpKey As Long, ByVal dwBlobType As Long, ByVal dwFlags As Long, ByVal pbData As String, ByRef pdwDataLen As Long) As Long
Private Declare Function CryptGenRandom Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwLen As Long, ByVal pbBuffer As String) As Long

Private Const SERVICE_PROVIDER As String = "Microsoft Enhanced Cryptographic Provider v1.0" & vbNullChar
Private Const KEY_CONTAINER As String = "GCN SSL Container" & vbNullChar '"Puddys Botmaker" & vbNullChar
Private Const PROV_RSA_FULL As Long = 1
Private Const CRYPT_NEWKEYSET As Long = 8
Private Const CRYPT_EXPORTABLE As Long = 1
Private Const CALG_MD5 As Long = 32771
Private Const CALG_RC4 As Long = 26625
Private Const HP_HASHVAL As Long = 2
Private Const SIMPLEBLOB As Long = 1
Private Const GEN_KEY_BITS As Long = &H800000

Private ProxyState As Integer
Private m_RemoteHost As String
Private m_RemotePort As Long

Private hCryptProv As Long
Private hClientWriteKey As Long
Private hClientReadKey As Long
Private hMasterKey As Long
Private lngType As Long
Private bRaisedConnected As Boolean

Public Property Get RemoteHost() As String
    RemoteHost = m_RemoteHost
End Property

Public Property Let RemoteHost(Host As String)
    m_RemoteHost = Host
End Property

Private Function ExportKeyBlob(ByRef StrMasterKey As String, ByRef StrReadKey As String, ByRef StrWriteKey As String, ByVal StrChallenge As String, ByVal StrConnectionID As String, ByVal StrPrivateKey As String) As String

    Dim lngReturnValue As Long
    Dim lngLength As Long
    Dim rgbBlob As String
    Dim hPrivateKey As Long
    
    Call CreateKey(hMasterKey, StrMasterKey)
    StrMasterKey = MD5_Hash(StrMasterKey)
    
    Call CreateKey(hClientReadKey, StrMasterKey & "0" & StrChallenge & StrConnectionID)
    Call CreateKey(hClientWriteKey, StrMasterKey & "1" & StrChallenge & StrConnectionID)
    
    StrReadKey = MD5_Hash(StrMasterKey & "0" & StrChallenge & StrConnectionID)
    StrWriteKey = MD5_Hash(StrMasterKey & "1" & StrChallenge & StrConnectionID)

    lngReturnValue = CryptImportKey(hCryptProv, StrPrivateKey, Len(StrPrivateKey), 0, 0, hPrivateKey)

    lngReturnValue = CryptExportKey(hMasterKey, hPrivateKey, SIMPLEBLOB, 0, vbNull, lngLength)
    rgbBlob = String(lngLength, 0)
    lngReturnValue = CryptExportKey(hMasterKey, hPrivateKey, SIMPLEBLOB, 0, rgbBlob, lngLength)
    
    If hPrivateKey <> 0 Then CryptDestroyKey hPrivateKey
    If hMasterKey <> 0 Then CryptDestroyKey hMasterKey

    ExportKeyBlob = ReverseString(Right(rgbBlob, 128))
End Function

Private Sub CreateKey(ByRef KeyName As Long, ByVal HashData As String)

    Dim lngParams As Long
    Dim lngReturnValue As Long
    Dim lngHashLen As Long
    Dim hHash As Long
    
    lngReturnValue = CryptCreateHash(hCryptProv, CALG_MD5, 0, 0, hHash)
    If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "Could not create a Hash Object (CryptCreateHash API)"
    
    lngReturnValue = CryptHashData(hHash, HashData, Len(HashData), 0)
    If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "Could not calculate a Hash Value (CryptHashData API)"
    
    lngParams = GEN_KEY_BITS Or CRYPT_EXPORTABLE
    lngReturnValue = CryptDeriveKey(hCryptProv, CALG_RC4, hHash, lngParams, KeyName)
    If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "Could not create a session key (CryptDeriveKey API)"
    
    If hHash <> 0 Then CryptDestroyHash hHash
End Sub

Private Function RC4_Encrypt(ByVal Plaintext As String) As String

    Dim lngLength As Long
    Dim lngReturnValue As Long
    
    lngLength = Len(Plaintext)
    lngReturnValue = CryptEncrypt(hClientWriteKey, 0, False, 0, Plaintext, lngLength, lngLength)

    RC4_Encrypt = Plaintext
End Function

Private Function RC4_Decrypt(ByVal Ciphertext As String) As String

    Dim lngLength As Long
    Dim lngReturnValue As Long
    
    lngLength = Len(Ciphertext)
    lngReturnValue = CryptDecrypt(hClientReadKey, 0, False, 0, Ciphertext, lngLength)

    RC4_Decrypt = Ciphertext
End Function

Private Function GenerateRandomBytes(ByVal Length As Long, ByRef TheString As String) As Boolean

    Dim i As Integer

    Randomize
    TheString = ""
    For i = 1 To Length
        TheString = TheString & Chr(Int(Rnd * 256))
    Next
    
    GenerateRandomBytes = CryptGenRandom(hCryptProv, Length, TheString)
End Function

Private Function MD5_Hash(ByVal TheString As String) As String

    Dim lngReturnValue As Long
    Dim strHash As String
    Dim hHash As Long
    Dim lngHashLen As Long
    
    lngReturnValue = CryptCreateHash(hCryptProv, CALG_MD5, 0, 0, hHash)
    lngReturnValue = CryptHashData(hHash, TheString, Len(TheString), 0)
    lngReturnValue = CryptGetHashParam(hHash, HP_HASHVAL, vbNull, lngHashLen, 0)
    strHash = String(lngHashLen, vbNullChar)
    lngReturnValue = CryptGetHashParam(hHash, HP_HASHVAL, strHash, lngHashLen, 0)
    
    If hHash <> 0 Then CryptDestroyHash hHash
    
    MD5_Hash = strHash
End Function

Public Sub ConnectSSL(Optional sUrl As String, Optional sPort As Long = 443, Optional Proxy)
    If sUrl <> "" Then m_RemoteHost = sUrl
    m_RemotePort = sPort
    
    If IsMissing(Proxy) = False Then
        ProxyState = 1
        Timer1.Enabled = True
        Socket.Close
        Socket.Connect Split(Proxy, ":")(0), Split(Proxy, ":")(1)
    Else
        ProxyState = 0
        Socket.Close
        Socket.Connect m_RemoteHost, sPort
    End If
    
    RaiseEvent SSLConnecting
End Sub

Public Sub CloseSSL()

    Socket.Close
    
    RaiseEvent SSLClose
End Sub

Private Sub CertToPrivateKey()

    Const lPbkLen As Long = 1024
    Dim lOffset As Long
    Dim lStart As Long
    Dim sBlkLen As String
    Dim sRevKey As String
    Dim ASNStart As Long
    Dim ASNKEY As String

    lOffset = CLng(lPbkLen \ 8)
    lStart = 5 + (lOffset \ 128) * 2

    ASNStart = InStr(1, ENCODED_CERT, Chr(48) & Chr(129) & Chr(137) & Chr(2) & Chr(129) & Chr(129) & Chr(0)) + lStart
    ASNKEY = Mid(ENCODED_CERT, ASNStart, 128)

    sRevKey = ReverseString(ASNKEY)

    sBlkLen = CStr(Hex(lPbkLen \ 256))
    If Len(sBlkLen) = 1 Then sBlkLen = "0" & sBlkLen

    Private_KEY = (HexToBin( _
            "06020000" & _
            "00A40000" & _
            "52534131" & _
            "00" & sBlkLen & "0000" & _
            "01000100") & sRevKey)
End Sub

Private Function VerifyMAC(ByVal DecryptedRecord As String) As Boolean

    Dim PrependedMAC As String
    Dim RecordData As String
    Dim CalculatedMAC As String
    
    PrependedMAC = Mid(DecryptedRecord, 1, 16)
    RecordData = Mid(DecryptedRecord, 17)
    
    CalculatedMAC = MD5_Hash(CLIENT_READ_KEY & RecordData & RecvSequence)
    
    Call IncrementRecv

    If CalculatedMAC = PrependedMAC Then
        VerifyMAC = True
    Else
        VerifyMAC = False
    End If
End Function

Private Function SendSequence() As String

    Dim TempString As String
    Dim TempSequence As Double
    Dim TempByte As Double
    Dim i As Integer
    
    TempSequence = SEND_SEQUENCE_NUMBER
    
    For i = 1 To 4
        TempByte = 256 * ((TempSequence / 256) - Int(TempSequence / 256))
        TempSequence = Int(TempSequence / 256)
        TempString = Chr(TempByte) & TempString
    Next
    
    SendSequence = TempString
End Function

Private Function RecvSequence() As String

    Dim TempString As String
    Dim TempSequence As Double
    Dim TempByte As Double
    Dim i As Integer
    TempSequence = RECV_SEQUENCE_NUMBER
    
    For i = 1 To 4
        TempByte = 256 * ((TempSequence / 256) - Int(TempSequence / 256))
        TempSequence = Int(TempSequence / 256)
        TempString = Chr(TempByte) & TempString
    Next
    
    RecvSequence = TempString
End Function

Private Sub SendClientHello(ByRef Socket As Winsock)

    iLayer = 0
    
    bRaisedConnected = False
    
    Call GenerateRandomBytes(16, CHALLENGE_DATA)
    
    SEND_SEQUENCE_NUMBER = 0
    RECV_SEQUENCE_NUMBER = 0
    
    CLIENT_HELLO = Chr(1) & _
                    Chr(0) & Chr(2) & _
                    Chr(0) & Chr(3) & _
                    Chr(0) & Chr(0) & _
                    Chr(0) & Chr(Len(CHALLENGE_DATA)) & _
                    Chr(1) & Chr(0) & Chr(128) & _
                    CHALLENGE_DATA
                    
    If Socket.State = 7 Then Socket.SendData AddRecordHeader(CLIENT_HELLO)
End Sub

Private Sub SendMasterKey(ByRef Socket As Winsock)

    iLayer = 1
    
    Call GenerateRandomBytes(32, MASTER_KEY)

    Call CertToPrivateKey

    Socket.SendData AddRecordHeader(Chr(2) & _
                                    Chr(1) & Chr(0) & Chr(128) & _
                                    Chr(0) & Chr(0) & _
                                    Chr(0) & Chr(128) & _
                                    Chr(0) & Chr(0) & _
                                    ExportKeyBlob(MASTER_KEY, CLIENT_READ_KEY, CLIENT_WRITE_KEY, CHALLENGE_DATA, CONNECTION_ID, Private_KEY))
End Sub

Private Sub SendClientFinish(ByRef Socket As Winsock)
    'TODO: Make sure this isnt crippling my connection
    'reference to botmaker server
    iLayer = 2
    Call SSLSend(Socket, Chr(3) & CONNECTION_ID)
End Sub

Private Sub SSLSend(ByRef Socket As Winsock, ByVal Plaintext As String)

    Dim SSLRecord As String
    Dim OtherPart As String
    Dim SendAnother As Boolean
    
    If Len(Plaintext) > 32751 Then
        SendAnother = True
        Plaintext = Mid(Plaintext, 1, 32751)
        OtherPart = Mid(Plaintext, 32752)
    Else
        SendAnother = False
    End If
    
    SSLRecord = AddMACData(Plaintext)
    SSLRecord = RC4_Encrypt(SSLRecord)
    SSLRecord = AddRecordHeader(SSLRecord)
    
    Socket.SendData SSLRecord
    
    If SendAnother = True Then
        Call SSLSend(Socket, OtherPart)
    End If
End Sub

Private Function AddMACData(ByVal Plaintext As String) As String

    AddMACData = MD5_Hash(CLIENT_WRITE_KEY & Plaintext & SendSequence) & Plaintext
End Function

Private Function AddRecordHeader(ByVal RecordData As String) As String

    Dim FirstChar As String
    Dim LastChar As String
    Dim TheLen As Long
        
    TheLen = Len(RecordData)
    
    FirstChar = Chr(128 + (TheLen \ 256))
    LastChar = Chr(TheLen Mod 256)

    AddRecordHeader = FirstChar & LastChar & RecordData
    
    Call IncrementSend
End Function

Private Sub IncrementSend()

    SEND_SEQUENCE_NUMBER = SEND_SEQUENCE_NUMBER + 1
    If SEND_SEQUENCE_NUMBER = 4294967296# Then SEND_SEQUENCE_NUMBER = 0
End Sub

Private Sub IncrementRecv()

    RECV_SEQUENCE_NUMBER = RECV_SEQUENCE_NUMBER + 1
    If RECV_SEQUENCE_NUMBER = 4294967296# Then RECV_SEQUENCE_NUMBER = 0
End Sub

Private Function BytesToLen(ByVal TwoBytes As String) As Long

    Dim FirstByteVal As Long
    FirstByteVal = Asc(Left(TwoBytes, 1))
    If FirstByteVal >= 128 Then FirstByteVal = FirstByteVal - 128
    
    BytesToLen = 256 * FirstByteVal + Asc(Right(TwoBytes, 1))
End Function

Private Function HexToBin(ByVal HexString As String) As String

    Dim BinString As String
    Dim i As Integer
    For i = 1 To Len(HexString) Step 2
        BinString = BinString & Chr(Val("&H" & Mid(HexString, i, 2)))
    Next i
    HexToBin = BinString
End Function

Private Sub socket_Close()

    iLayer = 0
    
    bRaisedConnected = False
End Sub

Private Sub socket_Connect()
    If ProxyState = 1 Then
        Timer1.Enabled = False
        Socket.SendData "CONNECT " & m_RemoteHost & ":" & m_RemotePort & " HTTP/1.1" & vbCrLf & vbCrLf
        Timer1.Enabled = True
    Else
        Call SendClientHello(Socket)
    End If
End Sub

Private Sub socket_DataArrival(ByVal bytesTotal As Long)

    Dim sData As String
    Dim lReachLen As Long
    
If ProxyState = 1 Then
    Timer1.Enabled = False
        Socket.GetData sData
        
        If Mid(sData, 10, 3) = "200" Then
            ProxyState = 0
            Call SendClientHello(Socket)
        Else
            RaiseEvent SSLProxyError
            RaiseEvent SSLClose
            Socket.Close
        End If
Else
    Do
    
        If iSeekLen = 0 Then
            If bytesTotal >= 2 Then
                Socket.GetData sData, vbString, 2
                iSeekLen = BytesToLen(sData)
                bytesTotal = bytesTotal - 2
            Else
                Exit Sub
            End If
        End If
        
        If bytesTotal >= iSeekLen Then
            Socket.GetData sData, vbString, iSeekLen
            bytesTotal = bytesTotal - iSeekLen
        Else
            Exit Sub
        End If
        
        
        Select Case iLayer
            Case 0:
                ENCODED_CERT = Mid(sData, 12, BytesToLen(Mid(sData, 6, 2)))
                CONNECTION_ID = Right(sData, BytesToLen(Mid(sData, 10, 2)))
                Call IncrementRecv
                Call SendMasterKey(Socket)
            Case 1:
                sData = RC4_Decrypt(sData)
                If Right(sData, Len(CHALLENGE_DATA)) = CHALLENGE_DATA Then
                    If VerifyMAC(sData) Then Call SendClientFinish(Socket)
                Else
                    Socket.Close
                End If
             Case 2:
                sData = RC4_Decrypt(sData)
                If VerifyMAC(sData) = False Then Socket.Close
                iLayer = 3
             Case 3:
                sData = RC4_Decrypt(sData)
                If VerifyMAC(sData) Then
                sBuffer = sBuffer & Mid(sData, 17)
                End If
        End Select
    
        iSeekLen = 0

    Loop Until bytesTotal = 0
    
    If iLayer = 3 Then
        If bRaisedConnected = False Then
            bRaisedConnected = True
            RaiseEvent SSLConnect
        End If
    End If
    
    If sBuffer <> "" Then
    
        RaiseEvent SSLData(sBuffer)
    End If
End If
End Sub

Private Function ReverseString(ByVal TheString As String) As String

    Dim Reversed As String
    Dim i As Integer
    For i = Len(TheString) To 1 Step -1
        Reversed = Reversed & Mid(TheString, i, 1)
    Next i
    ReverseString = Reversed
End Function

Public Function SendSSL(sData As String)

    If Socket.State = sckConnected Then
        RaiseEvent SSLSendingData(sData)
        Call SSLSend(Socket, sData)
    Else
        Err.Raise "40006", , "Wrong protocol or connection state for the requested transaction or request."
    End If
End Function

Private Sub socket_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    RaiseEvent SSLError(Number, Description)
End Sub

Private Sub Timer1_Timer()
    Socket.Close
    RaiseEvent SSLProxyError
    RaiseEvent SSLClose
    Timer1.Enabled = False
End Sub

Private Sub UserControl_Initialize()
    Dim lngReturnValue As Long
    Dim TheAnswer As Long
    
    lngReturnValue = CryptAcquireContext(hCryptProv, KEY_CONTAINER, SERVICE_PROVIDER, PROV_RSA_FULL, CRYPT_NEWKEYSET)
    
    If lngReturnValue = 0 Then
        lngReturnValue = CryptAcquireContext(hCryptProv, KEY_CONTAINER, SERVICE_PROVIDER, PROV_RSA_FULL, 0)
        If lngReturnValue = 0 Then TheAnswer = MsgBox("SSL control has detected that you do not have the required High Encryption Pack installed." & vbCrLf & "Would like to download this pack from Microsoft's website?", 16 + vbYesNo)
    End If
    
    If TheAnswer = vbYes Then
        Call Shell("START http://www.microsoft.com/windows/ie/downloads/recommended/128bit/default.asp", vbHide)
        'Add: Close The Socket
    End If
    
    If TheAnswer = vbNo Then
        'Add: Close The Socket
    End If
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = 420
    UserControl.Width = 420
End Sub

Private Sub UserControl_Terminate()
    If hClientWriteKey <> 0 Then CryptDestroyKey hClientWriteKey
    If hClientReadKey <> 0 Then CryptDestroyKey hClientReadKey
    If hCryptProv <> 0 Then CryptReleaseContext hCryptProv, 0
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_RemoteHost = PropBag.ReadProperty("RemoteHost", "")
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("RemoteHost", m_RemoteHost, "")
End Sub
