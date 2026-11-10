Attribute VB_Name = "modWebSocket"
Option Explicit

' ============================================================================
' modWebSocket - a minimal WebSocket client for VBA, using only Windows itself
'
' TECHNIQUE 1 (pure VBA) - the plumbing. Imported prebuilt in the video and
' walked through in chunks.
'
' Everything here calls winhttp.dll, which ships with Windows (8 or later)
' and is the same DLL behind WinHttp.WinHttpRequest.5.1. Windows performs
' the TLS encryption, the WebSocket handshake, the frame headers and the
' masking. We only open, send, receive and close.
'
' Requires 64-bit or 32-bit Office 2010 or later (VBA7).
'
' KNOWN LIMITATION: WsReceiveText blocks the calling thread until a message
' arrives. Windows offers no way to ask "is a message waiting?" on a
' WebSocket handle, and its asynchronous mode delivers results on a worker
' thread, which VBA cannot safely handle. See the video for what that means.
' ============================================================================

' --- winhttp.dll ------------------------------------------------------------

Private Declare PtrSafe Function WinHttpOpen Lib "winhttp" ( _
    ByVal pszAgentW As LongPtr, ByVal dwAccessType As Long, _
    ByVal pszProxyW As LongPtr, ByVal pszProxyBypassW As LongPtr, _
    ByVal dwFlags As Long) As LongPtr

Private Declare PtrSafe Function WinHttpSetTimeouts Lib "winhttp" ( _
    ByVal hInternet As LongPtr, ByVal nResolveTimeout As Long, _
    ByVal nConnectTimeout As Long, ByVal nSendTimeout As Long, _
    ByVal nReceiveTimeout As Long) As Long

Private Declare PtrSafe Function WinHttpConnect Lib "winhttp" ( _
    ByVal hSession As LongPtr, ByVal pswzServerName As LongPtr, _
    ByVal nServerPort As Long, ByVal dwReserved As Long) As LongPtr

Private Declare PtrSafe Function WinHttpOpenRequest Lib "winhttp" ( _
    ByVal hConnect As LongPtr, ByVal pwszVerb As LongPtr, _
    ByVal pwszObjectName As LongPtr, ByVal pwszVersion As LongPtr, _
    ByVal pwszReferrer As LongPtr, ByVal ppwszAcceptTypes As LongPtr, _
    ByVal dwFlags As Long) As LongPtr

Private Declare PtrSafe Function WinHttpSetOption Lib "winhttp" ( _
    ByVal hInternet As LongPtr, ByVal dwOption As Long, _
    ByVal lpBuffer As LongPtr, ByVal dwBufferLength As Long) As Long

Private Declare PtrSafe Function WinHttpSendRequest Lib "winhttp" ( _
    ByVal hRequest As LongPtr, ByVal lpszHeaders As LongPtr, _
    ByVal dwHeadersLength As Long, ByVal lpOptional As LongPtr, _
    ByVal dwOptionalLength As Long, ByVal dwTotalLength As Long, _
    ByVal dwContext As LongPtr) As Long

Private Declare PtrSafe Function WinHttpReceiveResponse Lib "winhttp" ( _
    ByVal hRequest As LongPtr, ByVal lpReserved As LongPtr) As Long

Private Declare PtrSafe Function WinHttpQueryHeaders Lib "winhttp" ( _
    ByVal hRequest As LongPtr, ByVal dwInfoLevel As Long, _
    ByVal pwszName As LongPtr, ByRef lpBuffer As Long, _
    ByRef lpdwBufferLength As Long, ByRef lpdwIndex As Long) As Long

Private Declare PtrSafe Function WinHttpWebSocketCompleteUpgrade Lib "winhttp" ( _
    ByVal hRequest As LongPtr, ByVal pContext As LongPtr) As LongPtr

Private Declare PtrSafe Function WinHttpWebSocketSend Lib "winhttp" ( _
    ByVal hWebSocket As LongPtr, ByVal eBufferType As Long, _
    ByRef pvBuffer As Any, ByVal dwBufferLength As Long) As Long

Private Declare PtrSafe Function WinHttpWebSocketReceive Lib "winhttp" ( _
    ByVal hWebSocket As LongPtr, ByRef pvBuffer As Any, _
    ByVal dwBufferLength As Long, ByRef pdwBytesRead As Long, _
    ByRef peBufferType As Long) As Long

Private Declare PtrSafe Function WinHttpWebSocketClose Lib "winhttp" ( _
    ByVal hWebSocket As LongPtr, ByVal usStatus As Integer, _
    ByVal pvReason As LongPtr, ByVal dwReasonLength As Long) As Long

Private Declare PtrSafe Function WinHttpCloseHandle Lib "winhttp" ( _
    ByVal hInternet As LongPtr) As Long

' --- kernel32.dll (UTF-8 conversion and error codes) ------------------------

Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long

Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, _
    ByVal cchWideChar As Long, ByRef lpMultiByteStr As Any, ByVal cbMultiByte As Long, _
    ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long

Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" ( _
    ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpMultiByteStr As Any, _
    ByVal cbMultiByte As Long, ByVal lpWideCharStr As LongPtr, _
    ByVal cchWideChar As Long) As Long

' --- Constants --------------------------------------------------------------

Private Const WINHTTP_ACCESS_TYPE_DEFAULT_PROXY As Long = 0
Private Const INTERNET_DEFAULT_HTTPS_PORT As Long = 443
Private Const WINHTTP_FLAG_SECURE As Long = &H800000
Private Const WINHTTP_OPTION_UPGRADE_TO_WEB_SOCKET As Long = 114
Private Const WINHTTP_QUERY_STATUS_CODE As Long = 19
Private Const WINHTTP_QUERY_FLAG_NUMBER As Long = &H20000000
Private Const HTTP_STATUS_SWITCHING_PROTOCOLS As Long = 101

Private Const WEB_SOCKET_BINARY_MESSAGE As Long = 0
Private Const WEB_SOCKET_BINARY_FRAGMENT As Long = 1
Private Const WEB_SOCKET_UTF8_MESSAGE As Long = 2
Private Const WEB_SOCKET_UTF8_FRAGMENT As Long = 3
Private Const WEB_SOCKET_CLOSE As Long = 4
Private Const WEB_SOCKET_SUCCESS_CLOSE_STATUS As Integer = 1000

Private Const ERROR_SUCCESS As Long = 0
Private Const CP_UTF8 As Long = 65001
Private Const RECEIVE_BUFFER_SIZE As Long = 8192

' --- Module state -----------------------------------------------------------

Private m_hSession As LongPtr
Private m_hConnect As LongPtr
Private m_hWebSocket As LongPtr
Private m_IsOpen As Boolean
Private m_LastError As Long

' ---------------------------------------------------------------------------
' Open a secure WebSocket. serverName is the host on its own, for example
' "stream.data.alpaca.markets". resourcePath is everything after it, for
' example "/v2/test".
' ---------------------------------------------------------------------------
Public Function WsConnect(ByVal serverName As String, ByVal resourcePath As String) As Boolean
    Dim hRequest As LongPtr
    Dim httpVerb As String
    Dim statusCode As Long
    Dim statusSize As Long

    If m_IsOpen Then WsClose
    m_LastError = 0
    httpVerb = "GET"
    statusSize = 4

    m_hSession = WinHttpOpen(StrPtr("Excel VBA WebSocket"), WINHTTP_ACCESS_TYPE_DEFAULT_PROXY, 0, 0, 0)
    If m_hSession = 0 Then GoTo Failed

    ' Resolve 10s, connect 30s, send 30s, receive 0 = wait forever.
    ' A receive timeout does not mean "no data yet" - it cancels the request.
    WinHttpSetTimeouts m_hSession, 10000, 30000, 30000, 0

    m_hConnect = WinHttpConnect(m_hSession, StrPtr(serverName), INTERNET_DEFAULT_HTTPS_PORT, 0)
    If m_hConnect = 0 Then GoTo Failed

    hRequest = WinHttpOpenRequest(m_hConnect, StrPtr(httpVerb), StrPtr(resourcePath), 0, 0, 0, WINHTTP_FLAG_SECURE)
    If hRequest = 0 Then GoTo Failed

    ' Mark this request as a WebSocket upgrade. No value is passed, only the flag.
    If WinHttpSetOption(hRequest, WINHTTP_OPTION_UPGRADE_TO_WEB_SOCKET, 0, 0) = 0 Then GoTo Failed

    If WinHttpSendRequest(hRequest, 0, 0, 0, 0, 0, 0) = 0 Then GoTo Failed
    If WinHttpReceiveResponse(hRequest, 0) = 0 Then GoTo Failed

    If WinHttpQueryHeaders(hRequest, WINHTTP_QUERY_STATUS_CODE Or WINHTTP_QUERY_FLAG_NUMBER, _
                           0, statusCode, statusSize, 0) = 0 Then GoTo Failed

    If statusCode <> HTTP_STATUS_SWITCHING_PROTOCOLS Then
        m_LastError = statusCode
        GoTo Failed
    End If

    m_hWebSocket = WinHttpWebSocketCompleteUpgrade(hRequest, 0)
    If m_hWebSocket = 0 Then GoTo Failed

    ' From here on, the WebSocket handle replaces the request handle.
    WinHttpCloseHandle hRequest
    hRequest = 0

    m_IsOpen = True
    WsConnect = True
    Exit Function

Failed:
    ' If this shows 0 for a misspelt host, use Err.LastDllError instead.
    If m_LastError = 0 Then m_LastError = GetLastError()
    If hRequest <> 0 Then WinHttpCloseHandle hRequest
    CloseHandles
    WsConnect = False
End Function

' ---------------------------------------------------------------------------
' Send one text message (UTF-8).
' ---------------------------------------------------------------------------
Public Function WsSendText(ByVal messageText As String) As Boolean
    Dim messageBytes() As Byte
    Dim byteCount As Long
    Dim result As Long

    If Not m_IsOpen Then Exit Function

    byteCount = StringToUtf8(messageText, messageBytes)
    If byteCount = 0 Then Exit Function

    result = WinHttpWebSocketSend(m_hWebSocket, WEB_SOCKET_UTF8_MESSAGE, messageBytes(0), byteCount)
    If result <> ERROR_SUCCESS Then
        m_LastError = result
        Exit Function
    End If

    WsSendText = True
End Function

' ---------------------------------------------------------------------------
' Receive one whole text message. THIS BLOCKS until a message arrives, or
' until the connection ends. Returns "" if the connection closed.
' ---------------------------------------------------------------------------
Public Function WsReceiveText() As String
    Dim readBuffer(0 To RECEIVE_BUFFER_SIZE - 1) As Byte
    Dim messageBytes() As Byte
    Dim totalBytes As Long
    Dim bytesRead As Long
    Dim BufferType As Long
    Dim result As Long
    Dim i As Long

    If Not m_IsOpen Then Exit Function

    Do
        result = WinHttpWebSocketReceive(m_hWebSocket, readBuffer(0), RECEIVE_BUFFER_SIZE, bytesRead, BufferType)

        If result <> ERROR_SUCCESS Then
            m_LastError = result
            m_IsOpen = False
            Exit Function
        End If

        If BufferType = WEB_SOCKET_CLOSE Then
            m_IsOpen = False
            Exit Function
        End If

        If bytesRead > 0 Then
            ReDim Preserve messageBytes(0 To totalBytes + bytesRead - 1)
            For i = 0 To bytesRead - 1
                messageBytes(totalBytes + i) = readBuffer(i)
            Next i
            totalBytes = totalBytes + bytesRead
        End If
    Loop While BufferType = WEB_SOCKET_UTF8_FRAGMENT Or BufferType = WEB_SOCKET_BINARY_FRAGMENT

    WsReceiveText = Utf8ToString(messageBytes, totalBytes)
End Function

Public Function WsIsOpen() As Boolean
    WsIsOpen = m_IsOpen
End Function

Public Function WsLastError() As Long
    WsLastError = m_LastError
End Function

' ---------------------------------------------------------------------------
' Send a close frame, then release every handle.
' ---------------------------------------------------------------------------
Public Sub WsClose()
    If m_hWebSocket <> 0 And m_IsOpen Then
        WinHttpWebSocketClose m_hWebSocket, WEB_SOCKET_SUCCESS_CLOSE_STATUS, 0, 0
    End If
    m_IsOpen = False
    CloseHandles
End Sub

Private Sub CloseHandles()
    If m_hWebSocket <> 0 Then WinHttpCloseHandle m_hWebSocket
    If m_hConnect <> 0 Then WinHttpCloseHandle m_hConnect
    If m_hSession <> 0 Then WinHttpCloseHandle m_hSession
    m_hWebSocket = 0
    m_hConnect = 0
    m_hSession = 0
    m_IsOpen = False
End Sub

' --- Text conversion --------------------------------------------------------

Private Function StringToUtf8(ByVal sourceText As String, ByRef outBytes() As Byte) As Long
    Dim byteCount As Long

    byteCount = WideCharToMultiByte(CP_UTF8, 0, StrPtr(sourceText), Len(sourceText), ByVal 0&, 0, 0, 0)
    If byteCount <= 0 Then Exit Function

    ReDim outBytes(0 To byteCount - 1)
    WideCharToMultiByte CP_UTF8, 0, StrPtr(sourceText), Len(sourceText), outBytes(0), byteCount, 0, 0
    StringToUtf8 = byteCount
End Function

Private Function Utf8ToString(ByRef sourceBytes() As Byte, ByVal byteCount As Long) As String
    Dim charCount As Long
    Dim resultText As String

    If byteCount <= 0 Then Exit Function

    charCount = MultiByteToWideChar(CP_UTF8, 0, sourceBytes(0), byteCount, 0, 0)
    If charCount <= 0 Then Exit Function

    resultText = String$(charCount, vbNullChar)
    MultiByteToWideChar CP_UTF8, 0, sourceBytes(0), byteCount, StrPtr(resultText), charCount
    Utf8ToString = resultText
End Function
