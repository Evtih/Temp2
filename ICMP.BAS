Attribute VB_Name = "ICMP"
Option Explicit
Public IPHostName As String
Public RezultPing As String
' WSock32 UDTs

Type Inet_address
    Byte4 As String * 1
    Byte3 As String * 1
    Byte2 As String * 1
    Byte1 As String * 1
End Type

Public IPLong As Inet_address

Type WSAdata
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To 255) As Byte
    szSystemStatus(0 To 128) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Type Hostent
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type

Type IP_OPTION_INFORMATION
    TTL As Byte                   ' Time to Live (used for traceroute)
    Tos As Byte                   ' Type of Service (usually 0)
    Flags As Byte                 ' IP header Flags (usually 0)
    OptionsSize As Long           ' Size of Options data (usually 0, max 40)
    OptionsData As String * 128   ' Options data buffer
End Type

Public pIPo As IP_OPTION_INFORMATION

Type IP_ECHO_REPLY
    Address(0 To 3) As Byte           ' Replying Address
    Status As Long                    ' Reply Status
    RoundTripTime As Long             ' Round Trip Time in milliseconds
    DataSize As Integer               ' reply data size
    Reserved As Integer               ' for system use
    data As Long                      ' pointer to echo data
    Options As IP_OPTION_INFORMATION  ' Reply Options
End Type

Public pIPe As IP_ECHO_REPLY

' WSock32 Subroutines and Functions
Declare Function GetHostName Lib "wsock32.dll" Alias "gethostname" (ByVal hostname$, HostLen&) As Long
Declare Function GetHostByName& Lib "wsock32.dll" Alias "gethostbyname" (ByVal hostname$)
Declare Function WSAGetLastError Lib "wsock32.dll" () As Long
Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequired&, lpWSAData As WSAdata) As Long
Declare Function WSACleanup Lib "wsock32.dll" () As Long

' Kernel32 Subroutines and Functions
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

' ICMP Subroutines and Functions
    ' IcmpCreateFile will return a file handle
Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
     
     ' Pass the handle value from IcmpCreateFile to the IcmpCloseHandle.  It will return
     ' a boolean value indicating whether or not it closed successfully.
Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal HANDLE As Long) As Boolean
    
    ' IcmpHandle returned from IcmpCreateFile
    ' DestAddress is a pointer to the first entry in the hostent.h_addr_list
    ' RequestData is a null-terminated 64-byte string filled with ASCII 170 characters
    ' RequestSize is 64-bytes
    ' RequestOptions is a NULL at this time
    ' ReplyBuffer
    ' ReplySize
    ' Timeout is the timeout in milliseconds

Declare Function IcmpSendEcho Lib "ICMP" (ByVal IcmpHandle As Long, ByVal DestAddress As Long, _
    ByVal RequestData As String, ByVal RequestSize As Integer, RequestOptns As IP_OPTION_INFORMATION, _
     ReplyBuffer As IP_ECHO_REPLY, ByVal ReplySize As Long, ByVal TimeOut As Long) As Boolean
' WSock32 Variables
Dim iReturn As Long
Dim sLowByte As String
Dim sHighByte As String
Dim sMsg As String
Dim HostLen As Long
Public Host As String
Dim Hostent As Hostent
Dim PointerToPointer As Long
Dim ListAddress As Long
Dim WSAdata As WSAdata
Dim DotA As Long
Dim DotAddr As String
Dim ListAddr As Long
Dim MaxUDP As Long
Dim MaxSockets As Long
Dim Description As String
Dim Status As String

' ICMP Variables
Dim bReturn As Boolean, hIP As Long
Dim szBuffer As String
Dim Addr As Long
Dim RCode As String
Dim RespondingHost As String

' TRACERT Variables
Dim TraceRT As Boolean
Dim TTL As Integer

' WSock32 Constants
Const WS_VERSION_MAJOR = &H101 \ &H100 And &HFF&
Const WS_VERSION_MINOR = &H101 And &HFF&
Const MIN_SOCKETS_REQD = 0

'===========================================================
'               DECLARE
'===========================================================
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Declare Function SetWindowPos Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long) As Long
Public ItemNumb As Long
Public SName_Net As String
Public SName As String
Public ItemIDs() As String
Public MyOPCServerConnect As Boolean
'=========================================================
'
'=========================================================
Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" ( _
    ByVal hwnd As Long, _
    ByVal lpText As String, _
    ByVal lpCaption As String, _
    ByVal wType As Long) As Long
Public Const MB_OK = &H0&
Public Const MB_OKCANCEL = &H1&
Public Const MB_YESNO = &H4&
Public Const MB_YESNOCANCEL = &H3&

Public Const ARM1 = "ARM1"
Public Const ARM2 = "ARM2"
'=========================================================
'      OPCServerState
'=========================================================
Public Const OPCRunning = 1
Public Const OPCFailed = 2
Public Const OPCNoconfig = 3
Public Const OPCSuspended = 4
Public Const OPCTest = 5
Public Const OPCDisconnected = 6
Public OPCServState(1 To 6) As String

'=========================================================
'           ОБЪЯВЛЕНИЯ ДЛЯ OPC SERVER
'=========================================================
Public CurrentNode As Long
Public CurrentVar As Long
Public Sel As Long
Public Type PVarInfo
    Name  As String
    Index(0 To 255) As Long
End Type
Public pv As PVarInfo
Private j As Long
'------------------------------------------
Public groupsCollectionObj As OPCGroups  'коллеция OPC-групп
Public opcGroupObj As OPCGroup            'отдельная OPC-группа
Public itemsCollectionObj As OPCItems    'коллеция элементов данных
Public it As OPCItem
Public BR As OPCBrowser
  'Массив идентификаторов сервера
Public ServerHandles() As Long
  'Массив ошибок, заполняемый в процессе работы с объектами доступа
Public Errors() As Long
  'Массив, в котором будут находиться считанные значения OPC-элементов данных
Public ReadValue() As Variant
  
  'Массив идентификаторов клиента
Public ClientHandles() As Long
  'Массив начальных значений OPC-элементов данных
  'которые определяют начальные условия процесса
Public WriteValues() As Variant
  'Массив путей доступа к элементам данных
Public AccessPaths() As String
  'Массив флагов активности элементов данных
Public Active() As Boolean
  
  'Признак необходимости установки правильных начальных условий
Public writeRequired As Boolean
  'Массив признаков качества сигналов (OPC Quality)
Public pQuality As Variant
  'Массив меток времени сигналов (OPC Timestamp)
Public pTimestamp As Variant
  'Количество элементов данных
Public ItemCount As Integer
  'Счетчик, просто счетчик...
Public i As Integer
  'Количество элементов данных
Public MaxItems As Integer
  'Признак отсутствия каких-либо элементов данных
Public bInvalidItems As Boolean

Public asyncRead As Boolean
Public TagsCount As Long
   'Массив имен элементов данных
'Public ItemIDs() As String
Public ServerName As String
Public Segment As String 'ИМЯ АКТИВНОГО СЕГМЕНТА, К КОТОРОМУ ВЫП-НО ПОДКЛЮЧЕНИЕ
Public nod As String 'ИМЯ АКТИВНОГО СЕГМЕНТА, К КОТОРОМУ ВЫП-НО ПОДКЛЮЧЕНИЕ
Public lstSegments() As String
Public lstNodes() As String
Public retValTag As Variant
Public ConnectToOPCSer As Long


'===========================================================
Public Function PingHost() As Boolean
    RezultPing = ""
    Call vbWSAStartup               ' Initialize Winsock
    Call vbGetHostByName            ' ПОЛУЧИТЬ IPAddress ДЛЯ ПК
    If IPHostName = "" Then
        RezultPing = " нет в домене"
        GoTo EXT
    End If
    Call vbIcmpCreateFile           ' ПОЛУЧИТЬ ICMP ДИСКРИПТОР
    ' НИЖЕПРИВЕДЕННЫЕ ОПРЕДЛЕНИЯ TTL ДЛЯ ICMPEcho
    pIPo.TTL = 255
    Call vbIcmpSendEcho             ' ПЕРЕДАТЬ ICMP Echo Request
    Call vbIcmpCloseHandle          ' Закрыть ICMP ДИСКРИПТОР
    Call vbWSACleanup               ' закрыть Winsock
    If RezultPing = "" Then PingHost = True
EXT:
    
End Function
'--------------- FOR APP
Sub vbGetHostByName()
    Dim szString As String

    szString = String(64, &H0)
    Host = Host + Right$(szString, 64 - Len(Host))

    'не могу найти SOCKET_ERROR
    'If gethostbyname(Host) = SOCKET_ERROR Then              ' If WSock32 error, then tell me about it
    '    sMsg = "Winsock ОШИБКАr" & Str$(WSAGetLastError())
    '    MsgBox sMsg, vbOKOnly, "VB4032-ICMPEcho"
    'Else
        PointerToPointer = GetHostByName(Host)              ' Get the pointer to the address of the winsock hostent structure
        If PointerToPointer = 0 Then 'нет в домене хоста с таким именеи
            IPHostName = ""
        Else
            CopyMemory Hostent.h_name, ByVal _
                PointerToPointer, Len(Hostent)                      ' Copy Winsock structure to the VisualBasic structure
            ListAddress = Hostent.h_addr_list                   ' Get the ListAddress of the Address List
            CopyMemory ListAddr, ByVal ListAddress, 4           ' Copy Winsock structure to the VisualBasic structure
            CopyMemory IPLong, ByVal ListAddr, 4                ' Get the first list entry from the Address List
            CopyMemory Addr, ByVal ListAddr, 4
    
            IPHostName = Trim$(CStr(Asc(IPLong.Byte4)) + "." + CStr(Asc(IPLong.Byte3)) _
                + "." + CStr(Asc(IPLong.Byte2)) + "." + CStr(Asc(IPLong.Byte1)))
        End If
    'End If

End Sub

Function GetRCode() As Long
    If pIPe.Status = 0 Then RCode = "Success": GoTo LL1
    If pIPe.Status = 11001 Then RCode = "Buffer too Small"
    If pIPe.Status = 11002 Then RCode = "Dest Network Not Reachable"
    If pIPe.Status = 11003 Then RCode = "Dest Host Not Reachable"
    If pIPe.Status = 11004 Then RCode = "Dest Protocol Not Reachable"
    If pIPe.Status = 11005 Then RCode = "Dest Port Not Reachable"
    If pIPe.Status = 11006 Then RCode = "No Resources Available"
    If pIPe.Status = 11007 Then RCode = "Bad Option"
    If pIPe.Status = 11008 Then RCode = "Hardware Error"
    If pIPe.Status = 11009 Then RCode = "Packet too Big"
    If pIPe.Status = 11010 Then RCode = "Rqst Timed Out"
    If pIPe.Status = 11011 Then RCode = "Bad Request"
    If pIPe.Status = 11012 Then RCode = "Bad Route"
    If pIPe.Status = 11013 Then RCode = "TTL Exprd in Transit"
    If pIPe.Status = 11014 Then RCode = "TTL Exprd Reassemb"
    If pIPe.Status = 11015 Then RCode = "Parameter Problem"
    If pIPe.Status = 11016 Then RCode = "Source Quench"
    If pIPe.Status = 11017 Then RCode = "Option too Big"
    If pIPe.Status = 11018 Then RCode = " Bad Destination"
    If pIPe.Status = 11019 Then RCode = "Address Deleted"
    If pIPe.Status = 11020 Then RCode = "Spec MTU Change"
    If pIPe.Status = 11021 Then RCode = "MTU Change"
    If pIPe.Status = 11022 Then RCode = "Unload"
    If pIPe.Status = 11050 Then RCode = "General Failure"
LL1:
    RCode = RCode + " (" + CStr(pIPe.Status) + ")"

    DoEvents
        If pIPe.Status = 0 Then
            'RezultPing = RezultPing + "Reply " + RespondingHost + ":Bytes=" + Trim$(CStr(pIPe.DataSize)) + " RTT=" + Trim$(CStr(pIPe.RoundTripTime)) + "ms TTL=" + Trim$(CStr(pIPe.Options.TTL)) + Chr$(13) + Chr$(10)
            RezultPing = RezultPing + RespondingHost + ", " & Trim$(CStr(pIPe.RoundTripTime)) + " ms"
        Else
            RezultPing = RezultPing + "  Ответ " + RespondingHost + ": " + RCode ' + Chr$(13) + Chr$(10)
        End If
End Function

Public Function vbIcmpSendEcho() As Boolean
    Dim NbrOfPkts As Integer
    'ОПРЕДЕЛЕНИЕ КОЛИЧЕСТВА СИМВОЛОВ В ПАКЕТЕ
    szBuffer = "Get IPAdress"
    For NbrOfPkts = 1 To 1
        DoEvents
        bReturn = IcmpSendEcho(hIP, Addr, szBuffer, Len(szBuffer), pIPo, pIPe, Len(pIPe) + 8, 100)
        If bReturn Then
            'ЭТО IPAdress
            RespondingHost = CStr(pIPe.Address(0)) + "." + CStr(pIPe.Address(1)) + "." + CStr(pIPe.Address(2)) + "." + CStr(pIPe.Address(3))
            Call GetRCode
            If RCode Like "*0*" Then vbIcmpSendEcho = True
        Else        ' I hate it when this happens.  If I get an ICMP timeout
                    ' during a TRACERT, try again.
            RezultPing = RezultPing + "Request Timeout" ' + Chr$(13) + Chr$(10)
        End If
    Next NbrOfPkts
End Function

Public Sub vbWSAStartup()
    ' Subroutine to Initialize WSock32
    iReturn = WSAStartup(&H101, WSAdata)
    If iReturn <> 0 Then    ' If WSock32 error, then tell me about it
        MsgBox "WSock32.dll is not responding!", vbOKOnly, "VB4032-ICMPEcho"
    End If

    If LoByte(WSAdata.wVersion) < WS_VERSION_MAJOR Or (LoByte(WSAdata.wVersion) = WS_VERSION_MAJOR And HiByte(WSAdata.wVersion) < WS_VERSION_MINOR) Then
        sHighByte = Trim$(Str$(HiByte(WSAdata.wVersion)))
        sLowByte = Trim$(Str$(LoByte(WSAdata.wVersion)))
        
        sMsg = "WinSock Version " & sLowByte & "." & sHighByte
        sMsg = sMsg & " is not supported "
        MsgBox sMsg, vbOKOnly, "VB4032-ICMPEcho"
        End
    End If

    If WSAdata.iMaxSockets < MIN_SOCKETS_REQD Then
        sMsg = "This application requires a minimum of "
        sMsg = sMsg & Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
        MsgBox sMsg, vbOKOnly, "VB4032-ICMPEcho"
        End
    End If
    MaxSockets = WSAdata.iMaxSockets
    '  WSAdata.iMaxSockets is an unsigned short, so we have to convert it to a signed long
    If MaxSockets < 0 Then
        MaxSockets = 65536 + MaxSockets
    End If
    MaxUDP = WSAdata.iMaxUdpDg
    If MaxUDP < 0 Then
        MaxUDP = 65536 + MaxUDP
    End If
    '  Process the Winsock Description information
    Description = ""
    For i = 0 To 0 'WSADESCRIPTION_LEN
        If WSAdata.szDescription(i) = 0 Then Exit For
        Description = Description + Chr$(WSAdata.szDescription(i))
    Next i
    '  Process the Winsock Status information
    Status = ""
    For i = 0 To 0 'WSASYS_STATUS_LEN
        If WSAdata.szSystemStatus(i) = 0 Then Exit For
        Status = Status + Chr$(WSAdata.szSystemStatus(i))
    Next i
End Sub
Function HiByte(ByVal wParam As Integer)
    HiByte = wParam \ &H100 And &HFF&
End Function
Function LoByte(ByVal wParam As Integer)
    LoByte = wParam And &HFF&
End Function
Public Sub vbWSACleanup()
    ' Subroutine to perform WSACleanup
    iReturn = WSACleanup()
    If iReturn <> 0 Then       ' If WSock32 error, then tell me about it.
        sMsg = "WSock32 Error - " & Trim$(Str$(iReturn)) & " occurred in Cleanup"
        MsgBox sMsg, vbOKOnly, "VB4032-ICMPEcho"
        End
    End If
End Sub

Sub vbIcmpCloseHandle()
    bReturn = IcmpCloseHandle(hIP)
    If bReturn = False Then
        'MsgBox "ICMP Closed with Error", vbOKOnly, "VB4032-ICMPEcho"
    End If
End Sub
Sub vbIcmpCreateFile()
    hIP = IcmpCreateFile()
    If hIP = 0 Then
        'MsgBox "Unable to Create File Handle", vbOKOnly, "VBPing32"
    End If
End Sub


