Attribute VB_Name = "modSQL"
Option Explicit
Public cnME As ADODB.Connection
Public cmdME As ADODB.Command
Public rsPorogZero As New ADODB.Recordset
Public Type Tags_T
    Nam() As String
    Descr() As String
    Arhiv() As String * 1
    Shcala() As Single
    Ma() As Single
    Mi() As Single
End Type
Public Tags As Tags_T
'----------------------------
'ŒÔÂ‰ÂÎÂÌËÂ: ÂÒÚ¸ ÎË Ù‡ÈÎ ?
'----------------------------
Public Type OpenFileProperties
        cBytes As Byte
        fFixedDisk As Byte
        nErrCode As Integer
        Reserved1 As Integer
        Reserved2 As Integer
        szPathName(128) As Byte
End Type
Public Type OVERLAPPED
        Internal As Long
        InternalHigh As Long
        offset As Long
        OffsetHigh As Long
        hEvent As Long
End Type
Public Declare Function OpenFile Lib "Kernel32.dll" ( _
        ByVal lpFileName As String, _
        lpReOpenBuff As OpenFileProperties, _
        ByVal wStyle As Long) As Long
Public Declare Function ReadFile Lib "kernel32" ( _
        ByVal hFile As Long, _
        lpBuffer As Any, _
        ByVal nNumberOfBytesToRead As Long, _
        lpNumberOfBytesRead As Long, _
        lpOverlapped As OVERLAPPED) As Long

Public Function ConnectingToMESQL() As Boolean
On Error GoTo ERR2147467259
    If Not cmdME Is Nothing Then Set cmdME = Nothing
    If cnME Is Nothing Then Set cnME = New ADODB.Connection
    If rsPorogZero.state = adStateOpen Then rsPorogZero.Close '16.05.2014
    If cnME.state = adStateOpen Then cnME.Close: Set cnME = Nothing
    Set cnME = New ADODB.Connection
    cnME.ConnectionTimeout = 15
    cnME.CommandTimeout = 10
    cnME.Open "Provider=SQLOLEDB.1;UID=SA;PWD=48;APP=WS_GRAFIT;Data Source=Pacient;"
    If cnME.state = adStateOpen Then
        cnME.CursorLocation = adUseClient
        ConnectingToMESQL = True
        cnME.DefaultDatabase = "Kotelnaya"
        '=========================================================
        '   «¿√–”«»“‹ ƒ¿ÕÕ€≈ œŒ œŒ–Œ√¿Ã Õ≈◊”¬—“¬»“≈À‹ÕŒ—“»
        '=========================================================
        rsPorogZero.Open "SELECT Nam, PorogZero from PorogZero", cnME, adOpenStatic, adLockReadOnly
        
        GoTo ext
    End If
ERR2147467259:
    Set cnME = Nothing
ext:
    On Error GoTo 0
End Function

Public Function IsFileExists(ByVal ExistsFileName As String) As Boolean
Dim OpenFP As OpenFileProperties
Dim temp As Long
    temp = OpenFile(ExistsFileName, OpenFP, &H4000)
    If temp = 1 Then
        IsFileExists = True
    Else
        IsFileExists = False
    End If
End Function


Public Function GetIdentityID() As Long
Dim i As Long
Dim J As Long
Dim s As Long
        i = DateDiff("d", "01.01.2012", Date)
       J = (Hour(Time) * 3600) + (Minute(Time) * 60) + Second(Time)
       GetIdentityID = i * 10 ^ 5 + J
End Function

