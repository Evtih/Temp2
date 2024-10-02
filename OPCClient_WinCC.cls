VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OPCClientWinCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Arrays starts with base 1
Option Base 1

' Define objects for OPC functionality
Private MyOPCServer As OPCServer
Private WithEvents MyOPCGroup As OPCGroup
Attribute MyOPCGroup.VB_VarHelpID = -1

' Define Global variables in OPCClientWinCC
Private MyClientHandles() As Long
Private MyServerHandles() As Long
Private Errors() As Long

Const OPC_DS_CACHE = 1
'============================================
'   OPC Quality Code Overview
'============================================
''''0x04 = 4 = Config Error
''''0x08 = 8 = Not Connected
''''0x0C = 12 = Device Failure
''''0x10 = 16 = Sensor Failure
''''0x14 = 20 = Last Known
''''0x18 = 24 = Comm Failure
''''0x1C = 28 = Out of Service
''''0x20 = 32 = Initializing
''''
''''0x40 = 64 = Uncertain
''''0x44 = 68 = Last Usable
''''0x50 = 80 = Sensor Calibration
''''0x54 = 84 = EGU Exceeded
''''0x58 = 88 = Sub Normal
''''
''''0xC0 = 192 = Good
''''0xD8 = 216 = Local Override
Const OPC_QUALITY_GOOD = &HC0
Const OPC_QUALITY_BAD = &H0
Const OPC_QUALITY_UNCERTAIN = &H40

'*************************************************
'       Sub Connect OPC Server
'*************************************************
Public Function state() As Long
Exit Function
    On Error GoTo ErrorHandler
    state = MyOPCServer.state
    On Error GoTo 0
    Exit Function
ErrorHandler:
    MsgBox " ", vbCritical, "VBOPCClient"
    On Error GoTo 0
End Function


'*************************************************
'       Sub Connect OPC Server
'*************************************************
Public Sub connect(Net_Name As String, ServerName As String)
    On Error GoTo ErrorHandler
    
    ' Create new OPC Server Object
    Set MyOPCServer = New OPCServer
    
    ' Call Sub Connect OPC Server
    MyOPCServer.connect ServerName, Net_Name
    MyOPCServerConnect = True
    On Error GoTo 0
    Exit Sub
ErrorHandler:
    MyOPCServerConnect = False
    'MsgBox "Œÿ»¡ ¿ œŒƒ Àﬁ◊≈Õ»ﬂ   OPCServer.WinCC ", vbCritical, "VBOPCClient"
    On Error GoTo 0
End Sub

'*************************************************
'       Sub Disconnect OPC Server
'*************************************************
Public Sub Disconnect()
    On Error GoTo ErrorHandler
    
    ' Call Sub Disconnect OPC Server
    MyOPCServer.Disconnect
    
    ' Release OPC Server Object
    Set MyOPCServer = Nothing
    MyOPCServerConnect = False
    Exit Sub
ErrorHandler:
    MsgBox "Œÿ»¡ ¿ Œ“ Àﬁ◊≈Õ»ﬂ Œ“ OPCServer.WinCC", vbCritical, "VBOPCClient"
End Sub

'*************************************************
'       Sub Add OPC Group
'*************************************************
Public Sub AddGroup(GroupNm As String)
    On Error GoTo ErrorHandler
    
    ' Set UpdateRate for groups (for DataChange)
    MyOPCServer.OPCGroups.DefaultGroupUpdateRate = 1000
   
    ' Create group object for a server object
    Set MyOPCGroup = MyOPCServer.OPCGroups.Add(GroupNm)
    
    ' Call Sub Deactivate Callback for OPC Group
    Call DeactivateCallback
    On Error GoTo 0
    Exit Sub
ErrorHandler:
    MsgBox "Add OPC Group failed!", vbCritical, "VBOPCClient"
    On Error GoTo 0
End Sub

'*************************************************
'           Remove OPC Group
'*************************************************
Public Sub RemGroup(GroupNm As String)
    On Error GoTo ErrorHandler
    ' Remove OPC Group from OPC Server Object
    MyOPCServer.OPCGroups.Remove GroupNm
    ' Release OPC Group Object
    Set MyOPCGroup = Nothing
    On Error GoTo 0
    Exit Sub
ErrorHandler:
    MsgBox "Remove OPC Group failed!", vbCritical, "VBOPCClient"
    On Error GoTo 0
End Sub

''*************************************************
'       Sub Add OPC Items to OPC Group
'*************************************************
Public Sub AddItems(ItemIDs() As String, ItemNumb As Long)
Dim i As Long
Dim bInvalidItems As Boolean
ReDim MyClientHandles(ItemNumb)
    On Error GoTo ErrorHandler
    
    ' Set Client Handles for each OPC Item
    For i = 1 To ItemNumb
        MyClientHandles(i) = i
    Next i
              
    ' Call Sub Add Items (subordinate group)
    MyOPCGroup.OPCItems.AddItems ItemNumb, ItemIDs, MyClientHandles, MyServerHandles, Errors
  
  For i = LBound(Errors) To UBound(Errors)
    If Errors(i) <> 0 Then
        bInvalidItems = True
        'Exit For
    Else
        bInvalidItems = False
    End If
    
  Next i
    
    ' Call Sub Activate Callback (DataChange) for OPC Group
    ActivateCallback
    
    ' Release OPC Error Array
    Erase Errors
    On Error GoTo 0
    Exit Sub
ErrorHandler:
    MsgBox "Œÿ»¡ ¿ ƒŒ¡¿¬À≈Õ»ﬂ “›√¿", vbCritical, "VBOPCClient"
    On Error GoTo 0
End Sub

'*************************************************
'        Sub Remove OPC Items
'*************************************************
Public Sub RemItems(ItemNumb As Long)
Dim i As Long
    On Error GoTo ErrorHandler
    
    ' Delete Column "values" in Excel sheet1
    With frmOPCClient.flxOPCTags
        For i = 1 To ItemNumb
            .TextMatrix(i, 2) = ""
            .TextMatrix(i, 3) = ""
            .TextMatrix(i, 4) = ""
        Next i
        
        ' Delete Column "Writing values" in Excel sheet1
        'For i = 1 To 4
        '    Cells(6 + i, 9) = ""
        'Next i
    End With
    ' Delete items from group object: not necessary used because remove group object delete item objects
    MyOPCGroup.OPCItems.Remove ItemNumb, MyServerHandles, Errors
    
    ' Release Arrays
    Erase MyClientHandles
    Erase MyServerHandles
    Erase Errors
    On Error GoTo 0
    Exit Sub
ErrorHandler:
    MsgBox "Œÿ»¡ ¿ ”ƒ¿À≈Õ»ﬂ “›√¿", vbCritical, "VBOPCClient"
    On Error GoTo 0
End Sub

'*************************************************
'       Read OPC Items synchron
'*************************************************
Public Sub ReadSync()
Dim i As Long
Dim Values() As Variant
Dim Qualities()  As Integer
Dim TimeStamps() As Date
Dim s As String
' Local Server Handles Array
Dim SHandles() As Long
' Reallocate Array of local Server Handles
ReDim SHandles(ItemNumb)
    On Error GoTo ErrorHandler
    
    '«¿œŒÃÕ»“‹ ƒ»— –»œ“Œ–€ “›√Œ¬ (MyServerHandles) ¬ ÀŒ ¿À‹ÕŒÃ Ã¿——»¬≈
    For i = 1 To ItemNumb
        SHandles(i) = MyServerHandles(i)
    Next i
    
    ' Read Values synchron from OPC Server
    MyOPCGroup.SyncRead OPCCache, ItemNumb, SHandles, Values, Errors ', Qualities, TimeStamps
    
    For i = 1 To ItemNumb
        If Errors(i) <> 0 Then
           MsgBox "Œÿ»¡ ¿ ¬€œŒÀÕ≈Õ»ﬂ 'SyncRead' “›√¿ '" & ItemIDs(i) & "'", vbCritical, "VBOPCClient"
        End If
    Next i
    ' Write Values, TimeStamps and Qualities to Excel-Sheet
    With frmOPCClient.flxOPCTags
        For i = 1 To ItemNumb
            .TextMatrix(MyClientHandles(i), 1) = Values(i)
            Select Case Qualities(i)
                Case OPC_QUALITY_GOOD
                    .TextMatrix(MyClientHandles(i), 2) = "ıÓÓ¯ÂÂ"
                Case OPC_QUALITY_BAD
                    .TextMatrix(MyClientHandles(i), 2) = "ÔÎÓıÓÂ"
                Case OPC_QUALITY_UNCERTAIN
                    .TextMatrix(MyClientHandles(i), 2) = "ÌÂÓÔÂ‰ÂÎÂÌÓ"
                Case Else
                    .TextMatrix(MyClientHandles(i), 2) = Hex(Qualities(i))
            End Select
            .TextMatrix(MyClientHandles(i), 3) = TimeStamps(i)
            '23.01.2012 s = s & "~" & ItemIDs(i) _
                  & "~" & Values(i) _
                  & "~" & Qualities(i) _
                  & "~" & TimeStamps(i)
        Next i
    End With
    Erase Values
    Erase Errors
    Erase Qualities
    Erase TimeStamps
    '23.01.2012 s = "Req~" & Now & "~" & s & "||"
    'Open "Z:\Kotel.opc" For Output As #1
    '    Print #1, s
    'Close #1
    frmOPCClient.txtReadSync.Text = Now & " ReadSync: ¬˚ÔÓÎÌÂÌÓ"
    On Error GoTo 0
    Exit Sub
ErrorHandler:
    'MsgBox "Œÿ»¡ ¿ ¬€œŒÀÕ≈Õ»ﬂ 'SyncRead'", vbCritical, "ERROR"
    frmOPCClient.txtReadSync.Text = Now & " ReadSync: Œÿ»¡ ¿ ¬€œŒÀÕ≈Õ»ﬂ "
    On Error GoTo 0
    'Close #1
End Sub

''''''''*************************************************
''''''''       Sub Write OPC Items synchron
''''''''*************************************************
'''''''Public Sub SyncWrite()
'''''''Dim i As Long
'''''''Dim Values()
'''''''Dim SHandles() As Long
'''''''    ' Reallocate Arrays
'''''''    ReDim Values(4)
'''''''    ReDim SHandles(4)
'''''''    On Error GoTo ErrorHandler
'''''''
'''''''    ' loop for writing values
'''''''    For i = 1 To 4
'''''''        ' Read Values from column "Writing values" in Excel sheet1
'''''''        'Values(i) = Cells(6 + i, 9)
'''''''        If Values(i) = "" Then Values(i) = 0
'''''''        ' Get MyServerHandles assign to local SHandles array
'''''''        SHandles(i) = MyServerHandles(i + 2)
'''''''    Next i
'''''''
'''''''    ' Call Sub Write Item Values synchron to OPC Server
'''''''    MyOPCGroup.SyncWrite 4, SHandles, Values, Errors
'''''''    For i = 1 To 4
'''''''        If Errors(i) <> 0 Then
'''''''           MsgBox "Error SyncWrite OPC Item", vbCritical, "ERROR"
'''''''        End If
'''''''    Next i
'''''''
'''''''    ' Release Arrays
'''''''    Erase Errors
'''''''    Erase SHandles
'''''''    Erase Values
'''''''
'''''''    Exit Sub
'''''''ErrorHandler:
'''''''    MsgBox "Write OPC Items Failed", vbCritical, "ERROR"
'''''''End Sub

'*************************************************
'
'*************************************************
Private Sub MyOPCGroup_DataChange(ByVal TransactionID As Long, ByVal NumItems As Long, MyClientHandles() As Long, ItemValues() As Variant, Qualities() As Long, TimeStamps() As Date)
Dim i As Integer
    With frmOPCClient.flxOPCTags
        For i = 1 To NumItems
            .TextMatrix(MyClientHandles(i), 1) = ItemValues(i)
            Select Case Qualities(i)
                Case OPC_QUALITY_GOOD
                    .TextMatrix(MyClientHandles(i), 2) = "GOOD"
                Case OPC_QUALITY_BAD
                    .TextMatrix(MyClientHandles(i), 2) = "BAD"
                Case OPC_QUALITY_UNCERTAIN
                    .TextMatrix(MyClientHandles(i), 2) = "UNCERTAIN"
                Case Else
                    .TextMatrix(MyClientHandles(i), 2) = Hex(Qualities(i))
            End Select
            .TextMatrix(MyClientHandles(i), 3) = TimeStamps(i)
        Next i
    End With
End Sub

'*************************************************
'
'*************************************************
Private Sub ActivateCallback()
    ' Activate Group for DataChange
    MyOPCGroup.IsSubscribed = True
    ' Set Group to Active State
    MyOPCGroup.IsActive = True
End Sub

'*************************************************
'
'*************************************************
Private Sub DeactivateCallback()
    ' Deactivate Group Collection
    MyOPCServer.OPCGroups.DefaultGroupIsActive = False
     ' Group not receives DataChanges from OPC Server
    MyOPCGroup.IsSubscribed = False
End Sub


