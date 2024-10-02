VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmOPCClientBrowser 
   Caption         =   "Form1"
   ClientHeight    =   8736
   ClientLeft      =   132
   ClientTop       =   516
   ClientWidth     =   9408
   LinkTopic       =   "Form1"
   ScaleHeight     =   8736
   ScaleWidth      =   9408
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstI 
      Columns         =   1
      Height          =   9648
      Left            =   4500
      TabIndex        =   1
      Top             =   0
      Width           =   2295
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxTags 
      Height          =   9735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      _ExtentX        =   7641
      _ExtentY        =   17166
      _Version        =   393216
      Cols            =   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
End
Attribute VB_Name = "frmOPCClientBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents opcServerObj As OPCServer        'OPC-������
Attribute opcServerObj.VB_VarHelpID = -1

Private Sub Form_Load()
Dim sTemp As String
    Me.Left = 0: Me.Top = 0:
    Me.Height = 300
    Me.flxOPCTags.ColWidth(0) = 500
    Me.flxOPCTags.ColWidth(1) = 1500: flxOPCTags.ColWidth(2) = 1500
    Call ConnectToOPCServer
    If ConnectToOPCSer > 0 Then
      sTemp = "������ ��������� ���������� - "
      Select Case ConnectToOPCSer
        Case 1 '�� ������� ��� sServer
           i = MessageBox(Me.hwnd, sTemp & "�� ������� ��� Server", "������ OPC Server", MB_OK)
        Case 2 '� ������������ OPC SERVER ��� ���������
           i = MessageBox(Me.hwnd, sTemp & "� ������������ OPC SERVER ��� ���������", "����������� ���������", MB_OK)
        Case 3 '� ������������ �������� OPC SERVER ��� �����������
           i = MessageBox(Me.hwnd, sTemp & "� ������������ �������� OPC SERVER ��� �����������", "����������� ����������� � ��������", MB_OK)
        Case 4 '� ������������ OPC SERVER ��� �����
           i = MessageBox(Me.hwnd, sTemp & "� ������������ OPC SERVER ��� �����", "����������� �����", MB_OK)
      End Select
      Unload Me
      End
    End If
    opcGroupObj.SyncRead OPCCache, ItemCount, ServerHandles, ReadValue, Errors

End Sub

Public Function ConnectToOPCServer() As Long
Dim Seg As Long '������� ���������
Dim NodCount As Long '������� ����� � ��������
Dim TagS As Long '������� ����� � ����
    OPCServState(1) = "OPCRunning": OPCServState(2) = "OPCFailed"
    OPCServState(3) = "OPCNoconfig": OPCServState(4) = "OPCSuspended"
    OPCServState(5) = "OPCtest":  OPCServState(6) = "OPCDisconnected"
On Error GoTo Problem
    SName_Net = "ZVAN1"
    SName = "OPCServer.WinCC"
    Set opcServerObj = New OPCServer
    opcServerObj.connect SName_Net, SName
    If TypeName(opcServerObj) = TypeName(Nothing) Then    'IOPCAutoServer
        'mnuErrSucServer.Caption = "OPC SERVER HE �������"
        ConnectToOPCSer = 1 'no server
        Return
    End If
    '********************BROWSER ������������ �������**********
    Set BR = opcServerObj.CreateBrowser
    'BR.Filter = "*"
    '--------�������� ������ ��������� "UltraNet_�"
    BR.ShowBranches
    If BR.Count = 0 Then
        ConnectToOPCSer = 2 '��� ���������
        Exit Function
    End If
    ReDim lstSegments(BR.Count)
    For Seg = 1 To BR.Count
        lstSegments(Seg) = BR.Item(Seg) '"UltraNet_1",UltraNet_2....
    Next
    '--------�������� ������ ������ ����� ��� ��������
    Segment = BR.Item(1)
    BR.MoveDown Segment '"UltraNet_1"
    BR.ShowBranches '"03","05".......
    If BR.Count = 0 Then
        ConnectToOPCSer = 3 '��� ����� (������������)
        Exit Function
    End If
    ReDim lstNodes(UBound(lstSegments), BR.Count)
    lstNodes(1, 1) = BR.Item(1)
    '--------�������� ������ ����� ��� ������� ����
    nod = BR.Item(1)
    BR.MoveDown nod
    BR.ShowLeafs nod '�������� ������ �����
        If BR.Count = 0 Then
        ConnectToOPCSer = 4 '��� �����
        Exit Function
    End If
    TagsCount = BR.Count
    ReDim ItemIDs(TagsCount + 2)
    flxOPCTags.Rows = TagsCount + 2
    For TagS = 1 To TagsCount
        lstI.AddItem BR.Item(TagS)
        ItemIDs(TagS) = Segment & "." & nod & "." & BR.Item(TagS)
        flxOPCTags.TextMatrix(TagS, 0) = TagS
        flxOPCTags.TextMatrix(TagS, 1) = BR.Item(TagS)
    Next
    Set BR = Nothing
  '*********************************************
  '�������� ������ �� �������
  '��������� ��������� OPC-�����
  Set groupsCollectionObj = opcServerObj.OPCGroups
  '-------------------------------------------------------------------
  '���������� �������� (������) � ��������� OPC-�����
  Set opcGroupObj = groupsCollectionObj.Add("Group One")
  opcGroupObj.ClientHandle = 100
  'Debug.Print opcGroupObj.ServerHandle, opcGroupObj.
  
  '-------------------------------------------------------------------
  '��������� ��������� ��������� ������ � ����������� ������
  Set itemsCollectionObj = opcGroupObj.OPCItems
  '-------------------------------------------------------------------
  '���������� ��������, ������������ ��� ����������
  '��������� ������ � ������
  ReDim ClientHandles(TagsCount + 1)
  ReDim AccessPaths(TagsCount + 1)
  ReDim Active(TagsCount + 1)
  ReDim WriteValues(TagsCount + 1)
  ItemCount = 0
  For i = 1 To TagsCount
      Active(i) = True
      ItemCount = ItemCount + 1
      ClientHandles(ItemCount) = i + 100
      WriteValues(ItemCount) = 0
  Next i
  '-------------------------------------------------------------------
  '���������� ��������� ������ (������) � ������
  itemsCollectionObj.AddItems ItemCount, ItemIDs, ClientHandles, _
                ServerHandles, Errors, pQuality, pTimestamp
  '-------------------------------------------------------------------
  '�������� ������������ ������� ����� �������� ������ �����
  For i = LBound(Errors) To UBound(Errors)
    If Errors(i) <> 0 Then
        bInvalidItems = True
        Exit For
    Else
        bInvalidItems = False
    End If
  Next i
  '-------------------------------------------------------------------
  '������������ ������� ��������� - ��������� ����������� ������ �����
    If bInvalidItems = False Then
      '���������� ������ �������� ��������� ������ � ������ (�����������)
      opcGroupObj.SyncRead OPCCache, ItemCount, ServerHandles, ReadValue, Errors
      '��� ���� ��������� � ��������� ����������� �������
      '��������� �������� ���������� � �������
      'IsConnect = True
      'GetTags = True
      asyncRead = False
      opcGroupObj.IsSubscribed = False
    Else
    
    End If
    On Error GoTo 0
    Exit Function
Problem:
    On Error GoTo 0
End Function


