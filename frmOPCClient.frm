VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmOPCClient 
   BackColor       =   &H00FFFF00&
   Caption         =   "OPC Client  OPCServer.WinCC"
   ClientHeight    =   11835
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   15765
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   204
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   11835
   ScaleWidth      =   15765
   Begin VB.CommandButton cmdLogOPC 
      Caption         =   "Log OPC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14070
      TabIndex        =   19
      Top             =   150
      Width           =   1635
   End
   Begin VB.ListBox lstLogOPC 
      Height          =   9810
      Left            =   10470
      TabIndex        =   18
      Top             =   1110
      Width           =   10000
   End
   Begin VB.TextBox txtConectCount 
      Height          =   315
      Left            =   2700
      TabIndex        =   15
      Text            =   "0"
      Top             =   30
      Width           =   765
   End
   Begin VB.Timer tmrSaveFile 
      Interval        =   11111
      Left            =   9000
      Top             =   45
   End
   Begin VB.TextBox txtReadSync 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   90
      TabIndex        =   13
      Tag             =   "ВЫПОЛНЯЕТСЯ ПОДКЛЮЧЕНИЕ К OPCServer.WinCC: "
      Top             =   11430
      Width           =   11160
   End
   Begin VB.Timer tmrPing 
      Interval        =   1000
      Left            =   7515
      Top             =   45
   End
   Begin VB.CommandButton cmdPing 
      Caption         =   "Ping Host"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   11400
      TabIndex        =   11
      Top             =   11160
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   12840
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   11130
      Visible         =   0   'False
      Width           =   2490
   End
   Begin VB.TextBox txtWait 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   90
      TabIndex        =   4
      Tag             =   "ВЫПОЛНЯЕТСЯ ПОДКЛЮЧЕНИЕ К OPCServer.WinCC: "
      Top             =   11040
      Width           =   11160
   End
   Begin VB.CommandButton cmdConnect 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   780
      Picture         =   "frmOPCClient.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   45
      Width           =   855
   End
   Begin VB.CommandButton cmbClose 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   45
      Picture         =   "frmOPCClient.frx":0C42
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   45
      Width           =   735
   End
   Begin VB.CommandButton cmdDisconnect 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1635
      Picture         =   "frmOPCClient.frx":21A8
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   45
      Width           =   900
   End
   Begin VB.Timer tmrReadSync 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   7020
      Top             =   45
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxOPCTags 
      Height          =   9855
      Left            =   135
      TabIndex        =   0
      Top             =   1140
      Width           =   15570
      _ExtentX        =   27464
      _ExtentY        =   17383
      _Version        =   393216
      Rows            =   41
      Cols            =   6
      BackColorBkg    =   16777152
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
   End
   Begin VB.Label lblStartApp 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "    "
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   840
      Width           =   180
   End
   Begin VB.Label lblTypeRefresh 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "    "
      Height          =   195
      Left            =   2730
      TabIndex        =   16
      Top             =   420
      Width           =   750
   End
   Begin VB.Label lblTimer1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "    "
      Height          =   195
      Left            =   2730
      TabIndex        =   14
      Top             =   690
      Width           =   180
   End
   Begin VB.Label lblServerState 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "    "
      Height          =   192
      Left            =   6096
      TabIndex        =   12
      Top             =   696
      Width           =   204
   End
   Begin VB.Label lblOPC2 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "OPC"
      Height          =   195
      Left            =   9990
      TabIndex        =   9
      Top             =   240
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblOPC1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "OPC"
      Height          =   195
      Left            =   5400
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpARM2 
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   156
      Left            =   10500
      Top             =   612
      Width           =   132
   End
   Begin VB.Shape shpARM1 
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   156
      Left            =   5916
      Top             =   612
      Width           =   132
   End
   Begin VB.Label lblARM2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "ARM 2"
      Height          =   195
      Left            =   10350
      TabIndex        =   7
      Tag             =   "ARM2: "
      Top             =   0
      Width           =   585
   End
   Begin VB.Label lblARM1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "ARM 1"
      Height          =   195
      Left            =   5760
      TabIndex        =   6
      Tag             =   "ARM1: "
      Top             =   0
      Width           =   585
   End
   Begin VB.Label lblPC 
      BackColor       =   &H00FFFFC0&
      Caption         =   "C8_Conect"
      Height          =   195
      Left            =   4350
      TabIndex        =   5
      Top             =   0
      Width           =   975
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   5976
      X2              =   10506
      Y1              =   648
      Y2              =   648
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   10290
      Picture         =   "frmOPCClient.frx":2DEA
      Top             =   120
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   4560
      X2              =   5880
      Y1              =   648
      Y2              =   648
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5700
      Picture         =   "frmOPCClient.frx":558C
      Top             =   120
      Width           =   480
   End
   Begin VB.Image imgPC 
      Height          =   480
      Left            =   4290
      Picture         =   "frmOPCClient.frx":7D2E
      Top             =   120
      Width           =   480
   End
   Begin VB.Menu mnuDlitOPC 
      Caption         =   "ДЛИТЕЛЬНОСТЬ РАБОТЫ ОРС"
   End
End
Attribute VB_Name = "frmOPCClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================
'23.11.2011
'23.01.2012
'27.01.2012 ВАЖНО
'09.02.2012 ВАЖНО
'14.02.2012 ВАЖНО
'01.06.2012 ВАЖНО: ИСПРАВЛЕНЫ IP АДРЕСА НА МНЕМОСХЕМЕ
'08.08.2012 В RE() МУСОР ВОТ ТАКОГО ВИДА:
'                  2TXT °Ф™у@ф@}§у@    aеt,txtMяяяяяяяяяяяяяяеPIRA3M45_PV_acеIRA34
'                  1TXT ЇФ™у@ф@}§у@1е,txt|яяяяяяяяяяяяяяяяеFIR34|6_PV_actеIR346
'                  1TXT ЇФ™у@ф@}§у@е_corrl,txtяяяяяяеFIR34l4_PV_actеIR344
'09.11.2012 важно
'13.05.2013 ОШИБКА ДАЛЬШЕ ПО ПОДПРОГРАММЕ ПРИ ЗАПИСИ В ФАЙЛ....
'13.11.2013 ПЕРЕЗАПУСК ОРС ПО ДЛИТЕЛЬНОСТИ ЕГО РАБОТЫ
'03.02.2014 Srednee = 0
'21.02.2014
'08.04.2014 ВАЖНО ПЕРЕПОДКЛЮЧЕНИЕ К ОРС СЕРВЕРУ
'16.04.2014 ЗАГРУЗИТЬ ДАННЫЕ ПО ПОРОГАМ НЕЧУВСТВИТЕЛЬНОСТИ
'05.06.2015 ВАЖНО -> IP
'26.06.2015
'=========================================================
Option Explicit
Private OPCModule As OPCClientWinCC
Dim opcServerObj As OPCServer
Dim flxOPCTags_Tag As Date
Dim i As Long
Dim StartOPC As Date
Dim dlitOPC As Long

'FIR346_PV_act:  "Скорректированный расход пара в существующие паропроводы" 0...25
'PIRA345_PV_act":"Давление пара после РОУ-25" 0...1
'TIRA318_PV_act":"Температура пара после РОУ-25" 0...350

Private Sub cmbClose_Click()
    Unload Me
End Sub

Private Sub cmdLogOPC_Click()
    If lstLogOPC.Left < 0 Then
       lstLogOPC.Left = 10470
    Else
       lstLogOPC.Left = -11111
    End If
End Sub

Private Sub Form_Load()
    'SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    txtWait.Left = 22222
    lblStartApp.Caption = Now
    Call FillFlxOPCTags
    SName = "OPCServer.WinCC"
    SName_Net = "192.168.0.101" '05.06.2015
    ReDim ItemIDs(ItemNumb)
    
    ReDim ClientHandles(ItemNumb)
    
    'СОЗДАНИЕ МАССИВА ИМЕН ТЭГов:
    For i = 1 To ItemNumb
        ItemIDs(i) = flxOPCTags.TextMatrix(i, 0)
        ClientHandles(i) = i
    Next i
    flxOPCTags.ColWidth(0) = 2500
    flxOPCTags.ColWidth(1) = 1500
    flxOPCTags.ColWidth(2) = 1500
    flxOPCTags.ColWidth(3) = 2000
    flxOPCTags.ColWidth(4) = 6500
    flxOPCTags_Tag = DateAdd("N", -5, Now)
    'cmdConnect.Value = True
    OPCServState(1) = "OPCRunning": OPCServState(2) = "OPCFailed"
    OPCServState(3) = "OPCNoconfig": OPCServState(4) = "OPCSuspended"
    OPCServState(5) = "OPCtest":  OPCServState(6) = "OPCDisconnected"
End Sub

Private Sub FillFlxOPCTags()
    With flxOPCTags
    '''''    .ColWidth(0) = 1800
    '''''    .ColWidth(1) = 1000
    '''''    .ColWidth(2) = 1000
    '''''    .ColWidth(3) = 1600
    '''''    .ColWidth(4) = 5100
    '''''    .TextMatrix(0, 0) = "Тэг": .TextMatrix(0, 1) = "Значение"
    '''''    .TextMatrix(0, 2) = "Качество": .TextMatrix(0, 3) = "Д/Т опроса": .TextMatrix(0, 4) = "Описание"
    '''''    .TextMatrix(1, 0) = "FIR1_121_PV_act":      .TextMatrix(1, 4) = " Расход питательной воды на котел 1"
    '''''    .TextMatrix(2, 0) = "FIR1_127_PV_act_corr": .TextMatrix(2, 4) = "Скорректированный расход пара от Котла 1"
    '''''    .TextMatrix(3, 0) = "FIR1_129_PV_act_corr": .TextMatrix(3, 4) = "Скорректированный расход природного газа на котел 1"
    '''''    .TextMatrix(4, 0) = "FIR2_121_PV_act":      .TextMatrix(4, 4) = "Расход питательной воды на котел 2"
    '''''    .TextMatrix(5, 0) = "FIR2_127_PV_act_corr": .TextMatrix(5, 4) = "Скорректированный расход пара от Котла 2"
    '''''    .TextMatrix(6, 0) = "FIR2_129_PV_act_corr": .TextMatrix(6, 4) = "Скорректированный расход природного газа на котел 2"
    '''''    .TextMatrix(7, 0) = "FIR3_121_PV_act":      .TextMatrix(7, 4) = "Расход питательной воды на котел 3"
    '''''    .TextMatrix(8, 0) = "FIR3_127_PV_act_corr": .TextMatrix(8, 4) = "Скорректированный расход пара от Котла 3"
    '''''    .TextMatrix(9, 0) = "FIR3_129_PV_act_corr": .TextMatrix(9, 4) = "Скорректированный расход природного газа на котел 3"
    '''''    .TextMatrix(10, 0) = "FIR344_PV_act_corr":   .TextMatrix(10, 4) = "Скорректированный расход пара на ЦТП"
    '''''    .TextMatrix(11, 0) = "FIR346_PV_act":        .TextMatrix(11, 4) = "Скорректированный расход пара в существующие паропроводы"
    '''''    .TextMatrix(12, 0) = "FIR348_PV_act":        .TextMatrix(12, 4) = "Расход питательной воды в коллекторе №1 от ДПУ"
    '''''    .TextMatrix(13, 0) = "FIR350_PV_act":        .TextMatrix(13, 4) = "Расход питательной воды в коллекторе №2 от ДПУ"
    '''''    .TextMatrix(14, 0) = "FIR803_PV_act":        .TextMatrix(14, 4) = "Расход подпиточной воды"
    '''''    .TextMatrix(15, 0) = "PIR1_122_PV_act": .TextMatrix(15, 4) = "Давление питательной воды после узла питания Котла 1"
    '''''    .TextMatrix(16, 0) = "PIR1_126_PV_act": .TextMatrix(16, 4) = "Давление перегретого пара Котла 1"
    '''''    .TextMatrix(17, 0) = "PIR1_141_PV_act": .TextMatrix(17, 4) = "Разрежение дымовых газов перед котлом 1"
    '''''    .TextMatrix(18, 0) = "PIR2_122_PV_act": .TextMatrix(18, 4) = "Давление питательной воды после узла питания Котла 2"
    '''''    .TextMatrix(19, 0) = "PIR2_126_PV_act": .TextMatrix(19, 4) = "Давление перегретого пара Котла 2"
    '''''    .TextMatrix(20, 0) = "PIR2_141_PV_act": .TextMatrix(20, 4) = "Разрежение дымовых газов перед котлом 2"
    '''''    .TextMatrix(21, 0) = "PIR3_122_PV_act": .TextMatrix(21, 4) = "Давление питательной воды после узла питания Котла 3"
    '''''    .TextMatrix(22, 0) = "PIR3_126_PV_act": .TextMatrix(22, 4) = "Давление перегретого пара Котла 3"
    '''''    .TextMatrix(23, 0) = "PIR3_141_PV_act": .TextMatrix(23, 4) = "Разрежение дымовых газов перед котлом 3"
    '''''    .TextMatrix(24, 0) = "PIRA340_PV_act":  .TextMatrix(24, 4) = "Давление газа после ГРУ"
    '''''    .TextMatrix(25, 0) = "PIRA343_PV_act":  .TextMatrix(25, 4) = "Давление пара после РОУ-50"
    '''''    .TextMatrix(26, 0) = "PIRA345_PV_act":  .TextMatrix(26, 4) = "Давление пара после РОУ-25"
    '''''    .TextMatrix(27, 0) = "TIR1_101_PV_act": .TextMatrix(27, 4) = "Температура питательной воды до экономайзера Котла 1"
    '''''    .TextMatrix(28, 0) = "TIR1_105_PV_act": .TextMatrix(28, 4) = "Температура дымовых газов перед котлом 1"
    '''''    .TextMatrix(29, 0) = "TIR2_101_PV_act": .TextMatrix(29, 4) = "Температура питательной воды до экономайзера Котла 2"
    '''''    .TextMatrix(30, 0) = "TIR2_105_PV_act": .TextMatrix(30, 4) = "Температура дымовых газов перед котлом 2"
    '''''    .TextMatrix(31, 0) = "TIR3_101_PV_act": .TextMatrix(31, 4) = "Температура питательной воды до экономайзера Котла 3"
    '''''    .TextMatrix(32, 0) = "TIR3_105_PV_act": .TextMatrix(32, 4) = "Температура дымовых газов перед котлом 3"
    '''''    .TextMatrix(33, 0) = "TIR314_PV_act":   .TextMatrix(33, 4) = "Температура газа после ГРУ"
    '''''    .TextMatrix(34, 0) = "TIRA1_102_PV_act": .TextMatrix(34, 4) = "Температура перегретого пара Котла 1"
    '''''    .TextMatrix(35, 0) = "TIRA2_102_PV_act": .TextMatrix(35, 4) = "Температура перегретого пара Котла 2"
    '''''    .TextMatrix(36, 0) = "TIRA3_102_PV_act": .TextMatrix(36, 4) = "Температура перегретого пара Котла 3"
    '''''    .TextMatrix(37, 0) = "TIRA317_PV_act":   .TextMatrix(37, 4) = "Температура пара после РОУ-50"
    '''''    .TextMatrix(38, 0) = "TIRA318_PV_act":   .TextMatrix(38, 4) = "Температура пара после РОУ-25"
    
    
        Call ConnectingToMESQL
        Dim rD As ADODB.Recordset
        Set rD = New ADODB.Recordset
    '''''    rD.Open "select * from Tags ORDER BY Cod", cnME, adOpenKeyset, adLockOptimistic
    '''''    For i = 1 To 38
    '''''        rD.AddNew
    '''''        rD!Nam = .TextMatrix(i, 0)
    '''''        rD!Descr = .TextMatrix(i, 4)
    '''''        rD!cod = i
    '''''        rD.Update
    '''''    Next
    '''''    rD.Close
        rD.Open "select * from Tags ORDER BY Cod", cnME, adOpenKeyset
        ItemNumb = rD.RecordCount
        ReDim Tags.Nam(ItemNumb)
        ReDim Tags.Descr(ItemNumb)
        ReDim Tags.Arhiv(ItemNumb)
        ReDim Tags.Shcala(ItemNumb)
        ReDim Tags.Ma(ItemNumb)
        ReDim Tags.Mi(ItemNumb)
        For i = 1 To ItemNumb
            Tags.Nam(i) = rD!Nam
            Tags.Descr(i) = rD!Descr
            Tags.Arhiv(i) = IIf(IsNull(rD!Arhiv), "", rD!Arhiv)
            Tags.Shcala(i) = rD!Shcala
            rD.MoveNext
        Next
        rD.Close
    End With
    
    With flxOPCTags
        .Clear
        .ColWidth(0) = 1800
        .ColWidth(1) = 1000
        .ColWidth(2) = 1000
        .ColWidth(3) = 1600
        .ColWidth(4) = 5100
        .ColWidth(5) = 1000
        .TextMatrix(0, 0) = "Тэг":      .TextMatrix(0, 1) = "Значение"
        .TextMatrix(0, 2) = "Качество": .TextMatrix(0, 3) = "Д/Т опроса"
        .TextMatrix(0, 4) = "Описание": .TextMatrix(0, 5) = "Архив"
        .Rows = ItemNumb + 2
        For i = 1 To ItemNumb
            .TextMatrix(i, 0) = Tags.Nam(i)
            .TextMatrix(i, 4) = Tags.Descr(i)
            .TextMatrix(i, 5) = Tags.Arhiv(i)
        Next
    End With
   
End Sub

Private Sub cmdConnect_Click()
'''''    Host = "ARM1"
'''''    cmdPing.Value = True
'''''    lblARM1.Caption = lblARM1.Tag & Text3.Text
'''''    If Text3 Like "Request Timeout*" Or Text3 Like "*нет в домене*" Then
'''''        shpARM1.Visible = True
'''''    Else
'''''        shpARM1.Visible = False
'''''        SName_Net = ARM1
'''''        GoTo lblCon
'''''    End If
'''''    Host = "ARM2"
'''''    cmdPing.Value = True
'''''    lblARM2.Caption = lblARM2.Tag & Text3.Text
'''''    If Text3 Like "Request Timeout*" Or Text3 Like "*нет в домене*" Then
'''''        shpARM1.Visible = True
'''''        txtWait.Text="В СЕТИ ОТСУТСТВУЮТ 'ARM1' И 'ARM2'. НЕВОЗМОЖНО ПОДКЛЮЧЕНИЕ К АКТИВНОМУ OPCServer", vbCritical, "ПОДКЛЮЧЕНИЕ К OPCServer...."
'''''    Else
'''''        shpARM1.Visible = False
'''''        SName_Net = ARM2
'''''    End If

    'Call ConnectSlave
    'Call ConnectToOPCServer
End Sub

Public Function ConnectToOPCServer() As Long
On Error GoTo Problem
  ServerName = "OPCServer.WinCC"
  If opcServerObj Is Nothing Then
    Set opcServerObj = New OPCServer
  End If
  If lstLogOPC.ListCount < 32760 Then _
     lstLogOPC.AddItem Format(Now, "DD.MM hh:nn:ss ") & "подключение OPC к " & SName_Net '23.06.2015
  opcServerObj.connect ServerName, SName_Net '"OPCServer.WinCC", "192.168.0.101"
  txtConectCount.Text = txtConectCount.Text + 1
  lblTypeRefresh.Caption = "SyncRead"
  MyOPCServerConnect = True

  '*********************************************
  '       Создание группы на сервере
  '*********************************************
  If opcGroupObj Is Nothing Then
    Set opcGroupObj = opcServerObj.OPCGroups.Add("Grp1")
    
  End If
  '-------------------------------------------------------------------
  'Добавление элементов данных (ТЭГОВ) в группу
   opcGroupObj.OPCItems.AddItems ItemNumb, ItemIDs, ClientHandles, ServerHandles, Errors
  '-------------------------------------------------------------------
  'Проверка конфигурации сервера после создания группы тэгов
  For i = 1 To ItemNumb 'LBound(Errors) To UBound(Errors)
    If Errors(i) <> 0 Then
        bInvalidItems = True
        'Exit For
    Else
        bInvalidItems = False
    End If
  Next i
  '-------------------------------------------------------------------
  'конфигурация успешно проверена - выполнить контрольное чтение тэгов
    If bInvalidItems = False Then
        For i = 1 To ItemNumb
            If Errors(i) <> 0 Then
               MsgBox "Error SyncRead OPC Item", vbCritical, "ERROR"
            End If
        Next i
      'ВСЕ ТЭГИ ДОБАВЛЕНЫ И ПРОВЕРЕНЫ КОНТРОЛЬНЫМ ЧТЕНИЕМ
      'Установка признака готовности к запуску
      asyncRead = False
      tmrReadSync.Enabled = True
    Else
    
    End If
    'If SName_Net = "ARM1" Then
    If SName_Net = "192.168.0.101" Then
        lblOPC1.Visible = True
        lblOPC2.Visible = False
        lblServerState.Left = 6120
    Else
        lblOPC1.Visible = False
        lblOPC2.Visible = True
        lblServerState.Left = 10680
    End If
    If txtWait.Left > 0 Then txtWait.Left = -22222
    StartOPC = Now
    dlitOPC = 0
    ConnectToOPCServer = 1
    If lstLogOPC.ListCount < 32760 Then _
       lstLogOPC.AddItem Format(Now, "DD.MM hh:nn:SS ") & "подключен OPC к " & SName_Net '23.06.2015
    GoTo ext

Problem:
    If lstLogOPC.ListCount < 32760 Then _
       lstLogOPC.AddItem Format(Now, "DD.MM hh:nn:SS ") & "не подключено OPC к " & SName_Net & vbCrLf & Err.Description  '23.06.2015 ' 12.01.2016
     txtWait.Text = Now & " ConnectToOPCServer: " & Err.Description
     txtWait.Left = 100
     If SName_Net = "192.168.0.101" Then
        lblOPC1.Visible = True
        lblOPC2.Visible = False
        lblServerState.Left = 6120
    Else
        lblOPC1.Visible = False
        lblOPC2.Visible = True
        lblServerState.Left = 10680
    End If

ext:
    MousePointer = vbDefault
    On Error GoTo 0
End Function


Private Sub ConnectSlave()
    txtWait.Left = 100
    Me.Refresh
    MousePointer = 11
    txtConectCount.Text = txtConectCount.Text + 1
    lblTypeRefresh.Caption = "DataChange"
    '===============================================================
    ' Create new class object for OPCClientWinCC
    Set OPCModule = New OPCClientWinCC
    '===============================================================
    ' Call Connect in Class Module OPCClientWinCC with Servername
    '   - "ARM1"  "ARM2"
    '   - "OPCServer.WinCC"
    txtWait.Text = txtWait.Tag & SName_Net
    OPCModule.connect SName_Net, SName
    If MyOPCServerConnect = False Then
       txtWait = Now & ": ОШИБКА ПОДКЛЮЧЕНИЯ К OPCServer.WinCC "
       GoTo ext
    End If
    '===============================================================
    ' Call AddGroup in Class Module OPCClientWinCC with Groupname
    OPCModule.AddGroup "Group1"
    '===============================================================
    ' Call AddItems in class module OPCClientWinCC with array
    ' of ItemIDs and number of Items
    OPCModule.AddItems ItemIDs(), ItemNumb
    '===============================================================
    ' cmdConnect.Enabled = False
    cmdConnect.Enabled = True
    '===============================================================
    cmdDisconnect.Enabled = True
    txtWait.Left = 22222
    'tmrReadSync.Enabled = True
    If SName_Net = "192.168.0.101" Then '"ARM1" Then
        lblOPC1.Visible = True
        lblOPC2.Visible = False
        lblServerState.Left = 6120
    Else
        lblOPC1.Visible = False
        lblOPC2.Visible = True
        lblServerState.Left = 10680
    End If
    
ext:
    Me.MousePointer = 1
End Sub

Private Sub cmdDisconnect_Click()
    On Error GoTo err1
    Select Case lblTypeRefresh
        Case "SyncRead"
            tmrReadSync.Enabled = False
            lblOPC1.Visible = False
            lblOPC2.Visible = False
            If Not opcGroupObj Is Nothing Then '14.02.2012
               If (opcGroupObj.OPCItems.Count > 0) Then
                   opcGroupObj.OPCItems.Remove ItemNumb, ServerHandles, Errors
               End If
               If (opcServerObj.OPCGroups.Count > 0) Then '14.02.2012
                   opcServerObj.OPCGroups.Remove "Grp1"
               End If
               Set opcGroupObj = Nothing
            End If
            If Not opcServerObj Is Nothing Then
                opcServerObj.Disconnect
                Set opcServerObj = Nothing
                lblServerState.Caption = Format(Now, "dd.mm hh:nn") & " отключен от " & SName_Net '23.06.2015
                If lstLogOPC.ListCount < 32760 Then _
                   lstLogOPC.AddItem Format(Now, "dd.mm hh:nn:SS") & " отключен от " & SName_Net '23.06.2015
            End If
            MyOPCServerConnect = False
        Case "DataChange"
            '=================================================
            ' Call RemItems in Class Module ExcelOPCClass
            OPCModule.RemItems ItemNumb
            '=================================================
            ' Call RemGroup in Class Module ExcelOPCClass
            OPCModule.RemGroup "Group1"
            '=================================================
            ' Call Disconnect in Class Module ExcelOPCClass
            OPCModule.Disconnect
    End Select
    cmdConnect.Enabled = True
    cmdDisconnect.Enabled = False
    SName_Net = ""
    Exit Sub
err1:
    '08.04.2014 If (Err.Number = 462)  Then
    If (Err.Number = 462) Or (Err.Number = -2147467259) Then
    'The remote server machine does not exist or is unavailable
        Set opcGroupObj = Nothing
        If Not opcServerObj Is Nothing Then
           opcServerObj.Disconnect
           Set opcServerObj = Nothing
           lblServerState.Caption = Format(Now, "dd.mm hh:nn") & " ошибка отключения от " & SName_Net '23.06.2015
           If lstLogOPC.ListCount < 32760 Then _
              lstLogOPC.AddItem Format(Now, "dd.mm hh:nn:SS") & " ошибка отключения от " & SName_Net '23.06.2015
        End If
        MyOPCServerConnect = False
        cmdConnect.Enabled = True
        cmdDisconnect.Enabled = False
        SName_Net = ""
    End If
End Sub

Private Sub cmdPing_Click()
    On Error GoTo exopc:
        cmdPing.Enabled = False
        'Me.MousePointer = 11
        Text3 = ""
        Call PingHost
        
        Text3 = RezultPing
        'cmdPing.Enabled = True
        'Me.MousePointer = 0
exopc:
        On Error GoTo 0
        'Host = ""
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    cmdDisconnect.Value = True
    Set frmOPCClient = Nothing
    rsPorogZero.Close
    Set rsPorogZero = Nothing
End Sub



Private Sub tmrPing_Timer()
    On Error GoTo err1
    tmrPing.Enabled = False
'''''    '=====================================01.06.2012====================
'''''    ' ПРИ ПЕРВОМ СРАБАТЫВАНИИ ТВЙМЕРА............
'''''    '===================================================================
'''''    If SName_Net = "" Then
'''''        shpARM1.Visible = False
'''''        SName_Net = ARM1
'''''        GoTo BOOT
'''''    End If
    '=====================================01.06.2012====================
   If MyOPCServerConnect Then
    '    lblOPC1.Caption = OPCServState(OPCModule.state)
    End If
    '===================================================================
    '
    '===================================================================
    '01.06.2012 Host = "ARM3"
    Host = SName_Net '08.04.2014 "ARM1" '01.06.2012
    cmdPing.Value = True
    If SName_Net = "192.168.0.101" Then '"ARM1" Then
        lblARM1.Caption = lblARM1.Tag & Format(Now, "hh:nn:ss") & " " & Text3.Text
    Else
        lblARM2.Caption = lblARM2.Tag & Format(Now, "hh:nn:ss") & " " & Text3.Text
    End If
    If Text3 Like "Request Timeout*" Or Text3 Like "*нет в домене*" Then
        shpARM1.Visible = True
        '-------------------------------------------------
        ' ОТКЛЮЧИТЬСЯ ОТ СЕРВЕРА, Т.К. УЗЕЛ НЕ ПИНГУЕТСЯ
        '-------------------------------------------------
        'If (SName_Net = ARM1) Then
        If (SName_Net = "192.168.0.101") Then
            cmdDisconnect.Value = True
        End If
    Else
        shpARM1.Visible = False
        '08.04.2014 If MyOPCServerConnect = False Then SName_Net = ARM1
    End If
    '===================================================================
    '
    '===================================================================
BOOT:
    '05.06.2015 Host = "ARM2" '01.06.2012
    Host = "192.168.0.102"
    cmdPing.Value = True
    lblARM2.Caption = lblARM2.Tag & Format(Now, "hh:nn:ss") & " " & Text3.Text
    If Text3 Like "Request Timeout*" Or Text3 Like "*нет в домене*" Then
        shpARM2.Visible = True
        '-------------------------------------------------
        ' ОТКЛЮЧИТЬСЯ ОТ СЕРВЕРА, Т.К. УЗЕЛ НЕ ПИНГУЕТСЯ
        '-------------------------------------------------
       'If (SName_Net = ARM2) Then
        If (SName_Net = "192.168.0.102") Then
            cmdDisconnect.Value = True
        End If
    Else
        shpARM2.Visible = False
        'If MyOPCServerConnect = False Then SName_Net = ARM2
        If (MyOPCServerConnect = False) Then
            If SName_Net = "" Then
                SName_Net = "192.168.0.101"
            ElseIf SName_Net = "192.168.0.101" Then
                SName_Net = "192.168.0.102"
            ElseIf SName_Net = "192.168.0.102" Then
                SName_Net = "192.168.0.101"
            End If
        End If
    End If
    '===================================================================
    '               ОБА УЗЛА НЕ ПИНГУЮТСЯ
    '===================================================================
    If shpARM1.Visible And shpARM2.Visible Then
        cmdConnect.Enabled = False
    End If

err1:
    If (MyOPCServerConnect = False) And (SName_Net <> "") Then
         'Call ConnectSlave
         '08.04.2014 Call ConnectToOPCServer
         If ConnectToOPCServer = 0 Then
            'If SName_Net = "ARM1" Then
            If SName_Net = "192.168.0.101" Then
                lblOPC1.Visible = True
                lblOPC2.Visible = False
                lblServerState.Left = 6120
                SName_Net = "192.168.0.102" '"ARM2"
            Else
                lblOPC1.Visible = False
                lblOPC2.Visible = True
                lblServerState.Left = 10680
                SName_Net = "192.168.0.101" '"ARM1"
            End If
            
         
         End If
         
    End If
    On Error GoTo 0
    tmrPing.Enabled = True
End Sub

Private Sub tmrReadSync_Timer()
Dim Qualities(555)  As Integer
Dim s As String
Dim sT As String
Dim sDT As String
Dim RE
Dim dtTEMP As Date
Dim Srednee As Single
Dim K As Long
Dim IdentityID As Long
Dim dtIF As Date
    ReDim pTimestamp(555)
    lblTimer1.Caption = Now
    On Error GoTo ErrRead
    
    '=========================================================
    '      OPCServerState
    '=========================================================
    Select Case opcServerObj.ServerState
        Case 1
            lblServerState.Caption = "OPCRunning"
        Case 2
            lblServerState.Caption = "OPCFailed"
        Case 3
            lblServerState.Caption = "OPCNoconfig"
        Case 4
            lblServerState.Caption = "OPCSuspended"
        Case 5
            lblServerState.Caption = "OPCTest"
        Case 6
            lblServerState.Caption = "OPCDisconnected"
        Case Else
            lblServerState.Caption = opcServerObj.ServerState
    End Select
    '=========================================================
    '  opcGroupObj.SyncRead  OPCCache.....:
    '  ПРИ ОБРЫВЕ СЕТИ ПРОДОЛЖАЕТСЯ ВЫВОД СОСТОЯНИЯ ИЗ OPCCache.
    '  ПРИ ЭТОМ pTimestamp() ОБНОВЛЯЕТСЯ.
    '  OPCDevice
    '=========================================================
    opcGroupObj.SyncRead OPCDevice, ItemNumb, ServerHandles, ReadValue, Errors, Qualities, pTimestamp
    opcGroupObj.SyncRead OPCCache, ItemNumb, ServerHandles, ReadValue, Errors, Qualities, pTimestamp
    '=========================================================
    '           OPC Quality Code Overview
    '=========================================================
    '0x04 = 4  = Config Error
    '0x08 = 8  = Not Connected
    '0x0C = 12 = Device Failure
    '0x10 = 16 = Sensor Failure
    '0x14 = 20 = Last Known
    '0x18 = 24 = Comm Failure
    '0x1C = 28 = Out of Service
    '0x20 = 32 = Initializing
    '
    '0x40 = 64 = Uncertain
    '0x44 = 68 = Last Usable
    '0x50 = 80 = Sensor Calibration
    '0x54 = 84 = EGU Exceeded
    '0x58 = 88 = Sub Normal
    '
    '0xC0 = 192 = Good
    '0xD8 = 216 = Local Override

    For i = 1 To ItemNumb
        If Errors(i) <> 0 Then
           MsgBox "Error SyncRead OPC Item", vbCritical, "ERROR"
        End If
    Next i
    IdentityID = GetIdentityID
    For i = 1 To ItemNumb
        flxOPCTags.TextMatrix(i, 1) = FormatNumber(ReadValue(i), 4)
        flxOPCTags.TextMatrix(i, 3) = pTimestamp(i)
        Select Case Qualities(i)
            Case 192
                flxOPCTags.TextMatrix(i, 2) = "GOOD"
            Case 24
                flxOPCTags.TextMatrix(i, 2) = "Comm Failure"
            Case 64
                flxOPCTags.TextMatrix(i, 2) = "UNCERTAIN"
            Case Else
                'flxOPCTags.TextMatrix(i, 4) = Hex(Qualities(i))
                flxOPCTags.TextMatrix(i, 2) = Qualities(i)
        End Select
        'GoTo NEX '13.05.2013
        
        's = s & "~" & ItemIDs(i) _
              & "~" & ReadValue(i) _
              & "~" & Qualities(i) _
              & "~" & pTimestamp(i)
        If Tags.Arhiv(i) <> "A" Then GoTo NEX
        '===================================================================
        'DD~MM~YYYY~HH~P~
        '   DD-ДЕНЬ СОЗДАНИЯ И ЗАПИСИ В ФАЙЛ
        '   ММ-МЕСЯЦ СОЗДАНИЯ И ЗАПИСИ В ФАЙЛ
        '   HH-ЧАС СУТОК СОЗДАНИЯ И ЗАПИСИ В ФАЙЛ
        '   P -1Я ИЛИ 2Я ПОЛОВИНА ЧАСА (HH), К КОТОРОМУ ОТНОСИТСЯ ЗАПИСЬ В ФАЙЛЕ
        '===================================================================
        If IsFileExists("Z:\" & Tags.Nam(i) & ".txt") Then
            Open "Z:\" & Tags.Nam(i) & ".txt" For Input As #1
                Input #1, s
            Close #1
            '------------------------------------------------------------------
            ' ЕСЛИ ДЕНЬ, МЕСЯЦ И ПОЛОВИНА ЧАСА СОЗДАНИЯ (ЗАПИСИ) ФАЙЛА
            ' НЕ СООТВЕТСТВУЕТ НАСТОЯЩЕМУ - ЗНАЧИТ ФАЙЛ УСТАРЕВШИЙ
            ' И НАДО НАЧАТЬ ЗАПИСЬ ПО ДАННОМУ ТЭГУ В НОВЫЙ ФАЙЛ.
            '------------------------------------------------------------------
            dtTEMP = Now '02.08 + 1
            '02.08.2012 If (CLng(Left(s, 2)) <> Day(dtTEMP)) Or _
               (CLng(Mid$(s, 4, 2)) <> Month(dtTEMP)) Or _
               (CLng(Mid$(s, 12, 2)) <> Hour(dtTEMP)) Or _
               (CLng(Mid$(s, 15, 1)) <> Int(Minute(dtTEMP) / 30) + 1) Then
            dtIF = Mid$(s, 17, 10) & " " & Mid$(s, 12, 2) & ":00"
'            If Mid$(s, 15, 1) = "1" Then
'                dtIF = DateAdd("n", 30, dtIF)
'            Else
'                dtIF = DateAdd("h", 1, dtIF)
'            End If
            If (Day(dtIF) <> Day(dtTEMP)) Or _
               (Month(dtIF) <> Month(dtTEMP)) Or _
               (Hour(dtIF) <> Hour(dtTEMP)) Or _
               (CLng(Mid$(s, 15, 1)) <> Int(Minute(dtTEMP) / 30) + 1) Then
               '---------------------------------------------------------------
               ' УСРЕДНИТЬ И ПЕРЕНЕСТИ СРЕДНЕ ЗНАЧЕНИЕ В МЕСЯЧНУЮ ТАБЛИЦУ
               '---------------------------------------------------------------
                sT = "K" & Mid$(s, 4, 2) & "_" & Mid$(s, 7, 4)
                Call ConnectingToMESQL
                Dim rD As ADODB.Recordset
                Set rD = New ADODB.Recordset
                rD.Open "SP_TABLES '" & sT & "'", cnME
                If rD.EOF = False Then GoTo lblADD
                Set cmdME = New ADODB.Command
                cmdME.ActiveConnection = cnME
                cmdME.CommandText = "CREATE TABLE [" & sT & "] (Sutky int,DT Datetime, Tag nvarchar(32),Val Numeric(9,1), Cod INT PRIMARY KEY)"
                cmdME.Execute
                
lblADD:         RE = Split(Mid$(s, 28), "~")
                Dim K2 As Long
                K2 = 0
                Srednee = 0 '03.02.2014
                '2TXT °Ф™у@ф@}§у@    aеt,txtMяяяяяяяяяяяяяяеPIRA3M45_PV_acеIRA34
                '1TXT ЇФ™у@ф@}§у@1е,txt|яяяяяяяяяяяяяяяяеFIR34|6_PV_actеIR346
                '1TXT ЇФ™у@ф@}§у@е_corrl,txtяяяяяяеFIR34l4_PV_actеIR344
               
                For K = 0 To UBound(RE) ' - 1
                    RE(K) = Replace(RE(K), ".", ",")
                    If IsNumeric(RE(K)) Then
                        If (RE(K) <> 0) Then '16.05.2014
                            Srednee = Srednee + CSng(RE(K))
                            K2 = K2 + 1
                        End If
                    End If
                Next
                If (K2 > 0) Then Srednee = Srednee / K2
                rD.Close
                rD.Open "SELECT * FROM [" & sT & "] WHERE Cod=0", cnME, adOpenKeyset, adLockOptimistic
                rD.AddNew
                    rD!Cod = IdentityID + i
                    'Debug.Print Tags.Nam(i)
                    rD!Tag = Tags.Nam(i)
                    rD!Val = Srednee
                    '02.08.2012sDT = Left(s, 2) & "." & Mid$(s, 4, 2) & "." & Mid$(s, 7, 4) & " " & Mid$(s, 12, 2) & ":00"
                    sDT = Mid$(s, 17, 10) & " " & Mid$(s, 12, 2) & ":00"
                    If Mid$(s, 15, 1) = "1" Then
                        rD!DT = DateAdd("n", 30, CDate(sDT))
                    Else
                        rD!DT = DateAdd("h", 1, CDate(sDT))
                    End If
                    rD!Sutky = Left(s, 2) '02.08.2012Day(rD!DT) ' - 1
                    '02.08.2012 rD!DT = rD!DT - 1
                    
                    
                    '===============================
                    '   НАЧАЛО СУТОК В 07:00
                    '===============================
'                    If (Format(rD!DT, "HH:NN") >= "07:00") And (Format(rD!DT, "HH:NN") <= "23:59") Then
'                       rD!Sutky = Day(rD!DT) + 1
'                    Else
'                       rD!Sutky = Day(rD!DT)
'                    End If
                rD.Update
                Kill "Z:\" & Tags.Nam(i) & ".txt"
                '------------------------------------------------------------------------------
                ' ГОТОВИМ ЗАГОЛОВОК "DD~MM~YYYY~HH~P".
                ' ЗАГОЛОВОК ФАЙЛА ДОЛЖЕН СОДЕРЖАТЬ ДАТУ, К КОТОРОЙ БУДУТ ПРИНАДЛЕЖАТЬ ДАННЫЕ.
                ' Т.Е. ИМЯ ТАБЛИЦЫ БУДЕТ ФОРМИРОВАТЬСЯ ИЗ ЗАГОЛОВКА ФАЙЛА
                '------------------------------------------------------------------------------
                '30.07.2012 s = Format(Now, "DD~MM~YYYY~HH") & "~" & Int(Minute(Time) / 30) + 1
                dtTEMP = Now + 1
                '--------------------------------------
                '          СЛЕДУЮЩИЙ МЕСЯЦ
                '--------------------------------------
                If Month(Now) < Month(dtTEMP) Then
                    '03.08 If (Format(dtTEMP, "HH:NN") >= "07:00") And (Format(dtTEMP, "HH:NN") <= "23:59") Then
                    If (Format(dtTEMP, "HH:NN") >= "06:30") And (Format(dtTEMP, "HH:NN") <= "23:59") Then
                       dtTEMP = Now + 1
                    Else
                       dtTEMP = Now
                    End If
                    s = Format(dtTEMP, "DD~MM~YYYY~HH") & "~" & Int(Minute(dtTEMP) / 30) + 1
                Else
                    '03.08.If (Format(Now, "HH:NN") >= "07:00") And (Format(Now, "HH:NN") <= "23:59") Then
                    If (Format(Now, "HH:NN") >= "06:30") And (Format(Now, "HH:NN") <= "23:59") Then
                       dtTEMP = Now + 1
                       s = Format(dtTEMP, "DD~MM~YYYY~HH") & "~" & Int(Minute(dtTEMP) / 30) + 1
                    Else
                       s = Format(Now, "DD~MM~YYYY~HH") & "~" & Int(Minute(Time) / 30) + 1
                    End If
                End If
                s = s & "~" & Format(Now, "DD.MM.YYYY")
            End If
         '------------------------------------------------------
         '       НЕТ ФАЙЛА - ГОТОВИМ ЗАГОЛОВОК "DD~MM~P"
         '------------------------------------------------------
        Else
            dtTEMP = Now + 1
            '--------------------------------------
            '          СЛЕДУЮЩИЙ МЕСЯЦ
            '--------------------------------------
            If Month(Now) < Month(dtTEMP) Then
                '03.08If (Format(dtTEMP, "HH:NN") >= "07:00") And (Format(dtTEMP, "HH:NN") <= "23:59") Then
                If (Format(dtTEMP, "HH:NN") >= "06:30") And (Format(dtTEMP, "HH:NN") <= "23:59") Then
                   dtTEMP = Now + 1
                Else
                   dtTEMP = Now
                End If
                s = Format(dtTEMP, "DD~MM~YYYY~HH") & "~" & Int(Minute(dtTEMP) / 30) + 1
            Else
                '03.08If (Format(Now, "HH:NN") >= "07:00") And (Format(Now, "HH:NN") <= "23:59") Then
                If (Format(Now, "HH:NN") >= "06:30") And (Format(Now, "HH:NN") <= "23:59") Then
                   dtTEMP = Now + 1
                   s = Format(dtTEMP, "DD~MM~YYYY~HH") & "~" & Int(Minute(dtTEMP) / 30) + 1
                Else
                   s = Format(Now, "DD~MM~YYYY~HH") & "~" & Int(Minute(Time) / 30) + 1
                End If
            End If
            s = s & "~" & Format(Now, "DD.MM.YYYY")
        End If
        Open "Z:\" & Tags.Nam(i) & ".txt" For Output As #1
            '------------------------------16.05.2014--------------
            Dim PorogZero As Single
            Dim sReadVal As String
            Dim ReadVal As Single
            ReadVal = CSng(ReadValue(i))
            sReadVal = Replace(ReadValue(i), ",", ".")
            rsPorogZero.MoveFirst
            rsPorogZero.Find "Nam='" & Tags.Nam(i) & "'"
            If rsPorogZero.EOF Then
               s = s & "~" & Replace(ReadValue(i), ",", ".")
            ElseIf rsPorogZero!PorogZero = "-" Then
               s = s & "~" & Replace(ReadValue(i), ",", ".")
            ElseIf (CSng(rsPorogZero!PorogZero) <= ReadVal) Then
               s = s & "~" & Replace(ReadValue(i), ",", ".")
            Else
               s = s & "~0"
            End If
            '----------------------------16.05.2014----------------
            '16.05.2014 s = s & "~" & Replace(ReadValue(i), ",", ".")
            Print #1, s
        Close #1
NEX:
    Next
    flxOPCTags.TextMatrix(i, 0) = "flxOPCTags.Tag"
    flxOPCTags.TextMatrix(i, 3) = flxOPCTags_Tag
    flxOPCTags_Tag = Now
    GoTo ext

ErrRead:
    txtWait.Text = Now & " tmrReadSync_Timer:" & Err.Description
    txtWait.Left = 100
    '-----------------------27.01.2012-------------------------------------------
    ' 462: The remote server machine does not exist or is unavailable
    ' -2147467259: Automation error
    '----------------------------------------------------------------------------
    If (Err.Number = 462) Or (Err.Number = -2147467259) Then
        If Not opcServerObj Is Nothing Then
            cmdDisconnect.Value = True '09.02.2012
        End If
        MyOPCServerConnect = False
        SName_Net = ""
    End If
    '-----------------------27.01.2012-------------------------------------------
ext:
    On Error GoTo 0
    txtWait.Left = -22222
     Close #1 '16.05.2014
End Sub


Private Sub tmrSaveFile_Timer()
Dim s As String
On Error GoTo ErrRead
    
    '==============================13.11.2013========
    dlitOPC = DateDiff("H", StartOPC, Now)
    'dlitOPC = DateDiff("N", StartOPC, Now)
    mnuDlitOPC.Caption = "ДЛИТЕЛЬНОСТЬ РАБОТЫ ОРС " & dlitOPC & " ЧАС"
    If tmrReadSync.Enabled = False Then tmrReadSync.Enabled = True '08.04.2014
    If (dlitOPC >= 72) Then
        cmdDisconnect.Value = True
        cmdConnect.Value = True
        Exit Sub
    End If
    '==============================13.11.2013========
    'Debug.Print tmrReadSync.Enabled
    'If tmrReadSync.Enabled = False Then tmrReadSync.Enabled = True
    lblTimer1.Caption = Now
    For i = 1 To flxOPCTags.Rows - 1
        If flxOPCTags.TextMatrix(i, 0) = "" Then Exit For
        s = s & "~" & flxOPCTags.TextMatrix(i, 0) _
              & "~" & flxOPCTags.TextMatrix(i, 1) _
              & "~" & flxOPCTags.TextMatrix(i, 2) _
              & "~" & flxOPCTags.TextMatrix(i, 3)
    Next
    
    s = "Req~" & lblTimer1.Caption & "~" & lblTypeRefresh.Caption & "~" _
               & lblOPC1.Visible & "~" & lblOPC2.Visible & "~" _
               & lblARM1.Caption & "~" & lblARM2.Caption & "~" _
               & shpARM1.Visible & "~" & shpARM2.Visible & "~" _
               & lblServerState.Caption & s & "~||"
    Open "Z:\Kotel.opc" For Output As #1
        Print #1, s
    Close #1
    GoTo ext

ErrRead:
    txtWait.Text = Now & " tmrSaveFile_Timer:" & Err.Description
    txtWait.Left = 100
ext:
    On Error GoTo 0
    txtWait.Left = -22222

End Sub

