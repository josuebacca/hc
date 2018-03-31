VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5F09B5DF-6F4D-11D2-8355-4854E82A9183}#15.0#0"; "FECHA32.OCX"
Begin VB.Form FrmBoletaDeposito 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Boleta de Depósito"
   ClientHeight    =   7515
   ClientLeft      =   1110
   ClientTop       =   720
   ClientWidth     =   9030
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   9030
   StartUpPosition =   2  'CenterScreen
   Begin FechaCtl.Fecha TxtBolFecha 
      Height          =   300
      Left            =   3720
      TabIndex        =   1
      Top             =   135
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   529
      Separador       =   "/"
      Text            =   ""
      MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
   End
   Begin VB.CommandButton CmdEliminar 
      Caption         =   "&Eliminar"
      DisabledPicture =   "FrmBoletaDeposito.frx":0000
      Height          =   810
      Left            =   4725
      Picture         =   "FrmBoletaDeposito.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   88
      Top             =   6420
      Width           =   1020
   End
   Begin VB.TextBox TxtBanNomCor 
      Height          =   330
      Index           =   9
      Left            =   2670
      TabIndex        =   70
      Top             =   4635
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.TextBox TxtBanNomCor 
      Height          =   330
      Index           =   8
      Left            =   2670
      TabIndex        =   69
      Top             =   4305
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.TextBox TxtBanNomCor 
      Height          =   330
      Index           =   7
      Left            =   2670
      TabIndex        =   68
      Top             =   3960
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.TextBox TxtBanNomCor 
      Height          =   330
      Index           =   6
      Left            =   2670
      TabIndex        =   67
      Top             =   3645
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.TextBox TxtBanNomCor 
      Height          =   330
      Index           =   5
      Left            =   2670
      TabIndex        =   66
      Top             =   3300
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.TextBox TxtBanNomCor 
      Height          =   330
      Index           =   4
      Left            =   2670
      TabIndex        =   65
      Top             =   2970
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.TextBox TxtBanNomCor 
      Height          =   330
      Index           =   3
      Left            =   2670
      TabIndex        =   64
      Top             =   2640
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.TextBox TxtBanNomCor 
      Height          =   330
      Index           =   2
      Left            =   2670
      TabIndex        =   63
      Top             =   2310
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.TextBox TxtBanNomCor 
      Height          =   330
      Index           =   1
      Left            =   2670
      TabIndex        =   62
      Top             =   1980
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      DisabledPicture =   "FrmBoletaDeposito.frx":0BD4
      Height          =   810
      Left            =   6795
      Picture         =   "FrmBoletaDeposito.frx":0EDE
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   6420
      Width           =   1020
   End
   Begin VB.CommandButton CmdAnular 
      Caption         =   "&Anular"
      DisabledPicture =   "FrmBoletaDeposito.frx":17A8
      Height          =   810
      Left            =   3690
      Picture         =   "FrmBoletaDeposito.frx":1AB2
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   6420
      Width           =   1020
   End
   Begin VB.TextBox TxtBanCodInt 
      Enabled         =   0   'False
      Height          =   315
      Left            =   5895
      TabIndex        =   45
      Top             =   525
      Width           =   465
   End
   Begin VB.ComboBox CboBancoBoleta 
      Height          =   315
      ItemData        =   "FrmBoletaDeposito.frx":1DBC
      Left            =   1605
      List            =   "FrmBoletaDeposito.frx":1DBE
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   525
      Width           =   3255
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      DisabledPicture =   "FrmBoletaDeposito.frx":1DC0
      Height          =   810
      Left            =   7830
      Picture         =   "FrmBoletaDeposito.frx":20CA
      Style           =   1  'Graphical
      TabIndex        =   85
      Top             =   6420
      Width           =   1020
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "&Nueva"
      DisabledPicture =   "FrmBoletaDeposito.frx":23D4
      Height          =   810
      Left            =   2655
      Picture         =   "FrmBoletaDeposito.frx":26DE
      Style           =   1  'Graphical
      TabIndex        =   84
      Top             =   6420
      Width           =   1020
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Guardar"
      DisabledPicture =   "FrmBoletaDeposito.frx":29E8
      Enabled         =   0   'False
      Height          =   810
      Left            =   5760
      Picture         =   "FrmBoletaDeposito.frx":2CF2
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6420
      Width           =   1020
   End
   Begin VB.ComboBox CboCuentas 
      Height          =   315
      Left            =   7290
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   525
      Width           =   1485
   End
   Begin VB.TextBox TxtBoleta 
      Height          =   315
      Left            =   1605
      MaxLength       =   8
      TabIndex        =   0
      Top             =   120
      Width           =   1110
   End
   Begin TabDlg.SSTab SSTabABMCheque 
      Height          =   5130
      Left            =   195
      TabIndex        =   20
      Top             =   990
      Width           =   8670
      _ExtentX        =   15293
      _ExtentY        =   9049
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   706
      TabCaption(0)   =   "."
      TabPicture(0)   =   "FrmBoletaDeposito.frx":2FFC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(22)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(18)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(17)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(16)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(2)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Line1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "LblAnulada"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(14)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "TxtSucursal(7)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "TxtBanco(7)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "TxtNumeroCh(7)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "TxtSucursal(6)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "TxtBanco(6)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "TxtNumeroCh(6)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "TxtNumeroCh(5)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "TxtBanco(5)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "TxtSucursal(5)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "TxtNumeroCh(4)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "TxtBanco(4)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "TxtSucursal(4)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "TxtNumeroCh(3)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "TxtBanco(3)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "TxtSucursal(3)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "TxtNumeroCh(2)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "TxtBanco(2)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "TxtSucursal(2)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "TxtNumeroCh(1)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "TxtBanco(1)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "TxtSucursal(1)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "TxtSucursal(0)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "TxtBanco(0)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "TxtNumeroCh(0)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "TxtSucursal(9)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "TxtBanco(9)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "TxtNumeroCh(9)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "TxtSucursal(8)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "TxtBanco(8)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "TxtNumeroCh(8)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "TxtCodInt(0)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "TxtCodInt(1)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "TxtCodInt(2)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "TxtCodInt(3)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "TxtCodInt(4)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "TxtCodInt(5)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "TxtCodInt(6)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "TxtCodInt(7)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "TxtCodInt(8)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "TxtCodInt(9)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "TxtBanNomCor(0)"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "SumandoTODO"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "TxtEfvo"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "TxtValorNominalCh(8)"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "TxtValorNominalCh(9)"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "TxtValorNominalCh(0)"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "TxtValorNominalCh(1)"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "TxtValorNominalCh(2)"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "TxtValorNominalCh(3)"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "TxtValorNominalCh(4)"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "TxtValorNominalCh(5)"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "SumandoCh"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "TxtValorNominalCh(6)"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "TxtValorNominalCh(7)"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).ControlCount=   63
      Begin VB.TextBox TxtValorNominalCh 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   7
         Left            =   7110
         TabIndex        =   82
         Top             =   2970
         Width           =   1275
      End
      Begin VB.TextBox TxtValorNominalCh 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   6
         Left            =   7110
         TabIndex        =   81
         Top             =   2640
         Width           =   1275
      End
      Begin VB.TextBox SumandoCh 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   315
         Left            =   7110
         TabIndex        =   80
         Top             =   3990
         Width           =   1275
      End
      Begin VB.TextBox TxtValorNominalCh 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   5
         Left            =   7110
         TabIndex        =   79
         Top             =   2310
         Width           =   1275
      End
      Begin VB.TextBox TxtValorNominalCh 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   7110
         TabIndex        =   78
         Top             =   1980
         Width           =   1275
      End
      Begin VB.TextBox TxtValorNominalCh 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   7110
         TabIndex        =   77
         Top             =   1650
         Width           =   1275
      End
      Begin VB.TextBox TxtValorNominalCh 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   7110
         TabIndex        =   76
         Top             =   1320
         Width           =   1275
      End
      Begin VB.TextBox TxtValorNominalCh 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   7110
         TabIndex        =   75
         Top             =   990
         Width           =   1275
      End
      Begin VB.TextBox TxtValorNominalCh 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   7110
         TabIndex        =   74
         Top             =   660
         Width           =   1275
      End
      Begin VB.TextBox TxtValorNominalCh 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   9
         Left            =   7110
         TabIndex        =   73
         Top             =   3645
         Width           =   1275
      End
      Begin VB.TextBox TxtValorNominalCh 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   8
         Left            =   7110
         TabIndex        =   72
         Top             =   3315
         Width           =   1275
      End
      Begin VB.TextBox TxtEfvo 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   7110
         TabIndex        =   14
         Top             =   4335
         Width           =   1275
      End
      Begin VB.TextBox SumandoTODO 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   315
         Left            =   7110
         TabIndex        =   71
         Top             =   4725
         Width           =   1275
      End
      Begin VB.TextBox TxtBanNomCor 
         Height          =   330
         Index           =   0
         Left            =   2475
         TabIndex        =   61
         Top             =   645
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.TextBox TxtCodInt 
         Height          =   330
         Index           =   9
         Left            =   4575
         TabIndex        =   55
         Top             =   3645
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox TxtCodInt 
         Height          =   330
         Index           =   8
         Left            =   4575
         TabIndex        =   54
         Top             =   3315
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox TxtCodInt 
         Height          =   330
         Index           =   7
         Left            =   4575
         TabIndex        =   53
         Top             =   2970
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox TxtCodInt 
         Height          =   330
         Index           =   6
         Left            =   4575
         TabIndex        =   52
         Top             =   2640
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox TxtCodInt 
         Height          =   330
         Index           =   5
         Left            =   4575
         TabIndex        =   51
         Top             =   2310
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox TxtCodInt 
         Height          =   330
         Index           =   4
         Left            =   4575
         TabIndex        =   50
         Top             =   1980
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox TxtCodInt 
         Height          =   330
         Index           =   3
         Left            =   4575
         TabIndex        =   49
         Top             =   1650
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox TxtCodInt 
         Height          =   330
         Index           =   2
         Left            =   4575
         TabIndex        =   48
         Top             =   1320
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox TxtCodInt 
         Height          =   330
         Index           =   1
         Left            =   4575
         TabIndex        =   47
         Top             =   990
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox TxtCodInt 
         Height          =   330
         Index           =   0
         Left            =   4575
         TabIndex        =   46
         Top             =   660
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox TxtNumeroCh 
         Enabled         =   0   'False
         Height          =   315
         Index           =   8
         Left            =   345
         TabIndex        =   12
         Top             =   3315
         Width           =   1410
      End
      Begin VB.TextBox TxtBanco 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   8
         Left            =   1770
         TabIndex        =   44
         Top             =   3315
         Width           =   3435
      End
      Begin VB.TextBox TxtSucursal 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   8
         Left            =   5220
         TabIndex        =   43
         Top             =   3315
         Width           =   1875
      End
      Begin VB.TextBox TxtNumeroCh 
         Enabled         =   0   'False
         Height          =   315
         Index           =   9
         Left            =   345
         TabIndex        =   13
         Top             =   3645
         Width           =   1410
      End
      Begin VB.TextBox TxtBanco 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   9
         Left            =   1770
         TabIndex        =   42
         Top             =   3645
         Width           =   3435
      End
      Begin VB.TextBox TxtSucursal 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   9
         Left            =   5220
         TabIndex        =   41
         Top             =   3645
         Width           =   1875
      End
      Begin VB.TextBox TxtNumeroCh 
         Height          =   315
         Index           =   0
         Left            =   345
         TabIndex        =   4
         Top             =   660
         Width           =   1410
      End
      Begin VB.TextBox TxtBanco 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1770
         TabIndex        =   36
         Top             =   660
         Width           =   3435
      End
      Begin VB.TextBox TxtSucursal 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   5220
         TabIndex        =   35
         Top             =   660
         Width           =   1875
      End
      Begin VB.TextBox TxtSucursal 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   5220
         TabIndex        =   34
         Top             =   990
         Width           =   1875
      End
      Begin VB.TextBox TxtBanco 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   1770
         TabIndex        =   33
         Top             =   990
         Width           =   3435
      End
      Begin VB.TextBox TxtNumeroCh 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   345
         TabIndex        =   5
         Top             =   990
         Width           =   1410
      End
      Begin VB.TextBox TxtSucursal 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   5220
         TabIndex        =   32
         Top             =   1320
         Width           =   1875
      End
      Begin VB.TextBox TxtBanco 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   1770
         TabIndex        =   31
         Top             =   1320
         Width           =   3435
      End
      Begin VB.TextBox TxtNumeroCh 
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   345
         TabIndex        =   6
         Top             =   1320
         Width           =   1410
      End
      Begin VB.TextBox TxtSucursal 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   5220
         TabIndex        =   30
         Top             =   1650
         Width           =   1875
      End
      Begin VB.TextBox TxtBanco 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   1770
         TabIndex        =   29
         Top             =   1650
         Width           =   3435
      End
      Begin VB.TextBox TxtNumeroCh 
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   345
         TabIndex        =   7
         Top             =   1650
         Width           =   1410
      End
      Begin VB.TextBox TxtSucursal 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   5220
         TabIndex        =   28
         Top             =   1980
         Width           =   1875
      End
      Begin VB.TextBox TxtBanco 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   1770
         TabIndex        =   27
         Top             =   1980
         Width           =   3435
      End
      Begin VB.TextBox TxtNumeroCh 
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   345
         TabIndex        =   8
         Top             =   1980
         Width           =   1410
      End
      Begin VB.TextBox TxtSucursal 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   5
         Left            =   5220
         TabIndex        =   26
         Top             =   2310
         Width           =   1875
      End
      Begin VB.TextBox TxtBanco 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   5
         Left            =   1770
         TabIndex        =   25
         Top             =   2295
         Width           =   3435
      End
      Begin VB.TextBox TxtNumeroCh 
         Enabled         =   0   'False
         Height          =   315
         Index           =   5
         Left            =   345
         TabIndex        =   9
         Top             =   2310
         Width           =   1410
      End
      Begin VB.TextBox TxtNumeroCh 
         Enabled         =   0   'False
         Height          =   315
         Index           =   6
         Left            =   345
         TabIndex        =   10
         Top             =   2640
         Width           =   1410
      End
      Begin VB.TextBox TxtBanco 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   6
         Left            =   1770
         TabIndex        =   24
         Top             =   2640
         Width           =   3435
      End
      Begin VB.TextBox TxtSucursal 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   6
         Left            =   5220
         TabIndex        =   23
         Top             =   2640
         Width           =   1875
      End
      Begin VB.TextBox TxtNumeroCh 
         Enabled         =   0   'False
         Height          =   315
         Index           =   7
         Left            =   345
         TabIndex        =   11
         Top             =   2970
         Width           =   1410
      End
      Begin VB.TextBox TxtBanco 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   7
         Left            =   1770
         TabIndex        =   22
         Top             =   2970
         Width           =   3435
      End
      Begin VB.TextBox TxtSucursal 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   7
         Left            =   5220
         TabIndex        =   21
         Top             =   2970
         Width           =   1875
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Importe"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Index           =   14
         Left            =   7110
         TabIndex        =   83
         Top             =   315
         Width           =   1275
      End
      Begin VB.Label LblAnulada 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "ANULADA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   435
         Left            =   1770
         TabIndex        =   59
         Top             =   4350
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   6975
         X2              =   8460
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   5220
         TabIndex        =   57
         Top             =   4725
         Width           =   1875
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "EFECTIVO"
         Height          =   315
         Index           =   0
         Left            =   5220
         TabIndex        =   56
         Top             =   4320
         Width           =   1875
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Efecto Nº"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Index           =   16
         Left            =   345
         TabIndex        =   40
         Top             =   315
         Width           =   1410
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Banco"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Index           =   17
         Left            =   1770
         TabIndex        =   39
         Top             =   315
         Width           =   3435
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Plaza"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Index           =   18
         Left            =   5220
         TabIndex        =   38
         Top             =   315
         Width           =   1875
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CHEQUES"
         Height          =   315
         Index           =   22
         Left            =   5220
         TabIndex        =   37
         Top             =   3990
         Width           =   1875
      End
   End
   Begin VB.Label lblEstado 
      AutoSize        =   -1  'True
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   89
      Top             =   6195
      Width           =   570
   End
   Begin VB.Label LblDetalle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "-------"
      Height          =   195
      Left            =   4335
      TabIndex        =   87
      Top             =   7275
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
      Height          =   195
      Index           =   7
      Left            =   5220
      TabIndex        =   86
      Top             =   540
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha:"
      Height          =   195
      Index           =   1
      Left            =   3105
      TabIndex        =   19
      Top             =   165
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Banco:"
      Height          =   195
      Index           =   5
      Left            =   1005
      TabIndex        =   18
      Top             =   585
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cuenta:"
      Height          =   195
      Index           =   6
      Left            =   6615
      TabIndex        =   17
      Top             =   555
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Boleta de depósito:"
      Height          =   195
      Index           =   4
      Left            =   150
      TabIndex        =   16
      Top             =   165
      Width           =   1395
   End
End
Attribute VB_Name = "FrmBoletaDeposito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edita As Boolean
Dim Nuevo_NRO As Boolean

Sub Limpio_Grilla()
    For A = 0 To 9
      Me.TxtNumeroCh(A).Text = ""
      Me.TxtBanco(A).Text = ""
      Me.TxtSucursal(A).Text = ""
      Me.TxtValorNominalCh(A).Text = ""
      Me.SumandoCh.Text = ""
      Me.TxtBanNomCor(A).Text = ""
    Next A
End Sub

Private Sub Imprimir_Boleta_Cheque()
'    Dim NroLetras As String
'    Dim VEZ As Integer
'    Dim PRI_RENGLON As Double
'
'    Printer.FontName = "Courier New"
'    Printer.FontSize = 8
'
'    Printer.Orientation = 1
'    Printer.PaperSize = 1
'    'Modificado porque la Impresora Láser NO permite esto
'    'Printer.Height = (20.4 * 567)  '567 es una constante
'    'Printer.Width = (22.3 * 567)
'    Printer.ScaleMode = 7
'
'    Printer.FontBold = True
'
'   For VEZ = 0 To 1
'
'        If VEZ = 0 Then
'           PRI_RENGLON = 0
'        ElseIf VEZ = 1 Then
'           PRI_RENGLON = 7.6
'        End If
'
'        'Nro de Boleta en el Talón
'        P 1, PRI_RENGLON + 0.5
'        PP Trim(Me.TxtBoleta.Text)
'
'        'Nro de Boleta en el Cuerpo
'        P 7, PRI_RENGLON + 0.7
'        PP Trim(Me.TxtBoleta.Text)
'
'        'FECHA DEL TALON
'        'DIA
'        P 0.6, PRI_RENGLON + 2.2
'        PP Format(Day(TxtBolFecha.Text), "00")
'
'        'MES
'        P 1.6, PRI_RENGLON + 2.2
'        PP Format(Month(TxtBolFecha.Text), "00")
'
'        'AÑO
'        P 2.5, PRI_RENGLON + 2.2
'        PP Year(TxtBolFecha.Text)
'
'        'FECHA DEL CUERPO
'        'DIA
'        P 4.3, PRI_RENGLON + 1
'        PP Format(Day(TxtBolFecha.Text), "00")
'
'        'MES
'        P 5.3, PRI_RENGLON + 1
'        PP Format(Month(TxtBolFecha.Text), "00")
'
'        'AÑO
'        P 6.1, PRI_RENGLON + 1
'        PP Year(TxtBolFecha.Text)
'
'        'Importe en Letras
'        NroLetras = LeeNro(Trim(SumandoTODO.Text), 30, 111, "", "-", "*")
'
'        '1º Renglón
'        P 5.2, PRI_RENGLON + 1.8
'        PP Mid(NroLetras, 1, 30)
'
'        '2º Renglón
'        P 3.8, PRI_RENGLON + 2.3
'        PP Mid(NroLetras, 31, 40)
'
'        '2º Renglón
'        P 3.8, PRI_RENGLON + 2.8
'        PP Mid(NroLetras, 71, 40)
'
'        'DOMICILIO
'        P 5, PRI_RENGLON + 3.3
'        PP "Av.Hipólito Yrigoyen 490 - Córdoba"
'
'        'FIRMA
'        P 4.5, PRI_RENGLON + 3.8
'        PP "Por C.P.C.E."
'
'        Dim Renglon As Double
'        Dim Linea As Double
'        Renglon = 0.6
'
'        For I = 0 To 9
'
'           If Trim(TxtNumeroCh(I).Text) <> "" Then
'
'               'EFECTO
'               P 12.2, PRI_RENGLON + Trim(Renglon)
'               PP Trim(TxtNumeroCh(I).Text)
'
'               'BANCO
'               P 14.2, PRI_RENGLON + Trim(Renglon)
'               PP Trim(TxtBanNomCor(I).Text)
'
'               'PLAZA
'               P 16.3, PRI_RENGLON + Trim(Renglon)
'               If Trim(TxtBanNomCor(I).Text) = "PCIA.CBA." Then
'                  'Si es Pcia. de Cba. hay que poner la Sucursal
'                  PP TextoPrevioAlGuion(Me.TxtSucursal(I).Text)
'               Else 'En caso contrario solo el Cód. Postal
'                  PP TextoPostAlGuion(Me.TxtSucursal(I).Text)
'               End If
'
'               'IMPORTE
'               P 18.5, PRI_RENGLON + Trim(Renglon)
'               PP "$" & CompletarConEspaciosIzq(TxtValorNominalCh(I).Text, 11)
'               Renglon = Renglon + 0.43
'
'           End If
'
'        Next I
'
'        'IMPORTE TOTAL DEL TALON
'        P 1, PRI_RENGLON + 4.5
'        PP "$" & CompletarConEspaciosIzq(SumandoTODO.Text, 9)
'
'        'IMPORTE TOTAL DEL CUERPO
'        P 18.5, PRI_RENGLON + 5
'        PP "$" & CompletarConEspaciosIzq(SumandoTODO.Text, 11)
'
'    Next VEZ
'    Printer.EndDoc
'    Me.CmdSalir.SetFocus
End Sub

Private Sub Imprimir_Boleta_Efectivo()
'    Dim NroLetras As String
'    Dim VEZ As Integer
'    Dim PRI_RENGLON As Double
'
'    Printer.FontName = "Courier New"
'    Printer.FontSize = 8
'    Printer.Orientation = 1
'    Printer.PaperSize = 1
'    'Modificado porque la Impresora Láser NO permite esto
'    'Printer.Height = (20.4 * 567)  '567 es una constante
'    'Printer.Width = (22.3 * 567)
'    Printer.ScaleMode = 7
'
'    Printer.FontBold = True
'
'   For VEZ = 0 To 1
'
'        If VEZ = 0 Then
'           PRI_RENGLON = 0
'        ElseIf VEZ = 1 Then
'           PRI_RENGLON = 7.6
'        End If
'
'        'Nro de Boleta en el Talón
'        P 1, PRI_RENGLON + 0.5
'        PP Trim(Me.TxtBoleta.Text)
'
'        'Nro de Boleta en el Cuerpo
'        P 7, PRI_RENGLON + 0.7
'        PP Trim(Me.TxtBoleta.Text)
'
'        'FECHA DEL TALON
'        'DIA
'        P 0.6, PRI_RENGLON + 2.2
'        PP Format(Day(TxtBolFecha.Text), "00")
'
'        'MES
'        P 1.6, PRI_RENGLON + 2.2
'        PP Format(Month(TxtBolFecha.Text), "00")
'
'        'AÑO
'        P 2.5, PRI_RENGLON + 2.2
'        PP Year(TxtBolFecha.Text)
'
'        'CRUZ EN EL CUADRO EN EFECTIVO EN EL TALON
'        Printer.FontSize = 14
'        P 0.5, PRI_RENGLON + 3.2
'        PP "X"
'        Printer.FontSize = 8
'
'        'FECHA DEL CUERPO
'        'DIA
'        P 4.3, PRI_RENGLON + 1
'        PP Format(Day(TxtBolFecha.Text), "00")
'
'        'MES
'        P 5.3, PRI_RENGLON + 1
'        PP Format(Month(TxtBolFecha.Text), "00")
'
'        'AÑO
'        P 6.1, PRI_RENGLON + 1
'        PP Year(TxtBolFecha.Text)
'
'        'CRUZ EN EL CUADRO EN EFECTIVO EN EL CUERPO
'        Printer.FontSize = 14
'        P 11.2, PRI_RENGLON + 1
'        PP "X"
'        Printer.FontSize = 8
'
'        'Importe en Letras
'        NroLetras = LeeNro(Trim(SumandoTODO.Text), 30, 111, "", "-", "*")
'
'        '1º Renglón
'        P 5.2, PRI_RENGLON + 1.8
'        PP Mid(NroLetras, 1, 30)
'
'        '2º Renglón
'        P 3.8, PRI_RENGLON + 2.3
'        PP Mid(NroLetras, 31, 40)
'
'        '2º Renglón
'        P 3.8, PRI_RENGLON + 2.8
'        PP Mid(NroLetras, 71, 40)
'
'        'DOMICILIO
'        P 5, PRI_RENGLON + 3.3
'        PP "Av.Hipólito Yrigoyen 490 - Córdoba"
'
'        'FIRMA
'        P 4.5, PRI_RENGLON + 3.8
'        PP "Por C.P.C.E."
'
'        'EFECTO
'        Printer.FontSize = 12
'        P 12.4, PRI_RENGLON + 0.5
'        PP "EFECTIVO"
'        Printer.FontSize = 8
'
'        'IMPORTE TOTAL DEL TALON
'        P 1, PRI_RENGLON + 4.5
'        PP "$" & CompletarConEspaciosIzq(SumandoTODO.Text, 9)
'
'        'IMPORTE TOTAL DEL CUERPO
'        P 18.5, PRI_RENGLON + 5
'        PP "$" & CompletarConEspaciosIzq(SumandoTODO.Text, 11)
'
'    Next VEZ
'    Printer.EndDoc
'    Me.CmdSalir.SetFocus
End Sub

Private Sub Calculo_Importe()
    'Calculo del Importe de la Boleta de Depósito
    SumandoCh.Text = 0
    For A = 0 To 9
      If TxtNumeroCh(A).Text <> "" Then
         SumandoCh.Text = CDbl(SumandoCh.Text) + CDbl(TxtValorNominalCh(A).Text)
      End If
    Next
    SumandoCh.Text = Format(SumandoCh.Text, "#0.00")
End Sub
    
Private Sub CboBancoBoleta_LostFocus()
   'Consulto las Boletas de Depósito con el Nro. de Boleta, Delegación y Banco
   Me.TxtBanCodInt.Text = CStr(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex))
   Dim snp1 As ADODB.Recordset
   Set snp1 = New ADODB.Recordset
   Dim snp As ADODB.Recordset
   Set snp = New ADODB.Recordset
   
   Edita = False
   
   If ActiveControl.Name <> "CmdSalir" And ActiveControl.Name <> "CmdNuevo" Then
   
   'Significa que NO se imprimió BIEN el Nro y se le sugirió uno nuevo.
   'Nuevo_NRO = True
   
   If Trim(Me.TxtBoleta.Text) <> "" And Trim(Me.TxtBanCodInt.Text) <> "" And Nuevo_NRO = False Then
   
      Limpio_Grilla
        
      'Verifico que exista la boleta de deposito.
      sql = "SELECT BOL_FECHA,BOL_EFECVO,EBO_CODIGO,CTA_NROCTA"
      sql = sql & " FROM BOL_DEPOSITO"
      sql = sql & " WHERE BOL_NUMERO = " & XS(TxtBoleta.Text)
      sql = sql & " AND BAN_CODINT = " & XN(TxtBanCodInt.Text)
      snp.Open sql, DBConn, adOpenStatic, adLockOptimistic
      If snp.RecordCount > 0 Then
             If snp.RecordCount = 1 Then
                'Boleta Anulada
                If Trim(snp!EBO_CODIGO) = 2 Then Me.LblAnulada.Visible = True
                
                TxtBolFecha.Text = snp!BOL_FECHA
                TxtBolFecha.Enabled = False
                
                CboCuentas_GotFocus
                Call BuscaProx(snp!CTA_NROCTA, CboCuentas)
                
                TxtEfvo.Text = Format(snp!BOL_EFECVO, "0.00")
                
                sql = "SELECT C.CHE_NUMERO,C.BAN_CODINT,C.CHE_IMPORT, B.BAN_DESCRI,B.BAN_LOCALIDAD,"
                sql = sql & "B.BAN_SUCURSAL,B.BAN_NOMCOR,B.BAN_CODIGO,B.BAN_CODIGO,B.BAN_BANCO"
                sql = sql & " FROM CHEQUE C, BANCO B"
                sql = sql & " WHERE C.BAN_CODINT = B.BAN_CODINT"
                sql = sql & " AND C.BOL_NUMERO = " & XS(Me.TxtBoleta.Text)
                sql = sql & " AND C.BOL_BAN_CODINT = " & XN(Me.TxtBanCodInt.Text)
                snp1.Open sql, DBConn, adOpenStatic, adLockOptimistic
                If snp1.RecordCount > 0 Then
                   snp1.MoveFirst
                   Index = 0
                   Do While Not snp1.EOF
                      Edita = True
                      Me.TxtNumeroCh(Index).Text = Trim(snp1("CHE_NUMERO"))
                      Me.TxtNumeroCh(Index).Enabled = False
                      Me.TxtBanco(Index).Text = ChkNull(snp1("BAN_DESCRI"))
                      Me.TxtCodInt(Index).Text = Trim(snp1!BAN_CODINT)
                      Me.TxtBanNomCor(Index).Text = ChkNull(snp1!BAN_NOMCOR)
                      Me.TxtSucursal(Index).Text = ChkNull(snp1("BAN_LOCALIDAD")) + " - " + ChkNull(snp1("BAN_SUCURSAL"))
                      Me.TxtValorNominalCh(Index).Text = Format(ChkNull(snp1("CHE_IMPORT")), "#0.00")
                      Calculo_Importe
                      snp1.MoveNext
                      Index = Index + 1
                   Loop
                   TxtNumeroCh(0).Enabled = False
                End If
                snp1.Close
            Else
               'Hay mas de una Boleta de Depósito con este Nro. y de esa Delegación
               MsgBox "Hay más de una Boleta de Depósito con el mismo Nro., en la misma Delegación y del mismo Banco. Verifique!", 16, TIT_MSGBOX
            End If
        Else
            Edita = False
            'No existe ese Nro. de Boleta de Depósito
            Me.TxtBoleta.Enabled = True
            Me.CboBancoBoleta.Enabled = True
            Me.CboCuentas.Enabled = True
            Me.TxtBolFecha.Enabled = True
            
            'If FormLlamado = "CtaCte" Then
            '    If UCase(Trim(Format(Date - 1, "ddd"))) <> "DOM" Then
            '        TxtBolFecha.Text = Format(Date - 1, "dd/mm/yyyy")
            '    Else
            '        TxtBolFecha.Text = Format(Date - 3, "dd/mm/yyyy")
            '    End If
            'ElseIf FormLlamado = "VAL" Then
            '    TxtBolFecha.Text = Format(Date, "dd/mm/yyyy")
            'End If
            Me.TxtEfvo = ""
            LblAnulada.Visible = False
            CmdGrabar.Enabled = True
        End If
        snp.Close
     End If
  End If
End Sub

Private Sub CboCuentas_GotFocus()
    If Trim(CboBancoBoleta.Text) <> "" Then
        CboCuentas.Clear
        Call CargoCtaBancaria(CStr(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex)))
        CboCuentas.ListIndex = 0
    End If
End Sub

Private Sub CargoCtaBancaria(Banco As String)
    Set rec = New ADODB.Recordset
    sql = "SELECT CTA_NROCTA FROM CTA_BANCARIA"
    sql = sql & " WHERE BAN_CODINT=" & XN(Banco)
    sql = sql & " AND CTA_FECCIE IS NULL"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
     Do While rec.EOF = False
         CboCuentas.AddItem Trim(rec!CTA_NROCTA)
         rec.MoveNext
     Loop
    End If
    rec.Close
End Sub

Private Sub CboCuentas_LostFocus()
 'Consulto las Boletas de Depósito con el Nro. de Boleta, Delegación, Banco y Nro. de Cta.
 Dim snp1 As ADODB.Recordset
 Set snp1 = New ADODB.Recordset
 Dim snp As ADODB.Recordset
 Set snp = New ADODB.Recordset
   
 Edita = False
   
 If ActiveControl.Name <> "CmdSalir" And ActiveControl.Name <> "CmdNuevo" Then
   
   'Significa que NO se imprimió BIEN el Nro y se le sugirió uno nuevo.
   'Nuevo_NRO = True
   
   If Trim(Me.TxtBoleta.Text) <> "" And Trim(Me.TxtBanCodInt.Text) <> "" And Trim(Me.CboCuentas.Text) <> "" And Nuevo_NRO = False Then
   
      Limpio_Grilla
        
      'Verifico que exista la boleta de deposito.
      cSQL = "SELECT bol_fecha,bol_efecvo,ebo_codigo FROM bol_deposito " & _
             " WHERE BOL_NUMERO = " & XS(TxtBoleta.Text) & _
               " and BAN_CODINT = " & XN(TxtBanCodInt.Text) & _
               " and CTA_NROCTA = " & XS(CboCuentas.List(Me.CboCuentas.ListIndex))
      snp.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
       If snp.RecordCount > 0 Then
            If snp.RecordCount = 1 Then
                'Boleta Anulada
                If Trim(snp!EBO_CODIGO) = 2 Then Me.LblAnulada.Visible = True
                   
                   
                TxtBolFecha.Text = snp!BOL_FECHA
                TxtBolFecha.Enabled = False
                
                TxtEfvo.Text = Format(snp!BOL_EFECVO, "0.00")
                
                cSQL = "SELECT c.che_numero,c.ban_codint,c.che_import, b.ban_descri,b.ban_localidad," & _
                       "b.ban_sucursal,B.BAN_NOMCOR,B.BAN_CODIGO,b.ban_codigo,b.ban_banco " & _
                       " FROM cheque c,banco b " & _
                       " WHERE c.BAN_CODINT = b.BAN_CODINT " & _
                         " and c.BOL_NUMERO = " & XS(Me.TxtBoleta.Text) & _
                         " and c.CTA_NROCTA = " & XS(Me.CboCuentas.List(Me.CboCuentas.ListIndex)) & _
                         " and c.BOL_BAN_CODINT = " & XN(Me.TxtBanCodInt.Text)
                snp1.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
                If snp1.RecordCount > 0 Then
                   snp1.MoveFirst
                   Index = 0
                   Do While Not snp1.EOF
                      Edita = True
                      Me.TxtNumeroCh(Index).Text = Trim(snp1("che_numero"))
                      Me.TxtNumeroCh(Index).Enabled = False
                      Me.TxtBanco(Index).Text = ChkNull(snp1("ban_descri"))
                      Me.TxtCodInt(Index).Text = Trim(snp1!BAN_CODINT)
                      Me.TxtBanNomCor(Index).Text = ChkNull(snp1!BAN_NOMCOR)
                      Me.TxtSucursal(Index).Text = ChkNull(snp1("ban_localidad")) + " - " + ChkNull(snp1("ban_sucursal"))
                      Me.TxtValorNominalCh(Index).Text = Format(ChkNull(snp1("che_import")), "#0.00")
                      Calculo_Importe
                      snp1.MoveNext
                      Index = Index + 1
                   Loop
                   TxtNumeroCh(0).Enabled = False
                End If
                snp1.Close
            Else
                'Hay mas de una Boleta de Depósito con este Nro.
                MsgBox "Hay más de una Boleta de Depósito con el mismo Nro., en la misma Delegación, del mismo Banco y Depositado en la misma Cta. Bancaria. Verifique!", 16, TIT_MSGBOX
            End If

        Else
            Edita = False
            'No existe ese Nro. de Boleta de Depósito
            Me.TxtBoleta.Enabled = True
            Me.CboBancoBoleta.Enabled = True
            Me.CboCuentas.Enabled = True
            Me.TxtBolFecha.Enabled = True
            
            'If FormLlamado = "CtaCte" Then
            '    If UCase(Trim(Format(Date - 1, "ddd"))) <> "DOM" Then
            '        TxtBolFecha.Text = Format(Date - 1, "dd/mm/yyyy")
            '    Else
            '        TxtBolFecha.Text = Format(Date - 3, "dd/mm/yyyy")
            '    End If
            'ElseIf FormLlamado = "VAL" Then
            '    TxtBolFecha.Text = Format(Date, "dd/mm/yyyy")
            'End If
            
            Me.TxtEfvo = ""
            LblAnulada.Visible = False
            CmdGrabar.Enabled = True
            CboBancoBoleta.SetFocus
        End If
        snp.Close
     End If
  End If
    TxtNumeroCh(0).Enabled = True
    TxtNumeroCh(0).SetFocus
End Sub

Private Sub CmdAnular_Click()

    If Me.TxtBoleta.Text <> "" Then
        If MsgBox("Seguro desea ANULAR la Boleta de Depósito ?", 36, TIT_MSGBOX) = vbYes Then
            
            'ACTUALIZAR LOS CHEQUES
             For i = 0 To 9
              If TxtValorNominalCh(i).Text <> "" Then
                'Inserto en Cheque
                sql = "UPDATE CHEQUE SET BOL_NUMERO = NULL "
                sql = sql & ",BOL_BAN_CODINT = NULL "
                sql = sql & ",CTA_NROCTA = NULL "
                sql = sql & " WHERE CHE_NUMERO = " & XS(Me.TxtNumeroCh(i).Text)
                sql = sql & " AND BAN_CODINT = " & XN(Me.TxtCodInt(i).Text)
                DBConn.Execute sql
                
                'Cambio en Cheque_Estados 1 ES CHEQUES EN CARTERA
                sql = "INSERT INTO CHEQUE_ESTADOS "
                sql = sql & "(CHE_NUMERO,BAN_CODINT,ECH_CODIGO,CES_FECHA,CES_DESCRI)"
                sql = sql & " VALUES (" & XS(Me.TxtNumeroCh(i).Text) & ","
                sql = sql & XN(Me.TxtCodInt(i).Text)
                sql = sql & ", 1," & XDQ(Date) & ",'CHEQUE EN CARTERA')"
                DBConn.Execute sql
              End If
            Next i
            
            'ACTUALIZA LA BOLETA DE DEPOSITO
            sql = "UPDATE BOL_DEPOSITO SET EBO_CODIGO =  2, "
            sql = sql & "BOL_FECEST =" & XDQ(Date)
            sql = sql & "WHERE BOL_NUMERO = " & XS(Me.TxtBoleta.Text)
            DBConn.Execute sql
            
'            'ACTUALIZO EL SALDO DE LA CUENTA BANCARIA
'            sql = "UPDATE CTA_BANCARIA"
'            sql = sql & " SET CTA_SALACT = CTA_SALACT - " & XN(CDbl(SumandoTODO.Text))
'            sql = sql & " Where BAN_CODINT = " & XN(Me.TxtBanCodInt.Text)
'            sql = sql & " AND CTA_NROCTA = " & XS(CboCuentas.List(CboCuentas.ListIndex))
'            DBConn.Execute sql
            
            CmdNuevo_Click
        End If
    End If
End Sub

Private Sub CmdAnular_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   LblDetalle.Caption = "Al Anular la Boleta de Depósito los cheques vuelven a estar en CARTERA y el estado de la Boleta es 2 ANULADO"
End Sub

Private Sub cmdEliminar_Click()
    If Me.TxtBoleta.Text <> "" Then
        If MsgBox("Seguro desea ELIMINAR la Boleta de Depósito ?", 36, TIT_MSGBOX) = vbYes Then
            On Error GoTo CLAVOSE
            DBConn.BeginTrans
            lblEstado.Caption = "Borrando..."
            
            sql = "SELECT EBO_CODIGO"
            sql = sql & " FROM BOL_DEPOSITO "
            sql = sql & " WHERE"
            sql = sql & " BOL_NUMERO = " & XS(Me.TxtBoleta.Text)
            sql = sql & " AND BAN_CODINT=" & XN(Me.TxtBanCodInt.Text)
            sql = sql & " AND CTA_NROCTA=" & XS(CboCuentas.List(CboCuentas.ListIndex))
            rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If rec.EOF = False Then
                If rec!EBO_CODIGO <> 2 Then
'                    'ACTUALIZO EL SALDO DE LA CUENTA BANCARIA
'                    sql = "UPDATE CTA_BANCARIA"
'                    sql = sql & " SET CTA_SALACT = CTA_SALACT - " & XN(CDbl(SumandoTODO.Text))
'                    sql = sql & " WHERE BAN_CODINT = " & XN(Me.TxtBanCodInt.Text)
'                    sql = sql & " AND CTA_NROCTA = " & XS(CboCuentas.List(CboCuentas.ListIndex))
'                    DBConn.Execute sql
                End If
            End If
            rec.Close
            
            'ACTUALIZAR LOS CHEQUES
             For i = 0 To 9
              If TxtValorNominalCh(i).Text <> "" Then
                'Inserto en Cheque
                sql = "UPDATE CHEQUE SET BOL_NUMERO = NULL "
                sql = sql & ",BOL_BAN_CODINT = NULL "
                sql = sql & ",CTA_NROCTA = NULL "
                sql = sql & " WHERE CHE_NUMERO = " & XS(Me.TxtNumeroCh(i).Text)
                sql = sql & " AND BAN_CODINT = " & XN(Me.TxtCodInt(i).Text)
                DBConn.Execute sql
                
                'Cambio en Cheque_Estados 1 ES CHEQUES EN CARTERA
                sql = "INSERT INTO CHEQUE_ESTADOS "
                sql = sql & "(CHE_NUMERO,BAN_CODINT,ECH_CODIGO,CES_FECHA,CES_DESCRI) "
                sql = sql & " VALUES (" & XS(Me.TxtNumeroCh(i).Text)
                sql = sql & "," & XN(Me.TxtCodInt(i).Text)
                sql = sql & ", 1," & XDQ(Date) & ",'CHEQUE EN CARTERA')"
                DBConn.Execute sql
              End If
            Next i
            
            'ACTUALIZA LA BOLETA DE DEPOSITO
            sql = "DELETE FROM BOL_DEPOSITO"
            sql = sql & " WHERE  BOL_NUMERO = " & XS(Me.TxtBoleta.Text)
            DBConn.Execute sql
            
            DBConn.CommitTrans
            lblEstado.Caption = ""
            CmdNuevo_Click
        End If
    End If
    Exit Sub
CLAVOSE:
    If rec.State = 1 Then rec.Close
    DBConn.RollbackTrans
    lblEstado.Caption = ""
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub CmdEliminar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LblDetalle.Caption = " Al Eliminar la Boleta de Depósito los cheques vuelven a estar en CARTERA y la Boleta se puede volver a CARGAR"
End Sub

Private Sub CmdGrabar_Click()
    Dim Minuta As Integer
    Dim snp As ADODB.Recordset
    Set snp = New ADODB.Recordset
    On Error GoTo ErrorTrans
    
    If Not Edita Then
    
        If Me.SumandoTODO.Text <> "" Then
        
            DBConn.BeginTrans
            lblEstado.Caption = "Guardando..."
'            sql = "SELECT MIA_CODIGO FROM CTA_BANCARIA WHERE CTA_NROCTA = " & XS(Me.CboCuentas.List(Me.CboCuentas.ListIndex))
'            snp.Open sql, DBConn, adOpenStatic, adLockOptimistic
'            If snp.RecordCount > 0 Then
'               Minuta = Trim(snp!MIA_CODIGO)
'            End If
'            snp.Close
           
            'Guardo la Boleta de Depósito
             sql = "INSERT INTO BOL_DEPOSITO "
             sql = sql & "(BOL_NUMERO,CTA_NROCTA,BAN_CODINT,BOL_FECHA,"
             sql = sql & " BOL_EFECVO,EBO_CODIGO,BOL_FECEST,BOL_TOTAL)"
             sql = sql & " VALUES ("
             sql = sql & XN(Me.TxtBoleta.Text) & ","
             sql = sql & XS(Me.CboCuentas.List(Me.CboCuentas.ListIndex)) & ","
             sql = sql & XN(Me.TxtBanCodInt.Text) & "," & XDQ(TxtBolFecha.Text) & ","
             sql = sql & XN(TxtEfvo) & ",1," & XDQ(TxtBolFecha.Text) & ","
             sql = sql & XN(SumandoTODO.Text) & ")"
             DBConn.Execute sql
             
             
            For i = 0 To 9
              If TxtValorNominalCh(i).Text <> "" Then
                'Inserto en Cheque
                sql = "UPDATE CHEQUE SET BOL_NUMERO = " & XN(Me.TxtBoleta.Text)
                sql = sql & ",BOL_BAN_CODINT = " & XN(Me.TxtBanCodInt.Text)
                sql = sql & ",CTA_NROCTA = " & XS(Me.CboCuentas.List(Me.CboCuentas.ListIndex))
                sql = sql & " WHERE CHE_NUMERO = " & XS(Me.TxtNumeroCh(i).Text)
                sql = sql & " AND BAN_CODINT = " & XN(Me.TxtCodInt(i).Text)
                DBConn.Execute sql
                
                'Cambio en Cheque_Estados 2 ES CHEQUES DEPOSITADO
                sql = "INSERT INTO CHEQUE_ESTADOS"
                sql = sql & "(ECH_CODIGO,BAN_CODINT,CHE_NUMERO,CES_FECHA,CES_DESCRI) "
                sql = sql & " VALUES ( 2," & XN(Me.TxtCodInt(i).Text) & ","
                sql = sql & XS(Me.TxtNumeroCh(i).Text) & ","
                sql = sql & XDQ(Date) & ","
                sql = sql & "'CHEQUE DEPOSITADO')"
                DBConn.Execute sql
              End If
            Next i
'            'ACTUALIZO EL SALDO DE LA CUENTA BANCARIA
'            sql = "UPDATE CTA_BANCARIA"
'            sql = sql & " SET CTA_SALACT = CTA_SALACT + " & XN(CDbl(SumandoTODO.Text))
'            sql = sql & " Where BAN_CODINT = " & XN(Me.TxtBanCodInt.Text)
'            sql = sql & " And CTA_NROCTA = " & XS(CboCuentas.List(CboCuentas.ListIndex))
'            DBConn.Execute sql
            
            DBConn.CommitTrans
            lblEstado.Caption = ""
        End If
        CmdNuevo_Click
        MousePointer = 0
    End If
    Exit Sub
    
ErrorTrans:
  Beep
  DBConn.RollbackTrans
  lblEstado.Caption = ""
  Screen.MousePointer = 0
  MsgBox "Error intentando Grabar la Boleta de Depósito. " & Chr(13) & Err.Description, 16, TIT_MSGBOX
End Sub


Private Sub CmdGrabar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     LblDetalle.Caption = "Graba la Boleta de Depósito."
End Sub

Private Sub cmdImprimir_Click()
 If Me.TxtBoleta.Text <> "" And Me.SumandoTODO.Text <> "" Then
    If TxtEfvo.Text = "" Then
       Imprimir_Boleta_Cheque
    Else
       Imprimir_Boleta_Efectivo
    End If
    If MsgBox("Verifique la Impresión. Desea GRABAR la Boleta de Depósito ?", 36, TIT_MSGBOX) = vbYes Then
       Nuevo_NRO = False
       CmdGrabar_Click
    Else
       Nuevo_NRO = True
       SelecTexto Me.TxtBoleta
       TxtBoleta.SetFocus
    End If
 Else
    MsgBox "Ingrese el Nro. de Boleta a Imprimir", 16, TIT_MSGBOX
    TxtBoleta.SetFocus
 End If
End Sub

Private Sub CmdImprimir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LblDetalle.Caption = "Imprime la Boleta de Depósito."
End Sub

Private Sub CmdNuevo_Click()
  Me.TxtBoleta.Enabled = True
  Me.CboBancoBoleta.Enabled = True
  Me.CboCuentas.Enabled = True
  Me.TxtBolFecha.Enabled = True
  
  Me.TxtBoleta.Text = ""
  Me.CboBancoBoleta.ListIndex = 0
  Me.TxtBanCodInt.Text = Trim(Mid(CboBancoBoleta, 102, 5))
  Me.CboCuentas.Clear
  
'  If FormLlamado = "CtaCte" Then
'     If UCase(Trim(Format(Date - 1, "ddd"))) <> "DOM" Then
'        TxtBolFecha.Text = Format(Date - 1, "dd/mm/yyyy")
'     Else
'        TxtBolFecha.Text = Format(Date - 3, "dd/mm/yyyy")
'     End If
'  ElseIf FormLlamado = "VAL" Then
     TxtBolFecha.Text = Format(Date, "dd/mm/yyyy")
'  End If
    
  For A = 0 To 9
    Me.TxtNumeroCh(A).Text = ""
    Me.TxtBanco(A).Text = ""
    Me.TxtSucursal(A).Text = ""
    Me.TxtValorNominalCh(A).Text = ""
    Me.SumandoCh.Text = ""
    Me.TxtBanNomCor(A).Text = ""
  Next A
  
  Me.TxtEfvo = ""
  Me.TxtBoleta.SetFocus
  LblAnulada.Visible = False
End Sub

Private Sub CmdNuevo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LblDetalle.Caption = "Nueva Boleta de Depósito."
End Sub

Private Sub CmdSalir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LblDetalle.Caption = "Sale del Formulario de la Boleta de Depósito."
End Sub

Private Sub SumandoCh_Change()
    SumandoTODO = Format(CDbl(IIf(Trim(SumandoCh) = "", 0, Trim(SumandoCh))) + CDbl(IIf(Trim(TxtEfvo) = "", 0, Trim(TxtEfvo))), "#0.00")
End Sub

Private Sub TxtBoleta_GotFocus()
    lblEstado.Caption = "<F1> Buscar Boleta de Depósito"
End Sub

Private Sub TxtBoleta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        Call ConsultaBoleta.Parametros(Me, TxtBoleta, TxtBolFecha, CboBancoBoleta, CboCuentas)
        ConsultaBoleta.Show vbModal
    End If
End Sub

Private Sub TxtBoleta_LostFocus()
   'Consulto las Boletas de Depósito con el Nro. de Boleta
   Dim snp1 As ADODB.Recordset
   Set snp1 = New ADODB.Recordset
   Dim snp As ADODB.Recordset
   Set snp = New ADODB.Recordset
   Edita = False
   'Significa que NO se imprimió BIEN el Nro y se le sugirió uno nuevo.
   'Nuevo_NRO = True
   lblEstado.Caption = ""
   If ActiveControl.Name <> "CmdSalir" And ActiveControl.Name <> "CmdNuevo" Then
   
      If Trim(Me.TxtBoleta.Text) <> "" And Nuevo_NRO = False Then
   
         Limpio_Grilla
           
         'Verifico que exista la boleta de deposito.
         sql = "SELECT BOL_FECHA, BOL_EFECVO, EBO_CODIGO,"
         sql = sql & " CTA_NROCTA,BAN_CODINT"
         sql = sql & " FROM BOL_DEPOSITO"
         sql = sql & " WHERE BOL_NUMERO = " & XS(TxtBoleta.Text)
         snp.Open sql, DBConn, adOpenStatic, adLockOptimistic
         If snp.EOF = False Then
              If snp.RecordCount = 1 Then 'Hay solo una Boleta con este Nro.
                   'Boleta Anulada
                   If Trim(snp!EBO_CODIGO) = 2 Then Me.LblAnulada.Visible = True
                   Call BuscaCodigoProxItemData(CInt(snp!BAN_CODINT), CboBancoBoleta)
                   CboCuentas_GotFocus
                   CboBancoBoleta_LostFocus
                   Call BuscaProx(snp!CTA_NROCTA, CboCuentas)
                   
                   TxtBolFecha.Text = snp!BOL_FECHA
                   'TxtBolFecha.Enabled = True
                
                   TxtEfvo.Text = Format(ChkNull(snp!BOL_EFECVO), "0.00")
                   
                   sql = "SELECT C.CHE_NUMERO, C.BAN_CODINT, C.CHE_IMPORT, B.BAN_DESCRI, B.BAN_LOCALIDAD,"
                   sql = sql & "B.BAN_SUCURSAL,B.BAN_NOMCOR,B.BAN_CODIGO,B.BAN_BANCO"
                   sql = sql & " FROM CHEQUE C, BANCO B"
                   sql = sql & " WHERE C.BAN_CODINT = B.BAN_CODINT"
                   sql = sql & " AND C.BOL_NUMERO = " & XS(Me.TxtBoleta.Text)
                   snp1.Open sql, DBConn, adOpenStatic, adLockOptimistic
                   If snp1.EOF = False Then
                      snp1.MoveFirst
                      Index = 0
                      Do While Not snp1.EOF
                         Edita = True
                         Me.TxtNumeroCh(Index).Text = Trim(snp1("CHE_NUMERO"))
                         Me.TxtNumeroCh(Index).Enabled = False
                         Me.TxtBanco(Index).Text = ChkNull(snp1("BAN_DESCRI"))
                         Me.TxtCodInt(Index).Text = Trim(snp1!BAN_CODINT)
                         Me.TxtBanNomCor(Index).Text = ChkNull(snp1!BAN_NOMCOR)
                         Me.TxtSucursal(Index).Text = ChkNull(snp1("BAN_LOCALIDAD")) + " - " + ChkNull(snp1("BAN_SUCURSAL"))
                         Me.TxtValorNominalCh(Index).Text = Format(ChkNull(snp1("CHE_IMPORT")), "#0.00")
                         Calculo_Importe
                         snp1.MoveNext
                         Index = Index + 1
                      Loop
                      TxtNumeroCh(0).Enabled = False
                   End If
                   snp1.Close
                Else
                   'Hay mas de una Boleta de Depósito con este Nro.
                   MsgBox "Hay más de una Boleta de Depósito con el mismo Nro. Verifique!", 16, TIT_MSGBOX
                End If
         Else
                Edita = False
                'No existe ese Nro. de Boleta de Depósito
                Me.TxtBoleta.Enabled = True
                Me.CboBancoBoleta.Enabled = True
                Me.CboCuentas.Enabled = True
                Me.TxtBolFecha.Enabled = True
                
                'If FormLlamado = "CtaCte" Then
                '    If UCase(Trim(Format(Date - 1, "ddd"))) <> "DOM" Then
                '        TxtBolFecha.Text = Format(Date - 1, "dd/mm/yyyy")
                '    Else
                '        TxtBolFecha.Text = Format(Date - 3, "dd/mm/yyyy")
                '    End If
                'ElseIf FormLlamado = "VAL" Then
                '    If TxtBolFecha.Text = "" Then TxtBolFecha.Text = Format(Date, "dd/mm/yyyy")
                'End If
            
                Me.TxtEfvo = ""
                LblAnulada.Visible = False
                CmdGrabar.Enabled = True
                TxtBolFecha.SetFocus
         End If
         snp.Close
      End If
   End If
End Sub

Private Sub TxtEfvo_Change()
    SumandoTODO = Format(CDbl(IIf(Trim(SumandoCh) = "", 0, Trim(SumandoCh))) + CDbl(IIf(Trim(TxtEfvo) = "", 0, Trim(TxtEfvo))), "0.00")
End Sub

Private Sub TxtEfvo_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(Me.TxtEfvo.Text, KeyAscii)
End Sub

Private Sub TxtNumeroCh_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        Call ConsultaCheque.Parametros(Me, TxtNumeroCh(Index))
        ConsultaCheque.Show vbModal
    End If
End Sub

Private Sub CmdSalir_Click()
    Unload Me
    Set FrmBoletaDeposito = Nothing
End Sub

Private Sub TxtNumeroCh_GotFocus(Index As Integer)
    With TxtNumeroCh(Index)
    .SelStart = 0
    .SelLength = Len(.Text)
    End With
    lblEstado.Caption = "<F1> Buscar Cheques en Cartera"
End Sub

Private Sub TxtNumeroCh_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub TxtBoleta_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub TxtBolFecha_LostFocus()
    If Me.TxtBolFecha.Text = "" Then Me.TxtBolFecha.Text = Format(Date, "dd/mm/yyyy")
End Sub

Private Sub TxtNumeroCh_LostFocus(Index As Integer)
 
 If ActiveControl.Name <> "CmdSalir" And ActiveControl.Name <> "CmdNuevo" Then
    lblEstado.Caption = ""
   If Trim(TxtNumeroCh(Index).Enabled) = True And Trim(TxtNumeroCh(Index).Text) <> "" Then
      
      If Len(TxtNumeroCh(Index).Text) < 10 Then TxtNumeroCh(Index).Text = CompletarConCeros(TxtNumeroCh(Index).Text, 10)
      
      INDICE = Index
      
      If Trim(Index) > 0 Then
        For i = 0 To 9
           If Trim(Index) <> Trim(i) Then
               If Trim(TxtNumeroCh(i).Text) = Trim(TxtNumeroCh(Index).Text) Then
                  MsgBox "El Nro. de Cheque está repetido. Verifique!", 16, TIT_MSGBOX
                  TxtNumeroCh(Index).Text = ""
                  TxtNumeroCh(Index).SetFocus
                  Exit Sub
               End If
           End If
        Next i
      End If
      
      Index = INDICE
      
      Dim snp As ADODB.Recordset
      Set snp = New ADODB.Recordset
      
      'Estado 1 = En Cartera
      sql = "SELECT B.BAN_CODINT,C.CHE_IMPORT,B.BAN_DESCRI,"
      sql = sql & "B.BAN_LOCALIDAD,B.BAN_SUCURSAL,B.BAN_NOMCOR,B.BAN_BANCO,B.BAN_CODIGO"
      sql = sql & " FROM CHEQUE C, BANCO B, CHEQUE_ESTADOS CE"
      sql = sql & " WHERE C.BAN_CODINT = B.BAN_CODINT"
      sql = sql & " AND C.BAN_CODINT = CE.BAN_CODINT"
      sql = sql & " AND C.CHE_NUMERO=CE.CHE_NUMERO"
      sql = sql & " AND CE.ECH_CODIGO = 1 "
      sql = sql & " AND C.CHE_NUMERO = " + XS(Me.TxtNumeroCh(Index).Text)
      snp.Open sql, DBConn, adOpenStatic, adLockOptimistic
      If snp.RecordCount > 0 Then
         
         Me.TxtBanco(Index).Text = ChkNull(snp("BAN_DESCRI"))
         Me.TxtCodInt(Index).Text = Trim(snp!BAN_CODINT)
         'Me.TxtSucursal(Index).Text = ChkNull(snp("ban_banco")) + " - " + ChkNull(snp("ban_localidad")) + " - " + ChkNull(snp("ban_sucursal")) + " - " + ChkNull(snp("ban_codigo"))
         'Me.TxtSUCURSAL(Index).Text = ChkNull(snp("ban_localidad")) + " - " + ChkNull(snp("ban_sucursal"))
         Me.TxtSucursal(Index).Text = ChkNull(snp("BAN_SUCURSAL")) + " - " + ChkNull(snp("BAN_CODIGO"))
         'Me.TxtPlaza(Index).Text = ChkNull(snp("BAN_CODIGO"))
         Me.TxtBanNomCor(Index).Text = ChkNull(snp("BAN_NOMCOR"))
         Me.TxtValorNominalCh(Index).Text = Format(ChkNull(snp("CHE_IMPORT")), "#0.00")
         snp.Close
         Calculo_Importe
         If Trim(Index) < 9 Then
            TxtNumeroCh(Index + 1).Enabled = True
            TxtNumeroCh(Index + 1).SetFocus
         End If
         Exit Sub
      Else
         snp.Close
         MsgBox "Nro. de Cheque NO disponible.!", 16, TIT_MSGBOX
         TxtNumeroCh(Index).Text = ""
         TxtNumeroCh(Index).SetFocus
      End If
    End If
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If Not (Me.ActiveControl.Name = "TxtNumeroCh" And TxtNumeroCh(Index).Text = "") Then
      If KeyAscii = vbKeyReturn Then SendKeys ("{TAB}")
   End If
End Sub

Private Sub Form_Load()

    Nuevo_NRO = False
    
    Screen.MousePointer = 1
    Set rec = New ADODB.Recordset
    
    
    CboBancoBoleta.Clear
    CargoBanco
'    sql = "SELECT B.BAN_DESCRI,B.BAN_CODINT FROM BANCO B WHERE " & _
'    "(SELECT COUNT(*) FROM CTA_BANCARIA C WHERE B.BAN_CODINT = C.BAN_CODINT) > 0"
'    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'    If Rec.RecordCount > 0 Then
'        Rec.MoveFirst
'        Do While Not Rec.EOF
'            CboBancoBoleta.AddItem Trim(Rec!ban_descri) & Space(100 - Len(Trim(Rec!ban_descri))) & Chr(9) & Trim(Rec!BAN_CODINT)
'            Rec.MoveNext
'        Loop
'        CboBancoBoleta.ListIndex = 0
'    End If
'    Rec.Close
    
    Me.TxtBanCodInt.Text = Trim(Mid(CboBancoBoleta.List(CboBancoBoleta.ListIndex), 102, 10))
    Edita = True
    
    Call Centrar_pantalla(Me)
    lblEstado.Caption = ""
    CboCuentas.Clear
    
    If FormLlamado = "CtaCte" Then
        If UCase(Trim(Format(Date - 1, "ddd"))) <> "DOM" Then
            TxtBolFecha.Text = Format(Date - 1, "dd/mm/yyyy")
        Else
            TxtBolFecha.Text = Format(Date - 3, "dd/mm/yyyy")
        End If
    ElseIf FormLlamado = "VAL" Then
        TxtBolFecha.Text = Format(Date, "dd/mm/yyyy")
    End If
    
End Sub

Private Sub CargoBanco()
    sql = "SELECT DISTINCT B.BAN_CODINT, B.BAN_DESCRI"
    sql = sql & " FROM BANCO B, CTA_BANCARIA CB"
    sql = sql & " WHERE B.BAN_CODINT=CB.BAN_CODINT"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            CboBancoBoleta.AddItem Trim(rec!BAN_DESCRI)
            CboBancoBoleta.ItemData(CboBancoBoleta.NewIndex) = Trim(rec!BAN_CODINT)
            rec.MoveNext
        Loop
        CboBancoBoleta.ListIndex = 0
    End If
    rec.Close
End Sub

