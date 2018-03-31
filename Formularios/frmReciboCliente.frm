VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{5F09B5DF-6F4D-11D2-8355-4854E82A9183}#15.0#0"; "Fecha32.ocx"
Begin VB.Form frmReciboCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recibo de Cliente"
   ClientHeight    =   6390
   ClientLeft      =   0
   ClientTop       =   750
   ClientWidth     =   11715
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   11715
   Begin Crystal.CrystalReport Rep 
      Left            =   1965
      Top             =   5925
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   450
      Left            =   8070
      TabIndex        =   8
      Top             =   5910
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Cancelar"
      Height          =   450
      Left            =   9870
      TabIndex        =   9
      Top             =   5910
      Width           =   870
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      Height          =   450
      Left            =   8970
      TabIndex        =   7
      Top             =   5910
      Width           =   870
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   450
      Left            =   10755
      TabIndex        =   10
      Top             =   5910
      Width           =   870
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   5835
      Left            =   15
      TabIndex        =   11
      Top             =   45
      Width           =   11640
      _ExtentX        =   20532
      _ExtentY        =   10292
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   512
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Datos"
      TabPicture(0)   =   "frmReciboCliente.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "tabComprobantes"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "tabValores"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FrameRemito"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "FrameRecibo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "frmReciboCliente.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "GrdModulos"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "frameBuscar"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin TabDlg.SSTab tabComprobantes 
         Height          =   3840
         Left            =   -74895
         TabIndex        =   34
         Top             =   1935
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   6773
         _Version        =   393216
         Style           =   1
         Tabs            =   1
         TabsPerRow      =   2
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "C&omprobantes Pendientes"
         TabPicture(0)   =   "frmReciboCliente.frx":0038
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame5"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         Begin VB.Frame Frame5 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3480
            Left            =   75
            TabIndex        =   37
            Top             =   315
            Width           =   5550
            Begin VB.TextBox txtSaldo 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   3855
               TabIndex        =   39
               Top             =   2625
               Width           =   1275
            End
            Begin VB.TextBox txtImporteApagar 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FF0000&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   420
               Left            =   3855
               TabIndex        =   38
               Top             =   3000
               Width           =   1290
            End
            Begin MSFlexGridLib.MSFlexGrid GrillaAplicar 
               Height          =   2415
               Left            =   105
               TabIndex        =   40
               Top             =   195
               Width           =   5325
               _ExtentX        =   9393
               _ExtentY        =   4260
               _Version        =   393216
               Cols            =   7
               FixedCols       =   0
               RowHeightMin    =   300
               BackColorSel    =   16761024
               AllowBigSelection=   -1  'True
               FocusRect       =   0
               HighLight       =   2
               SelectionMode   =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Saldo:"
               Height          =   195
               Left            =   3225
               TabIndex        =   42
               Top             =   2670
               Width           =   450
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Importe a Pagar:"
               Height          =   195
               Left            =   2445
               TabIndex        =   41
               Top             =   2985
               Width           =   1230
            End
         End
      End
      Begin TabDlg.SSTab tabValores 
         Height          =   3840
         Left            =   -69150
         TabIndex        =   6
         Top             =   1935
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   6773
         _Version        =   393216
         Style           =   1
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "&Valores"
         TabPicture(0)   =   "frmReciboCliente.frx":0054
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Moneda"
         TabPicture(1)   =   "frmReciboCliente.frx":0070
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame4"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Valores a Cuenta"
         TabPicture(2)   =   "frmReciboCliente.frx":008C
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame6"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         Begin VB.Frame Frame4 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3480
            Left            =   -74880
            TabIndex        =   55
            Top             =   315
            Width           =   5532
            Begin VB.CommandButton cmdAgregarEfectivo 
               Caption         =   "Agregar"
               Height          =   345
               Left            =   2175
               TabIndex        =   61
               Top             =   705
               Width           =   885
            End
            Begin VB.TextBox txtEftImporte 
               Height          =   330
               Left            =   1125
               TabIndex        =   60
               Top             =   705
               Width           =   1005
            End
            Begin VB.ComboBox cboMoneda 
               Height          =   315
               Left            =   1125
               Style           =   2  'Dropdown List
               TabIndex        =   59
               Top             =   345
               Width           =   1950
            End
            Begin VB.TextBox txtTotalEfectivo 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FF0000&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   405
               Left            =   2745
               Locked          =   -1  'True
               TabIndex        =   58
               Top             =   2460
               Width           =   1290
            End
            Begin VB.CommandButton cmdAceptarMoneda 
               Caption         =   "Aceptar"
               Height          =   360
               Left            =   2115
               TabIndex        =   57
               Top             =   2925
               Width           =   960
            End
            Begin VB.CommandButton cmdCancelarMoneda 
               Caption         =   "Cancelar"
               Height          =   360
               Left            =   3090
               TabIndex        =   56
               Top             =   2925
               Width           =   960
            End
            Begin MSFlexGridLib.MSFlexGrid GrillaEfectivo 
               Height          =   1320
               Left            =   1110
               TabIndex        =   65
               Top             =   1110
               Width           =   2925
               _ExtentX        =   5159
               _ExtentY        =   2328
               _Version        =   393216
               Cols            =   3
               FixedCols       =   0
               RowHeightMin    =   300
               BackColorSel    =   16761024
               AllowBigSelection=   -1  'True
               FocusRect       =   0
               HighLight       =   2
               SelectionMode   =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Importe:"
               Height          =   195
               Index           =   2
               Left            =   420
               TabIndex        =   64
               Top             =   765
               Width           =   630
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "Moneda:"
               Height          =   195
               Left            =   420
               TabIndex        =   63
               Top             =   390
               Width           =   630
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "Total:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   2190
               TabIndex        =   62
               Top             =   2475
               Width           =   405
            End
         End
         Begin VB.Frame Frame6 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3480
            Left            =   -74925
            TabIndex        =   47
            Top             =   315
            Width           =   5535
            Begin VB.CommandButton cmdAgregarACta 
               Caption         =   "A&gregar"
               Height          =   420
               Left            =   3285
               TabIndex        =   51
               Top             =   2865
               Width           =   1065
            End
            Begin VB.TextBox txtSaldoACta 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1455
               Locked          =   -1  'True
               TabIndex        =   50
               Top             =   2535
               Width           =   1185
            End
            Begin VB.TextBox txtImporteACta 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1455
               TabIndex        =   49
               Top             =   2925
               Width           =   1185
            End
            Begin VB.CommandButton cmaAceptarACta 
               Caption         =   "A&ceptar"
               Height          =   420
               Left            =   4365
               TabIndex        =   48
               Top             =   2865
               Width           =   1065
            End
            Begin MSFlexGridLib.MSFlexGrid GrillaAFavor 
               Height          =   2175
               Left            =   150
               TabIndex        =   52
               Top             =   210
               Width           =   5235
               _ExtentX        =   9234
               _ExtentY        =   3836
               _Version        =   393216
               Cols            =   7
               FixedCols       =   0
               RowHeightMin    =   300
               BackColorSel    =   16761024
               AllowBigSelection=   -1  'True
               FocusRect       =   0
               HighLight       =   2
               SelectionMode   =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               Caption         =   "Saldo:"
               Height          =   195
               Left            =   975
               TabIndex        =   54
               Top             =   2595
               Width           =   450
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               Caption         =   "Importe:"
               Height          =   195
               Left            =   795
               TabIndex        =   53
               Top             =   2970
               Width           =   630
            End
         End
         Begin VB.Frame Frame2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3480
            Left            =   75
            TabIndex        =   43
            Top             =   315
            Width           =   5550
            Begin VB.TextBox txtTotalValores 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FF0000&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   405
               Left            =   4065
               Locked          =   -1  'True
               TabIndex        =   44
               TabStop         =   0   'False
               Top             =   3000
               Width           =   1290
            End
            Begin MSFlexGridLib.MSFlexGrid grillaValores 
               Height          =   2430
               Left            =   90
               TabIndex        =   45
               Top             =   225
               Width           =   5310
               _ExtentX        =   9366
               _ExtentY        =   4286
               _Version        =   393216
               Cols            =   7
               FixedCols       =   0
               RowHeightMin    =   300
               BackColorSel    =   16761024
               FocusRect       =   0
               HighLight       =   0
               SelectionMode   =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Label LblDineroaCta 
               AutoSize        =   -1  'True
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   285
               Left            =   165
               TabIndex        =   66
               Top             =   2670
               Width           =   75
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               Caption         =   "Total:"
               Height          =   195
               Left            =   3540
               TabIndex        =   46
               Top             =   3000
               Width           =   420
            End
         End
      End
      Begin VB.Frame FrameRemito 
         Caption         =   "Cliente..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1545
         Left            =   -70395
         TabIndex        =   19
         Top             =   360
         Width           =   6900
         Begin VB.TextBox txtDomici 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1350
            MaxLength       =   50
            TabIndex        =   35
            Top             =   840
            Width           =   5145
         End
         Begin VB.TextBox txtCliRazSoc 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2160
            MaxLength       =   50
            TabIndex        =   5
            Tag             =   "Descripción"
            Top             =   480
            Width           =   4335
         End
         Begin VB.TextBox txtCodCliente 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1350
            MaxLength       =   40
            TabIndex        =   4
            Top             =   480
            Width           =   780
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio:"
            Height          =   195
            Left            =   600
            TabIndex        =   36
            Top             =   870
            Width           =   675
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Recibimos de:"
            Height          =   195
            Left            =   285
            TabIndex        =   32
            Top             =   540
            Width           =   990
         End
      End
      Begin VB.Frame FrameRecibo 
         Caption         =   "Recibo..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1545
         Left            =   -74895
         TabIndex        =   25
         Top             =   360
         Width           =   4485
         Begin VB.TextBox txtNroSucursal 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1305
            MaxLength       =   4
            TabIndex        =   1
            Top             =   570
            Width           =   555
         End
         Begin VB.TextBox txtNroRecibo 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1890
            MaxLength       =   8
            TabIndex        =   2
            Top             =   570
            Width           =   1065
         End
         Begin VB.ComboBox cboRecibo 
            Height          =   315
            Left            =   1305
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   225
            Width           =   2400
         End
         Begin FechaCtl.Fecha FechaRecibo 
            Height          =   285
            Left            =   1305
            TabIndex        =   3
            Top             =   945
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin VB.Label lblEstadoRecibo 
            AutoSize        =   -1  'True
            Caption         =   "EST. RECIBO"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   1305
            TabIndex        =   30
            Top             =   1275
            Width           =   1005
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Left            =   660
            TabIndex        =   29
            Top             =   1260
            Width           =   555
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            Height          =   195
            Left            =   600
            TabIndex        =   28
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   720
            TabIndex        =   27
            Top             =   945
            Width           =   495
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   855
            TabIndex        =   26
            Top             =   240
            Width           =   360
         End
      End
      Begin VB.Frame frameBuscar 
         Caption         =   "Buscar por..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         Left            =   285
         TabIndex        =   20
         Top             =   480
         Width           =   11115
         Begin VB.TextBox txtCliente 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2805
            MaxLength       =   40
            TabIndex        =   12
            Top             =   300
            Width           =   750
         End
         Begin VB.TextBox txtDesCli 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3585
            MaxLength       =   50
            TabIndex        =   13
            Tag             =   "Descripción"
            Top             =   300
            Width           =   3990
         End
         Begin VB.CommandButton CmdBuscAprox 
            Caption         =   "Buscar"
            Height          =   450
            Left            =   8280
            MaskColor       =   &H80000006&
            Picture         =   "frmReciboCliente.frx":00A8
            TabIndex        =   17
            ToolTipText     =   "Buscar "
            Top             =   915
            UseMaskColor    =   -1  'True
            Width           =   2085
         End
         Begin VB.ComboBox cboRecibo1 
            Height          =   315
            Left            =   2805
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   975
            Width           =   2400
         End
         Begin FechaCtl.Fecha FechaHasta 
            Height          =   315
            Left            =   5310
            TabIndex        =   15
            Top             =   660
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin FechaCtl.Fecha FechaDesde 
            Height          =   330
            Left            =   2805
            TabIndex        =   14
            Top             =   660
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   582
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   3
            Left            =   2145
            TabIndex        =   24
            Top             =   345
            Width           =   555
         End
         Begin VB.Label lblFechaDesde 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde:"
            Height          =   195
            Left            =   1710
            TabIndex        =   23
            Top             =   690
            Width           =   990
         End
         Begin VB.Label lblFechaHasta 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Left            =   4260
            TabIndex        =   22
            Top             =   705
            Width           =   960
         End
         Begin VB.Label lbltipoFac 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Recibo:"
            Height          =   195
            Left            =   1815
            TabIndex        =   21
            Top             =   1005
            Width           =   885
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   3645
         Left            =   255
         TabIndex        =   18
         Top             =   2040
         Width           =   11160
         _ExtentX        =   19685
         _ExtentY        =   6429
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorSel    =   16761024
         AllowBigSelection=   -1  'True
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ordenado por :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   -74820
         TabIndex        =   31
         Top             =   570
         Width           =   1065
      End
   End
   Begin VB.Label lblEstado 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   135
      TabIndex        =   33
      Top             =   5985
      Width           =   660
   End
End
Attribute VB_Name = "frmReciboCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim TotFac As Double
Dim Estado As Integer
Dim mBorroTransfe As Boolean
Dim mImprimoRecibo As Boolean
 
Private Function SumaGrilla(Grilla As MSFlexGrid, COLUMNA As Integer) As String
    Dim Suma As Double
    Suma = 0
    For i = 1 To Grilla.Rows - 1
        Suma = Suma + CDbl(Chk0(Grilla.TextMatrix(i, COLUMNA)))
    Next
    SumaGrilla = Valido_Importe(CStr(Suma))
End Function

Private Sub cmdImprimir_Click()
    If txtCodCliente.Text = "" Or GrillaAplicar.Rows = 1 Then Exit Sub
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Buscando Recibo..."

    SQL = "DELETE FROM TMP_RECIBO_CLIENTE"
    DBConn.Execute SQL
    i = 1
    
    ReciboFacturas
    ReciboComprobante
    ReciboCheques
    ReciboMoneda

    DBConn.Execute "DELETE FROM TMP_RESUMEN_CUENTA_BANCO"
    Rep.WindowState = crptMaximized
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & SERVIDOR
    Rep.SelectionFormula = ""

    Rep.WindowTitle = "Recibo"
    Rep.ReportFileName = DRIVE & DirReport & "rptrecibo.rpt"
    
    If mImprimoRecibo = True Then
        'MANDO RECIBO A PANTALLA
        Rep.Destination = crptToWindow
    Else
        'MANDO RECIBO A IMPRESORA
        Rep.Destination = crptToPrinter
    End If
    
    Rep.Action = 1
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    Rep.SelectionFormula = ""
End Sub

Private Sub ReciboFacturas()
    Set Rec1 = New ADODB.Recordset
    'BUSCO FACTURAS_PROVEEDOR
    SQL = "SELECT C.CLI_RAZSOC, C.CLI_DOMICI, L.LOC_DESCRI, PR.PRO_DESCRI, C.CLI_CUIT, C.CLI_INGBRU"
    SQL = SQL & ",CI.IVA_DESCRI, TC.TCO_ABREVIA, FR.FCL_SUCURSAL, FR.FCL_NUMERO, FR.FCL_FECHA ,F.FCL_TOTAL, FR.REC_IMPORTE, R.REC_TOTAL"
    SQL = SQL & " FROM CLIENTE C, RECIBO_CLIENTE R ,CONDICION_IVA CI ,LOCALIDAD L"
    SQL = SQL & " , PROVINCIA PR, TIPO_COMPROBANTE TC, FACTURAS_RECIBO_CLIENTE FR,"
    SQL = SQL & " FACTURA_CLIENTE F"
    SQL = SQL & " WHERE R.REC_NUMERO=" & XN(txtNroRecibo.Text)
    SQL = SQL & " AND R.REC_SUCURSAL=" & XN(txtNroSucursal.Text)
    SQL = SQL & " AND R.TCO_CODIGO=" & XN(cboRecibo.ItemData(cboRecibo.ListIndex))
    SQL = SQL & " AND R.REC_NUMERO=FR.REC_NUMERO"
    SQL = SQL & " AND R.REC_SUCURSAL=FR.REC_SUCURSAL"
    SQL = SQL & " AND R.TCO_CODIGO=FR.TCO_CODIGO"
    SQL = SQL & " AND R.CLI_CODIGO=C.CLI_CODIGO"
    SQL = SQL & " AND C.LOC_CODIGO=L.LOC_CODIGO"
    SQL = SQL & " AND C.PAI_CODIGO=L.PAI_CODIGO"
    SQL = SQL & " AND C.PRO_CODIGO=L.PRO_CODIGO"
    SQL = SQL & " AND L.PRO_CODIGO=PR.PRO_CODIGO"
    SQL = SQL & " AND L.PAI_CODIGO=PR.PAI_CODIGO"
    SQL = SQL & " AND C.IVA_CODIGO=CI.IVA_CODIGO"
    SQL = SQL & " AND FR.FCL_TCO_CODIGO=TC.TCO_CODIGO"
    SQL = SQL & " AND FR.FCL_TCO_CODIGO=F.TCO_CODIGO"
    SQL = SQL & " AND FR.FCL_SUCURSAL=F.FCL_SUCURSAL"
    SQL = SQL & " AND FR.FCL_NUMERO=F.FCL_NUMERO"

    'BUSCAR NOTA_DEBITO_PROVEEDOR
    SQL = SQL & " UNION ALL"
    SQL = SQL & " SELECT C.CLI_RAZSOC, C.CLI_DOMICI, L.LOC_DESCRI, PR.PRO_DESCRI, C.CLI_CUIT, C.CLI_INGBRU"
    SQL = SQL & ",CI.IVA_DESCRI, TC.TCO_ABREVIA, FR.FCL_SUCURSAL, FR.FCL_NUMERO, FR.FCL_FECHA ,N.NDC_TOTAL, FR.REC_IMPORTE, R.REC_TOTAL"
    SQL = SQL & " FROM CLIENTE C, RECIBO_CLIENTE R ,CONDICION_IVA CI ,LOCALIDAD L"
    SQL = SQL & " , PROVINCIA PR, TIPO_COMPROBANTE TC, FACTURAS_RECIBO_CLIENTE FR,"
    SQL = SQL & " NOTA_DEBITO_CLIENTE N"
    SQL = SQL & " WHERE R.REC_NUMERO=" & XN(txtNroRecibo.Text)
    SQL = SQL & " AND R.REC_SUCURSAL=" & XN(txtNroSucursal.Text)
    SQL = SQL & " AND R.TCO_CODIGO=" & XN(cboRecibo.ItemData(cboRecibo.ListIndex))
    SQL = SQL & " AND R.REC_NUMERO=FR.REC_NUMERO"
    SQL = SQL & " AND R.REC_SUCURSAL=FR.REC_SUCURSAL"
    SQL = SQL & " AND R.TCO_CODIGO=FR.TCO_CODIGO"
    SQL = SQL & " AND R.CLI_CODIGO=C.CLI_CODIGO"
    SQL = SQL & " AND C.LOC_CODIGO=L.LOC_CODIGO"
    SQL = SQL & " AND C.PAI_CODIGO=L.PAI_CODIGO"
    SQL = SQL & " AND C.PRO_CODIGO=L.PRO_CODIGO"
    SQL = SQL & " AND L.PRO_CODIGO=PR.PRO_CODIGO"
    SQL = SQL & " AND L.PAI_CODIGO=PR.PAI_CODIGO"
    SQL = SQL & " AND C.IVA_CODIGO=CI.IVA_CODIGO"
    SQL = SQL & " AND FR.FCL_TCO_CODIGO=TC.TCO_CODIGO"
    SQL = SQL & " AND FR.FCL_TCO_CODIGO=N.TCO_CODIGO"
    SQL = SQL & " AND FR.FCL_SUCURSAL=N.NDC_SUCURSAL"
    SQL = SQL & " AND FR.FCL_NUMERO=N.NDC_NUMERO"
    Rec1.Open SQL, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        Do While Rec1.EOF = False
            
            SQL = "INSERT INTO TMP_RECIBO_CLIENTE ("
            SQL = SQL & "REC_NUMERO,REC_FECHA,CLI_RAZSOC,CLI_DOMICI,CLI_CUIT,CLI_INGBRU,"
            SQL = SQL & "LOC_DESCRI,PRO_DESCRI,IVA_DESCRI,TCO_ABREVIA,COM_FECHA,COM_NUMERO,COM_IMPORTE,"
            SQL = SQL & "REC_TOTAL,FAC_ABREVIA,FAC_NUMERO,FAC_FECHA,FAC_IMPORTE,FAC_TOTAL,REC_ITEM) VALUES ("
            SQL = SQL & XS(Format(txtNroSucursal.Text, "0000") & "-" & Format(txtNroRecibo.Text, "00000000")) & ","
            SQL = SQL & XDQ(FechaRecibo.Text) & ","
            SQL = SQL & XS(Rec1!CLI_RAZSOC) & ","
            SQL = SQL & XS(Rec1!CLI_DOMICI) & ","
            SQL = SQL & XS(Format(Rec1!CLI_CUIT, "##-########-#")) & ","
            SQL = SQL & XS(Format(Rec1!CLI_INGBRU, "###-#####-##")) & ","
            SQL = SQL & XS(Rec1!LOC_DESCRI) & ","
            SQL = SQL & XS(Rec1!PRO_DESCRI) & ","
            SQL = SQL & XS(Rec1!IVA_DESCRI) & ","
            SQL = SQL & "NULL,"
            SQL = SQL & "NULL,"
            SQL = SQL & "NULL,"
            SQL = SQL & "NULL,"
            SQL = SQL & XN(Rec1!REC_TOTAL) & ","
            SQL = SQL & XS(Rec1!TCO_ABREVIA) & ","
            SQL = SQL & XS(Format(Rec1!FCL_SUCURSAL, "0000") & "-" & Format(Rec1!FCL_NUMERO, "00000000")) & ","
            SQL = SQL & XS(Rec1!FCL_FECHA) & ","
            SQL = SQL & XN(Rec1!REC_IMPORTE) & ","
            SQL = SQL & XN(Rec1!FCL_TOTAL) & ","
            SQL = SQL & i & ")"
            DBConn.Execute SQL
            
            i = i + 1
            Rec1.MoveNext
        Loop
    End If
    Rec1.Close
End Sub

Private Sub ReciboComprobante()
    Set Rec1 = New ADODB.Recordset
    SQL = "SELECT C.CLI_RAZSOC, C.CLI_DOMICI, L.LOC_DESCRI, PR.PRO_DESCRI, C.CLI_CUIT, C.CLI_INGBRU"
    SQL = SQL & ",CI.IVA_DESCRI, TC.TCO_ABREVIA, DR.DRE_COMFECHA, DR.DRE_COMSUCURSAL ,DR.DRE_COMNUMERO, DR.DRE_COMIMP, R.REC_TOTAL"
    SQL = SQL & " FROM CLIENTE C, DETALLE_RECIBO_CLIENTE DR, RECIBO_CLIENTE R ,CONDICION_IVA CI"
    SQL = SQL & " ,LOCALIDAD L, PROVINCIA PR, TIPO_COMPROBANTE TC"
    SQL = SQL & " WHERE R.REC_NUMERO=" & XN(txtNroRecibo.Text)
    SQL = SQL & " AND R.REC_SUCURSAL=" & XN(txtNroSucursal.Text)
    SQL = SQL & " AND R.TCO_CODIGO=" & XN(cboRecibo.ItemData(cboRecibo.ListIndex))
    SQL = SQL & " AND R.REC_NUMERO=DR.REC_NUMERO"
    SQL = SQL & " AND R.REC_SUCURSAL=DR.REC_SUCURSAL"
    SQL = SQL & " AND R.TCO_CODIGO=DR.TCO_CODIGO"
    SQL = SQL & " AND R.CLI_CODIGO=C.CLI_CODIGO"
    SQL = SQL & " AND C.LOC_CODIGO=L.LOC_CODIGO"
    SQL = SQL & " AND C.PAI_CODIGO=L.PAI_CODIGO"
    SQL = SQL & " AND C.PRO_CODIGO=L.PRO_CODIGO"
    SQL = SQL & " AND L.PRO_CODIGO=PR.PRO_CODIGO"
    SQL = SQL & " AND L.PAI_CODIGO=PR.PAI_CODIGO"
    SQL = SQL & " AND C.IVA_CODIGO=CI.IVA_CODIGO"
    SQL = SQL & " AND DR.DRE_TCO_CODIGO=TC.TCO_CODIGO"
    Rec1.Open SQL, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        Do While Rec1.EOF = False
                        
            SQL = "INSERT INTO TMP_RECIBO_CLIENTE ("
            SQL = SQL & "REC_NUMERO,REC_FECHA,CLI_RAZSOC,CLI_DOMICI,CLI_CUIT,CLI_INGBRU,"
            SQL = SQL & "LOC_DESCRI,PRO_DESCRI,IVA_DESCRI,TCO_ABREVIA,COM_FECHA,COM_NUMERO,COM_IMPORTE,"
            SQL = SQL & "REC_TOTAL,REC_ITEM) VALUES ("
            SQL = SQL & XS(Format(txtNroSucursal.Text, "0000") & "-" & Format(txtNroRecibo.Text, "00000000")) & ","
            SQL = SQL & XDQ(FechaRecibo.Text) & ","
            SQL = SQL & XS(Rec1!CLI_RAZSOC) & ","
            SQL = SQL & XS(Rec1!CLI_DOMICI) & ","
            SQL = SQL & XS(Format(Rec1!CLI_CUIT, "##-########-#")) & ","
            SQL = SQL & XS(Format(Rec1!CLI_INGBRU, "###-#####-##")) & ","
            SQL = SQL & XS(Rec1!LOC_DESCRI) & ","
            SQL = SQL & XS(Rec1!PRO_DESCRI) & ","
            SQL = SQL & XS(Rec1!IVA_DESCRI) & ","
            SQL = SQL & XS(Rec1!TCO_ABREVIA) & ","
            SQL = SQL & XDQ(Rec1!DRE_COMFECHA) & ","
            SQL = SQL & XS(Rec1!DRE_COMSUCURSAL & "-" & Format(Rec1!DRE_COMNUMERO, "00000000")) & ","
            SQL = SQL & XN(Rec1!DRE_COMIMP) & ","
            SQL = SQL & XN(Rec1!REC_TOTAL) & ","
            SQL = SQL & i & ")"
            DBConn.Execute SQL
            
            i = i + 1
            Rec1.MoveNext
        Loop
    End If
    Rec1.Close
End Sub

Private Sub ReciboCheques()
    Set Rec1 = New ADODB.Recordset
    'PARA CHEQUES DE TERCEROS
    SQL = "SELECT C.CLI_RAZSOC, C.CLI_DOMICI, L.LOC_DESCRI, PR.PRO_DESCRI, C.CLI_CUIT, C.CLI_INGBRU"
    SQL = SQL & ",CI.IVA_DESCRI, B.BAN_NOMCOR, CH.CHE_FECVTO ,DR.CHE_NUMERO, CH.CHE_IMPORT, R.REC_TOTAL"
    SQL = SQL & " FROM CLIENTE C, DETALLE_RECIBO_CLIENTE DR, RECIBO_CLIENTE R ,CONDICION_IVA CI"
    SQL = SQL & " ,LOCALIDAD L, PROVINCIA PR, CHEQUE CH, BANCO B"
    SQL = SQL & " WHERE R.REC_NUMERO=" & XN(txtNroRecibo.Text)
    SQL = SQL & " AND R.REC_SUCURSAL=" & XN(txtNroSucursal.Text)
    SQL = SQL & " AND R.TCO_CODIGO=" & XN(cboRecibo.ItemData(cboRecibo.ListIndex))
    SQL = SQL & " AND R.REC_NUMERO=DR.REC_NUMERO"
    SQL = SQL & " AND R.REC_SUCURSAL=DR.REC_SUCURSAL"
    SQL = SQL & " AND R.TCO_CODIGO=DR.TCO_CODIGO"
    SQL = SQL & " AND R.CLI_CODIGO=C.CLI_CODIGO"
    SQL = SQL & " AND C.LOC_CODIGO=L.LOC_CODIGO"
    SQL = SQL & " AND C.PAI_CODIGO=L.PAI_CODIGO"
    SQL = SQL & " AND C.PRO_CODIGO=L.PRO_CODIGO"
    SQL = SQL & " AND L.PRO_CODIGO=PR.PRO_CODIGO"
    SQL = SQL & " AND L.PAI_CODIGO=PR.PAI_CODIGO"
    SQL = SQL & " AND C.IVA_CODIGO=CI.IVA_CODIGO"
    SQL = SQL & " AND DR.BAN_CODINT=CH.BAN_CODINT"
    SQL = SQL & " AND DR.CHE_NUMERO=CH.CHE_NUMERO"
    SQL = SQL & " AND CH.BAN_CODINT=B.BAN_CODINT"
    Rec1.Open SQL, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        Do While Rec1.EOF = False
            
            SQL = "INSERT INTO TMP_RECIBO_CLIENTE ("
            SQL = SQL & "REC_NUMERO,REC_FECHA,CLI_RAZSOC,CLI_DOMICI,CLI_CUIT,CLI_INGBRU,"
            SQL = SQL & "LOC_DESCRI,PRO_DESCRI,IVA_DESCRI,TCO_ABREVIA,COM_FECHA,COM_NUMERO,COM_IMPORTE,"
            SQL = SQL & "REC_TOTAL,REC_ITEM) VALUES ("
            SQL = SQL & XS(Format(txtNroSucursal.Text, "0000") & "-" & Format(txtNroRecibo.Text, "00000000")) & ","
            SQL = SQL & XDQ(FechaRecibo.Text) & ","
            SQL = SQL & XS(Rec1!CLI_RAZSOC) & ","
            SQL = SQL & XS(Rec1!CLI_DOMICI) & ","
            SQL = SQL & XS(Format(Rec1!CLI_CUIT, "##-########-#")) & ","
            SQL = SQL & XS(Format(Rec1!CLI_INGBRU, "###-#####-##")) & ","
            SQL = SQL & XS(Rec1!LOC_DESCRI) & ","
            SQL = SQL & XS(Rec1!PRO_DESCRI) & ","
            SQL = SQL & XS(Rec1!IVA_DESCRI) & ","
            SQL = SQL & XS(Rec1!BAN_NOMCOR) & ","
            SQL = SQL & XDQ(Rec1!CHE_FECVTO) & ","
            SQL = SQL & XS(Rec1!CHE_NUMERO) & ","
            SQL = SQL & XN(Rec1!CHE_IMPORT) & ","
            SQL = SQL & XN(Rec1!REC_TOTAL) & ","
            SQL = SQL & i & ")"
            DBConn.Execute SQL
            
            i = i + 1
            Rec1.MoveNext
        Loop
    End If
    Rec1.Close
    'PARA CHEQUES PROPIOS
    SQL = "SELECT C.CLI_RAZSOC, C.CLI_DOMICI, L.LOC_DESCRI, PR.PRO_DESCRI, C.CLI_CUIT, C.CLI_INGBRU"
    SQL = SQL & ",CI.IVA_DESCRI, B.BAN_NOMCOR, CH.CHEP_FECVTO ,DR.CHE_NUMERO, CH.CHEP_IMPORT, R.REC_TOTAL"
    SQL = SQL & " FROM CLIENTE C,DETALLE_RECIBO_CLIENTE DR, RECIBO_CLIENTE R ,CONDICION_IVA CI"
    SQL = SQL & " ,LOCALIDAD L, PROVINCIA PR, CHEQUE_PROPIO CH, BANCO B"
    SQL = SQL & " WHERE R.REC_NUMERO=" & XN(txtNroRecibo.Text)
    SQL = SQL & " AND R.REC_SUCURSAL=" & XN(txtNroSucursal.Text)
    SQL = SQL & " AND R.TCO_CODIGO=" & XN(cboRecibo.ItemData(cboRecibo.ListIndex))
    SQL = SQL & " AND R.REC_NUMERO=DR.REC_NUMERO"
    SQL = SQL & " AND R.REC_SUCURSAL=DR.REC_SUCURSAL"
    SQL = SQL & " AND R.TCO_CODIGO=DR.TCO_CODIGO"
    SQL = SQL & " AND R.CLI_CODIGO=C.CLI_CODIGO"
    SQL = SQL & " AND C.LOC_CODIGO=L.LOC_CODIGO"
    SQL = SQL & " AND C.PAI_CODIGO=L.PAI_CODIGO"
    SQL = SQL & " AND C.PRO_CODIGO=L.PRO_CODIGO"
    SQL = SQL & " AND L.PRO_CODIGO=PR.PRO_CODIGO"
    SQL = SQL & " AND L.PAI_CODIGO=PR.PAI_CODIGO"
    SQL = SQL & " AND C.IVA_CODIGO=CI.IVA_CODIGO"
    SQL = SQL & " AND DR.BAN_CODINT=CH.BAN_CODINT"
    SQL = SQL & " AND DR.CHE_NUMERO=CH.CHEP_NUMERO"
    SQL = SQL & " AND CH.BAN_CODINT=B.BAN_CODINT"
    Rec1.Open SQL, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        Do While Rec1.EOF = False
            
            SQL = "INSERT INTO TMP_RECIBO_CLIENTE ("
            SQL = SQL & "REC_NUMERO,REC_FECHA,CLI_RAZSOC,CLI_DOMICI,CLI_CUIT,CLI_INGBRU,"
            SQL = SQL & "LOC_DESCRI,PRO_DESCRI,IVA_DESCRI,TCO_ABREVIA,COM_FECHA,COM_NUMERO,COM_IMPORTE,"
            SQL = SQL & "REC_TOTAL,REC_ITEM) VALUES ("
            SQL = SQL & XS(Format(txtNroSucursal.Text, "0000") & "-" & Format(txtNroRecibo.Text, "00000000")) & ","
            SQL = SQL & XDQ(FechaRecibo.Text) & ","
            SQL = SQL & XS(Rec1!CLI_RAZSOC) & ","
            SQL = SQL & XS(Rec1!CLI_DOMICI) & ","
            SQL = SQL & XS(Format(Rec1!CLI_CUIT, "##-########-#")) & ","
            SQL = SQL & XS(Format(Rec1!CLI_INGBRU, "###-#####-##")) & ","
            SQL = SQL & XS(Rec1!LOC_DESCRI) & ","
            SQL = SQL & XS(Rec1!PRO_DESCRI) & ","
            SQL = SQL & XS(Rec1!IVA_DESCRI) & ","
            SQL = SQL & XS(Rec1!BAN_NOMCOR) & ","
            SQL = SQL & XDQ(Rec1!CHEP_FECVTO) & ","
            SQL = SQL & XS(Rec1!CHE_NUMERO) & ","
            SQL = SQL & XN(Rec1!CHEP_IMPORT) & ","
            SQL = SQL & XN(Rec1!OPG_TOTAL) & ","
            SQL = SQL & i & ")"
            DBConn.Execute SQL
            
            i = i + 1
            Rec1.MoveNext
        Loop
    End If
    Rec1.Close
End Sub

Private Sub ReciboMoneda()
    Set Rec1 = New ADODB.Recordset
    SQL = "SELECT C.CLI_RAZSOC, C.CLI_DOMICI, L.LOC_DESCRI, PR.PRO_DESCRI, C.CLI_CUIT, C.CLI_INGBRU"
    SQL = SQL & ", M.MON_DESCRI, DR.DRE_MONIMP, R.REC_TOTAL, CI.IVA_DESCRI"
    SQL = SQL & " FROM CLIENTE C, DETALLE_RECIBO_CLIENTE DR, RECIBO_CLIENTE R"
    SQL = SQL & " ,LOCALIDAD L, PROVINCIA PR, MONEDA M, CONDICION_IVA CI"
    SQL = SQL & " WHERE R.REC_NUMERO=" & XN(txtNroRecibo.Text)
    SQL = SQL & " AND R.REC_SUCURSAL=" & XN(txtNroSucursal.Text)
    SQL = SQL & " AND R.TCO_CODIGO=" & XN(cboRecibo.ItemData(cboRecibo.ListIndex))
    SQL = SQL & " AND R.REC_NUMERO=DR.REC_NUMERO"
    SQL = SQL & " AND R.REC_SUCURSAL=DR.REC_SUCURSAL"
    SQL = SQL & " AND R.TCO_CODIGO=DR.TCO_CODIGO"
    SQL = SQL & " AND R.CLI_CODIGO=C.CLI_CODIGO"
    SQL = SQL & " AND DR.MON_CODIGO=M.MON_CODIGO"
    SQL = SQL & " AND C.LOC_CODIGO=L.LOC_CODIGO"
    SQL = SQL & " AND C.PAI_CODIGO=L.PAI_CODIGO"
    SQL = SQL & " AND C.PRO_CODIGO=L.PRO_CODIGO"
    SQL = SQL & " AND L.PRO_CODIGO=PR.PRO_CODIGO"
    SQL = SQL & " AND L.PAI_CODIGO=PR.PAI_CODIGO"
    SQL = SQL & " AND C.IVA_CODIGO=CI.IVA_CODIGO"
    
    Rec1.Open SQL, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        Do While Rec1.EOF = False
            
            SQL = "INSERT INTO TMP_RECIBO_CLIENTE ("
            SQL = SQL & "REC_NUMERO,REC_FECHA,CLI_RAZSOC,CLI_DOMICI,CLI_CUIT,CLI_INGBRU,"
            SQL = SQL & "LOC_DESCRI,PRO_DESCRI,IVA_DESCRI,TCO_ABREVIA,COM_IMPORTE,"
            SQL = SQL & "REC_TOTAL,REC_ITEM) VALUES ("
            SQL = SQL & XS(Format(txtNroSucursal.Text, "0000") & "-" & Format(txtNroRecibo.Text, "00000000")) & ","
            SQL = SQL & XDQ(FechaRecibo.Text) & ","
            SQL = SQL & XS(Rec1!CLI_RAZSOC) & ","
            SQL = SQL & XS(Rec1!CLI_DOMICI) & ","
            SQL = SQL & XS(Format(Rec1!CLI_CUIT, "##-########-#")) & ","
            SQL = SQL & XS(Format(Rec1!CLI_INGBRU, "###-#####-##")) & ","
            SQL = SQL & XS(Rec1!LOC_DESCRI) & ","
            SQL = SQL & XS(Rec1!PRO_DESCRI) & ","
            SQL = SQL & XS(Rec1!IVA_DESCRI) & ","
            SQL = SQL & XS(Rec1!MON_DESCRI) & ","
            SQL = SQL & XN(Rec1!DRE_MONIMP) & ","
            SQL = SQL & XN(Rec1!REC_TOTAL) & ","
            SQL = SQL & i & ")"
            DBConn.Execute SQL
            
            i = i + 1
            Rec1.MoveNext
        Loop
    End If
    Rec1.Close
End Sub

Private Sub cmaAceptarACta_Click()
    txtSaldoACta.Text = ""
    txtImporteACta.Text = ""
    tabValores.Tab = 0
End Sub

Private Sub cmdAceptarMoneda_Click()
    If GrillaEfectivo.Rows > 1 Then
        'CARGO EN GRILLA VALORES
        For i = 1 To GrillaEfectivo.Rows - 1
            grillaValores.AddItem "EFT" & Chr(9) & _
                                  GrillaEfectivo.TextMatrix(i, 1) & Chr(9) & _
                                  "" & Chr(9) & _
                                  GrillaEfectivo.TextMatrix(i, 0) & Chr(9) & _
                                  "" & Chr(9) & _
                                  GrillaEfectivo.TextMatrix(i, 2)
        Next
        txtTotalValores.Text = Valido_Importe(CStr(SumaGrilla(grillaValores, 1)))
        grillaValores.HighLight = flexHighlightAlways
        GrillaEfectivo.Rows = 1
        txtTotalEfectivo.Text = ""
        tabValores.Tab = 0
    End If
    cmdGrabar.Enabled = True
    cmdGrabar.SetFocus
End Sub

Private Sub cmdAgregarACta_Click()
    If GrillaAFavor.Rows > 1 Then
        If grillaValores.Rows > 1 Then
            For i = 1 To grillaValores.Rows - 1
                If GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 5) = grillaValores.TextMatrix(i, 5) _
                    And (GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 1)) = (grillaValores.TextMatrix(i, 4)) _
                    And CDate(GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 2)) = CDate(grillaValores.TextMatrix(i, 2)) _
                    And (GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 6)) = (grillaValores.TextMatrix(i, 6)) Then
                   MsgBox "El Valor ya fue ingresado", vbInformation, TIT_MSGBOX
                   txtSaldoACta.Text = ""
                   txtImporteACta.Text = ""
                   GrillaAFavor.SetFocus
                   Exit Sub
                End If
            Next
        End If
                
        'CARGO EN GRILLA VALORES
        grillaValores.AddItem "A-CTA" & Chr(9) & Valido_Importe(txtImporteACta) & Chr(9) & GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 2) _
                                & Chr(9) & GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 0) & Chr(9) & GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 1) & Chr(9) & _
                                GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 5) & Chr(9) & GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 6)

        'ARREGLO EL SALDO DEL DINERO A CTA
        GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 4) = Valido_Importe(CStr(CDbl(GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 4)) - CDbl(Chk0(txtImporteACta.Text))))
        
        txtTotalValores.Text = Valido_Importe(CStr(SumaGrilla(grillaValores, 1)))
        grillaValores.HighLight = flexHighlightAlways

        txtSaldoACta.Text = ""
        txtImporteACta.Text = ""
        GrillaAFavor.SetFocus
    End If
End Sub

Private Function ValidoIngCheques() As Boolean
'    For I = 1 To GrillaCheques.Rows - 1
'        If TxtCodInt.Text = GrillaCheques.TextMatrix(I, 7) And _
'           TxtCheNumero.Text = GrillaCheques.TextMatrix(I, 4) Then
'
'           ValidoIngCheques = False
'           Exit Function
'        End If
'    Next
'    ValidoIngCheques = True
End Function

Private Sub LimpiarCheques()
    'TxtBANCO.Text = ""
    'TxtLOCALIDAD.Text = ""
    'TxtSUCURSAL.Text = ""
    'txtCodigo.Text = ""
    'TxtCheNumero.Text = ""
    'TxtCheFecVto.Text = ""
    'TxtCheImport.Text = ""
    'TxtCodInt.Text = ""
    'TxtBanDescri.Text = ""
    'frameBanco.Enabled = False
    'cmdAgregarCheque.Enabled = False
End Sub

Private Function BuscarTipoDocAbre(Codigo As String) As String
    Set Rec1 = New ADODB.Recordset
    SQL = "SELECT TCO_ABREVIA"
    SQL = SQL & " FROM TIPO_COMPROBANTE"
    SQL = SQL & " WHERE TCO_CODIGO = " & XN(Codigo)
    Rec1.Open SQL, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        BuscarTipoDocAbre = Rec1!TCO_ABREVIA
    Else
        BuscarTipoDocAbre = ""
    End If
    Rec1.Close
End Function

Private Sub cmdAgregarEfectivo_Click()
    'VALIDO QUE EL CHEQUE NO SE HAYA CARGADO
    'If GrillaEfectivo.Rows > 1 Then
    '    If ValidoIngMoneda = False Then
    '        MsgBox "La Moneda ya fue ingresada", vbCritical, TIT_MSGBOX
    '        txtEftImporte.Text = ""
    '        cboMoneda.SetFocus
    '        Exit Sub
    '    End If
    'End If
    
    Dim TotalDineroaCta As String
    TotalDineroaCta = "0"
    
    If grillaValores.Rows > 1 Then
       For i = 1 To grillaValores.Rows - 1
          TotalDineroaCta = CDbl(TotalDineroaCta) + CDbl(grillaValores.TextMatrix(i, 1))
       Next i
       If txtEftImporte.Text = "" Then
          txtEftImporte.Text = Format(txtImporteApagar.Text - CDbl(TotalDineroaCta), "0.00")
       End If
    Else
       If txtEftImporte.Text = "" Then
          txtEftImporte.Text = txtImporteApagar.Text
       End If
    End If
    
    'CARGO GRILLA
    GrillaEfectivo.AddItem cboMoneda.Text & Chr(9) & txtEftImporte.Text & Chr(9) & cboMoneda.ItemData(cboMoneda.ListIndex)
                                                   
    GrillaEfectivo.HighLight = flexHighlightAlways
    txtTotalEfectivo.Text = Valido_Importe(CStr(SumaGrilla(GrillaEfectivo, 1)))
    'txtEftImporte.Text = ""
    'cboMoneda.SetFocus
    cmdAceptarMoneda.SetFocus
End Sub

Private Function ValidoIngMoneda() As Boolean
    For i = 1 To GrillaEfectivo.Rows - 1
        If cboMoneda.ItemData(cboMoneda.ListIndex) = GrillaEfectivo.TextMatrix(i, 2) Then
           
           ValidoIngMoneda = False
           Exit Function
        End If
    Next
    ValidoIngMoneda = True
End Function

Private Function ValidoIngFactura(Combo As ComboBox, Grilla As MSFlexGrid, NROFAC As String, NroSuc As String) As Boolean
    For i = 1 To Grilla.Rows - 1
        If Combo.ItemData(Combo.ListIndex) = Grilla.TextMatrix(i, 4) And _
           NROFAC = Right(Grilla.TextMatrix(i, 1), 8) And _
           NroSuc = Left(Grilla.TextMatrix(i, 1), 4) Then
           
           ValidoIngFactura = False
           Exit Function
        End If
    Next
    ValidoIngFactura = True
End Function

Private Sub CmdBuscAprox_Click()
    GrdModulos.Rows = 1
    GrdModulos.HighLight = flexHighlightNever
    lblEstado.Caption = "Buscando..."
    Screen.MousePointer = vbHourglass
    
    Set Rec1 = New ADODB.Recordset
    SQL = "SELECT RC.REC_NUMERO, RC.REC_SUCURSAL,"
    SQL = SQL & " RC.REC_FECHA, RC.TCO_CODIGO, TC.TCO_ABREVIA, C.CLI_RAZSOC, RC.REC_TOTAL "
    SQL = SQL & " FROM RECIBO_CLIENTE RC, CLIENTE C,  TIPO_COMPROBANTE TC"
    SQL = SQL & " WHERE RC.TCO_CODIGO=TC.TCO_CODIGO"
    SQL = SQL & "   AND RC.CLI_CODIGO=C.CLI_CODIGO"
    If txtCliente.Text <> "" Then SQL = SQL & " AND RC.CLI_CODIGO=" & XN(txtCliente.Text)
    If FechaDesde.Text <> "" Then SQL = SQL & " AND RC.REC_FECHA>=" & XDQ(FechaDesde.Text)
    If FechaHasta.Text <> "" Then SQL = SQL & " AND RC.REC_FECHA<=" & XDQ(FechaHasta.Text)
    If cboRecibo1.List(cboRecibo1.ListIndex) <> "(Todos)" Then SQL = SQL & " AND RC.TCO_CODIGO=" & XN(cboRecibo1.ItemData(cboRecibo1.ListIndex))
    SQL = SQL & " ORDER BY RC.REC_SUCURSAL, RC.REC_NUMERO"
    
    Rec1.Open SQL, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        Do While Rec1.EOF = False
            GrdModulos.AddItem Rec1!TCO_ABREVIA & Chr(9) & Format(Rec1!REC_SUCURSAL, "0000") & "-" & Format(Rec1!REC_NUMERO, "00000000") _
                               & Chr(9) & Rec1!REC_FECHA & Chr(9) & Rec1!CLI_RAZSOC _
                               & Chr(9) & Rec1!TCO_CODIGO & Chr(9) & Valido_Importe(Rec1!REC_TOTAL)
            Rec1.MoveNext
        Loop
        GrdModulos.HighLight = flexHighlightAlways
        GrdModulos.SetFocus
        GrdModulos.Col = 0
    Else
        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        MsgBox "No se encontraron datos... ", vbExclamation, TIT_MSGBOX
        txtCliente.SetFocus
    End If
    Rec1.Close
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
End Sub

Private Sub cmdBuscarCliente_Click()
    frmBuscar.TipoBusqueda = 1
    frmBuscar.TxtDescriB = ""
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Text <> "" Then
        frmBuscar.grdBuscar.Col = 0
        txtCodCliente.Text = frmBuscar.grdBuscar.Text
        txtCodCliente_LostFocus
    Else
        txtCodCliente.SetFocus
    End If
End Sub

Private Sub cmdCancelarMoneda_Click()
    GrillaEfectivo.Rows = 1
    txtTotalEfectivo.Text = ""
    tabValores.Tab = 0
End Sub

Private Sub cmdGrabar_Click()
    If ValidarRecibo = False Then Exit Sub
    If MsgBox("¿Confirma Recibo?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    
    On Error GoTo HayError
    DBConn.BeginTrans
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando..."
    
    SQL = "SELECT EST_CODIGO"
    SQL = SQL & " FROM RECIBO_CLIENTE"
    SQL = SQL & " WHERE TCO_CODIGO = " & XN(cboRecibo.ItemData(cboRecibo.ListIndex))
    SQL = SQL & "   AND REC_NUMERO = " & XN(txtNroRecibo.Text)
    SQL = SQL & "   AND REC_SUCURSAL = " & XN(txtNroSucursal.Text)
    rec.Open SQL, DBConn, adOpenStatic, adLockOptimistic
    
    If rec.EOF = True Then
        
        'CABEZA DEL RECIBO
        SQL = "INSERT INTO RECIBO_CLIENTE ("
        SQL = SQL & " TCO_CODIGO, REC_NUMERO, REC_SUCURSAL, REC_FECHA,"
        SQL = SQL & " EST_CODIGO, CLI_CODIGO,"
        SQL = SQL & " REC_NUMEROTXT, REC_TOTAL)"
        SQL = SQL & " VALUES ("
        SQL = SQL & XN(cboRecibo.ItemData(cboRecibo.ListIndex)) & ", "
        SQL = SQL & XN(txtNroRecibo.Text) & ","
        SQL = SQL & XN(txtNroSucursal.Text) & ","
        SQL = SQL & XDQ(FechaRecibo.Text) & ","
        SQL = SQL & "3,"                          'ESTADO DEFINITIVO
        SQL = SQL & XN(txtCodCliente.Text) & ","
        SQL = SQL & XS(Format(txtNroRecibo.Text, "00000000")) & ","
        SQL = SQL & XN(txtImporteApagar.Text) & ")"
        DBConn.Execute SQL
        
        'DETALLE DEL RECIBO
        For i = 1 To grillaValores.Rows - 1
            SQL = "INSERT INTO DETALLE_RECIBO_CLIENTE"
            SQL = SQL & " (TCO_CODIGO, REC_NUMERO, REC_SUCURSAL, REC_FECHA,"
            SQL = SQL & " DRE_NROITEM, MON_CODIGO,"
            SQL = SQL & " DRE_MONIMP, DRE_TCO_CODIGO, DRE_COMFECHA, DRE_COMNUMERO,"
            SQL = SQL & " DRE_COMSUCURSAL, DRE_COMIMP)"
            SQL = SQL & " VALUES ("
            SQL = SQL & XN(cboRecibo.ItemData(cboRecibo.ListIndex)) & ","
            SQL = SQL & XN(txtNroRecibo.Text) & ","
            SQL = SQL & XN(txtNroSucursal.Text) & ","
            SQL = SQL & XDQ(FechaRecibo.Text) & ","
            SQL = SQL & XN(CStr(i)) & ","
            
            If grillaValores.TextMatrix(i, 0) = "EFT" Then
                SQL = SQL & XN(grillaValores.TextMatrix(i, 5)) & "," 'MONEDA
                SQL = SQL & XN(grillaValores.TextMatrix(i, 1)) & "," 'IMPORTE
            Else
                SQL = SQL & "NULL,NULL,"
            End If
            
            If grillaValores.TextMatrix(i, 0) = "COMP" Or grillaValores.TextMatrix(i, 0) = "A-CTA" Then
                SQL = SQL & XN(grillaValores.TextMatrix(i, 5)) & ","
                SQL = SQL & XDQ(grillaValores.TextMatrix(i, 2)) & ","
                SQL = SQL & XN(Right(grillaValores.TextMatrix(i, 4), 8)) & "," 'NUMERO COMPROBANTE
                SQL = SQL & XN(Left(grillaValores.TextMatrix(i, 4), 4)) & ","  'NUMERO SUCURSAL
                SQL = SQL & XN(grillaValores.TextMatrix(i, 1)) & ")"
            Else
                SQL = SQL & "NULL,NULL,NULL,NULL,NULL)"
            End If
            DBConn.Execute SQL
        Next
        
        'FACTURAS Y NOTA DE DEBITO CANCELADAS EN EL RECIBO
        For i = 1 To GrillaAplicar.Rows - 1
           If CDbl(txtImporteApagar.Text) > 0 Then
                SQL = "INSERT INTO FACTURAS_RECIBO_CLIENTE"
                SQL = SQL & " (TCO_CODIGO, REC_NUMERO, REC_SUCURSAL, REC_FECHA,"
                SQL = SQL & " FCL_TCO_CODIGO, FCL_NUMERO, FCL_SUCURSAL,"
                SQL = SQL & " FCL_FECHA,REC_IMPORTE,REC_ABONA,REC_SALDO)"
                SQL = SQL & " VALUES ("
                SQL = SQL & XN(cboRecibo.ItemData(cboRecibo.ListIndex)) & ","
                SQL = SQL & XN(txtNroRecibo.Text) & ","
                SQL = SQL & XN(txtNroSucursal.Text) & ","
                SQL = SQL & XDQ(FechaRecibo) & ","
                SQL = SQL & XN(GrillaAplicar.TextMatrix(i, 6)) & ","           'TIPO FACTURA O NOTA DEBITO
                SQL = SQL & XN(Right(GrillaAplicar.TextMatrix(i, 1), 8)) & "," 'NUMERO FACTURA O NOTA DEBITO
                SQL = SQL & XN(Left(GrillaAplicar.TextMatrix(i, 1), 4)) & ","  'NUMERO SUCURSAL
                SQL = SQL & XDQ(GrillaAplicar.TextMatrix(i, 2)) & ","          'FECHA FACTURA O NOTA DEBITO
                
                'Comparo para ver si me queda saldo
                If CDbl(txtImporteApagar.Text) > Valido_Importe(GrillaAplicar.TextMatrix(i, 5)) Then
                   'Importe TOTAL de la Factura
                   txtImporteApagar.Text = CDbl(txtImporteApagar.Text) - _
                                           Valido_Importe(GrillaAplicar.TextMatrix(i, 5))
                   GrillaAplicar.TextMatrix(i, 4) = GrillaAplicar.TextMatrix(i, 3)
                   GrillaAplicar.TextMatrix(i, 5) = "0,00"
                   
                Else
                   'Importe del SALDO
                   GrillaAplicar.TextMatrix(i, 4) = txtImporteApagar.Text
                   GrillaAplicar.TextMatrix(i, 5) = Format(CDbl(GrillaAplicar.TextMatrix(i, 5)) - CDbl(GrillaAplicar.TextMatrix(i, 4)), "0.00")
                   txtImporteApagar.Text = "0,00"
                End If
                
                SQL = SQL & XN(GrillaAplicar.TextMatrix(i, 3)) & _
                     ", " & XN(GrillaAplicar.TextMatrix(i, 4)) & _
                     ", " & XN(GrillaAplicar.TextMatrix(i, 5)) & ")"
                DBConn.Execute SQL
           End If
        Next
        
        'ACTUALIZO EL SALDO DE LAS FACTURAS ELEGIDAS
        For i = 1 To GrillaAplicar.Rows - 1
           If Trim(GrillaAplicar.TextMatrix(i, 4)) <> "0,00" Then
                SQL = "UPDATE FACTURA_CLIENTE"
                SQL = SQL & " SET FCL_SALDO = " & XN(GrillaAplicar.TextMatrix(i, 5))
                SQL = SQL & " WHERE TCO_CODIGO=" & XN(GrillaAplicar.TextMatrix(i, 6))
                SQL = SQL & "   AND FCL_NUMERO=" & XN(Right(GrillaAplicar.TextMatrix(i, 1), 8))  'NUMERO FACTURA
                SQL = SQL & "   AND FCL_SUCURSAL=" & XN(Left(GrillaAplicar.TextMatrix(i, 1), 4)) 'NUMERO SUCURSAL
                DBConn.Execute SQL
           End If
        Next

        'ACTUALIZO EL DINERO A CUENTA (RECIBO_CLIENTE_SALDO)
        For i = 1 To GrillaAFavor.Rows - 1
            If GrillaAFavor.TextMatrix(i, 5) <> "19" Then '19 ANTICIPO DE COBRO
                SQL = "UPDATE RECIBO_CLIENTE_SALDO"
                SQL = SQL & " SET REC_SALDO = " & XN(GrillaAFavor.TextMatrix(i, 4))
                SQL = SQL & " WHERE TCO_CODIGO = " & XN(GrillaAFavor.TextMatrix(i, 5))
                SQL = SQL & "   AND REC_NUMERO = " & XN(Right(GrillaAFavor.TextMatrix(i, 1), 8)) 'NUMERO RECIBO
                SQL = SQL & "   AND REC_SUCURSAL = " & XN(Left(GrillaAFavor.TextMatrix(i, 1), 4)) 'NUMERO SUCURSAL
                DBConn.Execute SQL
            Else
                SQL = "UPDATE ANTICIPO_COBRO"
                SQL = SQL & " SET ANC_SALDO = " & XN(GrillaAFavor.TextMatrix(i, 4))
                SQL = SQL & " WHERE ANC_NUMERO = " & XN(Right(GrillaAFavor.TextMatrix(i, 1), 8)) 'NUMERO ANTICIPO
                SQL = SQL & "   AND ANC_SUCURSAL = " & XN(Left(GrillaAFavor.TextMatrix(i, 1), 4)) 'NUMERO SUCURSAL
                SQL = SQL & "   AND ANC_FECHA = " & XDQ(GrillaAFavor.TextMatrix(i, 2))
                DBConn.Execute SQL
            End If
        Next

        'VERIFICO SI HAY DINERO A CUENTA
        If CDbl(txtImporteApagar.Text) > 0 Then
            SQL = "INSERT INTO RECIBO_CLIENTE_SALDO"
            SQL = SQL & " (TCO_CODIGO, REC_NUMERO, REC_SUCURSAL, REC_FECHA,"
            SQL = SQL & " REC_TOTSALDO, REC_SALDO)"
            SQL = SQL & " VALUES ("
            SQL = SQL & XN(cboRecibo.ItemData(cboRecibo.ListIndex)) & ","
            SQL = SQL & XN(txtNroRecibo.Text) & ","
            SQL = SQL & XN(txtNroSucursal.Text) & ","
            SQL = SQL & XDQ(FechaRecibo.Text) & ","
            SQL = SQL & XN(CDbl(txtTotalValores.Text)) & ","
            SQL = SQL & XN(CDbl(txtImporteApagar.Text)) & ")"
            DBConn.Execute SQL
        End If
                                                
        'ACTUALIZO LA TABLA PARAMENTROS Y LE SUMO UNO AL RECIBO QUE CORRESPONDA
        Select Case cboRecibo.ItemData(cboRecibo.ListIndex)
            Case 12
                SQL = "UPDATE PARAMETROS SET RECIBO_C=" & XN(txtNroRecibo)
                DBConn.Execute SQL
        End Select
        
        DBConn.CommitTrans
        mBorroTransfe = False
        
        'If MsgBox("Desea Imprimir el Recibo?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        '    'MANDO RECIBO A IMPRESORA
        '    mImprimoRecibo = False
        '    cmdImprimir_Click
        'End If
    Else 'SI EXISTE
        MsgBox "El Recibo ya fue Registrado", vbCritical, TIT_MSGBOX
        DBConn.CommitTrans
        mBorroTransfe = True
    End If
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    rec.Close
    cmdNuevo_Click
    Exit Sub
    
HayError:
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    If rec.State = 1 Then rec.Close
    DBConn.RollbackTrans
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Function ValidarRecibo() As Boolean
    
    If txtNroSucursal.Text = "" Or txtNroRecibo.Text = "" Then
        MsgBox "Debe ingresar el número de Recibo", vbCritical, TIT_MSGBOX
        txtNroSucursal.SetFocus
        ValidarRecibo = False
        Exit Function
    End If
    If FechaRecibo.Text = "" Then
        MsgBox "Debe ingresar la fecha del Recibo", vbCritical, TIT_MSGBOX
        FechaRecibo.SetFocus
        ValidarRecibo = False
        Exit Function
    End If
    If txtCodCliente.Text = "" Then
        MsgBox "Debe ingresar un Cliente", vbCritical, TIT_MSGBOX
        txtCodCliente.SetFocus
        ValidarRecibo = False
        Exit Function
    End If
    If grillaValores.Rows = 1 Then
        MsgBox "Debe ingresar Valores Recibidos", vbCritical, TIT_MSGBOX
        ValidarRecibo = False
        Exit Function
    End If
    If GrillaAplicar.Rows = 1 Then
        MsgBox "No tiene Facturas pendientes", vbCritical, TIT_MSGBOX
        'cmdAgregarFactura.SetFocus
        ValidarRecibo = False
        Exit Function
    End If
    'If CDbl(txtSaldo.Text) > CDbl(txtTotalValores.Text) Then
    '    MsgBox "El Total de Facturas supera al Total de Valores Recibidos", vbCritical, TIT_MSGBOX
    '    ValidarRecibo = False
    '    Exit Function
    'End If
    If CDbl(txtSaldo.Text) < CDbl(txtTotalValores.Text) Then
        If MsgBox("El Total de Valores Recibidos supera al Total de Facturas," & Chr(13) & _
                "deja el importe (" & Format(CStr(CDbl(txtTotalValores.Text) - CDbl(txtSaldo.Text)), "#,##0.00") & _
                ") como dinero a cuenta?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then

            'cmdAgregarFactura.SetFocus
            ValidarRecibo = False
            Exit Function
        End If
    End If
    ValidarRecibo = True
End Function

Private Sub cmdNuevo_Click()
    Estado = 1
    If mBorroTransfe = True Then
       'VERIFICO SI HAY UNA TRASFERENCIA CARGADA
       'SI HAY LA BORRO DE LA TABLA DEBCRE_BANCARIOS
       For i = 1 To grillaValores.Rows - 1
           If grillaValores.TextMatrix(i, 0) = "COMP" Then
               If grillaValores.TextMatrix(i, 5) = 30 Then
                   DBConn.Execute "DELETE FROM DEBCRE_BANCARIOS WHERE DCB_NUMERO = " & XN(Right(Trim(grillaValores.TextMatrix(i, 4)), 8))
               End If
           End If
       Next
    End If
    cmdGrabar.Enabled = True
    FrameRecibo.Enabled = True
    FrameRemito.Enabled = True
'    TxtCheNumero.Text = ""
'    GrillaCheques.Rows = 1
'    GrillaCheques.HighLight = flexHighlightNever
    txtEftImporte.Text = ""
    GrillaEfectivo.Rows = 1
    GrillaEfectivo.HighLight = flexHighlightNever
    GrillaAplicar.Rows = 1
    GrillaAplicar.HighLight = flexHighlightNever
    
    GrillaAFavor.Rows = 1
    GrillaAFavor.HighLight = flexHighlightNever
    LblDineroaCta.Caption = ""
    
    grillaValores.Rows = 1
    grillaValores.HighLight = flexHighlightNever
    
    txtCodCliente.Text = ""
    txtNroRecibo.Text = ""
    txtNroSucursal.Text = ""
    txtSaldo.Text = ""
    txtImporteApagar.Text = ""
    'FechaRendicion.Text = Date
    cboRecibo.ListIndex = 0
    'txtTotalCheques.Text = ""
    txtTotalEfectivo.Text = ""
    txtTotalValores.Text = ""
    
    'txtTotalComprobante.Text = ""
    tabValores.Tab = 0
    tabComprobantes.Tab = 0
    
    'MANDO RECIBO A PANTALLA
    mImprimoRecibo = True
    
    'CARGO ESTADO
    Call BuscoEstado(1, lblEstadoRecibo) 'ESTADO PENDIENTE
    tabDatos.Tab = 0
    cboRecibo.SetFocus
    txtNroSucursal_LostFocus
    txtNroRecibo_LostFocus
    FechaRecibo_LostFocus
    txtCodCliente.SetFocus
End Sub

'Private Sub cmdNuevoCheque_Click()
'    FrmCargaCheques.Show vbModal
'    'TxtCheNumero.SetFocus
'End Sub
'
'Private Sub cmdQuitarVal_Click()
'
'End Sub

Private Sub QuitoDineroACta()
    For i = 1 To GrillaAFavor.Rows - 1
        If GrillaAFavor.TextMatrix(i, 5) = grillaValores.TextMatrix(grillaValores.RowSel, 5) _
            And CLng(GrillaAFavor.TextMatrix(i, 1)) = CLng(grillaValores.TextMatrix(grillaValores.RowSel, 4)) _
            And CDate(GrillaAFavor.TextMatrix(i, 2)) = CDate(grillaValores.TextMatrix(grillaValores.RowSel, 2)) Then
            
            'ARREGLO EL SALDO DEL DINERO A CTA
            GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 4) = Valido_Importe(CStr(CDbl(GrillaAFavor.TextMatrix(i, 4)) + CDbl(grillaValores.TextMatrix(grillaValores.RowSel, 1))))
           Exit For
        End If
    Next
End Sub

Private Sub cmdSalir_Click()
    'If MsgBox("Seguro que desea Salir", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        Set frmReciboCliente = Nothing
        Unload Me
    'End If
End Sub

Private Sub FechaRecibo_LostFocus()
    If FechaRecibo.Text = "" Then
        FechaRecibo.Text = Date
    End If
End Sub

Private Sub Form_Activate()
    txtCodCliente.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 And _
       ActiveControl.Name <> "txtCodCliente" And _
       ActiveControl.Name <> "txtCliRazSoc" And _
       ActiveControl.Name <> "txtCliente" And _
       ActiveControl.Name <> "txtDesCli" Then tabDatos.Tab = 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then MySendKeys Chr(9)
    If KeyAscii = vbKeyEscape Then cmdSalir_Click
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    Set Rec2 = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
   
    Me.Left = 0
    Me.Top = 0
    
    tabDatos.Tab = 0
    tabValores.Tab = 0
    tabComprobantes.Tab = 0
    
    'CONFIGURO GRILLAS
    configurogrillas
    
    'CARGO COMBO CON LOS TIPOS DE RECIBO
    LlenarComboRecibo
    
    'CARGO COMBO CON LAS PROVINCIAS
    LLenarComboMoneda
    
    'CARGO COMBO CON COMPROBANTES PARA USO DE PAGO
    'Call CargoComboBox(cboComprobantes, "TIPO_COMPROBANTE", "TCO_CODIGO", "TCO_DESCRI")
    'cboComprobantes.ListIndex = 0

    'CARGO ESTADO
    Call BuscoEstado(1, lblEstadoRecibo) 'ESTADO PENDIENTE
    Estado = 1
    '------------------------
    'frameBanco.Enabled = False
    'cmdAgregarCheque.Enabled = False
    'cmdAgregarEfectivo.Enabled = False
    'FechaRendicion.Text = Date
    txtNroRecibo.Enabled = True
    lblEstado.Caption = ""
    mBorroTransfe = False
    
    'MANDO RECIBO A PANTALLA
    mImprimoRecibo = True
    
    txtNroSucursal_LostFocus
    txtNroRecibo_LostFocus
    FechaRecibo_LostFocus
End Sub

Private Sub configurogrillas()
    
    'GRILLA EFECTIVO
    GrillaEfectivo.FormatString = "Moneda|>Importe|Cód.Moneda"
    GrillaEfectivo.ColWidth(0) = 1900 'MONEDA
    GrillaEfectivo.ColWidth(1) = 1000 'IMPORTE
    GrillaEfectivo.ColWidth(2) = 0    'CODIGO MONEDA
    GrillaEfectivo.Rows = 1
    GrillaEfectivo.HighLight = flexHighlightNever
    GrillaEfectivo.BorderStyle = flexBorderNone
    GrillaEfectivo.row = 0
    For i = 0 To GrillaEfectivo.Cols - 1
        GrillaEfectivo.Col = i
        GrillaEfectivo.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        GrillaEfectivo.CellBackColor = &H808080    'GRIS OSCURO
        GrillaEfectivo.CellFontBold = True
    Next

    'GRILLA Aplicar A
    GrillaAplicar.FormatString = "^Comp.|^Número|^Fecha|>Total|>Abona|>Saldo|Cod.Comprob"
    GrillaAplicar.ColWidth(0) = 700  'COMPROBANTE
    GrillaAplicar.ColWidth(1) = 1250 'NUMERO
    GrillaAplicar.ColWidth(2) = 1000 'FECHA
    GrillaAplicar.ColWidth(3) = 700  'TOTAL
    GrillaAplicar.ColWidth(4) = 700  'ABONA
    GrillaAplicar.ColWidth(5) = 700  'SALDO
    GrillaAplicar.ColWidth(6) = 0    'CODIGO COMPROBANTE
    GrillaAplicar.Rows = 1
    GrillaAplicar.HighLight = flexHighlightNever
    GrillaAplicar.BorderStyle = flexBorderNone
    GrillaAplicar.row = 0
    For i = 0 To GrillaAplicar.Cols - 1
        GrillaAplicar.Col = i
        GrillaAplicar.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        GrillaAplicar.CellBackColor = &H808080    'GRIS OSCURO
        GrillaAplicar.CellFontBold = True
    Next
    
    'GRILLA BUSQUEDA
    GrdModulos.FormatString = "^Tipo|^Nro Recibo|^Fecha Recibo|Cliente|Tipo Recibo|>Monto"
    GrdModulos.ColWidth(0) = 1000 'TIPO RECIBO
    GrdModulos.ColWidth(1) = 1600 'NRO RECIBO
    GrdModulos.ColWidth(2) = 1600 'FECHA RECIBO
    GrdModulos.ColWidth(3) = 5000 'CLIENTE
    GrdModulos.ColWidth(4) = 0    'TIPO RECIBO (TCO_CODIGO)
    GrdModulos.ColWidth(5) = 1600 'MONTO
    GrdModulos.Cols = 6
    GrdModulos.Rows = 1
    GrdModulos.HighLight = flexHighlightNever
    GrdModulos.BorderStyle = flexBorderNone
    GrdModulos.row = 0
    For i = 0 To GrdModulos.Cols - 1
        GrdModulos.Col = i
        GrdModulos.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        GrdModulos.CellBackColor = &H808080    'GRIS OSCURO
        GrdModulos.CellFontBold = True
    Next
    
    'GRILLA VALORES
    grillaValores.FormatString = "^Tipo|>Importe||Descripción|Número||"
    grillaValores.ColWidth(0) = 700  'TIPO DE VALOR (CHE,EFT...)
    grillaValores.ColWidth(1) = 1000 'IMPORTE
    grillaValores.ColWidth(2) = 0    'FECHA
    grillaValores.ColWidth(3) = 2000 'DESCRIPCIÓN
    grillaValores.ColWidth(4) = 1350 'NÚMERO
    grillaValores.ColWidth(5) = 0    'CÓDIGO
    grillaValores.ColWidth(6) = 0    'REPRESENTADA
    grillaValores.Rows = 1
    grillaValores.HighLight = flexHighlightNever
    grillaValores.BorderStyle = flexBorderNone
    grillaValores.row = 0
    For i = 0 To grillaValores.Cols - 1
        grillaValores.Col = i
        grillaValores.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        grillaValores.CellBackColor = &H808080    'GRIS OSCURO
        grillaValores.CellFontBold = True
    Next
    
    'GRILLA A FAVOR
    GrillaAFavor.FormatString = "^Comp.|^Número|^Fecha|>Total|>Saldo|codigo comprobante|REPRESENTADA"
    GrillaAFavor.ColWidth(0) = 850  'COMPROBANTE
    GrillaAFavor.ColWidth(1) = 1300 'NUMERO
    GrillaAFavor.ColWidth(2) = 1000 'FECHA
    GrillaAFavor.ColWidth(3) = 1000 'TOTAL
    GrillaAFavor.ColWidth(4) = 1000 'SALDO
    GrillaAFavor.ColWidth(5) = 0    'CODIGO COMPROBANTE
    GrillaAFavor.ColWidth(6) = 0    'REPRESENTADA
    GrillaAFavor.Rows = 1
    GrillaAFavor.HighLight = flexHighlightNever
    GrillaAFavor.BorderStyle = flexBorderNone
    GrillaAFavor.row = 0
    For i = 0 To GrillaAFavor.Cols - 1
        GrillaAFavor.Col = i
        GrillaAFavor.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        GrillaAFavor.CellBackColor = &H808080    'GRIS OSCURO
        GrillaAFavor.CellFontBold = True
    Next
End Sub

Private Sub LlenarComboRecibo()
    SQL = "SELECT * FROM TIPO_COMPROBANTE"
    SQL = SQL & " WHERE TCO_DESCRI LIKE 'RECIBO C%'"
    SQL = SQL & " ORDER BY TCO_DESCRI"
    rec.Open SQL, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cboRecibo1.AddItem "(Todos)"
        Do While rec.EOF = False
            cboRecibo.AddItem rec!TCO_DESCRI
            cboRecibo.ItemData(cboRecibo.NewIndex) = rec!TCO_CODIGO
            cboRecibo1.AddItem rec!TCO_DESCRI
            cboRecibo1.ItemData(cboRecibo1.NewIndex) = rec!TCO_CODIGO
            rec.MoveNext
        Loop
        cboRecibo.ListIndex = 0
        cboRecibo1.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub LLenarComboMoneda()
    SQL = "SELECT * FROM MONEDA ORDER BY MON_DESCRI"
    rec.Open SQL, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboMoneda.AddItem rec!MON_DESCRI
            cboMoneda.ItemData(cboMoneda.NewIndex) = rec!MON_CODIGO
            rec.MoveNext
        Loop
        cboMoneda.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub GrdModulos_dblClick()
     If GrdModulos.Rows > 1 Then
        mBorroTransfe = False
        cmdNuevo_Click
        Call BuscaCodigoProxItemData(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 4)), cboRecibo)
        txtNroRecibo.Text = Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 8)
        txtNroSucursal.Text = Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 4)
        FechaRecibo.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 2)
        tabDatos.Tab = 0
        Call BuscarRecibo(GrdModulos.TextMatrix(GrdModulos.RowSel, 4), Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 8), Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 4))
     End If
End Sub

Private Sub GrdModulos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then GrdModulos_dblClick
End Sub

Private Sub GrillaAFavor_DblClick()
    If GrillaAFavor.Rows > 1 Then
        txtSaldoACta.Text = Valido_Importe(GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 4))
        txtImporteACta.Text = Valido_Importe(GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 4))
        txtImporteACta.SetFocus
    End If
End Sub

Private Sub GrillaAFavor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If GrillaAFavor.Rows > 1 Then
           GrillaAFavor_DblClick
        End If
    End If
End Sub

Private Sub GrillaEfectivo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If GrillaEfectivo.Rows > 2 Then
           GrillaEfectivo.RemoveItem GrillaEfectivo.RowSel
        Else
           GrillaEfectivo.Rows = 1
           GrillaEfectivo.HighLight = flexHighlightNever
           cboMoneda.SetFocus
        End If
        txtTotalEfectivo.Text = SumaGrilla(GrillaEfectivo, 1)
        'txtTotalValores.Text = Valido_Importe(CStr(CDbl(SumaGrilla(GrillaCheques, 6)) + CDbl(SumaGrilla(GrillaEfectivo, 1))))
    End If
End Sub

Private Sub grillaValores_DblClick()
    If grillaValores.Rows > 1 Then
        If MsgBox("¿Seguro que desea eliminar?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
            If grillaValores.Rows > 2 Then
                If grillaValores.TextMatrix(grillaValores.RowSel, 0) = "A-CTA" Then
                    QuitoDineroACta
                ElseIf grillaValores.TextMatrix(grillaValores.RowSel, 0) = "COMP" Then
                    'VEO SI ES UNA TRASFERENCIA
                    'SI ES LA BORRO DE LA TABLA DEBCRE_BANCARIOS
                    If grillaValores.TextMatrix(grillaValores.RowSel, 5) = 30 Then
                        DBConn.Execute "DELETE FROM DEBCRE_BANCARIOS WHERE DCB_NUMERO = " & XN(Right(Trim(grillaValores.TextMatrix(grillaValores.RowSel, 4)), 8))
                    End If
                End If
                grillaValores.RemoveItem grillaValores.RowSel
                txtTotalValores.Text = SumaGrilla(grillaValores, 1)
            Else
                If grillaValores.TextMatrix(grillaValores.RowSel, 0) = "A-CTA" Then
                    QuitoDineroACta
                ElseIf grillaValores.TextMatrix(grillaValores.RowSel, 0) = "COMP" Then
                    'VEO SI ES UNA TRASFERENCIA
                    'SI ES LA BORRO DE LA TABLA DEBCRE_BANCARIOS
                    If grillaValores.TextMatrix(grillaValores.RowSel, 5) = 30 Then
                        DBConn.Execute "DELETE FROM DEBCRE_BANCARIOS WHERE DCB_NUMERO = " & XN(Right(Trim(grillaValores.TextMatrix(grillaValores.RowSel, 4)), 8))
                    End If
                End If
                grillaValores.Rows = 1
                txtTotalValores.Text = ""
                grillaValores.HighLight = flexHighlightNever
            End If
        End If
    End If
End Sub

Private Sub tabComprobantes_Click(PreviousTab As Integer)
    If tabComprobantes.Tab = 1 Then
        GrillaAplicar.SetFocus
    End If
    If tabComprobantes.Tab = 0 Then
        'If Me.tabComprobantes.Visible = True Then cmdAgregarFactura.SetFocus
        If GrillaAplicar.Rows > 1 Then
          ' txtTotalAplicar.Text = Valido_Importe(SumaGrilla(GrillaAplicar, 1))
        End If
    End If
End Sub

Private Sub tabDatos_Click(PreviousTab As Integer)
  If tabDatos.Tab = 1 Then
    'LimpiarBusqueda
    cmdGrabar.Enabled = False
    If Me.Visible = True Then txtCliente.SetFocus
  End If
End Sub

Private Sub LimpiarBusqueda()
    txtCliente.Text = ""
    txtDesCli.Text = ""
    FechaDesde.Text = ""
    FechaHasta.Text = ""
    GrdModulos.Rows = 1
    GrdModulos.HighLight = flexHighlightNever
   ' cboBuscaRep.ListIndex = -1
    cboRecibo1.ListIndex = 0
   ' chkRepresentada.Value = Unchecked
End Sub

Private Sub tabValores_Click(PreviousTab As Integer)
    If tabValores.Tab = 1 Then
       BuscaProx "PESOS", cboMoneda
       txtEftImporte.SetFocus
    ElseIf tabValores.Tab = 2 Then
       If GrillaAFavor.Rows > 1 Then
          GrillaAFavor.Col = 0
          GrillaAFavor.row = 1
          GrillaAFavor.SetFocus
       End If
    End If
End Sub

Private Sub txtCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarClientes "txtCliente", "CODIGO"
    End If
End Sub

Private Sub txtCliRazSoc_Change()
    If txtCliRazSoc.Text = "" Then
        txtCodCliente.Text = ""
        txtDomici.Text = ""
    End If
End Sub

Private Sub txtCliRazSoc_GotFocus()
    SelecTexto txtCliRazSoc
End Sub

Private Sub txtCliRazSoc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarClientes "txtCodCliente", "CODIGO"
    End If
End Sub

Private Sub txtCliRazSoc_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtCliRazSoc_LostFocus()
    If txtCodCliente.Text = "" And txtCliRazSoc.Text <> "" Then
        SQL = "SELECT C.CLI_CODIGO,C.CLI_RAZSOC,C.CLI_DOMICI"
        SQL = SQL & "  FROM CLIENTE C "
        SQL = SQL & " WHERE C.CLI_RAZSOC LIKE '" & txtCliRazSoc.Text & "%'"
        If rec.State = 1 Then rec.Close
        rec.Open SQL, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            If rec.RecordCount > 1 Then
                BuscarClientes "txtCodCliente", "CADENA", Trim(txtCliRazSoc.Text)
                If rec.State = 1 Then rec.Close
                txtCliRazSoc.SetFocus
            Else
                txtCodCliente.Text = rec!CLI_CODIGO
                txtCliRazSoc.Text = rec!CLI_RAZSOC
                txtDomici.Text = ChkNull(rec!CLI_DOMICI)
                If Estado = 1 Then
                    If BuscarFactura(txtCodCliente) = False Then
                        MsgBox "El Cliente NO tiene facturas pendientes de pago. Verifique!", vbExclamation, TIT_MSGBOX
                        txtCodCliente.Text = ""
                        txtCodCliente.SetFocus
                        FrameRecibo.Enabled = True
                    Else
                        Call BuscarSaldosAFavor(txtCodCliente)
                        FrameRecibo.Enabled = False
                        txtImporteApagar.SetFocus
                        'GrillaAplicar.SetFocus
                    End If
                End If
            End If
        Else
            MsgBox "No se encontro el Cliente", vbExclamation, TIT_MSGBOX
            txtCodCliente.SetFocus
        End If
        If rec.State = 1 Then rec.Close
    End If
End Sub

Private Sub txtCliente_Change()
    If txtCliente.Text = "" Then
        txtDesCli.Text = ""
    End If
End Sub

Private Sub txtCliente_GotFocus()
    SelecTexto txtCliente
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCliente_LostFocus()
    If txtCliente.Text <> "" Then
        Set rec = New ADODB.Recordset
        SQL = "SELECT CLI_RAZSOC FROM CLIENTE"
        SQL = SQL & " WHERE CLI_CODIGO=" & XN(txtCliente)
        rec.Open SQL, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtDesCli.Text = rec!CLI_RAZSOC
        Else
            MsgBox "El Cliente no existe", vbExclamation, TIT_MSGBOX
            txtDesCli.Text = ""
            txtCliente.SetFocus
        End If
        If rec.State = 1 Then rec.Close
    End If
End Sub

Private Sub cmdBuscarCli_Click()
    frmBuscar.TipoBusqueda = 1
    frmBuscar.TxtDescriB = ""
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Text <> "" Then
        frmBuscar.grdBuscar.Col = 0
        txtCliente.Text = frmBuscar.grdBuscar.Text
        txtCliente.SetFocus
        txtCliente_LostFocus
    Else
        txtCliente.SetFocus
    End If
End Sub

Private Sub txtCodCliente_Change()
    If txtCodCliente.Text = "" Then
        txtCliRazSoc.Text = ""
        txtDomici.Text = ""
    End If
End Sub

Private Sub txtCodCliente_GotFocus()
    SelecTexto txtCodCliente
End Sub

Private Sub txtCodCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarClientes "txtCodCliente", "CODIGO"
    End If
End Sub

Private Sub txtCodCliente_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCodCliente_LostFocus()
    If txtCodCliente.Text <> "" Then
        Set rec = New ADODB.Recordset
        SQL = "SELECT C.CLI_RAZSOC,C.CLI_DOMICI"
        SQL = SQL & " FROM CLIENTE C"
        SQL = SQL & " WHERE"
        SQL = SQL & " C.CLI_CODIGO=" & XN(txtCodCliente)
        rec.Open SQL, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtCliRazSoc.Text = rec!CLI_RAZSOC
            txtDomici.Text = ChkNull(rec!CLI_DOMICI)
            If Estado = 1 Then
                If BuscarFactura(txtCodCliente) = False Then
                    MsgBox "El Cliente NO tiene facturas pendientes de pago. Verifique!", vbExclamation, TIT_MSGBOX
                    txtCodCliente.Text = ""
                    txtCodCliente.SetFocus
                    FrameRecibo.Enabled = True
                Else
                    Call BuscarSaldosAFavor(txtCodCliente)
                    FrameRecibo.Enabled = False
                    GrillaAplicar.SetFocus
                End If
            End If
        Else
            MsgBox "El Cliente no existe", vbExclamation, TIT_MSGBOX
            FrameRecibo.Enabled = True
            txtCliRazSoc.Text = ""
            txtCodCliente.SetFocus
        End If
        If rec.State = 1 Then rec.Close
    End If
End Sub

Private Sub txtDesCli_Change()
    If txtDesCli.Text = "" Then
        txtCliente.Text = ""
    End If
End Sub

Private Sub txtDesCli_GotFocus()
    SelecTexto txtDesCli
End Sub

Private Sub txtDesCli_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarClientes "txtCliente", "CODIGO"
    End If
End Sub

Private Sub txtDesCli_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtDesCli_LostFocus()
    If txtCliente.Text = "" And txtDesCli.Text <> "" Then
        SQL = "SELECT C.CLI_CODIGO,C.CLI_RAZSOC,C.CLI_DOMICI"
        SQL = SQL & " FROM CLIENTE C"
        SQL = SQL & " WHERE"
        SQL = SQL & " C.CLI_RAZSOC LIKE '" & txtDesCli.Text & "%'"
        If rec.State = 1 Then rec.Close
        rec.Open SQL, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            If rec.RecordCount > 1 Then
                BuscarClientes "txtCliente", "CADENA", Trim(txtDesCli.Text)
                If rec.State = 1 Then rec.Close
                txtDesCli.SetFocus
            Else
                txtCliente.Text = rec!CLI_CODIGO
                txtDesCli.Text = rec!CLI_RAZSOC
            End If
        Else
            MsgBox "No se encontro el Cliente", vbExclamation, TIT_MSGBOX
            txtCliente.SetFocus
        End If
        If rec.State = 1 Then rec.Close
    End If
End Sub

Private Sub txtEftImporte_Change()
    'If txtEftImporte.Text = "" Then
    '    cmdAgregarEfectivo.Enabled = False
    'Else
    '    cmdAgregarEfectivo.Enabled = True
    'End If
End Sub

Private Sub txtEftImporte_GotFocus()
    
    
    SelecTexto txtEftImporte
End Sub

Private Sub txtEftImporte_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtEftImporte, KeyAscii)
End Sub

Private Sub txtEftImporte_LostFocus()
    If txtEftImporte.Text <> "" Then
        txtEftImporte.Text = Valido_Importe(txtEftImporte.Text)
        cmdAgregarEfectivo.Enabled = True
        cmdAgregarEfectivo.SetFocus
    End If
End Sub

Private Sub txtImporteACta_Change()
    If txtSaldoACta.Text <> "" And txtImporteACta.Text <> "" Then
        cmdAgregarACta.Enabled = True
    Else
        cmdAgregarACta.Enabled = False
    End If
End Sub

Private Sub txtImporteACta_GotFocus()
    SelecTexto txtImporteACta
End Sub

Private Sub txtImporteACta_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtImporteACta, KeyAscii)
End Sub

Private Sub txtImporteACta_LostFocus()
    If txtSaldoACta.Text <> "" Then
        If txtImporteACta.Text = "" Then
            txtImporteACta.Text = txtSaldoACta.Text
        ElseIf CDbl(txtImporteACta.Text) > CDbl(txtSaldoACta.Text) Then
            MsgBox "Importe mayor al Saldo. Verifique!", vbCritical, TIT_MSGBOX
            txtImporteACta.Text = txtSaldoACta.Text
            txtImporteACta.SetFocus
        End If
        txtImporteACta.Text = Valido_Importe(txtImporteACta)
    End If
End Sub

Private Sub txtImporteApagar_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtImporteApagar, KeyAscii)
End Sub

Private Sub txtImporteApagar_LostFocus()
  If Me.ActiveControl.Name <> "CmdSalir" Then
    txtImporteApagar.Text = Valido_Importe(txtImporteApagar.Text)
    
    If GrillaAFavor.Rows > 1 Then
       tabValores.Tab = 2
       GrillaAFavor.Col = 0
       GrillaAFavor.row = 1
    Else
       tabValores.Tab = 1
       BuscaProx "PESOS", cboMoneda
       txtEftImporte.SetFocus
    End If
  End If
End Sub

Private Function BuscarFactura(CodCli As String) As Boolean
        GrillaAplicar.Rows = 1
        Set Rec1 = New ADODB.Recordset
        Dim TotalDeuda As Double
        TotalDeuda = 0
        'BUSCA LAS FACTURAS
        SQL = "SELECT FCL_NUMERO AS NUMERO, FCL_SUCURSAL AS SUCURSAL, "
        SQL = SQL & " FCL_FECHA AS FECHA, FCL_TOTAL AS TOTAL, FCL_SALDO AS SALDO"
        SQL = SQL & " ,TCO_CODIGO AS TIPO, TCO_ABREVIA AS ABREVIA"
        SQL = SQL & " FROM SALDO_FACTURAS_CLIENTE_V"
        SQL = SQL & " WHERE "
        SQL = SQL & " CLI_CODIGO=" & XN(CodCli)
        SQL = SQL & " UNION ALL"
        
        'BUSCA LAS NOTA DE DEBITO
        SQL = SQL & " SELECT NDC_NUMERO AS NUMERO, NDC_SUCURSAL AS SUCURSAL, "
        SQL = SQL & " NDC_FECHA AS FECHA, NDC_TOTAL AS TOTAL, NDC_SALDO AS SALDO"
        SQL = SQL & " ,TCO_CODIGO AS TIPO, TCO_ABREVIA AS ABREVIA"
        SQL = SQL & " FROM SALDO_NOTA_DEBITO_CLIENTE_V"
        SQL = SQL & " WHERE "
        SQL = SQL & " CLI_CODIGO=" & XN(CodCli)
        SQL = SQL & " ORDER BY FECHA , NUMERO ASC"
        
        Rec1.Open SQL, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            Do While Rec1.EOF = False
                If Rec1!Saldo > 0 Then
                    GrillaAplicar.AddItem Rec1!ABREVIA & Chr(9) & Format(Rec1!Sucursal, "0000") & "-" & Format(Rec1!Numero, "00000000") _
                                    & Chr(9) & Rec1!Fecha & Chr(9) & Valido_Importe(Rec1!TOTAL) _
                                    & Chr(9) & "0,00" & Chr(9) & Valido_Importe(Rec1!Saldo) & Chr(9) & Rec1!Tipo
                    TotalDeuda = CDbl(TotalDeuda) + Valido_Importe(Rec1!Saldo)
                End If
                Rec1.MoveNext
            Loop
            GrillaAplicar.HighLight = flexHighlightAlways
            BuscarFactura = True
            txtSaldo.Text = Format(TotalDeuda, "0.00")
        Else
            BuscarFactura = False
        End If
        Rec1.Close
End Function

Private Sub BuscarSaldosAFavor(CodCli As String)
        GrillaAFavor.Rows = 1
        Set Rec1 = New ADODB.Recordset
        SQL = "SELECT RS.TCO_CODIGO, RS.REC_NUMERO, RS.REC_SUCURSAL, RS.REC_FECHA,"
        SQL = SQL & " RS.REC_TOTSALDO,RS.REC_SALDO, T.TCO_ABREVIA"
        SQL = SQL & " FROM RECIBO_CLIENTE_SALDO RS, RECIBO_CLIENTE R,TIPO_COMPROBANTE T"
        SQL = SQL & " WHERE RS.TCO_CODIGO = T.TCO_CODIGO"
        SQL = SQL & "   AND RS.TCO_CODIGO = R.TCO_CODIGO"
        SQL = SQL & "   AND RS.REC_NUMERO = R.REC_NUMERO"
        SQL = SQL & "   AND RS.REC_SUCURSAL = R.REC_SUCURSAL"
        SQL = SQL & "   AND RS.REC_FECHA = R.REC_FECHA"
        SQL = SQL & "   AND RS.REC_SALDO > 0"
        SQL = SQL & "   AND R.CLI_CODIGO = " & XN(CodCli)
        SQL = SQL & " ORDER BY RS.TCO_CODIGO,RS.REC_SUCURSAL,RS.REC_NUMERO, RS.REC_FECHA"
        Rec1.Open SQL, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            GrillaAFavor.HighLight = flexHighlightAlways
            Do While Rec1.EOF = False
               If Rec1!REC_SALDO > 0 Then
                  GrillaAFavor.AddItem Rec1!TCO_ABREVIA & Chr(9) & Format(Rec1!REC_SUCURSAL, "0000") & "-" & Format(Rec1!REC_NUMERO, "00000000") & Chr(9) & Rec1!REC_FECHA & Chr(9) & Valido_Importe(Rec1!REC_TOTSALDO) & Chr(9) & Valido_Importe(Rec1!REC_SALDO) & Chr(9) & Rec1!TCO_CODIGO
               End If
               Rec1.MoveNext
            Loop
        End If
        Rec1.Close
        
        If GrillaAFavor.Rows > 1 Then
           LblDineroaCta.Caption = "El Cliente tiene Dinero a Cuenta"
        Else
           LblDineroaCta.Caption = ""
        End If
                        
        'BUSCO ANTICIPOS DE COBRO
        'sql = "SELECT A.ANC_NUMERO, A.ANC_FECHA, A.ANC_SUCURSAL,"
        'sql = sql & " A.ANC_MONTO,A.ANC_SALDO"
        'sql = sql & " FROM ANTICIPO_COBRO A, CLIENTE C"
        'sql = sql & " WHERE"
        'sql = sql & " A.CLI_CODIGO=C.CLI_CODIGO"
        'sql = sql & " AND A.ANC_SALDO > 0"
        'sql = sql & " AND A.CLI_CODIGO=" & XN(CodCli)
        'sql = sql & " ORDER BY A.ANC_FECHA,A.ANC_NUMERO"
        'Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        'If Rec1.EOF = False Then
        '    GrillaAFavor.HighLight = flexHighlightAlways
        '    Do While Rec1.EOF = False
        '        If Rec1!ANC_SALDO > 0 Then
        '            GrillaAFavor.AddItem "ANT-COB" & Chr(9) & Format(Rec1!ANC_SUCURSAL, "0000") & "-" & Format(Rec1!ANC_NUMERO, "00000000") _
        '                            & Chr(9) & Rec1!ANC_FECHA & Chr(9) & Valido_Importe(Rec1!ANC_MONTO) _
        '                            & Chr(9) & Valido_Importe(Rec1!ANC_SALDO) & Chr(9) & "19" 'TIPO DE COMPROBANTE NRO 19
        '        End If
        '        Rec1.MoveNext
        '    Loop
        'End If
        'Rec1.Close
End Sub

Private Function BuscoComprobanteEnRecibo() As Boolean
'    Set Rec2 = New ADODB.Recordset
'
'    sql = "SELECT DR.REC_NUMERO"
'    sql = sql & " FROM DETALLE_RECIBO_CLIENTE DR, RECIBO_CLIENTE RC"
'    sql = sql & " WHERE"
'    sql = sql & " DR.DRE_TCO_CODIGO=" & XN(cboComprobantes.ItemData(cboComprobantes.ListIndex))
'    sql = sql & " AND DR.DRE_COMNUMERO=" & XN(txtNroComprobantes)
'    sql = sql & " AND DR.DRE_COMSUCURSAL=" & XN(txtNroCompSuc)
'    sql = sql & " AND RC.CLI_CODIGO=" & XN(txtCodCliente)
'    sql = sql & " AND DR.REC_NUMERO=RC.REC_NUMERO"
'    sql = sql & " AND DR.REC_SUCURSAL=RC.REC_SUCURSAL"
'    sql = sql & " AND DR.TCO_CODIGO=RC.TCO_CODIGO"
'    sql = sql & " AND RC.EST_CODIGO=3"
'    Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
'
'    If Rec2.EOF = False Then
'        BuscoComprobanteEnRecibo = False
'    Else
'        BuscoComprobanteEnRecibo = True
'    End If
'    Rec2.Close
    
End Function

Private Sub txtNroRecibo_Change()
    If txtNroRecibo.Text = "" Then
        FechaRecibo.Text = ""
    End If
End Sub

Private Sub txtNroRecibo_GotFocus()
    SelecTexto txtNroRecibo
End Sub

Private Sub txtNroRecibo_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroRecibo_LostFocus()
    If txtNroRecibo.Text = "" Then
        'BUSCO EL NUMERO DE RECIBO QUE CORRESPONDE
        txtNroRecibo.Text = Format(BuscoUltimoRecibo(cboRecibo.ItemData(cboRecibo.ListIndex)), "00000000")
    Else
        If txtNroSucursal.Text = "" Then
            txtNroSucursal.Text = Sucursal
        End If
        txtNroRecibo.Text = Format(txtNroRecibo.Text, "00000000")
        Call BuscarRecibo(CStr(cboRecibo.ItemData(cboRecibo.ListIndex)), _
                          txtNroRecibo, txtNroSucursal)
    End If
End Sub

Private Function BuscoUltimoRecibo(TipoRec As Integer) As String
    'ACA BUSCA EL NUMERO DE REMITO SIGUIENTE AL ULTIMO CARGADO
    SQL = "SELECT (RECIBO_C) + 1 AS REC_C"
    SQL = SQL & " FROM PARAMETROS"
    rec.Open SQL, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Select Case TipoRec
            Case 12
                BuscoUltimoRecibo = IIf(IsNull(rec!REC_C), 1, rec!REC_C)
        End Select
    End If
    rec.Close
End Function

Private Sub BuscarRecibo(TipoRec As String, NroRec As String, NroSuc As String)
    Set Rec2 = New ADODB.Recordset
    
    SQL = "SELECT * "
    SQL = SQL & "  FROM RECIBO_CLIENTE"
    SQL = SQL & " WHERE TCO_CODIGO = " & XN(TipoRec)
    SQL = SQL & "   AND REC_NUMERO = " & XN(NroRec)
    SQL = SQL & "   AND REC_SUCURSAL = " & XN(NroSuc)
    Rec2.Open SQL, DBConn, adOpenStatic, adLockOptimistic
    If Rec2.EOF = False Then
        If Rec2.RecordCount > 2 Then
            Rec2.Close
            tabDatos.Tab = 1
            Exit Sub
        End If
        'CABEZA DEL RECIDO
        FechaRecibo.Text = Rec2!REC_FECHA
        'FechaRendicion.Text = Rec2!REC_FECHA_RENDICION
        'CARGO ESTADO
        Call BuscoEstado(CInt(Rec2!EST_CODIGO), lblEstadoRecibo)
        Estado = CInt(Rec2!EST_CODIGO)
        txtCodCliente.Text = Rec2!CLI_CODIGO
        txtCodCliente_LostFocus
        
        'DETALLE_DEL RECIBO CHEQUES
        Set rec = New ADODB.Recordset
        SQL = "SELECT *"
        SQL = SQL & " FROM DETALLE_RECIBO_CLIENTE"
        SQL = SQL & " WHERE TCO_CODIGO =" & XN(TipoRec)
        SQL = SQL & "   AND REC_NUMERO =" & XN(NroRec)
        SQL = SQL & "   AND REC_SUCURSAL =" & XN(NroSuc)
        rec.Open SQL, DBConn, adOpenStatic, adLockOptimistic
        
        If rec.EOF = False Then
            Do While rec.EOF = False
                'If Not IsNull(rec!BAN_CODINT) Then 'BANCO
                '    Call BuscarCheque(rec!BAN_CODINT, rec!CHE_NUMERO)
                'Else
                If Not IsNull(rec!MON_CODIGO) Then 'MONEDA
                    grillaValores.AddItem "EFT" & Chr(9) & Valido_Importe(rec!DRE_MONIMP) _
                                    & Chr(9) & "" & Chr(9) & BuscarMoneda(rec!MON_CODIGO) _
                                    & Chr(9) & "" & Chr(9) & rec!MON_CODIGO
                                    
                ElseIf Not IsNull(rec!DRE_TCO_CODIGO) Then 'COMPROBANTE
                    Dim QueEs As String
                    If rec!DRE_TCO_CODIGO >= 10 And rec!DRE_TCO_CODIGO <= 13 Then
                        QueEs = "A-CTA"
                    ElseIf (rec!DRE_TCO_CODIGO = 19) Then
                        QueEs = "A-CTA"
                    Else
                        QueEs = "COMP"
                    End If
                    grillaValores.AddItem QueEs & Chr(9) & Valido_Importe(rec!DRE_COMIMP) _
                                    & Chr(9) & rec!DRE_COMFECHA & Chr(9) & BuscarTipoDocAbre(rec!DRE_TCO_CODIGO) _
                                    & Chr(9) & Format(ChkNull(rec!DRE_COMSUCURSAL), "0000") & "-" & Format(rec!DRE_COMNUMERO, "00000000") _
                                    & Chr(9) & rec!DRE_TCO_CODIGO
                End If
                rec.MoveNext
            Loop
            
            grillaValores.HighLight = flexHighlightAlways
            txtTotalValores.Text = SumaGrilla(grillaValores, 1)
        End If
        rec.Close
                   
        'DETALLE_DEL RECIBO FACTURA
        SQL = "SELECT * "
        SQL = SQL & " FROM FACTURAS_RECIBO_CLIENTE"
        SQL = SQL & " WHERE TCO_CODIGO=" & XN(TipoRec)
        SQL = SQL & "   AND REC_NUMERO=" & XN(NroRec)
        SQL = SQL & "   AND REC_SUCURSAL=" & XN(NroSuc)
        rec.Open SQL, DBConn, adOpenStatic, adLockOptimistic
        
        If rec.EOF = False Then
            Do While rec.EOF = False
                GrillaAplicar.AddItem BuscarTipoDocAbre(rec!FCL_TCO_CODIGO) & Chr(9) & _
                            Format(rec!FCL_SUCURSAL, "0000") & "-" & Format(rec!FCL_NUMERO, "00000000") & Chr(9) & rec!FCL_FECHA _
                             & Chr(9) & Valido_Importe(rec!REC_IMPORTE) & Chr(9) & Valido_Importe(rec!REC_ABONA) & Chr(9) & Valido_Importe(rec!REC_SALDO) & Chr(9) & rec!FCL_TCO_CODIGO
                            
                rec.MoveNext
            Loop
            GrillaAplicar.HighLight = flexHighlightAlways
            txtImporteApagar.Text = SumaGrilla(GrillaAplicar, 4)
        End If
        FrameRecibo.Enabled = False
        FrameRemito.Enabled = False
        rec.Close
        cmdNuevo.SetFocus
        cmdGrabar.Enabled = False
        mBorroTransfe = False
    End If
    Rec2.Close
End Sub

Private Function BuscarCheque(Codigo As String, NroChe As String) As String
    
    Set Rec1 = New ADODB.Recordset
    SQL = "SELECT B.BAN_DESCRI,C.CHE_IMPORT,C.CHE_FECVTO"
    SQL = SQL & " FROM BANCO B, CHEQUE C"
    SQL = SQL & " WHERE C.BAN_CODINT=" & XN(Codigo)
    SQL = SQL & " AND C.CHE_NUMERO=" & XS(NroChe)
    SQL = SQL & " AND C.BAN_CODINT=B.BAN_CODINT"
    Rec1.Open SQL, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        grillaValores.AddItem "CHE" & Chr(9) & Valido_Importe(Rec1!CHE_IMPORT) & Chr(9) & Rec1!CHE_FECVTO _
                           & Chr(9) & Rec1!BAN_DESCRI & Chr(9) & NroChe & Chr(9) & Codigo
    End If
    Rec1.Close
End Function

Private Function BuscarMoneda(Codigo As String) As String
    
    Set Rec1 = New ADODB.Recordset
    SQL = "SELECT MON_DESCRI"
    SQL = SQL & " FROM MONEDA"
    SQL = SQL & " WHERE MON_CODIGO=" & XN(Codigo)
    Rec1.Open SQL, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        BuscarMoneda = Rec1!MON_DESCRI
    Else
        BuscarMoneda = ""
    End If
    Rec1.Close
End Function

Private Sub txtNroSucursal_GotFocus()
    SelecTexto txtNroSucursal
End Sub

Private Sub txtNroSucursal_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroSucursal_LostFocus()
    If txtNroSucursal.Text = "" Then
        txtNroSucursal.Text = Sucursal
    Else
        txtNroSucursal.Text = Format(txtNroSucursal.Text, "0000")
    End If
End Sub

Public Sub BuscarClientes(Txt As String, mQuien As String, Optional mCadena As String)
    Dim cSQL As String
    Dim hSQL As String
    Dim B As CBusqueda
    Dim i, posicion As Integer
    Dim cadena As String
    
    Set B = New CBusqueda
    With B
        cSQL = "SELECT CLI_RAZSOC, CLI_CODIGO"
        cSQL = cSQL & " FROM CLIENTE C"
        If mQuien = "CADENA" Then
            cSQL = cSQL & " WHERE CLI_RAZSOC LIKE '" & Trim(mCadena) & "%'"
        End If
        
        hSQL = "Nombre, Código"
        .SQL = cSQL
        .Headers = hSQL
        .Field = "CLI_RAZSOC"
        campo1 = .Field
        .Field = "CLI_CODIGO"
        campo2 = .Field
        .OrderBy = "CLI_RAZSOC"
        camponumerico = False
        .Titulo = "Busqueda de Clientes :"
        .MaxRecords = 1
        .Show

        ' utilizar la coleccion de datos devueltos
        If .ResultFields.Count > 0 Then
            If Txt = "txtCodCliente" Then
                txtCodCliente.Text = .ResultFields(2)
                txtCodCliente_LostFocus
            Else
                txtCliente.Text = .ResultFields(2)
                txtCliente_LostFocus
            End If
        End If
    End With
    
    Set B = Nothing
End Sub


