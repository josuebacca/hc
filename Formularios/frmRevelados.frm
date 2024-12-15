VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{5F09B5DF-6F4D-11D2-8355-4854E82A9183}#15.0#0"; "Fecha32.ocx"
Begin VB.Form frmRevelados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Revelados..."
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11280
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   7845
   ScaleWidth      =   11280
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   450
      Left            =   8080
      TabIndex        =   1
      Top             =   7380
      Width           =   990
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   450
      Left            =   10080
      TabIndex        =   3
      Top             =   7380
      Width           =   990
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      Height          =   450
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7380
      Width           =   990
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   450
      Left            =   9080
      TabIndex        =   2
      Top             =   7380
      Width           =   990
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   7300
      Left            =   45
      TabIndex        =   6
      Top             =   45
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   12885
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   5
      TabHeight       =   512
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Datos"
      TabPicture(0)   =   "frmRevelados.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "freCliente"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "freRemito"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "SSTab1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "frmRevelados.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GrdModulos"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "frameBuscar"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin TabDlg.SSTab SSTab1 
         Height          =   5055
         Left            =   120
         TabIndex        =   52
         Top             =   2160
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   8916
         _Version        =   393216
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "Detalle del Revelado Actual"
         TabPicture(0)   =   "frmRevelados.frx":0038
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame9"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Frame10"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Revelados Anteriores"
         TabPicture(1)   =   "frmRevelados.frx":0054
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "MSFlexGrid1"
         Tab(1).ControlCount=   1
         Begin VB.Frame Frame10 
            Caption         =   "Datos de la Entrega"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   1215
            Left            =   5520
            TabIndex        =   101
            Top             =   3720
            Width           =   5295
            Begin VB.ComboBox Combo1 
               Height          =   315
               Left            =   360
               TabIndex        =   115
               Text            =   "Combo1"
               Top             =   240
               Width           =   2175
            End
            Begin VB.CheckBox Check1 
               Caption         =   "Entregado"
               Height          =   255
               Left            =   360
               TabIndex        =   114
               Top             =   720
               Width           =   2175
            End
            Begin FechaCtl.Fecha Fecha1 
               Height          =   285
               Left            =   3915
               TabIndex        =   110
               Top             =   240
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   503
               Separador       =   "/"
               Text            =   ""
               MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
            End
            Begin FechaCtl.Fecha Fecha2 
               Height          =   285
               Left            =   3915
               TabIndex        =   112
               Top             =   720
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   503
               Separador       =   "/"
               Text            =   ""
               MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               Caption         =   "Fecha Entrega:"
               Height          =   195
               Left            =   2760
               TabIndex        =   113
               Top             =   765
               Width           =   1095
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               Caption         =   "Fecha Retiro:"
               Height          =   195
               Left            =   2880
               TabIndex        =   111
               Top             =   285
               Width           =   960
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Datos del Revelado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   1215
            Left            =   120
            TabIndex        =   100
            Top             =   3720
            Width           =   5295
            Begin VB.TextBox Text11 
               Height          =   315
               Left            =   3960
               TabIndex        =   109
               Top             =   720
               Width           =   975
            End
            Begin VB.TextBox Text10 
               Height          =   315
               Left            =   3960
               TabIndex        =   107
               Top             =   240
               Width           =   975
            End
            Begin VB.TextBox Text9 
               Height          =   315
               Left            =   1560
               TabIndex        =   105
               Top             =   720
               Width           =   975
            End
            Begin VB.TextBox Text8 
               Height          =   315
               Left            =   1560
               TabIndex        =   103
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               Caption         =   "Seña:"
               Height          =   195
               Left            =   3495
               TabIndex        =   108
               Top             =   780
               Width           =   420
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "Valor de la Foto:"
               Height          =   195
               Left            =   2760
               TabIndex        =   106
               Top             =   300
               Width           =   1155
            End
            Begin VB.Label Label17 
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
               Left            =   1080
               TabIndex        =   104
               Top             =   780
               Width           =   405
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "Fotos Reveladas:"
               Height          =   195
               Left            =   240
               TabIndex        =   102
               Top             =   300
               Width           =   1245
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Datos del Rollo Fotográfico"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   3375
            Left            =   120
            TabIndex        =   59
            Top             =   360
            Width           =   10695
            Begin VB.TextBox Text6 
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   1440
               TabIndex        =   99
               Top             =   3000
               Width           =   9135
            End
            Begin VB.Frame Frame4 
               Caption         =   "Formato"
               ForeColor       =   &H8000000D&
               Height          =   975
               Left            =   7080
               TabIndex        =   68
               Top             =   240
               Width           =   3495
               Begin VB.TextBox Text4 
                  Height          =   315
                  Left            =   2640
                  TabIndex        =   88
                  Top             =   450
                  Width           =   615
               End
               Begin VB.CommandButton Command16 
                  BackColor       =   &H8000000E&
                  Caption         =   "OTRO"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   1920
                  Style           =   1  'Graphical
                  TabIndex        =   87
                  Top             =   360
                  Width           =   615
               End
               Begin VB.CommandButton Command15 
                  BackColor       =   &H8000000E&
                  Caption         =   "125"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   1320
                  Style           =   1  'Graphical
                  TabIndex        =   86
                  Top             =   360
                  Width           =   615
               End
               Begin VB.CommandButton Command14 
                  BackColor       =   &H8000000E&
                  Caption         =   "110"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   720
                  Style           =   1  'Graphical
                  TabIndex        =   85
                  Top             =   360
                  Width           =   615
               End
               Begin VB.CommandButton Command8 
                  BackColor       =   &H8000000E&
                  Caption         =   "135"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   120
                  MaskColor       =   &H00FFFFFF&
                  Style           =   1  'Graphical
                  TabIndex        =   84
                  Top             =   360
                  Width           =   615
               End
               Begin VB.Label Label9 
                  AutoSize        =   -1  'True
                  Caption         =   "Elegido"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H8000000D&
                  Height          =   195
                  Left            =   2640
                  TabIndex        =   89
                  Top             =   240
                  Width           =   645
               End
            End
            Begin VB.Frame Frame8 
               Caption         =   "Papel"
               ForeColor       =   &H8000000D&
               Height          =   855
               Left            =   7080
               TabIndex        =   83
               Top             =   2040
               Width           =   3495
               Begin VB.TextBox Text5 
                  Height          =   315
                  Left            =   2640
                  TabIndex        =   96
                  Top             =   380
                  Width           =   615
               End
               Begin VB.CommandButton Command19 
                  BackColor       =   &H8000000E&
                  Caption         =   "OTRO"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   1560
                  Style           =   1  'Graphical
                  TabIndex        =   95
                  Top             =   300
                  Width           =   735
               End
               Begin VB.CommandButton Command18 
                  BackColor       =   &H8000000E&
                  Caption         =   "Mate"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   840
                  Style           =   1  'Graphical
                  TabIndex        =   94
                  Top             =   300
                  Width           =   735
               End
               Begin VB.CommandButton Command17 
                  BackColor       =   &H8000000E&
                  Caption         =   "Brillo"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   120
                  MaskColor       =   &H00FFFFFF&
                  Style           =   1  'Graphical
                  TabIndex        =   93
                  Top             =   300
                  Width           =   735
               End
               Begin VB.Label Label12 
                  AutoSize        =   -1  'True
                  Caption         =   "Elegido"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H8000000D&
                  Height          =   195
                  Left            =   2640
                  TabIndex        =   97
                  Top             =   165
                  Width           =   645
               End
            End
            Begin VB.Frame Frame7 
               Caption         =   "Tamaño"
               ForeColor       =   &H8000000D&
               Height          =   855
               Left            =   120
               TabIndex        =   76
               Top             =   2040
               Width           =   6855
               Begin VB.TextBox Text3 
                  Height          =   315
                  Left            =   5160
                  TabIndex        =   81
                  Top             =   450
                  Width           =   1575
               End
               Begin VB.CommandButton Command13 
                  BackColor       =   &H8000000E&
                  Caption         =   "OTRO"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   495
                  Left            =   3840
                  Style           =   1  'Graphical
                  TabIndex        =   80
                  Top             =   240
                  Width           =   1215
               End
               Begin VB.CommandButton Command12 
                  BackColor       =   &H8000000E&
                  Caption         =   "15 x 21"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   495
                  Left            =   2640
                  Style           =   1  'Graphical
                  TabIndex        =   79
                  Top             =   240
                  Width           =   1215
               End
               Begin VB.CommandButton Command11 
                  BackColor       =   &H8000000E&
                  Caption         =   "13 x 18"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   495
                  Left            =   1440
                  Style           =   1  'Graphical
                  TabIndex        =   78
                  Top             =   240
                  Width           =   1215
               End
               Begin VB.CommandButton Command7 
                  BackColor       =   &H8000000E&
                  Caption         =   "10 x 15"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   495
                  Left            =   120
                  MaskColor       =   &H00FFFFFF&
                  Style           =   1  'Graphical
                  TabIndex        =   77
                  Top             =   240
                  Width           =   1335
               End
               Begin VB.Label Label8 
                  AutoSize        =   -1  'True
                  Caption         =   "Tamaño Elegido"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H8000000D&
                  Height          =   195
                  Left            =   5160
                  TabIndex        =   82
                  Top             =   240
                  Width           =   1380
               End
            End
            Begin VB.Frame Frame6 
               Caption         =   "Trabajo a Realizar"
               ForeColor       =   &H8000000D&
               Height          =   855
               Left            =   120
               TabIndex        =   70
               Top             =   1200
               Width           =   6855
               Begin VB.CommandButton Command10 
                  BackColor       =   &H8000000E&
                  Caption         =   "Rev. y Copia"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   495
                  Left            =   120
                  MaskColor       =   &H00FFFFFF&
                  Style           =   1  'Graphical
                  TabIndex        =   74
                  Top             =   240
                  Width           =   1600
               End
               Begin VB.CommandButton Command9 
                  BackColor       =   &H8000000E&
                  Caption         =   "ReImpresión"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   495
                  Left            =   1800
                  Style           =   1  'Graphical
                  TabIndex        =   73
                  Top             =   240
                  Width           =   1600
               End
               Begin VB.CommandButton Command6 
                  BackColor       =   &H8000000E&
                  Caption         =   "OTRO"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   495
                  Left            =   3450
                  Style           =   1  'Graphical
                  TabIndex        =   72
                  Top             =   240
                  Width           =   1600
               End
               Begin VB.TextBox Text2 
                  Height          =   315
                  Left            =   5160
                  TabIndex        =   71
                  Top             =   450
                  Width           =   1575
               End
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  Caption         =   "Trabajo Elegido"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H8000000D&
                  Height          =   195
                  Left            =   5160
                  TabIndex        =   75
                  Top             =   240
                  Width           =   1350
               End
            End
            Begin VB.Frame Frame5 
               Caption         =   "Exposición"
               ForeColor       =   &H8000000D&
               Height          =   855
               Left            =   7080
               TabIndex        =   69
               Top             =   1200
               Width           =   3495
               Begin VB.OptionButton Option3 
                  Caption         =   "36 Fotos"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   2280
                  TabIndex        =   92
                  Top             =   360
                  Width           =   1095
               End
               Begin VB.OptionButton Option2 
                  Caption         =   "24 Fotos"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   1200
                  TabIndex        =   91
                  Top             =   360
                  Width           =   1095
               End
               Begin VB.OptionButton Option1 
                  Caption         =   "12 Fotos"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   90
                  Top             =   360
                  Value           =   -1  'True
                  Width           =   1095
               End
            End
            Begin VB.Frame Frame3 
               Caption         =   "Marca"
               ForeColor       =   &H8000000D&
               Height          =   975
               Left            =   120
               TabIndex        =   60
               Top             =   240
               Width           =   6855
               Begin VB.TextBox Text1 
                  Height          =   375
                  Left            =   5160
                  TabIndex        =   66
                  Top             =   450
                  Width           =   1575
               End
               Begin VB.CommandButton Command5 
                  BackColor       =   &H8000000E&
                  Caption         =   "OTRO"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   615
                  Left            =   4080
                  Style           =   1  'Graphical
                  TabIndex        =   65
                  Top             =   240
                  Width           =   975
               End
               Begin VB.CommandButton Command4 
                  BackColor       =   &H8000000E&
                  Height          =   615
                  Left            =   3120
                  Picture         =   "frmRevelados.frx":0070
                  Style           =   1  'Graphical
                  TabIndex        =   64
                  Top             =   240
                  Width           =   975
               End
               Begin VB.CommandButton Command3 
                  BackColor       =   &H8000000E&
                  Height          =   615
                  Left            =   2160
                  Picture         =   "frmRevelados.frx":261A
                  Style           =   1  'Graphical
                  TabIndex        =   63
                  Top             =   240
                  Width           =   975
               End
               Begin VB.CommandButton Command2 
                  BackColor       =   &H8000000E&
                  Height          =   615
                  Left            =   1200
                  Picture         =   "frmRevelados.frx":3160
                  Style           =   1  'Graphical
                  TabIndex        =   62
                  Top             =   240
                  Width           =   975
               End
               Begin VB.CommandButton Command1 
                  BackColor       =   &H8000000E&
                  Height          =   615
                  Left            =   120
                  MaskColor       =   &H00FFFFFF&
                  Picture         =   "frmRevelados.frx":410E
                  Style           =   1  'Graphical
                  TabIndex        =   61
                  Top             =   240
                  Width           =   1095
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Marca Elegida"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H8000000D&
                  Height          =   195
                  Left            =   5160
                  TabIndex        =   67
                  Top             =   240
                  Width           =   1230
               End
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "Observaciones:"
               ForeColor       =   &H8000000D&
               Height          =   195
               Left            =   240
               TabIndex        =   98
               Top             =   3000
               Width           =   1110
            End
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
            Height          =   4290
            Left            =   -74760
            TabIndex        =   58
            Top             =   600
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   7567
            _Version        =   393216
            Cols            =   13
            FixedCols       =   0
            BackColorSel    =   8388736
            AllowBigSelection=   -1  'True
            FocusRect       =   0
            HighLight       =   2
            SelectionMode   =   1
         End
      End
      Begin VB.Frame frameBuscar 
         Caption         =   "Buscar por..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1950
         Left            =   -74600
         TabIndex        =   28
         Top             =   540
         Width           =   10410
         Begin VB.CheckBox chkCliente 
            Caption         =   "Cliente"
            Height          =   195
            Left            =   878
            TabIndex        =   45
            Top             =   315
            Width           =   855
         End
         Begin VB.CheckBox chkFecha 
            Caption         =   "Fecha"
            Height          =   195
            Left            =   900
            TabIndex        =   44
            Top             =   975
            Width           =   810
         End
         Begin VB.TextBox txtCliente 
            Height          =   300
            Left            =   3240
            MaxLength       =   40
            TabIndex        =   43
            Top             =   255
            Width           =   975
         End
         Begin VB.TextBox txtDesCli 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   4725
            MaxLength       =   50
            TabIndex        =   42
            Tag             =   "Descripción"
            Top             =   255
            Width           =   4620
         End
         Begin VB.CommandButton CmdBuscAprox 
            Height          =   1395
            Left            =   9660
            MaskColor       =   &H000000FF&
            Picture         =   "frmRevelados.frx":510C
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Buscar "
            Top             =   225
            UseMaskColor    =   -1  'True
            Width           =   555
         End
         Begin VB.CheckBox chkVendedor 
            Caption         =   "Empleado"
            Height          =   195
            Left            =   900
            TabIndex        =   38
            Top             =   645
            Width           =   1035
         End
         Begin VB.TextBox txtDesVen 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4725
            TabIndex        =   37
            Top             =   675
            Width           =   4620
         End
         Begin VB.TextBox txtVendedor 
            Height          =   300
            Left            =   3240
            TabIndex        =   36
            Top             =   667
            Width           =   990
         End
         Begin VB.CommandButton cmdBuscarCli 
            Height          =   315
            Left            =   4290
            MaskColor       =   &H000000FF&
            Picture         =   "frmRevelados.frx":78AE
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Buscar Cliente"
            Top             =   255
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdBuscarVendedor 
            Height          =   315
            Left            =   4290
            MaskColor       =   &H000000FF&
            Picture         =   "frmRevelados.frx":7BB8
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Buscar Vendedor"
            Top             =   660
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.Frame Frame1 
            Caption         =   "Estado Revelado"
            Height          =   495
            Left            =   840
            TabIndex        =   29
            Top             =   1320
            Width           =   8535
            Begin VB.OptionButton optPen 
               Caption         =   "Pendientes"
               Height          =   195
               Left            =   1200
               TabIndex        =   33
               Top             =   200
               Value           =   -1  'True
               Width           =   1575
            End
            Begin VB.OptionButton optDef 
               Caption         =   "Definitivos"
               Height          =   195
               Left            =   3075
               TabIndex        =   32
               Top             =   200
               Width           =   1455
            End
            Begin VB.OptionButton optAnu 
               Caption         =   "Anulados"
               Height          =   195
               Left            =   4845
               TabIndex        =   31
               Top             =   200
               Width           =   1455
            End
            Begin VB.OptionButton optTod 
               Caption         =   "Todos"
               Height          =   195
               Left            =   6600
               TabIndex        =   30
               Top             =   200
               Width           =   1455
            End
         End
         Begin FechaCtl.Fecha FechaHasta 
            Height          =   285
            Left            =   5745
            TabIndex        =   40
            Top             =   1080
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin FechaCtl.Fecha FechaDesde 
            Height          =   330
            Left            =   3240
            TabIndex        =   41
            Top             =   1080
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
            Left            =   2625
            TabIndex        =   49
            Top             =   300
            Width           =   525
         End
         Begin VB.Label lblFechaDesde 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde:"
            Height          =   195
            Left            =   2145
            TabIndex        =   48
            Top             =   1125
            Width           =   1005
         End
         Begin VB.Label lblFechaHasta 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Left            =   4695
            TabIndex        =   47
            Top             =   1140
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Empleado:"
            Height          =   195
            Index           =   0
            Left            =   2415
            TabIndex        =   46
            Top             =   705
            Width           =   750
         End
      End
      Begin VB.Frame freRemito 
         Caption         =   "Revelado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1710
         Left            =   8160
         TabIndex        =   20
         Top             =   360
         Width           =   2835
         Begin VB.TextBox txtNroSucursal 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   840
            MaxLength       =   4
            TabIndex        =   22
            Top             =   480
            Width           =   555
         End
         Begin VB.TextBox txtNroRemito 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1410
            MaxLength       =   8
            TabIndex        =   21
            Top             =   480
            Width           =   1005
         End
         Begin FechaCtl.Fecha FechaRemito 
            Height          =   285
            Left            =   840
            TabIndex        =   23
            Top             =   945
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Left            =   240
            TabIndex        =   27
            Top             =   1395
            Width           =   540
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            Height          =   195
            Left            =   180
            TabIndex        =   26
            Top             =   525
            Width           =   600
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   285
            TabIndex        =   25
            Top             =   990
            Width           =   495
         End
         Begin VB.Label lblEstadoRemito 
            AutoSize        =   -1  'True
            Caption         =   "EST. Revelado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   840
            TabIndex        =   24
            Top             =   1410
            Width           =   1785
         End
      End
      Begin VB.Frame freCliente 
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1710
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   7845
         Begin VB.TextBox txtcodpos 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   930
            TabIndex        =   56
            Top             =   935
            Width           =   975
         End
         Begin VB.TextBox txtlocalidad 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1995
            MaxLength       =   50
            TabIndex        =   55
            Top             =   935
            Width           =   5700
         End
         Begin VB.TextBox Text7 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6480
            TabIndex        =   54
            Top             =   595
            Width           =   1215
         End
         Begin VB.TextBox txtCUIT 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   930
            TabIndex        =   15
            Top             =   1275
            Width           =   1455
         End
         Begin VB.CommandButton cmdNuevoCliente 
            Height          =   315
            Left            =   2625
            MaskColor       =   &H000000FF&
            Picture         =   "frmRevelados.frx":7EC2
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Agregar Cliente"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdBuscarCliente 
            Height          =   315
            Left            =   2160
            MaskColor       =   &H000000FF&
            Picture         =   "frmRevelados.frx":824C
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Buscar Cliente"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox txtCondicionIVA 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2415
            TabIndex        =   12
            Top             =   1275
            Width           =   3135
         End
         Begin VB.TextBox TxtCodigoCli 
            Height          =   300
            Left            =   930
            MaxLength       =   40
            TabIndex        =   11
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtRazSocCli 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3105
            MaxLength       =   50
            TabIndex        =   10
            Tag             =   "Descripción"
            Top             =   240
            Width           =   4590
         End
         Begin VB.TextBox txtDomici 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   930
            MaxLength       =   50
            TabIndex        =   9
            Top             =   595
            Width           =   4620
         End
         Begin VB.TextBox txtIngBrutos 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6480
            TabIndex        =   8
            Top             =   1275
            Width           =   1215
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Localidad:"
            Height          =   195
            Left            =   120
            TabIndex        =   57
            Top             =   975
            Width           =   735
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Teléfono:"
            Height          =   195
            Left            =   5760
            TabIndex        =   53
            Top             =   640
            Width           =   675
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "DNI:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   450
            TabIndex        =   19
            Top             =   285
            Width           =   330
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio:"
            Height          =   195
            Left            =   180
            TabIndex        =   18
            Top             =   630
            Width           =   675
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "C.U.I.T.:"
            Height          =   195
            Left            =   255
            TabIndex        =   17
            Top             =   1320
            Width           =   600
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Ing. Brutos:"
            Height          =   195
            Left            =   5625
            TabIndex        =   16
            Top             =   1320
            Width           =   810
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   4530
         Left            =   -74640
         TabIndex        =   50
         Top             =   2595
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   7990
         _Version        =   393216
         Cols            =   13
         FixedCols       =   0
         BackColorSel    =   8388736
         AllowBigSelection=   -1  'True
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ordenado por :"
         Height          =   195
         Index           =   0
         Left            =   -74820
         TabIndex        =   51
         Top             =   570
         Width           =   1065
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "<F1> Buscar Revelado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   3600
      TabIndex        =   5
      Top             =   7440
      Width           =   2595
   End
   Begin VB.Label lblEstado 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   150
      TabIndex        =   4
      Top             =   7455
      Width           =   750
   End
End
Attribute VB_Name = "frmRevelados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Centrar_pantalla Me
End Sub

