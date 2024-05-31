VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form FrmCreditos 
   Caption         =   "Créditos Hard Motor's"
   ClientHeight    =   9105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11730
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9105
   ScaleWidth      =   11730
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtdias 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   10005
      MaxLength       =   2
      TabIndex        =   83
      Top             =   1275
      Width           =   705
   End
   Begin VB.Frame Frame6 
      Caption         =   "Datos Vendedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   840
      Left            =   585
      TabIndex        =   74
      Top             =   1605
      Width           =   8415
      Begin VB.TextBox i_codven 
         Height          =   315
         Left            =   945
         TabIndex        =   1
         Top             =   390
         Width           =   1185
      End
      Begin VB.Label lblnom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vendedor :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   24
         Left            =   60
         TabIndex        =   76
         Tag             =   "4"
         Top             =   420
         Width           =   795
      End
      Begin VB.Label lblnomven 
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2220
         TabIndex        =   75
         Top             =   390
         Width           =   6045
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Crédito de Producto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   6570
      Left            =   30
      TabIndex        =   17
      Top             =   2475
      Width           =   11640
      Begin MSComctlLib.ListView LV_VEN 
         Height          =   285
         Left            =   4695
         TabIndex        =   78
         Top             =   930
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         TextBackground  =   -1  'True
         _Version        =   393217
         ForeColor       =   8388608
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.CommandButton SALIR 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ce&rrar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   10470
         Picture         =   "FrmCreditos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   77
         TabStop         =   0   'False
         Tag             =   "9999"
         Top             =   5610
         Width           =   1140
      End
      Begin VB.CommandButton grabar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Grabar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   10470
         Picture         =   "FrmCreditos.frx":0876
         Style           =   1  'Graphical
         TabIndex        =   73
         TabStop         =   0   'False
         Tag             =   "9999"
         Top             =   3615
         Width           =   1110
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   540
         Left            =   3075
         TabIndex        =   70
         Top             =   930
         Visible         =   0   'False
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   953
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         TextBackground  =   -1  'True
         _Version        =   393217
         ForeColor       =   8388608
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.TextBox txt_key 
         DataField       =   "ART_KEY"
         DataSource      =   "Data1"
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
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   2
         Top             =   510
         Width           =   1260
      End
      Begin MSComctlLib.ListView LV_CLI 
         Height          =   270
         Left            =   2160
         TabIndex        =   63
         Top             =   975
         Visible         =   0   'False
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   476
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         TextBackground  =   -1  'True
         _Version        =   393217
         ForeColor       =   8388608
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.TextBox txtPeriodos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1785
         MaxLength       =   2
         TabIndex        =   5
         Text            =   "10"
         Top             =   5070
         Width           =   570
      End
      Begin VB.Frame Frame5 
         Caption         =   "Lista   ( $ )"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1995
         Left            =   3675
         TabIndex        =   31
         Top             =   1110
         Width           =   2055
         Begin VB.Image Image1 
            Height          =   210
            Left            =   240
            Picture         =   "FrmCreditos.frx":1710
            Top             =   1560
            Width           =   1665
         End
         Begin VB.Label lblpvtaLista 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "2990.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   360
            Left            =   330
            TabIndex        =   35
            Top             =   585
            Width           =   1260
         End
         Begin VB.Label Label6 
            Caption         =   "P.V.P. CASH"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   450
            TabIndex        =   32
            Top             =   330
            Width           =   1140
         End
      End
      Begin VB.TextBox txtTipoCambio 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   480
         Left            =   10665
         MaxLength       =   4
         TabIndex        =   3
         Text            =   "3.29"
         Top             =   420
         Width           =   900
      End
      Begin VB.Frame Frame4 
         Caption         =   "Soles   ( S/. )"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   2010
         Left            =   5940
         TabIndex        =   19
         Top             =   1095
         Width           =   2055
         Begin VB.Label lblInicialSol 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1974.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   315
            Left            =   285
            TabIndex        =   36
            Top             =   1485
            Width           =   1260
         End
         Begin VB.Label lblpvtaSol 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "10328.96"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   360
            Left            =   285
            TabIndex        =   34
            Top             =   540
            Width           =   1260
         End
         Begin VB.Label lblIMSol 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1859.21"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   315
            Left            =   285
            TabIndex        =   33
            Top             =   1035
            Width           =   1260
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Dólares ( $ )"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   1995
         Left            =   1455
         TabIndex        =   18
         Top             =   1110
         Width           =   2055
         Begin VB.TextBox txtInicialDol 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   360
            Left            =   345
            MaxLength       =   4
            TabIndex        =   4
            Text            =   "600"
            Top             =   1440
            Width           =   1260
         End
         Begin VB.Label Label5 
            Caption         =   "P.V.P. FINANCIADO"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   180
            TabIndex        =   30
            Top             =   345
            Width           =   1785
         End
         Begin VB.Label lblIMDol 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "3138.66"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   315
            Left            =   360
            TabIndex        =   22
            Top             =   1065
            Width           =   1260
         End
         Begin VB.Label lblpvtaDol 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "3139.50"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   360
            Left            =   360
            TabIndex        =   21
            Top             =   600
            Width           =   1260
         End
      End
      Begin VB.CommandButton cmdimp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Imprimir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   10485
         Picture         =   "FrmCreditos.frx":1904
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   4605
         Width           =   1110
      End
      Begin VB.CommandButton cancelar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Limpiar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   10470
         Picture         =   "FrmCreditos.frx":2906
         Style           =   1  'Graphical
         TabIndex        =   72
         TabStop         =   0   'False
         Tag             =   "9999"
         Top             =   2640
         Width           =   1110
      End
      Begin VB.Label lblIM 
         BackStyle       =   0  'Transparent
         Caption         =   "18"
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
         Height          =   240
         Left            =   645
         TabIndex        =   65
         ToolTipText     =   "Doble Click para cambiar %"
         Top             =   2325
         Width           =   315
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Producto Seleccionado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   3195
         TabIndex        =   69
         Top             =   315
         Width           =   1920
      End
      Begin VB.Label lblProd 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3120
         TabIndex        =   68
         Top             =   495
         Width           =   5910
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Cod. Interno"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   20
         Left            =   1860
         TabIndex        =   67
         Top             =   315
         Width           =   1050
      End
      Begin VB.Label lblnom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "% :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   21
         Left            =   945
         TabIndex        =   66
         Tag             =   "4"
         Top             =   2295
         Width           =   345
      End
      Begin VB.Label Label3 
         Caption         =   "Busca Producto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   180
         Left            =   120
         TabIndex        =   64
         Top             =   360
         Width           =   1380
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   45
         Picture         =   "FrmCreditos.frx":30B4
         Top             =   525
         Width           =   1710
      End
      Begin VB.Label lblCuotaSol 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1171.37"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   345
         Left            =   6210
         TabIndex        =   62
         Top             =   5985
         Width           =   1290
      End
      Begin VB.Label lblCuotaDol 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "356.04"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   345
         Left            =   1785
         TabIndex        =   61
         Top             =   5985
         Width           =   1290
      End
      Begin VB.Label lblnom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "CUOTA :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   20
         Left            =   585
         TabIndex        =   60
         Tag             =   "4"
         Top             =   6030
         Width           =   765
      End
      Begin VB.Label Label8 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2415
         TabIndex        =   59
         Top             =   5595
         Width           =   255
      End
      Begin VB.Label lblTIM 
         Caption         =   "2.95"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1815
         TabIndex        =   58
         ToolTipText     =   "Doble Click para cambiar %"
         Top             =   5565
         Width           =   555
      End
      Begin VB.Label lblnom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "T.I.M. :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   19
         Left            =   675
         TabIndex        =   57
         Tag             =   "4"
         Top             =   5565
         Width           =   660
      End
      Begin VB.Label lblMAF1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3044.86"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   315
         Left            =   8220
         TabIndex        =   56
         Top             =   5010
         Width           =   1290
      End
      Begin VB.Label lblnom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "MAF :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   18
         Left            =   7275
         TabIndex        =   55
         Tag             =   "4"
         Top             =   5040
         Width           =   540
      End
      Begin VB.Label lblnom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Meses"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   17
         Left            =   2475
         TabIndex        =   54
         Tag             =   "4"
         Top             =   5085
         Width           =   615
      End
      Begin VB.Label lblnom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Periodo :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   16
         Left            =   465
         TabIndex        =   53
         Tag             =   "4"
         Top             =   5070
         Width           =   870
      End
      Begin VB.Label lblLocador 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "251.51"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   315
         Left            =   8220
         TabIndex        =   52
         Top             =   4125
         Width           =   1260
      End
      Begin VB.Label lblFGarantia 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "253.94"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   315
         Left            =   8205
         TabIndex        =   51
         Top             =   3705
         Width           =   1260
      End
      Begin VB.Image Image2 
         Height          =   540
         Left            =   8475
         Picture         =   "FrmCreditos.frx":32B6
         Top             =   2820
         Width           =   1620
      End
      Begin VB.Label lblAbonoHM 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2539.41"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   390
         Left            =   8190
         TabIndex        =   50
         Top             =   3255
         Width           =   1260
      End
      Begin VB.Label lblnom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "LOCADOR ACC. :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   15
         Left            =   6240
         TabIndex        =   48
         Tag             =   "4"
         ToolTipText     =   "Doble Click para cambiar %"
         Top             =   4140
         Width           =   1560
      End
      Begin VB.Label lblnom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "F. GARANTÍA :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   14
         Left            =   6435
         TabIndex        =   47
         Tag             =   "4"
         ToolTipText     =   "Doble Click para cambiar %"
         Top             =   3720
         Width           =   1350
      End
      Begin VB.Label lblCOM 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "505.36"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   315
         Left            =   4260
         TabIndex        =   46
         Top             =   4275
         Width           =   1260
      End
      Begin VB.Label lblDIF 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2539.50"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   315
         Left            =   4245
         TabIndex        =   45
         Top             =   3855
         Width           =   1260
      End
      Begin VB.Label lblnom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "COM :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   13
         Left            =   3630
         TabIndex        =   44
         Tag             =   "4"
         Top             =   4305
         Width           =   555
      End
      Begin VB.Label lblnom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "DIF :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   12
         Left            =   3735
         TabIndex        =   43
         Tag             =   "4"
         Top             =   3900
         Width           =   435
      End
      Begin VB.Label lblBoleta 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3644.86"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   345
         Left            =   1785
         TabIndex        =   42
         Top             =   4545
         Width           =   1290
      End
      Begin VB.Label lblMAF 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3044.86"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   315
         Left            =   1785
         TabIndex        =   41
         Top             =   4125
         Width           =   1260
      End
      Begin VB.Label lblGA 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "203.16"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   315
         Left            =   1785
         TabIndex        =   40
         Top             =   3720
         Width           =   1260
      End
      Begin VB.Label lblGG 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "302.20"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   315
         Left            =   1785
         TabIndex        =   39
         Top             =   3315
         Width           =   1260
      End
      Begin VB.Label Label2 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9405
         TabIndex        =   38
         Top             =   1965
         Width           =   465
      End
      Begin VB.Label lblPorPVta 
         Alignment       =   2  'Center
         Caption         =   "19.11"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   8220
         TabIndex        =   37
         Top             =   1950
         Width           =   1110
      End
      Begin VB.Label lblnom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Boleta :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   11
         Left            =   570
         TabIndex        =   29
         Tag             =   "4"
         Top             =   4710
         Width           =   750
      End
      Begin VB.Label lblnom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "MAF :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   10
         Left            =   765
         TabIndex        =   28
         Tag             =   "4"
         Top             =   4125
         Width           =   540
      End
      Begin VB.Label lblnom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "G. A. :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   9
         Left            =   705
         TabIndex        =   27
         Tag             =   "4"
         ToolTipText     =   "Doble Click para cambiar %"
         Top             =   3735
         Width           =   585
      End
      Begin VB.Label lblnom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "G. G. :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   7
         Left            =   720
         TabIndex        =   26
         Tag             =   "4"
         ToolTipText     =   "Doble Click para cambiar %"
         Top             =   3330
         Width           =   555
      End
      Begin VB.Label lblnom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Inicial :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   6
         Left            =   585
         TabIndex        =   25
         Tag             =   "4"
         Top             =   2715
         Width           =   690
      End
      Begin VB.Label lblnom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "I.M."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   5
         Left            =   240
         TabIndex        =   24
         Tag             =   "4"
         Top             =   2295
         Width           =   360
      End
      Begin VB.Label lblnom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "P. Venta :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   4
         Left            =   315
         TabIndex        =   23
         Tag             =   "4"
         Top             =   1875
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Cambio :"
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
         Index           =   0
         Left            =   9120
         TabIndex        =   20
         Top             =   540
         Width           =   1500
      End
      Begin VB.Label Label4 
         BackColor       =   &H000000C0&
         Caption         =   "HARDCORP :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   6090
         TabIndex        =   49
         ToolTipText     =   "Doble Click para cambiar %"
         Top             =   3255
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1575
      Left            =   600
      TabIndex        =   6
      Top             =   15
      Width           =   8370
      Begin VB.TextBox i_codcli 
         Height          =   315
         Left            =   945
         TabIndex        =   0
         Top             =   375
         Width           =   1185
      End
      Begin VB.Label lbldire 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   930
         TabIndex        =   16
         Top             =   1155
         Width           =   7335
      End
      Begin VB.Label lblnom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Dirección :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   105
         TabIndex        =   15
         Tag             =   "4"
         Top             =   1200
         Width           =   750
      End
      Begin VB.Label lbltelefono 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6765
         TabIndex        =   14
         Top             =   780
         Width           =   1500
      End
      Begin VB.Label lblruc 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3585
         TabIndex        =   13
         Top             =   780
         Width           =   1635
      End
      Begin VB.Label lblnom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "RUC :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   3090
         TabIndex        =   12
         Tag             =   "4"
         Top             =   840
         Width           =   420
      End
      Begin VB.Label lbldni 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
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
         TabIndex        =   11
         Top             =   780
         Width           =   1290
      End
      Begin VB.Label lblCli 
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2220
         TabIndex        =   10
         Top             =   390
         Width           =   6045
      End
      Begin VB.Label lblnom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cliente :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   255
         TabIndex        =   9
         Tag             =   "4"
         Top             =   420
         Width           =   600
      End
      Begin VB.Label lblnom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "DNI :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   480
         TabIndex        =   8
         Tag             =   "4"
         Top             =   810
         Width           =   375
      End
      Begin VB.Label lblnom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Teléfono :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   8
         Left            =   5910
         TabIndex        =   7
         Top             =   825
         Width           =   735
      End
   End
   Begin MSMask.MaskEdBox i_fecha 
      Height          =   285
      Left            =   9780
      TabIndex        =   80
      Tag             =   "36"
      Top             =   435
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   14737632
      ForeColor       =   128
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox i_fechaDias 
      Height          =   285
      Left            =   9825
      TabIndex        =   81
      Tag             =   "36"
      Top             =   2025
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   14737632
      ForeColor       =   128
      Enabled         =   0   'False
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "Fecha Inicio Pago"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   210
      Left            =   9570
      TabIndex        =   84
      Top             =   1725
      Width           =   1695
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Dias de Gracia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   180
      Left            =   9660
      TabIndex        =   82
      Top             =   960
      Width           =   1380
   End
   Begin VB.Label Fecha 
      Alignment       =   2  'Center
      Caption         =   "Fecha Activación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   180
      Left            =   9525
      TabIndex        =   79
      Top             =   150
      Width           =   1605
   End
End
Attribute VB_Name = "FrmCreditos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPorGG As String
Dim vPorGA As String
Dim vPorHM As String
Dim vPorFGaran As String
Dim vPorLocador As String
Dim loc_key As Integer
Dim cnt As New ADODB.Connection

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub Limpiar()
    lblcli.Caption = ""
    lbldni.Caption = ""
    lblruc.Caption = ""
    lbltelefono.Caption = ""
    lbldire.Caption = ""
End Sub
Private Sub LimpiarValores()
    lblProd.Caption = ""
    lblpvtaDol.Caption = "0.00"
    lblpvtaLista.Caption = "0.00"
    lblpvtaSol.Caption = "0.00"
    lblIMDol.Caption = "0.00"
    lblIMSol.Caption = "0.00"
    txtInicialDol.Text = "0"
    lblInicialSol.Caption = "0"
    lblPorPVta.Caption = "0.00"
    lblGG.Caption = "0.00"
    lblGA.Caption = "0.00"
    lblMAF.Caption = "0.00"
    lblDIF.Caption = "0.00"
    lblCOM.Caption = "0.00"
    lblAbonoHM.Caption = "0.00"
    lblFGarantia.Caption = "0.00"
    lblLocador.Caption = "0.00"
    lblMAF1.Caption = "0.00"
    lblBoleta.Caption = "0.00"
    txtPeriodos.Text = "0"
    lblCuotaDol.Caption = "0.00"
    lblCuotaSol.Caption = "0.00"
End Sub

Private Sub cancelar_Click()
    txt_key.Text = ""
    i_codcli.Text = ""
    i_codven.Text = ""
    txtdias.Text = ""
    i_fecha.Text = Format(LK_FECHA_DIA, "dd/mm/yy")
    txtdias.Text = "0"
    i_fechaDias.Text = Format(i_fechaDias.Text, "__/__/__")
    Azul i_codcli, i_codcli
End Sub

Private Sub cmdimp_Click()
    'If MsgBox("Está seguro de Imprimir Cotización", vbInformation + vbYesNo, Pub_Titulo) = vbYes Then
        PrintForm
    'End If
End Sub

Private Sub Form_Load()
On Error GoTo LERROR
Dim strCnn As String
   strCnn = "provider=SQLOLEDB;Data source=(Local);initial catalog=bdatos;password=anteromariano;user id=sa"
         
    cnt.ConnectionTimeout = 15
    cnt.Open strCnn
    
    PUB_CP = "C"
    Call LlenaConstantes
    Call Limpiar
    Call LimpiarValores
    txtdias.Text = "0"
    i_fecha.Text = Format(LK_FECHA_DIA, "dd/mm/yy")
    Exit Sub
LERROR:
    MsgBox Err.Description, vbCritical, Pub_Titulo
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cnt.Close
    Set cnt = Nothing
End Sub

Private Sub grabar_Click()
    If Len(Trim(i_codcli.Text)) = 0 Then
        MsgBox "Ingrese Cliente...", vbExclamation, Pub_Titulo
        Azul i_codcli, i_codcli
        Exit Sub
    End If
    If Not IsDate(i_fecha.Text) Then
        MsgBox "Fecha Activación Incorrecta...", vbExclamation, Pub_Titulo
        i_fecha.SetFocus
        i_fecha.SelStart = 0
        i_fecha.SelLength = 8
        Exit Sub
    End If
     If Not IsDate(i_fechaDias.Text) Then
        MsgBox "Fecha Inicio Pago no ha sido Establecida...", vbExclamation, Pub_Titulo
        Azul txtdias, txtdias
        Exit Sub
    End If
    If Len(Trim(i_codven.Text)) = 0 Then
        MsgBox "Ingrese Vendedor...", vbExclamation, Pub_Titulo
        Azul i_codven, i_codven
        Exit Sub
    End If
    If Val(lblCuotaDol.Caption) = 0 Then
        MsgBox "Cotización no Procede...Realice los Cálculos..", vbExclamation, Pub_Titulo
        Azul txt_key, txt_key
        Exit Sub
    End If
    If MsgBox("Está seguro de Grabar Cotización", vbInformation + vbYesNo, Pub_Titulo) = vbYes Then
        Call GrabaCotizacion
    End If
End Sub

Private Sub i_codcli_Change()
    If i_codcli.Text = "" Then
        LV_CLI.Visible = False
        Call Limpiar
    End If
End Sub

Private Sub i_codcli_GotFocus()
    loc_key = 0
End Sub

Private Sub i_codcli_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As Object    ' Variable FoundItem.
If Not LV_CLI.Visible Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And i_codcli.Text = "" Then
  loc_key = 1
  Set LV_CLI.SelectedItem = LV_CLI.ListItems(loc_key)
  LV_CLI.ListItems.item(loc_key).Selected = True
  LV_CLI.ListItems.item(loc_key).EnsureVisible
  GoTo fin
End If

If KeyCode = 40 Then  ' flecha abajo
  loc_key = loc_key + 1
  If loc_key > LV_CLI.ListItems.count Then loc_key = LV_CLI.ListItems.count
  GoTo POSICION
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo POSICION
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > LV_CLI.ListItems.count Then loc_key = LV_CLI.ListItems.count
 GoTo POSICION
End If
If KeyCode = 33 Then
 loc_key = loc_key - 17
 If loc_key < 1 Then loc_key = 1
 GoTo POSICION
End If
GoTo fin
POSICION:
  LV_CLI.ListItems.item(loc_key).Selected = True
  LV_CLI.ListItems.item(loc_key).EnsureVisible
  i_codcli.Text = Trim(LV_CLI.ListItems.item(loc_key).Text) & " "
  i_codcli.SelStart = Len(i_codcli.Text)
fin:
End Sub

Private Sub i_codcli_KeyPress(KeyAscii As Integer)
Dim VAR  As String
Dim valor As String
Dim tf As Integer
Dim wcta1 As String * 12
Dim wcta2 As String * 12
Dim WDIAS As Integer
Dim wCONDI As Integer
Dim wdia1 As String
Dim wDIA2 As String
Dim i
Dim itmFound As Object    ' Variable FoundItem.
wCONDI = -1
If KeyAscii = 27 Then
  i_codcli.Text = ""
 Exit Sub
End If
If KeyAscii <> 13 Then
   GoTo fin
End If
If KeyAscii = 13 And Left(i_codcli.Text, 1) = "+" Then GoTo buscar
On Error GoTo OJO
pu_cp = PUB_CP
pu_codclie = Val(i_codcli.Text)
On Error GoTo 0
If Len(i_codcli.Text) = 0 Then
   Exit Sub
End If
If pu_codclie <> 0 And IsNumeric(i_codcli.Text) = True Then
   If Len(Trim(i_codcli.Text)) = LK_DIG_RUC Then ' LONG DEL RUC
        PUB_RUC = Trim(i_codcli.Text)
        SQ_OPER = 4
        pu_codcia = LK_CODCIA
        LEER_CLI_LLAVE
        If cli_ruc.EOF Then
           MsgBox "R.U.C. No Existe ", 48, Pub_Titulo
           Exit Sub
        End If
        i_codcli.Text = cli_ruc!cli_codclie
   End If
   On Error GoTo OJO
   SQ_OPER = 1
   pu_codclie = Val(i_codcli.Text)
   pu_codcia = LK_CODCIA
   pu_cp = PUB_CP
   LEER_CLI_LLAVE
   On Error GoTo 0
   wcta1 = ""
   wcta2 = ""
   If cli_llave.EOF Then
    Azul i_codcli, i_codcli
    MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
    GoTo fin
   Else
      If Trim(cli_llave("cli_estado")) = "" Then 'And LK_CODTRA = 2401 Then 'mic
            Azul i_codcli, i_codcli
            MsgBox "CLIENTE DESACTIVO ...", 48, Pub_Titulo
            i_codcli.SetFocus
            GoTo fin
      End If
      lblcli.Caption = Trim(cli_llave!CLI_NOMBRE)
      lbldni.Caption = Trim(Nulo_Valors(cli_llave!cli_RUC_ESPOSA))
      lblruc.Caption = Trim(Nulo_Valors(cli_llave!cli_ruc_esposo))
      lbltelefono.Caption = Trim(Nulo_Valors(cli_llave!CLI_TELEF1))
      lbldire.Caption = Trim(Nulo_Valors(cli_llave!CLI_CASA_DIREC))
      i_fecha.SetFocus
      i_fecha.SelStart = 0
      i_fecha.SelLength = 8
   End If
   
Else
On Error GoTo sigue
   If loc_key <> 0 Then valor = UCase(LV_CLI.ListItems.item(loc_key).Text)
   If Trim(UCase(i_codcli.Text)) = Left(valor, Len(Trim(i_codcli.Text))) Then
   Else
      Exit Sub
   End If
   If loc_key = 0 Then Exit Sub
   i_codcli.Text = Trim(LV_CLI.ListItems.item(loc_key).SubItems(1))
   pu_codclie = Val(i_codcli.Text)
   pu_cp = PUB_CP
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   LEER_CLI_LLAVE
   lblcli.Caption = Trim(cli_llave!CLI_NOMBRE)
   lbldni.Caption = Trim(Nulo_Valors(cli_llave!cli_RUC_ESPOSA))
   lblruc.Caption = Trim(Nulo_Valors(cli_llave!cli_ruc_esposo))
   lbltelefono.Caption = Trim(Nulo_Valors(cli_llave!CLI_TELEF1))
   lbldire.Caption = Trim(Nulo_Valors(cli_llave!CLI_CASA_DIREC))
   i_fecha.SetFocus
   i_fecha.SelStart = 0
   i_fecha.SelLength = 8
End If
LV_CLI.Visible = False
'If Nulo_Valor0(SUT_LLAVE!SUT_FLAG_CC) <> 1 And PUB_TIPMOV = 10 And pub_signo_car <> 0 And Nulo_Valors(cli_llave!CLI_estado) <> "A" Then
   'MsgBox "              !!! O J O !!! " + Chr(13) + "Cliente No está ACTIVO en el Sistema.", 48, Pub_Titulo
'End If

Exit Sub
sigue:
If Err.Number = 35600 Then
  Exit Sub
End If
fin:
OJO:
Exit Sub
buscar:
If Left(i_codcli.Text, 2) = "++" Then
 VAR = Mid(i_codcli.Text, 3, Len(i_codcli.Text))
 numarchi = alta_vista_nombre(LV_CLI, VAR, PUB_CP, "D")
Else
 VAR = Mid(i_codcli.Text, 2, Len(i_codcli.Text))
 numarchi = alta_vista_nombre(LV_CLI, VAR, PUB_CP)
End If
If numarchi = 0 Then
  LV_CLI.Visible = False
  MsgBox "Alta Vista: No Existe .. Esta descripcion..", 48, Pub_Titulo
Else
  LV_CLI.Visible = True
  i_codcli.SetFocus
End If
loc_key = 1
Exit Sub


End Sub


Private Sub i_codcli_KeyUp(KeyCode As Integer, Shift As Integer)
Dim VAR
If Len(i_codcli.Text) = 0 Or IsNumeric(i_codcli.Text) Then
   LV_CLI.Visible = False
   Exit Sub
End If
If LV_CLI.Visible = False Or Len(Trim(i_codcli.Text)) = 1 Then
   loc_key = 0
    VAR = Asc(i_codcli.Text)
    VAR = VAR + 1
    If VAR = 33 Or VAR = 91 Then
       VAR = "ZZZZZZZZ"
    Else
       VAR = Chr(VAR)
    End If
    numarchi = 1
    archi = "SELECT CLI_CODCLIE , CLI_CODCIA, CLI_CP, CLI_NOMBRE, CLI_CASA_DIREC,CLI_ZONA_NEW, CLI_CASA_NUM, TAB_NOMLARGO  FROM CLIENTES,TABLAS WHERE (CLI_ESTADO = 'A') AND (TAB_CODCIA = '00') AND (TAB_TIPREG = 35) AND (TAB_NUMTAB = CLI_ZONA_NEW) AND CLI_CP = '" & PUB_CP & "' AND CLI_CODCIA = '" & LK_CODCIA & "' AND CLI_NOMBRE BETWEEN '" & i_codcli.Text & "' AND  '" & VAR & "' ORDER BY CLI_NOMBRE"
    PROC_LISVIEW LV_CLI
    loc_key = 0
    If LV_CLI.Visible Then
     loc_key = 1
    End If
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As Object    ' Variable FoundItem.
If LV_CLI.Visible Then
  Set itmFound = LV_CLI.FindItem(LTrim(i_codcli.Text), lvwText, , lvwPartial)
  If itmFound Is Nothing Then
  Else
   itmFound.EnsureVisible
   itmFound.Selected = True
   loc_key = itmFound.Tag
   If loc_key + 8 > LV_CLI.ListItems.count Then
      LV_CLI.ListItems.item(LV_CLI.ListItems.count).EnsureVisible
   Else
     LV_CLI.ListItems.item(loc_key + 8).EnsureVisible
   End If
  End If
  Exit Sub
End If
End Sub


Private Sub i_codcli_LostFocus()
Dim xcuenta As Integer
Dim defecto As Integer
Dim key_descto As String * 2
Dim tf As Integer
If Not IsNumeric(i_codcli.Text) Then
 Exit Sub
End If
Dim ww_num As Integer
If cli_llave.EOF = False And Val(i_codcli.Text) > 0 Then
   If cli_llave!cli_codclie <> Val(i_codcli.Text) Then
      i_codcli.Text = ""
      Exit Sub
   End If
End If

'LV_CLI.Visible = False
'    If fracli.Visible Then
'        t_nombre.SetFocus
'    End If
End Sub

Private Sub i_codven_Change()
    If i_codcli.Text = "" Then
        LV_VEN.Visible = False
        lblnomven.Caption = ""
    End If
End Sub

Private Sub i_codven_GotFocus()
    Azul i_codven, i_codven
End Sub

Private Sub i_codven_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As Object    ' Variable FoundItem.

If Not LV_VEN.Visible Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And i_codcli.Text = "" Then
  loc_key = 1
  Set LV_VEN.SelectedItem = LV_VEN.ListItems(loc_key)
  LV_VEN.ListItems.item(loc_key).Selected = True
  LV_VEN.ListItems.item(loc_key).EnsureVisible
  GoTo fin
End If

If KeyCode = 40 Then  ' flecha abajo
  loc_key = loc_key + 1
  If loc_key > LV_VEN.ListItems.count Then loc_key = LV_VEN.ListItems.count
  GoTo POSICION
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo POSICION
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > LV_VEN.ListItems.count Then loc_key = LV_VEN.ListItems.count
 GoTo POSICION
End If
If KeyCode = 33 Then
 loc_key = loc_key - 17
 If loc_key < 1 Then loc_key = 1
 GoTo POSICION
End If
GoTo fin
POSICION:
'  KeyCode = 0
  LV_VEN.ListItems.item(loc_key).Selected = True
  LV_VEN.ListItems.item(loc_key).EnsureVisible
  i_codven.Text = Trim(LV_VEN.ListItems.item(loc_key).Text) & " "
  DoEvents
  i_codven.SelStart = Len(i_codven.Text)
  DoEvents
fin:
End Sub

Private Sub i_codven_KeyPress(KeyAscii As Integer)
Dim valor As String
Dim tf As Integer
Dim i
Dim itmFound As Object    ' Variable FoundItem.

If KeyAscii = 27 Then
  i_codven.Text = ""
  LV_VEN.Visible = False
  Exit Sub
End If
If KeyAscii <> 13 Then
   GoTo fin
End If
On Error GoTo OJO
PUB_CODVEN = Val(i_codven.Text)
On Error GoTo 0
If Len(i_codven.Text) = 0 Then
   Exit Sub
End If
If PUB_CODVEN <> 0 And IsNumeric(i_codven.Text) = True Then
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   LEER_VEN_LLAVE
   If ven_llave.EOF Then
    Azul i_codven, i_codven
    MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
    i_codven.SetFocus
    GoTo fin
   Else
      lblnomven.Caption = ven_llave(2)
      Azul txt_key, txt_key
   End If
Else
On Error GoTo sigue
   If loc_key > 0 Then valor = UCase(LV_VEN.ListItems.item(loc_key).Text)
   If Trim(UCase(i_codven.Text)) = Left(valor, Len(Trim(i_codven.Text))) Then
   Else
      Exit Sub
   End If
   
   If loc_key = 0 Then Exit Sub
   
   i_codven.Text = Trim(LV_VEN.ListItems.item(loc_key).SubItems(1))
   PUB_CODVEN = Val(i_codven.Text)
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   LEER_VEN_LLAVE
   lblnomven.Caption = Trim(ven_llave(2))
   Azul txt_key, txt_key
End If
LV_VEN.Visible = False
fin:
sigue:
OJO:
End Sub

Private Sub i_codven_KeyUp(KeyCode As Integer, Shift As Integer)
Dim VAR
Dim pos1, CARAC, i, izq, der


If Len(i_codven.Text) = 0 Or IsNumeric(i_codven.Text) Then
   LV_VEN.Visible = False
   Exit Sub
End If
If LV_VEN.Visible = False Or Len(i_codven.Text) = 1 Then
    loc_key = 0
    VAR = Asc(i_codven.Text)
    VAR = VAR + 1
    If VAR = 33 Or VAR = 91 Then
       VAR = "ZZZZZZZZ"
    Else
       VAR = Chr(VAR)
    End If
    numarchi = 2
    archi = "SELECT VEM_CODVEN , VEM_CODCIA, VEM_NOMBRE  FROM VEMAEST WHERE  VEM_CODCIA = '" & LK_CODCIA & "' AND VEM_NOMBRE BETWEEN '" & i_codven.Text & "' AND  '" & VAR & "' ORDER BY VEM_NOMBRE"
    PROC_LISVIEW LV_VEN
    loc_key = 0
    If LV_VEN.Visible Then
     loc_key = 1
    End If
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As Object    ' Variable FoundItem.
If LV_VEN.Visible Then
  Set itmFound = LV_VEN.FindItem(LTrim(i_codcli.Text), lvwText, , lvwPartial)
  If itmFound Is Nothing Then
  Else
   itmFound.EnsureVisible
   itmFound.Selected = True
   loc_key = itmFound.Tag
   If loc_key + 8 > LV_VEN.ListItems.count Then
      LV_VEN.ListItems.item(LV_VEN.ListItems.count).EnsureVisible
   Else
     LV_VEN.ListItems.item(loc_key + 8).EnsureVisible
   End If
   DoEvents
  End If
  Exit Sub
End If
Exit Sub
End Sub
Private Sub i_fecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not IsDate(i_fecha.Text) Then
            MsgBox "Fecha NO procede ...", 48, Pub_Titulo
            i_fecha.SetFocus
            Exit Sub
        End If
        Azul txtdias, txtdias
    End If
End Sub


Private Sub Label4_DblClick()
    Call LlenaPorcentaje(5, " H.M.", Val(vPorHM))
End Sub

Private Sub lblIM_DblClick()
   Call LlenaPorcentaje(1, " I.M.", Val(lblIM.Caption))
End Sub

Private Sub lblnom_DblClick(index As Integer)
    Select Case index
    
    Case 7
        Call LlenaPorcentaje(2, " G.G.", Val(vPorGG))
    Case 9
        Call LlenaPorcentaje(3, " G.A.", Val(vPorGA))
    Case 14
        Call LlenaPorcentaje(6, "F. Garantia", Val(vPorFGaran))
    Case 15
        Call LlenaPorcentaje(7, " Locador", Val(vPorLocador))
    End Select
    
End Sub

Private Sub lblTIM_DblClick()
    Call LlenaPorcentaje(4, " T.I.M.", Val(lblTIM.Caption))
End Sub

Private Sub ListView1_DblClick()
 loc_key = ListView1.SelectedItem.index
 txt_key.Text = Trim(ListView1.ListItems.item(loc_key).Text) & " "
 txt_key_KeyPress 13
End Sub

Private Sub ListView1_GotFocus()
If loc_key <> 0 Then
 Set ListView1.SelectedItem = ListView1.ListItems(loc_key)
 ListView1.ListItems.item(loc_key).Selected = True
 ListView1.ListItems.item(loc_key).EnsureVisible
End If
End Sub

Private Sub ListView1_ItemClick(ByVal item As MSComctlLib.ListItem)
If loc_key <> 0 Then
 loc_key = ListView1.SelectedItem.index
 txt_key.Text = Trim(ListView1.ListItems.item(loc_key).Text) & " "
End If
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  ListView1.Visible = False
  txt_key.Text = ""
  txt_key.SetFocus
  Exit Sub
End If
If KeyAscii <> 13 Then
 Exit Sub
End If
ListView1_DblClick
End Sub

Private Sub ListView1_LostFocus()
ListView1.Visible = False
End Sub

Private Sub LV_CLI_DblClick()
loc_key = LV_CLI.SelectedItem.index
i_codcli.Text = Trim(LV_CLI.ListItems.item(loc_key).Text) & " "
i_codcli_KeyPress 13
End Sub

Private Sub LV_CLI_GotFocus()
If loc_key <> 0 Then
 Set LV_CLI.SelectedItem = LV_CLI.ListItems(loc_key)
 LV_CLI.ListItems.item(loc_key).Selected = True
 LV_CLI.ListItems.item(loc_key).EnsureVisible
End If
End Sub

Private Sub LV_CLI_ItemClick(ByVal item As MSComctlLib.ListItem)
If loc_key <> 0 Then
  loc_key = LV_CLI.SelectedItem.index
  i_codcli.Text = Trim(LV_CLI.ListItems.item(loc_key).Text) & " "
End If
End Sub

Private Sub LV_CLI_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 If i_codcli.Visible Then
    LV_CLI.Visible = False
    i_codcli.Text = ""
    i_codcli.SetFocus
 End If
End If
If KeyAscii = 13 Then
 If i_codcli.Visible Then
    i_codcli_KeyPress 13
 End If
End If
End Sub

Private Sub LV_CLI_LostFocus()
LV_CLI.Visible = False
End Sub

Private Sub salir_Click()
    Unload Me
End Sub

Private Sub Text1_Change()
    
End Sub

Private Sub Txt_key_Change()
    If txt_key = "" Then
        Call LimpiarValores
    End If
End Sub

Private Sub txt_key_GotFocus()
    If ListView1.Visible Then
        ListView1.Visible = False
    End If
End Sub

Private Sub txt_key_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As Object    ' Variable FoundItem.
'mic para buscador

If Not ListView1.Visible Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And txt_key.Text = "" Then
  loc_key = 1
  Set ListView1.SelectedItem = ListView1.ListItems(loc_key)
  ListView1.ListItems.item(loc_key).Selected = True
  ListView1.ListItems.item(loc_key).EnsureVisible
  GoTo fin
End If

If KeyCode = 40 Then  ' flecha abajo
  loc_key = loc_key + 1
  If loc_key > ListView1.ListItems.count Then loc_key = ListView1.ListItems.count
  GoTo POSICION
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo POSICION
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > ListView1.ListItems.count Then loc_key = ListView1.ListItems.count
 GoTo POSICION
End If
If KeyCode = 33 Then
 loc_key = loc_key - 17
 If loc_key < 1 Then loc_key = 1
 GoTo POSICION
End If
GoTo fin
POSICION:
  ListView1.ListItems.item(loc_key).Selected = True
  ListView1.ListItems.item(loc_key).EnsureVisible
  txt_key.Text = Trim(ListView1.ListItems.item(loc_key).Text) & " "
  DoEvents
  txt_key.SelStart = Len(txt_key.Text)
  DoEvents
fin:
End Sub

Private Sub txt_key_KeyPress(KeyAscii As Integer)
Dim valor As String
Dim tf As Integer
Dim i
Dim itmFound As Object

If KeyAscii = 27 Then
 txt_key.Text = ""
End If
If KeyAscii <> 13 Then
   GoTo fin
End If
On Error GoTo ERROR_CODIGO
  'pu_codclie = Val(txt_key.Text)
On Error GoTo 0
If Len(txt_key.Text) = 0 Then
   Exit Sub
End If

If pu_codclie <> 0 And IsNumeric(txt_key.Text) = True Then
   SQ_OPER = 1
   PUB_CODCIA = LK_CODCIA
   On Error GoTo ERROR_CODIGO
    PUB_KEY = txt_key.Text
    'lblProd.Caption = Trim(UCase(ListView1.ListItems.Item(loc_key).Text))
    LEER_ART_LLAVE
   On Error GoTo 0
   If art_LLAVE.EOF Then
     MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
     Azul txt_key, txt_key
     GoTo fin
   Else
     If pu_codclie = 1 Then
       MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
       Azul txt_key, txt_key
       GoTo fin
     End If
'     LLENA_ARTI 1
'     BLOQUEA_TEXT frmARTI.txt_key
'     frmARTI.cmdModificar.SetFocus
'     BLOQUEA_TEXT txtnombre
'     cmdCancelar.Enabled = True
    lblProd.Caption = Trim(UCase(art_LLAVE!ART_NOMBRE))
    Call LLENA_ARTI
   End If
Else
   If loc_key > ListView1.ListItems.count Or loc_key = 0 Then
     Exit Sub
   End If
   valor = Trim(UCase(ListView1.ListItems.item(loc_key).Text))
   lblProd.Caption = valor
   If Trim(UCase(txt_key.Text)) = Left(valor, Len(Trim(txt_key.Text))) Then
   Else
      Exit Sub
   End If
   txt_key.Text = Trim(ListView1.ListItems.item(loc_key).SubItems(1))
   Call LLENA_ARTI
'   LLENA_ARTI 0
'   BLOQUEA_TEXT frmARTI.txt_key
'   frmARTI.cmdModificar.SetFocus
'   BLOQUEA_TEXT txtnombre
'   cmdCancelar.Enabled = True
End If
dale:
ListView1.Visible = False
fin:
Exit Sub
ERROR_CODIGO:
MsgBox "Codigo NO Valido .... ", 48, Pub_Titulo
Azul txt_key, txt_key
End Sub

Private Sub txt_key_KeyUp(KeyCode As Integer, Shift As Integer)
Dim VAR
Dim ws_codcia As String * 2

If Len(txt_key.Text) = 0 Or IsNumeric(txt_key.Text) = True Then
   ListView1.Visible = False
   Exit Sub
End If
If (ListView1.Visible = False And KeyCode <> 13 Or Len(txt_key.Text) = 1) Or (Left(txt_key.Text, 1) = "%" And Trim(Len(txt_key.Text)) > 1) Then
    If txt_key.Text = "" Then txt_key.Text = " "
    VAR = Asc(txt_key.Text)
    VAR = VAR + 1
    If VAR = 33 Or VAR = 91 Then
       VAR = "ZZZZZZZZ"
    ElseIf VAR = 58 Then
       VAR = "A"
    Else
       VAR = Chr(VAR)
    End If
    ws_codcia = LK_CODCIA
    If LK_EMP_PTO = "A" Then
      ws_codcia = "00"
    End If
        
    numarchi = 0
    If Left(txt_key.Text, 1) <> "%" Then
    ''  archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK ,PRE_EQUIV FROM ARTI, ARTICULO, PRECIOS WHERE (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND ART_KEY <> 0 AND ART_KEY  <> 1 and ART_CODCIA = '" & ws_codcia & "' AND ART_NOMBRE BETWEEN '" & Txt_key.Text & "' AND  '" & var & "' ORDER BY ART_NOMBRE"
        archi = "SELECT ARTI.ART_KEY,ARTI.ART_CODCIA, ARTI.ART_NOMBRE, ARTI.ART_ALTERNO, ARTICULO.ARM_STOCK, PRECIOS.PRE_EQUIV, TABLAS.TAB_NOMLARGO AS DIVI, TABLAS_1.TAB_NOMLARGO AS LINEA, PRECIOS.PRE_PRE1, PRECIOS.PRE_PRE4,ARTI.ART_CUENTA_CONTAB,ARTI.ART_CUENTA_CONTAB_C,PRECIOS.PRE_PRE11,PRECIOS.PRE_PRE22,ARTI.ART_FAMILIA,ARTI.ART_SUBFAM "
        archi = archi & "FROM ARTI INNER JOIN ARTICULO ON ARTI.ART_KEY = ARTICULO.ARM_CODART AND ARTI.ART_CODCIA = ARTICULO.ARM_CODCIA INNER JOIN PRECIOS ON ARTI.ART_KEY = PRECIOS.PRE_CODART AND ARTI.ART_CODCIA = PRECIOS.PRE_CODCIA INNER JOIN TABLAS ON ARTI.ART_CODCIA = TABLAS.TAB_CODCIA AND ARTI.ART_FAMILIA = TABLAS.TAB_NUMTAB INNER JOIN TABLAS TABLAS_1 ON ARTI.ART_CODCIA = TABLAS_1.TAB_CODCIA AND ARTI.ART_SUBFAM = TABLAS_1.TAB_NUMTAB "
        archi = archi & "WHERE (TABLAS.TAB_TIPREG = 122) AND (TABLAS_1.TAB_TIPREG = 123) AND (PRECIOS.PRE_FLAG_UNIDAD = 'A') AND ARTI.ART_CODCIA = '" & LK_CODCIA & "' AND ARTI.ART_CALIDAD = 1 AND (ARTI.ART_FLAG_STOCK = 'M' OR ARTI.ART_FLAG_STOCK = 'P') AND ARTI.ART_NOMBRE BETWEEN '" & Trim(txt_key.Text) & "%' AND  '" & VAR & "' ORDER BY ARTI.ART_NOMBRE"
    Else
        If KeyCode = 13 Then
        archi = "SELECT ARTI.ART_KEY,ARTI.ART_CODCIA, ARTI.ART_NOMBRE, ARTI.ART_ALTERNO, ARTICULO.ARM_STOCK, PRECIOS.PRE_EQUIV, TABLAS.TAB_NOMLARGO AS DIVI, TABLAS_1.TAB_NOMLARGO AS LINEA, PRECIOS.PRE_PRE1, PRECIOS.PRE_PRE4,ARTI.ART_CUENTA_CONTAB,ARTI.ART_CUENTA_CONTAB_C,PRECIOS.PRE_PRE11,PRECIOS.PRE_PRE22,ARTI.ART_FAMILIA,ARTI.ART_SUBFAM "
        archi = archi & "FROM ARTI INNER JOIN ARTICULO ON ARTI.ART_KEY = ARTICULO.ARM_CODART AND ARTI.ART_CODCIA = ARTICULO.ARM_CODCIA INNER JOIN PRECIOS ON ARTI.ART_KEY = PRECIOS.PRE_CODART AND ARTI.ART_CODCIA = PRECIOS.PRE_CODCIA INNER JOIN TABLAS ON ARTI.ART_CODCIA = TABLAS.TAB_CODCIA AND ARTI.ART_FAMILIA = TABLAS.TAB_NUMTAB INNER JOIN TABLAS TABLAS_1 ON ARTI.ART_CODCIA = TABLAS_1.TAB_CODCIA AND ARTI.ART_SUBFAM = TABLAS_1.TAB_NUMTAB "
        archi = archi & "WHERE (TABLAS.TAB_TIPREG = 122) AND (TABLAS_1.TAB_TIPREG = 123) AND (PRECIOS.PRE_FLAG_UNIDAD = 'A') AND ARTI.ART_CODCIA = '" & LK_CODCIA & "' AND ARTI.ART_CALIDAD = 1 AND (ARTI.ART_FLAG_STOCK = 'M'  OR ARTI.ART_FLAG_STOCK = 'P) AND ARTI.ART_NOMBRE like '" & Trim(txt_key.Text) & "%' ORDER BY ARTI.ART_NOMBRE"
        Else
            Exit Sub
        End If
    End If
    PROC_LISVIEW ListView1, 3000
    loc_key = 0
    If ListView1.Visible Then
    loc_key = 1
    End If
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As Object     ' Variable FoundItem.
If ListView1.Visible Then
  Set itmFound = ListView1.FindItem(LTrim(txt_key.Text), lvwText, , lvwPartial)
  If itmFound Is Nothing Then
  Else
   itmFound.EnsureVisible
   itmFound.Selected = True
   loc_key = itmFound.Tag
   If loc_key + 8 > ListView1.ListItems.count Then
      ListView1.ListItems.item(ListView1.ListItems.count).EnsureVisible
   Else
     ListView1.ListItems.item(loc_key + 8).EnsureVisible
   End If
   DoEvents
  End If
  Exit Sub
End If
End Sub


Private Sub txtdias_Change()
    If Trim(txtdias.Text) = "" Then
        txtdias.Text = "0"
    End If
End Sub

Private Sub txtdias_KeyPress(KeyAscii As Integer)
    SOLO_ENTERO KeyAscii
    If KeyAscii = 13 Then
        i_fechaDias.Text = Format(DateAdd("d", Val(txtdias.Text), i_fecha.Text), "dd/mm/yy")
        Azul i_codven, i_codven
    End If
End Sub

Private Sub txtInicialDol_KeyPress(KeyAscii As Integer)
    SOLO_ENTERO KeyAscii
    If KeyAscii = 13 Then
        If Val(txt_key) > 0 And Val(lblpvtaDol.Caption) > 0 And Val(lblpvtaLista.Caption) > 0 Then
            If Val(txtTipoCambio.Text) > 0 And _
                Val(txtInicialDol.Text) > 0 And _
                Val(txtPeriodos.Text) > 0 Then
                
                LLENA_ARTI
            Else
                If Val(txtTipoCambio.Text) = 0 Then
                    txtTipoCambio.Text = "0.00"
                End If
                If Val(txtInicialDol.Text) = 0 Then
                    txtInicialDol.Text = "0"
                End If
                If Val(txtPeriodos.Text) = 0 Then
                    txtPeriodos = "0"
                End If
            End If
        End If
        Azul txtPeriodos, txtPeriodos
    End If
End Sub

Private Sub txtPeriodos_KeyPress(KeyAscii As Integer)
    SOLO_ENTERO KeyAscii
    If KeyAscii = 13 Then
        If Val(txt_key) > 0 And Val(lblpvtaDol.Caption) > 0 And Val(lblpvtaLista.Caption) > 0 Then
            If Val(txtTipoCambio.Text) > 0 And _
                Val(txtInicialDol.Text) > 0 And _
                Val(txtPeriodos.Text) > 0 Then
                
                LLENA_ARTI
            
            Else
                If Val(txtTipoCambio.Text) = 0 Then
                    txtTipoCambio.Text = "0.00"
                End If
                If Val(txtInicialDol.Text) = 0 Then
                    txtInicialDol.Text = "0"
                End If
                If Val(txtPeriodos.Text) = 0 Then
                    txtPeriodos = "0"
                End If
            End If
        End If
        Azul txt_key, txt_key
    End If
End Sub

Private Sub txtTipoCambio_KeyPress(KeyAscii As Integer)
    SOLO_DECIMAL txtTipoCambio, KeyAscii
    If KeyAscii = 13 Then
        If Val(txt_key) > 0 And Val(lblpvtaDol.Caption) > 0 And Val(lblpvtaLista.Caption) > 0 Then
            If Val(txtTipoCambio.Text) > 0 And _
                Val(txtInicialDol.Text) > 0 And _
                Val(txtPeriodos.Text) > 0 Then
                
                LLENA_ARTI
            Else
                If Val(txtTipoCambio.Text) = 0 Then
                    txtTipoCambio.Text = "0.00"
                End If
                If Val(txtInicialDol.Text) = 0 Then
                    txtInicialDol.Text = "0"
                End If
                If Val(txtPeriodos.Text) = 0 Then
                    txtPeriodos = "0"
                End If
            End If
        End If
        Azul txtInicialDol, txtInicialDol
    End If
End Sub

'Private Sub LlenaMotos()
'Dim SQL As String
'Dim RDQMotos As rdoQuery
'Dim RDRMotos As rdoResultset
'
'SQL = "Select art_key,art_nombre from Arti where art_codcia='" & LK_CODCIA & "' and art_familia=1 and art_subfam=2"
'Set RDQMotos = CN.CreateQuery("", SQL)
'Set RDRMotos = RDQMotos.OpenResultset(rdOpenKeyset, rdConcurValues)
'RDRMotos.Requery
'
'cboProd.Clear
'Do While Not RDRMotos.EOF
'    cboProd.AddItem Trim(RDRMotos!ART_NOMBRE)
'    cboProd.ItemData(cboProd.NewIndex) = RDRMotos!ART_KEY
'    RDRMotos.MoveNext
'Loop
'End Sub

Private Sub LLENA_ARTI()
Dim SQL As String
Dim RDQMotos As rdoQuery
Dim RDRMotos As rdoResultset
Dim diferencia As String
'Dim tasa As Double, C As Integer
'Dim VA As Double,P As Double, Arriba As Double, ABAJO As Double, expo As Double
Dim mess As String

'SQL = "Select art_key,art_nombre from Arti where art_codcia='" & LK_CODCIA & "' and art_familia=1 and art_subfam=2"
 SQL = "SELECT ARTI.ART_key,ARTI.ART_NOMBRE,ARTI.ART_ALTERNO,PRECIOS.PRE_UNIDAD,ARTICULO.ARM_STOCK,PRECIOS.PRE_PRE11,PRECIOS.PRE_PRE22 "
    SQL = SQL & "FROM ARTI INNER JOIN PRECIOS ON ARTI.ART_KEY = PRECIOS.PRE_CODART AND ARTI.ART_CODCIA = PRECIOS.PRE_CODCIA "
    SQL = SQL & " INNER JOIN ARTICULO ON ARTI.ART_KEY = ARTICULO.ARM_CODART AND ARTI.ART_CODCIA = ARTICULO.ARM_CODCIA "
    SQL = SQL & "WHERE ARTI.ART_CODCIA = '" & LK_CODCIA & "' AND ARTI.ART_KEY=" & txt_key


Set RDQMotos = CN.CreateQuery("", SQL)
Set RDRMotos = RDQMotos.OpenResultset(rdOpenKeyset, rdConcurValues)
RDRMotos.Requery
If RDRMotos.EOF Then Exit Sub
If Val(Nulo_Valor0(RDRMotos!pre_pre11)) > 0 And Val(Nulo_Valor0(RDRMotos!PRE_PRE22)) > 0 Then
    txt_key.Text = RDRMotos!ART_KEY
    lblProd.Caption = Trim(RDRMotos!ART_NOMBRE)
    lblpvtaDol.Caption = Format(Nulo_Valor0(RDRMotos!pre_pre11), "0.00") 'Precio Financiado
    lblpvtaLista.Caption = Format(Nulo_Valor0(RDRMotos!PRE_PRE22), "0.00") 'Precio Cash
    lblIMDol.Caption = Format(Val(lblpvtaDol.Caption) * ((Val(lblIM.Caption) / 100)), "0.00")
    lblpvtaSol.Caption = Format(Val(lblpvtaDol.Caption) * Val(txtTipoCambio.Text), "0.00")
    lblIMSol.Caption = Format(Val(lblpvtaSol.Caption) * ((Val(lblIM.Caption) / 100)), "0.00")
    If Val(txtTipoCambio.Text) = 0 Or _
        Val(txtInicialDol.Text) = 0 Or _
        Val(txtPeriodos.Text) = 0 Then
        
        ListView1.Visible = False
        'MsgBox "Faltan datos para el cálculo...", vbExclamation, Pub_Titulo
        Azul txtTipoCambio, txtTipoCambio
        Exit Sub
    End If
    lblPorPVta.Caption = Format((Val(txtInicialDol.Text) / Val(lblpvtaDol.Caption)) * 100, "0.00")
    lblInicialSol.Caption = Format(Val(txtInicialDol.Text) * Val(txtTipoCambio.Text), "0.00")
    diferencia = Format(Val(lblpvtaDol.Caption) - Val(txtInicialDol.Text), "0.00")
    lblGG.Caption = Format(Val(diferencia) * ((Val(vPorGG) / 100)), "0.00")
    lblGA.Caption = Format(Val(diferencia) * ((Val(vPorGA) / 100)), "0.00")
    lblMAF.Caption = Format(Val(diferencia) + Val(lblGG.Caption) + Val(lblGA.Caption), "0.00")
    lblBoleta.Caption = Format(Val(lblpvtaDol.Caption) + Val(lblGG.Caption) + Val(lblGA.Caption), "0.00")
    lblDIF.Caption = Format(Val(diferencia), "0.00")
    lblCOM.Caption = Format(Val(lblGG.Caption) + Val(lblGA.Caption), "0.00")
    lblAbonoHM.Caption = Format(Val(lblMAF.Caption) * ((Val(vPorHM) / 100)), "0.00")
    lblFGarantia.Caption = Format(Val(lblMAF.Caption) * ((Val(vPorFGaran) / 100)), "0.00")
    lblLocador.Caption = Format(Val(lblMAF.Caption) * ((Val(vPorLocador) / 100)), "0.00")
    lblMAF1.Caption = Format(Val(lblAbonoHM.Caption) + Val(lblFGarantia.Caption) + Val(lblLocador.Caption), "0.00")
    lblCuotaDol.Caption = Format(Pago(Val(lblTIM.Caption), Val(txtPeriodos.Text), Val(lblMAF.Caption)), "0.00")
    lblCuotaSol.Caption = Format(Val(lblCuotaDol.Caption) * Val(txtTipoCambio.Text), "0.00")
    Azul txtTipoCambio, txtTipoCambio
Else
    Call LimpiarValores
    ListView1.Visible = False
    If Val(Nulo_Valor0(RDRMotos!pre_pre11)) <= 0 Then
        mess = "- No existe Precio Financiado...." & vbCrLf & vbCrLf
    End If
    If Val(Nulo_Valor0(RDRMotos!PRE_PRE22)) <= 0 Then
        mess = mess & "- No existe Precio Cash...."
    End If
    MsgBox mess, vbExclamation, Pub_Titulo
    Azul txt_key, txt_key
End If
End Sub

Private Sub LlenaPorcentaje(ByVal opt As Integer, titu As String, PorDefault As Double)
    Dim valor As String
    valor = Trim(InputBox("Ingrese Porcentaje a Calcular...", Pub_Titulo & " --- " & titu, PorDefault))
    If valor = "" Then Exit Sub
    If IsNumeric(valor) = True And Val(valor) > 0 Then
    
        Select Case opt
        
        Case 1 'I.M.
            lblIM.Caption = valor
        Case 2 'G.G.
            vPorGG = valor
        Case 3 'G.A.
            vPorGA = valor
        Case 4 'T.I.M.
            lblTIM.Caption = valor
        Case 5 'HM
            vPorHM = valor
        Case 6 'F. Garantia
            vPorFGaran = valor
        Case 7 'Locador
            vPorLocador = valor
        End Select
        Azul txt_key, txt_key
    Else
        If Not IsNumeric(valor) Then
            MsgBox "Dato incorrecto....Verifique", vbExclamation, Pub_Titulo
        End If
    End If
End Sub
Private Sub LlenaConstantes()
    lblIM.Caption = "18"
    lblTIM.Caption = "2.95"
    vPorGG = "11.9"
    vPorGA = "8"
    vPorHM = "83.4"
    vPorFGaran = "8.34"
    vPorLocador = "8.26"
    txtTipoCambio.Text = "3.29"
End Sub

Private Function Pago(ByVal Portasa As Double, ByVal c As Integer, ByVal VA As Double) As Double

    Dim p As Double, Arriba As Double, ABAJO As Double, expo As Double, Tasa As Double

    Tasa = Portasa / 100
    expo = Format((1 + Tasa) ^ c, "0.0000000")
    Arriba = Format(Tasa * expo, "0.0000000")
    ABAJO = Format(expo - 1, "0.0000000")
    p = Format(VA * (Arriba / ABAJO), "0.00")
    Pago = p
End Function

Private Sub GrabaCotizacion()
On Error GoTo LERROR
    Dim NumFac As Long, NumOper As Long, NumSec
    Dim Sql3 As String, Sql4 As String
    NumFac = GeneraNumFac
    NumOper = GeneraNumOperacion
    'INSERCIÓN DEL REGISTRO EN LA TABLA FACART
     Sql3 = "INSERT INTO FACART (FAR_TIPMOV,FAR_CODCIA,FAR_NUMSER,FAR_FBG,FAR_NUMFAC,FAR_NUMSEC, " & _
             "FAR_FECHA,FAR_NUMOPER,FAR_CODART,FAR_ESTADO,FAR_SIGNO_ARM,FAR_PRECIO, " & _
             "FAR_COSPRO,FAR_EQUIV,FAR_NUMSER_C,FAR_NUMFAC_C,FAR_CANTIDAD,FAR_DESCRI," & _
             "FAR_FECHA_PRO,FAR_FECHA_COMPRA,FAR_FLAG_SO,FAR_TRANSITO,FAR_NUMOPER2," & _
             "FAR_CONCEPTO,FAR_CODUSU,FAR_COSTEO,FAR_DIAS,FAR_IMPTO,FAR_TOT_DESCTO," & _
             "FAR_DESCTO,FAR_BRUTO,FAR_PORDESCTO1,FAR_TIPO_CAMBIO,FAR_CODVEN,FAR_MONEDA,FAR_COSPRO_ANT," & _
             "FAR_COSPRO_SUP,FAR_CODCLIE,FAR_CP) " & _
             "VALUES('103','" & LK_CODCIA & "','0','','" & NumFac & "','1','" & LK_FECHA_DIA & "','" & NumOper & "','" & Val(txt_key.Text) & "','" & _
             UCase("N") & "','0','" & CDbl("0" & Val(lblpvtaDol.Caption)) & "','" & _
             CDbl("0" & txtInicialDol.Text) & "','" & Val(txtPeriodos.Text) & "','0','0','1','UNIDAD','" & Format(i_fecha.Text, "dd/mm/yyyy") & "','" & Format(i_fechaDias.Text, "dd/mm/yyyy") & "','','','" & _
             NumOper & "','Cotización - Créditos','" & LK_CODUSU & "','A','" & Val(txtdias.Text) & "'," & _
             "'" & CDbl("0" & lblGG.Caption) & "','" & CDbl("0" & lblGA.Caption) & "'," & _
             "'" & CDbl("0" & lblMAF.Caption) & "','" & CDbl("0" & lblCuotaDol.Caption) & "'," & _
             "'" & CDbl("0" & lblAbonoHM.Caption) & "','" & CDbl("0" & txtTipoCambio.Text) & "'," & _
             "'" & Val(i_codven.Text) & "','D','" & CDbl("0" & lblFGarantia.Caption) & "'," & _
             "'" & CDbl("0" & lblLocador.Caption) & "','" & Val(i_codcli.Text) & "','C')"
     
     'INSERCIÓN DEL REGISTRO EN LA TABLA ALLOG
     Sql4 = "INSERT INTO ALLOG (ALL_CODCIA,ALL_FECHA_DIA,ALL_NUMOPER," & _
             "ALL_CODTRA,ALL_FLAG_EXT,ALL_CODART,ALL_SECUENCIA,ALL_FBG,ALL_CP," & _
             "ALL_CANTIDAD,ALL_NUMSER,ALL_NUMFAC,ALL_TIPMOV,ALL_NUMSER_C,ALL_NUMFAC_C," & _
             "ALL_FECHA_PRO,ALL_CODUSU,ALL_CONCEPTO,ALL_NUMOPER2,ALL_TIPO_CAMBIO,ALL_MONEDA_CAJA," & _
             "ALL_CODCLIE,ALL_CODVEN,ALL_IMPORTE_AMORT,ALL_HORA) " & _
             "VALUES('" & LK_CODCIA & "','" & LK_FECHA_DIA & "','" & NumOper & "','2404','N','" & _
             Val(txt_key.Text) & "','1','','C','1','0','" & NumFac & "','103','0','0','" & Format(i_fecha.Text, "dd/mm/yyyy") & "','" & _
             LK_CODUSU & "','Cotización - Créditos','" & NumOper & "'," & _
             "'" & Val(txtTipoCambio.Text) & "','D','" & Val(i_codcli.Text) & "'," & _
             "'" & Val(i_codven.Text) & "','" & CDbl("0" & lblCuotaDol.Caption) & "','" & Format(Now, "HH:mm:ss") & "')"
             
    cnt.BeginTrans
    cnt.Execute Sql3 + " " + Sql4, , adExecuteNoRecords
    cnt.CommitTrans
    MsgBox "Grabación se realizó Satisfactoriamente." & vbCrLf & _
        "Nº Cotización :" & CStr(NumFac), vbInformation, Pub_Titulo
    Exit Sub
LERROR:
    MsgBox Err.Description, vbCritical, Pub_Titulo
    cnt.RollbackTrans
End Sub

Private Function GeneraNumFac() As Long
    Dim VNumFac As rdoResultset
    pub_cadena = "select ISNULL(MAX(FAR_NUMFAC),0) + 1 from facart WHERE  FAR_TIPMOV='103' AND " & _
        "FAR_CODCIA='" & LK_CODCIA & "' AND FAR_NUMSER='0' AND FAR_FBG=''"
    Set VNumFac = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)
    GeneraNumFac = VNumFac(0).Value
End Function
Private Function GeneraNumOperacion() As Long
    Dim vnumoper As rdoResultset
    
    pub_cadena = "select ISNULL(MAX(ALL_NUMOPER),0) + 1 from ALLOG " & _
        "WHERE ALL_CODCIA='" & LK_CODCIA & "' AND " & _
        "ALL_FECHA_DIA='" & LK_FECHA_DIA & "'"
    Set vnumoper = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)
    GeneraNumOperacion = vnumoper(0).Value
End Function

