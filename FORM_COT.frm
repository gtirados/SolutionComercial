VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F6E4F630-E903-11D5-8BB9-0080AD40A177}#1.18#0"; "oscontrolsuser.ocx"
Begin VB.Form FORM_COT 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Cotizaciones"
   ClientHeight    =   4890
   ClientLeft      =   1500
   ClientTop       =   1140
   ClientWidth     =   6600
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "FORM_COT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4890
   ScaleWidth      =   6600
   Tag             =   "55"
   WindowState     =   2  'Maximized
   Begin VB.Frame Fracli 
      Caption         =   "Datos:"
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   120
      TabIndex        =   102
      Tag             =   "9999"
      Top             =   8160
      Visible         =   0   'False
      Width           =   7215
      Begin VB.TextBox t_nombre 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   105
         Tag             =   "9999"
         Top             =   240
         Width           =   3615
      End
      Begin VB.TextBox t_direc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   104
         Tag             =   "9999"
         Top             =   600
         Width           =   6015
      End
      Begin VB.TextBox t_doc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5760
         TabIndex        =   103
         Tag             =   "9999"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Lnom 
         AutoSize        =   -1  'True
         Caption         =   "Nombre: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Left            =   240
         TabIndex        =   108
         Tag             =   "9999"
         Top             =   255
         Width           =   645
      End
      Begin VB.Label Lnom 
         AutoSize        =   -1  'True
         Caption         =   "Dirección :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Left            =   240
         TabIndex        =   107
         Tag             =   "9999"
         Top             =   600
         Width           =   765
      End
      Begin VB.Label Lnom 
         AutoSize        =   -1  'True
         Caption         =   "Doc. Ident. :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Left            =   4800
         TabIndex        =   106
         Tag             =   "9999"
         Top             =   240
         Width           =   885
      End
   End
   Begin VB.Frame fraprecios 
      BackColor       =   &H00C0C0C0&
      Height          =   915
      Left            =   480
      TabIndex        =   85
      Tag             =   "9898"
      Top             =   6840
      Visible         =   0   'False
      Width           =   8595
      Begin VB.Label lblprecio 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   240
         Index           =   11
         Left            =   8325
         TabIndex        =   101
         Top             =   375
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lblprecio 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Index           =   10
         Left            =   8280
         TabIndex        =   100
         Top             =   135
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label Label8 
         BackColor       =   &H00808000&
         Caption         =   "Prec. Dolares"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   1
         Left            =   7320
         TabIndex        =   99
         Top             =   600
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H008B4914&
         Caption         =   "Prec. Soles"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Index           =   0
         Left            =   480
         TabIndex        =   98
         Top             =   360
         Width           =   2025
      End
      Begin VB.Label lblprecio 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.0000"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   240
         Index           =   5
         Left            =   7080
         TabIndex        =   97
         Top             =   480
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lblprecio 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.0000"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   240
         Index           =   6
         Left            =   7320
         TabIndex        =   96
         Top             =   375
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lblprecio 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   240
         Index           =   7
         Left            =   7440
         TabIndex        =   95
         Top             =   480
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lblprecio 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   240
         Index           =   8
         Left            =   7680
         TabIndex        =   94
         Top             =   360
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lblprecio 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   240
         Index           =   9
         Left            =   7200
         TabIndex        =   93
         Top             =   375
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lblprecio 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Index           =   4
         Left            =   7320
         TabIndex        =   92
         Top             =   135
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lblprecio 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Index           =   3
         Left            =   7200
         TabIndex        =   91
         Top             =   135
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lblprecio 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Index           =   2
         Left            =   7440
         TabIndex        =   90
         Top             =   480
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lblprecio 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0.0000"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Index           =   1
         Left            =   4200
         TabIndex        =   89
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lblprecio 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.0000"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Index           =   0
         Left            =   2640
         TabIndex        =   88
         Top             =   480
         Width           =   990
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Menor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2760
         TabIndex        =   87
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mayor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4440
         TabIndex        =   86
         Top             =   240
         Width           =   855
      End
   End
   Begin ComctlLib.ListView ListView2 
      Height          =   495
      Left            =   7470
      TabIndex        =   53
      Top             =   7350
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   10235904
      BackColor       =   16118252
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   495
      Left            =   5670
      TabIndex        =   17
      Top             =   7335
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   10235904
      BackColor       =   16118252
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.CommandButton siguiente 
      Height          =   405
      Left            =   11025
      Picture         =   "FORM_COT.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   6570
      Width           =   420
   End
   Begin VB.CommandButton anterior 
      Height          =   405
      Left            =   10545
      Picture         =   "FORM_COT.frx":0D84
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   6585
      Width           =   420
   End
   Begin VB.TextBox textovar 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      TabIndex        =   72
      Top             =   7155
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.CommandButton c_condi 
      Caption         =   "Condiciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2085
      TabIndex        =   51
      Top             =   7650
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame f1 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   60
      TabIndex        =   15
      Top             =   -60
      Width           =   11775
      Begin VB.CheckBox chkAprobacion 
         Caption         =   "Check1"
         Height          =   195
         Left            =   10425
         TabIndex        =   84
         Top             =   1245
         Visible         =   0   'False
         Width           =   195
      End
      Begin OSControlsUser.OSFindItem txtChofer 
         Height          =   285
         Left            =   4320
         TabIndex        =   77
         Top             =   990
         Visible         =   0   'False
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   503
         Locked          =   0   'False
      End
      Begin VB.TextBox i_dias 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1605
         TabIndex        =   1
         Top             =   1605
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox i_condi 
         BackColor       =   &H00F5F1EC&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   315
         ItemData        =   "FORM_COT.frx":1486
         Left            =   120
         List            =   "FORM_COT.frx":1488
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   405
         Width           =   2730
      End
      Begin VB.ComboBox i_destino 
         BackColor       =   &H00F5F1EC&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   315
         ItemData        =   "FORM_COT.frx":148A
         Left            =   4305
         List            =   "FORM_COT.frx":1494
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1590
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.ComboBox i_fbg 
         BackColor       =   &H00F5F1EC&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   315
         ItemData        =   "FORM_COT.frx":14AD
         Left            =   2475
         List            =   "FORM_COT.frx":14B7
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1590
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.TextBox Txt_key 
         Height          =   315
         Left            =   120
         MaxLength       =   10
         TabIndex        =   3
         Top             =   990
         Width           =   945
      End
      Begin VB.ComboBox moneda 
         BackColor       =   &H00F5F1EC&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   315
         ItemData        =   "FORM_COT.frx":14C1
         Left            =   120
         List            =   "FORM_COT.frx":14CB
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1575
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.TextBox txtruc 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7290
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   585
         Width           =   1515
      End
      Begin VB.TextBox txtcli 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2955
         TabIndex        =   4
         Top             =   420
         Width           =   1305
      End
      Begin MSMask.MaskEdBox txtfecha 
         Height          =   330
         Left            =   10380
         TabIndex        =   59
         Top             =   1545
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   14737632
         ForeColor       =   128
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lcodart 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Aprobación :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   18
         Left            =   9240
         TabIndex        =   83
         Tag             =   "9999"
         Top             =   1230
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblChofer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   5520
         TabIndex        =   79
         Top             =   990
         Width           =   3210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Chofer:"
         BeginProperty Font 
            Name            =   "Verdana"
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
         Left            =   4305
         TabIndex        =   78
         Top             =   765
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label lcodart 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Limite de Credito."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   6
         Left            =   8820
         TabIndex        =   66
         Tag             =   "9999"
         Top             =   150
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.Label lblcred 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   10410
         TabIndex        =   65
         Tag             =   "9999"
         Top             =   150
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label lbldisp 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   10410
         TabIndex        =   64
         Tag             =   "9999"
         Top             =   900
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label lcodart 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Deuda Actual :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   12
         Left            =   9075
         TabIndex        =   63
         Tag             =   "9999"
         Top             =   510
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label lblDeuda 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   10410
         TabIndex        =   62
         Tag             =   "9999"
         Top             =   525
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label lcodart 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Disponible :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   7
         Left            =   9330
         TabIndex        =   61
         Tag             =   "9999"
         Top             =   885
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label lcodart 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   16
         Left            =   9720
         TabIndex        =   60
         Tag             =   "9999"
         Top             =   1620
         Width           =   570
      End
      Begin VB.Label lcodart 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dias"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   11
         Left            =   1605
         TabIndex        =   48
         Tag             =   "9999"
         Top             =   1335
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lcodart 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Condición Venta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   47
         Tag             =   "9999"
         Top             =   165
         Width           =   1395
      End
      Begin VB.Label lcodart 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Almacen :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   9
         Left            =   4305
         TabIndex        =   46
         Tag             =   "9999"
         Top             =   1335
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.Label lblven 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1110
         TabIndex        =   45
         Top             =   1020
         Width           =   2895
      End
      Begin VB.Label lcodart 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fact./Bolet."
         BeginProperty Font 
            Name            =   "Verdana"
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
         Left            =   2475
         TabIndex        =   44
         Tag             =   "9999"
         Top             =   1335
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
         BeginProperty Font 
            Name            =   "Verdana"
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
         Left            =   120
         TabIndex        =   43
         Top             =   750
         Width           =   900
      End
      Begin VB.Label lblcli 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4305
         TabIndex        =   32
         Top             =   345
         Width           =   4410
      End
      Begin VB.Label lcodart 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Moneda : "
         BeginProperty Font 
            Name            =   "Verdana"
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
         Left            =   120
         TabIndex        =   27
         Tag             =   "9999"
         Top             =   1335
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lcodart 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Raz. Soc. :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   2940
         TabIndex        =   16
         Tag             =   "9999"
         Top             =   180
         Width           =   945
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   5190
      Left            =   10125
      TabIndex        =   50
      Top             =   1890
      Width           =   1695
      Begin VB.CommandButton SALIR 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cerrar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   330
         MaskColor       =   &H00800000&
         Picture         =   "FORM_COT.frx":14E4
         Style           =   1  'Graphical
         TabIndex        =   71
         TabStop         =   0   'False
         Tag             =   "9999"
         Top             =   3330
         Width           =   1110
      End
      Begin VB.CommandButton cmdimp 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   330
         Picture         =   "FORM_COT.frx":1D5A
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   2550
         Width           =   1110
      End
      Begin VB.CommandButton cancelar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   330
         Picture         =   "FORM_COT.frx":2D5C
         Style           =   1  'Graphical
         TabIndex        =   69
         TabStop         =   0   'False
         Tag             =   "9999"
         Top             =   1770
         Width           =   1110
      End
      Begin VB.CommandButton cmdconsulta 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Consulta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   330
         Picture         =   "FORM_COT.frx":350A
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   1005
         Width           =   1110
      End
      Begin VB.CommandButton cmdIngreso 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ingreso"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   360
         Picture         =   "FORM_COT.frx":3D6C
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   225
         Width           =   1110
      End
      Begin VB.TextBox txtdoc 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   570
         TabIndex        =   9
         Top             =   4320
         Width           =   1065
      End
      Begin VB.TextBox tserie 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   90
         TabIndex        =   8
         Top             =   4320
         Width           =   465
      End
      Begin VB.Label lcodart 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº de Cotiz:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   255
         Index           =   5
         Left            =   195
         TabIndex        =   52
         Tag             =   "9999"
         Top             =   4035
         Width           =   1260
      End
   End
   Begin VB.TextBox txtatte 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6165
      TabIndex        =   49
      Top             =   8295
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Frame condi 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2580
      TabIndex        =   35
      Top             =   7350
      Visible         =   0   'False
      Width           =   6495
      Begin VB.CommandButton Command1 
         Caption         =   "&Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         TabIndex        =   42
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox oferta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   41
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox tiempo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   39
         Top             =   480
         Width           =   3375
      End
      Begin VB.TextBox forma 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   37
         Top             =   120
         Width           =   3375
      End
      Begin VB.Label Label4 
         Caption         =   "Validez de la Oferta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   795
         Width           =   1695
      End
      Begin VB.Label c_entrega 
         Caption         =   "Tiempo de Entrega:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label c_forma 
         Caption         =   "Forma de Pago:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   330
         TabIndex        =   36
         Top             =   165
         Width           =   1335
      End
   End
   Begin ComctlLib.ProgressBar Barra 
      Height          =   210
      Left            =   5055
      TabIndex        =   12
      Tag             =   "0"
      Top             =   6780
      Visible         =   0   'False
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   370
      _Version        =   327682
      Appearance      =   0
      Min             =   77
      Max             =   91
   End
   Begin VB.Frame ESTADO 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   5190
      Left            =   75
      TabIndex        =   10
      Tag             =   "100"
      Top             =   1890
      Width           =   10020
      Begin VB.TextBox txtobservaciones 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1440
         TabIndex        =   81
         Top             =   195
         Width           =   8490
      End
      Begin VB.TextBox txtDescto 
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
         Left            =   3543
         TabIndex        =   75
         Top             =   4545
         Width           =   1320
      End
      Begin VB.ComboBox PRECIOS 
         BackColor       =   &H00F5F1EC&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   480
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   1440
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ComboBox UNIDAD 
         BackColor       =   &H00F5F1EC&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   1200
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txttotal 
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
         Left            =   8235
         TabIndex        =   20
         Top             =   4545
         Width           =   1320
      End
      Begin VB.TextBox txtigv 
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
         Left            =   5580
         TabIndex        =   19
         Top             =   4545
         Width           =   1320
      End
      Begin VB.TextBox txtvalorv 
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
         Left            =   1266
         TabIndex        =   18
         Top             =   4545
         Width           =   1320
      End
      Begin ComctlLib.ProgressBar PB 
         Height          =   210
         Left            =   4995
         TabIndex        =   14
         Top             =   4935
         Visible         =   0   'False
         Width           =   4440
         _ExtentX        =   7832
         _ExtentY        =   370
         _Version        =   327682
         Appearance      =   0
      End
      Begin MSFlexGridLib.MSFlexGrid grid_fac 
         Height          =   3600
         Left            =   105
         TabIndex        =   7
         Tag             =   "9999"
         Top             =   555
         Width           =   9840
         _ExtentX        =   17357
         _ExtentY        =   6350
         _Version        =   393216
         Rows            =   3
         FixedRows       =   2
         BackColor       =   16777215
         BackColorFixed  =   4210752
         ForeColorFixed  =   16777215
         BackColorBkg    =   16118252
         GridColorFixed  =   12632256
         FocusRect       =   2
         HighLight       =   2
         GridLinesFixed  =   1
         AllowUserResizing=   3
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
      End
      Begin VB.Label LBLSIT 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4740
         TabIndex        =   82
         Top             =   4935
         Width           =   5040
      End
      Begin VB.Label lcodart 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones:"
         BeginProperty Font 
            Name            =   "Verdana"
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
         Left            =   75
         TabIndex        =   80
         Tag             =   "9999"
         Top             =   225
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descto:"
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
         Height          =   330
         Index           =   4
         Left            =   2592
         TabIndex        =   76
         Tag             =   "9999"
         Top             =   4545
         Width           =   945
      End
      Begin VB.Label lcodart 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F2 = Grabar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Index           =   13
         Left            =   210
         TabIndex        =   58
         Tag             =   "9999"
         Top             =   4905
         Width           =   870
      End
      Begin VB.Label lcodart 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F4 = Cancelar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Index           =   14
         Left            =   2475
         TabIndex        =   57
         Tag             =   "9999"
         Top             =   4935
         Width           =   1005
      End
      Begin VB.Label lcodart 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F3 = Condición"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Index           =   15
         Left            =   1245
         TabIndex        =   56
         Tag             =   "9999"
         Top             =   4920
         Width           =   1065
      End
      Begin VB.Label lcodart 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F5 = Ingreso"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Index           =   17
         Left            =   3645
         TabIndex        =   55
         Tag             =   "9999"
         Top             =   4935
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Condición :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3210
         TabIndex        =   54
         Top             =   705
         Width           =   975
      End
      Begin VB.Label nomarti 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   180
         TabIndex        =   29
         Top             =   4230
         Width           =   7335
      End
      Begin VB.Label unid 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7560
         TabIndex        =   31
         Top             =   4230
         Width           =   1080
      End
      Begin VB.Label stock 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8685
         TabIndex        =   30
         Top             =   4230
         Width           =   1095
      End
      Begin VB.Label i_moneda 
         AutoSize        =   -1  'True
         Caption         =   "S/."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   7845
         TabIndex        =   28
         Top             =   4590
         Width           =   300
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total :"
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
         Height          =   330
         Index           =   3
         Left            =   6906
         TabIndex        =   23
         Tag             =   "9999"
         Top             =   4545
         Width           =   1320
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "I.G.V. :"
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
         Height          =   330
         Index           =   2
         Left            =   4869
         TabIndex        =   22
         Tag             =   "9999"
         Top             =   4545
         Width           =   705
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor Venta:"
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
         Height          =   330
         Index           =   1
         Left            =   105
         TabIndex        =   21
         Tag             =   "9999"
         Top             =   4545
         Width           =   1155
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Detalle del Pedido :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   11
         Tag             =   "9999"
         Top             =   690
         Width           =   2160
      End
      Begin VB.Label momen 
         Caption         =   "Un Momento ..."
         Height          =   375
         Left            =   3360
         TabIndex        =   13
         Top             =   1560
         Width           =   1575
      End
   End
   Begin Crystal.CrystalReport Reportes 
      Left            =   360
      Top             =   7080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label lcodart 
      Caption         =   "Serie"
      Height          =   255
      Index           =   3
      Left            =   7440
      TabIndex        =   34
      Tag             =   "9999"
      Top             =   240
      Width           =   525
   End
   Begin VB.Label lcodart 
      Caption         =   "Nº. Doc"
      Height          =   255
      Index           =   0
      Left            =   8040
      TabIndex        =   33
      Tag             =   "9999"
      Top             =   240
      Width           =   1125
   End
End
Attribute VB_Name = "FORM_COT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PSFAR_TRANS As rdoQuery
Dim FAR_TRANS As rdoResultset

Dim VAR_ACTIVAR As Integer
Dim WCOD_ORIGINAL As Currency
Dim WPASA As Boolean
Dim WSELE As String * 1
Dim llave1
Dim loc_key
Dim fila As Integer
Dim ws_bruto_d, ws_bruto_h As Currency
Dim SUM_D As Currency
Dim SUM_H As Currency
Dim PSTEMP_LLAVE As rdoQuery
Dim temp_llave As rdoResultset
Dim WMODO As String * 1
Dim LOC_ITEM As Integer
Dim cop_llave As rdoResultset
Dim PSCOP_LLAVE As rdoQuery
Dim LOC_CANCELA As Integer
Dim PSTEMP_MAYOR As rdoQuery
Dim temp_mayor As rdoResultset
Dim temporal
Dim wfila_act As Integer
Dim loc_ini As Integer
Dim loc_fin  As Integer
Dim Wsec As Integer
Dim PSLOC_WARTI As rdoQuery
Dim llave_sum_arti   As rdoResultset
Dim PRE_ETIQUETA(6) As String * 20
Dim LOC_TIPMOV As Integer
'====================
Dim FACTOR_DESCTO As Double

' Agregado
Dim blnConsulta As Boolean
Option Explicit

Private Sub anterior_Click()
If Val(txtdoc.Text) <= 0 Then Exit Sub
 txtdoc.Text = Val(txtdoc.Text) - 1
 PUB_NUMSER = Val(tserie.Text)
 PUB_NUMFAC = Val(txtdoc.Text)
 LLENA_DOCU
End Sub

Private Sub c_condi_Click()
If condi.Visible Then
 condi.Visible = False
Else
 condi.Visible = True
 forma.SetFocus
 
End If
End Sub

Private Sub cancelar_Click()
cmdimp.Visible = False
WMODO = ""
cmdIngreso.Caption = "&Ingreso"
f1.Enabled = False
ESTADO.Enabled = False
PB.Visible = False
fila = 0
SUM_D = 0
SUM_H = 0
LIMPIA_DATOS
CABE_MAN
f1.Enabled = False
cmdIngreso.Enabled = True
'grid_fac.SetFocus

End Sub

Private Sub cmdconsulta_Click()
blnConsulta = True ' Agregado
cmdimp.Visible = True
cmdIngreso.Enabled = False
tserie.Locked = False
txtdoc.Locked = False
tserie.Enabled = True
txtdoc.Enabled = True
'siguiente.Enabled = True
'anterior.Enabled = True
f1.Enabled = True
tserie.Text = "100"
tserie.Locked = True
Azul txtdoc, txtdoc

End Sub

Private Sub cmdimp_Click()
Call REP_CONSUL
End Sub

Private Sub cmdIngreso_Click()
Dim WMO As String
Dim RES_DEUDA As Currency
Dim wsumadol As Currency
Dim WTC As Currency
Dim ws_tot_debe, ws_tot_haber As Currency
Dim er As rdoError
Dim pub_mensaje As String
Const ingre = 2
Const MODIF = 1
Dim N As Integer
Dim LOC_SALDO_CAR As Currency
Dim FLAG As Boolean
Dim pub_mensaje_err As String
Dim WS_NRO_MOV, ws_nro_voucher As Long
Dim w_dh  As String

blnConsulta = False ' Agregado

If Left(cmdIngreso.Caption, 2) = "&G" Then
If Val(Left(LBLSIT.Caption, 2)) = 1 Then
  MsgBox "Pedido esta Procesado... no procede.", 48, Pub_Titulo
  Exit Sub
End If
If Trim(txtCli.Text) = "" Then
  MsgBox "Nombre del Cliente ", 48, Pub_Titulo
  txtCli.SetFocus
  Exit Sub
End If
If Val(txttotal.Text) <= 0 Then
  MsgBox "Ingrese Datos ", 48, Pub_Titulo
  grid_fac.SetFocus
  Exit Sub
End If

If grid_fac.rows = 3 Then
 If grid_fac.TextMatrix(2, 0) = "" Then
   MsgBox "Ingrese Datos de Productos ", 48, Pub_Titulo
   grid_fac.SetFocus
   Exit Sub
 End If
End If
pu_codcia = LK_CODCIA
pu_cp = "C"
pu_codclie = Val(txtCli.Text)
LEER_CLI_LLAVE
On Error GoTo 0
If cli_llave.EOF Then
  Azul txtCli, txtCli
  MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
  txtCli.SetFocus
  Exit Sub
End If

suma_grid
'chequeo el limite de credito
'If Not cli_llave.EOF And Nulo_Valors(par_llave!par_flag_cred) <> "A" And Nulo_Valor0(SUT_LLAVE!SUT_FLAG_CC) = 0 Then
'   If SUT_LLAVE!SUT_SIGNO_CAR = 1 Then
'      pu_codcia = LK_CODCIA
'      pub_deuda = CAR_TOT_CPX2("C", pu_codcia, cli_llave!cli_codclie)
'      If PUB_FLAG_VENCIDO = 1 And LK_FLAG_LIMITE <> "A" And LK_FLAG_LIMITE <> "C" Then
'          MsgBox "CLIENTE TIENE OBLIGACIONES VENCIDAS ... ", 48, Pub_Titulo
'          Exit Sub
'    End If
'    PUB_CAL_INI = LK_FECHA_DIA
'    PUB_CAL_FIN = LK_FECHA_DIA
'    pu_codcia = LK_CODCIA
'    PUB_CODCIA = LK_CODCIA
'    SQ_OPER = 1
'    LEER_CAL_LLAVE
'    WTC = 0
'    If Not cal_llave.EOF Then
'      WTC = Nulo_Valor0(cal_llave!cal_tipo_cambio)
'    End If
'    If WTC = 0 Then
'      MsgBox "Venta falta parametros ...INGRESE TIPO DE CAMBIO DEL DIA", 48, Pub_Titulo
'      Exit Sub
'    End If
'    If Trim(Left(moneda.Text, 1)) = "S" Then
'     wsumadol = Val(Nulo_Valor0(cli_llave!cli_limcre)) + Val(redondea((Nulo_Valor0(cli_llave!cli_limcre2) * WTC)))
'     RES_DEUDA = pub_deuda
'     WMO = "S/."
'    Else
 '    wsumadol = Val(redondea(Nulo_Valor0(cli_llave!cli_limcre) / WTC)) + Val(redondea(Val(Nulo_Valor0(cli_llave!cli_limcre2))))
 '    RES_DEUDA = redondea(Val(pub_deuda / WTC))
'     WMO = "US$."
'    End If
'If (RES_DEUDA + Val(txttotal.Text)) > wsumadol And LK_FLAG_LIMITE <> "B" And LK_FLAG_LIMITE <> "C" Then
'   MsgBox "LIMITE DE CREDITO EXCEDIDO ...SALDO POR ATENDER : " & WMO & " " & Format(wsumadol - RES_DEUDA, "0.00") & Chr(13) & "*** Venta No Procede ***", 48, Pub_Titulo
'   txtcli.Text = ""
'   Azul txtcli, txtcli
'   Exit Sub
'End If
'   End If
'End If

Barra.Visible = False

For fila = 2 To grid_fac.rows - 1
 If grid_fac.TextMatrix(fila, 1) <> "" Then
  If Val(grid_fac.TextMatrix(fila, 2)) <= 0 Then
    MsgBox "Verificar, cantidad en cero o menor. - " & grid_fac.TextMatrix(fila, 1) & " : " & grid_fac.TextMatrix(fila, 0), 48, Pub_Titulo
    grid_fac.SetFocus
    GoTo fin
  End If
  If Val(grid_fac.TextMatrix(fila, 4)) = 0 Then
    MsgBox "Verificar hay algun precio en 0 .", 48, Pub_Titulo
    grid_fac.SetFocus
    GoTo fin
  End If
End If
Next fila
Screen.MousePointer = 11
DoEvents
Barra.Visible = True
DoEvents
Barra.Min = 0
Barra.max = fila
Barra.Value = 0
exito = True
Barra.Value = 1
GoSub ACT1
'Call REP_CONSUL
fila = 1
SUM_D = 0
SUM_H = 0
CABE_MAN
LIMPIA_DATOS
fila = 0
'cancelar.SetFocus
CABE_MAN
Barra.Visible = False
f1.Enabled = False
cmdIngreso.Caption = "&Ingreso"

GoTo fin

ACT1:
If Trim(txtdoc.Text) <> 0 Then
 pub_cadena = "DELETE PEDIDOS WHERE PED_CODCIA = '" & LK_CODCIA & "' AND PED_NUMSER = " & Trim(tserie.Text) & " AND PED_NUMFAC = " & Trim(txtdoc.Text)
 CN.Execute pub_cadena, rdExecDirect
End If
fila = 1
FLAG = False
WS_NRO_MOV = 0
fila = 2
Do While FLAG = False
   If Trim(grid_fac.TextMatrix(fila, 1)) = "" Then GoTo pasa
    ' grabo todo
   temp_llave.AddNew
   temp_llave!PED_CODCIA = LK_CODCIA
   temp_llave!PED_FECHA = LK_FECHA_DIA
   temp_llave!PED_NUMSER = Trim(tserie.Text)
   temp_llave!PED_NUMFAC = Val(txtdoc.Text)
   temp_llave!PED_NUMSEC = WS_NRO_MOV
   temp_llave!PED_CANTIDAD = Val(grid_fac.TextMatrix(fila, 2))
   temp_llave!PED_PRECIO = Val(grid_fac.TextMatrix(fila, 4))
   temp_llave!PED_CODUSU = LK_CODUSU
   temp_llave!PED_IGV = Val(txtigv.Text)
   temp_llave!PED_BRUTO = Val(txtvalorv.Text)
   temp_llave!PED_ESTADO = "N"
   temp_llave!PED_CODUSU = LK_CODUSU
   temp_llave!PED_CODART = Val(grid_fac.TextMatrix(fila, 10))
   temp_llave!PED_UNIDAD = Trim(grid_fac.TextMatrix(fila, 3))
   temp_llave!PED_EQUIV = Val(grid_fac.TextMatrix(fila, 12))
   temp_llave!PED_NOMCLIE = Trim(FORM_COT.lblcli.Caption)
   temp_llave!PED_RUCCLIE = Trim(txtruc.Text) ' Trim(fbg.Text)
   temp_llave!PED_CODCLIE = Val(txtCli.Text)
   temp_llave!PED_TIPMOV = 201
   temp_llave!PED_HORA = Format(Now, "hh:mm:ss AMPM")
   Call FactorDescto(fila)
   temp_llave!ped_DESCTO = FACTOR_DESCTO
   temp_llave!ped_DESCTO_pre = Val(grid_fac.TextMatrix(fila, 19))
   
   temp_llave!PED_MONEDA = Left(Trim(moneda.Text), 1)
   temp_llave!PED_CONTACTO = txtatte.Text
   temp_llave!PED_FORMA = Trim(forma.Text)
   temp_llave!PED_TIEMPO = Trim(tiempo.Text)
   temp_llave!PED_OFERTA = Trim(grid_fac.TextMatrix(fila, 20))
   temp_llave!PED_SUBTOTAL = Val(grid_fac.TextMatrix(fila, 6))
   temp_llave!ped_CONDI = Val(Left(i_condi.Text, 2))
   temp_llave!ped_DIAS = Val(i_dias.Text)
   temp_llave!PED_CODVEN = Val(txt_key.Text)
   temp_llave!ped_DIRCLI = Val(Right(i_destino.Text, 8))
   temp_llave!ped_FBG = Trim(i_fbg.Text)
   temp_llave!PED_NUMPRE = Val(grid_fac.TextMatrix(fila, 18))
   temp_llave!PED_APROBADO = chkAprobacion.Value
   
   temp_llave.Update
pasa:
   fila = fila + 1
   WS_NRO_MOV = WS_NRO_MOV + 1
   If fila >= FORM_COT.grid_fac.rows Then
      FLAG = True
   End If
  
Loop

' Agregado para la impresion
Dim intRpta As Integer
intRpta = MsgBox("Desea imprimir la Cotizacion?", vbQuestion + vbYesNo)
If intRpta = vbYes Then
    Call REP_CONSUL
End If

Return

Screen.MousePointer = 1
Exit Sub
End If
' cuando pulsa Ingreso
Dim wser As String
Dim wnumfac As String

cmdIngreso.Caption = "&Grabar"
f1.Enabled = True
ESTADO.Enabled = True
LIMPIA_DATOS
CABE_MAN
WMODO = "I"
PSTEMP_MAYOR(0) = LK_CODCIA
temp_mayor.Requery
If temp_mayor.EOF Then
 wser = 100
 wnumfac = 1
Else
 wser = Nulo_Valors(temp_mayor!PED_NUMSER)
 wnumfac = Val(Nulo_Valor0(temp_mayor!PED_NUMFAC)) + 1
End If
tserie.Text = wser
txtdoc.Text = wnumfac

grid_fac.rows = grid_fac.rows + 1
grid_fac.RowHeight(grid_fac.rows - 1) = 285
grid_fac.TextMatrix(grid_fac.rows - 1, 0) = ""
grid_fac.TextMatrix(grid_fac.rows - 1, 1) = ""
grid_fac.TextMatrix(grid_fac.rows - 1, 2) = ""
grid_fac.TextMatrix(grid_fac.rows - 1, 3) = ""
grid_fac.TextMatrix(grid_fac.rows - 1, 4) = ""
grid_fac.TextMatrix(grid_fac.rows - 1, 5) = ""
grid_fac.TextMatrix(grid_fac.rows - 1, 6) = ""
grid_fac.TextMatrix(grid_fac.rows - 1, 7) = ""
grid_fac.TextMatrix(grid_fac.rows - 1, 8) = ""
grid_fac.TextMatrix(grid_fac.rows - 1, 9) = ""
grid_fac.TextMatrix(grid_fac.rows - 1, 10) = ""
grid_fac.TextMatrix(grid_fac.rows - 1, 11) = ""
grid_fac.TextMatrix(grid_fac.rows - 1, 12) = ""

grid_fac.TextMatrix(grid_fac.rows - 1, 14) = ""
grid_fac.TextMatrix(grid_fac.rows - 1, 15) = ""
grid_fac.TextMatrix(grid_fac.rows - 1, 16) = ""
grid_fac.TextMatrix(grid_fac.rows - 1, 17) = ""

If i_condi.ListCount > 0 And f1.Enabled = True Then
  i_condi.SetFocus
 ' SendKeys "%{up}"
End If
If moneda.ListCount > 0 And moneda.ListIndex = -1 And f1.Enabled = True Then moneda.ListIndex = 0

Exit Sub




Error_fatal:
    pub_mensaje = "Se ha producido un error " & "al abrir la conexión:" & Err & " - " & Error & vbCr
    For Each er In rdoErrors
        pub_mensaje = pub_mensaje & er.Description & ":" & er.Number & vbCr
        MsgBox pub_mensaje
    Next er
    CN.Execute "Rollback Transaction", rdExecDirect
'    Resume AbandonCn
Exit Sub

errorr:
 MsgBox pub_mensaje_err, 48, Pub_Titulo
fin:
Screen.MousePointer = 0

Exit Sub
SALE:
If Err.Number = 6 Then
  MsgBox "Verficar Importe.", 48, Pub_Titulo
  If textovar.Visible Then Azul3 textovar, textovar
  FORM_COT.Barra.Visible = False
  Screen.MousePointer = 0
  grid_fac.SetFocus
Else
  MsgBox Err.Description, 48, Pub_Titulo
End If

End Sub





Private Sub Command1_Click()
condi.Visible = False
c_condi.SetFocus
End Sub

Private Sub Form_Activate()
If i_condi.ListCount > 0 And i_condi.ListIndex = -1 Then
 i_condi.ListIndex = 0
End If

If moneda.ListCount > 0 And moneda.ListIndex = -1 And f1.Enabled = True Then moneda.ListIndex = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
  If Left(cmdIngreso.Caption, 2) = "&G" Then
     cmdIngreso_Click
  End If
ElseIf KeyCode = 114 Then
 If i_condi.Enabled Then
   i_condi.SetFocus
   SendKeys "%{up}"
  End If
ElseIf KeyCode = 115 Then
  cancelar_Click
ElseIf KeyCode = 116 Then
  If Left(cmdIngreso.Caption, 2) = "&G" Then
  Else
    cmdIngreso_Click
  End If
End If
End Sub

Private Sub Form_Load()
'On Error GoTo SALE
If CONST_CIACENTRAL = "01" Then
    chkAprobacion.Enabled = True
Else
    chkAprobacion.Enabled = False
End If
Wsec = 0
LOC_CANCELA = 0
fila = 0
wfila_act = 0
WSELE = ""
Dim ws_indice As Integer
Dim cade
WMODO = ""
Dim PSPRO_V As rdoQuery
Dim PRO_V As rdoResultset


pub_cadena = "SELECT * FROM PEDIDOS WHERE PED_CODCIA = ? AND PED_TIPMOV = ? AND PED_NUMSER = ? and PED_NUMFAC = ?  ORDER BY PED_NUMSEC"
Set PSLOC_WARTI = CN.CreateQuery("", pub_cadena)
PSLOC_WARTI(0) = 0
PSLOC_WARTI(1) = 0
PSLOC_WARTI(2) = 0
PSLOC_WARTI(3) = 0
Set llave_sum_arti = PSLOC_WARTI.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM PEDIDOS WHERE PED_CODCIA = ? AND PED_TIPMOV = 201  ORDER BY  PED_NUMFAC DESC "
Set PSTEMP_MAYOR = CN.CreateQuery("", pub_cadena)
PSTEMP_MAYOR(0) = LK_CODCIA
PSTEMP_MAYOR.MaxRows = 1
Set temp_mayor = PSTEMP_MAYOR.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM PEDIDOS WHERE  PED_TIPMOV = 201 ORDER BY PED_CODCIA"
Set PSTEMP_LLAVE = CN.CreateQuery("", pub_cadena)
Set temp_llave = PSTEMP_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
temp_llave.Requery

fila = 0
DoEvents
LIMPIA_DATOS
CABE_MAN
SQ_OPER = 2
PUB_TIPREG = 45
PUB_CODCIA = LK_CODCIA
LEER_TAB_LLAVE
fila = 0
Do Until tab_mayor.EOF
    PRE_ETIQUETA(fila) = Trim(tab_mayor!tab_NOMLARGO)
    fila = fila + 1
    tab_mayor.MoveNext
Loop
cmdimp.Visible = False
txtruc.MaxLength = LK_DIG_RUC

carga_venta

Exit Sub
SALE:
MsgBox "Depurar: " & Err.Description, 48, Pub_Titulo
Resume Next
End Sub

Private Sub forma_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       tiempo.SetFocus
    End If
End Sub


Private Sub grid_fac_EnterCell()
    textovar.Visible = False
    textovar.Text = Trim(grid_fac.TextMatrix(grid_fac.Row, grid_fac.COL))
    textovar.Left = ESTADO.Left + grid_fac.Left + grid_fac.CellLeft
    textovar.Width = grid_fac.CellWidth
    textovar.Height = grid_fac.CellHeight
    textovar.Top = ESTADO.Top + grid_fac.Top + grid_fac.CellTop '+ 1000 '
    If grid_fac.COL = 1 Or grid_fac.COL = 20 Then
        If Trim(grid_fac.TextMatrix(grid_fac.Row, 1)) = "" Or Trim(grid_fac.TextMatrix(grid_fac.Row, 20)) = "" Then
         textovar.Visible = True
         textovar.SetFocus
        End If
    End If
    If Trim(grid_fac.TextMatrix(grid_fac.Row, 1)) <> "" Then
        If Val(grid_fac.TextMatrix(grid_fac.Row, 12)) <> 0 Then
          stock.Caption = Format(Val(grid_fac.TextMatrix(grid_fac.Row, 15)) / Val(grid_fac.TextMatrix(grid_fac.Row, 12)), "0.00")
        End If
        unid.Caption = grid_fac.TextMatrix(grid_fac.Row, 16)
        nomarti.Caption = grid_fac.TextMatrix(grid_fac.Row, 0)
    Else
        stock.Caption = ""
        unid.Caption = ""
        nomarti.Caption = ""
    End If
End Sub

Private Sub grid_fac_KeyPress(KeyAscii As Integer)
Dim a As Integer
Dim t, WC
Dim wprecios As String * 12
Static CONS
Dim wactivo As Integer
If KeyAscii <> 13 Then Exit Sub
If grid_fac.rows <= 1 Then Exit Sub
'If grid_fac.COL = 1 Then Exit Sub
If grid_fac.COL >= 6 Then Exit Sub

If grid_fac.COL = 2 Then
 If Trim(grid_fac.TextMatrix(grid_fac.Row, 1)) = "" Then
    grid_fac.SetFocus
    Exit Sub
 End If
End If
If grid_fac.COL = 3 Then
    If Trim(grid_fac.TextMatrix(grid_fac.Row, 1)) = "" Then
     grid_fac.SetFocus
     Exit Sub
    End If
    UNIDAD.Left = grid_fac.Left + grid_fac.CellLeft
    UNIDAD.Width = grid_fac.CellWidth
    UNIDAD.Top = grid_fac.Top + grid_fac.CellTop
    SQ_OPER = 2
    pu_codcia = LK_CODCIA
    PUB_CODART = Val(grid_fac.TextMatrix(grid_fac.Row, 10))
    LEER_PRE_LLAVE
    UNIDAD.Clear
    UNIDAD.Visible = True
    wactivo = 0
    Do Until pre_mayor.EOF
     UNIDAD.AddItem Trim(pre_mayor!pre_unidad) & String(30, " ") & pre_mayor!pre_secuencia
     If pre_mayor!PRE_FLAG_UNIDAD <> "A" Then
       wactivo = pre_mayor.AbsolutePosition - 1
     End If
     pre_mayor.MoveNext
    Loop
    On Error GoTo pasa
    UNIDAD.ListIndex = 0 ''wactivo
    grid_fac.TextMatrix(grid_fac.Row, 4) = ""
    grid_fac.TextMatrix(grid_fac.Row, 13) = wactivo
    On Error GoTo 0
    UNIDAD.Visible = True
    UNIDAD.SetFocus
    ''SendKeys "%{up}"
     Exit Sub
End If
If grid_fac.COL = 4 Then
    If Trim(grid_fac.TextMatrix(grid_fac.Row, 1)) = "" Then
      grid_fac.SetFocus
      Exit Sub
    End If
    PRECIOS.Left = grid_fac.Left + grid_fac.CellLeft
    PRECIOS.Width = grid_fac.CellWidth + 600
    PRECIOS.Top = grid_fac.Top + grid_fac.CellTop

    SQ_OPER = 1
    pu_codcia = LK_CODCIA
    PUB_CODART = Val(grid_fac.TextMatrix(grid_fac.Row, 10))
    PUB_SECUEN = Val(Trim(Right(UNIDAD.Text, 3)))
    grid_fac.TextMatrix(grid_fac.Row, 18) = PUB_SECUEN
    LEER_PRE_LLAVE
    PRECIOS.Clear
    PRECIOS.Visible = True
    Do Until pre_llave.EOF
      If Left(moneda.Text, 1) = "S" Then
          wprecios = pre_llave!pre_pre1
          If Val(wprecios) <> 0 Then PRECIOS.AddItem Left(PRE_ETIQUETA(0), 8) & "= " & wprecios & String(60, " ") & "1"
          wprecios = pre_llave!PRE_PRE2
          If Val(wprecios) <> 0 Then PRECIOS.AddItem Left(PRE_ETIQUETA(1), 8) & "= " & wprecios & String(60, " ") & "1"
          'wprecios = pre_llave!PRE_PRE3
          'If Val(wprecios) <> 0 Then PRECIOS.AddItem Left(PRE_ETIQUETA(2), 8) & "= " & wprecios & String(60, " ") & "1"
          'wprecios = pre_llave!PRE_PRE4
          'If Val(wprecios) <> 0 Then PRECIOS.AddItem Left(PRE_ETIQUETA(3), 8) & "= " & wprecios & String(60, " ") & "1"
          'If LK_EMP <> "3AA" Then
          ' wprecios = pre_llave!PRE_PRE5
          ' If Val(wprecios) <> 0 Then PRECIOS.AddItem Left(PRE_ETIQUETA(4), 8) & "= " & wprecios & String(60, " ") & "1"
          'End If
          'wprecios = pre_llave!PRE_PRE6
          'If Val(wprecios) <> 0 Then PRECIOS.AddItem Left(PRE_ETIQUETA(5), 8) & "= " & wprecios & String(60, " ") & "1"
       Else
          wprecios = pre_llave!pre_pre11
          If Val(wprecios) <> 0 Then PRECIOS.AddItem Left(PRE_ETIQUETA(0), 8) & "= " & wprecios & String(60, " ") & "1"
          wprecios = pre_llave!PRE_PRE22
          If Val(wprecios) <> 0 Then PRECIOS.AddItem Left(PRE_ETIQUETA(0), 8) & "= " & wprecios & String(60, " ") & "1"
          wprecios = pre_llave!PRE_PRE33
          If Val(wprecios) <> 0 Then PRECIOS.AddItem Left(PRE_ETIQUETA(0), 8) & "= " & wprecios & String(60, " ") & "1"
          wprecios = pre_llave!PRE_PRE44
          If Val(wprecios) <> 0 Then PRECIOS.AddItem Left(PRE_ETIQUETA(0), 8) & "= " & wprecios & String(60, " ") & "1"
          If LK_EMP <> "3AA" Then
            wprecios = pre_llave!PRE_PRE55
            If Val(wprecios) <> 0 Then PRECIOS.AddItem Left(PRE_ETIQUETA(0), 8) & "= " & wprecios & String(60, " ") & "1"
          End If
          wprecios = pre_llave!PRE_PRE55
          If Val(wprecios) <> 0 Then PRECIOS.AddItem Left(PRE_ETIQUETA(0), 8) & "= " & wprecios & String(60, " ") & "1"
       End If
     pre_llave.MoveNext
    Loop
    On Error GoTo pasa
   ' If PRECIOS.ListCount <= 0 Then
    '  PRECIOS.Visible = False
     ' MsgBox "Definir precios....", 48, Pub_Titulo
      'grid_fac.COL = 1
      'grid_fac.SetFocus
      'Exit Sub
    'End If
    PRECIOS.ListIndex = 0
    On Error GoTo 0
    
    PRECIOS.Visible = True
    PRECIOS.SetFocus
    'SendKeys "%{up}"
     Exit Sub
End If
If grid_fac.COL = 5 Then
 If Trim(grid_fac.TextMatrix(grid_fac.Row, 1)) = "" Then
    grid_fac.SetFocus
    Exit Sub
 End If
End If


    textovar.Left = ESTADO.Left + grid_fac.Left + grid_fac.CellLeft
    textovar.Width = grid_fac.CellWidth
    textovar.Height = grid_fac.CellHeight
    textovar.Top = ESTADO.Top + grid_fac.Top + grid_fac.CellTop '+ 1000 '
    textovar.Text = grid_fac.TextMatrix(grid_fac.Row, grid_fac.COL)
    wfila_act = grid_fac.Row
    textovar.Visible = True
    Azul textovar, textovar
    textovar.SetFocus
Exit Sub
pasa:
Resume Next
End Sub

Private Sub grid_fac_KeyUp(KeyCode As Integer, Shift As Integer)
Dim WC
Dim a, WF As Integer
Dim tf, t, tC
Dim SALE As Boolean

If KeyCode = 46 Then
If grid_fac.rows <= 2 Then Exit Sub
If grid_fac.rows <= 3 Then
    pub_mensaje = MsgBox("Quitar el Producto para la Orden de Compra ", vbYesNo + vbExclamation + vbDefaultButton2, Pub_Titulo)
    If pub_mensaje = vbNo Then
      grid_fac.SetFocus
      Exit Sub
    End If
    CABE_MAN
Else
   pub_mensaje = MsgBox("Quitar el Producto para la Orden de Compra ", vbYesNo + vbExclamation + vbDefaultButton2, Pub_Titulo)
   If pub_mensaje = vbNo Then
     grid_fac.SetFocus
     Exit Sub
   Else
   '  grid_fac.RowHeight(grid_fac.Row) = 1
   grid_fac.RemoveItem (grid_fac.Row)
   grid_fac.Row = grid_fac.Row
   grid_fac.Refresh
   suma_grid
   grid_fac.SetFocus
   End If
End If
End If
'grid_fac.SetFocus
Exit Sub



End Sub



Private Sub i_condi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Left(Trim(Right(Trim(i_condi.Text), 3)), 1) = "1" And Trim(Right(Trim(i_condi.Text), 2)) = "CC" Then ' FA
           'i_dias.Text = 0
           'i_dias.Locked = False
           'Azul i_dias, i_dias
           
           i_dias.Text = "0"
           i_dias.Locked = True
           i_dias_KeyPress 13
        Else
           i_dias.Locked = False
        End If
        txtCli.SetFocus
    End If
End Sub

Private Sub i_condi_LostFocus()
PUB_CODTRA = 2401
PUB_SECUENCIA = Trim(Left(i_condi.Text, 2))
SQ_OPER = 1
LEER_SUT_LLAVE
pub_signo_car = SUT_LLAVE!SUT_SIGNO_CAR
End Sub

Private Sub i_destino_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
     grid_fac_EnterCell
     textovar.Visible = False
     grid_fac.SetFocus
    End If
End Sub

Private Sub i_dias_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 i_fbg.SetFocus
 SendKeys "%{up}"
End If
End Sub

Private Sub i_fbg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 grid_fac.SetFocus
 SendKeys "%{up}"
End If

End Sub

Private Sub ListView1_DblClick()
' loc_key = ListView1.SelectedItem.Index
' TEXTOVAR.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
' TEXTOVAR_KeyPress 13
End Sub

Private Sub ListView1_GotFocus()
'If loc_key <> 0 Then
' Set ListView1.SelectedItem = ListView1.ListItems(loc_key)
' ListView1.ListItems.Item(loc_key).Selected = True
' ListView1.ListItems.Item(loc_key).EnsureVisible
'End If

End Sub

Private Sub ListView1_ItemClick(ByVal item As ComctlLib.ListItem)
'If loc_key <> 0 Then
' loc_key = ListView1.SelectedItem.Index
' TEXTOVAR.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
'End If
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
'If KeyAscii = 27 Then
' ListView1.Visible = False
' TEXTOVAR.Text = ""
' TEXTOVAR.SetFocus
' Exit Sub
'End If
'If KeyAscii <> 13 Then
' Exit Sub
'End If
'ListView1_DblClick
End Sub

Private Sub ListView1_LostFocus()
ListView1.Visible = False
End Sub

Private Sub moneda_Click()
If Not cmdIngreso.Enabled Then Exit Sub
For fila = 2 To grid_fac.rows - 1
     PUB_CODART = Val(grid_fac.TextMatrix(fila, 10))
     If PUB_CODART > 0 Then
       pu_codcia = LK_CODCIA
       PUB_SECUENCIA = Val(grid_fac.TextMatrix(fila, 11))
       SQ_OPER = 1
       LEER_PRE_LLAVE
       If Left(moneda.Text, 1) = "S" Then
          If Val(grid_fac.TextMatrix(fila, 14)) = 1 Then grid_fac.TextMatrix(fila, 13) = Val(pre_llave!pre_pre1)
          If Val(grid_fac.TextMatrix(fila, 14)) = 2 Then grid_fac.TextMatrix(fila, 13) = Val(pre_llave!PRE_PRE2)
          If Val(grid_fac.TextMatrix(fila, 14)) = 3 Then grid_fac.TextMatrix(fila, 13) = Val(pre_llave!PRE_PRE3)
          If Val(grid_fac.TextMatrix(fila, 14)) = 4 Then grid_fac.TextMatrix(fila, 13) = Val(pre_llave!PRE_PRE4)
          If Val(grid_fac.TextMatrix(fila, 14)) = 5 Then grid_fac.TextMatrix(fila, 13) = Val(pre_llave!PRE_PRE5)
       Else
          If Val(grid_fac.TextMatrix(fila, 14)) = 1 Then grid_fac.TextMatrix(fila, 13) = Val(pre_llave!pre_pre11)
          If Val(grid_fac.TextMatrix(fila, 14)) = 2 Then grid_fac.TextMatrix(fila, 13) = Val(pre_llave!PRE_PRE22)
          If Val(grid_fac.TextMatrix(fila, 14)) = 3 Then grid_fac.TextMatrix(fila, 13) = Val(pre_llave!PRE_PRE33)
          If Val(grid_fac.TextMatrix(fila, 14)) = 4 Then grid_fac.TextMatrix(fila, 13) = Val(pre_llave!PRE_PRE44)
          If Val(grid_fac.TextMatrix(fila, 14)) = 5 Then grid_fac.TextMatrix(fila, 13) = Val(pre_llave!PRE_PRE55)
       End If
       grid_fac.TextMatrix(fila, 4) = redondea(Val(grid_fac.TextMatrix(fila, 13)) * (100 - Val(grid_fac.TextMatrix(fila, 5))) / 100)
     End If
Next fila
If grid_fac.rows <> 2 Then suma_grid


If Left(moneda.Text, 1) = "S" Then
 i_moneda.Caption = "S/."
 grid_fac.TextMatrix(1, 4) = "S/."
Else
 i_moneda.Caption = "US$."
 grid_fac.TextMatrix(1, 4) = "US$."
End If
End Sub

Private Sub moneda_GotFocus()
If moneda.ListCount = 1 Then moneda_KeyPress 13
End Sub

Private Sub moneda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Azul i_dias, i_dias
End If
End Sub


Private Sub PRECIOS_GotFocus()
On Error GoTo SALE
grid_fac.TextMatrix(grid_fac.Row, 13) = Format(Val(Mid(PRECIOS.Text, 10, Len(Trim(PRECIOS.Text)) - 10)), "0.00")
grid_fac.TextMatrix(grid_fac.Row, 14) = Val(Right(PRECIOS.Text, 3))
SALE:
Exit Sub
End Sub

Private Sub PRECIOS_KeyPress(KeyAscii As Integer)
On Error Resume Next

If KeyAscii = 27 Then
 PRECIOS.Visible = False
 grid_fac.SetFocus
End If
If KeyAscii <> 13 Then Exit Sub
'SQ_OPER = 1
'pu_codcia = LK_CODCIA
'PUB_CODART = Val(grid_fac.TextMatrix(grid_fac.Row, 10))
'PUB_SECUEN = Val(Right(UNIDAD.Text, 4))
'LEER_PRE_LLAVE

grid_fac.TextMatrix(grid_fac.Row, 4) = Format(Val(Mid(PRECIOS.Text, 10, Len(Trim(PRECIOS.Text)) - 10)), "0.00")
grid_fac.TextMatrix(grid_fac.Row, 13) = Format(Val(Mid(PRECIOS.Text, 10, Len(Trim(PRECIOS.Text)) - 10)), "0.00")
grid_fac.TextMatrix(grid_fac.Row, 14) = Val(Right(PRECIOS.Text, 3))

PRECIOS.Visible = False
suma_grid
textovar.Visible = False
'If Trim(grid_fac.TextMatrix(grid_fac.Rows - 1, 1)) <> "" Then
'  grid_fac.Rows = grid_fac.Rows + 1
'  grid_fac.RowHeight(grid_fac.Rows - 1) = 285
'  grid_fac.Row = grid_fac.Rows - 1
'Else
' If grid_fac.Row < grid_fac.Rows - 1 Then
'    grid_fac.Row = grid_fac.Row + 1
' End If
'End If
grid_fac.COL = 5
textovar.Visible = True
textovar.SetFocus
Exit Sub
grid_fac.COL = 5
grid_fac_KeyPress 13


End Sub

Private Sub PRECIOS_KeyUp(KeyCode As Integer, Shift As Integer)
Dim ww As String
Dim wpre As Currency
If KeyCode = 45 Then
'seIf grid_fac.COL = 6 Then
ww = InputBox("Digite Precios :", "Ingreso de Precio", "0")
wpre = Val(ww)
grid_fac.TextMatrix(grid_fac.Row, 4) = wpre
' Comentado 05052004
If Val(grid_fac.TextMatrix(grid_fac.Row, 13)) <> 0 Then
  'grid_fac.TextMatrix(grid_fac.Row, 5) = redondea((Val(grid_fac.TextMatrix(grid_fac.Row, 13)) - wpre) * 100 / Val(Val(grid_fac.TextMatrix(grid_fac.Row, 13))))
  'grid_fac.TextMatrix(grid_fac.Row, 5) = 0
End If
PRECIOS.Visible = False
suma_grid
grid_fac.COL = grid_fac.COL + 1 ' Agregado
grid_fac.SetFocus






End If
End Sub

Private Sub salir_Click()
Unload FORM_COT
End Sub


Public Sub LIMPIA_DATOS()
LBLSIT.Caption = ""
grid_fac.Enabled = True
lblcli.Caption = ""
txtatte.Text = ""

txtCli.Text = ""
txtruc.Text = ""
'tserie.Text = ""
'txtdoc.Text = ""
grid_fac.Clear

txtigv.Text = ""
txtvalorv.Text = ""
txttotal.Text = ""
textovar.Visible = False
stock.Caption = ""
unid.Caption = ""
nomarti.Caption = ""
oferta.Text = ""
forma.Text = ""
tiempo.Text = ""
i_destino.Clear
End Sub

Public Sub CABE_MAN()
grid_fac.Cols = 21
grid_fac.rows = 2
grid_fac.Clear
fila = 0
grid_fac.ColWidth(0) = 3600 ' nombre arti
grid_fac.ColWidth(1) = 900 ' codigo arti
grid_fac.ColWidth(2) = 800 ' cantidad
grid_fac.ColWidth(3) = 900 ' unidad
grid_fac.ColWidth(4) = 900 ' precio
grid_fac.ColWidth(5) = 500  ' decto. %
grid_fac.ColWidth(6) = 1000 ' sub total
grid_fac.ColWidth(7) = 0  ' peso
grid_fac.ColWidth(8) = 0
grid_fac.ColWidth(9) = 0
grid_fac.ColWidth(10) = 0 '  COD ORIGINAL
grid_fac.ColWidth(11) = 0 '  PRE_SECUENCIA
grid_fac.ColWidth(12) = 0 '  PRE_EQUIV
grid_fac.ColWidth(13) = 0 '  PRE_PRECIO COLOCADO
grid_fac.ColWidth(14) = 0 '  numero de PRE_PRECIO
grid_fac.ColWidth(15) = 0 '  numero de arm_stock
grid_fac.ColWidth(16) = 0 '  numero de pre_unidad
grid_fac.ColWidth(17) = 0
grid_fac.ColWidth(18) = 0 ' NUMERO DE SECUENCIA EN PRECIOS
'agregado por mic
grid_fac.ColWidth(19) = 0  'descuento en cantidad
grid_fac.ColWidth(20) = 0 'COLORES



grid_fac.TextMatrix(0, 0) = "Articulo"
grid_fac.TextMatrix(0, 1) = "Codigo"
grid_fac.TextMatrix(0, 2) = "Cantidad"
grid_fac.TextMatrix(0, 3) = "Unidad"
grid_fac.TextMatrix(0, 4) = "Precios"
grid_fac.TextMatrix(0, 5) = "Dscto"
grid_fac.TextMatrix(1, 5) = "  (%)"
grid_fac.TextMatrix(0, 6) = "Sub Total"
grid_fac.TextMatrix(0, 7) = "Peso(Kg)"
grid_fac.TextMatrix(0, 8) = ""
grid_fac.TextMatrix(0, 9) = ""
grid_fac.TextMatrix(0, 19) = "Descto"
grid_fac.TextMatrix(1, 19) = "Valor"
grid_fac.TextMatrix(1, 20) = ""
grid_fac.RowHeight(1) = 320


End Sub
Public Sub suma_grid()
'On Error GoTo SALE
Dim RES_DEUDA As Currency
Dim wsumadol As Currency
Dim WSIMBOL As String
Dim WTC As Currency
Dim WF As Integer
Dim DesctoTotal As Double

WF = 2
Dim fx As Integer
Dim wcantid As Currency
Dim wpeso As Currency
fx = 1
SUM_H = 0
SUM_D = 0
wcantid = 0
Do While fx = 1
    'If Left(grid_fac.TextMatrix(WF, 0), 1) <> "T" Then
      SUM_D = SUM_D + Val(grid_fac.TextMatrix(WF, 4))
      SUM_H = SUM_H + Val(Val(grid_fac.TextMatrix(WF, 2)) * Val(grid_fac.TextMatrix(WF, 4)))
      wcantid = wcantid + Val(grid_fac.TextMatrix(WF, 2))
      wpeso = wpeso + Val(grid_fac.TextMatrix(WF, 7))
      
      grid_fac.TextMatrix(WF, 6) = Format(Val(grid_fac.TextMatrix(WF, 2)) * Val(grid_fac.TextMatrix(WF, 4)), "0.00")
      Call FactorDescto(WF)
      grid_fac.TextMatrix(WF, 19) = Format(FACTOR_DESCTO * Val(grid_fac.TextMatrix(WF, 6)) / 100, "#0.00")
      DesctoTotal = DesctoTotal + Val(grid_fac.TextMatrix(WF, 19))
    'End If
    WF = WF + 1
    If WF = grid_fac.rows Then
        fx = 0
    Else
        If Trim(grid_fac.TextMatrix(WF, 0)) = "" Then fx = 0
    End If
Loop
   fila = WF - 1
   grid_fac.TextMatrix(1, 0) = "Totales = "
   grid_fac.TextMatrix(1, 6) = Format(SUM_H, "####0.00")
   grid_fac.TextMatrix(1, 2) = Format(wcantid, "####0.00")
   grid_fac.TextMatrix(1, 7) = Format(wpeso, "####0.00")
   
   
   WS_NETO = SUM_H
   'txtvalorv.Text = Format(SUM_H, "####0.00") 'quite / ((100 + LK_IGV) / 100
   'txtigv.Text = Format((SUM_H - DesctoTotal) * (LK_IGV) / 100, "#####0.00")
   'txtdescto.Text = Format(DesctoTotal, "#####0.00")
   'txttotal.Text = Format((SUM_H - DesctoTotal + Val(txtigv.Text)), "#####0.00")
   Dim SUM_H1 As Double, DesctoTotal1 As Double
   
   SUM_H1 = SUM_H / ((100 + LK_IGV) / 100)
   DesctoTotal1 = DesctoTotal / ((100 + LK_IGV) / 100)
   
   txtvalorv.Text = Format(SUM_H1, "####0.00")  'quite / ((100 + LK_IGV) / 100
   txtdescto.Text = Format(DesctoTotal1, "#####0.00")
   txtigv.Text = Format((SUM_H1 - DesctoTotal1) * (LK_IGV) / 100, "#####0.00")
   txttotal.Text = Format((SUM_H1 - DesctoTotal1 + Val(txtigv.Text)), "#####0.00")
   
'If cli_llave.EOF Then Exit Sub
'If SUT_LLAVE!SUT_SIGNO_CAR <> 1 Then Exit Sub
'If Nulo_Valor0(SUT_LLAVE!SUT_FLAG_CC) = 1 Then Exit Sub
WTC = 1
If Left(moneda.Text, 1) = "D" Then
 PUB_CAL_INI = LK_FECHA_DIA
 PUB_CAL_FIN = LK_FECHA_DIA
 pu_codcia = LK_CODCIA
 PUB_CODCIA = LK_CODCIA
 SQ_OPER = 1
 LEER_CAL_LLAVE
 WTC = 0
 WSIMBOL = ""
 If Not cal_llave.EOF Then
   WTC = Nulo_Valor0(cal_llave!cal_tipo_cambio)
 End If
 If WTC = 0 Then
   MsgBox "Ingresar el Tipo de Cambio", 48, Pub_Titulo
   GoTo PA
 End If
End If
If Left(moneda.Text, 1) = "S" Then
  WSIMBOL = "S/."
 ' wsumadol = Val(Nulo_Valor0(cli_llave!cli_limcre))
 ' RES_DEUDA = pub_deuda
Else
 WSIMBOL = "US$."
 'wsumadol = Val(redondea(Nulo_Valor0(cli_llave!cli_limcre) / WTC)) + Val(redondea(Val(Nulo_Valor0(cli_llave!cli_limcre2))))
 'RES_DEUDA = redondea(Val(pub_deuda / WTC))
End If
'lblcred.Caption = Format(Nulo_Valor0(cli_llave!cli_limcre), "#,##0.00")
'lblDeuda.Caption = Format(RES_DEUDA + WS_NETO, "#,##0.00")
'lbldisp.Caption = Format(wsumadol - RES_DEUDA - WS_NETO, "##,###,###.00")

PA:
'If LK_FLAG_EXED = "A" Then
'   If Trim(Left(moneda.Text, 1)) = "S" Then
'     If (RES_DEUDA + WS_NETO) > Nulo_Valor0(cli_llave!cli_limcre) Then
'         MsgBox "El Monto del Pedido supera el Credito... ", 48, Pub_Titulo
'         If TEXTOVAR.Visible Then TEXTOVAR.SetFocus
'     End If
'   Else
'     If (RES_DEUDA + WS_NETO) > Nulo_Valor0(cli_llave!cli_limcre2) Then
'         'If ws_bruto_bak <> WS_BRUTO Then MsgBox "Credito Excedido... "
'      End If
'   End If
'End If
 
  
Exit Sub
SALE:
cancelar_Click
'MsgBox "Verficar Importe.", 48, Pub_Titulo
'Resume Next
'If TEXTOVAR.Visible Then Azul3 TEXTOVAR, TEXTOVAR
End Sub
Public Sub suma_subtotal()
If WMODO = "I" Then Exit Sub

Dim WF As Integer
Dim WFIN As Integer
Dim WINI As Integer

Dim fx As Integer
Exit Sub
End Sub
Private Sub MuestrPrecios()
  SQ_OPER = 2
  LEER_PRE_LLAVE
  If Not pre_mayor.EOF Then
    ''If i_ds.Text = "S" Then 'Soles'
        lblprecio(0).Caption = Format(pre_mayor("PRE_PRE1"), "0.00")
        lblprecio(1).Caption = Format(pre_mayor("PRE_PRE2"), "0.00")
        lblprecio(2).Caption = Format(pre_mayor("PRE_PRE3"), "0.0000")
        lblprecio(3).Caption = Format(pre_mayor("PRE_PRE4"), "0.0000")
        lblprecio(4).Caption = Format(pre_mayor("PRE_PRE5"), "0.0000")
        lblprecio(10).Caption = Format(pre_mayor("PRE_PRE6"), "0.0000")
    ''Else 'Dolares'
        lblprecio(5).Caption = Format(pre_mayor("PRE_PRE11"), "0.0000")
        lblprecio(6).Caption = Format(pre_mayor("PRE_PRE22"), "0.0000")
        lblprecio(7).Caption = Format(pre_mayor("PRE_PRE33"), "0.0000")
        lblprecio(8).Caption = Format(pre_mayor("PRE_PRE44"), "0.0000")
        lblprecio(9).Caption = Format(pre_mayor("PRE_PRE55"), "0.0000")
        lblprecio(11).Caption = Format(pre_mayor("PRE_PRE66"), "0.0000")
    ''End If
  End If
End Sub


Private Sub Consistencias(wsGrid As MSFlexGrid, wsTexto As TextBox, wsKeyAscii As Integer, Optional ConsisVal, Optional ConsisCol)
  Static valor
  Dim car As String
 ' NUMEROS CON DECIMALES
    car = Chr$(wsKeyAscii)
    car = UCase$(Chr$(wsKeyAscii))
    wsKeyAscii = Asc(car)
    If wsKeyAscii = 45 Then
      If wsTexto.Text <> "" Then
         Beep
         wsKeyAscii = 0
         Exit Sub
      End If
    End If
    If wsKeyAscii = 46 Then
      If InStr(1, wsTexto.Text, ".") <> 0 Then
        Beep
        wsKeyAscii = 0
        Exit Sub
      End If
    End If
    
    If car < "0" Or car > "9" Then
      If wsKeyAscii <> 8 And wsKeyAscii <> 13 And car <> "." Then
          wsKeyAscii = 0
          Beep
          Exit Sub
        End If
    End If

End Sub

Public Sub CABE_ING()
grid_fac.Cols = 6
grid_fac.rows = 3
grid_fac.Clear
grid_fac.MergeCells = 4
grid_fac.MergeCol(0) = True
grid_fac.MergeCol(1) = True
grid_fac.MergeCol(2) = True
grid_fac.MergeCol(3) = True
grid_fac.MergeCol(4) = False
grid_fac.MergeCol(5) = False
grid_fac.MergeRow(2) = False
grid_fac.RowHeight(0) = 285
grid_fac.RowHeight(1) = 285
grid_fac.RowHeight(2) = 285

fila = 0
grid_fac.ColWidth(0) = 400
grid_fac.ColWidth(1) = 1400
grid_fac.ColWidth(2) = 2500
grid_fac.ColWidth(3) = 0
grid_fac.ColWidth(4) = 1500
grid_fac.ColWidth(5) = 1500

grid_fac.TextMatrix(0, 0) = "Item"
grid_fac.TextMatrix(0, 1) = "Cuenta"
grid_fac.TextMatrix(0, 2) = "Descripcion"
grid_fac.TextMatrix(0, 3) = "Glosa"
grid_fac.TextMatrix(0, 4) = "Debe"
grid_fac.TextMatrix(0, 5) = "Haber"
grid_fac.TextMatrix(1, 0) = "Item"
grid_fac.TextMatrix(1, 1) = "Cuenta"
grid_fac.TextMatrix(1, 2) = "Descripcion"
grid_fac.TextMatrix(1, 3) = "Glosa"

'grid_fac.MergeCol
'grid_fac.MergeRow(2) = True



End Sub

Private Sub siguiente_Click()
 txtdoc.Text = Val(txtdoc.Text) + 1
 PUB_NUMSER = Val(tserie.Text)
 PUB_NUMFAC = Val(txtdoc.Text)
 LLENA_DOCU

End Sub

Private Sub t_direc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Fracli.Visible = False
  lblcli.Caption = Trim(t_nombre.Text) & " - " & Trim(t_doc.Text) & " - " & Trim(t_direc.Text)
  'NUMERO = FORM_COT.txtcli.WhatsThisHelpID
  'cli_llave.Edit
  'cli_llave!CLI_NOMBRE = Trim(t_nombre.Text)
  'cli_llave!CLI_NOMBRE_ESPOSO = Trim(t_nombre.Text)
  'cli_llave!cli_RUC_ESPOSA = Trim(t_doc.Text)
  'cli_llave!CLI_CASA_DIREC = Trim(t_direc.Text)
  'cli_llave.Update
  avanza_campo
End If
If KeyAscii = 27 Then
  txtCli.Text = 0
  Fracli.Visible = False
  txtCli.SetFocus
End If
txt_key.SetFocus
End Sub

Private Sub t_doc_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then Azul t_direc, t_direc
If KeyAscii = 27 Then
  txtCli.Text = 0
  Fracli.Visible = False
  txtCli.SetFocus
End If
End Sub

Private Sub t_nombre_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then Azul t_doc, t_doc
If KeyAscii = 27 Then
  txtCli.Text = 0
  Fracli.Visible = False
  txtCli.SetFocus
End If

End Sub

Private Sub textovar_Change()
If Not textovar.Visible Then Exit Sub
'BLOQUEADO POR MIC
'If grid_fac.COL = 5 Then
' grid_fac.TextMatrix(grid_fac.Row, 4) = redondea(Val(grid_fac.TextMatrix(grid_fac.Row, 13)) * (100 - Val(textovar.Text)) / 100)
'End If
If grid_fac.COL = 1 Then
    grid_fac.TextMatrix(grid_fac.Row, 0) = ""
    grid_fac.TextMatrix(grid_fac.Row, 0) = ""
    grid_fac.TextMatrix(grid_fac.Row, 1) = ""
    grid_fac.TextMatrix(grid_fac.Row, 2) = ""
    grid_fac.TextMatrix(grid_fac.Row, 3) = ""
    grid_fac.TextMatrix(grid_fac.Row, 4) = ""
    grid_fac.TextMatrix(grid_fac.Row, 5) = ""
    grid_fac.TextMatrix(grid_fac.Row, 6) = ""
    grid_fac.TextMatrix(grid_fac.Row, 7) = ""
    grid_fac.TextMatrix(grid_fac.Row, 8) = ""
    grid_fac.TextMatrix(grid_fac.Row, 9) = ""
    grid_fac.TextMatrix(grid_fac.Row, 10) = ""
    grid_fac.TextMatrix(grid_fac.Row, 11) = ""
    grid_fac.TextMatrix(grid_fac.Row, 12) = ""
    grid_fac.TextMatrix(grid_fac.Row, 14) = ""
    grid_fac.TextMatrix(grid_fac.Row, 15) = ""
    grid_fac.TextMatrix(grid_fac.Row, 16) = ""
    grid_fac.TextMatrix(grid_fac.Row, 17) = ""
    grid_fac.Text = textovar.Text
    stock.Caption = ""
    unid.Caption = ""
    nomarti.Caption = ""
    suma_grid
Else
 If grid_fac.COL = 2 Then
  grid_fac.Text = textovar.Text
 Else
  grid_fac.Text = Format(textovar.Text, "0.00")
 End If
 suma_grid
 suma_subtotal
End If
End Sub

Private Sub TEXTOVAR_GotFocus()
'temporal = grid_fac.TextMatrix(grid_fac.Row, grid_fac.COL)
End Sub

Private Sub textovar_KeyDown(KeyCode As Integer, Shift As Integer)

' busca arti
If Not ListView1.Visible Then
If KeyCode = 40 Then  ' flecha abajo
  If grid_fac.Row = grid_fac.rows - 1 Then Exit Sub
  If Trim(grid_fac.Text) <> "" Then Exit Sub
  grid_fac.Row = grid_fac.Row + 1
  grid_fac.SetFocus
  Exit Sub
End If
If KeyCode = 38 Then
 If Trim(grid_fac.Text) <> "" Then Exit Sub
 grid_fac.Row = grid_fac.Row - 1
 grid_fac.SetFocus
 Exit Sub
End If
If KeyCode = 39 Then
If Trim(grid_fac.Text) <> "" Then Exit Sub
 grid_fac.COL = grid_fac.COL + 1
 grid_fac.SetFocus
 Exit Sub
End If
End If
If grid_fac.COL <> 1 Then Exit Sub
Dim strFindMe As String
Dim itmFound As ListItem    ' Variable FoundItem.
If Not ListView1.Visible Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And textovar.Text = "" Then
  loc_key = 1
  Set ListView1.SelectedItem = ListView1.ListItems(loc_key)
'  LISTVIEW1.Visible = False
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
'  KeyCode = 0
  ListView1.ListItems.item(loc_key).Selected = True
  ListView1.ListItems.item(loc_key).EnsureVisible
  textovar.Text = Trim(ListView1.ListItems.item(loc_key).Text) & " "
  DoEvents
  textovar.SelStart = Len(textovar.Text)
  DoEvents
fin:

End Sub

Private Sub textovar_KeyPress(KeyAscii As Integer)
'SOLO_DECIMAL TEXTOVAR, KeyAscii
If KeyAscii = 27 Then
  If textovar.Text = "" Then
    textovar.Visible = False
    grid_fac.SetFocus
    Exit Sub
  End If
  textovar.Text = "" ' temporal
  'TEXTOVAR.Visible = False
  'grid_fac.SetFocus
  ListView1.Visible = False
  Exit Sub
End If
If grid_fac.COL = 2 Or grid_fac.COL = 4 Then Consistencias grid_fac, textovar, KeyAscii  'Or grid_fac.COL = 5
'If grid_fac.COL = 5 Then Consistencias grid_fac, textovar, KeyAscii, 10, 42
If KeyAscii <> 13 Then Exit Sub

If grid_fac.COL = 2 Then
 If Trim(textovar.Text) = "" Then Exit Sub
 textovar.Visible = False
 'If Val(arm_llave!ARM_STOCK) - Val(grid_fac.TextMatrix(grid_fac.Row, 2)) < 0 Then
 '     MsgBox "Stock es :" & Format(arm_llave!ARM_STOCK, "0.00") & "  /  Aplicando la cantidad : " & Format(Val(arm_llave!ARM_STOCK) - Val(grid_fac.TextMatrix(grid_fac.Row, 2)), "0.00"), 48, Pub_Titulo
 'End If
 grid_fac.COL = 3
 If Trim(grid_fac.Text) <> "" Then
   grid_fac.SetFocus
   Exit Sub
 End If
 grid_fac_KeyPress 13
 Exit Sub
End If
If grid_fac.COL = 6 Then
' grid_fac.TextMatrix(grid_fac.Row, 6) = textovar.Text
' suma_grid
' textovar.Visible = False
' grid_fac.SetFocus
' Exit Sub
End If
If grid_fac.COL = 5 Then
 textovar.Visible = False
 If Trim(grid_fac.TextMatrix(grid_fac.rows - 1, 1)) <> "" Then
   grid_fac.rows = grid_fac.rows + 1
   grid_fac.RowHeight(grid_fac.rows - 1) = 285
   grid_fac.Row = grid_fac.rows - 1
 Else
  If grid_fac.Row < grid_fac.rows - 1 Then
     grid_fac.Row = grid_fac.Row + 1
  End If
 End If
 grid_fac.COL = 1
 textovar.Visible = True
 textovar.SetFocus
 Exit Sub
End If




If grid_fac.COL <> 1 Then Exit Sub

Dim valor As String
Dim tf As Integer
Dim i, car
Dim itmFound As ListItem
car = Chr(KeyAscii)
KeyAscii = Asc(UCase(car))
If KeyAscii = 27 Then
 ListView1.Visible = False
 textovar.Text = ""
End If
If KeyAscii <> 13 Then
   GoTo fin
End If
VAR_ACTIVAR = 0
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
  PUB_KEY = 0
Else
 On Error GoTo mucho
 PUB_KEY = Val(textovar.Text)
 On Error GoTo 0
 If Len(textovar.Text) = 0 Then
    Exit Sub
 End If
 If IsNumeric(textovar.Text) = False Then
   PUB_KEY = 0
 End If
End If

If PUB_KEY <> 0 Then
    SQ_OPER = 1
    PUB_KEY = textovar.Text
    pu_codcia = LK_CODCIA
    LEER_ART_LLAVE
    If art_LLAVE.EOF Then
       MsgBox "Codigo NO Existe.", 48, Pub_Titulo
       Azul textovar, textovar
       GoTo fin
    End If
    If art_LLAVE!art_flag_stock <> "P" Then
       'MsgBox "Producto no es Mercaderia.", 48, Pub_Titulo
       'Azul textovar, textovar
       'GoTo fin
    End If
    WCOD_ORIGINAL = art_LLAVE!ART_KEY
    SQ_OPER = 1
    pu_codcia = LK_CODCIA
    PUB_CODART = WCOD_ORIGINAL
    LEER_ARM_LLAVE
    SQ_OPER = 1
    pu_codcia = LK_CODCIA
    PUB_SECUEN = 0
    LEER_PRE_LLAVE
    grid_fac.TextMatrix(grid_fac.Row, 16) = pre_llave!pre_unidad
    grid_fac.TextMatrix(grid_fac.Row, 15) = arm_llave!ARM_STOCK
    grid_fac.TextMatrix(grid_fac.Row, 12) = pre_llave!PRE_EQUIV
    grid_fac.TextMatrix(grid_fac.Row, 11) = pre_llave!pre_secuencia
    grid_fac.TextMatrix(grid_fac.Row, 0) = art_LLAVE!ART_NOMBRE
    grid_fac.TextMatrix(grid_fac.Row, 10) = art_LLAVE!ART_KEY
    ListView1.Visible = False
    textovar.Visible = False
    grid_fac.COL = 2
    If Trim(grid_fac.Text) <> "" Then
      grid_fac.SetFocus
      Exit Sub
    End If
    textovar.Visible = True
    textovar.SetFocus
    Exit Sub
Else
  If ListView1.Visible = False And VAR_ACTIVAR <> 99 And textovar.Text <> "" And LK_FLAG_ORIGINAL <> "A" And LK_FLAG_ALTERNO = "A" Then
IR_ALTERNO:
     SQ_OPER = 3
     pu_alterno = textovar.Text
     pu_codcia = LK_CODCIA
     LEER_ART_LLAVE
     If art_llave_alt.EOF Then
       MsgBox "Codigo No Existe ...", 48, Pub_Titulo
       Azul textovar, textovar
       Exit Sub
     End If
     If art_llave_alt!art_flag_stock <> "M" Then
       MsgBox "Producto no es Mercaderia.", 48, Pub_Titulo
       Azul textovar, textovar
       GoTo fin
     End If
     ListView1.Visible = False
     WCOD_ORIGINAL = art_llave_alt!ART_KEY
     SQ_OPER = 1
     pu_codcia = LK_CODCIA
     PUB_CODART = WCOD_ORIGINAL
     LEER_ARM_LLAVE
     SQ_OPER = 1
     pu_codcia = LK_CODCIA
     PUB_SECUEN = 0
     LEER_PRE_LLAVE
     grid_fac.TextMatrix(grid_fac.Row, 16) = pre_llave!pre_unidad
     grid_fac.TextMatrix(grid_fac.Row, 15) = arm_llave!ARM_STOCK
     grid_fac.TextMatrix(grid_fac.Row, 12) = pre_llave!PRE_EQUIV
     grid_fac.TextMatrix(grid_fac.Row, 11) = pre_llave!pre_secuencia
    
     grid_fac.TextMatrix(grid_fac.Row, 0) = art_llave_alt!ART_NOMBRE
     grid_fac.TextMatrix(grid_fac.Row, 10) = art_llave_alt!ART_KEY
     textovar.Visible = False
     ListView1.Visible = False
     grid_fac.COL = 2
     If Trim(grid_fac.Text) <> "" Then
       grid_fac.SetFocus
       Exit Sub
     End If
     textovar.Visible = True
     Azul textovar, textovar
     Exit Sub
  Else
    If loc_key > ListView1.ListItems.count Or loc_key = 0 Then
     Exit Sub
    End If
    valor = UCase(ListView1.ListItems.item(loc_key).Text)
    If Trim(UCase(textovar.Text)) = Left(valor, Len(Trim(textovar.Text))) And Len(Trim(textovar.Text)) <> 0 Then
      If VAR_ACTIVAR = 0 And LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
        textovar.Text = Trim(ListView1.ListItems.item(loc_key))
        GoTo IR_ALTERNO
      End If
      If VAR_ACTIVAR <> 99 Then
       textovar.Text = Trim(ListView1.ListItems.item(loc_key).SubItems(1))
      Else
       textovar.Text = Trim(ListView1.ListItems.item(loc_key))
      End If
      SQ_OPER = 1
      pu_codcia = LK_CODCIA
      If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
       PUB_KEY = Val(ListView1.ListItems.item(loc_key).SubItems(1))
      Else
       PUB_KEY = textovar.Text
      End If
      LEER_ART_LLAVE
      VAR_ACTIVAR = 0
      If art_LLAVE.EOF Then
        MsgBox "Codigo No Existe ...", 48, Pub_Titulo
        Azul3 textovar, textovar
        Exit Sub
      End If
      If art_LLAVE!art_flag_stock <> "P" Then
       'MsgBox "Producto no es Mercaderia.", 48, Pub_Titulo
       'Azul3 textovar, textovar
      ' GoTo fin
      End If
      WCOD_ORIGINAL = art_LLAVE!ART_KEY
      SQ_OPER = 1
      pu_codcia = LK_CODCIA
      PUB_CODART = WCOD_ORIGINAL
      LEER_ARM_LLAVE
      SQ_OPER = 1
      pu_codcia = LK_CODCIA
      PUB_SECUEN = 0
      LEER_PRE_LLAVE
      grid_fac.TextMatrix(grid_fac.Row, 16) = pre_llave!pre_unidad
      grid_fac.TextMatrix(grid_fac.Row, 15) = arm_llave!ARM_STOCK
      grid_fac.TextMatrix(grid_fac.Row, 12) = pre_llave!PRE_EQUIV
      grid_fac.TextMatrix(grid_fac.Row, 11) = pre_llave!pre_secuencia
      ListView1.Visible = False
      grid_fac.TextMatrix(grid_fac.Row, 0) = art_LLAVE!ART_NOMBRE
      grid_fac.TextMatrix(grid_fac.Row, 10) = art_LLAVE!ART_KEY
      grid_fac.COL = 2
      If Trim(grid_fac.Text) <> "" Then
        grid_fac.SetFocus
        Exit Sub
      End If
      textovar.Visible = True
      textovar.SetFocus
     
      Exit Sub
    Else
      Exit Sub
    End If
    
  End If
End If
dale:
ListView1.Visible = False
fin:
mucho:
Exit Sub
ERROR_CODIGO:
MsgBox "Codigo NO Valido .... ", 48, Pub_Titulo
Azul3 textovar, textovar
  

Exit Sub

End Sub

Private Sub textovar_KeyUp(KeyCode As Integer, Shift As Integer)
If grid_fac.COL <> 1 Then Exit Sub
' busca arti
Dim VAR
''If KeyCode = 13 Then Exit Sub ' Comentado
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
  If Len(textovar.Text) = 0 Or Trim(textovar.Text) = "" Then
    ListView1.Visible = False
    Exit Sub
  End If
  If textovar.Text = "*" And KeyCode = 106 Then
   VAR_ACTIVAR = 99
   Exit Sub
  ElseIf textovar.Text = "" Then
   VAR_ACTIVAR = 0
   Exit Sub
  End If
  If VAR_ACTIVAR <> 99 Then
    Exit Sub
  End If
  If Left(textovar.Text, 1) = "*" Then
   textovar.Text = Mid(textovar.Text, 2, Len(textovar.Text))
   textovar.SelStart = Len(textovar.Text)
  End If
Else
 If Len(textovar.Text) = 0 Or IsNumeric(textovar.Text) = True Then
   ListView1.Visible = False
   Exit Sub
 End If
End If
If ListView1.Visible = False And KeyCode <> 13 Or Len(textovar.Text) = 1 Then
    VAR = Asc(textovar.Text)
    VAR = VAR + 1
    If VAR = 33 Or VAR = 91 Then
       VAR = "ZZZZZZZZ"
    Else
       VAR = Chr(VAR)
    End If
    If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
      numarchi = 3
      archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND  ART_CODCIA = '" & LK_CODCIA & "' AND ART_CALIDAD = 1 AND ART_FLAG_STOCK = 'M' AND ART_ALTERNO BETWEEN '" & textovar.Text & "' AND  '" & VAR & "' ORDER BY ART_ALTERNO"
    Else
      numarchi = 0
      'archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND  (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND ART_CODCIA = '" & LK_CODCIA & "' AND ART_CALIDAD = 1 AND ART_FLAG_STOCK = 'M' AND ART_NOMBRE BETWEEN '" & textovar.Text & "' AND  '" & var & "' ORDER BY ART_NOMBRE"
'      archi = "SELECT ARTI.ART_KEY,ARTI.ART_CODCIA, ARTI.ART_NOMBRE, ARTI.ART_ALTERNO, ARTICULO.ARM_STOCK, PRECIOS.PRE_EQUIV, TABLAS.TAB_NOMLARGO AS DIVI, TABLAS_1.TAB_NOMLARGO AS LINEA "
'      archi = archi & "FROM ARTI INNER JOIN ARTICULO ON ARTI.ART_KEY = ARTICULO.ARM_CODART AND ARTI.ART_CODCIA = ARTICULO.ARM_CODCIA INNER JOIN PRECIOS ON ARTI.ART_KEY = PRECIOS.PRE_CODART AND ARTI.ART_CODCIA = PRECIOS.PRE_CODCIA INNER JOIN TABLAS ON ARTI.ART_CODCIA = TABLAS.TAB_CODCIA AND ARTI.ART_FAMILIA = TABLAS.TAB_NUMTAB INNER JOIN TABLAS TABLAS_1 ON ARTI.ART_CODCIA = TABLAS_1.TAB_CODCIA AND ARTI.ART_SUBFAM = TABLAS_1.TAB_NUMTAB "
'      archi = archi & "WHERE (TABLAS.TAB_TIPREG = 122) AND (TABLAS_1.TAB_TIPREG = 123) AND (PRECIOS.PRE_FLAG_UNIDAD = 'A') AND ARTI.ART_CODCIA = '" & LK_CODCIA & "' AND ARTI.ART_NOMBRE BETWEEN '" & Trim(TEXTOVAR.Text) & "%' AND  '" & VAR & "' ORDER BY ARTI.ART_NOMBRE"
    archi = "SELECT ARTI.ART_KEY,ARTI.ART_CODCIA, ARTI.ART_NOMBRE, ARTI.ART_ALTERNO, ARTICULO.ARM_STOCK, PRECIOS.PRE_EQUIV, TABLAS.TAB_NOMLARGO AS DIVI, TABLAS_1.TAB_NOMLARGO AS LINEA,PRECIOS.PRE_PRE1, PRECIOS.PRE_PRE4,ARTI.ART_CUENTA_CONTAB,ARTI.ART_CUENTA_CONTAB_C,PRECIOS.PRE_COSTO,PRECIOS.PRE_PRE22,ART_FAMILIA,ART_SUBFAM "
    archi = archi & "FROM ARTI INNER JOIN ARTICULO ON ARTI.ART_KEY = ARTICULO.ARM_CODART AND ARTI.ART_CODCIA = ARTICULO.ARM_CODCIA INNER JOIN PRECIOS ON ARTI.ART_KEY = PRECIOS.PRE_CODART AND ARTI.ART_CODCIA = PRECIOS.PRE_CODCIA INNER JOIN TABLAS ON ARTI.ART_CODCIA = TABLAS.TAB_CODCIA AND ARTI.ART_FAMILIA = TABLAS.TAB_NUMTAB INNER JOIN TABLAS TABLAS_1 ON ARTI.ART_CODCIA = TABLAS_1.TAB_CODCIA AND ARTI.ART_SUBFAM = TABLAS_1.TAB_NUMTAB "
    archi = archi & "WHERE (TABLAS.TAB_TIPREG = 122) AND (TABLAS_1.TAB_TIPREG = 123) AND (PRECIOS.PRE_FLAG_UNIDAD = 'A') AND ARTI.ART_CODCIA = '" & LK_CODCIA & "' AND ARTI.ART_NOMBRE BETWEEN '" & Trim(textovar.Text) & "%' AND  '" & VAR & "' AND ARTI.ART_situacion <> 1 ORDER BY ARTI.ART_NOMBRE"
    End If
   ' If Len(TEXTOVAR.text) > 1 And ListView1.ListItems.count = 0 Then
   ' Else
     PROC_LISVIEW ListView1
   ' End If
    Exit Sub
End If

Dim ws_codcia As String

If (ListView1.Visible = False And KeyCode <> 13 Or Len(textovar.Text) = 1) Or (Left(textovar.Text, 1) = "%" And Trim(Len(textovar.Text)) > 1) Then
    If textovar.Text = "" Then textovar.Text = " "
    VAR = Asc(textovar.Text)
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
    If Left(textovar.Text, 1) <> "%" Then
    ''  archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK ,PRE_EQUIV FROM ARTI, ARTICULO, PRECIOS WHERE (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND ART_KEY <> 0 AND ART_KEY  <> 1 and ART_CODCIA = '" & ws_codcia & "' AND ART_NOMBRE BETWEEN '" & TEXTOVAR.Text & "' AND  '" & var & "' ORDER BY ART_NOMBRE"
        archi = "SELECT ARTI.ART_KEY,ARTI.ART_CODCIA, ARTI.ART_NOMBRE, ARTI.ART_ALTERNO, ARTICULO.ARM_STOCK, PRECIOS.PRE_EQUIV, TABLAS.TAB_NOMLARGO AS DIVI, TABLAS_1.TAB_NOMLARGO AS LINEA, PRECIOS.PRE_PRE1, PRECIOS.PRE_PRE2 "
        archi = archi & "FROM ARTI INNER JOIN ARTICULO ON ARTI.ART_KEY = ARTICULO.ARM_CODART AND ARTI.ART_CODCIA = ARTICULO.ARM_CODCIA INNER JOIN PRECIOS ON ARTI.ART_KEY = PRECIOS.PRE_CODART AND ARTI.ART_CODCIA = PRECIOS.PRE_CODCIA INNER JOIN TABLAS ON ARTI.ART_CODCIA = TABLAS.TAB_CODCIA AND ARTI.ART_FAMILIA = TABLAS.TAB_NUMTAB INNER JOIN TABLAS TABLAS_1 ON ARTI.ART_CODCIA = TABLAS_1.TAB_CODCIA AND ARTI.ART_SUBFAM = TABLAS_1.TAB_NUMTAB "
        archi = archi & "WHERE (TABLAS.TAB_TIPREG = 122) AND (TABLAS_1.TAB_TIPREG = 123) AND (PRECIOS.PRE_FLAG_UNIDAD = 'A') AND ARTI.ART_CODCIA = '" & LK_CODCIA & "' AND ARTI.ART_CALIDAD = 1 AND ARTI.ART_FLAG_STOCK = 'M' AND ARTI.ART_NOMBRE BETWEEN '" & Trim(textovar.Text) & "%' AND  '" & VAR & "' ORDER BY ARTI.ART_NOMBRE"
    Else
        If KeyCode = 13 Then
            archi = "SELECT ARTI.ART_KEY,ARTI.ART_CODCIA, ARTI.ART_NOMBRE, ARTI.ART_ALTERNO, ARTICULO.ARM_STOCK, PRECIOS.PRE_EQUIV, TABLAS.TAB_NOMLARGO AS DIVI, TABLAS_1.TAB_NOMLARGO AS LINEA, PRECIOS.PRE_PRE1, PRECIOS.PRE_PRE2  "
            archi = archi & "FROM ARTI INNER JOIN ARTICULO ON ARTI.ART_KEY = ARTICULO.ARM_CODART AND ARTI.ART_CODCIA = ARTICULO.ARM_CODCIA INNER JOIN PRECIOS ON ARTI.ART_KEY = PRECIOS.PRE_CODART AND ARTI.ART_CODCIA = PRECIOS.PRE_CODCIA INNER JOIN TABLAS ON ARTI.ART_CODCIA = TABLAS.TAB_CODCIA AND ARTI.ART_FAMILIA = TABLAS.TAB_NUMTAB INNER JOIN TABLAS TABLAS_1 ON ARTI.ART_CODCIA = TABLAS_1.TAB_CODCIA AND ARTI.ART_SUBFAM = TABLAS_1.TAB_NUMTAB "
            archi = archi & "WHERE (TABLAS.TAB_TIPREG = 122) AND (TABLAS_1.TAB_TIPREG = 123) AND (PRECIOS.PRE_FLAG_UNIDAD = 'A') AND ARTI.ART_CODCIA = '" & LK_CODCIA & "' AND ARTI.ART_CALIDAD = 1 AND ARTI.ART_FLAG_STOCK = 'M' AND ARTI.ART_NOMBRE like '" & Trim(textovar.Text) & "%' ORDER BY ARTI.ART_NOMBRE"
        Else
            Exit Sub
        End If
    End If
PROC_LISVIEW ListView1, 1000
    loc_key = 0
    If ListView1.Visible Then
    loc_key = 1
    'fraprecios.Visible = True
    End If
    Exit Sub
End If


If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As ListItem    ' Variable FoundItem.
If ListView1.Visible Then
  Set itmFound = ListView1.FindItem(LTrim(textovar.Text), lvwText, , lvwPartial)
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

Private Sub textovar_LostFocus()
'TEXTOVAR.Visible = False
'If TEXTOVAR.Visible Then
'   TEXTOVAR.Visible = False
'   grid_fac.Row = wfila_act
'   grid_fac.SetFocus
   Exit Sub
'End If

End Sub

Public Sub LLENADOS(cont As ListBox, tip As Integer)
Dim CONTA As Integer
    CONTA = -1
    PUB_TIPREG = tip
    SQ_OPER = 2
    LEER_TAB_LLAVE
    cont.ToolTipText = "TAB_TIPREG = " & tip
    cont.Clear
    cont.AddItem " "
    Do Until tab_mayor.EOF
        cont.AddItem tab_mayor!tab_NOMLARGO & String(60, " ") & tab_mayor!TAB_NUMTAB
        CONTA = CONTA + 1
        tab_mayor.MoveNext
    Loop
End Sub

Private Sub tiempo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   oferta.SetFocus
End If

End Sub

Private Sub Txt_key_Change()
If txt_key.Text = "" Then lblven.Caption = ""

End Sub

Private Sub txt_key_GotFocus()
 Azul txt_key, txt_key
End Sub
Private Sub txt_key_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As ListItem    ' Variable FoundItem.
If Not ListView2.Visible Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And txt_key.Text = "" Then
  loc_key = 1
  Set ListView2.SelectedItem = ListView2.ListItems(loc_key)
  ListView2.ListItems.item(loc_key).Selected = True
  ListView2.ListItems.item(loc_key).EnsureVisible
  GoTo fin
End If

If KeyCode = 40 Then  ' flecha abajo
  loc_key = loc_key + 1
  If loc_key > ListView2.ListItems.count Then loc_key = ListView2.ListItems.count
  GoTo POSICION
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo POSICION
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > ListView2.ListItems.count Then loc_key = ListView2.ListItems.count
 GoTo POSICION
End If
If KeyCode = 33 Then
 loc_key = loc_key - 17
 If loc_key < 1 Then loc_key = 1
 GoTo POSICION
End If
GoTo fin
POSICION:
  ListView2.ListItems.item(loc_key).Selected = True
  ListView2.ListItems.item(loc_key).EnsureVisible
  txt_key.Text = Trim(ListView2.ListItems.item(loc_key).Text) & " "
  DoEvents
  txt_key.SelStart = Len(txt_key.Text)
  DoEvents
fin:

End Sub
Private Sub txt_key_KeyPress(KeyAscii As Integer)
Dim valor As String
Dim tf As Integer
Dim i
Dim itmFound As ListItem
If KeyAscii = 27 Then
 txt_key.Text = ""
End If
'If KeyAscii <> 13 Then
'   GoTo fin
'End If
If KeyAscii = 13 Then
     grid_fac_EnterCell
     textovar.Visible = False
     grid_fac.SetFocus
End If
pu_codclie = Val(txt_key.Text)
If Len(txt_key.Text) = 0 Or txt_key.Locked Then
   Exit Sub
End If
If pu_codclie <> 0 And IsNumeric(txt_key.Text) = True Then
   loc_key = 0
   On Error GoTo mucho
   PUB_CODVEN = Val(txt_key.Text)
   pu_codcia = LK_CODCIA
   SQ_OPER = 1
   LEER_VEN_LLAVE
   On Error GoTo 0
   If ven_llave.EOF Then
     Azul txt_key, txt_key
     MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
     txt_key.SetFocus
     GoTo fin
   End If
   lblven.Caption = Trim(ven_llave!VEM_NOMBRE)
   ListView1.Visible = False
   'Azul txtcli, txtcli Comentado
   ' grid_fac.SetFocus ' Modificado 04052004 (Ody)
   Screen.MousePointer = 0
   Exit Sub
Else
   If loc_key > ListView2.ListItems.count Or loc_key = 0 Then
     Exit Sub
   End If
   valor = UCase(ListView2.ListItems.item(loc_key).Text)
   If Trim(UCase(txt_key.Text)) = Left(valor, Len(Trim(txt_key.Text))) Then
   Else
      Exit Sub
   End If
   txt_key.Text = Trim(ListView2.ListItems.item(loc_key).SubItems(1))
   PUB_CODVEN = Val(txt_key.Text)
   pu_codcia = LK_CODCIA
   SQ_OPER = 1
   LEER_VEN_LLAVE
   On Error GoTo 0
   If ven_llave.EOF Then
     Azul txt_key, txt_key
     MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
     txt_key.SetFocus
     GoTo fin
   End If
   lblven.Caption = Trim(ven_llave!VEM_NOMBRE)
   ListView2.Visible = False
   txtchofer.TEXTO = ven_llave!VEM_TRNKEY
   txtChofer_ShowData ven_llave!VEM_TRNKEY
   'txtchofer.SetFocus
   
End If
dale:
mucho:
ListView1.Visible = False
fin:
End Sub

Private Sub txt_key_KeyUp(KeyCode As Integer, Shift As Integer)
Dim VAR
If Len(txt_key.Text) = 0 Or txt_key.Locked = True Or IsNumeric(txt_key.Text) = True Then
   ListView2.Visible = False
   Exit Sub
End If
If ListView2.Visible = False And KeyCode <> 13 Or Len(txt_key.Text) = 1 Then
    VAR = Asc(txt_key.Text)
    VAR = VAR + 1
    If VAR = 33 Or VAR = 91 Then
       VAR = "ZZZZZZZZ"
    Else
       VAR = Chr(VAR)
    End If
    numarchi = 9
    archi = "SELECT * FROM VEMAEST WHERE  VEM_CODCIA = '" & LK_CODCIA & "' AND VEM_NOMBRE BETWEEN '" & txt_key.Text & "' AND  '" & VAR & "' ORDER BY VEM_NOMBRE"
    PROC_LISVIEW ListView2
    loc_key = 1
    If ListView2.Visible = False Then
        loc_key = 0
    End If
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As ListItem    ' Variable FoundItem.
If ListView2.Visible Then
  Set itmFound = ListView2.FindItem(LTrim(txt_key.Text), lvwText, , lvwPartial)
  If itmFound Is Nothing Then
  Else
   itmFound.EnsureVisible
   itmFound.Selected = True
   loc_key = itmFound.Tag
   If loc_key + 8 > ListView2.ListItems.count Then
      ListView2.ListItems.item(ListView2.ListItems.count).EnsureVisible
   Else
     ListView2.ListItems.item(loc_key + 8).EnsureVisible
   End If
   DoEvents
  End If
  Exit Sub
End If
End Sub

Private Sub txtatte_KeyPress(KeyAscii As Integer)
On Error GoTo SALE
If KeyAscii = 13 Then
 grid_fac.Row = 2
 grid_fac.COL = 2
 grid_fac.COL = 1
 textovar.Visible = True
 textovar.SetFocus
End If
Exit Sub
SALE:
End Sub

Private Sub txtcli_Change()
 If txtCli.Text = "" Then
   lblcli.Caption = ""
   i_destino.Clear
 End If
End Sub

Private Sub txtcli_LostFocus()
Dim WMO As String
Dim RES_DEUDA As Currency
Dim wsumadol As Currency
Dim WTC As Currency

If SUT_LLAVE.EOF Then Exit Sub

pub_cadena = "SELECT * FROM DIRCLI WHERE CODCIA=? AND CODCLI=? AND CP=?"
Set PSFAR_TRANS = CN.CreateQuery("", pub_cadena)
PSFAR_TRANS.rdoParameters(0) = LK_CODCIA
PSFAR_TRANS.rdoParameters(1) = Val(txtCli.Text)
PSFAR_TRANS.rdoParameters(2) = "C"
Set FAR_TRANS = PSFAR_TRANS.OpenResultset(rdOpenKeyset, rdConcurValues)
i_destino.Clear
Do Until FAR_TRANS.EOF
  i_destino.AddItem Trim(FAR_TRANS!dircomp) & String(80, " ") & Trim(FAR_TRANS!DIRCLI)
  FAR_TRANS.MoveNext
Loop
If i_destino.ListCount > 0 Then i_destino.ListIndex = 0
If Not cli_llave.EOF Then
    For fila = 0 To i_fbg.ListCount - 1
        i_fbg.ListIndex = fila
        If Trim(i_fbg.Text) = Trim(cli_llave!CLI_TIPO) Then Exit For
    Next fila
    If Left(Trim(Right(Trim(i_condi.Text), 3)), 1) = "1" And Trim(Right(Trim(i_condi.Text), 2)) = "FA" Then
        i_dias.Text = Val(Nulo_Valor0(cli_llave!CLI_AUTO1))
    Else
        i_dias.Text = ""
    End If
End If
' Comentado Ody
If Not blnConsulta Then
If Not cli_llave.EOF Then
    If LK_FLAG_EXED = "A" Then
     If Nulo_Valor0(SUT_LLAVE!SUT_FLAG_CC) <> 1 And pub_signo_car <> 0 And Nulo_Valor0(cli_llave!cli_limcre) = 0 And Nulo_Valor0(cli_llave!cli_limcre2) = 0 Then
         MsgBox "Cliente No tiene Limite de Credito.", 48, Pub_Titulo
     End If
    End If
   If SUT_LLAVE!SUT_SIGNO_CAR = 1 Then
      pu_codcia = LK_CODCIA
      pub_deuda = CAR_TOT_CPX2("C", pu_codcia, cli_llave!cli_codclie)
      ' Quitar Mensajes
      If PUB_FLAG_VENCIDO = 1 Or PUB_FLAG_VENCIDO_VISTA = 1 Then
        If PUB_FLAG_VENCIDO_VISTA = 1 Then
          MsgBox "OJO... Tiene Documentos Pendientes.", 48, Pub_Titulo
        Else
          MsgBox "Cliente Tiene Obligaciones Vencidas. << Moroso>>.", 48
        End If
         MsgBox pub_mensaje, 48, Pub_Titulo
      End If
      If Nulo_Valors(cli_llave!CLI_TIPO_BLOQ1) = "1" Then
         MsgBox "Cliente con Credito Bloqueado ... (No procede su Venta al Credito)", 48, Pub_Titulo
         txtCli.Text = ""
         Azul txtCli, txtCli
         Exit Sub
      End If
      i_dias.Text = Val(Nulo_Valor0(cli_llave!CLI_AUTO1))
'      If PUB_FLAG_DOC > Nulo_Valor0(cli_llave!CLI_AUTO1) Then
'         MsgBox "Cliente alcanzo el tope de Documentos " & Chr(13) & "Emiitidos : " & PUB_FLAG_DOC & Chr(13) & "Autorizados : " & Trim(Nulo_Valor0(cli_llave!CLI_AUTO1)) & Chr(13) & " No procede la Venta", 48, Pub_Titulo
'         i_codcli.Text = ""
'         Exit Sub
'      End If

'   End If
End If

End If
End If

' ver si usuario tiene acceso

If Not cli_llave.EOF And Nulo_Valors(par_llave!par_flag_cred) <> "A" And Nulo_Valor0(SUT_LLAVE!SUT_FLAG_CC) = 0 Then
   If SUT_LLAVE!SUT_SIGNO_CAR = 1 Then
      pu_codcia = LK_CODCIA
      pub_deuda = CAR_TOT_CPX2("C", pu_codcia, cli_llave!cli_codclie)
      If PUB_FLAG_VENCIDO = 1 And LK_FLAG_LIMITE <> "A" And LK_FLAG_LIMITE <> "C" Then
          MsgBox "CLIENTE TIENE OBLIGACIONES VENCIDAS ... ", 48, Pub_Titulo
          Exit Sub
    End If
    PUB_CAL_INI = LK_FECHA_DIA
    PUB_CAL_FIN = LK_FECHA_DIA
    pu_codcia = LK_CODCIA
    PUB_CODCIA = LK_CODCIA
    SQ_OPER = 1
    LEER_CAL_LLAVE
    WTC = 0
    If Not cal_llave.EOF Then
      WTC = Nulo_Valor0(cal_llave!cal_tipo_cambio)
    End If
    If WTC = 0 Then
      MsgBox "Venta falta parametros ...INGRESE TIPO DE CAMBIO DEL DIA", 48, Pub_Titulo
      Exit Sub
    End If
    If Trim(Left(moneda.Text, 1)) = "S" Then
     wsumadol = Val(Nulo_Valor0(cli_llave!cli_limcre)) + Val(redondea((Nulo_Valor0(cli_llave!cli_limcre2) * WTC)))
     RES_DEUDA = pub_deuda
     WMO = "S/."
    Else
     wsumadol = Val(redondea(Nulo_Valor0(cli_llave!cli_limcre) / WTC)) + Val(redondea(Val(Nulo_Valor0(cli_llave!cli_limcre2))))
     RES_DEUDA = redondea(Val(pub_deuda / WTC))
     WMO = "US$."
    End If
If (RES_DEUDA + Val(txttotal.Text)) > wsumadol And LK_FLAG_LIMITE <> "B" And LK_FLAG_LIMITE <> "C" Then
   MsgBox "LIMITE DE CREDITO EXCEDIDO ...SALDO POR ATENDER : " & WMO & " " & Format(wsumadol - RES_DEUDA, "0.00") & Chr(13) & "*** Venta No Procede ***", 48, Pub_Titulo
   txtCli.Text = ""
   Azul txtCli, txtCli
   Exit Sub
End If
   End If
End If

     
End Sub

Private Sub txtdoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 PUB_NUMSER = Val(tserie.Text)
 PUB_NUMFAC = Val(txtdoc.Text)
 LLENA_DOCU
End If
End Sub

Private Sub txtRuc_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
  If Len(Trim(txtruc.Text)) <> LK_DIG_RUC Then
    MsgBox "R.U.C. No procede ", 48, Pub_Titulo
    Azul txtruc, txtruc
    Exit Sub
  Else
    moneda.SetFocus
  End If
End If

End Sub

Private Sub UNIDAD_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then
 UNIDAD.Visible = False
 grid_fac.SetFocus
End If


If KeyAscii <> 13 Then Exit Sub

SQ_OPER = 1
pu_codcia = LK_CODCIA
PUB_CODART = Val(grid_fac.TextMatrix(grid_fac.Row, 10))
PUB_SECUEN = Val(Right(UNIDAD.Text, 4))
LEER_PRE_LLAVE
If pre_llave.EOF Then Exit Sub
grid_fac.TextMatrix(grid_fac.Row, 3) = Trim(Left(UNIDAD.Text, 12))
grid_fac.TextMatrix(grid_fac.Row, 4) = pre_llave!pre_pre1 'Format(Val(grid_fac.TextMatrix(grid_fac.Row, 11)) / Val(grid_fac.TextMatrix(grid_fac.Row, 17)), "0.00")
grid_fac.TextMatrix(grid_fac.Row, 7) = redondea(Nulo_Valor0(pre_llave!pre_PESO) * Val(grid_fac.TextMatrix(grid_fac.Row, 2)))
grid_fac.TextMatrix(grid_fac.Row, 11) = pre_llave!pre_secuencia
grid_fac.TextMatrix(grid_fac.Row, 12) = pre_llave!PRE_EQUIV
grid_fac.TextMatrix(grid_fac.Row, 16) = pre_llave!pre_unidad
stock.Caption = Format(Val(grid_fac.TextMatrix(grid_fac.Row, 15)) / Val(grid_fac.TextMatrix(grid_fac.Row, 12)), "0.00")
unid.Caption = grid_fac.TextMatrix(grid_fac.Row, 16)
nomarti.Caption = grid_fac.TextMatrix(grid_fac.Row, 0)

UNIDAD.Visible = False
suma_grid
grid_fac.COL = 4
grid_fac_KeyPress 13

End Sub
Public Function REP_CONSUL() As Integer
Dim WMONEDA As String * 1
Dim wser As String * 3
Dim WSRUTA As String
Dim indice As Integer
Dim wm As Integer
Dim llave_rep01 As rdoResultset
Dim PS_REP01 As rdoQuery
Dim i As Integer
Dim valor
Dim loc_xl As Object
Dim loc_codtra As Integer
Dim wRuta As String
Dim WSNUMDOC As String
Dim numero_device As Integer
'If LK_EMP = "HER" Then
'  wRuta = "C:\ADMIN\STANDAR\"
'Else
LOC_TIPMOV = 201
If LK_EMP_PTO = "A" Then
  wRuta = PUB_RUTA_OTRO & "PTOVTA\"
Else
  wRuta = PUB_RUTA_OTRO
End If
If Left(moneda.Text, 1) = "S" Then
 WMONEDA = "S"
Else
 WMONEDA = "D"
End If

'End If
  If Trim(Nulo_Valors(par_llave!PAR_DEVICE_FBG)) <> "" Then
     'numero_device = 0
     'Reportes.PrinterName = Printers(numero_device).DeviceName
     'Reportes.PrinterDriver = Printers(numero_device).DriverName '"RASDD.DLL"
     'Reportes.PrinterPort = Printers(numero_device).Port
  End If

    FORM_COT.Reportes.Connect = PUB_ODBC
    FORM_COT.Reportes.Destination = crptToWindow  '= crptToPrinter
    FORM_COT.Reportes.WindowLeft = 2
    FORM_COT.Reportes.WindowTop = 70
    FORM_COT.Reportes.WindowWidth = 635
    FORM_COT.Reportes.WindowHeight = 390
    FORM_COT.Reportes.Formulas(1) = ""
    PUB_NETO = Val(txttotal.Text)
    PU_NUMSER = Val((tserie.Text))
    PU_NUMFAC = Val((txtdoc.Text))
    FORM_COT.Reportes.Formulas(1) = ""
    FORM_COT.Reportes.Formulas(1) = "SON_EFECTIVO=  ' " & CONVER_LETRAS(PUB_NETO, WMONEDA) & "'"
    FORM_COT.Reportes.WindowTitle = "COTIZACION:" & Format(PU_NUMSER, "000") & " - " & Format(PU_NUMFAC, "00000000")
    FORM_COT.Reportes.ReportFileName = wRuta + "COTI.RPT"
    pub_cadena = "{PEDIDOS.PED_TIPMOV} = " & LOC_TIPMOV & " AND {PEDIDOS.PED_CODCIA} = '" & LK_CODCIA & "' AND  {PEDIDOS.PED_NUMSER} = '" & PU_NUMSER & "' AND {PEDIDOS.PED_NUMFAC} = " & PU_NUMFAC
    FORM_COT.Reportes.SelectionFormula = pub_cadena
    'FORM_COT.Reportes.Destination = crptToPrinter
    FORM_COT.Reportes.Destination = crptToWindow
    On Error GoTo accion
    FORM_COT.Reportes.Action = 1
    On Error GoTo 0
Exit Function
accion:
 MsgBox Err.Description
 MsgBox "Intente Nuevamente, la impresion de Modo manual", 48, Pub_Titulo
 Exit Function

End Function

Private Sub txtcli_GotFocus()
 Azul txtCli, txtCli
End Sub

Private Sub txtcli_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As ListItem    ' Variable FoundItem.
If Not ListView1.Visible Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And txtCli.Text = "" Then
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
  txtCli.Text = Trim(ListView1.ListItems.item(loc_key).Text) & " "
  txtCli.SelStart = Len(txtCli.Text)
fin:

End Sub
Private Sub txtcli_KeyPress(KeyAscii As Integer)
Dim valor As String
Dim tf As Integer
Dim i
Dim itmFound As ListItem
On Error GoTo SALCODI
If KeyAscii = 27 Then
 txtCli.Text = ""
 lblcli.Caption = ""
End If
If KeyAscii <> 13 Then
   GoTo fin
End If
On Error GoTo CODI_ERR

'If Trim(txtcli.Text) = 109 Then  'cambiado gts
 ' Fracli.Top = txtcli.Top
 ' Fracli.Left = txtcli.Left + 300
 ' Fracli.Visible = True
 ' t_nombre.SetFocus
 ' Azul t_nombre, t_nombre
 ' Exit Sub
'End If


   SQ_OPER = 1
   On Error GoTo mucho
   pu_codcia = LK_CODCIA
   pu_cp = "C"
   'pu_codclie = Val(txtcli.Text)
   pu_codclie = Val(ListView1.ListItems.item(loc_key).SubItems(1))
   LEER_CLI_LLAVE
   On Error GoTo 0
   If cli_llave.EOF Then
     Azul txtCli, txtCli
    ' MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
     txtCli.SetFocus
     GoTo fin
   End If
   
   If pu_codclie <> 0 And IsNumeric(txtCli.Text) = True Then
   If Len(Trim(txtCli.Text)) = LK_DIG_RUC Then ' LONG DEL RUC
        pu_cp = "C"
        PUB_RUC = Trim(txtCli.Text)
        SQ_OPER = 4
        pu_codcia = LK_CODCIA
        LEER_CLI_LLAVE
        If cli_ruc.EOF Then
           MsgBox "R.U.C. No Existe ", 48, Pub_Titulo
           Exit Sub
        End If
  
        txtCli.Text = cli_ruc!cli_codclie
   End If
   
   ListView1.Visible = False
   txtCli.Text = cli_llave!cli_codclie
   FORM_COT.lblcli.Caption = cli_llave!CLI_NOMBRE
   If Trim(cli_llave!cli_ruc_esposo) <> "" Then
     txtruc.Text = cli_llave!cli_ruc_esposo
   Else
     txtruc.Text = cli_llave!cli_RUC_ESPOSA
   End If
   GoTo salta_dir
   Screen.MousePointer = 0
Else
   If loc_key > ListView1.ListItems.count Or loc_key = 0 Then
     Exit Sub
   End If
   valor = UCase(ListView1.ListItems.item(loc_key).Text)
   If Trim(UCase(txtCli.Text)) = Left(valor, Len(Trim(txtCli.Text))) Then
   Else
      Exit Sub
   End If

   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   pu_cp = "C"
   pu_codclie = Val(ListView1.ListItems.item(loc_key).SubItems(1))
   LEER_CLI_LLAVE
   On Error GoTo 0
   If cli_llave.EOF Then
     Azul txtCli, txtCli
     MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
     txtCli.SetFocus
     GoTo fin
   End If
   
   
   
   ListView1.Visible = False
   txtCli.Text = cli_llave!cli_codclie
   txt_key.Text = cli_llave("CLI_CIA_REF")
   FORM_COT.lblcli.Caption = cli_llave!CLI_NOMBRE
   If Trim(cli_llave!cli_ruc_esposo) <> "" Then
     txtruc.Text = cli_llave!cli_ruc_esposo
   Else
     txtruc.Text = cli_llave!cli_RUC_ESPOSA
   End If
   GoTo salta_dir
End If
dale:
ListView1.Visible = False
fin:
mucho:
CODI_ERR:
Exit Sub


SALCODI:
MsgBox Err.Description & " Intente Nuevamente ", 48, Pub_Titulo
Unload frmCLI
Exit Sub
salta_dir:
    txt_key.SetFocus
    txt_key_KeyPress 13
'i_destino.SetFocus
'SendKeys "%{up}"
End Sub

Private Sub txtcli_KeyUp(KeyCode As Integer, Shift As Integer)
Dim NADA
Dim VAR
If Len(txtCli.Text) = 0 Or IsNumeric(txtCli.Text) = True Then
   ListView1.Visible = False
   Exit Sub
End If
If ListView1.Visible = False And KeyCode <> 13 Or Len(txtCli.Text) = 1 Then
    If txtCli.Text = "" Then txtCli.Text = " "
    VAR = Asc(txtCli.Text)
    VAR = VAR + 1
    NADA = VAR
    If VAR = 33 Or VAR = 91 Then
       VAR = "ZZZZZZZZ"
    Else
       VAR = Chr(VAR)
    End If
    numarchi = 1
    'archi = "SELECT CLI_CODCLIE, CLI_CODCIA, CLI_CP, CLI_NOMBRE,CLI_CASA_DIREC,CLI_ZONA_NEW, CLI_CASA_NUM FROM CLIENTES WHERE  CLI_CP = 'C' AND CLI_CODCIA = '" & LK_CODCIA & "' AND CLI_NOMBRE BETWEEN '" & txtcli.Text & "' AND  '" & VAR & "' ORDER BY CLI_NOMBRE"
    archi = "SELECT CLI_CODCLIE , CLI_CODCIA, CLI_CP, CLI_NOMBRE, CLI_CASA_DIREC,CLI_ZONA_NEW, CLI_CASA_NUM, TAB_NOMLARGO  FROM CLIENTES,TABLAS WHERE (TAB_CODCIA = '00') AND (TAB_TIPREG = 35) AND (TAB_NUMTAB = CLI_ZONA_NEW) AND CLI_CP = 'C' AND CLI_CODCIA = '" & LK_CODCIA & "' AND CLI_NOMBRE BETWEEN '" & txtCli.Text & "' AND  '" & VAR & "' ORDER BY CLI_NOMBRE"
    PROC_LISVIEW ListView1
    loc_key = 1
    If NADA = 33 Or NADA = 91 Then
      If ListView1.Visible = False Then
        loc_key = 0
        MsgBox "No existe Datos ...", 48, Pub_Titulo
        txtCli.Text = ""
      End If
    End If
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As ListItem    ' Variable FoundItem.
If ListView1.Visible Then
  Set itmFound = ListView1.FindItem(LTrim(txtCli.Text), lvwText, , lvwPartial)
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
  End If
  Exit Sub
End If
End Sub


Public Sub LLENA_DOCU()
Dim MON As String
LIMPIA_DATOS
CABE_MAN

PSLOC_WARTI(0) = LK_CODCIA
PSLOC_WARTI(1) = 201
PSLOC_WARTI(2) = PUB_NUMSER
PSLOC_WARTI(3) = PUB_NUMFAC
llave_sum_arti.Requery
If llave_sum_arti.EOF Then
  tserie.Text = PUB_NUMSER
  txtdoc.Text = PUB_NUMFAC
  MsgBox "No Existe Cotización.", 48, Pub_Titulo
  Azul txtdoc, txtdoc
  Exit Sub
End If
txtigv.Text = llave_sum_arti!PED_IGV
txtvalorv.Text = llave_sum_arti!PED_BRUTO
txttotal.Text = Format(llave_sum_arti!PED_IGV + llave_sum_arti!PED_BRUTO, "0.00")
FORM_COT.lblcli.Caption = llave_sum_arti!PED_NOMCLIE
txtruc.Text = llave_sum_arti!PED_RUCCLIE
txtCli.Text = llave_sum_arti!PED_CODCLIE
txt_key.Text = llave_sum_arti!PED_CODVEN
i_dias.Text = llave_sum_arti!ped_DIAS

txtFecha.Text = Format(llave_sum_arti!PED_FECHA, "dd/mm/yyyy")
'If Trim(llave_sum_arti!ped_SITUACION) = "P" Then
'If Trim(llave_sum_arti!ped_situacion) = "A" Then
' LBLSIT.Caption = "01 P E D I D O   P R O C E S A D O - " & Trim(llave_sum_arti!PED_FORMA)
' LBLSIT.ForeColor = vbRed
'Else
' LBLSIT.Caption = "02 P E D I D O   P E N D I E N T E "
' LBLSIT.ForeColor = vbBlue
'End If
If Trim(llave_sum_arti!ped_FBG) = "F" Then
 i_fbg.ListIndex = 0
Else
 i_fbg.ListIndex = 1
End If

If llave_sum_arti!PED_MONEDA = "S" Then
 moneda.ListIndex = 0
 i_moneda.Caption = "S/."
 grid_fac.TextMatrix(1, 4) = "S/."
Else
 moneda.ListIndex = 1
 i_moneda.Caption = "US$."
 grid_fac.TextMatrix(1, 4) = "US$."
End If


txtcli_KeyPress 13
txtcli_LostFocus

For fila = 0 To i_destino.ListCount - 1
  If Val(Trim(Right(i_destino.List(fila), 8))) = Val(llave_sum_arti!ped_DIRCLI) Then
   i_destino.ListIndex = fila
  End If
Next fila
tserie.Text = PUB_NUMSER
txtdoc.Text = PUB_NUMFAC
txtatte.Text = Nulo_Valors(llave_sum_arti!PED_CONTACTO)
txtObservaciones.Text = Nulo_Valors(llave_sum_arti!PED_OFERTA)
chkAprobacion.Value = Val(Nulo_Valors(llave_sum_arti!PED_APROBADO))
'llave_sum_arti!PED_HORA

fila = 2
Do Until llave_sum_arti.EOF
   SQ_OPER = 1
   PUB_KEY = llave_sum_arti!PED_CODART
   pu_codcia = LK_CODCIA
   LEER_ART_LLAVE
   grid_fac.rows = grid_fac.rows + 1
   grid_fac.RowHeight(grid_fac.rows - 1) = 285
   grid_fac.TextMatrix(fila, 1) = art_LLAVE!art_alterno
   grid_fac.TextMatrix(fila, 0) = art_LLAVE!ART_NOMBRE
   If LK_EMP = "3AA" Then
     grid_fac.TextMatrix(fila, 11) = 0
     grid_fac.TextMatrix(fila, 14) = 1
   End If
   
   grid_fac.TextMatrix(fila, 2) = llave_sum_arti!PED_CANTIDAD
   grid_fac.TextMatrix(fila, 4) = llave_sum_arti!PED_PRECIO
   grid_fac.TextMatrix(fila, 10) = llave_sum_arti!PED_CODART
   grid_fac.TextMatrix(fila, 3) = llave_sum_arti!PED_UNIDAD
   grid_fac.TextMatrix(fila, 12) = llave_sum_arti!PED_EQUIV
   grid_fac.TextMatrix(fila, 5) = llave_sum_arti!ped_DESCTO
   fila = fila + 1
   llave_sum_arti.MoveNext
Loop
suma_grid
grid_fac.Enabled = True

ESTADO.Enabled = True
Azul txtdoc, txtdoc
cmdIngreso.Caption = "&Grabar"
cmdIngreso.Enabled = True
tserie.Enabled = False
txtdoc.Enabled = False
End Sub

Public Sub carga_venta()
SQ_OPER = 2
PUB_CODTRA = 2401
LEER_SUT_LLAVE
i_condi.Clear
Do Until SUT_MAYOR.EOF
 i_condi.AddItem Format(SUT_MAYOR!SUT_SECUENCIA, "00") & ".-" & SUT_MAYOR!sut_descripcion & String(180, " ") & SUT_MAYOR!SUT_SIGNO_CAR & SUT_MAYOR!sut_TIPDOC
 SUT_MAYOR.MoveNext
Loop
moneda.Clear
If LK_MONEDA = "S" Then
   moneda.AddItem "S = S/."
ElseIf LK_MONEDA = "D" Then
   moneda.AddItem "D = US$"
Else
   moneda.AddItem "S = S/."
   moneda.AddItem "D = US$"
End If
txtFecha.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")

End Sub
'++++++++++++++++++++++++++++++++++++++++++++++++++++++
'agregado po rmic
Private Sub FactorDescto(ByVal iRow As Integer)
Dim DesctosSuc As String
Dim valor As Integer
Dim Signo As String
Dim LenCadena As Integer
Dim iPos As Integer
Dim mSignos(1 To 12) As Integer
Dim CountSignos As Integer
Dim FactorAcum As Double
Dim i As Integer
Dim vPor As String

    DesctosSuc = Trim(grid_fac.TextMatrix(iRow, 5))
    LenCadena = Len(DesctosSuc)
    For i = 1 To LenCadena
        Signo = Mid(DesctosSuc, i, 1)
        If InStr(1, Signo, "+") <> 0 Then
            CountSignos = CountSignos + 1
            mSignos(CountSignos) = i
        End If
    Next i
    If CountSignos > 0 Then
        For i = 1 To CountSignos + 1
            If i = 1 Then
                vPor = Mid(DesctosSuc, 1, mSignos(i) - 1)
                FactorAcum = (100 - Val(vPor)) / 100
            ElseIf i = CountSignos + 1 Then
                vPor = Mid(DesctosSuc, mSignos(i - 1) + 1, LenCadena - mSignos(i - 1))
                FactorAcum = FactorAcum * ((100 - Val(vPor)) / 100)
            Else
                vPor = Mid(DesctosSuc, mSignos(i - 1) + 1, mSignos(i) - mSignos(i - 1) - 1)
                FactorAcum = FactorAcum * ((100 - Val(vPor)) / 100)
            End If
        Next i
        FactorAcum = 100 - 100 * FactorAcum
    Else
        FactorAcum = Val(DesctosSuc)
    End If
    FACTOR_DESCTO = FactorAcum
End Sub
'para mostrar chofer
Private Sub txtChofer_Cancel()
    txtchofer.TEXTO = ""
    lblChofer = ""
End Sub
Private Sub txtChofer_GetRegistros(ByVal oKeyFind As Variant)
Dim sSql As String
    sSql = "SELECT 'Razon Social de la Empresa'=TRN_NOMBRE ,'Codigo'=TRN_KEY FROM Transporte WHERE TRN_Codcia= '" & LK_CODCIA & "' AND TRN_Nombre LIKE '" & oKeyFind & "%' ORDER BY TRN_Nombre"
    txtchofer.TypeFind = NameField
    txtchofer.SetRecordset = OpenSQLForwardOnly(sSql)
End Sub
Private Sub txtChofer_GotFocus()
    txtchofer.ZOrder 0
End Sub
Private Sub txtChofer_ShowData(ByVal oKey As Variant)
    SQ_OPER = 1
    PUB_CODCHOFER = Val(oKey)
    pu_codcia = LK_CODCIA
    LEER_CHO_LLAVE
    If Not RSTRA.EOF Then
        'Call FormatLblDato(txtChofer, lblCP)
        lblChofer.Caption = Trim(UCase(RSTRA("TRN_Nombre")))
    End If
    'moneda.SetFocus
    SendKeys "%{up}"
End Sub
