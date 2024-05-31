VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDeliveryUltCompras 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ultimas Compras"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9450
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   9450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   480
      Left            =   8160
      TabIndex        =   9
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      Height          =   315
      Left            =   8730
      Picture         =   "frmDeliveryUltCompras.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   675
   End
   Begin MSComCtl2.DTPicker dtpDel 
      Height          =   315
      Left            =   5280
      TabIndex        =   5
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Format          =   100794369
      CurrentDate     =   41547
   End
   Begin MSComctlLib.ListView lvPedidos 
      Height          =   2655
      Left            =   30
      TabIndex        =   2
      Top             =   480
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   4683
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComCtl2.DTPicker dtpAl 
      Height          =   315
      Left            =   7200
      TabIndex        =   6
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Format          =   100794369
      CurrentDate     =   41547
   End
   Begin MSComctlLib.ListView lvDetalle 
      Height          =   2655
      Left            =   30
      TabIndex        =   8
      Top             =   3240
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   4683
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Al:"
      Height          =   195
      Left            =   6960
      TabIndex        =   4
      Top             =   180
      Width           =   240
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Del:"
      Height          =   195
      Left            =   4920
      TabIndex        =   3
      Top             =   180
      Width           =   360
   End
   Begin VB.Label lblCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   780
      TabIndex        =   1
      Top             =   120
      Width           =   4035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   180
      Width           =   675
   End
End
Attribute VB_Name = "frmDeliveryUltCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gPlaca As String
Public gAcepta As Boolean

Private Sub ConfigurarLV()

    With Me.lvPedidos
        .ColumnHeaders.Add , , "SERIE"
        .ColumnHeaders.Add , , "NUMERO"
        .ColumnHeaders.Add , , "DOCTO"
        .ColumnHeaders.Add , , "FECHA"
        .ColumnHeaders.Add , , "TOTAL"
        .Gridlines = True
        .LabelEdit = lvwManual
        .View = lvwReport
        .FullRowSelect = True
    End With

    With Me.lvDetalle
        .ColumnHeaders.Add , , "CANT", 800
        .ColumnHeaders.Add , , "PRODUCTO", 5000
        .ColumnHeaders.Add , , "UNIDAD"
        .ColumnHeaders.Add , , "PRECIO"
        .ColumnHeaders.Add , , "IMPORTE"
        .Gridlines = True
        .LabelEdit = lvwManual
        .View = lvwReport
        .FullRowSelect = True
    End With

End Sub

Private Sub CargarUltimosPedidos()
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_VENTAS_HISTORIAL_CAB"

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCLIE", adVarChar, adParamInput, 20, gPlaca)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHAINI", adDBTimeStamp, adParamInput, , Me.dtpDel.Value)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHAFIN", adDBTimeStamp, adParamInput, , Me.dtpAl.Value)

    Dim orsDoctos As ADODB.Recordset

    Set orsDoctos = oCmdEjec.Execute

    Me.lvPedidos.ListItems.Clear

    Dim itemO As Object

    Do While Not orsDoctos.EOF
        Set itemO = Me.lvPedidos.ListItems.Add(, , orsDoctos!serie)
        itemO.Tag = orsDoctos!TD
        itemO.SubItems(1) = orsDoctos!NUMERO
        itemO.SubItems(2) = orsDoctos!DOCTO
        itemO.SubItems(3) = orsDoctos!fecha
        itemO.SubItems(4) = orsDoctos!Total
    
        orsDoctos.MoveNext
    Loop
    If Me.lvPedidos.ListItems.count <> 0 Then Me.lvPedidos.SelectedItem.Selected = False
End Sub

Private Sub cmdAceptar_Click()

'    Dim cI   As Integer
'
'    Dim vENC As Boolean
'
'    Dim MSN  As String
'
'    vENC = False
'
'    Dim Fr As Integer
'
'    Dim Cf As Integer
'
'   ' FORMGEN.i_codven.Text = Me.lvDetalle.SelectedItem.SubItems(5)
'   ' FORMGEN.i_nomven.Caption = Me.lvDetalle.SelectedItem.SubItems(3)
'   ' FORMGEN.i_concepto.Text = Me.lvDetalle.SelectedItem.SubItems(2)
'    'FORMGEN.txtRequerimiento.Text = Me.lvDetalle.SelectedItem.SubItems(6)
'
'    For Cf = 1 To Me.lvDetalle.ListItems.count
'        cI = FORMGEN.grid_fac.Rows
'        Fr = cI
'
'        'FALTA AGREGAR EL NRO DE REQUERIMIENTO PARA PODER VALIDAR EN FUTUROS VALES
'       ' If Me.lvDetalle.ListItems(Cf).Checked Then
'           ' If cI = 3 Then
'
'                With FORMGEN.grid_fac
'                    .TextMatrix(cI - 1, 0) = Me.lvDetalle.ListItems(Cf).SubItems(1) 'NOMBRE PRODUCTO
'                    .TextMatrix(cI - 1, 1) = Me.lvDetalle.ListItems(Cf).Tag 'CODIGO DEL PRODUCTO
'                    .TextMatrix(cI - 1, 33) = Me.lvDetalle.ListItems(Cf).Tag 'CODIGO DEL PRODUCTO
'                    .TextMatrix(cI - 1, 16) = Me.lvDetalle.ListItems(Cf).Tag 'CODIGO DEL PRODUCTO
'                    .TextMatrix(cI - 1, 14) = 1
'                    .TextMatrix(cI - 1, 4) = Me.lvDetalle.ListItems(Cf).Text 'REQUERIDO
'                    .TextMatrix(cI - 1, 5) = Me.lvDetalle.ListItems(Cf).SubItems(2) & Space(20) & Space(7) & "1" 'UNIDA Y EQUIVALENCIA
'                    .TextMatrix(cI - 1, 6) = Me.lvDetalle.ListItems(Cf).SubItems(3) 'COSTO
'                    .TextMatrix(cI - 1, 34) = Me.lvDetalle.SelectedItem.SubItems(3) 'IDACTIVIDAD
'                    .TextMatrix(cI - 1, 3) = Me.lvDetalle.ListItems(Cf).SubItems(2) 'UNIDAD
'                    .TextMatrix(cI - 1, 8) = 0
'                    .TextMatrix(cI - 1, 26) = 5 'ped_llave!ped_TIPMOV
'                    .TextMatrix(cI - 1, 10) = 0
'                    .TextMatrix(cI - 1, 42) = 0
'                    .TextMatrix(cI - 1, 9) = Me.lvDetalle.SelectedItem.Text 'REQUERIMIENTO
'                End With
'
'                FORMGEN.grid_fac.Rows = FORMGEN.grid_fac.Rows + 1
''            Else
''
''                For Fr = 2 To FORMGEN.grid_fac.Rows - 1
''                    If Len(Trim(FORMGEN.grid_fac.TextMatrix(Fr, 16))) <> 0 Then
''                        If FORMGEN.grid_fac.TextMatrix(Fr, 16) = Me.lvDetalle.ListItems(Cf).Tag And FORMGEN.grid_fac.TextMatrix(Fr, 34) = Me.lvDetalle.SelectedItem.SubItems(1) Then
''                            vENC = True
''                            MSN = "Ubicado"
''                            Exit For
''                        Else
''                            vENC = False
''                        End If
''                    End If
''                Next
''
''                If vENC = False Then
''
''                    With FORMGEN.grid_fac
''                        .TextMatrix(cI - 1, 0) = Me.lvDetalle.ListItems(Cf).Text
''                        .TextMatrix(cI - 1, 1) = Me.lvDetalle.ListItems(Cf).Tag
''                        .TextMatrix(cI - 1, 33) = Me.lvDetalle.ListItems(Cf).Tag
''                        .TextMatrix(cI - 1, 16) = Me.lvDetalle.ListItems(Cf).Tag
''                        .TextMatrix(cI - 1, 14) = 1
''                        .TextMatrix(cI - 1, 4) = Me.lvDetalle.ListItems(Cf).SubItems(1)
''                        .TextMatrix(cI - 1, 5) = Me.lvDetalle.ListItems(Cf).SubItems(2) & Space(20) & Space(7) & "1"
''                        .TextMatrix(cI - 1, 6) = Me.lvDetalle.ListItems(Cf).SubItems(3)
''                        .TextMatrix(cI - 1, 34) = Me.lvDetalle.SelectedItem.SubItems(1)
''                        .TextMatrix(cI - 1, 3) = Me.lvDetalle.ListItems(Cf).SubItems(2)
''                        .TextMatrix(cI - 1, 8) = 0
''                        .TextMatrix(cI - 1, 26) = 5 'ped_llave!ped_TIPMOV
''                        .TextMatrix(cI - 1, 10) = 0
''                        .TextMatrix(cI - 1, 42) = 0
''                        .TextMatrix(cI - 1, 9) = Me.lvDetalle.SelectedItem.Text
''
''                        FORMGEN.grid_fac.Rows = FORMGEN.grid_fac.Rows + 1
''                    End With
''
''                    'Else
''                    ' vENC = False
''                End If
'         '   End If
'       ' End If
'
'    Next
'
'    FORMGEN.calcula_totales
    Unload Me
End Sub

Private Sub cmdSearch_Click()
CargarUltimosPedidos
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Me.dtpAl.Value = LK_FECHA_DIA
    Me.dtpDel.Value = DateAdd("m", -1, Me.dtpAl.Value)
    ConfigurarLV
    CargarUltimosPedidos

End Sub

Private Sub lvPedidos_ItemClick(ByVal item As MSComctlLib.ListItem)
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_VENTAS_HISTORIAL_DET"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , Me.lvPedidos.SelectedItem.SubItems(3))
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TIPODOCTO", adChar, adParamInput, 1, Me.lvPedidos.SelectedItem.Tag)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSER", adInteger, adParamInput, , Me.lvPedidos.SelectedItem.Text)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMFAC", adBigInt, adParamInput, , Me.lvPedidos.SelectedItem.SubItems(1))

    Dim oRSdet As ADODB.Recordset

    Set oRSdet = oCmdEjec.Execute
        
    Me.lvDetalle.ListItems.Clear

Dim itemd As Object
    Do While Not oRSdet.EOF
    Set itemd = Me.lvDetalle.ListItems.Add(, , oRSdet!cant)
    itemd.SubItems(1) = Trim(oRSdet!PRODUCTO)
   itemd.SubItems(2) = Trim(oRSdet!UNIDAD)
    itemd.SubItems(3) = oRSdet!PRECIO
    itemd.SubItems(4) = oRSdet!Importe
    itemd.Tag = oRSdet!Codigo
        oRSdet.MoveNext
    Loop
    
End Sub
