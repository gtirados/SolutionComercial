VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form frmGuiaRemision 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Guia de Remisión"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11775
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   11775
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "Procesar"
      Height          =   600
      Left            =   10320
      Picture         =   "frmGuiaRemision.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton cmdConsultar 
      Caption         =   "Consultar"
      Height          =   720
      Left            =   10080
      Picture         =   "frmGuiaRemision.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   990
   End
   Begin MSComctlLib.ListView lvSearch 
      Height          =   1695
      Left            =   840
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   2990
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSMask.MaskEdBox MasFechainicio 
      Height          =   345
      Left            =   840
      TabIndex        =   1
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   609
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Frame Frame2 
      Height          =   2655
      Left            =   120
      TabIndex        =   9
      Top             =   3840
      Width           =   11535
      Begin MSComctlLib.ListView lvDet 
         Height          =   2295
         Left            =   120
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   240
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   4048
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
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   11535
      Begin MSComctlLib.ListView lvCab 
         Height          =   2295
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   4048
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin MSMask.MaskEdBox MasFechaFin 
      Height          =   345
      Left            =   6480
      TabIndex        =   2
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   609
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtSearchCliente 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   8895
   End
   Begin VB.Label lblIdCliente 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      Height          =   195
      Left            =   10320
      TabIndex        =   12
      Top             =   840
      Width           =   555
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hasta:"
      Height          =   195
      Left            =   5880
      TabIndex        =   11
      Top             =   720
      Width           =   555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desde:"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   675
   End
End
Attribute VB_Name = "frmGuiaRemision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private vBuscar As Boolean 'variable para la busqueda de clientes
Dim loc_key  As Integer

Private Sub cmdconsultar_Click()
MostrarDocumentos
End Sub

Private Sub MostrarDocumentos()
Me.lvCab.ListItems.Clear
    MousePointer = vbHourglass

    On Error GoTo sSearch

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "[dbo].[USP_VENTAS_FILL_CAB]"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCLIENTE", adBigInt, adParamInput, , Me.lblIDcliente.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@INI", adChar, adParamInput, 8, FormatoFecha_yyyyMMdd(Me.MasFechainicio.Text))
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FIN", adChar, adParamInput, 8, FormatoFecha_yyyyMMdd(Me.MasFechaFin.Text))
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)

    Dim orsDatos As ADODB.Recordset

    Set orsDatos = oCmdEjec.Execute

    Dim itemx As Object

    Do While Not orsDatos.EOF
        Set itemx = Me.lvCab.ListItems.Add(, , orsDatos!serie)
        itemx.SubItems(1) = orsDatos!NUMERO
        itemx.SubItems(2) = orsDatos!tipo
        itemx.SubItems(3) = orsDatos!Fecha
        itemx.SubItems(4) = orsDatos!Total
        orsDatos.MoveNext
    Loop
    MousePointer = vbDefault
    Exit Sub
sSearch:
    MousePointer = vbDefault
    MsgBox Err.Description, vbCritical, Pub_Titulo

End Sub

Private Sub cmdProcesar_Click()
Dim sDocumentos As String
Dim itemx As Object

For Each itemx In Me.lvCab.ListItems
    If itemx.Checked Then
        If Len(Trim(sDocumentos)) = 0 Then
        sDocumentos = itemx.SubItems(2) + "-" + itemx.Text + "-" + itemx.SubItems(1)
        Else
        sDocumentos = sDocumentos + "," + itemx.SubItems(2) + "-" + itemx.Text + "-" + itemx.SubItems(1)
        End If
    End If
Next
frmGuiaRemisionMain.strDocumentos = sDocumentos
frmGuiaRemisionMain.lblCliente.Caption = txtSearchCliente.Text
frmGuiaRemisionMain.lblIDcliente.Caption = Me.lblIDcliente.Caption

frmGuiaRemisionMain.Show vbModal
End Sub

Private Sub Form_Load()
 vBuscar = True
 CentrarFormulario MDIForm1, Me
ConfiguraLV
Me.MasFechaFin.Text = LK_FECHA_DIA
Me.MasFechainicio.Text = DateAdd("m", -1, LK_FECHA_DIA)
End Sub

Private Sub ConfiguraLV()
With Me.lvSearch
    .FullRowSelect = True
    .Gridlines = True
    .LabelEdit = lvwManual
    .View = lvwReport
    .ColumnHeaders.Add , , "Codigo", 1000
    .ColumnHeaders.Add , , "Cliente", 5000
    .MultiSelect = False
End With


With Me.lvCab
    .FullRowSelect = True
    .Gridlines = True
    .LabelEdit = lvwManual
    .View = lvwReport
    .ColumnHeaders.Add , , "Serie"
    .ColumnHeaders.Add , , "Número"
    .ColumnHeaders.Add , , "Tipo"
    .ColumnHeaders.Add , , "Fecha"
    .ColumnHeaders.Add , , "Total"
    .MultiSelect = False
End With

With Me.lvDet
    .FullRowSelect = True
    .Gridlines = True
    .LabelEdit = lvwManual
    .View = lvwReport
    .ColumnHeaders.Add , , "Cantidad"
    .ColumnHeaders.Add , , "Producto"
    .ColumnHeaders.Add , , "Peso"
    .ColumnHeaders.Add , , "Peso Total"
    .MultiSelect = False
End With
End Sub



Private Sub lvCab_ItemClick(ByVal item As MSComctlLib.ListItem)
Me.lvDet.ListItems.Clear
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "[dbo].[USP_VENTAS_FILL_DET]"
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCLIENTE", adBigInt, adParamInput, , Me.lblIDcliente.Caption)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SERIE", adChar, adParamInput, 3, item.Text)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMERO", adBigInt, adParamInput, , item.SubItems(1))
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FBG", adChar, adParamInput, 1, item.SubItems(2))
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)

Dim oRSdet As ADODB.Recordset
Set oRSdet = oCmdEjec.Execute
Dim itemx As Object

Do While Not oRSdet.EOF
    Set itemx = Me.lvDet.ListItems.Add(, , oRSdet!Cantidad)
    itemx.SubItems(1) = oRSdet!PRODUCTO
    itemx.SubItems(2) = oRSdet!peso
    itemx.SubItems(3) = oRSdet!pesototal
    oRSdet.MoveNext
Loop

End Sub

Private Sub MasFechaFin_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdConsultar.SetFocus
End Sub

Private Sub MasFechainicio_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Me.MasFechaFin.SetFocus
    Me.MasFechaFin.SelStart = 0
    Me.MasFechaFin.SelLength = Len(Me.MasFechaFin.Text)
End If
End Sub

Private Sub txtSearchCliente_Change()
vBuscar = True
End Sub

Private Sub txtSearchCliente_GotFocus()
buscars = True
End Sub

Private Sub txtSearchCliente_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo SALE

    Dim strFindMe As String

    Dim itmFound  As Object ' ListItem    ' Variable FoundItem.

    If KeyCode = 40 Then  ' flecha abajo
        loc_key = loc_key + 1

        If loc_key > Me.lvSearch.ListItems.count Then loc_key = Me.lvSearch.ListItems.count
        GoTo POSICION
    End If

    If KeyCode = 38 Then
        loc_key = loc_key - 1

        If loc_key < 1 Then loc_key = 1
        GoTo POSICION
    End If

    If KeyCode = 34 Then
        loc_key = loc_key + 17

        If loc_key > Me.lvSearch.ListItems.count Then loc_key = Me.lvSearch.ListItems.count
        GoTo POSICION
    End If

    If KeyCode = 33 Then
        loc_key = loc_key - 17

        If loc_key < 1 Then loc_key = 1
        GoTo POSICION
    End If

    If KeyCode = 27 Then
        Me.lvSearch.Visible = False
'        Me.txtRS.Text = ""
'        Me.txtruc.Text = ""
'        Me.txtdireccion.Text = ""
    End If

    GoTo fin
POSICION:
    Me.lvSearch.ListItems.item(loc_key).Selected = True
    Me.lvSearch.ListItems.item(loc_key).EnsureVisible
    
    'txtRS.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
    'txtRS.SelStart = Len(txtRS.Text)
fin:

    Exit Sub

SALE:
End Sub

Private Sub txtSearchCliente_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)

    If KeyAscii = vbKeyReturn Then
        If vBuscar Then
            Me.lvSearch.ListItems.Clear
            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "[dbo].[USP_CLIENTES_FILL]"
            Set orspago = oCmdEjec.Execute(, Me.txtSearchCliente.Text)

            Dim item As Object
            
            If Not orspago.EOF Then

                Do While Not orspago.EOF
                    Set item = Me.lvSearch.ListItems.Add(, , orspago!Codigo)
                    item.SubItems(1) = Trim(orspago!cliente)
                    orspago.MoveNext
                Loop

                Me.lvSearch.Visible = True
                ' Me.lvSearch.ListItems(1).Selected = True
                loc_key = 1
                ' Me.lvSearch.ListItems(1).EnsureVisible
                vBuscar = False
            Else
                loc_key = -1

            End If
        
        Else
            Me.lblIDcliente.Caption = Me.lvSearch.ListItems(loc_key).Text
            Me.txtSearchCliente.Text = Me.lvSearch.ListItems(loc_key).SubItems(1)
            vBuscar = False
            '            Me.txtruc.Text = Me.ListView1.ListItems(loc_key).SubItems(2)
            '            Me.txtdireccion.Text = Me.ListView1.ListItems(loc_key).SubItems(3)
            '            Me.txtRS.Text = Me.ListView1.ListItems(loc_key).SubItems(1)
            Me.lvSearch.Visible = False

            Me.MasFechainicio.SetFocus
            Me.MasFechainicio.SelStart = 0
            Me.MasFechainicio.SelLength = Len(Me.MasFechaFin.Text)

            '            Me.txtDni.Text = Me.ListView1.ListItems(loc_key).Tag
            '            Me.txtruc.Tag = Me.ListView1.ListItems(loc_key)
            ' Me.lvDetalle.SetFocus
        End If

    End If

End Sub

