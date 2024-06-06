VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form frmGuiaRemisionMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Guia de Remisión"
   ClientHeight    =   9795
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9990
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
   MinButton       =   0   'False
   ScaleHeight     =   9795
   ScaleWidth      =   9990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   720
      Left            =   8520
      Picture         =   "frmGuiaRemisionMain.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   8880
      Width           =   1335
   End
   Begin VB.CheckBox chkDestinatario 
      Caption         =   "Otro Destinatario"
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   7200
      Width           =   2055
   End
   Begin VB.Frame Frame6 
      Height          =   1095
      Left            =   6960
      TabIndex        =   21
      Top             =   120
      Width           =   2895
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   1130
         TabIndex        =   26
         Top             =   525
         Width           =   165
      End
      Begin VB.Label lblNumero 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   1320
         TabIndex        =   25
         Top             =   480
         Width           =   1515
      End
      Begin VB.Label lblSerie 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "T001"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   60
         TabIndex        =   24
         Top             =   480
         Width           =   1035
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número"
         Height          =   195
         Left            =   1440
         TabIndex        =   23
         Top             =   120
         Width           =   675
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Serie"
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   120
         Width           =   450
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1575
      Left            =   120
      TabIndex        =   15
      Top             =   7200
      Width           =   9735
      Begin VB.TextBox txtDestinatarioRazonSocial 
         Enabled         =   0   'False
         Height          =   345
         Left            =   2160
         TabIndex        =   20
         Top             =   1080
         Width           =   7455
      End
      Begin VB.TextBox txtDestinatarioDocumento 
         Enabled         =   0   'False
         Height          =   345
         Left            =   2160
         TabIndex        =   19
         Top             =   720
         Width           =   2175
      End
      Begin VB.ComboBox ComTipoDocumento 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmGuiaRemisionMain.frx":076A
         Left            =   2160
         List            =   "frmGuiaRemisionMain.frx":0777
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Documento:"
         Height          =   195
         Left            =   360
         TabIndex        =   27
         Top             =   420
         Width           =   1470
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Razón Social:"
         Height          =   195
         Left            =   360
         TabIndex        =   18
         Top             =   1155
         Width           =   1170
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nro Documento:"
         Height          =   195
         Left            =   360
         TabIndex        =   17
         Top             =   795
         Width           =   1410
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1935
      Left            =   5040
      TabIndex        =   6
      Top             =   5280
      Width           =   4815
      Begin MSMask.MaskEdBox MasFechaTraslado 
         Height          =   345
         Left            =   2520
         TabIndex        =   31
         Top             =   1440
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSDataListLib.DataCombo datMotivo 
         Height          =   315
         Left            =   360
         TabIndex        =   7
         Top             =   480
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo datModalidad 
         Height          =   315
         Left            =   360
         TabIndex        =   8
         Top             =   1080
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Traslado:"
         Height          =   195
         Left            =   360
         TabIndex        =   35
         Top             =   1560
         Width           =   1365
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Modalidad de Traslado:"
         Height          =   195
         Left            =   360
         TabIndex        =   10
         Top             =   840
         Width           =   1995
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Motivo de Traslado:"
         Height          =   195
         Left            =   360
         TabIndex        =   9
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Transportista"
      Height          =   1935
      Left            =   120
      TabIndex        =   5
      Top             =   5280
      Width           =   4815
      Begin MSDataListLib.DataCombo datTransporte 
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label lblPlacaTransporte 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   4575
      End
      Begin VB.Label lblLicenciaTransporte 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   4575
      End
      Begin VB.Label lblDniTransporte 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   4575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Listado de Productos"
      Height          =   4095
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   9735
      Begin MSComctlLib.ListView lvListado 
         Height          =   3255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   5741
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
      Begin VB.Label lblBultos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   34
         Top             =   3600
         Width           =   1155
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Peso Total:"
         Height          =   195
         Left            =   6480
         TabIndex        =   33
         Top             =   3675
         Width           =   960
      End
      Begin VB.Label lblPesoTotal 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   7440
         TabIndex        =   32
         Top             =   3600
         Width           =   2115
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Cliente"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.Label lblDireccion 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   6435
      End
      Begin VB.Label lblCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6435
      End
   End
   Begin VB.Label lblIDcliente 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Left            =   3480
      TabIndex        =   30
      Top             =   9000
      Width           =   555
   End
End
Attribute VB_Name = "frmGuiaRemisionMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strDocumentos As String



Private Sub chkDestinatario_Click()
Me.ComTipoDocumento.Enabled = Me.chkDestinatario.Value
Me.txtDestinatarioDocumento.Enabled = Me.chkDestinatario.Value
Me.txtDestinatarioRazonSocial.Enabled = Me.chkDestinatario.Value
If Me.chkDestinatario.Value = False Then
    Me.ComTipoDocumento.ListIndex = 0
    Me.txtDestinatarioDocumento.Text = ""
    Me.txtDestinatarioRazonSocial.Text = ""
End If
End Sub

Private Sub cmdGrabar_Click()

    If Me.lvListado.ListItems.count = 0 Then
        MsgBox "Debe agregar articulos para la Guia.", vbCritical, Pub_Titulo
        Exit Sub

    End If

    If Me.datTransporte.BoundText = "-1" Then
        MsgBox "Debe elegir el Transportista.", vbCritical, Pub_Titulo
        Me.datTransporte.SetFocus
        Exit Sub

    End If

    If Len(Trim(Me.lblLicenciaTransporte.Caption)) = 0 Then
        MsgBox "El transportista debe contar con Licencia.", vbCritical, Pub_Titulo
        Exit Sub

    End If

    If Len(Trim(Me.lblDniTransporte.Caption)) = 0 Then
        MsgBox "El transportista debe contar con dni.", vbCritical, Pub_Titulo
        Exit Sub

    End If

    If Len(Trim(Me.lblPlacaTransporte.Caption)) = 0 Then
        MsgBox "El transportista debe contar con Placa.", vbCritical, Pub_Titulo
        Exit Sub

    End If

    If Me.datMotivo.BoundText = "-1" Then
        MsgBox "Debe elegir el Motivo de traslado.", vbCritical, Pub_Titulo
        Me.datMotivo.SetFocus
        Exit Sub

    End If

    If Me.datModalidad.BoundText = "-1" Then
        MsgBox "Debe elegir la Modalidad de Traslado.", vbCritical, Pub_Titulo
        Me.datModalidad.SetFocus
        Exit Sub

    End If

    If Me.chkDestinatario.Value And Me.ComTipoDocumento.ListIndex = 0 Then
        MsgBox "Debe elegir el destinatario.", vbCritical, Pub_Titulo
        Me.ComTipoDocumento.SetFocus
        Exit Sub

    End If

    If Me.chkDestinatario.Value And Len(Trim(Me.txtDestinatarioDocumento.Text)) = 0 Then
        MsgBox "Debe ingresar el Número de Documento del destinatario.", vbCritical, Pub_Titulo
        Me.txtDestinatarioDocumento.SetFocus
        Exit Sub

    End If

    If Me.chkDestinatario.Value And Len(Trim(Me.txtDestinatarioRazonSocial.Text)) = 0 Then
        MsgBox "Debe ingresar la Razón Social del destinatario.", vbCritical, Pub_Titulo
        Me.txtDestinatarioRazonSocial.SetFocus
        Exit Sub

    End If

    If Not IsDate(Me.MasFechaTraslado.Text) Then
        MsgBox "Debe ingresar una fecha de Traslado correcta.", vbCritical, Pub_Titulo
        Me.MasFechaTraslado.SetFocus
        Me.MasFechaTraslado.SelStart = 0
        Me.MasFechaTraslado.SelLength = Len(Me.MasFechaTraslado.Text)
Exit Sub
    End If

    grabarGuia

End Sub

Private Sub grabarGuia()

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "[dbo].[USP_GUIA_REGISTRAR]"
    
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCLIENTE", adBigInt, adParamInput, , Me.lblIDcliente.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SERIE", adChar, adParamInput, 4, Me.lblSerie.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMERO", adBigInt, adParamInput, , Me.lblNumero.Caption)
    
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHAEMISION", adChar, adParamInput, 8, FormatoFecha_yyyyMMdd(CStr(LK_FECHA_DIA)))
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHATRASLADO", adChar, adParamInput, 8, FormatoFecha_yyyyMMdd(Me.MasFechaTraslado.Text))
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODIGOMOTIVOTRASLADO", adChar, adParamInput, 2, Me.datMotivo.BoundText)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODIGOMODALIDADTRASLADO", adChar, adParamInput, 2, Me.datModalidad.BoundText)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDTRASPORTISTA", adInteger, adParamInput, , Me.datTransporte.BoundText)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PESO", adDouble, adParamInput, , CDbl(Me.lblPesoTotal.Caption))
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@BULTOS", adInteger, adParamInput, , Me.lblBultos.Caption)
    
    Dim itemx  As Object

    Dim xDatos As String

    xDatos = ""
    
    If Me.lvListado.ListItems.count <> 0 Then
        xDatos = "<r>"

        For Each itemx In Me.lvListado.ListItems

            xDatos = xDatos & "<d "
            xDatos = xDatos & "idp=""" & itemx.Tag & """ "
            xDatos = xDatos & "cant=""" & itemx.Text & """ "
            xDatos = xDatos & "/>"
        Next
        xDatos = xDatos & "</r>"

    End If
    
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PRODUCTOS", adLongVarChar, adParamInput, Len(xDatos), xDatos)
    
    If Me.chkDestinatario.Value Then
        If Me.ComTipoDocumento.ListIndex = 1 Then
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DESTINATARIOTIPO", adChar, adParamInput, 1, "1")
        Else
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DESTINATARIOTIPO", adChar, adParamInput, 1, "6")

        End If

        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DESTINATARIONUMERO", adVarChar, adParamInput, 11, Me.txtDestinatarioDocumento.Text)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DESTINATARIORS", adVarChar, adParamInput, 100, Me.txtDestinatarioRazonSocial.Text)

    End If

    Dim oRSsave As ADODB.Recordset

    Set oRSsave = oCmdEjec.Execute

    Dim arr() As String

    If Not oRSsave.EOF Then
        arr = Split(oRSsave.Fields(0).Value, "=")
        
        If UBound(arr) - LBound(arr) = 1 Then
            If arr(0) = "0" Then
                MsgBox arr(1), vbInformation, Pub_Titulo
                Unload Me
            Else
                MsgBox arr(0), vbCritical, Pub_Titulo
            End If
        
        Else
            MsgBox "No se puede identificar el mensaje de error."

        End If
        

    End If
    
End Sub

Private Sub datTransporte_Change()
If Me.datTransporte.BoundText = "" Then Exit Sub
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "[dbo].[USP_TRANSPORTISTA_DATOS]"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDTRANSPORTE", adBigInt, adParamInput, , Me.datTransporte.BoundText)
    
    Dim oRSda As ADODB.Recordset
    Set oRSda = oCmdEjec.Execute
    
    If Not oRSda.EOF Then
        Me.lblDniTransporte.Caption = oRSda!dni
        Me.lblLicenciaTransporte.Caption = oRSda!licencia
        Me.lblPlacaTransporte.Caption = oRSda!placa
    End If
End Sub

Private Sub Form_Load()
Me.MasFechaTraslado.Text = LK_FECHA_DIA
    ConfigurarLv
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "[dbo].[USP_VENTAS_DATOS]"

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@VALORES", adLongVarChar, adParamInput, Len(strDocumentos), strDocumentos)

    Dim orsDatos As ADODB.Recordset

    Set orsDatos = oCmdEjec.Execute

    Dim itemx   As Object

    Dim vBultos As Integer

    Dim vPeso   As Double
    
    vPeso = 0
    vBultos = 0

    Dim vEncontrado As Boolean

    vEncontrado = False

    Do While Not orsDatos.EOF

        For Each itemx In Me.lvListado.ListItems

            If itemx.Tag = orsDatos!idproducto Then
                vEncontrado = True
                Exit For

            End If

        Next
        
        If vEncontrado = False Then
            Set itemx = Me.lvListado.ListItems.Add(, , orsDatos!Cantidad)
            itemx.Tag = orsDatos!idproducto
            itemx.SubItems(1) = orsDatos!PRODUCTO
            itemx.SubItems(2) = orsDatos!peso
            itemx.SubItems(3) = orsDatos!pesototal
        Else
            itemx.Text = Val(itemx.Text) + orsDatos!Cantidad
            itemx.SubItems(3) = Val(itemx.SubItems(3)) + orsDatos!pesototal
        End If

        vBultos = vBultos + orsDatos!Cantidad
        vPeso = vPeso + orsDatos!pesototal
        orsDatos.MoveNext
        vEncontrado = False
    Loop
    
    Me.lblBultos.Caption = vBultos
    Me.lblPesoTotal.Caption = vPeso
    CargarCombos
    obtenerNumeracion

End Sub

Private Sub obtenerNumeracion()
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "[dbo].[USP_NUMERACION_DOCUMENTOS]"
 oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TIPO", adChar, adParamInput, 1, "G")
    
    Dim oRSnum As ADODB.Recordset
    
    Set oRSnum = oCmdEjec.Execute
    
    If Not oRSnum.EOF Then
        Me.lblNumero.Caption = oRSnum!NUMERO
    End If
    
End Sub
Private Sub ConfigurarLv()
With Me.lvListado
    .FullRowSelect = True
    .Gridlines = True
    .LabelEdit = lvwManual
    .View = lvwReport
    .ColumnHeaders.Add , , "Cantidad"
    .ColumnHeaders.Add , , "Producto", 5000
    .ColumnHeaders.Add , , "Peso"
    .ColumnHeaders.Add , , "Peso Total"
    .MultiSelect = False
End With
End Sub

Private Sub CargarCombos()
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "[dbo].[USP_MOTIVO_MODALIDAD_LIST]"
       oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)

    Dim orsDatos As ADODB.Recordset

    Set orsDatos = oCmdEjec.Execute
    Set Me.datModalidad.RowSource = orsDatos
    datModalidad.ListField = "DESMODALIDAD"
    datModalidad.BoundColumn = "CODMODALIDAD"
    datModalidad.BoundText = orsDatos!CODMODALIDAD
    
    Dim orsTEMP As ADODB.Recordset
    
    Set orsTEMP = orsDatos.NextRecordset
    
    Set Me.datMotivo.RowSource = orsTEMP
    datMotivo.ListField = "DESMOTIVO"
    datMotivo.BoundColumn = "CODMOTIVO"
    datMotivo.BoundText = orsTEMP!CODMOTIVO
    
    Set orsTEMP = orsTEMP.NextRecordset
    Set Me.datTransporte.RowSource = orsTEMP
    datTransporte.ListField = "NOMTRANSPORTE"
    datTransporte.BoundColumn = "CODTRANSPORTE"
    datTransporte.BoundText = -1

End Sub
