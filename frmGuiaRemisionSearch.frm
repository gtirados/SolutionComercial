VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmGuiaRemisionSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consultas de Guia de Remisión Electrónicas"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12735
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
   ScaleWidth      =   12735
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   11535
      Begin MSComctlLib.ListView lvCab 
         Height          =   2295
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   195
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
   Begin VB.Frame Frame2 
      Height          =   2655
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   11535
      Begin MSComctlLib.ListView lvDet 
         Height          =   2295
         Left            =   120
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   195
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
   Begin VB.CommandButton cmdConsultar 
      Caption         =   "Consultar"
      Height          =   720
      Left            =   8760
      Picture         =   "frmGuiaRemisionSearch.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   990
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   10440
      Top             =   5640
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton cmdPdf 
      Height          =   600
      Left            =   11760
      Picture         =   "frmGuiaRemisionSearch.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1440
      Width           =   855
   End
   Begin MSMask.MaskEdBox MasFechainicio 
      Height          =   345
      Left            =   1080
      TabIndex        =   0
      Top             =   405
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   609
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MasFechaFin 
      Height          =   345
      Left            =   6720
      TabIndex        =   1
      Top             =   405
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   609
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desde:"
      Height          =   195
      Left            =   360
      TabIndex        =   9
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hasta:"
      Height          =   195
      Left            =   6120
      TabIndex        =   8
      Top             =   480
      Width           =   555
   End
End
Attribute VB_Name = "frmGuiaRemisionSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConsultar_Click()
MostrarDocumentos
End Sub

Private Sub cmdPdf_Click()

    If Me.lvCab.SelectedItem.SubItems(5) = "" Then
        MsgBox "No puede visualizar una Guia sin Respuesta por SUNAT.", vbCritical, Pub_Titulo
        Exit Sub

    End If

    On Error GoTo xFile

    MousePointer = vbHourglass

    ' Ruta del archivo PDF en el servidor
    Dim sURL As String, sRUC As String, sNombreArchivo As String, sCarpeta As String
    
    sURL = Leer_Ini(App.Path & "\config.ini", "URL", "http://51.89.237.222/gtsoftware/")
    sRUC = Leer_Ini(App.Path & "\config.ini", "RUC", "20559765381")
    sCarpeta = Leer_Ini(App.Path & "\config.ini", "CARPETA", "c:\")
    sNombreArchivo = sRUC + "-09-" + Me.lvCab.SelectedItem.Text + "-" + Me.lvCab.SelectedItem.SubItems(1)
    
    Dim URL As String

    URL = sURL + "files/guia_electronica/PDF/" + sNombreArchivo + ".pdf"
    
    ' Ruta donde se guardará el archivo en tu máquina local
    Dim archivoLocal As String

    archivoLocal = sCarpeta + sNombreArchivo + ".pdf"
    
    ' Llama a la función para descargar el archivo
    If DescargarArchivo(URL, archivoLocal) Then
        MsgBox "Archivo descargado con éxito."
        frmGuiaRemisionPDF.xRuta = archivoLocal
        frmGuiaRemisionPDF.Show vbModal
    Else
        MsgBox "Error al descargar el archivo."

    End If

    MousePointer = vbDefault
    Exit Sub
xFile:
    MsgBox Err.Description, vbCritical, Pub_Titulo

End Sub


Private Function DescargarArchivo(URL As String, archivoLocal As String) As Boolean
    On Error GoTo ErrorHandler

    ' Configura el control Inet
    Inet1.Cancel ' Cancela cualquier operación en curso
    Inet1.Protocol = icHTTP ' Usa el protocolo HTTP
    Inet1.URL = URL ' Especifica la URL del archivo

    ' Descarga el archivo
    Dim datos() As Byte
    datos = Inet1.OpenURL(URL, icByteArray)
    
    ' Guarda el archivo en la ruta especificada
    Dim numFile As Integer
    numFile = FreeFile
    Open archivoLocal For Binary Access Write As numFile
    Put numFile, , datos
    Close numFile
    
    DescargarArchivo = True
    Exit Function

ErrorHandler:
    DescargarArchivo = False
End Function



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
ConfiguraLV
Me.MasFechaFin.Text = LK_FECHA_DIA
Me.MasFechainicio.Text = DateAdd("m", -1, LK_FECHA_DIA)
End Sub


Private Sub ConfiguraLV()

With Me.lvCab
    .FullRowSelect = True
    .Gridlines = True
    .LabelEdit = lvwManual
    .View = lvwReport
    .ColumnHeaders.Add , , "Serie", 700
    .ColumnHeaders.Add , , "Número", 1000
    .ColumnHeaders.Add , , "Cliente", 5000
    .ColumnHeaders.Add , , "Fecha"
    .ColumnHeaders.Add , , "Total"
    .ColumnHeaders.Add , , "Rpta SUNAT"
    .MultiSelect = False
End With

With Me.lvDet
    .FullRowSelect = True
    .Gridlines = True
    .LabelEdit = lvwManual
    .View = lvwReport
    .ColumnHeaders.Add , , "Cantidad", 1000
    .ColumnHeaders.Add , , "Producto", 7500
    .ColumnHeaders.Add , , "Peso", 1000
    .ColumnHeaders.Add , , "Peso Total", 1200
    .MultiSelect = False
End With
End Sub

Private Sub MostrarDocumentos()
    Me.lvCab.ListItems.Clear
    MousePointer = vbHourglass

    On Error GoTo sSearch

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "[dbo].[USP_GUIA_REMISION_FILL_CAB]"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@INI", adChar, adParamInput, 8, FormatoFecha_yyyyMMdd(Me.MasFechainicio.Text))
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FIN", adChar, adParamInput, 8, FormatoFecha_yyyyMMdd(Me.MasFechaFin.Text))
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)

    Dim orsDatos As ADODB.Recordset

    Set orsDatos = oCmdEjec.Execute

    Dim itemx As Object

    Do While Not orsDatos.EOF
        Set itemx = Me.lvCab.ListItems.Add(, , orsDatos!serie)
        itemx.SubItems(1) = orsDatos!NUMERO
        itemx.SubItems(2) = orsDatos!CLIENTE
        itemx.SubItems(3) = orsDatos!Fecha
        itemx.SubItems(4) = orsDatos!PESOTOTAL
        itemx.SubItems(5) = orsDatos!RPTASUNAT
        orsDatos.MoveNext
    Loop
    MousePointer = vbDefault
    Exit Sub
sSearch:
    MousePointer = vbDefault
    MsgBox Err.Description, vbCritical, Pub_Titulo

End Sub

Private Sub lvCab_ItemClick(ByVal Item As MSComctlLib.ListItem)
MostrarDetalle Item.Text, Item.SubItems(1)
End Sub

Private Sub MostrarDetalle(cSerie As String, cNumero As Integer)
Me.lvDet.ListItems.Clear
    MousePointer = vbHourglass

    On Error GoTo sSearch

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "[dbo].[USP_GUIA_REMISION_FILL_DET]"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SERIE", adChar, adParamInput, 4, cSerie)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMERO", adBigInt, adParamInput, , cNumero)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)

    Dim orsDatos As ADODB.Recordset

    Set orsDatos = oCmdEjec.Execute

    Dim itemx As Object

    Do While Not orsDatos.EOF
        Set itemx = Me.lvDet.ListItems.Add(, , orsDatos!Cantidad)
        itemx.SubItems(1) = orsDatos!PRODUCTO
        itemx.SubItems(2) = orsDatos!PESO
        itemx.SubItems(3) = orsDatos!PESOTOTAL
        orsDatos.MoveNext
    Loop
    MousePointer = vbDefault
    Exit Sub
sSearch:
    MousePointer = vbDefault
    MsgBox Err.Description, vbCritical, Pub_Titulo
End Sub
