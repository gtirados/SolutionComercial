Attribute VB_Name = "ModDatos"
Public ParaLot_count As Integer
Public ParaLot_codlot(100) As String
Public ParaLot_lotcant(100) As Currency
Public ParaLot_lotchasis(100) As String
Public ParaLot_lotproced(100) As String
Public ParaLot_lotanio(100) As String
Public ParaLot_lotpoliza(100) As String
'Public ParaLot_fechalot(100) As String

Public Sub TODO_DATOS()
'Dim JALA_LOTE As rdoResultset
Dim LT_SALDO As Currency
Dim xcuenta As Integer
Dim xcuenta2 As Integer
Dim LT_CANTIDAD As Currency
Dim pasa_act As String * 1
With FORMGEN
.det_lot.ColWidth(0) = 0 ' codigo de cia
.det_lot.ColWidth(1) = 400 ' codigo Interno  de Articulo
.det_lot.ColWidth(2) = 400 ' codigo nro motor
.det_lot.ColWidth(3) = 400 ' cantidad
.det_lot.ColWidth(4) = 200 ' fila del grid_fac
.det_lot.ColWidth(5) = 200 ' stock de lote

For xcuenta = 2 To .grid_fac.Rows - 1
  PUB_CODART = Val(.grid_fac.TextMatrix(xcuenta, 16))
  If PUB_CODART = 0 Then Exit For
  If Val(.grid_fac.TextMatrix(xcuenta, 47)) = 1 And Val(.grid_fac.TextMatrix(xcuenta, 48)) = 1 Then GoTo sigue
      LT_CANTIDAD = Val(.grid_fac.TextMatrix(xcuenta, 4))
      CONFIGURAR_DATOS LK_CODCIA, Val(.grid_fac.TextMatrix(xcuenta, 16)), LT_CANTIDAD, .grid_fac.TextMatrix(xcuenta, 5), Val(.grid_fac.TextMatrix(xcuenta, 14)), xcuenta, 1
      .vfila = xcuenta
      .vViene = True
      .cmdltaceptar_Click
      .vViene = False
    '  pasa_act = ""
    '  GoSub chequeo_lista ' chequeo que si esta en la lista
    '  If Trim(pasa_act) <> "A" Then GoTo sigue_for
    '  pub_cadena = "SELECT * FROM LOTE WHERE LOT_CODCIA = '" & LK_CODCIA & "' AND LOT_CODART = " & PUB_CODART & " AND LOT_SALDOS <> 0  ORDER BY LOT_FECHA_VCTO"
    '  Set JALA_LOTE = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
    '  LT_CANTIDAD = Val(grid_fac.TextMatrix(xcuenta, 4))
    '  LT_SALDO = 0
    '  Do Until JALA_LOTE.EOF
    '    If LT_CANTIDAD <= 0 Then Exit Do
    '    If LT_CANTIDAD <= JALA_LOTE!LOt_SALDOS Then
    '       LT_SALDO = LT_CANTIDAD
    '       LT_CANTIDAD = 0
    '       det_lot.Rows = det_lot.Rows + 1
    '       det_lot.TextMatrix(det_lot.Rows - 1, 0) = LK_CODCIA
    '       det_lot.TextMatrix(det_lot.Rows - 1, 1) = PUB_CODART
    '       det_lot.TextMatrix(det_lot.Rows - 1, 2) = Trim(JALA_LOTE!LOT_NROLOTE)
    '       det_lot.TextMatrix(det_lot.Rows - 1, 3) = LT_SALDO
    '       det_lot.TextMatrix(det_lot.Rows - 1, 4) = xcuenta
    '       det_lot.TextMatrix(det_lot.Rows - 1, 5) = JALA_LOTE!LOt_SALDOS
    '       det_lot.TextMatrix(det_lot.Rows - 1, 6) = JALA_LOTE!lot_fecha_vcto
    '    Else
    '       LT_SALDO = JALA_LOTE!LOt_SALDOS
    '       det_lot.Rows = det_lot.Rows + 1
    '       det_lot.TextMatrix(det_lot.Rows - 1, 0) = LK_CODCIA
    '       det_lot.TextMatrix(det_lot.Rows - 1, 1) = PUB_CODART
    '       det_lot.TextMatrix(det_lot.Rows - 1, 2) = Trim(JALA_LOTE!LOT_NROLOTE)
    '       det_lot.TextMatrix(det_lot.Rows - 1, 3) = LT_SALDO
    '       det_lot.TextMatrix(det_lot.Rows - 1, 4) = xcuenta
    '       det_lot.TextMatrix(det_lot.Rows - 1, 5) = JALA_LOTE!LOt_SALDOS
    '       det_lot.TextMatrix(det_lot.Rows - 1, 6) = JALA_LOTE!lot_fecha_vcto
    '       LT_CANTIDAD = LT_CANTIDAD - JALA_LOTE!LOt_SALDOS
    '    End If
    '    JALA_LOTE.MoveNext
    '  Loop
  'End If
'sigue_for:
sigue:
Next xcuenta
End With
Exit Sub
' chequeo Lista
'chequeo_lista:
'With FORMGEN
'pasa_act = "A"
'
'For xcuenta2 = 1 To .det_lot.Rows - 1
'   If Val(.det_lot.TextMatrix(xcuenta2, 1)) = PUB_CODART And Val(.det_lot.TextMatrix(xcuenta2, 4)) = xcuenta Then
'      pasa_act = ""
'      Exit For
'   End If
'Next xcuenta2
'End With
'Return
End Sub


Public Sub verif_datos(lot_codart As Currency, lot_fila As Integer)
Dim xcuenta As Integer
For xcuenta = 1 To 100
 ParaLot_codlot(xcuenta) = ""
 ParaLot_lotcant(xcuenta) = 0
 ParaLot_lotchasis(xcuenta) = ""
 ParaLot_lotproced(xcuenta) = ""
 ParaLot_lotanio(xcuenta) = ""
 ParaLot_lotpoliza(xcuenta) = ""
 'ParaLot_fechalot(xcuenta) = 0
Next xcuenta
ParaLot_count = 0
With FORMGEN
For xcuenta = 1 To .det_lot.Rows - 1
 If lot_codart = Val(.det_lot.TextMatrix(xcuenta, 1)) And lot_fila = Val(.det_lot.TextMatrix(xcuenta, 4)) Then
    If Val(.det_lot.TextMatrix(xcuenta, 3)) <> 0 Then
     ParaLot_count = ParaLot_count + 1
     ParaLot_codlot(ParaLot_count) = Trim(.det_lot.TextMatrix(xcuenta, 2))
     ParaLot_lotcant(ParaLot_count) = Format(.det_lot.TextMatrix(xcuenta, 3), "0.00")
     ParaLot_lotchasis(ParaLot_count) = Trim(.det_lot.TextMatrix(xcuenta, 10))
     ParaLot_lotproced(ParaLot_count) = Trim(.det_lot.TextMatrix(xcuenta, 11))
     ParaLot_lotanio(ParaLot_count) = Trim(.det_lot.TextMatrix(xcuenta, 12))
     ParaLot_lotpoliza(ParaLot_count) = Trim(.det_lot.TextMatrix(xcuenta, 13))
     'ParaLot_fechalot(ParaLot_count) = .det_lot.TextMatrix(xcuenta, 6)
    End If
 End If
Next xcuenta
End With
End Sub
Public Sub verif_datos1(lot_codart As Currency, lot_fila As Integer)
Dim xcuenta As Integer
For xcuenta = 1 To 100
 ParaLot_codlot(xcuenta) = ""
 ParaLot_lotcant(xcuenta) = 0
Next xcuenta
ParaLot_count = 0
With FORMGEN
For xcuenta = 1 To .det_lot1.Rows - 1
 If lot_codart = Val(.det_lot1.TextMatrix(xcuenta, 1)) And lot_fila = Val(.det_lot1.TextMatrix(xcuenta, 4)) Then
    If Val(.det_lot1.TextMatrix(xcuenta, 3)) <> 0 Then
     ParaLot_count = ParaLot_count + 1
     ParaLot_codlot(ParaLot_count) = Trim(.det_lot1.TextMatrix(xcuenta, 2))
     ParaLot_lotcant(ParaLot_count) = Format(.det_lot1.TextMatrix(xcuenta, 3), "0.00")
     ParaLot_lotchasis(ParaLot_count) = Trim(.det_lot1.TextMatrix(xcuenta, 14))
    End If
 End If
Next xcuenta
End With
End Sub
Public Function CHEQUEO_DETALLE() As String
' CHEQUEO PARA VER SI ESTAN TODOS LO ARTICULOS EDITAR SOLO EN INGRESOS

Dim xcuenta As Integer
Dim xcuenta2 As Integer
Dim pasa_act As String
Dim FILA_GRID As Integer
'Dim NRO_LOTE  As String
Dim wencuentra As String * 1
CHEQUEO_DETALLE = ""
wencuentra = "A"
pasa_act = ""
With FORMGEN
For xcuenta = 2 To .grid_fac.Rows - 1
    PUB_CODART = Val(.grid_fac.TextMatrix(xcuenta, 16))
    FILA_GRID = xcuenta
    If PUB_CODART = 0 Then
     Exit For
    End If
    If Val(.grid_fac.TextMatrix(xcuenta, 47)) = 1 And Val(.grid_fac.TextMatrix(xcuenta, 48)) = 1 Then GoTo sigue
    If Val(.grid_fac.TextMatrix(xcuenta, 47)) = 2 Or _
        (Val(.grid_fac.TextMatrix(xcuenta, 47)) = 1 And Val(.grid_fac.TextMatrix(xcuenta, 48)) <> 2) Then
        For xcuenta2 = 1 To .det_lot1.Rows - 1
          If PUB_CODART = Val(.det_lot1.TextMatrix(xcuenta2, 1)) And FILA_GRID = Val(.det_lot1.TextMatrix(xcuenta2, 4)) Then
            wencuentra = "A"
            pasa_act = ""
            Exit For
          Else
            wencuentra = ""
            pasa_act = "Fila:" & Format(xcuenta - 1, "00") & " " & Trim(.grid_fac.TextMatrix(xcuenta, 0))
          End If
        Next xcuenta2
        If .det_lot1.Rows <= 1 Then
           wencuentra = ""
           pasa_act = "Todos."
        End If
    Else
        'If gridlt.TextMatrix(xcuenta, 3) = 0 Then GoTo sigue_for
        For xcuenta2 = 1 To .det_lot.Rows - 1
          If PUB_CODART = Val(.det_lot.TextMatrix(xcuenta2, 1)) And FILA_GRID = Val(.det_lot.TextMatrix(xcuenta2, 4)) Then
            wencuentra = "A"
            pasa_act = ""
            Exit For
          Else
            wencuentra = ""
            pasa_act = "Fila:" & Format(xcuenta - 1, "00") & " " & Trim(.grid_fac.TextMatrix(xcuenta, 0))
          End If
        Next xcuenta2
        If .det_lot.Rows <= 1 Then
           wencuentra = ""
           pasa_act = "Todos."
        End If
    End If
sigue:
Next xcuenta
If wencuentra <> "A" Then
  CHEQUEO_DETALLE = pasa_act
End If
End With
End Function
Public Sub PASA_BOT_ACEPTAR()
' PASA DE LA MUESTRA AL DETALLATE AL PULSAR EL BOTON ACEPTAR
'Dim JALA_DATO As rdoResultset
Dim LT_SALDO As Currency
Dim xcuenta As Integer
Dim xcuenta2 As Integer
Dim LT_CANTIDAD As Currency
Dim pasa_act As String * 1
Dim FILA_GRID As Integer
Dim NRO_MOTOR  As String
Dim wencuentra As String * 1

With FORMGEN
    If Val(.grid_fac.TextMatrix(.grid_fac.Row, 47)) = 2 Or _
            (Val(.grid_fac.TextMatrix(.grid_fac.Row, 47)) = 1 And Val(.grid_fac.TextMatrix(.grid_fac.Row, 48)) <> 2) Then
            
        For xcuenta = 2 To .gridlt.Rows - 1
            PUB_CODART = Val(.gridlt.TextMatrix(xcuenta, 9))
            FILA_GRID = Val(.gridlt.TextMatrix(xcuenta, 10))
            NRO_MOTOR = Trim(.gridlt.TextMatrix(xcuenta, 12))
            If PUB_CODART = 0 Then Exit For
            'If gridlt.TextMatrix(xcuenta, 3) = 0 Then GoTo sigue_for
            wencuentra = ""
        
            For xcuenta2 = 1 To .det_lot1.Rows - 1
              If PUB_CODART = Val(.det_lot1.TextMatrix(xcuenta2, 1)) And FILA_GRID = Val(.det_lot1.TextMatrix(xcuenta2, 4)) And NRO_MOTOR = Trim(.det_lot1.TextMatrix(xcuenta2, 9)) Then
                NRO_MOTOR = Trim(.gridlt.TextMatrix(xcuenta, 0))
                .gridlt.TextMatrix(xcuenta, 12) = NRO_MOTOR
                .det_lot1.TextMatrix(xcuenta2, 2) = NRO_MOTOR
                .det_lot1.TextMatrix(xcuenta2, 9) = NRO_MOTOR
                .det_lot1.TextMatrix(xcuenta2, 3) = Val(.gridlt.TextMatrix(xcuenta, 3)) * Val(.gridlt.TextMatrix(xcuenta, 11))
                '.det_lot.TextMatrix(xcuenta2, 6) = Trim(.gridlt.TextMatrix(xcuenta, 4))
                .det_lot1.TextMatrix(xcuenta2, 10) = Trim(.gridlt.TextMatrix(xcuenta, 1))
                .det_lot1.TextMatrix(xcuenta2, 11) = Trim(.gridlt.TextMatrix(xcuenta, 6))
                .det_lot1.TextMatrix(xcuenta2, 12) = .gridlt.TextMatrix(xcuenta, 7)
                .det_lot1.TextMatrix(xcuenta2, 13) = Trim(.gridlt.TextMatrix(xcuenta, 8))
                .det_lot1.TextMatrix(xcuenta2, 14) = Trim(.gridlt.TextMatrix(xcuenta, 13))
                wencuentra = "A"
                Exit For
              End If
            Next xcuenta2
            If wencuentra <> "A" Then
               .det_lot1.Rows = .det_lot1.Rows + 1
               .det_lot1.TextMatrix(.det_lot1.Rows - 1, 0) = LK_CODCIA
               .det_lot1.TextMatrix(.det_lot1.Rows - 1, 1) = PUB_CODART
               .det_lot1.TextMatrix(.det_lot1.Rows - 1, 2) = NRO_MOTOR
               .det_lot1.TextMatrix(.det_lot1.Rows - 1, 3) = Val(.gridlt.TextMatrix(xcuenta, 3)) * Val(.gridlt.TextMatrix(xcuenta, 11))
               .det_lot1.TextMatrix(.det_lot1.Rows - 1, 4) = FILA_GRID
               .det_lot1.TextMatrix(.det_lot1.Rows - 1, 5) = Val(.gridlt.TextMatrix(xcuenta, 2)) * Val(.gridlt.TextMatrix(xcuenta, 11))
               '.det_lot1.TextMatrix(.det_lot1.Rows - 1, 6) = Trim(.gridlt.TextMatrix(xcuenta, 4))
               .det_lot1.TextMatrix(.det_lot1.Rows - 1, 7) = Trim(.gridlt.TextMatrix(xcuenta, 5))
               .det_lot1.TextMatrix(.det_lot1.Rows - 1, 8) = Trim(.gridlt.TextMatrix(xcuenta, 11))
               .det_lot1.TextMatrix(.det_lot1.Rows - 1, 9) = NRO_MOTOR
               .det_lot1.TextMatrix(.det_lot1.Rows - 1, 10) = Trim(.gridlt.TextMatrix(xcuenta, 1))
               .det_lot1.TextMatrix(.det_lot1.Rows - 1, 11) = Trim(.gridlt.TextMatrix(xcuenta, 6))
               .det_lot1.TextMatrix(.det_lot1.Rows - 1, 12) = .gridlt.TextMatrix(xcuenta, 7)
               .det_lot1.TextMatrix(.det_lot1.Rows - 1, 13) = Trim(.gridlt.TextMatrix(xcuenta, 8))
               .det_lot1.TextMatrix(.det_lot1.Rows - 1, 14) = Trim(.gridlt.TextMatrix(xcuenta, 13))
            End If
        Next xcuenta
    Else
        For xcuenta = 2 To .gridlt.Rows - 1
            PUB_CODART = Val(.gridlt.TextMatrix(xcuenta, 9))
            FILA_GRID = Val(.gridlt.TextMatrix(xcuenta, 10))
            NRO_MOTOR = Trim(.gridlt.TextMatrix(xcuenta, 12))
            If PUB_CODART = 0 Then Exit For
            'If gridlt.TextMatrix(xcuenta, 3) = 0 Then GoTo sigue_for
            wencuentra = ""
                 
            For xcuenta2 = 1 To .det_lot.Rows - 1
              If PUB_CODART = Val(.det_lot.TextMatrix(xcuenta2, 1)) And FILA_GRID = Val(.det_lot.TextMatrix(xcuenta2, 4)) And NRO_MOTOR = Trim(.det_lot.TextMatrix(xcuenta2, 9)) Then
                NRO_MOTOR = Trim(.gridlt.TextMatrix(xcuenta, 0))
                .gridlt.TextMatrix(xcuenta, 12) = NRO_MOTOR
                .det_lot.TextMatrix(xcuenta2, 2) = NRO_MOTOR
                .det_lot.TextMatrix(xcuenta2, 9) = NRO_MOTOR
                .det_lot.TextMatrix(xcuenta2, 3) = Val(.gridlt.TextMatrix(xcuenta, 3)) * Val(.gridlt.TextMatrix(xcuenta, 11))
                '.det_lot.TextMatrix(xcuenta2, 6) = Trim(.gridlt.TextMatrix(xcuenta, 4))
                .det_lot.TextMatrix(xcuenta2, 10) = Trim(.gridlt.TextMatrix(xcuenta, 1))
                .det_lot.TextMatrix(xcuenta2, 11) = Trim(.gridlt.TextMatrix(xcuenta, 6))
                .det_lot.TextMatrix(xcuenta2, 12) = .gridlt.TextMatrix(xcuenta, 7)
                .det_lot.TextMatrix(xcuenta2, 13) = Trim(.gridlt.TextMatrix(xcuenta, 8))
                wencuentra = "A"
                Exit For
              End If
            Next xcuenta2
            If wencuentra <> "A" Then
               .det_lot.Rows = .det_lot.Rows + 1
               .det_lot.TextMatrix(.det_lot.Rows - 1, 0) = LK_CODCIA
               .det_lot.TextMatrix(.det_lot.Rows - 1, 1) = PUB_CODART
               .det_lot.TextMatrix(.det_lot.Rows - 1, 2) = NRO_MOTOR
               .det_lot.TextMatrix(.det_lot.Rows - 1, 3) = Val(.gridlt.TextMatrix(xcuenta, 3)) * Val(.gridlt.TextMatrix(xcuenta, 11))
               .det_lot.TextMatrix(.det_lot.Rows - 1, 4) = FILA_GRID
               .det_lot.TextMatrix(.det_lot.Rows - 1, 5) = Val(.gridlt.TextMatrix(xcuenta, 2)) * Val(.gridlt.TextMatrix(xcuenta, 11))
               '.det_lot.TextMatrix(.det_lot.Rows - 1, 6) = Trim(.gridlt.TextMatrix(xcuenta, 4))
               .det_lot.TextMatrix(.det_lot.Rows - 1, 7) = Trim(.gridlt.TextMatrix(xcuenta, 5))
               .det_lot.TextMatrix(.det_lot.Rows - 1, 8) = Trim(.gridlt.TextMatrix(xcuenta, 11))
               .det_lot.TextMatrix(.det_lot.Rows - 1, 9) = NRO_MOTOR
               .det_lot.TextMatrix(.det_lot.Rows - 1, 10) = Trim(.gridlt.TextMatrix(xcuenta, 1))
               .det_lot.TextMatrix(.det_lot.Rows - 1, 11) = Trim(.gridlt.TextMatrix(xcuenta, 6))
               .det_lot.TextMatrix(.det_lot.Rows - 1, 12) = .gridlt.TextMatrix(xcuenta, 7)
               .det_lot.TextMatrix(.det_lot.Rows - 1, 13) = Trim(.gridlt.TextMatrix(xcuenta, 8))
            End If
    'sigue_for:
        Next xcuenta
    End If
End With
End Sub

Public Sub ASIGNA_NEW_DATO(LT_CODART As Currency, LT_CANTIDAD As Currency, LT_DESCRIP As String, LT_EQUIV As Integer, FILA_GRID As Integer) ', NREG As Integer)
With FORMGEN
    Dim J As Integer
    'For J = 1 To NREG
    If Val(.grid_fac.TextMatrix(.grid_fac.Row, 47)) = 2 Or _
            (Val(.grid_fac.TextMatrix(.grid_fac.Row, 47)) = 1 And Val(.grid_fac.TextMatrix(.grid_fac.Row, 48)) <> 2) Then
        
        .det_lot1.Rows = .det_lot1.Rows + 1
        .det_lot1.TextMatrix(.det_lot1.Rows - 1, 0) = LK_CODCIA
        .det_lot1.TextMatrix(.det_lot1.Rows - 1, 1) = LT_CODART
        .det_lot1.TextMatrix(.det_lot1.Rows - 1, 2) = "(*)" 'Format(LT_CODART, "0") & Format(LK_FECHA_DIA, "ddmm") & Format(Now, "HHMMSS")
        If Val(.det_lot1.TextMatrix(.det_lot1.Rows - 1, 3)) = 0 Then
            .det_lot1.TextMatrix(.det_lot1.Rows - 1, 3) = Val(.grid_fac.TextMatrix(.grid_fac.Row, 4)) / Val(.grid_fac.TextMatrix(.grid_fac.Row, 4))  'LT_CANTIDAD * LT_EQUIV
        Else
            .det_lot1.TextMatrix(.det_lot1.Rows - 1, 3) = "0.00"
        End If
        .det_lot1.TextMatrix(.det_lot1.Rows - 1, 4) = FILA_GRID
        .det_lot1.TextMatrix(.det_lot1.Rows - 1, 5) = "0.00"
        '.det_lot1.TextMatrix(.det_lot1.Rows - 1, 6) = LK_FECHA_DIA
        .det_lot1.TextMatrix(.det_lot1.Rows - 1, 7) = LT_DESCRIP
        .det_lot1.TextMatrix(.det_lot1.Rows - 1, 8) = LT_EQUIV
        .det_lot1.TextMatrix(.det_lot1.Rows - 1, 9) = .det_lot1.TextMatrix(.det_lot1.Rows - 1, 2)
        .det_lot1.TextMatrix(.det_lot1.Rows - 1, 10) = ""
        .det_lot1.TextMatrix(.det_lot1.Rows - 1, 11) = ""
        .det_lot1.TextMatrix(.det_lot1.Rows - 1, 12) = ""
        .det_lot1.TextMatrix(.det_lot1.Rows - 1, 13) = ""
        .det_lot1.TextMatrix(.det_lot1.Rows - 1, 14) = ""
    Else
        .det_lot.Rows = .det_lot.Rows + 1
        .det_lot.TextMatrix(.det_lot.Rows - 1, 0) = LK_CODCIA
        .det_lot.TextMatrix(.det_lot.Rows - 1, 1) = LT_CODART
        .det_lot.TextMatrix(.det_lot.Rows - 1, 2) = "(*)" 'Format(LT_CODART, "0") & Format(LK_FECHA_DIA, "ddmm") & Format(Now, "HHMMSS")
        If Val(.det_lot.TextMatrix(.det_lot.Rows - 1, 3)) = 0 Then
            .det_lot.TextMatrix(.det_lot.Rows - 1, 3) = Val(.grid_fac.TextMatrix(.grid_fac.Row, 4)) / Val(.grid_fac.TextMatrix(.grid_fac.Row, 4))  'LT_CANTIDAD * LT_EQUIV
        Else
            .det_lot.TextMatrix(.det_lot.Rows - 1, 3) = "0.00"
        End If
        .det_lot.TextMatrix(.det_lot.Rows - 1, 4) = FILA_GRID
        .det_lot.TextMatrix(.det_lot.Rows - 1, 5) = "0.00"
        '.det_lot.TextMatrix(.det_lot.Rows - 1, 6) = LK_FECHA_DIA
        .det_lot.TextMatrix(.det_lot.Rows - 1, 7) = LT_DESCRIP
        .det_lot.TextMatrix(.det_lot.Rows - 1, 8) = LT_EQUIV
        .det_lot.TextMatrix(.det_lot.Rows - 1, 9) = .det_lot.TextMatrix(.det_lot.Rows - 1, 2)
        .det_lot.TextMatrix(.det_lot.Rows - 1, 10) = ""
        .det_lot.TextMatrix(.det_lot.Rows - 1, 11) = ""
        .det_lot.TextMatrix(.det_lot.Rows - 1, 12) = ""
        .det_lot.TextMatrix(.det_lot.Rows - 1, 13) = ""
    End If
    'Next J
End With
End Sub

Public Sub CONFIGURAR_DATOS(LT_CODCIA As String, LT_CODART As Currency, LT_CANTIDAD As Currency, LT_DESCRIP As String, LT_EQUIV As Integer, FILA_GRID As Integer, Optional NOMUESTRA)

Dim xcta As Integer
Dim XCTA2 As Integer
Dim JALA_DATO As rdoResultset
Dim LT_SALDO As Currency
Dim flag_datos As String * 1
Dim WCANTIDAD_INI  As Currency
WCANTIDAD_INI = LT_CANTIDAD

If LK_CODTRA = 2401 Then
    pub_cadena = "SELECT * FROM DATOS WHERE DAT_CODCIA = '" & LT_CODCIA & "' AND DAT_CODART =" & LT_CODART & " AND DAT_STOCK > 0 AND DAT_ESTADO='N'  ORDER BY DAT_MOTOR" 'ORDER BY LOT_FECHA_VCTO
Else
    pub_cadena = "SELECT * FROM DATOS WHERE DAT_CODCIA = '" & LT_CODCIA & "' AND DAT_CODART =" & LT_CODART & "  AND DAT_ESTADO='N' AND DAT_STOCK<0 ORDER BY DAT_MOTOR"
End If
Set JALA_DATO = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)   ' rdConcurReadOnly) ', rdConcurLock)
LT_SALDO = LT_CANTIDAD
With FORMGEN
    .gridlt.Clear
    .gridlt.Rows = 2
    .gridlt.Cols = 14
    .gridlt.ColWidth(0) = 2200 'NRO MOTOR
    If Val(.grid_fac.TextMatrix(.grid_fac.Row, 47)) = 2 Or _
        (Val(.grid_fac.TextMatrix(.grid_fac.Row, 47)) = 1 And Val(.grid_fac.TextMatrix(.grid_fac.Row, 48)) <> 2) Then
        .gridlt.ColWidth(5) = 0
    Else
        .gridlt.ColWidth(5) = 800 'CHASIS ----ANTES ERA DESCRIP DE EQUIV
    End If
    .gridlt.ColWidth(2) = 1000 'stock
    .gridlt.ColWidth(3) = 1000 'cantidad solicitada
    .gridlt.ColWidth(4) = 0 '1200  era :'fecha vencimiento
    If Val(.grid_fac.TextMatrix(.grid_fac.Row, 47)) = 2 Or _
        (Val(.grid_fac.TextMatrix(.grid_fac.Row, 47)) = 1 And Val(.grid_fac.TextMatrix(.grid_fac.Row, 48)) <> 2) Then
        .gridlt.ColWidth(1) = 0 'DESCRIP DE EQUIV    ---ANTES ERA chasis
        .gridlt.ColWidth(6) = 0 'procedencia
        .gridlt.ColWidth(7) = 0 'anio
        .gridlt.ColWidth(8) = 0 'poliza
        .gridlt.ColWidth(13) = 1400 'Nro Telefono

    Else
        .gridlt.ColWidth(1) = 1400 'DESCRIP DE EQUIV    ---ANTES ERA chasis
        .gridlt.ColWidth(6) = 1400 'procedencia
        .gridlt.ColWidth(7) = 800 'anio
        .gridlt.ColWidth(8) = 1400 'poliza
        .gridlt.ColWidth(13) = 0 ' Nro Telefono
    End If
    
    .gridlt.ColWidth(9) = 0 ' codigo de arti
    .gridlt.ColWidth(10) = 0 ' FILA DE GRID_FAC
    .gridlt.ColWidth(11) = 0 ' EQUIV
    .gridlt.ColWidth(12) = 0 ' Nro de motor original
     If Val(.grid_fac.TextMatrix(.grid_fac.Row, 47)) = 2 Or _
        (Val(.grid_fac.TextMatrix(.grid_fac.Row, 47)) = 1 And Val(.grid_fac.TextMatrix(.grid_fac.Row, 48)) <> 2) Then
        'Or (Val(grid_fac.TextMatrix(grid_fac.Row, 47)) = 1 And Val(grid_fac.TextMatrix(grid_fac.Row, 48)) > 1) Then
        .gridlt.TextMatrix(0, 0) = "Nº IMEI"
    Else
        .gridlt.TextMatrix(0, 0) = "Nº MOTOR"
    End If
    .gridlt.TextMatrix(0, 1) = "CHASIS"
    .gridlt.TextMatrix(0, 2) = "STOCK"
    .gridlt.TextMatrix(0, 3) = "SOLICITA"
    .gridlt.TextMatrix(0, 5) = "UNIDAD"
    .gridlt.TextMatrix(0, 6) = "PROCED"
    .gridlt.TextMatrix(0, 7) = "AÑO"
    .gridlt.TextMatrix(0, 8) = "POLIZA"
    'gridlt.TextMatrix(0, 4) = "FECHAVCTO"
    .gridlt.TextMatrix(0, 10) = "fila del grid_fac"
    .gridlt.TextMatrix(0, 13) = "Teléfono"
    ' chequeo existe datos en el detalle para mostrar
    '------------------------------------------------
End With
otra_vez:
With FORMGEN
    .gridlt.Rows = 2
    flag_datos = ""
    If Val(.grid_fac.TextMatrix(FILA_GRID, 47)) = 2 Or _
        (Val(.grid_fac.TextMatrix(FILA_GRID, 47)) = 1 And Val(.grid_fac.TextMatrix(FILA_GRID, 48)) <> 2) Then
        
        For xcta = 1 To .det_lot1.Rows - 1
         If Val(.det_lot1.TextMatrix(xcta, 1)) = LT_CODART And FILA_GRID = Val(.det_lot1.TextMatrix(xcta, 4)) Then
           flag_datos = "A"
           .gridlt.Rows = .gridlt.Rows + 1
           .gridlt.TextMatrix(.gridlt.Rows - 1, 0) = .det_lot1.TextMatrix(xcta, 2) ' NRO LOTE
           .gridlt.TextMatrix(.gridlt.Rows - 1, 5) = .det_lot1.TextMatrix(xcta, 7)  ' DESCRIP UNIDAD
           .gridlt.TextMatrix(.gridlt.Rows - 1, 2) = Format(Val(.det_lot1.TextMatrix(xcta, 5)) / Val(.det_lot1.TextMatrix(xcta, 8)), "0.0000")   ' SALDO INCIAL
           .gridlt.TextMatrix(.gridlt.Rows - 1, 3) = Format(Val(.det_lot1.TextMatrix(xcta, 3)) / Val(.det_lot1.TextMatrix(xcta, 8)), "0.0000")  'CANTIDAD PARA MOVER
           '.gridlt.TextMatrix(.gridlt.Rows - 1, 4) = det_lot.TextMatrix(xcta, 6) ' FECHA VCTO
           .gridlt.TextMatrix(.gridlt.Rows - 1, 1) = Trim(.det_lot1.TextMatrix(xcta, 10))  'chasis
           .gridlt.TextMatrix(.gridlt.Rows - 1, 6) = Trim(.det_lot1.TextMatrix(xcta, 11))  'procedencia
           .gridlt.TextMatrix(.gridlt.Rows - 1, 7) = Trim(.det_lot1.TextMatrix(xcta, 12))  'anio
           .gridlt.TextMatrix(.gridlt.Rows - 1, 8) = Trim(.det_lot1.TextMatrix(xcta, 13))  'poliza
           .gridlt.TextMatrix(.gridlt.Rows - 1, 9) = LT_CODART
           .gridlt.TextMatrix(.gridlt.Rows - 1, 10) = FILA_GRID
           .gridlt.TextMatrix(.gridlt.Rows - 1, 11) = .det_lot1.TextMatrix(xcta, 8)  ' equiv de unidad
           .gridlt.TextMatrix(.gridlt.Rows - 1, 12) = .det_lot1.TextMatrix(xcta, 2)  ' ARTICULO
           .gridlt.TextMatrix(.gridlt.Rows - 1, 13) = .det_lot1.TextMatrix(xcta, 14)
         End If
        Next xcta
    Else
        For xcta = 1 To .det_lot.Rows - 1
         If Val(.det_lot.TextMatrix(xcta, 1)) = LT_CODART And FILA_GRID = Val(.det_lot.TextMatrix(xcta, 4)) Then
           flag_datos = "A"
           .gridlt.Rows = .gridlt.Rows + 1
           .gridlt.TextMatrix(.gridlt.Rows - 1, 0) = .det_lot.TextMatrix(xcta, 2) ' NRO LOTE
           .gridlt.TextMatrix(.gridlt.Rows - 1, 5) = .det_lot.TextMatrix(xcta, 7)  ' DESCRIP UNIDAD
           .gridlt.TextMatrix(.gridlt.Rows - 1, 2) = Format(Val(.det_lot.TextMatrix(xcta, 5)) / Val(.det_lot.TextMatrix(xcta, 8)), "0.0000")   ' SALDO INCIAL
           .gridlt.TextMatrix(.gridlt.Rows - 1, 3) = Format(Val(.det_lot.TextMatrix(xcta, 3)) / Val(.det_lot.TextMatrix(xcta, 8)), "0.0000")  'CANTIDAD PARA MOVER
           '.gridlt.TextMatrix(.gridlt.Rows - 1, 4) = det_lot.TextMatrix(xcta, 6) ' FECHA VCTO
           .gridlt.TextMatrix(.gridlt.Rows - 1, 1) = Trim(.det_lot.TextMatrix(xcta, 10))  'chasis
           .gridlt.TextMatrix(.gridlt.Rows - 1, 6) = Trim(.det_lot.TextMatrix(xcta, 11))  'procedencia
           .gridlt.TextMatrix(.gridlt.Rows - 1, 7) = Trim(.det_lot.TextMatrix(xcta, 12))  'anio
           .gridlt.TextMatrix(.gridlt.Rows - 1, 8) = Trim(.det_lot.TextMatrix(xcta, 13))  'poliza
           .gridlt.TextMatrix(.gridlt.Rows - 1, 9) = LT_CODART
           .gridlt.TextMatrix(.gridlt.Rows - 1, 10) = FILA_GRID
           .gridlt.TextMatrix(.gridlt.Rows - 1, 11) = .det_lot.TextMatrix(xcta, 8)  ' equiv de unidad
           .gridlt.TextMatrix(.gridlt.Rows - 1, 12) = .det_lot.TextMatrix(xcta, 2)  ' ARTICULO
         End If
        Next xcta
    End If
New_Dato:
    If flag_datos = "A" Then
     GoTo mues
    Else
     If JALA_DATO.EOF Then
      If LK_CODTRA = 1401 Or LK_CODTRA = 2403 Then
        For J = 1 To Val(.grid_fac.TextMatrix(.grid_fac.Row, 4))
             ASIGNA_NEW_DATO LT_CODART, LT_CANTIDAD, LT_DESCRIP, LT_EQUIV, FILA_GRID ', FORMGEN.grid_fac.TextMatrix(FORMGEN.grid_fac.Row, 4)
        Next J
        GoTo otra_vez
      End If
     End If
    End If
    xcta = 1
    If Not JALA_DATO.EOF Then
      LT_CANTIDAD = LT_CANTIDAD * LT_EQUIV
    End If
    Do Until JALA_DATO.EOF
       If LT_CANTIDAD <= 0 Then
           LT_SALDO = 0
           LT_CANTIDAD = 0
           GoSub muestra
           GoTo OTRO
       End If
       If LT_CANTIDAD <= JALA_DATO!dat_stock Then
           LT_SALDO = LT_CANTIDAD
           LT_CANTIDAD = 0
           GoSub muestra
       Else
           If JALA_DATO.RowCount = JALA_DATO.AbsolutePosition Then
              LT_SALDO = LT_CANTIDAD  'JALA_DATO!LOT_SALDOS
              LT_CANTIDAD = 0
           Else
              LT_SALDO = JALA_DATO!dat_stock
              'LT_CANTIDAD = 0
           End If
           GoSub muestra
           LT_CANTIDAD = LT_CANTIDAD - JALA_DATO!dat_stock
       End If
OTRO:
      JALA_DATO.MoveNext
    Loop
End With
mues:
With FORMGEN
    PUB_KEY = LT_CODART
    pu_codcia = LT_CODCIA
    SQ_OPER = 1
    LEER_ART_LLAVE
    If Not art_LLAVE.EOF Then
      .lblltcod.Caption = art_LLAVE!art_alterno
      .lblltnom.Caption = Trim(art_LLAVE!ART_NOMBRE)
    End If
    .lblltcantidad.Caption = Format(WCANTIDAD_INI, "0.000")
    .lblltunidad.Caption = Trim(LT_DESCRIP) & String(80, "  ") & LT_EQUIV
    If Not IsMissing(NOMUESTRA) Then
     If NOMUESTRA = 1 Then Exit Sub
    End If
    
    .fralotes.Visible = True
    If .gridlt.Rows > 2 Then
     If Val(.grid_fac.TextMatrix(.grid_fac.Row, 47)) = 2 Or _
        (Val(.grid_fac.TextMatrix(.grid_fac.Row, 47)) = 1 And _
        Val(.grid_fac.TextMatrix(.grid_fac.Row, 48)) <> 2) Then
        
        .fralotes.Width = 5820
        .gridlt.Width = 5655
     Else
        .fralotes.Width = 10380
        .gridlt.Width = 10215
     End If
     .gridlt.Row = 2
     .gridlt.COL = 0
    End If
    .gridlt.SetFocus
    Exit Sub
End With
muestra:
With FORMGEN
     xcta = xcta + 1
     .gridlt.Rows = .gridlt.Rows + 1
     If Val(LT_SALDO) <> 0 Then
     .gridlt.TextMatrix(xcta, 0) = "X"
     Else
     .gridlt.TextMatrix(xcta, 0) = ""
     End If
     .gridlt.TextMatrix(xcta, 0) = Trim(JALA_DATO!DAT_MOTOR)
     .gridlt.TextMatrix(xcta, 5) = Trim(LT_DESCRIP)
     .gridlt.TextMatrix(xcta, 2) = Format(Val(JALA_DATO!dat_stock) / Val(Right(LT_EQUIV, 8)), "0.0000")
     If Val(JALA_DATO!dat_stock) = 0 Then
       '.gridlt.TextMatrix(xcta, 3) = Format(Val(LT_SALDO) / Val(Right(LT_EQUIV, 8)), "0.0000")
       If LT_CANTIDAD < 2 Then
        .gridlt.TextMatrix(xcta, 3) = Format(Val(LT_CANTIDAD), "0.0000")
       Else
        .gridlt.TextMatrix(xcta, 3) = Format(Val(LT_CANTIDAD) - 1, "0.0000")
       End If
     Else
        .gridlt.TextMatrix(xcta, 3) = "0.0000"
     End If
     'gridlt.TextMatrix(xcta, 4) = Format(JALA_DATO!lot_fecha_vcto, "dd/mm/yy")
     .gridlt.TextMatrix(xcta, 1) = Trim(Nulo_Valors(JALA_DATO!DAT_CHASIS))
     .gridlt.TextMatrix(xcta, 6) = Trim(Nulo_Valors(JALA_DATO!DAT_PROCED))
     .gridlt.TextMatrix(xcta, 7) = Trim(Nulo_Valors(JALA_DATO!DAT_ANIO))
     .gridlt.TextMatrix(xcta, 8) = Trim(Nulo_Valors(JALA_DATO!DAT_POLIZA))
     .gridlt.TextMatrix(xcta, 9) = JALA_DATO!DAT_CODART
     .gridlt.TextMatrix(xcta, 10) = FILA_GRID
     .gridlt.TextMatrix(xcta, 11) = LT_EQUIV
     .gridlt.TextMatrix(xcta, 12) = Trim(JALA_DATO!DAT_MOTOR)
     .gridlt.TextMatrix(xcta, 13) = Trim(JALA_DATO!DAT_CHASIS)
End With
Return
End Sub

Public Sub BORRA_ITEM_DATO(LT_CODART As Currency, FILA_GRID As Integer, Optional WNROLOTE)
Dim ix As Integer
With FORMGEN
    If Not IsMissing(WNROLOTE) Then
     For ix = 1 To .det_lot.Rows - 1
       If WNROLOTE = .det_lot.TextMatrix(ix, 9) And LT_CODART = Val(.det_lot.TextMatrix(ix, 1)) And FILA_GRID = Val(.det_lot.TextMatrix(ix, 4)) Then
         If .det_lot.Rows <= 2 Then
           .det_lot.Rows = 1
           Exit For
         Else
           .det_lot.RemoveItem (ix)
           ix = ix - 1
           Exit For
         End If
       End If
     Next ix
     Exit Sub
    End If
    
    
    For ix = 1 To .det_lot.Rows - 1
      If LT_CODART = Val(.det_lot.TextMatrix(ix, 1)) And FILA_GRID = Val(.det_lot.TextMatrix(ix, 4)) Then
        If .det_lot.Rows <= 2 Then
          .det_lot.Rows = 1
          Exit For
        Else
          .det_lot.RemoveItem (ix)
          ix = ix - 1
        End If
      End If
    Next ix
End With
End Sub
Public Function Edit_Datos() As Boolean
 If LK_CODTRA = 1111 Then
    PSLOT_LLAVE(0) = LK_CODCIA
    PSLOT_LLAVE(1) = arm_llave!ARM_CODART
    PUB_MOTOR = Trim(FORMGEN.grid_fac.TextMatrix(fila, 44))
    PUB_DATO_CANT = Val(FORMGEN.grid_fac.TextMatrix(fila, 45))
    PUB_CHASIS = Trim(FORMGEN.grid_fac.TextMatrix(fila, 49))
    PSLOT_LLAVE(2) = Trim(FORMGEN.grid_fac.TextMatrix(fila, 44)) 'PUB_MOTOR
    lot_llave.Requery
    If lot_llave.EOF Then
        CN.Execute "Rollback Transaction", rdExecDirect
        If Val(FORMGEN.grid_fac.TextMatrix(fila, 47)) = 2 Then
            MsgBox "Nro de IMEI no Encontrado  : " & Trim(FORMGEN.grid_fac.TextMatrix(fila, 44)), 48, Pub_Titulo
        Else
            MsgBox "Nro de Motor no Encontrado  : " & Trim(FORMGEN.grid_fac.TextMatrix(fila, 44)), 48, Pub_Titulo
        End If
        Edit_Datos = True
        Exit Function
    End If
   ' If PUB_TIPMOV = 20 Then
   '     If lot_llave!dat_stock = 0 Then
   '         CN.Execute "Rollback Transaction", rdExecDirect
   '         MsgBox "Proceso no procede...Hubo una Venta antes.Extorne primero la Venta.", 48, Pub_Titulo
   '         Edit_Datos = True
   '         Exit Function
   '     End If
   ' End If
    lot_llave.Edit
    lot_llave!DAT_ESTADO = "E"
    lot_llave!dat_stock = lot_llave!dat_stock + (Val(FORMGEN.grid_fac.TextMatrix(fila, 45)) * pub_signo_arm)
    PUB_DATO_CANT = lot_llave!dat_stock
    lot_llave.Update
    If PUB_TIPMOV <> 20 Then
        'DECLARO PARA QUE NO QUEDE EN LAS OTRAS VARIABLES ---JC
        Dim WNUMSER As Long
        Dim wnumfac As Long
        
        PUB_PROCED = Trim(Nulo_Valors(lot_llave!DAT_PROCED))
        PUB_ANIO = Trim(Nulo_Valors(lot_llave!DAT_ANIO))
        PUB_POLIZA = Trim(Nulo_Valors(lot_llave!DAT_POLIZA))
        WNUMSER = Val(lot_llave!DAT_NUMSER)
        wnumfac = Val(lot_llave!DAT_NUMFAC)
        lot_llave.AddNew
        lot_llave!DAT_CODCIA = LK_CODCIA
        lot_llave!DAT_MOTOR = PUB_MOTOR
        lot_llave!DAT_CODART = arm_llave!ARM_CODART
        lot_llave!DAT_CHASIS = PUB_CHASIS
        lot_llave!DAT_PROCED = PUB_PROCED
        lot_llave!DAT_ANIO = PUB_ANIO
        lot_llave!DAT_POLIZA = PUB_POLIZA
        lot_llave!dat_stock = PUB_DATO_CANT
        lot_llave!DAT_NUMSER = WNUMSER
        lot_llave!DAT_NUMFAC = wnumfac
        lot_llave.Update
    End If
    Edit_Datos = False
End If
End Function
Public Sub New_Edit_datos(F As Integer, Optional ByVal condi As Integer = 0)
If LK_CODTRA = 1111 Then
  ParaLot_count = 0
  GoTo SINLOTE
End If
If condi = 1 Then
    GoTo Otro_lot
End If
If Val(FORMGEN.grid_fac.TextMatrix(F, 47)) = 1 And Val(FORMGEN.grid_fac.TextMatrix(F, 48)) = 2 Then
    verif_datos Val(arm_llave!ARM_CODART), F
Else
    verif_datos1 Val(arm_llave!ARM_CODART), F
End If

FORMGEN.asigna_flag_lote = ""
If ParaLot_count = 1 Then FORMGEN.asigna_flag_lote = "A"
Otro_lot:
    If Val(FORMGEN.grid_fac.TextMatrix(F, 47)) = 2 Or _
        (Val(FORMGEN.grid_fac.TextMatrix(F, 47)) = 1 And Val(FORMGEN.grid_fac.TextMatrix(F, 48)) <> 2) Then
        
      PUB_DATO_CANT = ParaLot_lotcant(ParaLot_count)
      PUB_MOTOR = ParaLot_codlot(ParaLot_count)
      PUB_CHASIS = ParaLot_lotchasis(ParaLot_count)
      PUB_PROCED = ""
      PUB_ANIO = ""
      PUB_POLIZA = ""
    Else
      PUB_DATO_CANT = ParaLot_lotcant(ParaLot_count)
      PUB_MOTOR = ParaLot_codlot(ParaLot_count)
      PUB_CHASIS = ParaLot_lotchasis(ParaLot_count)
      PUB_PROCED = ParaLot_lotproced(ParaLot_count)
      PUB_ANIO = ParaLot_lotanio(ParaLot_count)
      PUB_POLIZA = ParaLot_lotpoliza(ParaLot_count)
    End If
      'PUB_FECHA_LOT = ParaLot_fechalot(ParaLot_count)
      ParaLot_count = ParaLot_count - 1
SINLOTE:
      far_llave.AddNew
      
       'Datos
      PSLOT_LLAVE(0) = LK_CODCIA
      PSLOT_LLAVE(1) = arm_llave!ARM_CODART
      PSLOT_LLAVE(2) = PUB_MOTOR
      lot_llave.Requery
      If lot_llave.EOF Then
        If LK_CODTRA = 1111 And PUB_TIPMOV = 20 Then
            PUB_DATO_CANT = 1
        Else
            lot_llave.AddNew
            lot_llave!DAT_CODCIA = LK_CODCIA
            lot_llave!DAT_MOTOR = PUB_MOTOR
            lot_llave!DAT_CODART = arm_llave!ARM_CODART
            lot_llave!DAT_CHASIS = PUB_CHASIS
            lot_llave!DAT_PROCED = PUB_PROCED
            lot_llave!DAT_ANIO = PUB_ANIO
            lot_llave!DAT_POLIZA = PUB_POLIZA
            lot_llave!DAT_NUMSER = FORMGEN.i_numser.Text
            lot_llave!DAT_NUMFAC = Val(FORMGEN.i_numfac.Text)
            
            'lot_llave!lot_fecha_vcto = PUB_FECHA_LOT
            lot_llave!dat_stock = 0
        End If
      Else
        lot_llave.Edit
      End If
      'lot_llave!LOT_CODCLIE = 0
      If LK_CODTRA <> 1111 Then
        lot_llave!dat_stock = lot_llave!dat_stock + (PUB_DATO_CANT * pub_signo_arm)
        lot_llave.Update
      End If
      far_llave!far_cantidad_p = PUB_DATO_CANT
      far_llave!far_Motor = Trim(PUB_MOTOR)
      far_llave!far_Chasis = Trim(PUB_CHASIS)
      If FORMGEN.asigna_flag_lote = "A" Then
        far_llave!FAR_estado2 = "N"
      Else
       If ParaLot_count = 1 Then
          far_llave!FAR_estado2 = "N"
       Else
          far_llave!FAR_estado2 = "L"
       End If
      End If
End Sub
Public Function VerificaMotor() As Boolean
    Dim vpuedo As Boolean
    Dim VERIFI_MOTOR As rdoResultset
    Dim J As Integer

    pub_cadena = "SELECT DAT_MOTOR FROM DATOS WHERE DAT_ESTADO<>'E' ORDER BY DAT_NUMSER,DAT_NUMFAC"
    Set VERIFI_MOTOR = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)
                    
    If VERIFI_MOTOR.EOF Then
       VerificaMotor = True
       Exit Function
    End If
    vpuedo = True
With FORMGEN
    For J = 2 To .gridlt.Rows - 1
        VERIFI_MOTOR.MoveFirst
        Do While Not VERIFI_MOTOR.EOF
            If Trim(.gridlt.TextMatrix(J, 0)) = Trim(VERIFI_MOTOR!DAT_MOTOR) Then
                vpuedo = False
                Exit Do
            Else
                vpuedo = True
            End If
            VERIFI_MOTOR.MoveNext
        Loop
        If vpuedo = False Then
            Exit For
        End If
    Next J
    If vpuedo = False Then
        If Val(.grid_fac.TextMatrix(.grid_fac.Row, 47)) = 2 Then
             MsgBox "Proceso no Procede..." & vbCrLf & _
                 "IMEI: " + Trim(.gridlt.TextMatrix(J, 0)) + " existente.", vbExclamation, Pub_Titulo
        Else
            MsgBox "Proceso no Procede..." & vbCrLf & _
                 "Motor: " + Trim(.gridlt.TextMatrix(J, 0)) + " existente.", vbExclamation, Pub_Titulo
        End If
       .gridlt.Row = J
       .gridlt.COL = 0
       .gridlt.SetFocus
        VerificaMotor = False
    Else
       VerificaMotor = True
    End If
End With
End Function
Public Function VerificaFono() As Boolean
    Dim vpuedo As Boolean
    Dim VERIFI_FONO As rdoResultset
    Dim J As Integer

    pub_cadena = "SELECT DAT_CHASIS FROM DATOS WHERE DAT_ESTADO<>'E' ORDER BY DAT_NUMSER,DAT_NUMFAC"
    Set VERIFI_FONO = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)
                    
    If VERIFI_FONO.EOF Then
       VerificaFono = True
       Exit Function
    End If
    vpuedo = True
With FORMGEN
    For J = 2 To .gridlt.Rows - 1
        VERIFI_FONO.MoveFirst
        Do While Not VERIFI_FONO.EOF
            If Trim(.gridlt.TextMatrix(J, 0)) = Trim(Nulo_Valors(VERIFI_FONO!DAT_CHASIS)) Then
                vpuedo = False
                Exit Do
            Else
                vpuedo = True
            End If
            VERIFI_FONO.MoveNext
        Loop
        If vpuedo = False Then
            Exit For
        End If
    Next J
    If vpuedo = False Then
        MsgBox "Proceso no Procede..." & vbCrLf & _
                "Teléfono: " + Trim(.gridlt.TextMatrix(J, 0)) + " existente.", vbExclamation, Pub_Titulo
       .gridlt.Row = J
       .gridlt.COL = 13
       .gridlt.SetFocus
        VerificaFono = False
    Else
       VerificaFono = True
    End If
End With
End Function
