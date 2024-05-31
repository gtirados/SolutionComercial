Attribute VB_Name = "ParaFormGen"
Dim PUB_CADENA1 As String
'***** JC
Public PSFAR_MOV25 As rdoQuery
Public FAR_MOV25 As rdoResultset
'****JC
Public Sub Mov25()
    PUB_CADENA1 = "SELECT isnull(MAX(FAR_NUMSEC),0) FROM FACART " & _
    "WHERE  FAR_CODCIA =? and FAR_NUMSER=? AND FAR_TIPMOV=25 AND " & _
    "FAR_NUMFAC=? AND FAR_CP = 'C'"
    
    Set PSFAR_MOV25 = CN.CreateQuery("", PUB_CADENA1)
    PSFAR_MOV25.rdoParameters(0) = PUB_CODCIA
    PSFAR_MOV25.rdoParameters(1) = Val(FORMGEN.i_numser_c.Text)
    'PSFAR_MOV25.rdoParameters(2) = PUB_TIPMOV
    PSFAR_MOV25.rdoParameters(2) = Val(FORMGEN.i_numfac_c.Text)
    Set FAR_MOV25 = PSFAR_MOV25.OpenResultset(rdOpenKeyset, rdConcurValues)
End Sub

Public Sub Graba_Caracu()
    With FORMGEN
        caa_histo.AddNew
        caa_histo!CAA_CODCLIE = PUB_CODCLIE
        caa_histo!caa_codcia = LK_CODCIA
        caa_histo!CAA_TIPDOC = PUB_TIPDOC
        caa_histo!CAA_CP = PUB_CP
        caa_histo!CAA_NUM_OPER = PUB_NUM_OPER_XXX
        If .grid_liq.Visible = True Then
           caa_histo!caa_INTVEN = .grid_liq.TextMatrix(1, 2)
           caa_histo!caa_DIASV = .grid_liq.TextMatrix(1, 6)
           caa_histo!caa_DIASA = .grid_liq.TextMatrix(1, 7)
           caa_histo!caa_tasav = Val(.i_tasav.Text)
        Else
           caa_histo!caa_INTVEN = 0
           caa_histo!caa_DIASV = 0
           caa_histo!caa_DIASA = 0
           caa_histo!caa_tasav = 0
        End If
        caa_histo!caa_TIPO_CAMBIO = LK_TIPO_CAMBIO
        caa_histo!caa_serdoc = PUB_SERDOC
        caa_histo!CAA_NUMDOC = PUB_NUMDOC
        caa_histo!CAA_FECHA = LK_FECHA_DIA
        caa_histo!CAA_FECHA_VCTO = PUB_FECHA_VCTO
        caa_histo!caa_situacion = PUB_SITUACION_ACT
        caa_histo!caa_concepto = PUB_CONCEPTO
        caa_histo!CAA_IMPORTE = .WS_IMPORTE_AMORT * pub_signo_car
        
        caa_histo!CAA_TOTAL = Abs(.WS_TOT) * pub_signo_car
        caa_histo!CAA_SALDO = Nulo_Valor0(cli_llave!cli_SALDO)
        caa_histo!caa_SALDO_car = .wS_saldo_car
        If pub_signo_car <> 0 Then caa_histo!CAA_SIGNO_CAJA = PUB_SECUENCIA
        'Else
        If pub_signo_car = 0 Then caa_histo!CAA_SIGNO_CAJA = 0
        'End If
        caa_histo!CAA_SIGNO_CAJA_REAL = pub_signo_caja
        caa_histo!CAA_SIGNO_CAR = pub_signo_car
        caa_histo!CAA_TIPMOV = PUB_TIPMOV
        caa_histo!CAA_hora = Now
        caa_histo!CAA_CODUSU = LK_CODUSU
        caa_histo!CAA_ESTADO = "N"
        If Not cli_llave.EOF And cli_llave.RowCount > 0 Then
           If cli_llave!cli_codclie = Val(.i_codcli.Text) Then caa_histo!CAa_NOMBRE = cli_llave!CLI_NOMBRE
           'End If
        End If
        caa_histo!CAA_NUMPLAN = pub_numplan
        If LK_CODTRA = 1111 Then caa_histo!CAA_ESTADO = "E"
        If LK_CODTRA = 1122 Then caa_histo!CAA_ESTADO = "E"
        If LK_CODTRA = 1133 Then caa_histo!CAA_ESTADO = "E"
        If fx = 32000 Then caa_histo!CAA_ESTADO = "E"
        If SUT_LLAVE!SUT_SIGNO_CAR = 2 And LK_CODTRA <> 2727 Then caa_histo!CAA_ESTADO = "E"
        If .i_fecha_compra.Visible = True Then caa_histo!CAA_FECHA_COBRO = CDate(.i_fecha_compra.Text)
        'Else
        If .i_fecha_compra.Visible = False Then caa_histo!CAA_FECHA_COBRO = LK_FECHA_DIA
        'End If
        If ws_flag_car = ingre Then
           caa_histo!CAa_NUM_CHEQUE = PUB_NUM_CHEQUE
           caa_histo!CAa_numser = PUB_NUMSER
           caa_histo!CAa_numfac = PUB_NUMFAC
           caa_histo!caa_numser_c = PUB_NUMSER_C
           caa_histo!caa_numfac_c = PUB_NUMFAC_C
           If LK_CODTRA = 2741 Then
            caa_histo!CAa_numser = PUB_NUMSER_C
            caa_histo!CAa_numfac = PUB_NUMFAC_C
           End If
           caa_histo!CAa_numGUIA = PUB_NUMGUIA
           caa_histo!CAa_SERGUIA = Val(.i_serguia.Text)
           caa_histo!CAA_FBG = PUB_FBG
           caa_histo!CAA_CODVEN = PUB_CODVEN
           caa_histo!caa_situacion = " "
           caa_histo!caa_FLAG_SO = PUB_SO
        Else
           caa_histo!CAa_NUM_CHEQUE = Nulo_Valors(car_llave!car_NUM_CHEQUE)
           caa_histo!CAa_numser = Nulo_Valor0(car_llave!car_NUMSER)
           caa_histo!CAa_numfac = Nulo_Valor0(car_llave!car_NUMFAC)
           If LK_CODTRA = 2412 Then
             caa_histo!caa_numser_c = PUB_NUMSER
             caa_histo!caa_numfac_c = PUB_NUMFAC
           Else
             caa_histo!caa_numser_c = PUB_NUMSER_C
             caa_histo!caa_numfac_c = PUB_NUMFAC_C
           End If
           If LK_CODTRA = 1122 Or LK_CODTRA = 1111 Then
              If .ww_codtra_ext = 2412 Or .ww_codtra_ext = 2410 Then
                 caa_histo!caa_numfac_c = .ww_numdoc
                 caa_histo!caa_numser_c = .ww_numser
              End If
           End If
           caa_histo!CAA_NOTA = PUB_FBG
           caa_histo!CAA_FBG = Nulo_Valors(car_llave!car_FBG)
           If LK_CODTRA = 2770 And i_codven.Visible = True Then caa_histo!CAA_CODVEN = PUB_CODVEN
           'Else
           If LK_CODTRA <> 2770 And (.i_codven.Visible = False Or .i_codven.Visible = True) Then caa_histo!CAA_CODVEN = Nulo_Valor0(car_llave!CAR_codven)
           'End If
           If PUB_SECUENCIA = 1 And LK_EMP = "HER" Then caa_histo!CAA_CODVEN = Nulo_Valor0(car_llave!CAR_codven)
              caa_histo!CAa_numGUIA = Nulo_Valor0(car_llave!car_numguia)
           caa_histo!CAa_SERGUIA = Nulo_Valor0(car_llave!car_SERguia)
           caa_histo!caa_situacion = Nulo_Valors(car_llave!CAR_SITUACION)
           caa_histo!caa_FLAG_SO = Nulo_Valors(car_llave!CAR_FLAG_SO)
           caa_histo!caa_signo_ccm = pub_signo_ccm
           caa_histo!caa_codban = PUB_CODBAN
        End If
        caa_histo!caa_codTRA = LK_CODTRA
        If WS_ESTADO = 2 And (LK_CODTRA = 2725 Or LK_CODTRA = 2770 Or LK_CODTRA = 2774) Then
           caa_histo!CAA_RECIBO = Val(.i_numfac.Text) 'ES EL MISMO CAMPO
           caa_histo!CAA_SERIE = Val(.i_numser.Text)
        Else
           caa_histo!CAA_RECIBO = 0
           caa_histo!CAA_SERIE = 0
        End If
        If LK_CODTRA = 2728 And LK_EMP = "PLA" Then caa_histo!caa_situacion = .PLAZA_FLAG_MANUAL
        'End If
        caa_histo.Update
    End With
End Sub
