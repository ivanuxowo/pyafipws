Attribute VB_Name = "Module1"
' Ejemplo de Uso de Interface COM con Web Service obtener Carta de Porte Electrónica
' para transporte ferroviario y automotor RG 5017/2021
' Más info en: http://www.sistemasagiles.com.ar/trac/wiki/CartadePorte
' 2021 (C) Mariano Reingart <reingart@gmail.com>

Sub Main()
    Dim WSAA As Object, WSCPE As Object
    On Error GoTo ManejoError
    ttl = 2400 ' tiempo de vida en segundos
    cache = "" ' Directorio para archivos temporales (dejar en blanco para usar predeterminado)
    proxy = "" ' usar "usuario:clave@servidor:puerto"

    Certificado = App.Path & "\reingart.crt"   ' certificado es el firmado por la afip
    ClavePrivada = App.Path & "\reingart.key"  ' clave privada usada para crear el cert.
    
    Set WSAA = CreateObject("WSAA")
    tra = WSAA.CreateTRA("wscpe", ttl)
    Debug.Print tra
    ' Generar el mensaje firmado (CMS)
    cms = WSAA.SignTRA(tra, Certificado, ClavePrivada)
    Debug.Print cms
    
    wsdl = "https://wsaahomo.afip.gov.ar/ws/services/LoginCms?wsdl" ' homologación
    ok = WSAA.Conectar(cache, wsdl, proxy)
    '' ta = WSAA.LoginCMS(cms) 'obtener ticket de acceso
    ta = ""
    Debug.Print ta
    Debug.Print "Token:", WSAA.Token
    Debug.Print "Sign:", WSAA.Sign
    
    ' Crear objeto interface Web Service de CTG
    Set WSCPE = CreateObject("WSCPEv2")
    ' Setear tocken y sing de autorización (pasos previos)
    WSCPE.Token = WSAA.Token
    WSCPE.Sign = WSAA.Sign
    
    ' CUIT (debe estar registrado en la AFIP)
    WSCPE.Cuit = "20267565393"
    
    ' Conectar al Servicio Web
    wrapper = ""
    ok = WSCPE.Conectar("", "https://fwshomo.afip.gov.ar/wscpe/services/soap?wsdl", proxy, wrapper) ' homologación
    ' produccion: https://serviciosjava.afip.gob.ar/cpe-ws/services/wscpe?wsdl
    
    ' ok = WSCPE.Dummy
    Debug.Print "AppServerStatus", WSCPE.AppServerStatus
    Debug.Print "DbServerStatus", WSCPE.DbServerStatus
    Debug.Print "AuthServerStatus", WSCPE.AuthServerStatus
    
    ok = WSCPE.CrearCPE()
    tipo_cpe = 74
    cuit_solicitante = 20111111112#
    sucursal = 1
    nro_orden = 1
    ok = WSCPE.AgregarCabecera(tipo_cpe, cuit_solicitante, sucursal, nro_orden)
    
    planta = 1
    cod_provincia_operador = 12
    cod_localidad_operador = 5544
    cod_provincia_productor = 12
    cod_localidad_productor = 5544
    ok = WSCPE.AgregarOrigen(planta, cod_provincia_operador, cod_localidad_operador, cod_provincia_productor, cod_localidad_productor)
    
    planta = 1
    cod_provincia = 12
    es_destino_campo = True
    cod_localidad = 3058
    cuit_destino = 20111111112#
    cuit_destinatario = 30000000006#
    ok = WSCPE.AgregarDestino(planta, cod_provincia, es_destino_campo, cod_localidad, cuit_destino, cuit_destinatario)
    
    certificado_coe = 330100025869#
    cuit_remitente_comercial_productor = 20111111112#
    corresponde_retiro_productor = True
    es_solicitante_campo = True
    ok = WSCPE.AgregarRetiroProductor(certificado_coe, cuit_remitente_comercial_productor, corresponde_retiro_productor, es_solicitante_campo)
        
    cuit_mercado_a_termino = 20222222223#
    cuit_corredor_venta_primaria = 20222222223#
    cuit_corredor_venta_secundaria = 20222222223#
    cuit_remitente_comercial_venta_secundaria = 20222222223#
    cuit_intermediario = 20222222223#
    cuit_remitente_comercial_venta_primaria = 20222222223#
    cuit_representante_entregador = 20222222223#
    cuit_representante_recibidor = 20222222223#
    ok = WSCPE.AgregarIntervinientes(cuit_mercado_a_termino, cuit_corredor_venta_primaria, cuit_corredor_venta_secundaria, cuit_remitente_comercial_venta_secundaria, cuit_intermediario, cuit_remitente_comercial_venta_primaria, cuit_representante_entregador, cuit_representante_recibidor)
    
    peso_tara = 1000
    cod_grano = 31
    peso_bruto = 1000
    cosecha = 910
    ok = WSCPE.AgregarDatosCarga(peso_tara, cod_grano, peso_bruto, cosecha)
    
    cuit_transportista = 20333333334#
    fecha_hora_partida = "2021-08-21T23:29:26.579557"
    codigo_turno = "00"
    dominio = "ZZZ000"
    km_recorrer = 500
    cuit_chofer = 20333333334#
    tarifa = 100.1
    cuit_pagador_flete = 20333333334#
    cuit_intermediario_flete = 20333333334#
    mercaderia_fumigada = True
    ok = WSCPE.AgregarTransporte(cuit_transportista, fecha_hora_partida, codigo_turno, dominio, km_recorrer, cuit_chofer, tarifa, cuit_pagador_flete, cuit_intermediario_flete, mercaderia_fumigada)
    
    ok = WSCPE.LoadTestXML(App.Path & "\autorizar.xml")
    
    archivo = App.Path & "\cpe.pdf"
    ok = WSCPE.AutorizarCPEAutomotor(archivo)
    Debug.Print "Numero de CTG:", WSCPE.NroCTG
    Debug.Print "Fecha de emision:", WSCPE.FechaEmision
    Debug.Print "Estado:", WSCPE.Estado
    Debug.Print "Fecha de inicio de estado:", WSCPE.FechaInicioEstado
    Debug.Print "Fecha de vencimiento:", WSCPE.FechaVencimiento
            
    Debug.Print WSCPE.XmlResponse
    Debug.Print WSCPE.ErrMsg

    If Not ok Then
        ' muestro los errores
        Dim MensajeError As Variant
        For Each MensajeError In WSCPE.Errores
            MsgBox MensajeError, vbCritical, "WSCPE: Errores"
        Next
    End If
       
    MsgBox "CTG: " & WSCPE.NroCTG, vbInformation, "AutorizarCTE:"
    
    ' Consulto los CTG generados (genera planilla Excel por AFIP)
    
    Dim nro_ctg As Variant
    nro_ctg = 10100000542#
    
    ok = WSCPE.LoadTestXML(App.Path & "\consultar.xml")
    If nro_ctg <> 0 Then
        ok = WSCPE.ConsultarCPEAutomotor(Null, Null, Null, Null, nro_ctg)
    Else
        ok = WSCPE.ConsultarCPEAutomotor(tipo_cpe, sucursal, nro_orden, cuit_solicitante)
    End If
    ' Obtengo la constacia CTG -debe estar confirmada- (documento PDF AFIP)
    
    Debug.Print WSCPE.XmlResponse
    Debug.Print "Numero de CTG:", WSCPE.NroCTG
    Debug.Print "Errores:", WSCPE.ErrMsg
        
        

    

Exit Sub
ManejoError:
    ' Si hubo error:
    Debug.Print Err.Description            ' descripción error afip
    Debug.Print Err.Number - vbObjectError ' codigo error afip
    Select Case MsgBox(Err.Description, vbCritical + vbRetryCancel, "Error:" & Err.Number - vbObjectError & " en " & Err.Source)
        Case vbRetry
            Debug.Assert False
            Resume
        Case vbCancel
            Debug.Print Err.Description
    End Select
    Debug.Print WSCPE.XmlRequest
    Debug.Assert False

End Sub
