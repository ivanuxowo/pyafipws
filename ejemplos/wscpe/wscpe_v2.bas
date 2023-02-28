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
    WSCPE.CUIT = "20267565393"
    
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
    cuit_solicitante = "20111111112"
    sucursal = 1
    nro_orden = 1
    planta = 1
    carta_porte = 1
    nro_ctg = 10100000542#
    observaciones = "Notas del transporte"
    ok = WSCPE.AgregarCabecera(tipo_cpe, cuit_solicitante, sucursal, nro_orden, planta, carta_porte, nro_ctg, observaciones)
    
    es_usuario_industrial = True
    cuit_titular_planta = "20200000006"
    domicilio_origen_tipo = 2
    domicilio_origen_orden = 1
    planta = 1
    ok = WSCPE.AgregarOrigen(es_usuario_industrial, cuit_titular_planta, domicilio_origen_tipo, domicilio_origen_orden, planta)
    
    cuit_destino = "20111111112"
    domicilio_destino_tipo = 1
    domicilio_destino_orden = 2
    planta = 1938
    cuit_destinatario = CUIT
    ok = WSCPE.AgregarDestino(cuit_destino, domicilio_destino_tipo, domicilio_destino_orden, planta, cuit_destinatario)
    
    corresponde_retiro_productor = True
    es_solicitante_campo = True
    certificado_coe = "330100025869"
    cuit_remitente_comercial_productor = "20111111112"
    ok = WSCPE.AgregarRetiroProductor(corresponde_retiro_productor, es_solicitante_campo, certificado_coe, cuit_remitente_comercial_productor)
        
    cuit_reminitente_comercial = "20111111112"
    cuit_mercado_a_termino = "20222222223"
    cuit_comisionista = "20222222223"
    cuit_corredor = "20400000000"
    ok = WSCPE.AgregarIntervinientes(cuit_reminitente_comercial, cuit_mercado_a_termino, cuit_comisionista, cuit_corredor)
    
    cod_grano = 23
    cod_derivado_granario = 136
    peso_bruto = 110
    peso_tara = 10
    tipo_embalaje = 1
    otro_embalaje = vbNull
    unidad_media = 1
    cantidad_unidades = vbNull
    kg_litro_m3 = vbNull
    lote = vbNull
    fecha_lote = vbNull
    ok = WSCPE.AgregarDatosCarga(cod_grano, cod_derivado_granario, peso_bruto, peso_tara, tipo_embalaje, otro_embalaje, unidad_media, cantidad_unidades, kg_litro_m3, lote, fecha_lote)
    
    cuit_transportista = "20333333334"
    cuit_transportista_tramo2 = "20222222223"
    nro_vagon = 55555556
    nro_precinto = 1
    nro_operativo = "1111111111"
    dominio = "AB001ST"
    fecha_hora_partida = "2021-08-21T23:29:26.579557"
    km_recorrer = 500
    codigo_turno = "00"
    cuit_chofer = "20333333334"
    tarifa = 100.1
    cuit_pagador_flete = "20333333334"
    mercaderia_fumigada = True
    cuit_intermediario_flete = "20333333334"
    codigo_ramal = False
    descripcion_ramal = "XXXXX"
    tarifa_referencia = vbNull
    ok = WSCPE.AgregarTransporte(cuit_transportista, cuit_transportista_tramo2, nro_vagon, nro_precinto, nro_operativo, dominio, fecha_hora_partida, km_recorrer, codigo_turno, cuit_chofer, tarifa, cuit_pagador_flete, mercaderia_fumigada, cuit_intermediario_flete, codigo_ramal, descripcion_ramal, tarifa_referencia)
    
    ' ok = WSCPE.LoadTestXML(App.Path & "\autorizar.xml")
    
    archivo = App.Path & "\cpe.pdf"
    ok = WSCPE.AutorizarCPEAutomotorDG(archivo)
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
