Public Class BdgInterface
    Implements IBodega

    'Public Sub ReconstruirDAALineas(ByVal data As DataReconstruirDAALineas, ByVal services As ServiceProvider) Implements IBodega.ReconstruirDAALineas
    '    If Not data.dtCabeceraAlbaran Is Nothing AndAlso data.dtCabeceraAlbaran.Rows.Count > 0 Then

    '        Dim dtCabDAA As DataTable = New BdgDAACabecera().SelOnPrimaryKey(data.IDDaa)

    '        Dim dataResult As New BdgDAACabecera.stCrearDAAInfo(BdgDAACabecera.enumTipoEnvioDAA.EMCSInterno)
    '        dataResult.Cabecera = dtCabDAA
    '        'dataResult.Contador = data.IDContador
    '        dataResult.IDDocumento = data.dtCabeceraAlbaran.Rows(0)("IDAlbaran")
    '        dataResult.FechaDocumento = data.dtCabeceraAlbaran.Rows(0)("FechaAlbaran")
    '        'dataResult.NumeroCertificado = dttDatosAlbaran.Rows(0)("NAlbaran")
    '        dataResult.IDModoTransporte = data.dtCabeceraAlbaran.Rows(0)("IDModoTransporte") & String.Empty
    '        dataResult.Transportista = data.dtCabeceraAlbaran.Rows(0)("EmpresaTransp") & String.Empty
    '        dataResult.NIFResponsableTransporte = data.dtCabeceraAlbaran.Rows(0)("DNIConductor") & String.Empty
    '        dataResult.ResponsableTransporte = data.dtCabeceraAlbaran.Rows(0)("Conductor") & String.Empty
    '        dataResult.FormaEnvio = data.dtCabeceraAlbaran.Rows(0)("IDFormaEnvio") & String.Empty
    '        dataResult.NumeroDocumento = data.dtCabeceraAlbaran.Rows(0)("NAlbaran") & String.Empty
    '        dataResult.Cliente = data.dtCabeceraAlbaran.Rows(0)("IDCliente") & String.Empty
    '        dataResult.Direccion = data.dtCabeceraAlbaran.Rows(0)("IDDireccion")
    '        dataResult.Matricula = data.dtCabeceraAlbaran.Rows(0)("Matricula") & String.Empty
    '        dataResult.Remolque = data.dtCabeceraAlbaran.Rows(0)("Remolque") & String.Empty
    '        dataResult.Precinto = data.dtCabeceraAlbaran.Rows(0)("Precinto") & String.Empty
    '        dataResult.Contenedor = data.dtCabeceraAlbaran.Rows(0)("NContenedor") & String.Empty

    '        dataResult.RegistrosEmpresas = New BdgDAACabecera.RegistroEmpresaInfo
    '        dataResult.RegistrosEmpresas.Add(data.dtCabeceraAlbaran.Rows(0)("IDAlbaran"), data.IDBaseDatosOrigen)

    '        dataResult.VistaOrigenLineas = BdgDAACabecera.CN_VistaDAAAlbaran
    '        dataResult.CampoAgrupacionOrigenLineasPorDefecto = "IDLineaAlbaran"
    '        dataResult.CampoID = "IDAlbaran"
    '        dataResult.Origen = BdgDAACabecera.enumOrigenDAA.Albaran

    '        dataResult.DireccionDestino = data.dtCabeceraAlbaran.Rows(0)("IDDireccion")
    '        dataResult.DireccionEntrega = data.dtCabeceraAlbaran.Rows(0)("IDDireccion")
    '        'dataResult.CodigoAduanaExportacion = data.CodigoAduanaExportacion

    '        'dataResult.OrigenExterno = data.dtLineasAlbaran

    '        Dim bdgDAAL As New BdgDAALinea
    '        Dim dtLineasDelete As DataTable = bdgDAAL.Filter(New GuidFilterItem("IDDaa", data.IDDaa))

    '        dataResult = ProcessServer.ExecuteTask(Of BdgDAACabecera.stCrearDAAInfo, BdgDAACabecera.stCrearDAAInfo)(AddressOf BdgDAACabecera.CrearDAALineas, dataResult, services)

    '        'TABLAS AUXILIARES: LINEAS, DOCUMENTO, ETC
    '        bdgDAAL.Delete(dtLineasDelete)
    '        bdgDAAL.Update(dataResult.Lineas)
    '        'ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf FinancieroGeneral.FormatoCuentaContable, data, services)


    '        'Dim datActBdg As New DataReconstruirDAALineas
    '        'datActBdg.IDBaseDatosOrigen = currentBBDD
    '        'datActBdg.IDBaseDatosDAA = drEmpresa("IDBaseDatos")
    '        'datActBdg.dtLineasAlbaran = BEDataEngine.Filter("NegBdgDAAAlbaranLineas", New NumberFilterItem("IDAlbaran", Doc.HeaderRow("IDAlbaran")))
    '        'datActBdg.IDAlbaran = Doc.HeaderRow("IDAlbaran")
    '    End If

    'End Sub

End Class


