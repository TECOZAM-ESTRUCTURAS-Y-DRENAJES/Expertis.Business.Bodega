Public Class BdgProcesoEntrada

    <Task()> Public Shared Function CrearDocumento(ByVal data As UpdatePackage, ByVal services As ServiceProvider) As BdgDocumentoEntrada
        Return New BdgDocumentoEntrada(data)
    End Function

    <Task()> Public Shared Sub AsignarValores(ByVal doc As BdgDocumentoEntrada, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of BdgDocumentoEntrada)(AddressOf AsignarClavePrimaria, doc, services)
        ProcessServer.ExecuteTask(Of BdgDocumentoEntrada)(AddressOf AsignarContador, doc, services)
        ProcessServer.ExecuteTask(Of BdgDocumentoEntrada)(AddressOf AsignarDatos, doc, services)
    End Sub

#Region "Comprobar Totales"
    <Task()> Public Shared Sub ComprobarTotales(ByVal doc As BdgDocumentoEntrada, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of BdgDocumentoEntrada)(AddressOf ComprobarTotalVariedades, doc, services)
        ProcessServer.ExecuteTask(Of BdgDocumentoEntrada)(AddressOf ComprobarTotalFincas, doc, services)
        ProcessServer.ExecuteTask(Of BdgDocumentoEntrada)(AddressOf ComprobarTotalCartillistas, doc, services)
        ProcessServer.ExecuteTask(Of BdgDocumentoEntrada)(AddressOf ComprobarTotalFacturacion, doc, services)
        ProcessServer.ExecuteTask(Of BdgDocumentoEntrada)(AddressOf ComprobarTotalDepositos, doc, services)
    End Sub

    <Task()> Public Shared Sub ComprobarTotalVariedades(ByVal doc As BdgDocumentoEntrada, ByVal services As ServiceProvider)
        Dim strKgsCabecera As String = "Neto"
        Dim strMsg As String = "El detalle de Variedades no coincide con el " & strKgsCabecera & " de cabecera."
        If (doc.dtVariedades Is Nothing AndAlso Nz(doc.HeaderRow(strKgsCabecera), 0) <> 0) Then
            ApplicationService.GenerateError(strMsg)
        End If
        Dim dblTotalVariedades As Double = Nz(doc.dtVariedades.Compute("SUM (Porcentaje)", Nothing), 0)
        If dblTotalVariedades <> 0 And dblTotalVariedades <> 100 Then
            ApplicationService.GenerateError(strMsg)
        End If
    End Sub

    <Task()> Public Shared Sub ComprobarTotalFincas(ByVal doc As BdgDocumentoEntrada, ByVal services As ServiceProvider)
        Dim strKgsCabecera As String = "Neto"
        Dim strMsg As String = "El detalle de Fincas no coincide con el " & strKgsCabecera & " de cabecera."
        If (doc.dtFincas Is Nothing AndAlso Nz(doc.HeaderRow(strKgsCabecera), 0) <> 0) Then
            ApplicationService.GenerateError(strMsg)
        End If
        Dim dblTotal As Double = Nz(doc.dtFincas.Compute("SUM (Porcentaje)", Nothing), 0)
        If dblTotal <> 0 And dblTotal <> 100 Then
            ApplicationService.GenerateError(strMsg)
        End If
    End Sub

    <Task()> Public Shared Sub ComprobarTotalCartillistas(ByVal doc As BdgDocumentoEntrada, ByVal services As ServiceProvider)
        Dim strKgsCabecera As String = "Declarado"
        Dim strMsg As String = "El detalle de Cartillistas no coincide con el " & strKgsCabecera & " de cabecera."
        If (doc.dtCartillistas Is Nothing AndAlso Nz(doc.HeaderRow(strKgsCabecera), 0) <> 0) Then
            ApplicationService.GenerateError(strMsg)
        End If
        Dim dblTotal As Double = Nz(doc.dtCartillistas.Compute("SUM (Porcentaje)", Nothing), 0)
        If dblTotal <> 0 And dblTotal <> 100 Then
            ApplicationService.GenerateError(strMsg)
        End If
    End Sub

    <Task()> Public Shared Sub ComprobarTotalFacturacion(ByVal doc As BdgDocumentoEntrada, ByVal services As ServiceProvider)
        If Not Nz(doc.HeaderRow("GestionManualFacturacion"), False) Then
            Return
        End If

        Dim strKgsCabecera As String = "Neto"
        Dim strMsg As String = "El detalle de Facturación no coincide con el " & strKgsCabecera & " de cabecera."
        If (doc.dtFacturacion Is Nothing AndAlso Nz(doc.HeaderRow(strKgsCabecera), 0) <> 0) Then
            ApplicationService.GenerateError(strMsg)
        End If

        Dim dblTotal As Double = Nz(doc.dtFacturacion.Compute("SUM (Declarado)", Nothing), 0)
        If dblTotal <> doc.HeaderRow(strKgsCabecera) Then
            ApplicationService.GenerateError(strMsg)
        End If
    End Sub

    <Task()> Public Shared Sub ComprobarTotalDepositos(ByVal doc As BdgDocumentoEntrada, ByVal services As ServiceProvider)
        If (doc.dtDepositos Is Nothing OrElse doc.dtDepositos.Rows.Count = 0) Then
            Return
        End If

        Dim strKgsCabecera As String = "Neto"
        Dim strMsg As String = "El total de depósitos no coincide con el total " & strKgsCabecera & "."
        If (doc.dtDepositos Is Nothing AndAlso Nz(doc.HeaderRow(strKgsCabecera), 0) <> 0) Then
            ApplicationService.GenerateError(strMsg)
        End If
        Dim dblTotal As Double = Nz(doc.dtDepositos.Compute("SUM (Cantidad)", Nothing), 0)
        If dblTotal <> Nz(doc.HeaderRow(strKgsCabecera), 0) Then
            ApplicationService.GenerateError(strMsg)
        End If
    End Sub

#End Region

#Region "Cabecera - asignar valores"

    <Task()> Public Shared Sub AsignarClavePrimaria(ByVal data As BdgDocumentoEntrada, ByVal services As ServiceProvider)
        If data.HeaderRow.RowState = DataRowState.Added Then
            If Length(data.HeaderRow(_E.IDEntrada)) = 0 Then data.HeaderRow(_E.IDEntrada) = AdminData.GetAutoNumeric()
        End If
    End Sub

    <Task()> Public Shared Sub AsignarContador(ByVal data As BdgDocumentoEntrada, ByVal services As ServiceProvider)
        If data.HeaderRow.RowState = DataRowState.Added Then
            If Length(data.HeaderRow(_E.IDContador)) > 0 Then
                data.HeaderRow(_E.NEntrada) = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, data.HeaderRow(_E.IDContador), services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarDatos(ByVal data As BdgDocumentoEntrada, ByVal services As ServiceProvider)
        data.HeaderRow(_E.TipoVariedad) = ProcessServer.ExecuteTask(Of String, Integer)(AddressOf TipoVariedad, data.HeaderRow(_E.IDVariedad), services)
        If data.HeaderRow.RowState = DataRowState.Added Then
            If Length(data.HeaderRow(_E.Fecha)) = 0 Then data.HeaderRow(_E.Fecha) = Today
            If Not IsDate(data.HeaderRow(_E.Hora)) Then data.HeaderRow(_E.Hora) = TimeOfDay
        End If
    End Sub

    <Task()> Public Shared Function TipoVariedad(ByVal IDVariedad As String, ByVal services As ServiceProvider) As Business.Bodega.BdgTipoVariedad
        Return New BdgVariedad().GetItemRow(IDVariedad)("TipoVariedad")
    End Function

#End Region

#Region "Analítica"

    <Task()> Public Shared Sub TratarAnalisis(ByVal doc As BdgDocumentoEntrada, ByVal services As ServiceProvider)
        If Not doc.dtAnalisis Is Nothing AndAlso doc.dtAnalisis.Rows.Count > 0 Then
            For Each dtrAnalisis As DataRow In doc.dtAnalisis.Rows
                Dim rwVariable As DataRow = New BdgVariable().GetItemRow(dtrAnalisis(_EA.IDVariable))
                If CType(rwVariable(_Vr.TipoVariable), BdgTipoVariable) = BdgTipoVariable.Numerica Then
                    If Length(dtrAnalisis(_EA.Valor)) > 0 Then
                        dtrAnalisis(_EA.ValorNumerico) = Double.Parse(dtrAnalisis(_EA.Valor))
                    Else : dtrAnalisis(_EA.ValorNumerico) = 0
                    End If
                End If
            Next
        End If
    End Sub

    <Task()> Public Shared Sub TratarFacturacion(ByVal doc As BdgDocumentoEntrada, ByVal services As ServiceProvider)
        For Each Dr As DataRow In doc.dtFacturacion.Select
            If Dr.RowState = DataRowState.Added Then
                If Length(Dr("IDEntradaFacturacion")) = 0 OrElse CType(Dr("IDEntradaFacturacion"), Guid).Equals(Guid.Empty) Then
                    Dr("IDEntradaFacturacion") = Guid.NewGuid
                End If
            End If
        Next
    End Sub

#End Region

#Region "Cartillista"

    <Task()> Public Shared Sub TratarCartillista(ByVal doc As BdgDocumentoEntrada, ByVal services As ServiceProvider)
        If doc.HeaderRow.RowState = DataRowState.Modified Then
            For Each dtrCartillista As DataRow In doc.dtCartillistas.Rows
                If Length(dtrCartillista(_EC.Vendimia)) = 0 Then dtrCartillista(_EC.Vendimia) = doc.HeaderRow(_EC.Vendimia)
                If Length(dtrCartillista(_EC.TipoVariedad)) = 0 Then dtrCartillista(_EC.TipoVariedad) = doc.HeaderRow(_EC.TipoVariedad)
            Next
        ElseIf doc.HeaderRow.RowState = DataRowState.Added Then
            Dim IntTipoVariedad As Integer = ProcessServer.ExecuteTask(Of String, Integer)(AddressOf TipoVariedad, doc.HeaderRow(_E.IDVariedad), services)
            Dim rwEC As DataRow = doc.dtCartillistas.NewRow
            rwEC(_EC.IDEntrada) = doc.HeaderRow(_E.IDEntrada)
            rwEC(_EC.IDCartillista) = doc.HeaderRow("IDCartillista")
            rwEC(_EC.Vendimia) = doc.HeaderRow("Vendimia")
            rwEC(_EC.TipoVariedad) = IntTipoVariedad
            rwEC(_EC.Porcentaje) = 100
            rwEC(_EC.Declarado) = doc.HeaderRow(_E.Declarado)
            rwEC(_EC.Talon) = Nz(doc.HeaderRow(_E.Talon), 1)
            doc.dtCartillistas.Rows.Add(rwEC)
        End If
    End Sub

#End Region

#Region "Depósitos"

    <Task()> Public Shared Sub TratarDepositos(ByVal doc As BdgDocumentoEntrada, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of BdgDocumentoEntrada)(AddressOf ActualizarVinos, doc, services)
        ProcessServer.ExecuteTask(Of BdgDocumentoEntrada)(AddressOf ActualizarMovimientos, doc, services)
    End Sub

    <Task()> Public Shared Sub ActualizarVinos(ByVal doc As BdgDocumentoEntrada, ByVal services As ServiceProvider)
        Dim DtChangesAdd As DataTable = doc.dtDepositos.GetChanges(DataRowState.Added)
        Dim DtChangesMod As DataTable = doc.dtDepositos.GetChanges(DataRowState.Modified)
        If (Not DtChangesMod Is Nothing AndAlso DtChangesMod.Rows.Count > 0) OrElse (Not DtChangesAdd Is Nothing AndAlso DtChangesAdd.Rows.Count > 0) Then
            Dim oVinoWC As New BdgWorkClass
            '//se asume que todas las lineas son de la misma entrada
            Dim ClsEntDep As New BdgEntradaDeposito
            For Each oRw As DataRow In doc.dtDepositos.Select(Nothing, Nothing, DataViewRowState.ModifiedCurrent)
                Dim intIDEntrada As Integer = oRw(_ED.IDEntrada, DataRowVersion.Original)
                Dim strIDDeposito As String = oRw(_ED.IDDeposito, DataRowVersion.Original)
                If strIDDeposito <> oRw(_ED.IDDeposito, DataRowVersion.Current) OrElse oRw(_ED.Cantidad, DataRowVersion.Original) = 0 Then
                    '//si cambia el deposito o la cantidad original era cero
                    Dim nwRwED As DataRow = doc.dtDepositos.NewRow
                    nwRwED(_ED.IDEntrada) = oRw(_ED.IDEntrada)
                    nwRwED(_ED.IDDeposito) = oRw(_ED.IDDeposito)
                    nwRwED(_ED.Cantidad) = oRw(_ED.Cantidad)
                    nwRwED(_ED.IDArticulo) = oRw(_ED.IDArticulo)
                    doc.dtDepositos.Rows.Add(nwRwED)
                    'ClsEntDep.Delete(oRw)
                    oRw.Delete()
                Else
                    Dim IDVino As Guid
                    Dim dblQ As Double = 0
                    Dim dblQN As Double = 0
                    If Length(oRw(_ED.IDVino)) > 0 Then
                        IDVino = oRw(_ED.IDVino)
                        dblQ = oRw(_ED.Cantidad, DataRowVersion.Original)
                        dblQN = oRw(_ED.Cantidad, DataRowVersion.Current)
                    End If
                    If Not IDVino.Equals(Guid.Empty) Then
                        Dim StInc As New BdgWorkClass.StIncrementarIDVino(IDVino, dblQN - dblQ)
                        ProcessServer.ExecuteTask(Of BdgWorkClass.StIncrementarIDVino)(AddressOf BdgWorkClass.IncrementarCantidadIDVino, StInc, services)
                        If dblQN = 0 Then
                            oRw(_ED.IDVino) = DBNull.Value
                        End If
                    End If
                End If
            Next
            For Each rwED As DataRow In doc.dtDepositos.Select(Nothing, Nothing, DataViewRowState.Added)
                Dim strIDArticulo As String
                If Length(rwED(_ED.IDArticulo)) = 0 Then
                    Dim intIDTipoUva As BdgTipoVariedad = doc.HeaderRow(_E.TipoVariedad)
                    Dim StArtUva As New BdgVendimia.StArticuloUva(doc.HeaderRow(_E.Vendimia), intIDTipoUva)
                    strIDArticulo = ProcessServer.ExecuteTask(Of BdgVendimia.StArticuloUva, String)(AddressOf BdgVendimia.ArticuloUva, StArtUva, services)
                Else
                    strIDArticulo = rwED(_ED.IDArticulo)
                End If
                Dim strLote As String
                If Length(rwED(_ED.Lote)) = 0 Then
                    strLote = ProcessServer.ExecuteTask(Of String, String)(AddressOf BdgWorkClass.GetLotePredeterminado, Nothing, services)
                Else
                    strLote = rwED(_ED.Lote)
                End If
                Dim dblQ As Double
                If Length(rwED(_ED.Cantidad)) = 0 Then
                    dblQ = 0
                Else
                    dblQ = rwED(_ED.Cantidad)
                End If
                If dblQ > 0 Then
                    Dim StCrear As New BdgWorkClass.StCrearVinoOrigenCantidad(rwED(_ED.IDDeposito), strIDArticulo, strLote, doc.HeaderRow(_E.Fecha), BdgOrigenVino.Uva, dblQ)
                    rwED(_ED.IDVino) = ProcessServer.ExecuteTask(Of BdgWorkClass.StCrearVinoOrigenCantidad, Guid)(AddressOf BdgWorkClass.CrearVinoOrigenCantidad, StCrear, services)
                End If
            Next
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarMovimientos(ByVal doc As BdgDocumentoEntrada, ByVal services As ServiceProvider)
        Dim DtChangesAdd As DataTable = doc.dtDepositos.GetChanges(DataRowState.Added)
        Dim DtChangesMod As DataTable = doc.dtDepositos.GetChanges(DataRowState.Modified)
        If (Not DtChangesMod Is Nothing AndAlso DtChangesMod.Rows.Count > 0) OrElse (Not DtChangesAdd Is Nothing AndAlso DtChangesAdd.Rows.Count > 0) Then
            If Length(doc.HeaderRow(_E.IDMovimiento)) = 0 OrElse doc.HeaderRow(_E.IDMovimiento) = 0 Then
                Dim StEj As New BdgWorkClass.StEjecutarMovimientos(doc.HeaderRow(_E.NEntrada), doc.HeaderRow(_E.Fecha))
                doc.HeaderRow(_E.IDMovimiento) = ProcessServer.ExecuteTask(Of BdgWorkClass.StEjecutarMovimientos, Integer)(AddressOf BdgWorkClass.EjecutarMovimientos, StEj, services)
            Else
                Dim StEj As New BdgWorkClass.StEjecutarMovimientosNumero(doc.HeaderRow(_E.IDMovimiento), doc.HeaderRow(_E.NEntrada), doc.HeaderRow(_E.Fecha))
                ProcessServer.ExecuteTask(Of BdgWorkClass.StEjecutarMovimientosNumero)(AddressOf BdgWorkClass.EjecutarMovimientosNumero, StEj, services)
            End If
        End If
    End Sub

#End Region

#Region "Fincas"

    <Task()> Public Shared Sub TratarFincas(ByVal doc As BdgDocumentoEntrada, ByVal services As ServiceProvider)
        If New BdgParametro().ValidarCupoFincas Then
            For Each DrFinca As DataRow In doc.dtFincas.Select
                If DrFinca.RowState = DataRowState.Added OrElse DrFinca.RowState = DataRowState.Modified Then
                    Dim FilEntradas As New Filter
                    FilEntradas.Add("IDCartillista", FilterOperator.Equal, doc.HeaderRow("IDCartillista"))
                    FilEntradas.Add("Vendimia", FilterOperator.Equal, doc.HeaderRow("Vendimia"))
                    FilEntradas.Add("IDEntrada", FilterOperator.NotEqual, doc.HeaderRow("IDEntrada"))
                    Dim DtEntradas As DataTable = New BdgEntrada().Filter(FilEntradas)
                    If Not DtEntradas Is Nothing AndAlso DtEntradas.Rows.Count > 0 Then
                        Dim FilEntFinca As New Filter(FilterUnionOperator.Or)
                        For Each DrEnt As DataRow In DtEntradas.Select
                            FilEntFinca.Add("IDEntrada", FilterOperator.Equal, DrEnt("IDEntrada"))
                        Next
                        Dim FilDeclarado As New Filter
                        FilDeclarado.Add(FilEntFinca)
                        FilDeclarado.Add(New GuidFilterItem("IDFinca", DrFinca("IDFinca")))
                        Dim DtDecla As DataTable = New BdgEntradaFinca().Filter(FilDeclarado)
                        If Not DtDecla Is Nothing AndAlso DtDecla.Rows.Count > 0 Then
                            Dim DblDeclarado As Double = DtDecla.Compute("SUM(Neto)", String.Empty)
                            DblDeclarado += DrFinca("Neto")
                            Dim FilCartFinca As New Filter
                            FilCartFinca.Add("IDCartillista", FilterOperator.Equal, doc.HeaderRow("IDCartillista"))
                            FilCartFinca.Add("Vendimia", FilterOperator.Equal, doc.HeaderRow("Vendimia"))
                            FilCartFinca.Add(New GuidFilterItem("IDFinca", DrFinca("IDFinca")))
                            Dim DtCartFinca As DataTable = New BdgCartillistaFinca().Filter(FilCartFinca)
                            If Not DtCartFinca Is Nothing AndAlso DtCartFinca.Rows.Count > 0 Then
                                Dim DblMaximo As Double = DtCartFinca.Rows(0)("Maximo")
                                If DblDeclarado > DblMaximo Then
                                    Dim DtFindFinca As DataTable = New BdgFinca().Filter(New GuidFilterItem("IDFinca", DrFinca("IDFinca")), String.Empty, "CFinca")
                                    ApplicationService.GenerateError("Se ha superado el cupo de la Finca: |.", DtFindFinca.Rows(0)("CFinca"))
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        End If
        If doc.HeaderRow.RowState = DataRowState.Added Then
            If Length(doc.HeaderRow(_E.IDFinca)) > 0 Then
                If Length(doc.HeaderRow(_E.IDFinca).ToString) > 0 AndAlso doc.HeaderRow(_E.IDEntrada) > 0 Then
                    Dim rwEF As DataRow = doc.dtFincas.NewRow
                    rwEF(_EF.IdEntrada) = doc.HeaderRow(_E.IDEntrada)
                    rwEF(_EF.IdFinca) = doc.HeaderRow(_E.IDFinca)
                    rwEF(_EF.Neto) = doc.HeaderRow(_E.Neto)
                    rwEF(_EF.Porcentaje) = 100
                    rwEF("Talon") = Nz(doc.HeaderRow(_E.Talon), 1)
                    doc.dtFincas.Rows.Add(rwEF)
                End If
            End If
        End If
    End Sub

#End Region

#Region "Proveedor"

    <Task()> Public Shared Sub TratarProveedores(ByVal doc As BdgDocumentoEntrada, ByVal services As ServiceProvider)
        If doc.HeaderRow.RowState = DataRowState.Added Then
            If Length(doc.HeaderRow(_E.IDFinca)) > 0 And Length(doc.HeaderRow(_E.IDCartillista)) > 0 Then
                Dim dtFP As DataTable = New BdgFincaProveedor().Filter(New GuidFilterItem("IdFinca", doc.HeaderRow(_E.IDFinca)))
                If dtFP.Rows.Count > 0 Then
                    For Each rwFP As DataRow In dtFP.Select
                        Dim rwEP As DataRow = doc.dtProveedores.NewRow
                        rwEP(_EP.IdEntrada) = doc.HeaderRow(_E.IDEntrada)
                        rwEP(_EP.IdProveedor) = rwFP(_FP.IdProveedor)
                        rwEP(_EF.Porcentaje) = rwFP(_FP.Porcentaje)
                        doc.dtProveedores.Rows.Add(rwEP)
                    Next
                Else
                    If Length(doc.HeaderRow(_E.IDCartillista)) > 0 AndAlso doc.HeaderRow(_E.IDEntrada) > 0 Then
                        Dim dtCar As DataTable = New BdgCartillista().SelOnPrimaryKey(doc.HeaderRow(_E.IDCartillista))
                        Dim IdProveedor As String
                        If dtCar.Rows.Count > 0 Then IdProveedor = dtCar.Rows(0)("IdProveedor")
                        Dim rwEP As DataRow = doc.dtProveedores.NewRow
                        rwEP(_EP.IdEntrada) = doc.HeaderRow(_E.IDEntrada)
                        rwEP(_EP.IdProveedor) = IdProveedor
                        rwEP(_EF.Porcentaje) = 100
                        doc.dtProveedores.Rows.Add(rwEP)
                    End If
                End If
            End If
        End If
    End Sub

#End Region

#Region "Variedad"

    <Task()> Public Shared Sub TratarVariedades(ByVal doc As BdgDocumentoEntrada, ByVal services As ServiceProvider)
        If doc.HeaderRow.RowState = DataRowState.Added Then
            If Length(doc.HeaderRow(_E.IDVariedad)) > 0 And doc.HeaderRow(_E.IDEntrada) > 0 Then
                Dim rwEV As DataRow = doc.dtVariedades.NewRow
                rwEV(_EV.IDEntrada) = doc.HeaderRow(_E.IDEntrada)
                rwEV(_EV.IDVariedad) = doc.HeaderRow(_E.IDVariedad)
                rwEV(_EV.Neto) = doc.HeaderRow(_E.Neto)
                rwEV(_EV.Porcentaje) = 100
                doc.dtVariedades.Rows.Add(rwEV)
            End If
        End If
    End Sub

#End Region

#Region "Post - update"

    <Task()> Public Shared Sub Recalculos(ByVal data As BdgDocumentoEntrada, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of BdgDocumentoEntrada)(AddressOf RecalcularEntradaFinca, data, services)
        ProcessServer.ExecuteTask(Of BdgDocumentoEntrada)(AddressOf RecalcularEntradaVariedad, data, services)
        ProcessServer.ExecuteTask(Of BdgDocumentoEntrada)(AddressOf RecalcularEntradaCartillista, data, services)
        'La cantidad asignada a cada depósito debe decidirla el usuario
        'ProcessServer.ExecuteTask(Of BdgDocumentoEntrada)(AddressOf RecalcularDepositos, data, services)
    End Sub

    <Task()> Public Shared Sub CambioFechaOperacionMovimiento(ByVal data As BdgDocumentoEntrada, ByVal services As ServiceProvider)
        If data.HeaderRow.RowState = DataRowState.Modified Then
            If Nz(data.HeaderRow("Fecha"), cnMinDate) <> Nz(data.HeaderRow("Fecha", DataRowVersion.Original), cnMinDate) AndAlso Length(data.HeaderRow("IDMovimiento")) > 0 Then
                Dim StModifEntrada As New StModificarFechaEntrada(data.HeaderRow("IDMovimiento"), data.HeaderRow("Fecha"), data.HeaderRow("IDEntrada"))
                ProcessServer.ExecuteTask(Of StModificarFechaEntrada)(AddressOf ModificarFechaEntrada, StModifEntrada, services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub RecalcularEntradaCartillista(ByVal data As BdgDocumentoEntrada, ByVal services As ServiceProvider)
        If Not data.HeaderRow.RowState = DataRowState.Deleted Then
            If data.HeaderRow.RowState = DataRowState.Modified AndAlso data.HeaderRow(_E.Declarado, DataRowVersion.Current) <> data.HeaderRow(_E.Declarado, DataRowVersion.Original) Then
                For Each rwEC As DataRow In data.dtCartillistas.Rows
                    rwEC(_EC.Declarado) = data.HeaderRow(_E.Declarado) * rwEC(_EC.Porcentaje) / 100
                Next
            End If
        End If
    End Sub

    <Task()> Public Shared Sub RecalcularEntradaFinca(ByVal data As BdgDocumentoEntrada, ByVal services As ServiceProvider)
        Dim PorcentajeTotal As Integer = 0
        Dim rwAux As DataRow
        Dim strKgsCabecera As String = _E.Neto
        If Not data.HeaderRow.RowState = DataRowState.Deleted Then
            For Each rwEF As DataRow In data.dtFincas.Select
                PorcentajeTotal += rwEF(_EF.Porcentaje)
                If rwEF(_EF.Porcentaje) = 100 Then rwAux = rwEF
                If rwEF.RowState = DataRowState.Modified AndAlso rwEF(_EF.IdFinca).Equals(data.HeaderRow(_E.IDFinca, DataRowVersion.Original)) Then
                    rwEF(_EF.IdFinca) = data.HeaderRow(_E.IDFinca)
                End If
            Next
            If PorcentajeTotal > 100 And Not rwAux Is Nothing Then
                Dim Porcentaje As Integer = 200 - PorcentajeTotal
                rwAux(_EF.Porcentaje) = Porcentaje
                rwAux(_EF.Neto) = data.HeaderRow(strKgsCabecera) * Porcentaje / 100
            End If
            If data.HeaderRow.RowState = DataRowState.Modified AndAlso data.HeaderRow(strKgsCabecera, DataRowVersion.Current) <> data.HeaderRow(strKgsCabecera, DataRowVersion.Original) Then
                For Each rwEF As DataRow In data.dtFincas.Select
                    rwEF(_EF.Neto) = data.HeaderRow(strKgsCabecera) * rwEF(_EF.Porcentaje) / 100
                Next
            End If
        End If
    End Sub

    <Task()> Public Shared Sub RecalcularEntradaVariedad(ByVal data As BdgDocumentoEntrada, ByVal services As ServiceProvider)
        Dim PorcentajeTotal As Integer = 0
        Dim rwAux As DataRow
        If Not data.HeaderRow.RowState = DataRowState.Deleted Then
            For Each rwEV As DataRow In data.dtVariedades.Rows
                PorcentajeTotal += rwEV(_EV.Porcentaje)
                If rwEV(_EF.Porcentaje) = 100 Then rwAux = rwEV
                If rwEV.RowState = DataRowState.Modified AndAlso rwEV(_EV.IDVariedad).Equals(data.HeaderRow(_E.IDVariedad, DataRowVersion.Original)) Then
                    rwEV(_EV.IDVariedad) = data.HeaderRow(_E.IDVariedad)
                End If
            Next
            If PorcentajeTotal > 100 And Not rwAux Is Nothing Then
                Dim Porcentaje As Integer = 200 - PorcentajeTotal
                rwAux(_EV.Porcentaje) = Porcentaje
                rwAux(_EV.Neto) = data.HeaderRow(_E.Neto) * Porcentaje / 100
            End If
            If data.HeaderRow.RowState = DataRowState.Modified AndAlso data.HeaderRow(_E.Neto, DataRowVersion.Current) <> data.HeaderRow(_E.Neto, DataRowVersion.Original) Then
                For Each rwEV As DataRow In data.dtVariedades.Rows
                    rwEV(_EV.Neto) = data.HeaderRow(_E.Neto) * rwEV(_EV.Porcentaje) / 100
                Next
            End If
        End If
    End Sub

    <Serializable()> _
    Public Class StModificarFechaEntrada
        Public Fecha As Date
        Public IDMovimiento As Integer
        Public IDEntrada As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDMovimiento As Integer, ByVal Fecha As Date, ByVal IDEntrada As String)
            Me.Fecha = Fecha
            Me.IDMovimiento = IDMovimiento
            Me.IDEntrada = IDEntrada
        End Sub
    End Class

    <Task()> Public Shared Sub ModificarFechaEntrada(ByVal data As StModificarFechaEntrada, ByVal services As ServiceProvider)
        Dim f As New Filter
        f.Add(New StringFilterItem(_E.IDEntrada, data.IDEntrada))
        Dim dtEntradaDeposito As DataTable = New BdgEntradaDeposito().Filter(f)
        For Each drEntradaDeposito As DataRow In dtEntradaDeposito.Rows
            '//Modificar la fecha en el vino
            Dim v As New BdgVino
            Dim dtVino As DataTable = v.SelOnPrimaryKey(drEntradaDeposito(_OV.IDVino))
            If dtVino.Rows.Count > 0 Then
                dtVino.Rows(0)(_V.Fecha) = data.Fecha
                v.Validate(dtVino)
                v.Update(dtVino)
            End If
        Next

        Dim dtOpHis As DataTable = AdminData.GetData("tbHistoricoMovimiento", New NumberFilterItem("IDMovimiento", data.IDMovimiento))
        For Each drMovimiento As DataRow In dtOpHis.Rows
            Dim dataCorreccion As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Corregir, drMovimiento("IDLineaMovimiento"), data.Fecha, False)
            ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento)(AddressOf ProcesoStocks.ActualizarMovimiento, dataCorreccion, services)
        Next
    End Sub

    <Task()> Public Shared Sub RecalcularEntradaFacturacion(ByVal data As BdgDocumentoEntrada, ByVal services As ServiceProvider)

        'Los datos de facturación se generan automáticamente a partir del detalle de los cartillistas.
        'También se tienen en cuenta los kilos sin papel de la cabecera (Neto - Declarado).
        'Si la gestión es manual es el usuario quien debe cuadrar las cantidades antes de guardar la entrada.

        If Not Nz(data.HeaderRow("GestionManualFacturacion"), False) Then
            For Each DrEntFact As DataRow In data.dtFacturacion.Rows
                DrEntFact.Delete()
            Next

            If Not data.dtCartillistas Is Nothing AndAlso data.dtCartillistas.Rows.Count > 0 Then
                Dim StrProveedor As String = String.Empty

                If Not data.dtCartillistas.Columns.Contains("IDProveedor") Then
                    data.dtCartillistas.Columns.Add("IDProveedor", GetType(String))
                End If
                Dim clsBdgCart As New BdgCartillista
                For Each DrCart As DataRow In data.dtCartillistas.Select()
                    Dim dtCart As DataTable = clsBdgCart.SelOnPrimaryKey(DrCart("IDCartillista"))
                    If dtCart.Rows.Count > 0 Then
                        DrCart("IDProveedor") = dtCart.Rows(0)("IDProveedor")
                    End If
                Next

                'Generar tantas líneas como diferentes proveedores existan
                For Each DrCart As DataRow In data.dtCartillistas.Select("", "IDProveedor")
                    If DrCart("IDProveedor") <> StrProveedor Then
                        Dim dtr As DataRow = data.dtFacturacion.NewRow
                        dtr("IDEntradaFacturacion") = Guid.NewGuid
                        dtr("IDEntrada") = data.HeaderRow("IDEntrada")
                        dtr("IDProveedor") = DrCart("IDProveedor")
                        dtr("Tipo") = enumTipoFacturacion.Declarado
                        dtr("TipoVariedad") = DrCart("TipoVariedad")
                        dtr("Declarado") = data.dtCartillistas.Compute("SUM(Declarado)", "IDProveedor = '" & DrCart("IDProveedor") & "'")
                        dtr("Porcentaje") = data.dtCartillistas.Compute("SUM(Porcentaje)", "IDProveedor = '" & DrCart("IDProveedor") & "'")
                        data.dtFacturacion.Rows.Add(dtr)

                        StrProveedor = DrCart("IDProveedor")
                    End If
                Next

                Dim dblExcedente As Double = data.HeaderRow("Neto") - data.HeaderRow("Declarado")
                If dblExcedente > 0 Then
                    'Añadir registros con el Excedente
                    For Each DrFact As DataRow In data.dtFacturacion.Select()
                        Dim dtr As DataRow = data.dtFacturacion.NewRow
                        dtr("IDEntradaFacturacion") = Guid.NewGuid
                        dtr("IDEntrada") = data.HeaderRow("IDEntrada")
                        dtr("IDProveedor") = DrFact("IDProveedor")
                        dtr("Tipo") = enumTipoFacturacion.Excedente
                        dtr("TipoVariedad") = DrFact("TipoVariedad")
                        dtr("Declarado") = (dblExcedente * DrFact("Porcentaje")) / 100
                        dtr("Porcentaje") = DrFact("Porcentaje")
                        data.dtFacturacion.Rows.Add(dtr)
                    Next

                    'Recalcular los porcentajes
                    Dim dblAcumPorc As Double
                    Dim i As Integer
                    Dim drsFact As DataRow() = data.dtFacturacion.Select()
                    For i = 0 To drsFact.Count - 1
                        If i <> drsFact.Count - 1 Then
                            drsFact(i)("Porcentaje") = Math.Round(drsFact(i)("Declarado") / data.HeaderRow("Neto") * 100, CntDecPorcentajeEntrada)
                            dblAcumPorc += drsFact(i)("Porcentaje")
                        Else
                            drsFact(i)("Porcentaje") = 100 - dblAcumPorc
                        End If
                    Next
                ElseIf dblExcedente < 0 Then
                    'Quitar la parte Proporcional ya que se Factura el Neto, no el Declarado.
                    For Each DrFact As DataRow In data.dtFacturacion.Select()
                        DrFact("Declarado") = DrFact("Declarado") + (dblExcedente * DrFact("Porcentaje")) / 100
                    Next
                End If
            End If
        End If
    End Sub

#End Region

#Region "BusinessRules"

    Public Const CntDecPorcentajeEntrada As Integer = 8

    <Task()> Public Shared Sub CambioPorcentaje(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        Dim DblCantidad As Double = 0
        Dim DblPctj As Double = 0

        If data.Context.ContainsKey("Declarado") AndAlso Length(data.Context("Declarado")) > 0 Then
            DblCantidad = data.Context("Declarado")
        ElseIf data.Context.ContainsKey("Neto") AndAlso Length(data.Context("Neto")) > 0 Then
            DblCantidad = data.Context("Neto")
        End If
        If Length(data.Value) > 0 Then
            DblPctj = data.Value
        Else : ApplicationService.GenerateError("El campo Porcentaje debe ser numérico.")
        End If
        Dim dblQ As Double = DblCantidad * DblPctj / 100
        data.Current("Porcentaje") = Math.Round(DblPctj, CntDecPorcentajeEntrada)
        data.Current("Cantidad") = dblQ
    End Sub

    <Task()> Public Shared Sub CambioCantidad(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        Dim DblCantidad As Double = 0
        Dim DblPctj As Double = 0
        Dim DblQ As Double = 0
        If data.Context.ContainsKey("Declarado") AndAlso Length(data.Context("Declarado")) > 0 Then
            DblCantidad = data.Context("Declarado")
        ElseIf data.Context.ContainsKey("Neto") AndAlso Length(data.Context("Neto")) > 0 Then
            DblCantidad = data.Context("Neto")
        End If
        If Length(data.Value) > 0 Then
            DblQ = data.Value
        Else : ApplicationService.GenerateError("El campo Declarado debe ser numérico.")
        End If
        If DblCantidad > 0 Then DblPctj = DblQ / DblCantidad * 100
        data.Current("Porcentaje") = Math.Round(DblPctj, CntDecPorcentajeEntrada)
        data.Current("Cantidad") = DblQ
    End Sub

#End Region

End Class