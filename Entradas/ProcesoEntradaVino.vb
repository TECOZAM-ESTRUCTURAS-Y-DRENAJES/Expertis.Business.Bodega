Public Class ProcesoEntradaVino

#Region " B.Rules generales "

    <Task()> Public Shared Sub CambioPorcentaje(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CalculoCantidad, data, services)
    End Sub

    <Task()> Public Shared Sub CambioCantidad(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CalculoPorcentaje, data, services)
    End Sub

    <Task()> Public Shared Sub CalculoPorcentaje(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If data.Context.ContainsKey("Cantidad") Then
            Dim QEntradaVino As Double = Nz(data.Context("Cantidad"), 0)
            Dim Cantidad As Double
            If IsNumeric(data.Current("Cantidad")) Then
                Cantidad = data.Current("Cantidad")
            End If

            If QEntradaVino > 0 Then
                data.Current("Porcentaje") = Cantidad / QEntradaVino * 100
            Else
                data.Current("Porcentaje") = 0
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CalculoCantidad(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If data.Context.ContainsKey("Cantidad") Then
            Dim QEntradaVino As Double = Nz(data.Context("Cantidad"), 0)
            Dim Porcentaje As Double
            If IsNumeric(data.Current("Porcentaje")) Then
                Porcentaje = data.Current("Porcentaje")
            End If
            data.Current("Cantidad") = QEntradaVino * Porcentaje / 100
        End If
    End Sub

#End Region

#Region "Creación de Documentos"

    <Task()> Public Shared Function CrearDocumento(ByVal data As UpdatePackage, ByVal services As ServiceProvider) As DocumentoEntradaVino
        Return New DocumentoEntradaVino(data)
    End Function

#End Region

#Region "Eventos Update"


#Region "Comprobar Totales"

    <Task()> Public Shared Sub ComprobarTotales(ByVal doc As DocumentoEntradaVino, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DocumentoEntradaVino)(AddressOf ComprobarTotalDepositos, doc, services)
        ProcessServer.ExecuteTask(Of DocumentoEntradaVino)(AddressOf ComprobarTotalFincas, doc, services)
        ProcessServer.ExecuteTask(Of DocumentoEntradaVino)(AddressOf ComprobarTotalVariedades, doc, services)
    End Sub

    <Task()> Public Shared Sub ComprobarTotalDepositos(ByVal doc As DocumentoEntradaVino, ByVal services As ServiceProvider)
        Dim QCabecera As String = "Cantidad"
        Dim strMsg As String = Engine.ParseFormatString(AdminData.GetMessageText("El total de depósitos no coincide con la total {0}."), QCabecera)
        If (doc.dtDepositos Is Nothing AndAlso Nz(doc.HeaderRow(QCabecera), 0) <> 0) Then
            ApplicationService.GenerateError(strMsg)
        End If
        Dim dblTotal As Double = (Aggregate c In doc.dtDepositos Where c.RowState <> DataRowState.Deleted Into Sum(CDbl(c("Cantidad"))))
        If dblTotal <> 0 AndAlso dblTotal <> Nz(doc.HeaderRow(QCabecera), 0) Then
            ApplicationService.GenerateError(strMsg)
        End If
    End Sub

    <Task()> Public Shared Sub ComprobarTotalFincas(ByVal doc As DocumentoEntradaVino, ByVal services As ServiceProvider)
        Dim QCabecera As String = "Cantidad"
        Dim strMsg As String = Engine.ParseFormatString(AdminData.GetMessageText("El detalle de Fincas no coincide con la {0} de cabecera."), QCabecera)
        If (doc.dtFincas Is Nothing AndAlso Nz(doc.HeaderRow(QCabecera), 0) <> 0) Then
            ApplicationService.GenerateError(strMsg)
        End If

        Dim dblTotal As Double = 0
        If Not doc.dtFincas Is Nothing AndAlso doc.dtFincas.Rows.Count > 0 Then
            dblTotal = (Aggregate c In doc.dtFincas Where c.RowState <> DataRowState.Deleted Into Sum(CDbl(c("Porcentaje"))))
        End If

        If dblTotal <> 0 AndAlso dblTotal <> 100 Then
            ApplicationService.GenerateError(strMsg)
        End If
    End Sub

    <Task()> Public Shared Sub ComprobarTotalVariedades(ByVal doc As DocumentoEntradaVino, ByVal services As ServiceProvider)
        Dim QCabecera As String = "Cantidad"
        Dim strMsg As String = Engine.ParseFormatString(AdminData.GetMessageText("El detalle de Variedades no coincide con la {0} de cabecera."), QCabecera)
        If (doc.dtVariedades Is Nothing AndAlso Nz(doc.HeaderRow(QCabecera), 0) <> 0) Then
            ApplicationService.GenerateError(strMsg)
        End If

        Dim dblTotalVariedades As Double = 0
        If Not doc.dtVariedades Is Nothing AndAlso doc.dtVariedades.Rows.Count > 0 Then
            dblTotalVariedades = (Aggregate c In doc.dtVariedades Where c.RowState <> DataRowState.Deleted Into Sum(CDbl(c("Porcentaje"))))
        End If
        If dblTotalVariedades <> 0 AndAlso dblTotalVariedades <> 100 Then
            ApplicationService.GenerateError(strMsg)
        End If
    End Sub

#End Region

#Region "Eventos Cabecera - Entrada Vino"

    <Task()> Public Shared Sub AsignarContador(ByVal data As DocumentoEntradaVino, ByVal services As ServiceProvider)
        If data.HeaderRow.RowState = DataRowState.Added Then
            If Length(data.HeaderRow(_EVn.IDContador)) > 0 Then
                data.HeaderRow("NEntrada") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, data.HeaderRow("IDContador"), services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioFechaEntradaVino(ByVal data As DocumentoEntradaVino, ByVal services As ServiceProvider)
        If data.HeaderRow.RowState = DataRowState.Modified AndAlso Nz(data.HeaderRow("Fecha"), cnMinDate) <> Nz(data.HeaderRow("Fecha", DataRowVersion.Original), cnMinDate) Then
            Dim VinoDepositos As List(Of DataRow) = (From c In data.dtDepositos Where Not c.IsNull("IDVino") AndAlso c("IDVino") <> Guid.Empty Select c).ToList
            If Not VinoDepositos Is Nothing AndAlso VinoDepositos.Count > 0 Then
                Dim v As New BdgVino
                For Each drDeposito As DataRow In VinoDepositos
                    Dim IDVino As Guid = drDeposito("IDVino")
                    '//Modificar la fecha en el vino
                    Dim dtVino As DataTable = v.SelOnPrimaryKey(IDVino)
                    If dtVino.Rows.Count > 0 Then
                        dtVino.Rows(0)(_V.Fecha) = data.HeaderRow("Fecha")
                        v.Validate(dtVino)
                        v.Update(dtVino)
                    End If
                Next
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CalcularImporte(ByVal Data As DocumentoEntradaVino, ByVal services As ServiceProvider)
        Select Case Data.HeaderRow(_EVn.TipoPrecio)
            Case TipoPrecioContrato.PorLitro
                Data.HeaderRow(_EVn.Importe) = Data.HeaderRow(_EVn.Cantidad) * Data.HeaderRow(_EVn.Precio)
            Case TipoPrecioContrato.PorHectoGrado
                Dim strVariableGrado As String = New BdgParametro().VariableGradoVino()
                Dim StGrado As New BdgEntradaVino.StGetValorGrado(Data.HeaderRow(_EVn.NEntrada), strVariableGrado)
                Dim dblValorGrado As Double = ProcessServer.ExecuteTask(Of BdgEntradaVino.StGetValorGrado, Double)(AddressOf BdgEntradaVino.GetValorGrado, StGrado, services)
                If Double.IsNaN(dblValorGrado) Then
                    Data.HeaderRow(_EVn.Importe) = 0
                Else
                    Data.HeaderRow(_EVn.Importe) = Data.HeaderRow(_EVn.Cantidad) * (dblValorGrado / 100) * Data.HeaderRow(_EVn.Precio)
                End If
        End Select
        'Portes
        If Data.HeaderRow(_EVn.Cantidad) > 0 Then
            If Length(Data.HeaderRow(_EVn.IDContratoLinea)) = 0 Then
                'No hay contrato. Por pantalla sólo se puede introducir el importe
                Data.HeaderRow(_EVn.PrecioPorte) = Data.HeaderRow(_EVn.ImportePorte) / Data.HeaderRow(_EVn.Cantidad)
            Else
                'Hay contrato. Se vuelca el precio del contrato en pantalla.
                Data.HeaderRow(_EVn.ImportePorte) = Data.HeaderRow(_EVn.PrecioPorte) * Data.HeaderRow(_EVn.Cantidad)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarMovimientos(ByVal data As DocumentoEntradaVino, ByVal services As ServiceProvider)
        If data.HeaderRow.RowState = DataRowState.Modified Then
            If AreDifferents(data.HeaderRow(_EVn.Importe, DataRowVersion.Original), data.HeaderRow(_EVn.Importe)) _
            OrElse AreDifferents(data.HeaderRow(_EVn.ImportePorte, DataRowVersion.Original), data.HeaderRow(_EVn.ImportePorte)) OrElse _
            Nz(data.HeaderRow("Fecha")) <> Nz(data.HeaderRow("Fecha", DataRowVersion.Original)) Then
                ProcessServer.ExecuteTask(Of DataRow)(AddressOf BdgEntradaVino.ActualizarMovimientosEntrada, data.HeaderRow, services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarDescArticulo(ByVal data As DocumentoEntradaVino, ByVal services As ServiceProvider)
        If Length(data.HeaderRow(_EVn.DescArticulo)) = 0 Then
            data.HeaderRow(_EVn.DescArticulo) = New Articulo().GetItemRow(data.HeaderRow(_EVn.IDArticulo))(_EVn.DescArticulo)
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarContratoLinea(ByVal data As DocumentoEntradaVino, ByVal services As ServiceProvider)
        If Length(data.HeaderRow(_EVn.IDContratoLinea)) > 0 Then
            Dim bdgContrato As New BdgContratoLinea
            Dim rwContrato As DataRow = bdgContrato.GetItemRow(data.HeaderRow(_EVn.IDContratoLinea))
            If data.HeaderRow.RowState = DataRowState.Added Then
                rwContrato("QRecibida") += data.HeaderRow(_EVn.Cantidad)
            ElseIf data.HeaderRow.RowState = DataRowState.Modified Then
                If Not (DirectCast(data.HeaderRow(_EVn.IDContratoLinea), Guid).Equals(data.HeaderRow(_EVn.IDContratoLinea, DataRowVersion.Original))) Then
                    If Not IsDBNull(data.HeaderRow(_EVn.IDContratoLinea, DataRowVersion.Original)) AndAlso Not data.HeaderRow(_EVn.IDContratoLinea, DataRowVersion.Original).Equals(Guid.Empty) Then
                        Dim rwcontratoOrigen As DataRow = bdgContrato.GetItemRow(data.HeaderRow(_EVn.IDContratoLinea, DataRowVersion.Original))
                        If Nz(rwContrato("QRecibida"), 0) <> Nz(rwContrato("QRecibida", DataRowVersion.Original), 0) Then
                            rwcontratoOrigen("QRecibida") -= (data.HeaderRow(_EVn.Cantidad) - data.HeaderRow(_EVn.Cantidad, DataRowVersion.Original))
                        Else
                            rwcontratoOrigen("QRecibida") -= data.HeaderRow(_EVn.Cantidad)
                        End If
                        bdgContrato.Update(rwcontratoOrigen.Table)
                    End If
                    If Nz(rwContrato("QRecibida"), 0) <> Nz(rwContrato("QRecibida", DataRowVersion.Original), 0) Then
                        rwContrato("QRecibida") += (data.HeaderRow(_EVn.Cantidad) - data.HeaderRow(_EVn.Cantidad, DataRowVersion.Original))
                    Else
                        rwContrato("QRecibida") += data.HeaderRow(_EVn.Cantidad)
                    End If
                Else
                    rwContrato("QRecibida") += (Nz(data.HeaderRow(_EVn.Cantidad), 0) - Nz(data.HeaderRow(_EVn.Cantidad, DataRowVersion.Original), 0))
                End If
            End If
            bdgContrato.Update(rwContrato.Table)
        End If
    End Sub

#End Region

#Region "Eventos Entrada Vino Deposito"

    <Task()> Public Shared Sub AsignarVinosDepositos(ByVal data As DocumentoEntradaVino, ByVal services As ServiceProvider)
        Dim Precio As Double
        If data.HeaderRow(_EVn.Cantidad) > 0 Then
            Precio = (data.HeaderRow(_EVn.Importe) + data.HeaderRow(_EVn.ImportePorte)) / data.HeaderRow(_EVn.Cantidad)
        End If

        Dim Depositos As EntityInfoCache(Of BdgDepositoInfo) = services.GetService(Of EntityInfoCache(Of BdgDepositoInfo))()
        Dim DepositosEnBlanco As List(Of DataRow) = (From dpto In data.dtDepositos _
                                                     Where dpto.RowState <> DataRowState.Deleted AndAlso Length(dpto("IDDeposito")) = 0 _
                                                     Select dpto).ToList

        If Not DepositosEnBlanco Is Nothing AndAlso DepositosEnBlanco.Count > 0 Then
            ApplicationService.GenerateError("No se ha especificado un Depósito en alguna de las líneas de depósitos.")
        End If

        Dim ArticulosEnBlanco As List(Of DataRow) = (From dpto In data.dtDepositos _
                                                     Where (dpto.RowState = DataRowState.Added OrElse dpto.RowState = DataRowState.Modified) AndAlso Length(dpto("IDArticulo")) = 0 _
                                                     Select dpto).ToList

        If Not ArticulosEnBlanco Is Nothing AndAlso ArticulosEnBlanco.Count > 0 Then
            ApplicationService.GenerateError("No se ha especificado el Artículo en alguna de las líneas de depósitos.")
        End If

        Dim DepositosTratar As List(Of DataRow) = (From dpto In data.dtDepositos _
                                                   Where (dpto.RowState = DataRowState.Added OrElse dpto.RowState = DataRowState.Modified) AndAlso Length(dpto("IDDeposito")) > 0 _
                                                   Select dpto).ToList
        If DepositosTratar Is Nothing OrElse DepositosTratar.Count = 0 Then Exit Sub

        For Each Dr As DataRow In DepositosTratar
            Dim IDArticulo As String = Dr("IDArticulo") & String.Empty

            Dim VinoEnt As Guid
            If Dr.IsNull(_EVD.IDVino) Then
                VinoEnt = Guid.Empty
            Else
                VinoEnt = Dr(_EVD.IDVino)
            End If

            Dim Cantidad As Double = Nz(Dr(_EVD.Cantidad), 0)
            Dim CantidadOld As Double = 0
            If Dr.RowState = DataRowState.Modified AndAlso Nz(Dr(_EVD.Cantidad), 0) <> Nz(Dr(_EVD.Cantidad, DataRowVersion.Original), 0) Then
                CantidadOld = Nz(Dr(_EVD.Cantidad, DataRowVersion.Original), 0)
            End If

            Dim IDDeposito As String = String.Empty
            If Length(Dr(_EVD.IDDeposito)) > 0 Then
                IDDeposito = Dr(_EVD.IDDeposito)
            End If

            If Dr.RowState = DataRowState.Added Then
                Dim DptoInfo As BdgDepositoInfo = Depositos.GetEntity(IDDeposito)
                If DptoInfo.TipoDeposito = TipoDeposito.Barricas AndAlso DptoInfo.UsarBarricaComoLote Then
                    If Length(Dr("IDBarrica")) = 0 Then
                        ApplicationService.GenerateError("No se ha especificado un Tipo de Barrica para el Depósito {0}", IDDeposito)
                    Else
                        Dr("Lote") = Dr("IDBarrica")
                    End If
                End If
            End If

            If VinoEnt.Equals(Guid.Empty) Then
                Dim Lote As String
                If Length(Dr("Lote")) > 0 Then
                    Lote = Dr("Lote")
                Else
                    Lote = data.HeaderRow(_EVn.Lote) & String.Empty
                End If

                If Length(Lote) = 0 Then
                    '//Lote predeterminado
                    Lote = ProcessServer.ExecuteTask(Of String, String)(AddressOf BdgWorkClass.GetLotePredeterminado, Lote, services)
                End If

                Dim StCrear As New BdgWorkClass.StCrearVinoOrigenCantidad(IDDeposito, IDArticulo, Lote, data.HeaderRow(_EVn.Fecha), BdgOrigenVino.Compra, Cantidad)
                Dr(_EVD.IDVino) = ProcessServer.ExecuteTask(Of BdgWorkClass.StCrearVinoOrigenCantidad, Guid)(AddressOf BdgWorkClass.CrearVinoOrigenCantidad, StCrear, services)
            Else
                Dim StInc As New BdgWorkClass.StIncrementarIDVino(VinoEnt, Cantidad - CantidadOld)
                ProcessServer.ExecuteTask(Of BdgWorkClass.StIncrementarIDVino)(AddressOf BdgWorkClass.IncrementarCantidadIDVino, StInc, services)
                If Cantidad = 0 Then Dr(_EVD.IDVino) = DBNull.Value
            End If
        Next
    End Sub

    <Task()> Public Shared Sub AsignarMovimientosDepositos(ByVal data As DocumentoEntradaVino, ByVal services As ServiceProvider)
        '//se asume que todas las lineas son de la misma entrada
        Dim Precio As Double
        If data.HeaderRow(_EVn.Cantidad) > 0 Then
            Precio = (data.HeaderRow(_EVn.Importe) + data.HeaderRow(_EVn.ImportePorte)) / data.HeaderRow(_EVn.Cantidad)
        End If
        data.HeaderRow.Table.AcceptChanges() '// cambiamos el estado para que cuando estemos añadiendo podamos actualizar el movimiento
        If Length(data.HeaderRow(_EVn.IDMovimiento)) = 0 OrElse data.HeaderRow(_EVn.IDMovimiento) = 0 Then
            Dim StEj As New BdgWorkClass.StEjecutarMovimientos(data.HeaderRow(_EVn.NEntrada), data.HeaderRow(_EVn.Fecha), Precio)
            data.HeaderRow(_EVn.IDMovimiento) = ProcessServer.ExecuteTask(Of BdgWorkClass.StEjecutarMovimientos, Integer)(AddressOf BdgWorkClass.EjecutarMovimientos, StEj, services)
            BusinessHelper.UpdateTable(data.HeaderRow.Table)
        Else
            Dim StEj As New BdgWorkClass.StEjecutarMovimientosNumero(data.HeaderRow(_EVn.IDMovimiento), data.HeaderRow(_EVn.NEntrada), data.HeaderRow(_EVn.Fecha), Precio)
            ProcessServer.ExecuteTask(Of BdgWorkClass.StEjecutarMovimientosNumero)(AddressOf BdgWorkClass.EjecutarMovimientosNumero, StEj, services)
        End If
    End Sub


    <Task()> Public Shared Sub ActualizarEstadoVino(ByVal data As DocumentoEntradaVino, ByVal services As ServiceProvider)
        Dim ActualizarEstadoVino As List(Of DataRow) = (From c In data.dtDepositos _
                                                        Where Not c.IsNull("IDEstadoVino") AndAlso _
                                                             (c.RowState = DataRowState.Added OrElse _
                                                             (c.RowState = DataRowState.Modified AndAlso c("IDEstadoVino") & String.Empty <> c("IDEstadoVino", DataRowVersion.Original) & String.Empty))).ToList()

        For Each drVino As DataRow In ActualizarEstadoVino
            Dim datEstado As New BdgVino.StModificarEstado(drVino("IDVino"), drVino("IDEstadoVino"))
            ProcessServer.ExecuteTask(Of BdgVino.StModificarEstado)(AddressOf BdgVino.ModificarEstado, datEstado, services)
        Next
    End Sub

#End Region

#Region "Eventos Entrada Vino Analisis"

    <Task()> Public Shared Sub AsignarValorNumericoAnalisis(ByVal data As DocumentoEntradaVino, ByVal services As ServiceProvider)
        For Each Dr As DataRow In data.dtAnalisis.Select
            Dim rwVariable As DataRow = New BdgVariable().GetItemRow(Dr(_EVA.IDVariable))
            If CType(rwVariable(_Vr.TipoVariable), BdgTipoVariable) = BdgTipoVariable.Numerica Then
                If IsNumeric(Dr(_EVA.Valor)) Then
                    Dr(_EVA.ValorNumerico) = Double.Parse(Dr(_EVA.Valor))
                Else
                    'Dr(_EVA.Valor) = 0
                    Dr(_EVA.ValorNumerico) = System.DBNull.Value
                End If
            End If
        Next
    End Sub

#End Region

#End Region

End Class