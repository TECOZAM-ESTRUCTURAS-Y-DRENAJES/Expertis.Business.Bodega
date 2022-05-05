Public Class ProcesoBdgOperacion

#Region "Tareas Update BdgOperacion"

    <Task()> Public Shared Function CrearDocumento(ByVal data As UpdatePackage, ByVal services As ServiceProvider) As DocumentoBdgOperacionReal
        Return New DocumentoBdgOperacionReal(data)
    End Function

#Region "Tareas cabecera de Operación"

#Region "Tareas Validaciones Documento"

    <Task()> Public Shared Sub ValidacionesGeneralesOperacion(ByVal data As DocumentoBdgOperacionReal, ByVal services As ServiceProvider)
        If Not data.dtOperacionVino Is Nothing Then
            ProcessServer.ExecuteTask(Of DocumentoBdgOperacionReal)(AddressOf ValidarOperacionCabecera, data, services)
            ProcessServer.ExecuteTask(Of DocumentoBdgOperacionReal)(AddressOf ValidacionesOrigenDestino, data, services)
            ProcessServer.ExecuteTask(Of DocumentoBdgOperacionReal)(AddressOf ValidarOperacionVinoMateriales, data, services)
            ProcessServer.ExecuteTask(Of DocumentoBdgOperacionReal)(AddressOf ValidarOperacionVinoMaterialesLotes, data, services)
            ProcessServer.ExecuteTask(Of DocumentoBdgOperacionReal)(AddressOf ValidarOperacionMaterialesGlobalesLotes, data, services)
        End If
    End Sub

    <Task()> Public Shared Sub ValidarOperacionCabecera(ByVal data As DocumentoBdgOperacionReal, ByVal services As ServiceProvider)
        If data.GetTipoOperRequiereOF AndAlso (Length(data.HeaderRow("IDOrden")) = 0 OrElse data.HeaderRow("CantidadOrden") <= 0) Then
            ApplicationService.GenerateError("Debe introducir Orden y cantidad de la Orden para el tipo de Operación seleccionado.")
        End If
    End Sub

#Region "Validaciones - Operacion Vino Origen/Destino"

    <Task()> Public Shared Sub ValidacionesOrigenDestino(ByVal data As DocumentoBdgOperacionReal, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DocumentoBdgOperacionReal)(AddressOf ValidarOrigenMerma, data, services)
        ProcessServer.ExecuteTask(Of DocumentoBdgOperacionReal)(AddressOf ValidarDestinoBarricas, data, services)
        ProcessServer.ExecuteTask(Of DocumentoBdgOperacionReal)(AddressOf ValidarDestinoJaulones, data, services)
        ProcessServer.ExecuteTask(Of DocumentoBdgOperacionReal)(AddressOf ValidarOrigenDestinoTipoOperacion, data, services)
        ProcessServer.ExecuteTask(Of DocumentoBdgOperacionReal)(AddressOf ValidarOrigenDestinoIDVino, data, services)
        ProcessServer.ExecuteTask(Of DocumentoBdgOperacionReal)(AddressOf ValidarDestinoArticuloTipoDeposito, data, services)
        ProcessServer.ExecuteTask(Of DocumentoBdgOperacionReal)(AddressOf ValidarOrigenArticuloFecha, data, services)
        ProcessServer.ExecuteTask(Of DocumentoBdgOperacionReal)(AddressOf ValidarOrigenDestinoCantidades, data, services)
        ProcessServer.ExecuteTask(Of DocumentoBdgOperacionReal)(AddressOf ValidarDestinoRepartoCoste, data, services)
    End Sub

    <Task()> Public Shared Sub ValidarOrigenMerma(ByVal data As DocumentoBdgOperacionReal, ByVal services As ServiceProvider)
        If Not data.GetOperacionVinoOrigen Is Nothing AndAlso data.GetOperacionVinoOrigen.Length > 0 Then
            Dim DblTotalOrigen As Double = ProcessServer.ExecuteTask(Of DocumentoBdgOperacionReal, Double)(AddressOf GetOperacionTotalOrigen, data, services)
            If DblTotalOrigen > 0 Then
                Dim DblTotalDestino As Double = ProcessServer.ExecuteTask(Of DocumentoBdgOperacionReal, Double)(AddressOf GetOperacionTotalDestino, data, services)
                Dim DblTotalPendiente As Double = ProcessServer.ExecuteTask(Of DocumentoBdgOperacionReal, Double)(AddressOf GetOperacionTotalPendiente, data, services)
                Dim StData As New Business.Bodega.BdgOperacion.dataAplicarRulesMermaOperacion(data.GetTipoOperPermitirMerma, data.GetTipoOperMermaMax, DblTotalPendiente, DblTotalOrigen, DblTotalDestino)
                Dim CancelMerma As Boolean = ProcessServer.ExecuteTask(Of Business.Bodega.BdgOperacion.dataAplicarRulesMermaOperacion, Boolean)(AddressOf Business.Bodega.BdgOperacion.ValidarMermaOperacion, StData, services)
                If Not CancelMerma Then
                    ApplicationService.GenerateError("La Merma supera el porcentaje máximo permitido | %.", Strings.Format(data.GetTipoOperMermaMax, "#,#0.00"))
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarDestinoBarricas(ByVal data As DocumentoBdgOperacionReal, ByVal services As ServiceProvider)
        If data.GetTipoOperTipoMov <> Business.Bodega.TipoMovimiento.SinMovimiento Then
            Dim strDepositosSinBarrica As String = String.Empty
            If Not data.GetOperacionVinoDestino Is Nothing AndAlso data.GetOperacionVinoDestino.Length > 0 Then
                For Each drDestino As DataRow In data.GetOperacionVinoDestino
                    If drDestino.RowState <> DataRowState.Deleted Then
                        If Length(drDestino("IDDeposito")) > 0 Then
                            Dim blnUsarBarricaComoLote As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf Business.Bodega.BdgDeposito.UsarBarricaComoLote, drDestino("IDDeposito"), services)
                            If blnUsarBarricaComoLote AndAlso Length(drDestino("IDBarrica")) = 0 Then
                                If Length(strDepositosSinBarrica) > 0 Then strDepositosSinBarrica &= ","
                                strDepositosSinBarrica &= drDestino("IDDeposito")
                            End If
                        End If
                    End If
                Next
                If Length(strDepositosSinBarrica) > 0 Then
                    If strDepositosSinBarrica.LastIndexOf(",") > 0 Then
                        ApplicationService.GenerateError("En los Depósitos '|', la Barrica es obligatoria.", strDepositosSinBarrica)
                    Else : ApplicationService.GenerateError("En el Depósito '|', la Barrica es obligatoria.", strDepositosSinBarrica)
                    End If
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarDestinoJaulones(ByVal data As DocumentoBdgOperacionReal, ByVal services As ServiceProvider)
        If Not data.GetOperacionVinoDestino Is Nothing AndAlso data.GetOperacionVinoDestino.Length > 0 Then
            Dim DestinosActuales As List(Of DataRow) = (From c In data.GetOperacionVinoDestino Where c.RowState <> DataRowState.Deleted Select c).ToList
            Dim lstBotelleroSinJaulon As New List(Of String)
            For Each drDestino As DataRow In DestinosActuales
                If Length(drDestino("IDDeposito")) > 0 Then
                    Dim blnRequerirJaulon As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf Business.Bodega.BdgDeposito.RequerirJaulon, drDestino("IDDeposito"), services)
                    If blnRequerirJaulon AndAlso (Length(drDestino("JaulonDesde")) = 0 OrElse Length(drDestino("JaulonHasta")) = 0) Then
                        If Not lstBotelleroSinJaulon.Contains(drDestino("IDDeposito")) Then lstBotelleroSinJaulon.Add(drDestino("IDDeposito"))
                    End If
                End If
            Next

            If Not lstBotelleroSinJaulon Is Nothing AndAlso lstBotelleroSinJaulon.Count > 0 Then
                If lstBotelleroSinJaulon.Count > 1 Then
                    ApplicationService.GenerateError("En los Depósitos {0}, los jaulones son obligatorios.", Quoted(Strings.Join(lstBotelleroSinJaulon.ToArray, ",")))
                Else : ApplicationService.GenerateError("En el Depósito {0}, los jaulones son obligatorios.", Quoted(lstBotelleroSinJaulon(0)))
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarOrigenDestinoTipoOperacion(ByVal data As DocumentoBdgOperacionReal, ByVal services As ServiceProvider)
        Dim DrOrigenAux() As DataRow = data.GetOperacionVinoOrigen
        Dim DrDestinoAux() As DataRow = data.GetOperacionVinoDestino
        Dim cOperacionNoPermitida As String = "Operación no permitida por el Tipo de Operación."
        Select Case data.GetTipoOperTipoMov
            Case Business.Bodega.TipoMovimiento.SinMovimiento, Business.Bodega.TipoMovimiento.SinOrigen
                If Not DrOrigenAux Is Nothing AndAlso DrOrigenAux.Length <> 0 Then ApplicationService.GenerateError(cOperacionNoPermitida)
                If DrDestinoAux Is Nothing OrElse DrDestinoAux.Length = 0 Then ApplicationService.GenerateError(cOperacionNoPermitida)
            Case Business.Bodega.TipoMovimiento.DeUnoAUno
                If DrOrigenAux Is Nothing OrElse DrOrigenAux.Length <> 1 Then ApplicationService.GenerateError(cOperacionNoPermitida)
                If DrDestinoAux Is Nothing OrElse DrDestinoAux.Length <> 1 Then ApplicationService.GenerateError(cOperacionNoPermitida)
            Case Business.Bodega.TipoMovimiento.DeUnoAVarios
                If DrOrigenAux Is Nothing OrElse DrOrigenAux.Length <> 1 Then ApplicationService.GenerateError(cOperacionNoPermitida)
                If DrDestinoAux Is Nothing OrElse DrDestinoAux.Length = 0 Then ApplicationService.GenerateError(cOperacionNoPermitida)
            Case Business.Bodega.TipoMovimiento.DeVariosAUno
                If DrOrigenAux Is Nothing OrElse DrOrigenAux.Length = 0 Then ApplicationService.GenerateError(cOperacionNoPermitida)
                If DrDestinoAux Is Nothing OrElse DrDestinoAux.Length <> 1 Then ApplicationService.GenerateError(cOperacionNoPermitida)
            Case Business.Bodega.TipoMovimiento.Salida
                If DrOrigenAux Is Nothing OrElse DrOrigenAux.Length = 0 Then ApplicationService.GenerateError(cOperacionNoPermitida)
                If Not DrDestinoAux Is Nothing AndAlso DrDestinoAux.Length > 0 Then ApplicationService.GenerateError(cOperacionNoPermitida)
            Case Business.Bodega.TipoMovimiento.DeVariosAVarios, Business.Bodega.TipoMovimiento.CrearOrigen
                If DrOrigenAux Is Nothing OrElse DrOrigenAux.Length <= 0 Then ApplicationService.GenerateError(cOperacionNoPermitida)
                If DrDestinoAux Is Nothing OrElse DrDestinoAux.Length <= 0 Then ApplicationService.GenerateError(cOperacionNoPermitida)
        End Select
    End Sub

    <Task()> Public Shared Sub ValidarOrigenDestinoIDVino(ByVal data As DocumentoBdgOperacionReal, ByVal services As ServiceProvider)
        'Comprobar que un Vino no esté en Origen y Destino
        For Each DrOri As DataRow In data.GetOperacionVinoOrigen
            Dim StVal As New BdgOperacion.StValidarVinoRepetidoOperacion(DrOri, data.GetOperacionVinoDestino)
            ProcessServer.ExecuteTask(Of BdgOperacion.StValidarVinoRepetidoOperacion)(AddressOf BdgOperacion.ValidarVinoRepetidoOperacion, StVal, services)

            StVal = New BdgOperacion.StValidarVinoRepetidoOperacion(DrOri, data.GetOperacionVinoOrigen)
            ProcessServer.ExecuteTask(Of BdgOperacion.StValidarVinoRepetidoOperacion)(AddressOf BdgOperacion.ValidarVinoRepetidoOperacion, StVal, services)
        Next
        For Each DrDest As DataRow In data.GetOperacionVinoDestino
            Dim StVal As New BdgOperacion.StValidarVinoRepetidoOperacion(DrDest, data.GetOperacionVinoOrigen)
            ProcessServer.ExecuteTask(Of BdgOperacion.StValidarVinoRepetidoOperacion)(AddressOf BdgOperacion.ValidarVinoRepetidoOperacion, StVal, services)

            StVal = New BdgOperacion.StValidarVinoRepetidoOperacion(DrDest, data.GetOperacionVinoDestino)
            ProcessServer.ExecuteTask(Of BdgOperacion.StValidarVinoRepetidoOperacion)(AddressOf BdgOperacion.ValidarVinoRepetidoOperacion, StVal, services)
        Next
    End Sub

    <Task()> Public Shared Sub ValidarDestinoArticuloTipoDeposito(ByVal data As DocumentoBdgOperacionReal, ByVal services As ServiceProvider)
        Dim AppParams As BdgParametrosOperaciones = services.GetService(Of BdgParametrosOperaciones)()
        Dim LoteDef As String = AppParams.LotePorDefecto
        Dim LoteExplicitoEnBotellero As Boolean = AppParams.LoteExplicitoEnBotellero

        Dim Depositos As EntityInfoCache(Of BdgDepositoInfo) = services.GetService(Of EntityInfoCache(Of BdgDepositoInfo))()

        For Each DrDestino As DataRow In data.GetOperacionVinoDestino
            'Comprobar que los artículos están bien configurados
            ProcessServer.ExecuteTask(Of String)(AddressOf BdgVino.ValidarArticuloVino, DrDestino("IDArticulo"), services)
            'comprobar los tipos de depositos permitidos por el tipo de operacion
            If Length(DrDestino("IDDeposito")) = 0 Then ApplicationService.GenerateError("No se ha especificado un depósito destino")

            Dim DptoInfo As BdgDepositoInfo = Depositos.GetEntity(DrDestino("IDDeposito"))
            If DptoInfo.TipoDeposito = TipoDeposito.Almacen OrElse DptoInfo.TipoDeposito = TipoDeposito.Botellero Then
                If LoteExplicitoEnBotellero AndAlso Length(DrDestino("Lote")) = 0 Then
                    ApplicationService.GenerateError("No se ha especificado un lote en el destino para el artículo |", DrDestino("IDArticulo"))
                End If
            ElseIf DptoInfo.TipoDeposito = TipoDeposito.Barricas Then
                If DptoInfo.UsarBarricaComoLote AndAlso data.GetTipoOperTipoMov <> Business.Bodega.TipoMovimiento.SinMovimiento Then
                    If Length(DrDestino("IDBarrica")) > 0 Then
                        DrDestino("Lote") = DrDestino("IDBarrica")
                    Else : ApplicationService.GenerateError("No se ha especificado un lote de barricas en el destino para el artículo |", DrDestino("IDArticulo"))
                    End If
                End If
            End If
            If Length(DrDestino("Lote")) = 0 Then DrDestino("Lote") = LoteDef
        Next
    End Sub

    <Task()> Public Shared Sub ValidarOrigenArticuloFecha(ByVal data As DocumentoBdgOperacionReal, ByVal services As ServiceProvider)
        For Each DrOrigen As DataRow In data.GetOperacionVinoOrigen
            'Comprobar que los artículos están bien configurados
            ProcessServer.ExecuteTask(Of String)(AddressOf BdgVino.ValidarArticuloVino, DrOrigen("IDArticulo"), services)
            'Comprobar que la Fecha de la Operación sea igual o posterior a la Fecha de los vinos de Origen.
            If DrOrigen.Table.Columns.Contains("Fecha") Then
                Dim StVal As New StValidarFechaOperacion(DrOrigen("Fecha"), data.HeaderRow("Fecha"))
                ProcessServer.ExecuteTask(Of StValidarFechaOperacion)(AddressOf ValidarFechaOperacion, StVal, services)
            Else
                Dim StVal As New StValidarFechaOperacion(ProcessServer.ExecuteTask(Of Guid, Date)(AddressOf BdgVino.ObtenerFechaVino, DrOrigen("IDVino"), services), data.HeaderRow("Fecha"))
                ProcessServer.ExecuteTask(Of StValidarFechaOperacion)(AddressOf ValidarFechaOperacion, StVal, services)
            End If
        Next
    End Sub

    <Serializable()> _
    Public Class StValidarFechaOperacion
        Public FechaVino As Date
        Public FechaOperacion As Date

        Public Sub New()
        End Sub

        Public Sub New(ByVal FechaVino As Date, ByVal FechaOperacion As Date)
            Me.FechaVino = FechaVino
            Me.FechaOperacion = FechaOperacion
        End Sub
    End Class

    <Task()> Public Shared Function ValidarFechaOperacion(ByVal data As StValidarFechaOperacion, ByVal services As ServiceProvider) As Boolean
        'Comprobar si la Fecha del Vino es posterior a la Fecha de la Operación.
        Dim AppParams As BdgParametrosOperaciones = services.GetService(Of BdgParametrosOperaciones)()
        If AppParams.ComprobarFechaOperacion Then
            If data.FechaVino > data.FechaOperacion Then
                ApplicationService.GenerateError("La Fecha de la Operación {0} es anterior a la Fecha del Vino Origen {1}.", Format(data.FechaOperacion, "dd/MM/yyyy"), Format(data.FechaVino, "dd/MM/yyyy"))
                Return False
            Else : Return True
            End If
        Else : Return True
        End If
    End Function

    <Task()> Public Shared Sub ValidarOrigenDestinoCantidades(ByVal data As DocumentoBdgOperacionReal, ByVal services As ServiceProvider)
        If data.GetTipoOperTipoMov <> Business.Bodega.TipoMovimiento.SinMovimiento AndAlso data.GetTipoOperTipoMov <> Business.Bodega.TipoMovimiento.SinOrigen AndAlso data.GetTipoOperTipoMov <> Business.Bodega.TipoMovimiento.Salida Then
            Dim IDUdMedida As String = String.Empty
            Dim dblQ As Double
            Dim dblQOrigen As Double = 0
            For Each DrOrigen As DataRow In data.GetOperacionVinoOrigen
                If Length(DrOrigen("IDUDMedida")) = 0 Then ApplicationService.GenerateError("El artículo no tiene una unidad de medida en el Origen")
                If Length(IDUdMedida) = 0 Then IDUdMedida = DrOrigen("IDUdMedida")
                Dim dblQO As Double = 0
                If Length(DrOrigen("Cantidad")) > 0 Then dblQO = DrOrigen("Cantidad")
                If Length(DrOrigen("Merma")) > 0 Then dblQO -= DrOrigen("Merma")
                If IDUdMedida <> DrOrigen("IDUdMedida") Then
                    Dim dataFactor As New ArticuloUnidadAB.DatosFactorConversion(DrOrigen("IDArticulo"), DrOrigen("IDUdMedida"), IDUdMedida, False)
                    Dim f As Double = ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, dataFactor, services)
                    If f = 0 Then ApplicationService.GenerateError(Engine.Global.ParseFormatString("No existe conversión entre las unidades: {0} y {1}.", DrOrigen("IDUdMedida"), IDUdMedida))
                    dblQ += dblQO * f
                Else : dblQ += dblQO
                End If
            Next
            dblQOrigen = dblQ

            Dim dblQD As Double
            Dim dblQDTotal As Double
            Dim dblQNetaTot As Double = dblQ
            Dim DicVinoFactor As New Dictionary(Of Guid, Double)
            For Each oRw As DataRow In data.GetOperacionVinoDestino
                'Dim dblQD As Double
                If oRw.IsNull("IDUdMedida") Then ApplicationService.GenerateError("El artículo no tiene una unidad de medida en el Destino")
                If IDUdMedida <> oRw(_V.IDUdMedida) Then
                    Dim dataFactor As New ArticuloUnidadAB.DatosFactorConversion(oRw(_V.IDArticulo), oRw(_V.IDUdMedida), IDUdMedida, False)
                    Dim f As Double = ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, dataFactor, services)
                    If f = 0 Then ApplicationService.GenerateError(Engine.Global.ParseFormatString("No existe conversión entre las unidades: {0} y {1}.", oRw(_V.IDUdMedida), IDUdMedida))
                    dblQD = oRw("Cantidad") * f
                Else : dblQD = oRw("Cantidad")
                End If
                dblQDTotal += dblQD
                dblQ -= dblQD
                If dblQNetaTot <> 0 Then DicVinoFactor.Add(oRw("IDVino"), dblQD / dblQNetaTot)
            Next

            If dblQ * Math.Sign(dblQ) >= 0.01 Then ApplicationService.GenerateError(Engine.Global.ParseFormatString("No coinciden las cantidades Origen y Destino.{0} QOrigen: {1} {2} {3} QDestino: {4} {5}", vbNewLine, dblQOrigen, IDUdMedida, vbNewLine, dblQDTotal, IDUdMedida))

            If Not DicVinoFactor Is Nothing AndAlso DicVinoFactor.Count > 0 Then
                Dim DataPrcBdgOper As FactoresRecalcularEstructura = services.GetService(Of FactoresRecalcularEstructura)()
                DataPrcBdgOper.VinoFactor = DicVinoFactor
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarDestinoRepartoCoste(ByVal data As DocumentoBdgOperacionReal, ByVal services As ServiceProvider)
        If data.GetUsarRepartoTipoOperacion Then
            If data.GetTipoOperTipoMov <> Business.Bodega.TipoMovimiento.SinMovimiento _
                        And data.GetTipoOperTipoMov <> Business.Bodega.TipoMovimiento.SinOrigen _
                        And data.GetTipoOperTipoMov <> Business.Bodega.TipoMovimiento.Salida Then
                Dim dblPorcentaje As Double = 0
                If Not data.GetOperacionVinoDestino Is Nothing AndAlso data.GetOperacionVinoDestino.Length > 0 Then
                    Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
                    For Each drDestino As DataRow In data.GetOperacionVinoDestino
                        If drDestino.RowState <> DataRowState.Deleted Then
                            Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(drDestino("IDArticulo"))
                            If Not ArtInfo Is Nothing Then
                                Dim f As New Filter
                                f.Add(New StringFilterItem("IDTipoOperacion", data.HeaderRow("IDTipoOperacion")))
                                f.Add(New StringFilterItem("IDTipo", ArtInfo.IDTipo))
                                f.Add(New StringFilterItem("IDFamilia", ArtInfo.IDFamilia))
                                Dim dtReparto As DataTable = New BdgTipoOperacionRepartoCoste().Filter(f)
                                If Not dtReparto Is Nothing AndAlso dtReparto.Rows.Count > 0 Then
                                    dblPorcentaje += dtReparto.Rows(0)("Porcentaje")
                                Else
                                    ApplicationService.GenerateError("El Tipo de Operación '|' tiene Reparto de Costes.|Compruebe que el Tipo y la Familia del Artículo '|' en Destino se corresponde con el establecido en el Tipo de Operación.", data.HeaderRow("IDTipoOperacion"), vbNewLine, drDestino("IDArticulo"))
                                End If
                            End If
                        End If
                    Next
                    If dblPorcentaje <> 100 Then
                        ApplicationService.GenerateError("El Tipo de Operación '|' tiene Reparto de Costes.|Compruebe que el Tipo y la Familia de los Artículos en Destino se corresponde con el establecido en el Tipo de Operación.", data.HeaderRow("IDTipoOperacion"), vbNewLine)
                    End If
                End If
            End If
        End If
    End Sub

#End Region

#Region "Validaciones - Materiales/Globales - Lotes/Globales"

    <Task()> Public Shared Sub ValidarOperacionVinoMateriales(ByVal data As DocumentoBdgOperacionReal, ByVal services As ServiceProvider)
        For Each DrMat As DataRow In data.dtOperacionVinoMaterial.Select
            ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf BdgVino.ValidarArticuloComponente, DrMat("IDArticulo"), services)
        Next
        For Each DrMat As DataRow In data.dtOperacionMaterial.Select
            ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf BdgVino.ValidarArticuloComponente, DrMat("IDArticulo"), services)
        Next
    End Sub

    <Task()> Public Shared Sub ValidarOperacionVinoMaterialesLotes(ByVal data As DocumentoBdgOperacionReal, ByVal services As ServiceProvider)
        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()

        Dim MaterialesConCantidad As List(Of DataRow) = (From c In data.dtOperacionVinoMaterial Where c.RowState <> DataRowState.Deleted AndAlso c.IsNull("IDOperacionMaterial") AndAlso ((Not c.IsNull("Cantidad") AndAlso c("Cantidad") <> 0) OrElse (Not c.IsNull("Merma") AndAlso c("Merma") <> 0))).ToList
        If MaterialesConCantidad Is Nothing OrElse MaterialesConCantidad.Count = 0 Then Exit Sub

        For Each DrMat As DataRow In MaterialesConCantidad
            Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(DrMat("IDArticulo"))
            If Not ArtInfo Is Nothing AndAlso ArtInfo.GestionStockPorLotes Then
                If Not data.dtOperacionVinoMaterialLote Is Nothing AndAlso data.dtOperacionVinoMaterialLote.Rows.Count > 0 Then
                    Dim LotesOperacionMaterial As List(Of DataRow) = (From c In data.dtOperacionVinoMaterialLote Where c.RowState <> DataRowState.Deleted AndAlso Not c.IsNull("IDVinoMaterial") AndAlso c("IDVinoMaterial") = DrMat("IDVinoMaterial")).ToList
                    If Not LotesOperacionMaterial Is Nothing AndAlso LotesOperacionMaterial.Count > 0 Then
                        Dim QLotesAsignada As Double = (Aggregate c In LotesOperacionMaterial Into Sum(CDbl(c("Cantidad"))))
                        If QLotesAsignada <> (Nz(DrMat("Cantidad"), 0) + Nz(DrMat("Merma"), 0)) Then
                            ApplicationService.GenerateError("El desglose de lotes del Artículo | no es correcto. Revise sus datos.", DrMat("IDArticulo"))
                        End If

                        If ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf Business.Negocio.Articulo.EsPrecinta, DrMat("IDArticulo"), services) Then
                            Dim LotesNuevos As List(Of DataRow) = (From c In LotesOperacionMaterial Where c.RowState = DataRowState.Added Select c).ToList
                            If Not LotesNuevos Is Nothing AndAlso LotesNuevos.Count > 0 Then
                                For Each DrLote As DataRow In LotesNuevos
                                    Dim FilPrecinta As New Filter
                                    FilPrecinta.Add("IDArticulo", DrMat("IDArticulo"))
                                    FilPrecinta.Add("IDAlmacen", DrMat("IDAlmacen"))
                                    FilPrecinta.Add("Lote", DrLote("Lote"))
                                    Dim StPrecinta As New BdgHistoricoPrecintas.stObtenerRangosPrecinta(DrMat("IDArticulo"), DrMat("IDAlmacen"), Nz(DrLote("SeriePrecinta"), String.Empty), String.Empty, String.Empty, DrLote("Lote"), DrLote("Ubicacion"), String.Empty, String.Empty, enumBoolean.Todos)
                                    Dim dttRangosLibres As DataTable = ProcessServer.ExecuteTask(Of BdgHistoricoPrecintas.stObtenerRangosPrecinta, DataTable)(AddressOf BdgHistoricoPrecintas.ObtenerRangosLibresPrecinta, StPrecinta, services)
                                    If (dttRangosLibres Is Nothing OrElse dttRangosLibres.Rows.Count = 0) Then
                                        ApplicationService.GenerateError(String.Format("No hay rangos libres para el artículo |", DrMat("IDArticulo")))
                                    End If

                                    If Nz(DrLote("NDesde")) <> 0 AndAlso Nz(DrLote("NHasta")) <> 0 Then
                                        Dim BlnRangoLibre As Boolean = False
                                        For Each dtrLibre As DataRow In dttRangosLibres.Rows
                                            If (DrLote("NDesde") >= dtrLibre("NDesdePrecinta") AndAlso DrLote("NDesde") <= dtrLibre("NHastaPrecinta")) Then
                                                'si el desde está contenido en un rango, comprobaos que el hasta también lo está paa validarlo y devolver que se ha encontrado hueco
                                                If (DrLote("NHasta") >= dtrLibre("NDesdePrecinta") AndAlso DrLote("NHasta") <= dtrLibre("NHastaPrecinta")) Then
                                                    BlnRangoLibre = True 'encontrado rango libre
                                                End If
                                            End If
                                        Next
                                        If Not BlnRangoLibre Then ApplicationService.GenerateError("No hay rangos libres para el artículo |", DrMat("IDArticulo"))
                                    Else
                                        ApplicationService.GenerateError("Debe indicar el rango de precintas que desea asignar para el artículo {0}.", Quoted(DrMat("IDArticulo")))
                                    End If
                                Next
                            End If
                        End If
                    Else
                        ApplicationService.GenerateError("Debe establecer el desglose de lotes de los materiales. Artículo: |", DrMat("IDArticulo"))
                    End If
                Else
                    ApplicationService.GenerateError("Debe establecer el desglose de lotes de los materiales. Artículo: |", DrMat("IDArticulo"))
                End If
            End If
        Next
    End Sub

    <Task()> Public Shared Sub ValidarOperacionMaterialesGlobalesLotes(ByVal data As DocumentoBdgOperacionReal, ByVal services As ServiceProvider)
        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()

        Dim MaterialesConCantidad As List(Of DataRow) = (From c In data.dtOperacionMaterial Where (Not c.IsNull("Cantidad") AndAlso c("Cantidad") <> 0) OrElse (Not c.IsNull("Merma") AndAlso c("Merma") <> 0)).ToList
        If MaterialesConCantidad Is Nothing OrElse MaterialesConCantidad.Count = 0 Then Exit Sub

        For Each DrMat As DataRow In MaterialesConCantidad
            Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(DrMat("IDArticulo"))
            If Not ArtInfo Is Nothing AndAlso ArtInfo.GestionStockPorLotes Then
                If Not data.dtOperacionMaterialLote Is Nothing AndAlso data.dtOperacionMaterialLote.Rows.Count > 0 Then

                    Dim LotesOperacionMaterial As List(Of DataRow) = (From c In data.dtOperacionMaterialLote Where Not c.IsNull("IDOperacionMaterial") AndAlso c("IDOperacionMaterial") = DrMat("IDOperacionMaterial")).ToList
                    'Dim f As New Filter
                    'f.Add(New GuidFilterItem("IDOperacionMaterial", DrMat("IDOperacionMaterial")))
                    'Dim dtrLotesGlobales() As DataRow = data.dtOperacionMaterialLote.Select(f.Compose(New AdoFilterComposer), String.Empty)
                    If Not LotesOperacionMaterial Is Nothing AndAlso LotesOperacionMaterial.Count > 0 Then

                        Dim QLotesAsignada As Double = (Aggregate c In LotesOperacionMaterial Into Sum(CDbl(c("Cantidad"))))
                        'Dim dblTotal As Double = 0
                        'For Each dtrLotesGlobal As DataRow In dtrLotesGlobales
                        '    dblTotal += dtrLotesGlobal("Cantidad")
                        'Next
                        If QLotesAsignada <> (Nz(DrMat("Cantidad"), 0) + Nz(DrMat("Merma"), 0)) Then
                            'ApplicationService.GenerateError("Debe establecer el desglose de lotes de los materiales globales. Artículo: |", DrMat("IDArticulo"))
                            ApplicationService.GenerateError("El desglose de lotes del Artículo | en imputaciones globales no es correcto. Revise sus datos.", DrMat("IDArticulo"))
                        End If

                        If ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf Business.Negocio.Articulo.EsPrecinta, DrMat("IDArticulo"), services) Then
                            Dim LotesNuevos As List(Of DataRow) = (From c In LotesOperacionMaterial Where c.RowState = DataRowState.Added Select c).ToList
                            If Not LotesNuevos Is Nothing AndAlso LotesNuevos.Count > 0 Then
                                For Each DrLote As DataRow In LotesNuevos
                                    Dim FilPrecinta As New Filter
                                    FilPrecinta.Add("IDArticulo", DrMat("IDArticulo"))
                                    FilPrecinta.Add("IDAlmacen", DrMat("IDAlmacen"))
                                    FilPrecinta.Add("Lote", DrLote("Lote"))
                                    Dim StPrecinta As New BdgHistoricoPrecintas.stObtenerRangosPrecinta(DrMat("IDArticulo"), DrMat("IDAlmacen"), Nz(DrLote("SeriePrecinta"), String.Empty), String.Empty, String.Empty, DrLote("Lote"), DrLote("Ubicacion"), String.Empty, String.Empty, enumBoolean.Todos)
                                    Dim dttRangosLibres As DataTable = ProcessServer.ExecuteTask(Of BdgHistoricoPrecintas.stObtenerRangosPrecinta, DataTable)(AddressOf BdgHistoricoPrecintas.ObtenerRangosLibresPrecinta, StPrecinta, services)
                                    If (dttRangosLibres Is Nothing OrElse dttRangosLibres.Rows.Count = 0) Then
                                        ApplicationService.GenerateError(String.Format("No hay rangos libres para el artículo |", DrMat("IDArticulo")))
                                    End If

                                    Dim BlnRangoLibre As Boolean = False
                                    For Each dtrLibre As DataRow In dttRangosLibres.Rows
                                        If (DrLote("NDesde") >= dtrLibre("NDesdePrecinta") AndAlso DrLote("NDesde") <= dtrLibre("NHastaPrecinta")) Then
                                            'si el desde está contenido en un rango, comprobaos que el hasta también lo está paa validarlo y devolver que se ha encontrado hueco
                                            If (DrLote("NHasta") >= dtrLibre("NDesdePrecinta") AndAlso DrLote("NHasta") <= dtrLibre("NHastaPrecinta")) Then
                                                BlnRangoLibre = True 'encontrado rango libre
                                            End If
                                        End If
                                    Next
                                    If Not BlnRangoLibre Then ApplicationService.GenerateError("No hay rangos libres para el artículo |.", DrMat("IDArticulo"))
                                Next
                            End If
                        End If
                    Else : ApplicationService.GenerateError("Debe establecer el desglose de lotes de los materiales globales. Artículo: |", DrMat("IDArticulo"))
                    End If
                Else : ApplicationService.GenerateError("Debe establecer el desglose de lotes de los materiales globales. Artículo: |", DrMat("IDArticulo"))
                End If
                'Else : ApplicationService.GenerateError("Debe establecer el desglose de lotes de los materiales globales. Artículo: |", DrMat("IDArticulo"))
            End If
        Next
    End Sub

#End Region

#End Region

    <Task()> Public Shared Sub AsignarNumeroOperacion(ByVal data As DocumentoBdgOperacionReal, ByVal services As ServiceProvider)
        If data.HeaderRow.RowState = DataRowState.Added Then
            If Length(data.HeaderRow("IDContador")) > 0 Then

                Dim NOperacionProvis As String = data.HeaderRow("NOperacion") & String.Empty

                Dim StDatos As New Contador.DatosCounterValue
                StDatos.IDCounter = data.HeaderRow("IDContador")
                StDatos.TargetClass = New BdgOperacion
                StDatos.TargetField = "NOperacion"
                StDatos.DateField = "Fecha"
                StDatos.DateValue = data.HeaderRow("Fecha")
                'StDatos.IDEjercicio = fra.HeaderRow("IDEjercicio") & String.Empty
                data.HeaderRow("NOperacion") = ProcessServer.ExecuteTask(Of Contador.DatosCounterValue, String)(AddressOf Contador.CounterValue, StDatos, services)


                Dim datActualizar As New DataActualizarNOperacionRestoDocumento(data, NOperacionProvis, data.HeaderRow("NOperacion"))
                ProcessServer.ExecuteTask(Of DataActualizarNOperacionRestoDocumento)(AddressOf ActualizarNOperacionRestoDocumento, datActualizar, services)

            End If
        End If
    End Sub



    Public Class DataActualizarNOperacionRestoDocumento
        Public Doc As DocumentoBdgOperacionReal
        Public NOperacionProvisional As String
        Public NOperacionDefinitiva As String

        Public Sub New(ByVal Doc As DocumentoBdgOperacionReal, ByVal NOperacionProvisional As String, ByVal NOperacionDefinitiva As String)
            Me.Doc = Doc
            Me.NOperacionProvisional = NOperacionProvisional
            Me.NOperacionDefinitiva = NOperacionDefinitiva
        End Sub
    End Class
    <Task()> Public Shared Sub ActualizarNOperacionRestoDocumento(ByVal data As DataActualizarNOperacionRestoDocumento, ByVal services As ServiceProvider)
        '//Actualizamos el NOperacionPlan en el resto de entidades, si se ha cambiado el contador propuesto por un contador fijo
        If data.NOperacionProvisional <> data.NOperacionDefinitiva Then

            Dim datActualizar As New BdgGeneral.DataActualizarPKCabecera(data.Doc.dtOperacionVino, data.NOperacionProvisional, data.NOperacionDefinitiva, "NOperacion")
            ProcessServer.ExecuteTask(Of BdgGeneral.DataActualizarPKCabecera, DataTable)(AddressOf BdgGeneral.ActualizarPKCabecera, datActualizar, services)

            '//
            datActualizar.dtActualizar = data.Doc.dtOperacionMaterial
            ProcessServer.ExecuteTask(Of BdgGeneral.DataActualizarPKCabecera, DataTable)(AddressOf BdgGeneral.ActualizarPKCabecera, datActualizar, services)

            datActualizar.dtActualizar = data.Doc.dtOperacionMOD
            ProcessServer.ExecuteTask(Of BdgGeneral.DataActualizarPKCabecera, DataTable)(AddressOf BdgGeneral.ActualizarPKCabecera, datActualizar, services)

            datActualizar.dtActualizar = data.Doc.dtOperacionCentro
            ProcessServer.ExecuteTask(Of BdgGeneral.DataActualizarPKCabecera, DataTable)(AddressOf BdgGeneral.ActualizarPKCabecera, datActualizar, services)

            datActualizar.dtActualizar = data.Doc.dtOperacionVarios
            ProcessServer.ExecuteTask(Of BdgGeneral.DataActualizarPKCabecera, DataTable)(AddressOf BdgGeneral.ActualizarPKCabecera, datActualizar, services)

            '//
            datActualizar.dtActualizar = data.Doc.dtOperacionVinoMaterial
            ProcessServer.ExecuteTask(Of BdgGeneral.DataActualizarPKCabecera, DataTable)(AddressOf BdgGeneral.ActualizarPKCabecera, datActualizar, services)

            datActualizar.dtActualizar = data.Doc.dtOperacionVinoMOD
            ProcessServer.ExecuteTask(Of BdgGeneral.DataActualizarPKCabecera, DataTable)(AddressOf BdgGeneral.ActualizarPKCabecera, datActualizar, services)

            datActualizar.dtActualizar = data.Doc.dtOperacionVinoCentro
            ProcessServer.ExecuteTask(Of BdgGeneral.DataActualizarPKCabecera, DataTable)(AddressOf BdgGeneral.ActualizarPKCabecera, datActualizar, services)

            datActualizar.dtActualizar = data.Doc.dtOperacionVinoVarios
            ProcessServer.ExecuteTask(Of BdgGeneral.DataActualizarPKCabecera, DataTable)(AddressOf BdgGeneral.ActualizarPKCabecera, datActualizar, services)

            datActualizar.dtActualizar = data.Doc.dtOperacionVinoAnalisis
            ProcessServer.ExecuteTask(Of BdgGeneral.DataActualizarPKCabecera, DataTable)(AddressOf BdgGeneral.ActualizarPKCabecera, datActualizar, services)
        End If
    End Sub


    <Task()> Public Shared Sub ActualizarOperacionPlanRelacionada(ByVal data As DocumentoBdgOperacionReal, ByVal services As ServiceProvider)
        If data.HeaderRow.RowState = DataRowState.Added Then
            If (data.HeaderRow.Table.Columns.Contains("NOperacionPlan") AndAlso Length(data.HeaderRow("NOperacionPlan")) > 0) Then
                Dim stData As New BdgOperacionPlan.StCambiarEstadoOperacion(data.HeaderRow("NOperacionPlan"), BdgEstadoOperacionPlan.Confirmado)
                ProcessServer.ExecuteTask(Of BdgOperacionPlan.StCambiarEstadoOperacion)(AddressOf BdgOperacionPlan.CambiarEstadoOperacion, stData, services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub EjecutarMovimientos(ByVal data As DocumentoBdgOperacionReal, ByVal services As ServiceProvider)
        ' AdminData.CommitTx(True)
        'Actualizar de nuevo el documento con el idmovimiento asociado, verificar si con este comando nos valdrá.
        Dim FechaDoc As New Date(CDate(data.HeaderRow("Fecha")).Year, CDate(data.HeaderRow("Fecha")).Month, CDate(data.HeaderRow("Fecha")).Day)
        If data.HeaderRow.RowState = DataRowState.Added Then
            data = New DocumentoBdgOperacionReal(data.HeaderRow("NOperacion"))
            Dim StEj As New BdgWorkClass.StEjecutarMovimientos(data.HeaderRow("NOperacion"), FechaDoc)
            data.HeaderRow("IDMovimiento") = ProcessServer.ExecuteTask(Of BdgWorkClass.StEjecutarMovimientos, Integer)(AddressOf BdgWorkClass.EjecutarMovimientos, StEj, services)
            data.SetData()
        Else
            If Length(data.HeaderRow("IDMovimiento")) = 0 Then
                data = New DocumentoBdgOperacionReal(data.HeaderRow("NOperacion"))
                Dim StEjec As New BdgWorkClass.StEjecutarMovimientos(data.HeaderRow("NOperacion"), FechaDoc)
                data.HeaderRow("IDMovimiento") = ProcessServer.ExecuteTask(Of BdgWorkClass.StEjecutarMovimientos, Integer)(AddressOf BdgWorkClass.EjecutarMovimientos, StEjec, services)
                data.SetData()
            Else
                data = New DocumentoBdgOperacionReal(data.HeaderRow("NOperacion"))
                Dim StEj As New BdgWorkClass.StEjecutarMovimientosNumero(data.HeaderRow("IDMovimiento"), data.HeaderRow("NOperacion"), FechaDoc)
                ProcessServer.ExecuteTask(Of BdgWorkClass.StEjecutarMovimientosNumero)(AddressOf BdgWorkClass.EjecutarMovimientosNumero, StEj, services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioFechaOperacion(ByVal data As DocumentoBdgOperacionReal, ByVal services As ServiceProvider)
        If data.HeaderRow.RowState = DataRowState.Modified AndAlso (data.HeaderRow("Fecha") <> data.HeaderRow("Fecha", DataRowVersion.Original)) Then
            ProcessServer.ExecuteTask(Of DocumentoBdgOperacionReal)(AddressOf ProcesoBdgOperacion.ActualizarFechaDocumentoEnImputaciones, data, services)
            'Hay que comprobar que el tipo de operacion tenga puesto que controle los movimientos

            'Cambiamos la fecha a los movimientos asociados a Origen y Destino si la operación genera mvtos.
            If data.GetTipoOperTipoMov <> Business.Bodega.TipoMovimiento.SinMovimiento Then
                Dim StModif As New ProcesoBdgOperacion.StModificarFechaOperacion(data.HeaderRow("NOperacion"), data.HeaderRow("Fecha"), data.HeaderRow("IDMovimiento"), data.HeaderRow("IDTipoOperacion"))
                ProcessServer.ExecuteTask(Of ProcesoBdgOperacion.StModificarFechaOperacion)(AddressOf ProcesoBdgOperacion.ModificarFechaOperacion, StModif, services)
            End If

            'Cambiamos la fecha a los movimientos asociados a los materiales.
            Dim VinoMatLote As New BdgVinoMaterialLote
            If Not data.dtOperacionVinoMaterial Is Nothing AndAlso data.dtOperacionVinoMaterial.Rows.Count > 0 Then
                Dim FechaDoc As New Date(CDate(data.HeaderRow("Fecha")).Year, CDate(data.HeaderRow("Fecha")).Month, CDate(data.HeaderRow("Fecha")).Day)
                For Each drVinoMaterial As DataRow In data.dtOperacionVinoMaterial.Rows
                    If Length(drVinoMaterial("IDLineaMovimiento")) > 0 Then
                        Dim dataCorreccion As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Corregir, drVinoMaterial("IDLineaMovimiento"), FechaDoc, False)
                        ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento, StockUpdateData)(AddressOf ProcesoStocks.ActualizarMovimiento, dataCorreccion, services)
                    Else
                        If Not data.dtOperacionVinoMaterialLote Is Nothing AndAlso data.dtOperacionVinoMaterialLote.Rows.Count > 0 Then
                            Dim LotesMaterial As List(Of DataRow) = (From c In data.dtOperacionVinoMaterialLote Where Not c.IsNull("IDVinoMaterial") AndAlso c("IDVinoMaterial") = drVinoMaterial("IDVinoMaterial") Select c).ToList
                            If Not LotesMaterial Is Nothing AndAlso LotesMaterial.Count > 0 Then
                                For Each drVinoMaterialLote As DataRow In LotesMaterial
                                    If Length(drVinoMaterialLote("IDLineaMovimiento")) > 0 Then
                                        Dim dataCorreccion As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Corregir, drVinoMaterialLote("IDLineaMovimiento"), FechaDoc, False)
                                        ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento, StockUpdateData)(AddressOf ProcesoStocks.ActualizarMovimiento, dataCorreccion, services)
                                    End If
                                Next
                            End If
                        End If
                    End If
                Next
            End If
        End If
    End Sub

    <Task()> Public Shared Sub DeleteImputacionesVino(ByVal doc As DocumentoBdgOperacionReal, ByVal services As ServiceProvider)
        '//NOTA: debemos eliminar registros de entidades biznietas antes de guardar, registros que se han eliminado en el proceso del Update, después de crear el documento.
        Dim RegsEliminados As List(Of DataRow) = (From c In doc.dtOperacionVinoMaterialLote Where c.RowState = DataRowState.Deleted Select c).ToList()
        If RegsEliminados.Count > 0 Then
            Dim dtDelete As DataTable = doc.dtOperacionVinoMaterialLote.Clone
            For Each dr As DataRow In RegsEliminados
                dtDelete.ImportRow(dr)
                dr.AcceptChanges()
            Next
            dtDelete.RejectChanges()
            Dim OVML As New BdgVinoMaterialLote
            OVML.Delete(dtDelete)
        End If

        '//REVISAR: ¿hay que eliminar centro tasa y analisis variable?

    End Sub

    <Serializable()> _
    Public Class StModificarFechaOperacion
        Public NOperacion As String
        Public Fecha As Date
        Public IDMovimiento As Integer
        Public IDTipoOperacion As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal NOperacion As String, ByVal Fecha As Date, ByVal IDMovimiento As Integer, ByVal IDTipoOperacion As String)
            Me.NOperacion = NOperacion
            Me.Fecha = Fecha
            Me.IDMovimiento = IDMovimiento
            Me.IDTipoOperacion = IDTipoOperacion
        End Sub
    End Class

    <Task()> Public Shared Sub ModificarFechaOperacion(ByVal data As StModificarFechaOperacion, ByVal services As ServiceProvider)
        Dim dtOpHis As DataTable = AdminData.GetData("tbHistoricoMovimiento", New NumberFilterItem("IDMovimiento", data.IDMovimiento))
        For Each rwHM As DataRow In dtOpHis.Rows
            If Length(rwHM("IDLineaMovimiento")) > 0 Then
                Dim Fecha As New Date(data.Fecha.Year, data.Fecha.Month, data.Fecha.Day)
                Dim dataCorreccion As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Corregir, rwHM("IDLineaMovimiento"), Fecha, False)
                ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento)(AddressOf ProcesoStocks.ActualizarMovimiento, dataCorreccion, services)
            End If
        Next

        Dim TiposOperaciones As EntityInfoCache(Of BdgTipoOperacionInfo) = services.GetService(Of EntityInfoCache(Of BdgTipoOperacionInfo))()
        Dim TipoOp As BdgTipoOperacionInfo = TiposOperaciones.GetEntity(data.IDTipoOperacion)
        If TipoOp.TipoMovimiento <> Business.Bodega.TipoMovimiento.SinMovimiento _
            AndAlso TipoOp.TipoMovimiento <> Business.Bodega.TipoMovimiento.CrearOrigen Then
           
            Dim OpVino As New BdgOperacionVino
            Dim v As New BdgVino
            Dim f As New Filter
            f.Add(New StringFilterItem(_OV.NOperacion, data.NOperacion))
            f.Add(New BooleanFilterItem(_OV.Destino, True))
            Dim dtOVaux As DataTable = OpVino.Filter(f)
            For Each drDestino As DataRow In dtOVaux.Rows
                '//Modificar la fecha en el vino
                Dim dtVino As DataTable = v.SelOnPrimaryKey(drDestino(_OV.IDVino))
                If dtVino.Rows.Count > 0 Then
                    dtVino.Rows(0)(_V.Fecha) = data.Fecha
                    v.Validate(dtVino)
                    v.Update(dtVino)
                End If
            Next
        End If
    End Sub

#Region " Gestion con OFS"

    <Task()> Public Shared Sub ActualizarOF(ByVal data As DocumentoBdgOperacionReal, ByVal services As ServiceProvider)
        Dim ClsOF As BusinessHelper = BusinessHelper.CreateBusinessObject("OrdenFabricacion")
        For Each DrDest As DataRow In data.GetOperacionVinoDestino
            Dim IDOrden As Integer = 0
            Dim QOld As Double = 0
            Dim QNew As Double = 0
            If DrDest.HasVersion(DataRowVersion.Current) Then
                QNew = DrDest("Cantidad")
                If Not DrDest.IsNull("IDOrden") Then IDOrden = DrDest("IDOrden")
            End If
            If DrDest.HasVersion(DataRowVersion.Original) Then
                QOld = DrDest("Cantidad", DataRowVersion.Original)
                If IDOrden = 0 Then
                    If Not IsDBNull(DrDest("IDOrden", DataRowVersion.Original)) Then IDOrden = DrDest("IDOrden", DataRowVersion.Original)
                End If
            End If

            If IDOrden <> 0 Then
                Dim dd As New dataActualizarEstadoOF(IDOrden, QNew, QOld)
                dd.FechaFinReal = data.HeaderRow("Fecha")
                ProcessServer.ExecuteTask(Of dataActualizarEstadoOF)(AddressOf ProcesoBdgOperacion.ActualizarEstadoOF, dd, services)
            End If
        Next
    End Sub
    <Serializable()> _
    Public Class dataActualizarEstadoOF
        Public IDOrden As Integer
        Public QFabricada As Double
        Public QFabricadaOLD As Double
        Public FechaFinReal As Date

        Public Sub New(ByVal IDOrden As Integer, ByVal QFabricada As Double, ByVal QFabricadaOLD As Double)
            Me.IDOrden = IDOrden
            Me.QFabricada = QFabricada
            Me.QFabricadaOLD = QFabricadaOLD
        End Sub
    End Class
    <Task()> Public Shared Sub ActualizarEstadoOF(ByVal data As dataActualizarEstadoOF, ByVal services As ServiceProvider)
        Dim ClsOF As BusinessHelper = BusinessHelper.CreateBusinessObject("OrdenFabricacion")

        Dim DrOF As DataRow = ClsOF.GetItemRow(data.IDOrden)
        DrOF("QFabricada") += data.QFabricada - data.QFabricadaOLD
        DrOF("QIniciada") += data.QFabricada - data.QFabricadaOLD
        If DrOF("Estado") = enumofEstado.ofePlanificada And DrOF("QIniciada") <> 0 Then
            DrOF("Estado") = enumofEstado.ofeIniciada
        Else
            DrOF("Estado") = enumofEstado.ofePlanificada
        End If
        If DrOF("QFabricada") >= DrOF("QFabricar") Then
            If New Parametro().CierreAutomaticoOF() Then
                DrOF("Estado") = enumofEstado.ofeTerminada
            End If
        End If
        If DrOF("Estado") = enumofEstado.ofeTerminada Then
            DrOF("FechaFinReal") = data.FechaFinReal
        End If
        ClsOF.Update(DrOF.Table)
    End Sub

    '<Serializable()> _
    'Public Class clsCabeceraOF
    '    Public intIDOrden As Long
    '    Public dtOrdenFabricacion As DataTable

    '    Public Sub New()
    '    End Sub

    '    Public Sub New(ByVal intIDOrden As Long)
    '        Me.intIDOrden = intIDOrden
    '    End Sub
    'End Class

    '<Task()> Public Shared Sub CabeceraOF(ByVal data As clsCabeceraOF, ByVal services As ServiceProvider)
    '    Dim f As New Filter
    '    f.Add(New NumberFilterItem("IDOrden", data.intIDOrden))
    '    Dim dtOrden As DataTable = New BE.DataEngine().Filter("NegBdgOrdenCabecera", f)
    '    If Not dtOrden Is Nothing AndAlso dtOrden.Rows.Count > 0 Then
    '        data.dtOrdenFabricacion = dtOrden.Copy
    '    End If
    'End Sub

    '<Serializable()> _
    'Public Class clsDatosOfParaOperacion
    '    Public intIDOrden As Long
    '    Public dblCantidad As Double
    '    Public dtOrigen As DataTable
    '    Public dtDestino As DataTable
    '    Public dtMateriales As DataTable
    '    Public dtCentros As DataTable

    '    Public Sub New()
    '    End Sub

    '    Public Sub New(ByVal intIDOrden As Long, ByVal dblCantidad As Double)
    '        Me.intIDOrden = intIDOrden
    '        Me.dblCantidad = dblCantidad
    '    End Sub
    'End Class

    '<Task()> Public Shared Function DatosOfParaOperacion(ByVal data As clsDatosOfParaOperacion, ByVal services As ServiceProvider) As clsDatosOfParaOperacion
    '    'Origen
    '    Dim dataOrigenOF As New BdgOperacion.DataMaterialesOF(data.intIDOrden, data.dblCantidad, True, data.dtOrigen)
    '    ProcessServer.ExecuteTask(Of BdgOperacion.DataMaterialesOF)(AddressOf BdgOperacion.MaterialesOF, dataOrigenOF, services)
    '    data.dtOrigen = dataOrigenOF.dtMateriales
    '    For Each dr As DataRow In data.dtOrigen.Rows
    '        If Length(dr("IDAlmacen")) > 0 AndAlso Length(dr("IDArticulo")) > 0 AndAlso Length(dr("Lote")) > 0 AndAlso Length(dr("Ubicacion")) > 0 Then
    '            Dim dataVQ As New ProcesoStocks.DataVinoQ(dr("IDArticulo"), dr("Ubicacion"), dr("Lote"), dr("IDAlmacen"))
    '            dataVQ = ProcessServer.ExecuteTask(Of ProcesoStocks.DataVinoQ, ProcesoStocks.DataVinoQ)(AddressOf BdgWorkClass.GetIDVinoCantidad, dataVQ, services)
    '            If Not dataVQ Is Nothing Then
    '                dr("IDVino") = dataVQ.IDVino
    '            End If
    '        End If
    '    Next

    '    'Destino
    '    Dim dataDestinoOF As New clsCabeceraOF(data.intIDOrden)
    '    ProcessServer.ExecuteTask(Of clsCabeceraOF)(AddressOf CabeceraOF, dataDestinoOF, services)
    '    data.dtDestino = dataDestinoOF.dtOrdenFabricacion
    '    data.dtDestino.Columns.Add("IDVino", GetType(Guid))
    '    For Each dr As DataRow In data.dtDestino.Rows
    '        If Length(dr("IDAlmacen")) > 0 AndAlso Length(dr("IDArticulo")) > 0 AndAlso Length(dr("Lote")) > 0 AndAlso Length(dr("IDUbicacion")) > 0 Then
    '            Dim dataVQ As New ProcesoStocks.DataVinoQ(dr("IDArticulo"), dr("IDUbicacion"), dr("Lote"), dr("IDAlmacen"))
    '            dataVQ = ProcessServer.ExecuteTask(Of ProcesoStocks.DataVinoQ, ProcesoStocks.DataVinoQ)(AddressOf BdgWorkClass.GetIDVinoCantidad, dataVQ, services)
    '            If Not dataVQ Is Nothing Then
    '                dr("IDVino") = dataVQ.IDVino
    '            End If
    '        End If
    '    Next
    '    Return data
    'End Function

#End Region


    <Task()> Public Shared Sub ActualizarFechaDocumentoEnImputaciones(ByVal data As DocumentoBdgOperacionReal, ByVal services As ServiceProvider)
        If data.HeaderRow.RowState = DataRowState.Modified AndAlso (data.HeaderRow("Fecha") <> data.HeaderRow("Fecha", DataRowVersion.Original)) Then

            Dim datActualizaFecha As New DataActualizarFechaDt(data.dtOperacionMaterial, data.HeaderRow("Fecha"))
            ProcessServer.ExecuteTask(Of DataActualizarFechaDt)(AddressOf ActualizarFechaDt, datActualizaFecha, services)
            datActualizaFecha.dt = data.dtOperacionMOD
            ProcessServer.ExecuteTask(Of DataActualizarFechaDt)(AddressOf ActualizarFechaDt, datActualizaFecha, services)
            datActualizaFecha.dt = data.dtOperacionCentro
            ProcessServer.ExecuteTask(Of DataActualizarFechaDt)(AddressOf ActualizarFechaDt, datActualizaFecha, services)
            datActualizaFecha.dt = data.dtOperacionVarios
            ProcessServer.ExecuteTask(Of DataActualizarFechaDt)(AddressOf ActualizarFechaDt, datActualizaFecha, services)

            datActualizaFecha.dt = data.dtOperacionVinoMaterial
            ProcessServer.ExecuteTask(Of DataActualizarFechaDt)(AddressOf ActualizarFechaDt, datActualizaFecha, services)
            datActualizaFecha.dt = data.dtOperacionVinoMOD
            ProcessServer.ExecuteTask(Of DataActualizarFechaDt)(AddressOf ActualizarFechaDt, datActualizaFecha, services)
            datActualizaFecha.dt = data.dtOperacionVinoCentro
            ProcessServer.ExecuteTask(Of DataActualizarFechaDt)(AddressOf ActualizarFechaDt, datActualizaFecha, services)
            datActualizaFecha.dt = data.dtOperacionVinoVarios
            ProcessServer.ExecuteTask(Of DataActualizarFechaDt)(AddressOf ActualizarFechaDt, datActualizaFecha, services)

        End If
    End Sub

    <Serializable()> _
    Public Class DataActualizarFechaDt
        Public dt As DataTable
        Public FechaNew As Date
        Public Sub New(ByVal dt As DataTable, ByVal FechaNew As Date)
            Me.dt = dt
            Me.FechaNew = FechaNew
        End Sub
    End Class
    <Task()> Public Shared Sub ActualizarFechaDt(ByVal data As DataActualizarFechaDt, ByVal services As ServiceProvider)
        If Not data.dt Is Nothing AndAlso data.dt.Rows.Count > 0 AndAlso data.dt.Columns.Contains("Fecha") AndAlso data.FechaNew <> cnMinDate Then
            For Each dr As DataRow In data.dt.Rows
                dr("Fecha") = data.FechaNew
            Next
        End If
    End Sub

#End Region

#Region "Tareas OperacionVino"

    <Task()> Public Shared Sub ActualizarEstadoVino(ByVal data As DocumentoBdgOperacionReal, ByVal services As ServiceProvider)
        Dim ActualizarEstadoVino As List(Of DataRow) = (From c In data.dtOperacionVino _
                                                        Where Not c.IsNull("IDEstadoVino") AndAlso _
                                                             (c.RowState = DataRowState.Added OrElse _
                                                             (c.RowState = DataRowState.Modified AndAlso c("IDEstadoVino") & String.Empty <> c("IDEstadoVino", DataRowVersion.Original) & String.Empty))).ToList()

        For Each DrVino As DataRow In ActualizarEstadoVino
            Dim StEstado As New BdgVino.StModificarEstado(DrVino("IDVino"), DrVino("IDEstadoVino"))
            ProcessServer.ExecuteTask(Of BdgVino.StModificarEstado)(AddressOf BdgVino.ModificarEstado, StEstado, services)
        Next
    End Sub

    <Task()> Public Shared Sub AsignarDatosOperacionOrigen(ByVal data As DocumentoBdgOperacionReal, ByVal services As ServiceProvider)
        If data.GetOperacionVinoOrigen Is Nothing OrElse data.GetOperacionVinoOrigen.Count = 0 Then Exit Sub

        Dim OrigenesModificados As List(Of DataRow) = (From c In data.GetOperacionVinoOrigen _
                                                       Where c.RowState = DataRowState.Added OrElse c.RowState = DataRowState.Modified).ToList()
        For Each DrOrigen As DataRow In OrigenesModificados
            If DrOrigen.RowState = DataRowState.Added Then
                Dim StCambio As New BdgWorkClass.StCambiarOcupacion(DrOrigen("IDVino"), -Nz(DrOrigen("Cantidad"), 0))
                ProcessServer.ExecuteTask(Of BdgWorkClass.StCambiarOcupacion)(AddressOf BdgWorkClass.CambiarOcupacion, StCambio, services)
            ElseIf DrOrigen.RowState = DataRowState.Modified Then
                Dim dblQAnt As Double = 0
                Dim dblQAct As Double = 0
                If DrOrigen.HasVersion(DataRowVersion.Original) Then dblQAnt = DrOrigen("Cantidad", DataRowVersion.Original)
                If DrOrigen.HasVersion(DataRowVersion.Current) Then dblQAct = Nz(DrOrigen("Cantidad"), 0)
                If CType(DrOrigen("IDVino"), Guid).Equals(DrOrigen("IDVino", DataRowVersion.Original)) Then
                    dblQAct = dblQAct - dblQAnt
                    dblQAnt = 0
                End If
                If dblQAnt <> 0 Then
                    Dim StCambio As New BdgWorkClass.StCambiarOcupacion(DrOrigen("IDVino", DataRowVersion.Original), dblQAnt)
                    ProcessServer.ExecuteTask(Of BdgWorkClass.StCambiarOcupacion)(AddressOf BdgWorkClass.CambiarOcupacion, StCambio, services)
                End If
                If dblQAct <> 0 Then
                    Dim StCambio As New BdgWorkClass.StCambiarOcupacion(DrOrigen("IDVino"), -dblQAct)
                    ProcessServer.ExecuteTask(Of BdgWorkClass.StCambiarOcupacion)(AddressOf BdgWorkClass.CambiarOcupacion, StCambio, services)
                End If
            End If
        Next
    End Sub

    <Task()> Public Shared Sub AsignarDatosOperacionDestino(ByVal data As DocumentoBdgOperacionReal, ByVal services As ServiceProvider)
        If data.GetOperacionVinoDestino Is Nothing OrElse data.GetOperacionVinoDestino.Count = 0 Then Exit Sub

        Dim DestinosTratar As List(Of DataRow) = (From c In data.GetOperacionVinoDestino _
                                                       Where c.RowState = DataRowState.Added OrElse c.RowState = DataRowState.Modified OrElse c.RowState = DataRowState.Unchanged).ToList()
        If data.GetTipoOperTipoMov <> Business.Bodega.TipoMovimiento.SinMovimiento Then
            Dim DataPrcBdgOper As FactoresRecalcularEstructura = services.GetService(Of FactoresRecalcularEstructura)()
            Dim StVinoFactor As Dictionary(Of Guid, Double) = DataPrcBdgOper.VinoFactor

            For Each DrDestino As DataRow In DestinosTratar
                If DrDestino.RowState = DataRowState.Added Then
                    Dim oVinos() As VinoComponente
                    If DataPrcBdgOper Is Nothing OrElse StVinoFactor Is Nothing OrElse StVinoFactor.Count = 0 Then
                        oVinos = Nothing
                    Else
                        Dim DtOr As DataTable
                        If Not data.GetOperacionVinoOrigen Is Nothing AndAlso data.GetOperacionVinoOrigen.Length > 0 Then
                            DtOr = data.GetOperacionVinoOrigen.CopyToDataTable
                        End If
                        Dim StCalc As New StCalcularEstructura(StVinoFactor(DrDestino("IDVino")), DrDestino("IDArticulo"), DrDestino("IDUdMedida"), DtOr)
                        oVinos = ProcessServer.ExecuteTask(Of StCalcularEstructura, VinoComponente())(AddressOf CalcularEstructura, StCalc, services)
                    End If
                    If data.GetTipoOperTipoMov = Business.Bodega.TipoMovimiento.CrearOrigen Then
                        Dim StCrear As New BdgWorkClass.StCrearEstructuraOperacion(DrDestino("IDVino"), BdgOrigenVino.Interno, data.HeaderRow("NOperacion"), oVinos)
                        ProcessServer.ExecuteTask(Of BdgWorkClass.StCrearEstructuraOperacion)(AddressOf BdgWorkClass.CrearEstructuraOperacion, StCrear, services)
                    Else
                        Dim oldID As Guid
                        If (Length(DrDestino("IDVino")) > 0) Then oldID = DrDestino("IDVino")
                        Dim IDEstadoVino As String = DrDestino("IDEstadoVino") & String.Empty
                        Dim StCrear As New BdgWorkClass.StCrearVino(DrDestino("IDDeposito"), DrDestino("IDArticulo"), DrDestino("Lote"), data.HeaderRow("Fecha"), BdgOrigenVino.Interno, DrDestino("IDUdMedida"), DrDestino("Cantidad"), IDEstadoVino, data.HeaderRow("NOperacion"), oVinos, , Nz(DrDestino("IDBarrica")))
                        DrDestino("IDVino") = ProcessServer.ExecuteTask(Of BdgWorkClass.StCrearVino, Guid)(AddressOf BdgWorkClass.CrearVino, StCrear, services)
                        ' Dim Tables() As DataTable = New DataTable() {data.dtOperacionVinoAnalisis, data.dtOperacionVinoMaterial, data.dtOperacionVinoCentro, data.dtOperacionVinoMOD, data.dtOperacionVinoVarios}
                        Dim Tables As New List(Of DataTable) '=  New DataTable() {data.Analisis, data.Materiales, data.Centros, data.Operarios, data.Varios}
                        Tables.AddRange(New DataTable() {data.dtOperacionVinoAnalisis, data.dtOperacionVinoMaterial, data.dtOperacionVinoCentro, data.dtOperacionVinoMOD, data.dtOperacionVinoVarios})
                        Dim CostesIDVino As ProcesoBdgOperacion.DataCostesIDVino = services.GetService(Of ProcesoBdgOperacion.DataCostesIDVino)()
                        If Not CostesIDVino Is Nothing Then
                            If Not CostesIDVino.CostesElaboracion Is Nothing AndAlso CostesIDVino.CostesElaboracion.Rows.Count > 0 Then
                                Tables.Add(CostesIDVino.CostesElaboracion)
                            End If
                            If Not CostesIDVino.CostesEstanciaNave Is Nothing AndAlso CostesIDVino.CostesEstanciaNave.Rows.Count > 0 Then
                                Tables.Add(CostesIDVino.CostesEstanciaNave)
                            End If
                        End If
                        Dim AnalisisIDVino As ProcesoBdgOperacion.DataAnalisisIDVino = services.GetService(Of ProcesoBdgOperacion.DataAnalisisIDVino)()
                        If Not AnalisisIDVino Is Nothing Then
                            If Not AnalisisIDVino.Analisis Is Nothing AndAlso AnalisisIDVino.Analisis.Rows.Count > 0 Then
                                Tables.Add(AnalisisIDVino.Analisis)
                            End If
                        End If

                        Dim StCambiar As New DataCambiarIDVino(oldID, DrDestino("IDVino"), Tables)
                        ProcessServer.ExecuteTask(Of DataCambiarIDVino)(AddressOf CambiarIDVino, StCambiar, services)
                    End If
                ElseIf DrDestino.RowState = DataRowState.Modified Then
                    If Not DataPrcBdgOper Is Nothing AndAlso Not StVinoFactor Is Nothing AndAlso StVinoFactor.Count > 0 Then
                        Dim dtOrigenes As DataTable = data.dtOperacionVino.Clone
                        Dim dtOrigenesDelete As DataTable = data.dtOperacionVino.Clone
                        If Not data.GetOperacionVinoOrigen Is Nothing AndAlso data.GetOperacionVinoOrigen.Length > 0 Then
                            Dim OrigenesEliminados As List(Of DataRow) = (From c In data.GetOperacionVinoOrigen _
                                                                            Where c.RowState = DataRowState.Deleted).ToList()
                            For Each dr As DataRow In OrigenesEliminados
                                dtOrigenesDelete.ImportRow(dr)
                            Next

                            Dim OrigenesNoEliminados As List(Of DataRow) = (From c In data.GetOperacionVinoOrigen _
                                                                            Where c.RowState = DataRowState.Added OrElse _
                                                                                  c.RowState = DataRowState.Modified OrElse _
                                                                                  c.RowState = DataRowState.Unchanged).ToList()
                            For Each dr As DataRow In OrigenesNoEliminados
                                dtOrigenes.ImportRow(dr)
                            Next
                        End If
                        Dim StRec As New StReconstruirEstructura(StVinoFactor(DrDestino("IDVino")), DrDestino("IDVino"), DrDestino("IDArticulo"), DrDestino("IDUdMedida"), _
                                                                              dtOrigenes, dtOrigenesDelete, New BdgWorkClass, BdgOrigenVino.Interno, data.HeaderRow("NOperacion"))
                        ProcessServer.ExecuteTask(Of StReconstruirEstructura)(AddressOf ReconstruirEstructura, StRec, services)
                    End If

                    Dim dblQN As Double = DrDestino("Cantidad")
                    Dim dblQ As Double = DrDestino("Cantidad", DataRowVersion.Original)
                    If dblQN - dblQ <> 0 Then
                        Dim StInc As New BdgWorkClass.StIncrementarIDVino(DrDestino("IDVino"), dblQN - dblQ)
                        ProcessServer.ExecuteTask(Of BdgWorkClass.StIncrementarIDVino)(AddressOf BdgWorkClass.IncrementarCantidadIDVino, StInc, services)
                    End If

                    '//diferentes lineas de OperacionVino pueden tener como destino el mismo Vino, y eb cada linea se puede asignar un IDBarrica diferente
                    Dim IDBarricaOld As String = Nz(DrDestino("IDBarrica", DataRowVersion.Original), String.Empty)
                    If Length(IDBarricaOld) > 0 AndAlso IDBarricaOld <> Nz(DrDestino("IDBarrica"), String.Empty) Then
                        Dim ClsBdgVIno As New BdgVino
                        Dim drV As DataRow = ClsBdgVIno.GetItemRow(DrDestino("IDVino"))
                        drV("IDBarrica") = DrDestino("IDBarrica")
                        ClsBdgVIno.Update(drV.Table)
                    End If
                ElseIf DrDestino.RowState = DataRowState.Unchanged Then
                    If Not data.GetOperacionVinoOrigen Is Nothing AndAlso data.GetOperacionVinoOrigen.Length > 0 Then
                        Dim dtOrigenesDelete As DataTable = data.dtOperacionVino.Clone
                        Dim dtOrigenes As DataTable = data.dtOperacionVino.Clone

                        Dim OrigenesEliminados As List(Of DataRow) = (From c In data.GetOperacionVinoOrigen _
                                                                           Where c.RowState = DataRowState.Deleted).ToList()
                        For Each dr As DataRow In OrigenesEliminados
                            dtOrigenesDelete.ImportRow(dr)
                        Next

                        Dim OrigenesModificados As List(Of DataRow) = (From c In data.GetOperacionVinoOrigen _
                                                                       Where c.RowState = DataRowState.Added OrElse _
                                                                             c.RowState = DataRowState.Modified).ToList()
                        For Each dr As DataRow In OrigenesModificados
                            dtOrigenes.ImportRow(dr)
                        Next

                        If Not dtOrigenes Is Nothing AndAlso dtOrigenes.Rows.Count > 0 AndAlso Not dtOrigenesDelete Is Nothing AndAlso dtOrigenesDelete.Rows.Count > 0 Then
                            If Not DataPrcBdgOper Is Nothing AndAlso Not StVinoFactor Is Nothing AndAlso StVinoFactor.Count > 0 Then
                                Dim StRec As New StReconstruirEstructura(StVinoFactor(DrDestino("IDVino")), DrDestino("IDVino"), DrDestino("IDArticulo"), DrDestino("IDUdMedida"), _
                                                                                      dtOrigenes, dtOrigenesDelete, New BdgWorkClass, _
                                                                                      BdgOrigenVino.Interno, data.HeaderRow("NOperacion"))
                                ProcessServer.ExecuteTask(Of StReconstruirEstructura)(AddressOf ReconstruirEstructura, StRec, services)
                            End If
                        End If
                    End If
                End If
            Next
        End If
    End Sub

    <Serializable()> _
    Public Class DataCostesIDVino
        Public CostesElaboracion As DataTable       '//BdgVinoMaterial
        Public CostesVendimia As DataTable          '//BdgCosteVendimiaHist (histórico de BdgVinoMaterial)

        Public CostesEstanciaNave As DataTable      '//BdgVinoCentro
        Public CostesEstanciaNaveTasas As DataTable '//BdgVinoCentroTasa


        Public CostesElaboracionDel As DataTable    '// BdgCosteVendimia se elimina en cascada desde BdgVinoMaterial
        Public CostesEstanciaNaveDel As DataTable   '// BdgVinoCentroTasa se elimina en cascada desde BdgVinoCentro
    End Class

    <Serializable()> _
    Public Class DataAnalisisIDVino
        Public Analisis As DataTable
        Public AnalisisVariable As DataTable

        Public AnalisisDel As DataTable         '//Las variables se eliminan por borrado en cascada
    End Class

    <Serializable()> _
    Public Class StReconstruirEstructura
        Public Factor As Double
        Public IDVIno As Guid
        Public IDArticulo As String
        Public IDUDMedida As String
        Public CurOrigen As DataTable
        Public DeslOrigen As DataTable
        Public oVWC As BdgWorkClass
        Public Origen As BdgOrigenVino
        Public NOperacion As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal Factor As Double, ByVal IDVino As Guid, ByVal IDArticulo As String, ByVal IDUDMedida As String, _
                       ByVal CurOrigen As DataTable, ByVal DeslOrigen As DataTable, ByVal oVWC As BdgWorkClass, ByVal Origen As BdgOrigenVino, _
                       ByVal NOperacion As String)
            Me.Factor = Factor
            Me.IDVIno = IDVino
            Me.IDArticulo = IDArticulo
            Me.IDUDMedida = IDUDMedida
            Me.CurOrigen = CurOrigen
            Me.DeslOrigen = DeslOrigen
            Me.oVWC = oVWC
            Me.Origen = Origen
            Me.NOperacion = NOperacion
        End Sub
    End Class

    <Task()> Public Shared Sub ReconstruirEstructura(ByVal data As StReconstruirEstructura, ByVal services As ServiceProvider)
        Dim StCalc As New StCalcularEstructura(data.Factor, data.IDArticulo, data.IDUDMedida, data.CurOrigen)
        Dim oVinos() As VinoComponente = ProcessServer.ExecuteTask(Of StCalcularEstructura, VinoComponente())(AddressOf CalcularEstructura, StCalc, services)

        Dim IDVinosOrigenOLD As List(Of Guid)
        If Not data.DeslOrigen Is Nothing Then
            IDVinosOrigenOLD = (From c In data.DeslOrigen Select CType(c("IDVino", DataRowVersion.Original), Guid) Distinct).ToList
        End If
        Dim StCrear As New BdgWorkClass.StCrearEstructuraOperacion(data.IDVIno, data.Origen, data.NOperacion, oVinos, IDVinosOrigenOLD)
        ProcessServer.ExecuteTask(Of BdgWorkClass.StCrearEstructuraOperacion)(AddressOf BdgWorkClass.CrearEstructuraOperacion, StCrear, services)
    End Sub


    <Serializable()> _
  Public Class StCalcularEstructura
        Public FactorComp As Double
        Public IDArtDst As String
        Public IDUDMedidaDst As String
        Public Origen As DataTable

        Public Sub New()
        End Sub

        Public Sub New(ByVal FactorComp As Double, ByVal IDArtDst As String, ByVal IDUDMedidaDst As String, ByVal Origen As DataTable)
            Me.FactorComp = FactorComp
            Me.IDArtDst = IDArtDst
            Me.IDUDMedidaDst = IDUDMedidaDst
            Me.Origen = Origen
        End Sub
    End Class

    <Task()> Public Shared Function CalcularEstructura(ByVal data As StCalcularEstructura, ByVal services As ServiceProvider) As VinoComponente()
        Dim aVC(-1) As VinoComponente
        If Not data.Origen Is Nothing Then
            For Each oRw As DataRow In data.Origen.Rows
                Dim Merma As Double
                Dim Cantidad As Double

                If oRw.IsNull(_OV.Cantidad) Then
                    Cantidad = 0
                Else
                    Cantidad = oRw(_OV.Cantidad) * data.FactorComp
                End If
                If oRw.IsNull(_OV.Merma) Then
                    Merma = 0
                Else
                    Merma = oRw(_OV.Merma) * data.FactorComp
                End If
                If Merma > 0 Then Cantidad = Cantidad - Merma

                ReDim Preserve aVC(aVC.Length)
                aVC(aVC.Length - 1) = New VinoComponente(oRw(_OV.IDVino), Cantidad, 1)
                Dim dataFactor As New ArticuloUnidadAB.DatosFactorConversion(data.IDArtDst, oRw(_V.IDUdMedida), data.IDUDMedidaDst)
                aVC(aVC.Length - 1).Factor = ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, dataFactor, services)
                aVC(aVC.Length - 1).Merma = Merma
            Next
        End If

        Return aVC
    End Function

    <Serializable()> _
    Public Class DataCambiarIDVino
        Public OldID As Guid
        Public NewID As Guid
        Public Tables As List(Of DataTable)

        Public Sub New()
        End Sub

        Public Sub New(ByVal OldID As Guid, ByVal NewID As Guid, ByVal Tables As List(Of DataTable))
            Me.OldID = OldID
            Me.NewID = NewID
            Me.Tables = Tables
        End Sub
    End Class

    <Task()> Public Shared Sub CambiarIDVino(ByVal data As DataCambiarIDVino, ByVal services As ServiceProvider)
        For Each dt As DataTable In data.Tables
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                Dim RegsVinoModificar As List(Of DataRow) = (From c In dt Where c.RowState <> DataRowState.Deleted AndAlso Not c.IsNull("IDVino") AndAlso c("IDVino").Equals(data.OldID)).ToList()
                If Not RegsVinoModificar Is Nothing AndAlso RegsVinoModificar.Count > 0 Then
                    For Each oRw As DataRow In RegsVinoModificar
                        oRw("IDVino") = data.NewID
                    Next
                End If
            End If
        Next
    End Sub
#End Region

#Region "Tareas GlobalesLineas"

    <Task()> Public Shared Sub AsignarDatosGlobalesLineas(ByVal data As DocumentoBdgOperacionReal, ByVal services As ServiceProvider)
        If Not data.GetOperacionVinoDestino Is Nothing AndAlso data.GetOperacionVinoDestino.Length > 0 Then
            ProcessServer.ExecuteTask(Of DocumentoBdgOperacionReal)(AddressOf AsignarDatosGlobalesLineasMateriales, data, services)
            ProcessServer.ExecuteTask(Of DocumentoBdgOperacionReal)(AddressOf AsignarDatosGlobalesLineasMOD, data, services)
            ProcessServer.ExecuteTask(Of DocumentoBdgOperacionReal)(AddressOf AsignarDatosGlobalesLineasCentros, data, services)
            ProcessServer.ExecuteTask(Of DocumentoBdgOperacionReal)(AddressOf AsignarDatosGlobalesLineasVarios, data, services)
        End If
    End Sub

    <Task()> Public Shared Sub AsignarDatosGlobalesLineasMateriales(ByVal data As DocumentoBdgOperacionReal, ByVal services As ServiceProvider)
        If Not data.GetOperacionVinoDestino Is Nothing AndAlso data.GetOperacionVinoDestino.Length > 0 Then

            Dim CambiosEnVinoDestino As List(Of DataRow) = (From c In data.GetOperacionVinoDestino Where c.RowState <> DataRowState.Unchanged).ToList
            Dim CambiosEnMaterialesGlobales As List(Of DataRow) = (From c In data.dtOperacionMaterial Where c.RowState <> DataRowState.Unchanged).ToList
            If (CambiosEnVinoDestino Is Nothing OrElse CambiosEnVinoDestino.Count = 0) AndAlso _
                (CambiosEnMaterialesGlobales Is Nothing OrElse CambiosEnMaterialesGlobales.Count = 0) Then Exit Sub

            '//Eliminamos los materiales que vienen de globales y sus lotes
            If Not data.dtOperacionVinoMaterial Is Nothing AndAlso data.dtOperacionVinoMaterial.Rows.Count > 0 Then
                Dim MaterialesAsociadosAGlobales As List(Of DataRow) = (From c In data.dtOperacionVinoMaterial _
                                                                        Where c.RowState <> DataRowState.Deleted AndAlso _
                                                                                Not c.IsNull("IDOperacionMaterial") _
                                                                        Select c).ToList
                If Not MaterialesAsociadosAGlobales Is Nothing AndAlso MaterialesAsociadosAGlobales.Count > 0 Then
                    For Each drMaterial As DataRow In MaterialesAsociadosAGlobales
                        If Not data.dtOperacionVinoMaterialLote Is Nothing AndAlso data.dtOperacionVinoMaterialLote.Rows.Count > 0 Then
                            Dim LotesAsociados As List(Of DataRow) = (From c In data.dtOperacionVinoMaterialLote _
                                                                      Where c.RowState <> DataRowState.Deleted AndAlso _
                                                                            Not c.IsNull("IDVinoMaterial") AndAlso _
                                                                            c("IDVinoMaterial") = drMaterial("IDVinoMaterial") _
                                                                      Select c).ToList
                            If Not LotesAsociados Is Nothing AndAlso LotesAsociados.Count > 0 Then
                                For Each drLoteMaterial As DataRow In LotesAsociados
                                    drLoteMaterial.Delete()
                                Next
                            End If
                        End If

                        drMaterial.Delete()
                    Next
                End If
            End If


            If Not data.dtOperacionMaterial Is Nothing AndAlso data.dtOperacionMaterial.Rows.Count > 0 Then
                Dim LineasDestino As List(Of DataRow) = (From c In data.GetOperacionVinoDestino Where c.RowState <> DataRowState.Deleted Select c).ToList
                If Not LineasDestino Is Nothing AndAlso LineasDestino.Count > 0 Then

                    Dim datImputMat As New BdgGeneral.StImputacion(data.HeaderRow("IDTipoOperacion"), data.HeaderRow("NOperacion"), enumBdgOrigenOperacion.Real, False, data.HeaderRow("Fecha"))
                    datImputMat.DtMaterialesGlobal = data.dtOperacionMaterial : datImputMat.DtMateriales = data.dtOperacionVinoMaterial
                    datImputMat.DtMaterialesGlobalLotes = data.dtOperacionMaterialLote : datImputMat.DtMaterialesLotes = data.dtOperacionVinoMaterialLote
                    datImputMat.TotalLitrosDestino = Nz(data.GetOperacionVinoDestino.CopyToDataTable.Compute("SUM(Litros)", String.Empty), 0)
                    For Each DrDestVino As DataRow In LineasDestino
                        datImputMat.LineaDestino.Add(New BdgGeneral.StDestino(DrDestVino("IDVino"), DrDestVino("IDArticulo"), Nz(DrDestVino("IDEstructura"), String.Empty), DrDestVino("Litros"), DrDestVino("Cantidad")))
                    Next

                    '//Imputación de Materiales desde Imputaciones Globales  
                    ProcessServer.ExecuteTask(Of BdgGeneral.StImputacion, DataTable)(AddressOf BdgGeneral.ImputacionLineasMaterialesGenerales, datImputMat, services)
                    ProcessServer.ExecuteTask(Of BdgGeneral.StImputacion, DataTable)(AddressOf BdgGeneral.ImputacionLineasMaterialesLotes, datImputMat, services)

                End If
            End If

        End If
    End Sub

    <Task()> Public Shared Sub AsignarDatosGlobalesLineasMOD(ByVal data As DocumentoBdgOperacionReal, ByVal services As ServiceProvider)
        If Not data.GetOperacionVinoDestino Is Nothing AndAlso data.GetOperacionVinoDestino.Length > 0 Then

            Dim CambiosEnVinoDestino As List(Of DataRow) = (From c In data.GetOperacionVinoDestino Where c.RowState <> DataRowState.Unchanged).ToList
            Dim CambiosEnMODGlobales As List(Of DataRow) = (From c In data.dtOperacionMOD Where c.RowState <> DataRowState.Unchanged).ToList
            If (CambiosEnVinoDestino Is Nothing OrElse CambiosEnVinoDestino.Count = 0) AndAlso _
                (CambiosEnMODGlobales Is Nothing OrElse CambiosEnMODGlobales.Count = 0) Then Exit Sub


            If Not data.dtOperacionVinoMOD Is Nothing AndAlso data.dtOperacionVinoMOD.Rows.Count > 0 Then
                Dim MODAsociadosAGlobales As List(Of DataRow) = (From c In data.dtOperacionVinoMOD _
                                                                        Where Not c.IsNull("IDOperacionMOD") _
                                                                        Select c).ToList
                If Not MODAsociadosAGlobales Is Nothing AndAlso MODAsociadosAGlobales.Count > 0 Then
                    For Each drMOD As DataRow In MODAsociadosAGlobales
                        drMOD.Delete()
                    Next
                End If
            End If

            If Not data.dtOperacionMOD Is Nothing AndAlso data.dtOperacionMOD.Rows.Count > 0 Then
                Dim LineasDestino As List(Of DataRow) = (From c In data.GetOperacionVinoDestino Where c.RowState <> DataRowState.Deleted Select c).ToList
                If Not LineasDestino Is Nothing AndAlso LineasDestino.Count > 0 Then

                    Dim datImputMOD As New BdgGeneral.StImputacion(data.HeaderRow("IDTipoOperacion"), data.HeaderRow("NOperacion"), enumBdgOrigenOperacion.Real, False, data.HeaderRow("Fecha"))
                    datImputMOD.DtMODGlobal = data.dtOperacionMOD : datImputMOD.DtMOD = data.dtOperacionVinoMOD

                    datImputMOD.TotalLitrosDestino = Nz(data.GetOperacionVinoDestino.CopyToDataTable.Compute("SUM(Litros)", String.Empty), 0)
                    For Each DrDestVino As DataRow In LineasDestino
                        datImputMOD.LineaDestino.Add(New BdgGeneral.StDestino(DrDestVino("IDVino"), DrDestVino("IDArticulo"), Nz(DrDestVino("IDEstructura"), String.Empty), DrDestVino("Litros"), DrDestVino("Cantidad")))
                    Next

                    '//Imputación de MOD desde Imputaciones Globales  
                    ProcessServer.ExecuteTask(Of BdgGeneral.StImputacion)(AddressOf BdgGeneral.ImputacionLineasMODGenerales, datImputMOD, services)
                End If
            End If
        End If
    End Sub


    <Task()> Public Shared Sub AsignarDatosGlobalesLineasCentros(ByVal data As DocumentoBdgOperacionReal, ByVal services As ServiceProvider)
        If Not data.GetOperacionVinoDestino Is Nothing AndAlso data.GetOperacionVinoDestino.Length > 0 Then
            Dim CambiosEnVinoDestino As List(Of DataRow) = (From c In data.GetOperacionVinoDestino Where c.RowState <> DataRowState.Unchanged).ToList
            Dim CambiosEnCentrosGlobales As List(Of DataRow) = (From c In data.dtOperacionCentro Where c.RowState <> DataRowState.Unchanged).ToList
            If (CambiosEnVinoDestino Is Nothing OrElse CambiosEnVinoDestino.Count = 0) AndAlso _
                (CambiosEnCentrosGlobales Is Nothing OrElse CambiosEnCentrosGlobales.Count = 0) Then Exit Sub


            If Not data.dtOperacionVinoCentro Is Nothing AndAlso data.dtOperacionVinoCentro.Rows.Count > 0 Then
                Dim CentrosAsociadosAGlobales As List(Of DataRow) = (From c In data.dtOperacionVinoCentro _
                                                                        Where Not c.IsNull("IDOperacionCentro") _
                                                                        Select c).ToList
                If Not CentrosAsociadosAGlobales Is Nothing AndAlso CentrosAsociadosAGlobales.Count > 0 Then
                    For Each drCentro As DataRow In CentrosAsociadosAGlobales
                        drCentro.Delete()
                    Next
                End If
            End If

            If Not data.dtOperacionCentro Is Nothing AndAlso data.dtOperacionCentro.Rows.Count > 0 Then
                Dim LineasDestino As List(Of DataRow) = (From c In data.GetOperacionVinoDestino Where c.RowState <> DataRowState.Deleted Select c).ToList
                If Not LineasDestino Is Nothing AndAlso LineasDestino.Count > 0 Then

                    Dim datImputCentros As New BdgGeneral.StImputacion(data.HeaderRow("IDTipoOperacion"), data.HeaderRow("NOperacion"), enumBdgOrigenOperacion.Real, False, data.HeaderRow("Fecha"))
                    datImputCentros.DtCentrosGlobal = data.dtOperacionCentro : datImputCentros.DtCentros = data.dtOperacionVinoCentro

                    datImputCentros.TotalLitrosDestino = Nz(data.GetOperacionVinoDestino.CopyToDataTable.Compute("SUM(Litros)", String.Empty), 0)
                    For Each DrDestVino As DataRow In LineasDestino
                        datImputCentros.LineaDestino.Add(New BdgGeneral.StDestino(DrDestVino("IDVino"), DrDestVino("IDArticulo"), Nz(DrDestVino("IDEstructura"), String.Empty), DrDestVino("Litros"), DrDestVino("Cantidad")))
                    Next

                    '//Imputación de Centros desde Imputaciones Globales  
                    ProcessServer.ExecuteTask(Of BdgGeneral.StImputacion)(AddressOf BdgGeneral.ImputacionLineasCentrosGenerales, datImputCentros, services)
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarDatosGlobalesLineasVarios(ByVal data As DocumentoBdgOperacionReal, ByVal services As ServiceProvider)
        If Not data.GetOperacionVinoDestino Is Nothing AndAlso data.GetOperacionVinoDestino.Length > 0 Then
            Dim CambiosEnVinoDestino As List(Of DataRow) = (From c In data.GetOperacionVinoDestino Where c.RowState <> DataRowState.Unchanged).ToList
            Dim CambiosEnVariosGlobales As List(Of DataRow) = (From c In data.dtOperacionVarios Where c.RowState <> DataRowState.Unchanged).ToList
            If (CambiosEnVinoDestino Is Nothing OrElse CambiosEnVinoDestino.Count = 0) AndAlso _
                (CambiosEnVariosGlobales Is Nothing OrElse CambiosEnVariosGlobales.Count = 0) Then Exit Sub


            If Not data.dtOperacionVinoVarios Is Nothing AndAlso data.dtOperacionVinoVarios.Rows.Count > 0 Then
                Dim VariosAsociadosAGlobales As List(Of DataRow) = (From c In data.dtOperacionVinoVarios _
                                                                        Where Not c.IsNull("IDOperacionVarios") _
                                                                        Select c).ToList
                If Not VariosAsociadosAGlobales Is Nothing AndAlso VariosAsociadosAGlobales.Count > 0 Then
                    For Each drVarios As DataRow In VariosAsociadosAGlobales
                        drVarios.Delete()
                    Next
                End If
            End If

            If Not data.dtOperacionVarios Is Nothing AndAlso data.dtOperacionVarios.Rows.Count > 0 Then
                Dim LineasDestino As List(Of DataRow) = (From c In data.GetOperacionVinoDestino Where c.RowState <> DataRowState.Deleted Select c).ToList
                If Not LineasDestino Is Nothing AndAlso LineasDestino.Count > 0 Then

                    Dim datImputVarios As New BdgGeneral.StImputacion(data.HeaderRow("IDTipoOperacion"), data.HeaderRow("NOperacion"), enumBdgOrigenOperacion.Real, False, data.HeaderRow("Fecha"))
                    datImputVarios.DtVariosGlobal = data.dtOperacionVarios : datImputVarios.DtVarios = data.dtOperacionVinoVarios

                    datImputVarios.TotalLitrosDestino = Nz(data.GetOperacionVinoDestino.CopyToDataTable.Compute("SUM(Litros)", String.Empty), 0)
                    For Each DrDestVino As DataRow In LineasDestino
                        datImputVarios.LineaDestino.Add(New BdgGeneral.StDestino(DrDestVino("IDVino"), DrDestVino("IDArticulo"), Nz(DrDestVino("IDEstructura"), String.Empty), DrDestVino("Litros"), DrDestVino("Cantidad")))
                    Next

                    '//Imputación de Varios desde Imputaciones Globales  
                    ProcessServer.ExecuteTask(Of BdgGeneral.StImputacion)(AddressOf BdgGeneral.ImputacionLineasVariosGenerales, datImputVarios, services)
                End If
            End If
        End If
    End Sub

#End Region

#Region "Tareas VinoMaterial"

    <Task()> Public Shared Sub CrearMovimientosMaterialesVino(ByVal data As DocumentoBdgOperacionReal, ByVal services As ServiceProvider)
        data.dtOperacionVinoMaterialLote.AcceptChanges()
        Dim stCrearMov As New BdgVinoMaterialLote.StCrearMovimientos(data.HeaderRow("NOperacion"), data.dtOperacionVinoMaterialLote, data.dtOperacionVinoMaterial)
        ProcessServer.ExecuteTask(Of BdgVinoMaterialLote.StCrearMovimientos)(AddressOf BdgVinoMaterialLote.CrearMovimientos, stCrearMov, services)
    End Sub


    <Task()> Public Shared Sub AsignarDatosVinoMaterial(ByVal data As DocumentoBdgOperacionReal, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DataTable)(AddressOf BdgVinoMaterial.DescuentoMateriales, data.dtOperacionVinoMaterial, services)
    End Sub

#End Region

#Region "Tareas VinoCentro"

    <Task()> Public Shared Sub AsignarDatosVinoCentro(ByVal data As DocumentoBdgOperacionReal, ByVal services As ServiceProvider)
        If data.dtOperacionVinoCentro Is Nothing OrElse data.dtOperacionVinoCentro.Rows.Count = 0 Then Exit Sub

        Dim OperacionVinoCentroModif As List(Of DataRow) = (From c In data.dtOperacionVinoCentro Where c.RowState = DataRowState.Added OrElse c.RowState = DataRowState.Modified Select c).ToList
        If Not OperacionVinoCentroModif Is Nothing AndAlso OperacionVinoCentroModif.Count > 0 Then
            For Each Dr As DataRow In OperacionVinoCentroModif ' data.dtOperacionVinoCentro.Select(String.Empty, String.Empty, DataViewRowState.Added Or DataViewRowState.ModifiedCurrent Or DataViewRowState.ModifiedCurrent)
                If Dr.RowState = DataRowState.Added Then
                    If Length(Dr("IDVinoCentro")) = 0 Then Dr("IdVinoCentro") = Guid.NewGuid
                End If

                Dim StGet As New BdgVinoCentro.StGetTasa(Dr("IDCentro"), Dr("Fecha"), Dr("IDIncidencia") & String.Empty)
                Dim oTasa As TasaInfo = ProcessServer.ExecuteTask(Of BdgVinoCentro.StGetTasa, TasaInfo)(AddressOf BdgVinoCentro.GetTasa, StGet, services)
                Dr(_VC.UDTiempo) = oTasa.UdTiempo
                Dr(_VC.TasaD) = oTasa.TasaD
                Dr(_VC.TasaI) = oTasa.TasaI
                Dr(_VC.TasaF) = oTasa.TasaF
                Dr(_VC.TasaV) = oTasa.TasaV
                Dr(_VC.TasaFiscal) = oTasa.TasaFscl
                If data.dtOperacionVinoCentro.Columns.Contains("FechaFinCosteNave") Then Dr("FechaFinCosteNave") = Dr("FechaFinCosteNave")
                If data.dtOperacionVinoCentro.Columns.Contains("FechaInicioCosteNave") Then Dr("FechaInicioCosteNave") = Dr("FechaInicioCosteNave")
            Next
        End If

    End Sub

    <Task()> Public Shared Sub ActualizarVinoCentroTasa(ByVal data As DocumentoBdgOperacionReal, ByVal services As ServiceProvider)
        If Not data.dtOperacionVinoCentro Is Nothing AndAlso data.dtOperacionVinoCentro.Rows.Count > 0 Then
            Dim dtVinoCentrosActualizar As DataTable = data.dtOperacionVinoCentro.GetChanges(DataRowState.Added Or DataRowState.Modified)
            If Not dtVinoCentrosActualizar Is Nothing AndAlso dtVinoCentrosActualizar.Rows.Count > 0 Then
                Dim StCrear As New BdgVinoCentroTasa.StCrearVinoCentroTasa(data.HeaderRow("Fecha"), dtVinoCentrosActualizar)
                ProcessServer.ExecuteTask(Of BdgVinoCentroTasa.StCrearVinoCentroTasa)(AddressOf BdgVinoCentroTasa.CrearVinoCentroTasa, StCrear, services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarCostesVino(ByVal data As DocumentoBdgOperacionReal, ByVal services As ServiceProvider)
        Dim CostesIDVino As ProcesoBdgOperacion.DataCostesIDVino = services.GetService(Of ProcesoBdgOperacion.DataCostesIDVino)()
        If Not CostesIDVino.CostesElaboracion Is Nothing AndAlso CostesIDVino.CostesElaboracion.Rows.Count > 0 Then
            Dim dtCostesElaboracion As DataTable = CostesIDVino.CostesElaboracion.Clone
            For Each dr As DataRow In CostesIDVino.CostesElaboracion.Rows
                dtCostesElaboracion.Rows.Add(dr.ItemArray)
            Next
            Dim VM As New BdgVinoMaterial
            VM.Update(dtCostesElaboracion)
            CostesIDVino.CostesElaboracion.Clear()
        End If
        If Not CostesIDVino.CostesVendimia Is Nothing AndAlso CostesIDVino.CostesVendimia.Rows.Count > 0 Then
            Dim dtCostesElaboracionHist As DataTable = CostesIDVino.CostesVendimia.Clone
            For Each dr As DataRow In CostesIDVino.CostesVendimia.Rows
                dtCostesElaboracionHist.Rows.Add(dr.ItemArray)
            Next
            Dim CVH As New BdgCosteVendimiaHist
            CVH.Update(dtCostesElaboracionHist)
            CostesIDVino.CostesVendimia.Clear()
        End If

        If Not CostesIDVino.CostesEstanciaNave Is Nothing AndAlso CostesIDVino.CostesEstanciaNave.Rows.Count > 0 Then
            Dim dtCostesEstanciaNave As DataTable = CostesIDVino.CostesEstanciaNave.Clone
            For Each dr As DataRow In CostesIDVino.CostesEstanciaNave.Rows
                dtCostesEstanciaNave.Rows.Add(dr.ItemArray)
            Next
            Dim VM As New BdgVinoCentro
            VM.Update(dtCostesEstanciaNave)
            CostesIDVino.CostesEstanciaNave.Clear()
        End If

        If Not CostesIDVino.CostesEstanciaNaveTasas Is Nothing AndAlso CostesIDVino.CostesEstanciaNaveTasas.Rows.Count > 0 Then
            Dim dtCostesEstanciaNaveTasas As DataTable = CostesIDVino.CostesEstanciaNaveTasas.Clone
            For Each dr As DataRow In CostesIDVino.CostesEstanciaNaveTasas.Rows
                dtCostesEstanciaNaveTasas.Rows.Add(dr.ItemArray)
            Next
            Dim VM As New BdgVinoCentroTasa
            VM.Update(dtCostesEstanciaNaveTasas)
            CostesIDVino.CostesEstanciaNaveTasas.Clear()
        End If


    End Sub

    <Task()> Public Shared Sub DeleteCostesVino(ByVal data As Object, ByVal services As ServiceProvider)
        Dim CostesIDVino As ProcesoBdgOperacion.DataCostesIDVino = services.GetService(Of ProcesoBdgOperacion.DataCostesIDVino)()
        If Not CostesIDVino.CostesElaboracionDel Is Nothing AndAlso CostesIDVino.CostesElaboracionDel.Rows.Count > 0 Then
            For Each drDel As DataRow In CostesIDVino.CostesElaboracionDel.Rows
                BdgVinoMaterial.DeleteRowCascade(drDel, services)
            Next
        End If
        If Not CostesIDVino.CostesEstanciaNaveDel Is Nothing AndAlso CostesIDVino.CostesEstanciaNaveDel.Rows.Count > 0 Then
            For Each drDel As DataRow In CostesIDVino.CostesEstanciaNaveDel.Rows
                BdgVinoCentro.DeleteRowCascade(drDel, services)
            Next
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarAnalisisVinoNoOperacion(ByVal data As Object, ByVal services As ServiceProvider)
        Dim AnalisisIDVino As ProcesoBdgOperacion.DataAnalisisIDVino = services.GetService(Of ProcesoBdgOperacion.DataAnalisisIDVino)()
        If Not AnalisisIDVino.Analisis Is Nothing AndAlso AnalisisIDVino.Analisis.Rows.Count > 0 Then
            Dim dtAnalisis As DataTable = AnalisisIDVino.Analisis.Clone
            For Each dr As DataRow In AnalisisIDVino.Analisis.Rows
                dtAnalisis.Rows.Add(dr.ItemArray)
            Next
            Dim VA As New BdgVinoAnalisis
            VA.Update(dtAnalisis)
            AnalisisIDVino.Analisis.Clear()
        End If
        If Not AnalisisIDVino.AnalisisVariable Is Nothing AndAlso AnalisisIDVino.AnalisisVariable.Rows.Count > 0 Then
            Dim dtAnalisisVariable As DataTable = AnalisisIDVino.AnalisisVariable.Clone
            For Each dr As DataRow In AnalisisIDVino.AnalisisVariable.Rows
                dtAnalisisVariable.Rows.Add(dr.ItemArray)
            Next
            Dim VAV As New BdgVinoVariable
            VAV.Update(dtAnalisisVariable)
            AnalisisIDVino.AnalisisVariable.Clear()
        End If
    End Sub

    <Task()> Public Shared Sub DeleteAnalisisVinoNoOperacion(ByVal data As Object, ByVal services As ServiceProvider)
        Dim AnalisisIDVino As ProcesoBdgOperacion.DataAnalisisIDVino = services.GetService(Of ProcesoBdgOperacion.DataAnalisisIDVino)()
        If Not AnalisisIDVino.AnalisisDel Is Nothing AndAlso AnalisisIDVino.AnalisisDel.Rows.Count > 0 Then
            For Each drDel As DataRow In AnalisisIDVino.AnalisisDel.Rows
                BdgVinoAnalisis.DeleteRowCascade(drDel, services)
            Next
            AnalisisIDVino.AnalisisDel.Clear()
        End If
    End Sub

#End Region

#Region "Tareas VinoAnalisis"

    <Task()> Public Shared Sub AsignarDatosVinoAnalisisVariable(ByVal data As DocumentoBdgOperacionReal, ByVal services As ServiceProvider)
        If data.dtOperacionVinoAnalisis Is Nothing Then Exit Sub

        Dim NuevosAnalisis As List(Of DataRow) = (From c In data.dtOperacionVinoAnalisis Where c.RowState = DataRowState.Added Select c).ToList
        If Not NuevosAnalisis Is Nothing AndAlso NuevosAnalisis.Count > 0 Then
            For Each Dr As DataRow In NuevosAnalisis
                If Length(Dr("IdContador")) = 0 Then
                    Dim datacont As New Contador.DatosDefaultCounterValue(Dr, "BdgVinoAnalisis", _VA.NVinoAnalisis)
                    ProcessServer.ExecuteTask(Of Contador.DatosDefaultCounterValue)(AddressOf Contador.LoadDefaultCounterValue, datacont, services)
                End If
                If Length(Dr("IdContador")) > 0 Then
                    Dr("NVinoAnalisis") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, Dr("IdContador"), services)
                Else
                    ApplicationService.GenerateError("Debe indicar un contador para la entidad 'BdgVinoAnalisis'")
                End If
                If Length(Dr("IDVinoAnalisis")) = 0 OrElse CType(Dr("IDVinoAnalisis"), Guid).Equals(Guid.Empty) Then Dr("IDVinoAnalisis") = Guid.NewGuid()
            Next
        End If
    End Sub

#End Region

#Region "Tareas VinoMOD"

    <Task()> Public Shared Sub AsignarDatosVinoMOD(ByVal data As DocumentoBdgOperacionReal, ByVal services As ServiceProvider)
        If data.dtOperacionVinoMOD Is Nothing OrElse data.dtOperacionVinoMOD.Rows.Count = 0 Then Exit Sub

        Dim OperacionVinoMODModif As List(Of DataRow) = (From c In data.dtOperacionVinoMOD Where c.RowState = DataRowState.Added OrElse c.RowState = DataRowState.Modified Select c).ToList
        If Not OperacionVinoMODModif Is Nothing AndAlso OperacionVinoMODModif.Count > 0 Then
            For Each Dr As DataRow In OperacionVinoMODModif
                If Length(Dr("IDHora")) = 0 Then
                    Dim mstrHoraPred As String = ProcessServer.ExecuteTask(Of String, String)(AddressOf HoraCategoria.GetHoraPredeterminada, Dr("IDCategoria"), services)
                    Dr("IDHora") = mstrHoraPred
                End If

                Dim dataHora As New General.HoraCategoria.DatosPrecioHoraCatOper(Dr("IDCategoria") & String.Empty, Dr("IDHora") & String.Empty, Dr("Fecha"), Dr("IDOperario"))
                Dr("Tasa") = ProcessServer.ExecuteTask(Of General.HoraCategoria.DatosPrecioHoraCatOper, Double)(AddressOf General.HoraCategoria.ObtenerPrecioHoraCategoriaOperario, dataHora, services)
            Next
        End If
    End Sub

#End Region

#End Region

#Region "Tareas Generales"

    <Task()> Public Shared Function GetOperacionTotalOrigen(ByVal data As DocumentoBdgOperacionReal, ByVal services As ServiceProvider) As Double
        Dim DblTotal As Double = 0
        If Not data.GetOperacionVinoOrigen Is Nothing AndAlso data.GetOperacionVinoOrigen.Length > 0 Then
            DblTotal = Nz((Aggregate c In data.GetOperacionVinoOrigen Into Sum(CDbl(c("Litros")))), 0)
        End If
        Return DblTotal
    End Function

    <Task()> Public Shared Function GetOperacionTotalDestino(ByVal data As DocumentoBdgOperacionReal, ByVal services As ServiceProvider) As Double
        Dim DblTotal As Double = 0
        If Not Data.GetOperacionVinoDestino Is Nothing AndAlso Data.GetOperacionVinoDestino.Length > 0 Then
            DblTotal = Nz((Aggregate c In data.GetOperacionVinoDestino Into Sum(CDbl(c("Litros")))), 0)
        End If
        Return DblTotal
    End Function

    <Task()> Public Shared Function GetOperacionTotalPendiente(ByVal data As DocumentoBdgOperacionReal, ByVal services As ServiceProvider) As Double
        Dim DblTotal As Double = 0, DblTotalMermaNeg As Double = 0, DblTotalMermaPos As Double = 0
        Dim AppParams As BdgParametrosOperaciones = services.GetService(Of BdgParametrosOperaciones)()
        Dim IDUDMedidaLitros As String = AppParams.UnidadesCampoLitros

        If Not data.GetOperacionVinoOrigen Is Nothing AndAlso data.GetOperacionVinoOrigen.Count > 0 Then
            Dim OrigenesUdMedida As List(Of DataRow) = (From c In data.GetOperacionVinoOrigen Where Not c.IsNull("IDArticulo") AndAlso Not c.IsNull("IDUDMedida") Select c).ToList
            If Not OrigenesUdMedida Is Nothing AndAlso OrigenesUdMedida.Count > 0 Then
                For Each dr As DataRow In OrigenesUdMedida
                    If Length(dr("IDArticulo")) > 0 AndAlso Length(dr("IDUDMedida")) > 0 Then
                        Dim datFactor As New ArticuloUnidadAB.DatosFactorConversion(dr("IDArticulo"), dr("IDUDMedida"), IDUDMedidaLitros)
                        If Nz(dr("Merma"), 0) < 0 Then
                            DblTotalMermaNeg += Math.Abs(Nz(dr("Merma"), 0)) * ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, datFactor, services)
                        Else : DblTotalMermaPos += Nz(dr("Merma"), 0) * ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, datFactor, services)
                        End If
                    End If
                Next
            End If

        End If

        Dim DblTotalOrigen As Double = ProcessServer.ExecuteTask(Of DocumentoBdgOperacionReal, Double)(AddressOf GetOperacionTotalOrigen, data, services)
        Dim DblTotalDestino As Double = ProcessServer.ExecuteTask(Of DocumentoBdgOperacionReal, Double)(AddressOf GetOperacionTotalDestino, data, services)
        DblTotal = DblTotalOrigen - DblTotalDestino - DblTotalMermaPos + DblTotalMermaNeg
        Return DblTotal
    End Function

    '<Task()> Public Shared Function GetOperacionEstadoVino(ByVal data As DocumentoBdgOperacionReal, ByVal services As ServiceProvider) As String
    '    Dim IDEstadoVino As String = data.GetTipoOperEstadoVino
    '    If Length(IDEstadoVino) = 0 AndAlso Not data.GetOperacionVinoOrigen Is Nothing Then
    '        Dim EstadoVinoOrigenes As List(Of DataRow) = (From c In data.GetOperacionVinoOrigen _
    '                                                      Where Not c.IsNull("IDEstadoVino")).ToList()
    '        If Not EstadoVinoOrigenes Is Nothing AndAlso EstadoVinoOrigenes.Count > 0 Then
    '            IDEstadoVino = EstadoVinoOrigenes(0)("IDEstadoVino")
    '        End If
    '    End If
    '    Return IDEstadoVino
    'End Function

#End Region

End Class

<Serializable()> _
Public Class FactoresRecalcularEstructura
    Public VinoFactor As New Dictionary(Of Guid, Double)
End Class