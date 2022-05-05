<Transactional()> _
Public Class BdgStock
    Inherits ContextBoundObject
    Implements IStock

    Public Function SincronizarSalida(ByVal NumeroMovimiento As Integer, ByVal data As Negocio.StockData) As Negocio.ProcesoStocks.DataVinoQ Implements Negocio.IStock.SincronizarSalida
        If Not data Is Nothing Then
            Dim Services As New ServiceProvider
            Dim dataVQ As New ProcesoStocks.DataVinoQ(data.Articulo, data.Ubicacion, data.Lote, data.Almacen)
            dataVQ = ProcessServer.ExecuteTask(Of ProcesoStocks.DataVinoQ, ProcesoStocks.DataVinoQ)(AddressOf BdgWorkClass.GetIDVinoCantidad, dataVQ, Services)
            If dataVQ Is Nothing Then
                Dim oWC As New BdgWorkClass
                If data.Cantidad < 0 Then
                    Dim StCrear As New BdgWorkClass.StCrearVinoOrigenCantidad(data.Ubicacion, data.Articulo, data.Lote, data.FechaDocumento, BdgOrigenVino.Interno, -data.Cantidad)
                    StCrear.IDAlmacen = data.Almacen
                    ProcessServer.ExecuteTask(Of BdgWorkClass.StCrearVinoOrigenCantidad, Guid)(AddressOf BdgWorkClass.CrearVinoOrigenCantidad, StCrear, Services)
                    dataVQ = New ProcesoStocks.DataVinoQ(data.Articulo, data.Ubicacion, data.Lote, data.Almacen)
                    dataVQ = ProcessServer.ExecuteTask(Of ProcesoStocks.DataVinoQ, ProcesoStocks.DataVinoQ)(AddressOf BdgWorkClass.GetIDVinoCantidad, dataVQ, Services)
                End If
            Else
                Dim StCambio As New BdgWorkClass.StCambiarOcupacion(dataVQ.IDVino, -data.Cantidad)
                ProcessServer.ExecuteTask(Of BdgWorkClass.StCambiarOcupacion)(AddressOf BdgWorkClass.CambiarOcupacion, StCambio, Services)
            End If
            Return dataVQ
        End If
    End Function

    Public Function SincronizarEntrada(ByVal NumeroMovimiento As Integer, ByVal data As Negocio.StockData) As Negocio.ProcesoStocks.DataVinoQ Implements Negocio.IStock.SincronizarEntrada
        If Not data Is Nothing Then
            Dim Services As New ServiceProvider
            Dim oWC As New BdgWorkClass
            Dim dataVQ As New ProcesoStocks.DataVinoQ(data.Articulo, data.Ubicacion, data.Lote, data.Almacen)
            dataVQ = ProcessServer.ExecuteTask(Of ProcesoStocks.DataVinoQ, ProcesoStocks.DataVinoQ)(AddressOf BdgWorkClass.GetIDVinoCantidad, dataVQ, Services)
            If dataVQ Is Nothing Then
                Dim StCrear As New BdgWorkClass.StCrearVinoOrigenCantidad(data.Ubicacion, data.Articulo, data.Lote, data.FechaDocumento, BdgOrigenVino.Interno, data.Cantidad)
                StCrear.IDAlmacen = data.Almacen
                ProcessServer.ExecuteTask(Of BdgWorkClass.StCrearVinoOrigenCantidad, Guid)(AddressOf BdgWorkClass.CrearVinoOrigenCantidad, StCrear, Services)
                dataVQ = New ProcesoStocks.DataVinoQ(data.Articulo, data.Ubicacion, data.Lote, data.Almacen)
                dataVQ = ProcessServer.ExecuteTask(Of ProcesoStocks.DataVinoQ, ProcesoStocks.DataVinoQ)(AddressOf BdgWorkClass.GetIDVinoCantidad, dataVQ, Services)
            Else
                Dim StCambiar As New BdgWorkClass.StCambiarOcupacion(dataVQ.IDVino, data.Cantidad)
                ProcessServer.ExecuteTask(Of BdgWorkClass.StCambiarOcupacion)(AddressOf BdgWorkClass.CambiarOcupacion, StCambiar, Services)
            End If
            Return dataVQ
        End If
    End Function

    Public Function SincronizarEntradaTransferencia(ByVal NumeroMovimiento As Integer, ByVal dataentrada As Negocio.StockData, _
                                                      ByVal datasalida As Negocio.StockData, ByVal updateEntrada As StockUpdateData, _
                                                      ByVal updateSalida As StockUpdateData) As Negocio.ProcesoStocks.DataVinoQ Implements Negocio.IStock.SincronizarEntradaTransferencia

        If Not dataentrada Is Nothing Then
            Dim oWC As New BdgWorkClass

            Dim Services As New ServiceProvider
            Dim dataVQ_Origen As New ProcesoStocks.DataVinoQ(datasalida.Articulo, datasalida.Ubicacion, datasalida.Lote, datasalida.Almacen)

            If updateSalida.StockData.Traza.Equals(Guid.Empty) Then
                dataVQ_Origen = ProcessServer.ExecuteTask(Of ProcesoStocks.DataVinoQ, ProcesoStocks.DataVinoQ)(AddressOf BdgWorkClass.GetIDVinoCantidad, dataVQ_Origen, Services)
                If dataVQ_Origen Is Nothing Then
                    dataVQ_Origen = New ProcesoStocks.DataVinoQ
                    dataVQ_Origen.IDVino = Guid.Empty
                End If
            Else
                dataVQ_Origen = New ProcesoStocks.DataVinoQ
                dataVQ_Origen.IDVino = updateSalida.StockData.Traza
            End If
            Dim oArt As New Articulo
            Dim rwArt As DataRow = oArt.GetItemRow(dataentrada.Articulo)
            Dim strIDUdInt As String
            If Not rwArt.IsNull("IDUDInterna") Then strIDUdInt = rwArt("IDUDInterna")

            Dim strNOperacion As String = Right(CStr(updateEntrada.IDLineaMovimiento), 10)

            If dataVQ_Origen.IDVino.Equals(Guid.Empty) Then
                Dim StCrear As New BdgWorkClass.StCrearVinoOrigenCantidad(dataentrada.Ubicacion, dataentrada.Articulo, dataentrada.Lote, dataentrada.FechaDocumento, BdgOrigenVino.AlbaranTransferencia, dataentrada.Cantidad)
                StCrear.IDAlmacen = dataentrada.Almacen
                ProcessServer.ExecuteTask(Of BdgWorkClass.StCrearVinoOrigenCantidad, Guid)(AddressOf BdgWorkClass.CrearVinoOrigenCantidad, StCrear, Services)
            Else
                Dim StCrear As New BdgWorkClass.StCrearVino(dataentrada.Ubicacion, dataentrada.Articulo, dataentrada.Lote, dataentrada.FechaDocumento, BdgOrigenVino.AlbaranTransferencia, _
                                strIDUdInt, dataentrada.Cantidad, , strNOperacion, New VinoComponente() {New VinoComponente(dataVQ_Origen.IDVino, datasalida.Cantidad, 1)})
                StCrear.IDAlmacen = dataentrada.Almacen
                ProcessServer.ExecuteTask(Of BdgWorkClass.StCrearVino)(AddressOf BdgWorkClass.CrearVino, StCrear, Services)
            End If

            Dim dataVQ As New ProcesoStocks.DataVinoQ(dataentrada.Articulo, dataentrada.Ubicacion, dataentrada.Lote, dataentrada.Almacen)
            dataVQ = ProcessServer.ExecuteTask(Of ProcesoStocks.DataVinoQ, ProcesoStocks.DataVinoQ)(AddressOf BdgWorkClass.GetIDVinoCantidad, dataVQ, Services)
            Return dataVQ
        End If
    End Function

    Public Function SincronizarInventario(ByVal NumeroMovimiento As Integer, ByVal data As Negocio.StockData) As Negocio.ProcesoStocks.DataVinoQ Implements Negocio.IStock.SincronizarInventario
        If Not data Is Nothing Then
            Dim Services As New ServiceProvider
            Dim oWC As New BdgWorkClass

            Dim dataVQ As New ProcesoStocks.DataVinoQ(data.Articulo, data.Ubicacion, data.Lote, data.Almacen)
            dataVQ = ProcessServer.ExecuteTask(Of ProcesoStocks.DataVinoQ, ProcesoStocks.DataVinoQ)(AddressOf BdgWorkClass.GetIDVinoCantidad, dataVQ, Services)
            If dataVQ Is Nothing Then
                If data.Cantidad <> 0 Then 'Al hacer inventario vienen todos los lotes incluso los que ya tenían stock 0 y no se han tocado.
                    Dim StCrear As New BdgWorkClass.StCrearVinoOrigenCantidad(data.Ubicacion, data.Articulo, data.Lote, data.FechaDocumento, BdgOrigenVino.Interno, data.Cantidad)
                    StCrear.IDAlmacen = data.Almacen
                    ProcessServer.ExecuteTask(Of BdgWorkClass.StCrearVinoOrigenCantidad, Guid)(AddressOf BdgWorkClass.CrearVinoOrigenCantidad, StCrear, Services)
                    dataVQ = New ProcesoStocks.DataVinoQ(data.Articulo, data.Ubicacion, data.Lote, data.Almacen)
                    dataVQ = ProcessServer.ExecuteTask(Of ProcesoStocks.DataVinoQ, ProcesoStocks.DataVinoQ)(AddressOf BdgWorkClass.GetIDVinoCantidad, dataVQ, Services)
                End If
            ElseIf data.Cantidad - dataVQ.Cantidad <> 0 Then
                Dim StCambiar As New BdgWorkClass.StCambiarOcupacion(dataVQ.IDVino, data.Cantidad - dataVQ.Cantidad)
                ProcessServer.ExecuteTask(Of BdgWorkClass.StCambiarOcupacion)(AddressOf BdgWorkClass.CambiarOcupacion, StCambiar, Services)
            End If
            Return dataVQ
        End If
    End Function

    Public Function SincronizarEliminarMovimiento(ByVal NumeroMovimiento As Integer, ByVal IDLineaMovimiento As Integer, _
                                                  ByVal dataOriginal As Negocio.StockData) As Negocio.ProcesoStocks.DataVinoQ Implements Negocio.IStock.SincronizarEliminarMovimiento
        If Not dataOriginal Is Nothing AndAlso (dataOriginal.TipoMovimiento = enumTipoMovimiento.tmSalAlbaranVenta _
                                                    Or dataOriginal.TipoMovimiento = enumTipoMovimiento.tmEntAlbaranCompra _
                                                    Or dataOriginal.TipoMovimiento = enumTipoMovimiento.tmSalTransferencia _
                                                    Or dataOriginal.TipoMovimiento = enumTipoMovimiento.tmEntTransferencia) Then
            Dim Services As New ServiceProvider
            Dim strOperacion As String = Right(CStr(IDLineaMovimiento), 10)
            Dim Cantidad As Double = -dataOriginal.Cantidad

            'Buscamos el vino que está vivo ahora.
            Dim dataVQ As New ProcesoStocks.DataVinoQ(dataOriginal.Articulo, dataOriginal.Ubicacion, dataOriginal.Lote, dataOriginal.Almacen)
            dataVQ = ProcessServer.ExecuteTask(Of ProcesoStocks.DataVinoQ, ProcesoStocks.DataVinoQ)(AddressOf BdgWorkClass.GetIDVinoCantidad, dataVQ, Services)

            If dataVQ Is Nothing OrElse IsDBNull(dataVQ.IDVino) Then
                'Sino buscamos el vino original según el campo traza del movimiento.
                'Comprobamos que Artículo, Almacén, Lote y Ubicación son correctos.
                If Not IsDBNull(dataOriginal.Traza) Then
                    Dim oV As New BdgVino
                    Dim dtVino As DataTable = oV.SelOnPrimaryKey(dataOriginal.Traza)
                    If Not IsNothing(dtVino) AndAlso dtVino.Rows.Count > 0 Then
                        If dtVino.Rows(0)(_V.IDArticulo) = dataOriginal.Articulo And dtVino.Rows(0)(_V.IDDeposito) = dataOriginal.Ubicacion _
                                            And dtVino.Rows(0)(_V.Lote) = dataOriginal.Lote And dtVino.Rows(0)(_V.IDAlmacen) = dataOriginal.Almacen Then
                            dataVQ = New ProcesoStocks.DataVinoQ(dataOriginal.Articulo, dataOriginal.Ubicacion, dataOriginal.Lote, dataOriginal.Almacen)
                            dataVQ.IDVino = dataOriginal.Traza
                        End If
                    End If
                End If
            End If

            If dataVQ Is Nothing OrElse IsDBNull(dataVQ.IDVino) Then
                'Si no encontramos ningún vino lo creamos pero perdemos la trazabilidad y los costes
                Dim StCrear As New BdgWorkClass.StCrearVinoOrigenCantidad(dataOriginal.Ubicacion, dataOriginal.Articulo, dataOriginal.Lote, dataOriginal.FechaDocumento, BdgOrigenVino.Interno, -dataOriginal.Cantidad)
                StCrear.IDAlmacen = dataOriginal.Almacen
                ProcessServer.ExecuteTask(Of BdgWorkClass.StCrearVinoOrigenCantidad)(AddressOf BdgWorkClass.CrearVinoOrigenCantidad, StCrear, Services)
                dataVQ = New ProcesoStocks.DataVinoQ(dataOriginal.Articulo, dataOriginal.Ubicacion, dataOriginal.Lote, dataOriginal.Almacen)
                dataVQ = ProcessServer.ExecuteTask(Of ProcesoStocks.DataVinoQ, ProcesoStocks.DataVinoQ)(AddressOf BdgWorkClass.GetIDVinoCantidad, dataVQ, Services)
            Else
                If dataOriginal.TipoMovimiento = enumTipoMovimiento.tmEntTransferencia Then
                    If dataVQ.Cantidad = dataOriginal.Cantidad Then
                        Dim StDeshacer As New BdgWorkClass.StDeshacerVino(dataVQ.IDVino, strOperacion, String.Empty, dataOriginal.Cantidad)
                        ProcessServer.ExecuteTask(Of BdgWorkClass.StDeshacerVino)(AddressOf BdgWorkClass.DeshacerVino, StDeshacer, Services)
                    Else
                        Dim StCambio As New BdgWorkClass.StCambiarOcupacion(dataVQ.IDVino, Cantidad)
                        ProcessServer.ExecuteTask(Of BdgWorkClass.StCambiarOcupacion)(AddressOf BdgWorkClass.CambiarOcupacion, StCambio, Services)
                        Dim StSel As New BdgVinoEstructura.StSelOnVinoOperacion(dataVQ.IDVino, strOperacion)
                        Dim dtVE As DataTable = ProcessServer.ExecuteTask(Of BdgVinoEstructura.StSelOnVinoOperacion, DataTable)(AddressOf BdgVinoEstructura.SelOnVinoOperacion, StSel, Services)
                        If Not dtVE Is Nothing AndAlso dtVE.Rows.Count > 0 Then
                            Dim NuevaCantidad As Double = dtVE.Rows(0)("Cantidad") - dataOriginal.Cantidad
                            If NuevaCantidad = 0 Then
                                Dim StDeshacer As New BdgWorkClass.StDeshaceEstructura(dataVQ.IDVino, strOperacion)
                                ProcessServer.ExecuteTask(Of BdgWorkClass.StDeshaceEstructura)(AddressOf BdgWorkClass.DeshacerEstructura, StDeshacer, Services)
                            Else
                                Dim StCrearEst As New BdgWorkClass.StCrearEstructuraOperacion(dataVQ.IDVino, BdgOrigenVino.AlbaranTransferencia, strOperacion, _
                                                    New VinoComponente() {New VinoComponente(dtVE.Rows(0)("IDVinoComponente"), NuevaCantidad, 1)})
                                ProcessServer.ExecuteTask(Of BdgWorkClass.StCrearEstructuraOperacion)(AddressOf BdgWorkClass.CrearEstructuraOperacion, StCrearEst, Services)
                            End If
                        End If
                    End If
                Else
                    Dim StCambio As New BdgWorkClass.StCambiarOcupacion(dataVQ.IDVino, Cantidad)
                    ProcessServer.ExecuteTask(Of BdgWorkClass.StCambiarOcupacion)(AddressOf BdgWorkClass.CambiarOcupacion, StCambio, Services)
                End If
            End If
            Return dataVQ
        End If
    End Function

    Public Function SincronizarCorreccionMovimiento(ByVal NumeroMovimiento As Integer, ByVal data As Negocio.StockData, ByVal dataOriginal As Negocio.StockData) As Negocio.ProcesoStocks.DataVinoQ Implements Negocio.IStock.SincronizarCorreccionMovimiento
        If Not data Is Nothing Then
            If data.Cantidad <> dataOriginal.Cantidad Then
                Dim Services As New ServiceProvider
                Dim dataVQ As New ProcesoStocks.DataVinoQ(data.Articulo, data.Ubicacion, data.Lote, data.Almacen)
                dataVQ = ProcessServer.ExecuteTask(Of ProcesoStocks.DataVinoQ, ProcesoStocks.DataVinoQ)(AddressOf BdgWorkClass.GetIDVinoCantidad, dataVQ, Services)
                If Not dataVQ Is Nothing Then
                    Dim StCambiar As New BdgWorkClass.StCambiarOcupacion(dataVQ.IDVino, -(data.Cantidad - dataOriginal.Cantidad))
                    ProcessServer.ExecuteTask(Of BdgWorkClass.StCambiarOcupacion)(AddressOf BdgWorkClass.CambiarOcupacion, StCambiar, Services)
                End If
                Return dataVQ
            End If
        End If
    End Function

    Public Function SincronizarAjuste(ByVal NumeroMovimiento As Integer, ByVal data As Negocio.StockData) As Negocio.ProcesoStocks.DataVinoQ Implements Negocio.IStock.SincronizarAjuste
        'Al hacer ajuste sólo vienen los lotes a los que se ha modificado su stock.
        If Not data Is Nothing Then
            Dim Services As New ServiceProvider
            Dim oWC As New BdgWorkClass
            Dim dataVQ As New ProcesoStocks.DataVinoQ(data.Articulo, data.Ubicacion, data.Lote, data.Almacen)
            dataVQ = ProcessServer.ExecuteTask(Of ProcesoStocks.DataVinoQ, ProcesoStocks.DataVinoQ)(AddressOf BdgWorkClass.GetIDVinoCantidad, dataVQ, Services)
            If dataVQ Is Nothing Then
                Dim StCrear As New BdgWorkClass.StCrearVinoOrigenCantidad(data.Ubicacion, data.Articulo, data.Lote, data.FechaDocumento, BdgOrigenVino.Interno, data.Cantidad)
                StCrear.IDAlmacen = data.Almacen
                ProcessServer.ExecuteTask(Of BdgWorkClass.StCrearVinoOrigenCantidad, Guid)(AddressOf BdgWorkClass.CrearVinoOrigenCantidad, StCrear, Services)
                dataVQ = New ProcesoStocks.DataVinoQ(data.Articulo, data.Ubicacion, data.Lote, data.Almacen)
                dataVQ = ProcessServer.ExecuteTask(Of ProcesoStocks.DataVinoQ, ProcesoStocks.DataVinoQ)(AddressOf BdgWorkClass.GetIDVinoCantidad, dataVQ, Services)
            Else
                Dim StCambiar As New BdgWorkClass.StCambiarOcupacion(dataVQ.IDVino, data.Cantidad - dataVQ.Cantidad)
                ProcessServer.ExecuteTask(Of BdgWorkClass.StCambiarOcupacion)(AddressOf BdgWorkClass.CambiarOcupacion, StCambiar, Services)
            End If
            Return dataVQ
        End If
    End Function

    Private Function GetIDVinoParaSalida(ByVal IDArticulo As String, ByVal IDDeposito As String, ByVal Lote As String, ByVal IDAlmacen As String) As Guid
        Dim Services As New ServiceProvider
        Dim dataVQ As New ProcesoStocks.DataVinoQ(IDArticulo, IDDeposito, Lote, IDAlmacen)
        dataVQ = ProcessServer.ExecuteTask(Of ProcesoStocks.DataVinoQ, ProcesoStocks.DataVinoQ)(AddressOf BdgWorkClass.GetIDVinoCantidad, dataVQ, Services)

        If dataVQ Is Nothing Then ApplicationService.GenerateError("Situación inesperada: no existe un vino para el Artículo/Depósito/Lote/Almacén")
        Return dataVQ.IDVino
    End Function

    Public Function AltaEntradaVino(ByVal data As Negocio.DataAltEntVino) As Integer Implements Negocio.IStock.AltaEntradaVino
        Dim ClsEntVino As New BdgEntradaVino
        Dim DtEntVino As DataTable = ClsEntVino.AddNewForm
        DtEntVino.Rows(0)("IDArticulo") = data.IDArticulo
        DtEntVino.Rows(0)("IDProveedor") = data.IDProveedor
        DtEntVino.Rows(0)("Precio") = data.Precio
        DtEntVino.Rows(0)("Cantidad") = data.Cantidad
        DtEntVino.Rows(0)("Lote") = data.Lote
        DtEntVino.Rows(0)("Fecha") = data.Fecha
        DtEntVino.Rows(0)("NDaa") = data.NDaa
        DtEntVino.Rows(0)("AadReferenceCode") = data.ARC

        Dim IDAnalisis As String = New BdgParametro().AnalisisEntradaVino

        Dim oDE As New DataEngine
        Dim DtVariables As DataTable = oDE.Filter("frmBdgAnalisisVariable", New FilterItem("IDAnalisis", FilterOperator.Equal, IDAnalisis, FilterType.String))
        Dim DtAnalisis As DataTable = New BdgEntradaVinoAnalisis().AddNew
        For Each DrVar As DataRow In DtVariables.Select
            Dim DrNew As DataRow = DtAnalisis.NewRow
            DrNew("NEntrada") = DtEntVino.Rows(0)("NEntrada")
            DrNew("IDVariable") = DrVar("IDVariable")
            DrNew("Valor") = 0
            DrNew("ValorNumerico") = 0
            DtAnalisis.Rows.Add(DrNew)
        Next

        Dim UpdtPckg As New UpdatePackage
        UpdtPckg.Add(DtEntVino)
        UpdtPckg.Add(DtAnalisis)
        ClsEntVino.Update(UpdtPckg)
        Return DtEntVino.Rows(0)("NEntrada")
    End Function

    Public Sub ActualizarPrecioEntradaVino(ByVal data As Negocio.DataPrecioEntVino) Implements Negocio.IStock.ActualizarPrecioEntradaVino
        Dim ClsEntVino As New BdgEntradaVino
        Dim DtEnt As DataTable = ClsEntVino.Filter(New FilterItem("NEntrada", FilterOperator.Equal, data.NEntrada))
        If Not DtEnt Is Nothing AndAlso DtEnt.Rows.Count > 0 Then
            DtEnt.Rows(0)("Precio") = data.Precio
            ClsEntVino.Update(DtEnt)
        End If
    End Sub

    Public Sub ActualizarDAAARCEntradaVino(ByVal data As Negocio.DataActDAAARCEntVino) Implements Negocio.IStock.ActualizarDAAARCEntradaVino
        Dim ClsEntVino As New BdgEntradaVino
        Dim DtEnt As DataTable = ClsEntVino.Filter(New FilterItem("NEntrada", FilterOperator.Equal, data.NEntrada))
        If Not DtEnt Is Nothing AndAlso DtEnt.Rows.Count > 0 Then
            DtEnt.Rows(0)("NDaa") = data.NDaa
            DtEnt.Rows(0)("AadReferenceCode") = data.ARC
            ClsEntVino.Update(DtEnt)
        End If
    End Sub

    Public Sub ActualizarFechaEntradaVino(ByVal data As Negocio.DataFechaEntVino) Implements Negocio.IStock.ActualizarFechaEntradaVino
        Dim ClsEntVino As New BdgEntradaVino
        Dim DtEnt As DataTable = ClsEntVino.Filter(New FilterItem("NEntrada", FilterOperator.Equal, data.NEntrada))
        If Not DtEnt Is Nothing AndAlso DtEnt.Rows.Count > 0 Then
            DtEnt.Rows(0)("Fecha") = data.Fecha
            ClsEntVino.Update(DtEnt)
        End If
    End Sub

    Public Function CrearOperacionArticulosCompatibles(ByVal data As DataArtCompatiblesExp) As String Implements Negocio.IStock.CrearOperacionArticulosCompatibles
        Dim lstArtComp As New List(Of DataArtCompatiblesExp)
        lstArtComp.Add(data)
        Dim datPropuesta As New DataPrcPropuestaOperacionExpediciones(enumBdgOrigenOperacion.Real, OrigenOperacion.Expediciones, lstArtComp, , True)
        Dim BEDataEngine As New BE.DataEngine
        Dim rslt As DataPrcPropuestaOperacionResultLog = BEDataEngine.RunProcess(GetType(PrcPropuestaOperacion), datPropuesta)
        If Not rslt Is Nothing Then
            If Not rslt.logPropuesta.CreatedElements Is Nothing AndAlso rslt.logPropuesta.CreatedElements.Count > 0 Then
                Return rslt.logPropuesta.CreatedElements(0).NElement
            Else
                If Not rslt.logPropuesta.Errors Is Nothing AndAlso rslt.logPropuesta.Errors.Count > 0 Then
                    ApplicationService.GenerateError("No se ha podido generar el Stock para Expedir.", rslt.logPropuesta.Errors(0).MessageError)
                Else
                    ApplicationService.GenerateError("Ha ocurrido algún problema al generar la operación. No ha generado ningún resultado.")
                End If
            End If

        Else
            ApplicationService.GenerateError("Ha ocurrido algún problema al generar la operación. No ha generado ningún resultado.")
        End If
    End Function

End Class