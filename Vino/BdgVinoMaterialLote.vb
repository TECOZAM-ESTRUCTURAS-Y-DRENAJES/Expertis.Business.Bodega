Public Class BdgVinoMaterialLote

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgVinoMaterialLote"

#End Region


#Region " RegisterAddNewTasks "

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        data(_BdgVinoMaterialLote.IDVinoMaterialLote) = Guid.NewGuid
    End Sub

#End Region

#Region " RegisterDeleteTasks "

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf DeshacerMovimientos)
    End Sub

    <Task()> Public Shared Sub DeshacerMovimientos(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Not data Is Nothing Then
            If Length(data(_BdgVinoMaterialLote.IDLineaMovimiento)) > 0 Then
                Dim dataCorreccion As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Eliminar, data(_BdgVinoMaterialLote.IDLineaMovimiento), False)
                ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento)(AddressOf ProcesoStocks.ActualizarMovimiento, dataCorreccion, services)
            End If
        End If
    End Sub

#End Region

#Region "Register Validate Task"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        ' validateProcess.AddTask(Of DataTable)(AddressOf ValidarDatosObligatorios)
    End Sub

    <Task()> _
    Public Shared Function ValidarDatosObligatorios(ByVal dttDesglose As DataTable, ByVal services As ServiceProvider) As Boolean
        For Each dtr As DataRow In dttDesglose.Rows
            If Not ProcessServer.ExecuteTask(Of DataRow, Boolean)(AddressOf ValidarLineaPrecintasSalida, dtr, services) Then
                Return False
            End If
        Next
        Return True
    End Function

#End Region

#Region "Task"

    <Serializable()> _
    Public Class StCrearMovimientos
        Public NOperacion As String
        Public LotesMaterial As DataTable
        Public LineasMaterial As DataTable

        Public Sub New(ByVal nOperacion As String, ByVal lotesMaterial As DataTable, ByVal lineasMaterial As DataTable)
            Me.NOperacion = nOperacion
            Me.LotesMaterial = lotesMaterial
            Me.LineasMaterial = lineasMaterial
        End Sub

    End Class

    <Task()> Public Shared Sub CrearMovimientos(ByVal data As StCrearMovimientos, ByVal services As ServiceProvider)
        Dim almacenPredeterminado As String = String.Empty
        If (data.LotesMaterial Is Nothing) Then
            Return
        End If
        Dim bsnAML As New ArticuloAlmacenLote
        For Each VinoMaterialLoteRow As DataRow In data.LotesMaterial.Rows
            If Length(VinoMaterialLoteRow(_BdgVinoMaterialLote.IDVinoMaterialLote)) = 0 Then
                VinoMaterialLoteRow(_BdgVinoMaterialLote.IDVinoMaterialLote) = Guid.NewGuid
            End If
            If (Length(VinoMaterialLoteRow(_BdgVinoMaterialLote.IDLineaMovimiento)) = 0) Then
                Dim f As New Filter
                f.Add(_BdgVinoMaterialLote.IDVinoMaterial, VinoMaterialLoteRow(_BdgVinoMaterialLote.IDVinoMaterial))
                Dim VinoMaterialRows() As DataRow = data.LineasMaterial.Select(f.Compose(New AdoFilterComposer))
                If (VinoMaterialRows.Length > 0) Then
                    Dim VinoMaterialRow As DataRow = VinoMaterialRows(0)
                    Dim NumeroMovimiento As Integer = ProcessServer.ExecuteTask(Of Object, Integer)(AddressOf ProcesoStocks.NuevoNumeroMovimiento, Nothing, services)
                    Dim stock As New StockData
                    stock.Articulo = VinoMaterialRow(_VM.IDArticulo)
                    If (Not VinoMaterialRow.IsNull(_VM.IDAlmacen)) Then
                        stock.Almacen = VinoMaterialRow(_VM.IDAlmacen)
                    Else
                        If (Length(almacenPredeterminado) = 0) Then
                            almacenPredeterminado = New Parametro().AlmacenPredeterminado
                        End If
                        stock.Almacen = almacenPredeterminado
                    End If
                    stock.Cantidad = VinoMaterialLoteRow(_BdgVinoMaterialLote.Cantidad)
                    stock.Lote = VinoMaterialLoteRow(_BdgVinoMaterialLote.Lote)
                    stock.Ubicacion = VinoMaterialLoteRow(_BdgVinoMaterialLote.Ubicacion)
                    stock.TipoMovimiento = enumTipoMovimiento.tmSalFabrica
                    stock.PrecioA = Nz(VinoMaterialRow(_VM.Precio))
                    If (Length(VinoMaterialRow(_VM.NOperacion)) > 0) Then
                        stock.Documento = VinoMaterialRow(_VM.NOperacion)
                    Else
                        stock.Documento = data.NOperacion
                    End If
                    If (Not IsDate(VinoMaterialRow(_VM.Fecha)) OrElse VinoMaterialRow(_VM.Fecha) = cnMinDate) Then
                        stock.FechaDocumento = Today
                    Else
                        stock.FechaDocumento = New Date(CDate(VinoMaterialRow(_VM.Fecha)).Year, CDate(VinoMaterialRow(_VM.Fecha)).Month, CDate(VinoMaterialRow(_VM.Fecha)).Day)
                    End If

                    If (Length(VinoMaterialLoteRow(_BdgVinoMaterialLote.NDesdePrecinta)) > 0) Then _
                       stock.PrecintaUtilizadaDesde = VinoMaterialLoteRow(_BdgVinoMaterialLote.NDesdePrecinta)
                    If (Length(VinoMaterialLoteRow(_BdgVinoMaterialLote.NHastaPrecinta)) > 0) Then _
                        stock.PrecintaUtilizadaHasta = VinoMaterialLoteRow(_BdgVinoMaterialLote.NHastaPrecinta)
                    If (Length(VinoMaterialLoteRow(_BdgVinoMaterialLote.SeriePrecinta)) > 0) Then _
                        stock.PrecintaNSerie = VinoMaterialLoteRow(_BdgVinoMaterialLote.SeriePrecinta)
                    Dim fPrec As New Filter
                    fPrec.Add("Lote", VinoMaterialLoteRow(_BdgVinoMaterialLote.Lote))
                    fPrec.Add("IDAlmacen", stock.Almacen)
                    fPrec.Add("SeriePrecinta", stock.PrecintaNSerie)
                    fPrec.Add("IDArticulo", stock.Articulo)
                    Dim artalmlote As DataTable = bsnAML.Filter(fPrec)
                    If (Not artalmlote Is Nothing AndAlso artalmlote.Rows.Count > 0) Then
                        stock.PrecintaDesde = artalmlote.Rows(0)("NDesdePrecinta")
                        stock.PrecintaHasta = artalmlote.Rows(0)("NHastaPrecinta")
                    End If

                    Dim dataSalida As New DataNumeroMovimientoSinc(NumeroMovimiento, stock, False)
                    Dim updateData As StockUpdateData = ProcessServer.ExecuteTask(Of DataNumeroMovimientoSinc, StockUpdateData)(AddressOf ProcesoStocks.Salida, dataSalida, services)
                    If Not updateData Is Nothing Then
                        If updateData.Estado = EstadoStock.Actualizado Then
                            VinoMaterialLoteRow(_BdgVinoMaterialLote.IDLineaMovimiento) = updateData.IDLineaMovimiento 'VinoMaterialRow(_VM.IDLineaMovimiento) = updateData.IDLineaMovimiento
                            VinoMaterialRow(_VM.Precio) = updateData.StockData.PrecioA
                        ElseIf updateData.Estado = EstadoStock.NoActualizado Then
                            If Length(stock.Lote) > 0 Then
                                ApplicationService.GenerateError(updateData.Log & vbNewLine & "Lote relacionado: " & stock.Lote)
                            Else
                                ApplicationService.GenerateError(updateData.Log)
                            End If
                        End If
                    End If
                End If
            End If
        Next
        Dim bsnMe As New BdgVinoMaterialLote
        bsnMe.Update(data.LotesMaterial)
    End Sub

    <Serializable()> _
    Public Class StValidarDesglosePrecintas
        Public dttDesglose As DataTable

        Public Sub New(ByVal desglose As DataTable)
            Me.dttDesglose = desglose
        End Sub
    End Class

    <Task()> _
    Public Shared Function ValidarDesglosePrecintasSalida(ByVal data As StValidarDesglosePrecintas, ByVal services As ServiceProvider) As Boolean
        If (data Is Nothing OrElse data.dttDesglose.Rows.Count = 0) Then
            Return True
        End If
        Dim ValidarDesglosesCantidad As List(Of DataRow) = (From c In data.dttDesglose Where c.RowState <> DataRowState.Deleted AndAlso c("Cantidad") <> 0 Select c).ToList
        If Not ValidarDesglosesCantidad Is Nothing AndAlso ValidarDesglosesCantidad.Count > 0 Then
            For Each dtr As DataRow In ValidarDesglosesCantidad
                If Not ProcessServer.ExecuteTask(Of DataRow, Boolean)(AddressOf ValidarLineaPrecintasSalida, dtr, services) Then
                    Return False
                End If
            Next
        End If
        Return True
    End Function

    <Task()> _
    Public Shared Function ValidarLineaPrecintasSalida(ByVal dtr As DataRow, ByVal services As ServiceProvider) As Boolean
        Dim f As New Filter
        f.Add("IDArticulo", dtr("IDArticulo"))
        f.Add("IDAlmacen", dtr("IDAlmacen"))
        f.Add("Lote", dtr("Lote"))
        Dim data As New BdgHistoricoPrecintas.stObtenerRangosPrecinta(dtr("IDArticulo"), dtr("IDAlmacen"), dtr("SeriePrecinta") & String.Empty, String.Empty, String.Empty, dtr("Lote"), dtr("Ubicacion"), String.Empty, String.Empty, enumBoolean.Todos)

        Dim dttRangosLibres As DataTable = ProcessServer.ExecuteTask(Of BdgHistoricoPrecintas.stObtenerRangosPrecinta, DataTable) _
                                            (AddressOf BdgHistoricoPrecintas.ObtenerRangosLibresPrecinta, data, services)

        If (dttRangosLibres Is Nothing OrElse dttRangosLibres.Rows.Count = 0) Then
            ApplicationService.GenerateError(String.Format("No hay rangos libres para el artículo {0}", dtr("IDArticulo")))
            Return False

        End If

        For Each dtrLibre As DataRow In dttRangosLibres.Rows
            If (dtr("NDesde") >= dtrLibre("NDesdePrecinta") AndAlso dtr("NDesde") <= dtrLibre("NHastaPrecinta")) Then
                'si el desde está contenido en un rango, comprobaos que el hasta también lo está paa validarlo y devolver que se ha encontrado hueco
                If (dtr("NHasta") >= dtrLibre("NDesdePrecinta") AndAlso dtr("NHasta") <= dtrLibre("NHastaPrecinta")) Then
                    Return True 'encontrado rango libre
                End If
            End If
        Next
        ApplicationService.GenerateError(String.Format("No hay rangos libres para el artículo {0}", dtr("IDArticulo")))
        Return False
    End Function

#End Region

End Class


Public Class _BdgVinoMaterialLote
    Public Const IDVinoMaterialLote As String = "IDVinoMaterialLote"
    Public Const IDVinoMaterial As String = "IDVinoMaterial"
    Public Const Cantidad As String = "Cantidad"
    Public Const Lote As String = "Lote"
    Public Const Ubicacion As String = "Ubicacion"
    Public Const IDLineaMovimiento As String = "IDLineaMovimiento"
    Public Const SeriePrecinta As String = "SeriePrecinta"
    Public Const NDesdePrecinta As String = "NDesde"
    Public Const NHastaPrecinta As String = "NHasta"
End Class
