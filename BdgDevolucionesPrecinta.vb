Public Class BdgDevolucionesPrecinta

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgDevolucionesPrecintas"

#End Region

#Region "Tasks públicas"

    <Serializable()> _
    Public Class stDevolucionesPrecintas
        Public Data As DataTable
        Public Fecha As Date
        Public AplicarSegundaUnidad As Boolean = False

        Public Sub New(ByVal dttorigen As DataTable, ByVal fecha As Date)
            Me.Data = dttorigen
            Me.Fecha = fecha
        End Sub
    End Class

    <Task()> _
    Public Shared Function DevolucionPrecintas(ByVal data As stDevolucionesPrecintas, ByVal services As ServiceProvider) As StockUpdateData()
        'DEVOLUCIÓN
        Dim bdgDev As New BdgDevolucionesPrecinta
        If Not data Is Nothing AndAlso data.Data.Rows.Count > 0 Then
            Dim dttUpdate As DataTable = bdgDev.AddNew
            For Each dr As DataRow In data.Data.Rows
                Dim dtrUpdate As DataRow = dttUpdate.NewRow
                dtrUpdate("IDDevolucionPrecinta") = Guid.NewGuid
                dtrUpdate("IDArticulo") = dr("IDArticulo")
                dtrUpdate("Lote") = dr("Lote")
                dtrUpdate("IDAlmacen") = dr("IDAlmacen")
                dtrUpdate("Ubicacion") = dr("Ubicacion")
                'TODO - ¿ALMACENAR Nº MOVIMIENTO DE AJUSTE?
                dtrUpdate("SeriePrecinta") = dr("SeriePrecinta")
                dtrUpdate("NDesdePrecinta") = dr("NDesdePrecinta")
                dtrUpdate("NHastaPrecinta") = dr("NHastaPrecinta")
                dtrUpdate("NDesdePrecintaUtilizada") = dr("NDesdePrecintaSerie")
                dtrUpdate("NHastaPrecintaUtilizada") = dr("NHastaPrecintaSerie")
                dtrUpdate("Cantidad") = dr("NHastaPrecinta") - dr("NDesdePrecinta") + 1
                dtrUpdate("Fecha") = data.Fecha
                dttUpdate.Rows.Add(dtrUpdate)
            Next
            bdgDev.Update(dttUpdate)
        End If

        'AJUSTE
        Return ProcessServer.ExecuteTask(Of stDevolucionesPrecintas, StockUpdateData())(AddressOf AjusteNegativo, data, services)

    End Function

    <Task()> _
    Public Shared Function AjusteNegativo(ByVal data As stDevolucionesPrecintas, ByVal services As ServiceProvider) As StockUpdateData()
        If (data.Data Is Nothing OrElse data.Data.Rows.Count = 0) Then Return Nothing
        Dim SegundaUnidad As Boolean
        Dim stck(-1) As StockData
        Dim bsnAAL As New ArticuloAlmacenLote
        Dim updateData(-1) As StockUpdateData
        For Each dtrOrigen As DataRow In data.Data.Rows
            Dim stckData As New StockData()
            stckData.Articulo = dtrOrigen("IDArticulo")
            stckData.Almacen = dtrOrigen("IDAlmacen")
            stckData.Ubicacion = dtrOrigen("Ubicacion")

            stckData.TipoMovimiento = enumTipoMovimiento.tmSalAlbaranVenta 'EXPEDICIONES CLIENTE
            stckData.Cantidad = dtrOrigen("StockFisico")
            If data.AplicarSegundaUnidad Then
                SegundaUnidad = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, dtrOrigen("IDArticulo"), services)
            End If
            If SegundaUnidad Then
                stckData.Cantidad2 = CDbl(dtrOrigen("Cantidad2"))
            End If
            stckData.PrecintaNSerie = dtrOrigen("SeriePrecinta")
            stckData.PrecintaDesde = dtrOrigen("NDesdePrecintaSerie")
            stckData.PrecintaHasta = dtrOrigen("NHastaPrecintaSerie")
            stckData.PrecintaUtilizadaDesde = dtrOrigen("NDesdePrecinta")
            stckData.PrecintaUtilizadaHasta = dtrOrigen("NHastaPrecinta")
            stckData.FechaDocumento = data.Fecha
            stckData.Lote = dtrOrigen("Lote")

            Dim nuevomov As Integer = ProcessServer.ExecuteTask(Of Object, Integer)(AddressOf ProcesoStocks.NuevoNumeroMovimiento, Nothing, services)
            Dim datacrearmov As New DataNumeroMovimientoSinc(nuevomov, stckData)
            Dim result As StockUpdateData = ProcessServer.ExecuteTask(Of DataNumeroMovimientoSinc, StockUpdateData)(AddressOf ProcesoStocks.Salida, datacrearmov, services)

            ReDim Preserve updateData(UBound(updateData) + 1)
            updateData(UBound(updateData)) = result
        Next
        Return updateData
    End Function

    <Serializable()> _
   Public Class stDeshacerDevoluciones
        Public IDDevolucionesPrecinta() As Guid
        Public Fecha As Date
        Public AplicarSegundaUnidad As Boolean = False

        Public Sub New(ByVal IDDevolucionPrecinta() As Guid, ByVal fecha As Date)
            Me.IDDevolucionesPrecinta = IDDevolucionPrecinta
            Me.Fecha = fecha
        End Sub
    End Class

    <Task()> _
    Public Shared Function DeshacerDevoluciones(ByVal data As stDeshacerDevoluciones, ByVal services As ServiceProvider) As StockUpdateData()
        Dim bdgDev As New BdgDevolucionesPrecinta
        Dim UpdateData(-1) As StockUpdateData

        For Each id As Guid In data.IDDevolucionesPrecinta
            Dim dttDevolucion As DataTable = bdgDev.SelOnPrimaryKey(id)
            If (Not dttDevolucion Is Nothing AndAlso dttDevolucion.Rows.Count > 0) Then
                Dim stckData As New StockData()
                stckData.Articulo = dttDevolucion.Rows(0)("IDArticulo")
                stckData.Almacen = dttDevolucion.Rows(0)("IDAlmacen")
                stckData.Lote = dttDevolucion.Rows(0)("Lote")
                stckData.Ubicacion = dttDevolucion.Rows(0)("Ubicacion")
                stckData.PrecintaNSerie = dttDevolucion.Rows(0)("SeriePrecinta")
                stckData.PrecintaDesde = dttDevolucion.Rows(0)("NDesdePrecintaUtilizada")
                stckData.PrecintaHasta = dttDevolucion.Rows(0)("NHastaPrecintaUtilizada")
                stckData.PrecintaUtilizadaDesde = dttDevolucion.Rows(0)("NDesdePrecinta")
                stckData.PrecintaUtilizadaHasta = dttDevolucion.Rows(0)("NHastaPrecinta")
                stckData.Cantidad = dttDevolucion.Rows(0)("Cantidad")
                stckData.TipoMovimiento = enumTipoMovimiento.tmEntAjuste '???????
                stckData.FechaDocumento = data.Fecha
                'If Data.AplicarSegundaUnidad Then
                '    SegundaUnidad = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, dtrOrigen("IDArticulo"), services)
                'End If
                'If SegundaUnidad Then
                '    stckData.Cantidad2 = CDbl(dtrOrigen("Cantidad2"))
                'End If

                Dim nuevomov As Integer = ProcessServer.ExecuteTask(Of Object, Integer)(AddressOf ProcesoStocks.NuevoNumeroMovimiento, Nothing, services)
                Dim datacrearmov As New DataNumeroMovimientoSinc(nuevomov, stckData)
                Dim result As StockUpdateData = ProcessServer.ExecuteTask(Of DataNumeroMovimientoSinc, StockUpdateData)(AddressOf ProcesoStocks.Entrada, datacrearmov, services)

                ReDim Preserve UpdateData(UBound(UpdateData) + 1)
                UpdateData(UBound(UpdateData)) = result

                'y ahora, borramos el registro
                bdgDev.Delete(dttDevolucion, services)
            End If
        Next


        Return UpdateData

    End Function

#End Region

End Class
