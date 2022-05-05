Imports System.Collections.Generic

Public Class BdgCosteVino

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgCosteVino"

#End Region

#Region "Funciones Públicas"

    Protected Shared cnColumnaFabrica As String = "Fabrica"
    Protected Shared TABLA_TMP_COSTE_VINO As String = "tmpBdgVinoQTot"

    <Task()> Public Shared Sub DeleteTablaQTotalCostes(ByVal queryId As Integer, ByVal services As ServiceProvider)
        AdminData.Execute("DELETE FROM " & TABLA_TMP_COSTE_VINO & " WHERE idquery = " & queryId)
    End Sub

    <Task()> Public Shared Function CalculoCosteBodega(ByVal VOArticulo As ArticuloVino, ByVal services As ServiceProvider) As ArticuloVino
        Dim FechaDesde As Date = "1/1/1900"
        Dim FechaHasta As Date = DateAdd(DateInterval.Day, 1, Date.Today)
        Dim StData As New StCalculoCosteBodegaFechas(VOArticulo, FechaDesde, FechaHasta)
        VOArticulo = ProcessServer.ExecuteTask(Of StCalculoCosteBodegaFechas, ArticuloVino)(AddressOf CalculoCosteBodegaFechas, StData, services)
        Return VOArticulo
    End Function

    <Serializable()> _
    Public Class StCalculoCosteBodegaFechas
        Public VOArticulo As ArticuloVino
        Public FechaDesde As Date
        Public FechaHasta As Date


        Public Sub New()
        End Sub

        Public Sub New(ByVal VOArticulo As ArticuloVino, ByVal FechaDesde As Date, ByVal FechaHasta As Date)
            Me.VOArticulo = VOArticulo
            Me.FechaDesde = FechaDesde
            Me.FechaHasta = FechaHasta
        End Sub
    End Class

    <Task()> Public Shared Function CalculoCosteBodegaFechas(ByVal data As StCalculoCosteBodegaFechas, ByVal services As ServiceProvider) As ArticuloVino
        Dim dtCosteVino As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf CrearTbCOsteVino, Nothing, services)
        Dim dCantidadAcum, dCosteAcum, dCantidadActualAcum As Double
        If Not data.VOArticulo.Articulo Is Nothing AndAlso data.VOArticulo.Articulo.Rows.Count > 0 Then

            Dim IDQueryCoste As Integer = New Random().[Next](Integer.MaxValue)
            ProcessServer.ExecuteTask(Of Integer)(AddressOf DeleteTablaQTotalCostes, IDQueryCoste, services)

            Dim ArticulosSeleccionados As List(Of DataRow) = (From c In data.VOArticulo.Articulo Select c).ToList
            For Each drArticulo As DataRow In ArticulosSeleccionados 'Por cada articulo seleccionado 
                Dim datVinosArt As New DataGetVinosArticulo(drArticulo("IDArticulo"), data.FechaDesde, data.FechaHasta, Nz(drArticulo(cnColumnaFabrica), False))
                Dim Vinos As IList(Of VinoArticulo) = ProcessServer.ExecuteTask(Of DataGetVinosArticulo, IList(Of VinoArticulo))(AddressOf GetVinosArticulo, datVinosArt, services)

                dCosteAcum = 0
                dCantidadAcum = 0
                dCantidadActualAcum = 0
                For Each v As VinoArticulo In Vinos
                    Dim datVinoCst As New DataGetVinoCoste(v.idVino, data.FechaDesde, data.FechaHasta, True, True)
                    datVinoCst.IDQueryCoste = IDQueryCoste
                    Dim CosteTot As IList(Of VinoCosteTotal) = ProcessServer.ExecuteTask(Of DataGetVinoCoste, IList(Of VinoCosteTotal))(AddressOf GetVinoCosteTotales, datVinoCst, services)
                    If Not CosteTot Is Nothing AndAlso CosteTot.Count > 0 Then
                        'Rellenar la tabla con los resultados                    
                        Dim drCosteVino As DataRow = dtCosteVino.NewRow
                        drCosteVino("IDArticulo") = drArticulo("IDArticulo")
                        drCosteVino("IDVino") = CosteTot(0).IDVinoInicial
                        drCosteVino("Fecha") = v.Fecha
                        drCosteVino("IDDeposito") = v.IDDeposito
                        drCosteVino("Lote") = v.Lote
                        drCosteVino("IDAlmacen") = v.IDAlmacen

                        drCosteVino("CantidadActual") = v.QActual
                        drCosteVino("Cantidad") = CosteTot(0).QTot
                        drCosteVino("Coste") = CosteTot(0).CosteTotalAcum
                        drCosteVino("CosteUnit") = CosteTot(0).CosteTotalUnit

                        drCosteVino("FechaUltCalc") = Date.Today
                        dtCosteVino.Rows.Add(drCosteVino)

                        dCosteAcum += CosteTot(0).CosteTotalAcum
                        dCantidadAcum += CosteTot(0).QTot
                        dCantidadActualAcum += v.QActual
                    End If
                Next

                'Por cada articulo su coste unit
                If dCantidadAcum > 0 AndAlso dCosteAcum > 0 Then
                    drArticulo("CosteUnit") = dCosteAcum / dCantidadAcum
                Else
                    drArticulo("CosteUnit") = 0
                End If
                drArticulo("Coste") = dCosteAcum
                drArticulo("Cantidad") = dCantidadAcum
                drArticulo("CantidadActual") = dCantidadActualAcum
                drArticulo("Fecha") = Date.Today
                drArticulo("ImpEstandar") = Nz(drArticulo("Cantidad"), 0) * Nz(drArticulo("PrecioEstandarA"), 0)
                drArticulo("DifImporte") = Nz(drArticulo("Coste"), 0) - Nz(drArticulo("ImpEstandar"), 0)
                drArticulo("DifPrecio") = Nz(drArticulo("CosteUnit"), 0) - Nz(drArticulo("PrecioEstandarA"), 0)
                If drArticulo("PrecioEstandarA") <> 0 Then
                    drArticulo("PorcenDifPrecio") = (drArticulo("DifPrecio")) * 100 / Nz(drArticulo("PrecioEstandarA"), 0)
                End If
            Next
            data.VOArticulo.ArticuloDesglose = dtCosteVino 'Pasar el desglose a presentacion.

            ProcessServer.ExecuteTask(Of Integer)(AddressOf DeleteTablaQTotalCostes, IDQueryCoste, services)
        End If
        Return data.VOArticulo
    End Function



    <Serializable()> _
    Public Class DataGetVinosArticulo
        Public IDArticulo As String
        Public FechaDesde As Date
        Public FechaHasta As Date
        Public Fabrica As Boolean

        Public Sub New(ByVal IDArticulo As String, ByVal FechaDesde As Date, ByVal FechaHasta As Date, ByVal Fabrica As Boolean)
            Me.IDArticulo = IDArticulo
            Me.FechaDesde = FechaDesde
            Me.FechaHasta = FechaHasta
            Me.Fabrica = Fabrica
        End Sub
    End Class

    <Task()> Public Shared Function GetVinosArticulo(ByVal data As DataGetVinosArticulo, ByVal services As ServiceProvider) As IList(Of VinoArticulo)
        Dim dataCostContext As New CostWineDataContext(AdminData.GetConnectionString)
        If data.Fabrica Then
            Dim Vinos As IList(Of VinoArticulo) = dataCostContext.GetVinosArticuloFabrica(data.IDArticulo, data.FechaDesde, data.FechaHasta).ToList
            Return Vinos
        Else
            Dim Vinos As IList(Of VinoArticulo) = dataCostContext.GetVinosArticulo(data.IDArticulo, data.FechaHasta).ToList
            Return Vinos
        End If
    End Function


    <Serializable()> _
   Public Class DataGetVinoCoste
        Public IDVino As Guid
        Public FechaDesde As Date
        Public FechaHasta As Date
        Public ConsiderarNoTrazar As Boolean
        Public ConsiderarCosteInicial As Boolean
        Public IDQueryCoste As Integer

        Public Sub New(ByVal IDVino As Guid, ByVal FechaDesde As Date, ByVal FechaHasta As Date, ByVal ConsiderarNoTrazar As Boolean, ByVal ConsiderarCosteInicial As Boolean, Optional ByVal IDQueryCoste As Integer = 0)
            Me.IDVino = IDVino
            Me.FechaDesde = FechaDesde
            Me.FechaHasta = FechaHasta
            Me.ConsiderarNoTrazar = ConsiderarNoTrazar
            Me.ConsiderarCosteInicial = ConsiderarCosteInicial
            Me.IDQueryCoste = IDQueryCoste
        End Sub
    End Class


    <Task()> Public Shared Function GetVinoCosteTotales(ByVal data As DataGetVinoCoste, ByVal services As ServiceProvider) As IList(Of VinoCosteTotal)
        Dim Costes As IList(Of VinoCosteTotal)
        Dim datExpVino As New DataExplosionVino(data.IDVino, data.FechaHasta, data.ConsiderarNoTrazar, data.ConsiderarCosteInicial)
        datExpVino.IDQueryCoste = data.IDQueryCoste
        Dim rslt As DataExplosionVinoResult = ProcessServer.ExecuteTask(Of DataExplosionVino, DataExplosionVinoResult)(AddressOf BdgExplosionVino.Explosion, datExpVino, services)
        If Not rslt Is Nothing AndAlso Not rslt.Vinos Is Nothing AndAlso Not rslt.VinosEstructura Is Nothing Then
            Dim lstVinos As IList(Of Vino) = rslt.Vinos
            Dim lstVinosEstructura As IList(Of VinoEstructura) = rslt.VinosEstructura

            Dim QDepositoVino As Double = Nz(lstVinos(0).QDepositoVino, 0)

            Dim datPorcTmp As New BdgExplosionVino.DataRellenarTablaPorcentajes(lstVinos, lstVinosEstructura, False, True)
            Dim queryId As Integer = ProcessServer.ExecuteTask(Of BdgExplosionVino.DataRellenarTablaPorcentajes, Integer)(AddressOf BdgExplosionVino.RellenarTablaPorcentajes, datPorcTmp, services)

            Dim dataCostContext As New CostWineDataContext(AdminData.GetConnectionString)
            Costes = dataCostContext.GetVinoCosteTotales(queryId, data.IDVino, True, data.FechaDesde, data.FechaHasta, data.ConsiderarNoTrazar, data.ConsiderarCosteInicial).ToList
            'Costes = dataCostContext.GetVinoCosteTotales(data.IDVino, True, data.FechaDesde, data.FechaHasta).ToList
            ProcessServer.ExecuteTask(Of Integer)(AddressOf BdgExplosionVino.DeleteTablaPorcentajes, queryId, services)
        End If
        Return Costes

        'Dim dataCostContext As New CostWineDataContext(AdminData.GetConnectionString)
        'Dim Costes As IList(Of VinoCosteTotal) = dataCostContext.GetVinoCosteTotales(data.IDVino, True, data.FechaDesde, data.FechaHasta).ToList
        'Return Costes
    End Function


    <Serializable()> _
    Public Class DataVinoCosteDetalle
        Public Materiales As IList(Of VinoCosteMateriales)
        Public ManoObra As IList(Of VinoCosteMOD)
        Public Centros As IList(Of VinoCosteCentros)
        Public Varios As IList(Of VinoCosteVarios)
        Public Compras As IList(Of VinoCosteCompras)
        Public Tasas As IList(Of VinoCosteTasas)
        Public VendimiaElaboracion As IList(Of VinoCosteVendimiaElaboracion)
        Public EntradaUVA As IList(Of VinoCosteEntradaUVA)
        Public EstanciaNave As IList(Of VinoCosteEstanciaEnNave)
        Public CosteInicial As IList(Of VinoCosteInicial)
        Public Totales As IList(Of VinoCosteTotal)
    End Class

    <Task()> Public Shared Function GetVinoCosteDetalle(ByVal data As DataGetVinoCoste, ByVal services As ServiceProvider) As DataVinoCosteDetalle
        Dim Detalle As New DataVinoCosteDetalle
        Dim datExpVino As New DataExplosionVino(data.IDVino, data.FechaHasta, data.ConsiderarNoTrazar, data.ConsiderarCosteInicial)
        Dim rslt As DataExplosionVinoResult = ProcessServer.ExecuteTask(Of DataExplosionVino, DataExplosionVinoResult)(AddressOf BdgExplosionVino.Explosion, datExpVino, services)
        If Not rslt Is Nothing AndAlso Not rslt.Vinos Is Nothing AndAlso Not rslt.VinosEstructura Is Nothing Then
            Dim lstVinos As IList(Of Vino) = rslt.Vinos
            Dim lstVinosEstructura As IList(Of VinoEstructura) = rslt.VinosEstructura

            Dim QDepositoVino As Double = Nz(lstVinos(0).QDepositoVino, 0)

            Dim datPorcTmp As New BdgExplosionVino.DataRellenarTablaPorcentajes(lstVinos, lstVinosEstructura, False, True)
            Dim queryId As Integer = ProcessServer.ExecuteTask(Of BdgExplosionVino.DataRellenarTablaPorcentajes, Integer)(AddressOf BdgExplosionVino.RellenarTablaPorcentajes, datPorcTmp, services)

            Dim dataCostContext As New CostWineDataContext(AdminData.GetConnectionString)
            Dim cstDet As IMultipleResults = dataCostContext.GetVinoCosteDetalle(queryId, data.IDVino, False, data.FechaDesde, data.FechaHasta, data.ConsiderarNoTrazar, data.ConsiderarCosteInicial)
            If Not cstDet Is Nothing Then
                Detalle.Materiales = cstDet.GetResult(Of VinoCosteMateriales).ToList
                Detalle.ManoObra = cstDet.GetResult(Of VinoCosteMOD).ToList
                Detalle.Centros = cstDet.GetResult(Of VinoCosteCentros).ToList
                Detalle.Varios = cstDet.GetResult(Of VinoCosteVarios).ToList
                Detalle.Compras = cstDet.GetResult(Of VinoCosteCompras).ToList
                Detalle.Tasas = cstDet.GetResult(Of VinoCosteTasas).ToList
                Detalle.VendimiaElaboracion = cstDet.GetResult(Of VinoCosteVendimiaElaboracion).ToList
                Detalle.EntradaUVA = cstDet.GetResult(Of VinoCosteEntradaUVA).ToList
                Detalle.EstanciaNave = cstDet.GetResult(Of VinoCosteEstanciaEnNave).ToList
                Detalle.CosteInicial = cstDet.GetResult(Of VinoCosteInicial).ToList
                Detalle.Totales = cstDet.GetResult(Of VinoCosteTotal).ToList
            End If

            ProcessServer.ExecuteTask(Of Integer)(AddressOf BdgExplosionVino.DeleteTablaPorcentajes, queryId, services)
        End If
        Return Detalle
    End Function


    '<Serializable()> _
    '    Public Class DataActualizarCosteEstandar
    '    Public dtArticulos As DataTable
    '    Public Fecha As Date
    '    Public Esquema As String

    '    Public Sub New(ByVal dtArticulos As DataTable, ByVal Fecha As Date, ByVal Esquema As String)
    '        Me.dtArticulos = dtArticulos
    '        Me.Fecha = Fecha
    '        Me.Esquema = Esquema
    '    End Sub
    'End Class

    '<Serializable()> _
    'Public Class DataActualizarCosteEstandarResult
    '    Public ArticulosActualizados As List(Of String)
    '    Public Errores As List(Of ClassErrors)

    '    Public Sub New()

    '    End Sub

    '    Public Sub New(ByVal ArticulosActualizados As List(Of String), ByVal Errores As List(Of ClassErrors))
    '        Me.ArticulosActualizados = ArticulosActualizados
    '        Me.Errores = Errores
    '    End Sub
    'End Class

    '<Task()> Public Shared Function ActualizacionCosteEstandar(ByVal data As DataActualizarCosteEstandar, ByVal services As ServiceProvider) As DataActualizarCosteEstandarResult
    '    Dim rslt As New DataActualizarCosteEstandarResult
    '    Dim IDs As String = String.Empty
    '    If Not IsNothing(data.dtArticulos) AndAlso data.dtArticulos.Rows.Count > 0 Then
    '        Dim ff As New Filter
    '        Dim dtArtCompleto As DataTable
    '        Dim clsArticulo As New Articulo
    '        Dim pTipoMov As Integer? = New BdgParametro().TipoMovimiento0
    '        For Each drArticulo As DataRow In data.dtArticulos.Select
    '            Try
    '                ff.Clear()
    '                ff.Add("IDArticulo", FilterOperator.Equal, drArticulo("IDArticulo"))
    '                dtArtCompleto = clsArticulo.Filter(ff)
    '                If Not IsNothing(dtArtCompleto) AndAlso dtArtCompleto.Rows.Count > 0 Then
    '                    IDs = IDs & " - " & drArticulo("IDArticulo")
    '                    dtArtCompleto.Rows(0)("PrecioEstandarA") = Nz(drArticulo("CosteUnit"), 0)
    '                    clsArticulo.ApplyBusinessRule("PrecioEstandarA", dtArtCompleto.Rows(0)("PrecioEstandarA"), dtArtCompleto.Rows(0))
    '                    dtArtCompleto.Rows(0)("FechaEstandar") = Date.Today

    '                    ProcessServer.ExecuteTask(Of Object)(AddressOf Comunes.BeginTransaction, Nothing, services)
    '                    If Not pTipoMov Is Nothing AndAlso pTipoMov <> 0 Then
    '                        Dim stData As New DataGenerarMovimiento(pTipoMov, data.Esquema, dtArtCompleto.Rows(0)("IDArticulo"), data.Fecha, dtArtCompleto.Rows(0)("PrecioEstandarA"), dtArtCompleto.Rows(0)("PrecioEstandarB"))
    '                        ProcessServer.ExecuteTask(Of DataGenerarMovimiento)(AddressOf GenerarMovimiento, stData, services)
    '                    End If

    '                    clsArticulo.Update(dtArtCompleto)

    '                    ProcessServer.ExecuteTask(Of Boolean)(AddressOf Business.General.Comunes.CommitTransaction, False, services)

    '                    If rslt.ArticulosActualizados Is Nothing Then rslt.ArticulosActualizados = New List(Of String)
    '                    rslt.ArticulosActualizados.Add(drArticulo("IDArticulo"))
    '                End If
    '            Catch ex As Exception
    '                If rslt.Errores Is Nothing Then rslt.Errores = New List(Of ClassErrors)
    '                Dim err As New ClassErrors(drArticulo("IDArticulo"), ex.Message)
    '                rslt.Errores.Add(err)
    '            End Try
    '        Next
    '    End If
    '    Return rslt
    'End Function


#Region " Actualizar Coste Estandar "

    <Serializable()> _
    Public Class DataActualizarCosteEstandar
        Public dtArticulos As DataTable
        Public Fecha As Date

        Public Sub New(ByVal dtArticulos As DataTable, ByVal Fecha As Date)
            Me.dtArticulos = dtArticulos
            Me.Fecha = Fecha
        End Sub
    End Class

    <Serializable()> _
    Public Class DataActualizarCosteEstandarResult
        Public ArticulosActualizados As List(Of String)
        Public Errores As List(Of ClassErrors)

        Public Sub New()

        End Sub

        Public Sub New(ByVal ArticulosActualizados As List(Of String), ByVal Errores As List(Of ClassErrors))
            Me.ArticulosActualizados = ArticulosActualizados
            Me.Errores = Errores
        End Sub
    End Class

    <Task()> Public Shared Function ActualizacionCosteEstandar(ByVal data As DataActualizarCosteEstandar, ByVal services As ServiceProvider) As DataActualizarCosteEstandarResult
        Dim rslt As New DataActualizarCosteEstandarResult
        Dim Esquema As String = ProcessServer.ExecuteTask(Of Object, String)(AddressOf Comunes.GetEsquemaBD, Nothing, services)
        If Not data.dtArticulos Is Nothing AndAlso data.dtArticulos.Rows.Count > 0 Then
            Dim ff As New Filter

            Dim Art As New Articulo
            Dim AppParams As ParametroStocks = services.GetService(Of ParametroStocks)()
            Dim FechaEstandar As Date = data.Fecha
            For Each drArticulo As DataRow In data.dtArticulos.Select
                Try
                    ff.Clear()
                    ff.Add("IDArticulo", FilterOperator.Equal, drArticulo("IDArticulo"))
                    Dim dtArticulo As DataTable = Art.Filter(ff)
                    If dtArticulo.Rows.Count > 0 Then
                        dtArticulo.Rows(0)("PrecioEstandarA") = Nz(drArticulo("CosteUnit"), 0)
                        dtArticulo.Rows(0)("PrecioEstandarB") = ProcessServer.ExecuteTask(Of Double, Double)(AddressOf Articulo.CalcularImporteEnMonedaB, dtArticulo.Rows(0)("PrecioEstandarA"), services)
                        dtArticulo.Rows(0)("FechaEstandar") = FechaEstandar

                        AdminData.BeginTx()
                        Art.Update(dtArticulo)
                        If AppParams.TipoMovimientoCantidad0 <> 0 AndAlso AppParams.TipoMovimientoCantidad0 > enumTipoMovimiento.tmSalContraActivos Then
                            Dim ProcInfo As ArticuloCosteEstandar.ProcInfoActualizarPrecioEstandar = services.GetService(Of ArticuloCosteEstandar.ProcInfoActualizarPrecioEstandar)()
                            ProcInfo.PermitirMovtoCantidad0 = True
                            ProcInfo.RecalcularPrecioStdPosteriores = True
                            Dim stData As New ArticuloCosteEstandar.DataGenerarMovimiento(Esquema, dtArticulo.Rows(0)("IDArticulo"), FechaEstandar, dtArticulo.Rows(0)("PrecioEstandarA"), dtArticulo.Rows(0)("PrecioEstandarB"))
                            ProcessServer.ExecuteTask(Of ArticuloCosteEstandar.DataGenerarMovimiento)(AddressOf ArticuloCosteEstandar.GenerarMovimiento, stData, services)
                        End If


                        AdminData.CommitTx(True)

                        If rslt.ArticulosActualizados Is Nothing Then rslt.ArticulosActualizados = New List(Of String)
                        rslt.ArticulosActualizados.Add(drArticulo("IDArticulo"))
                    End If
                Catch ex As Exception
                    AdminData.RollBackTx()
                    If rslt.Errores Is Nothing Then rslt.Errores = New List(Of ClassErrors)
                    Dim err As New ClassErrors(drArticulo("IDArticulo"), ex.Message)
                    rslt.Errores.Add(err)
                End Try
            Next
        End If
        Return rslt
    End Function



#End Region


    <Task()> Public Shared Function CrearTbCOsteVino(ByVal data As Object, ByVal services As ServiceProvider) As DataTable
        Dim dt As New DataTable
        dt.Columns.Add("IDArticulo", GetType(String))
        dt.Columns.Add("IDVino", GetType(System.Guid))
        dt.Columns.Add("Fecha", GetType(Date))
        dt.Columns.Add("Cantidad", GetType(Double))
        dt.Columns.Add("Coste", GetType(Double))
        dt.Columns.Add("CosteUnit", GetType(Double))
        dt.Columns.Add("FechaUltCalc", GetType(Date))
        dt.Columns.Add("IDDeposito", GetType(String))
        dt.Columns.Add("Lote", GetType(String))
        dt.Columns.Add("IDAlmacen", GetType(String))
        dt.Columns.Add("CantidadActual", GetType(Double))
        Return dt
    End Function

    <Serializable()> _
    Public Class DataActTasaCoste
        Public Origen As String
        Public FechaDesde As DateTime
        Public FechaHasta As DateTime

        Public Sub New()
        End Sub
        Public Sub New(ByVal Origen As String, ByVal FechaDesde As DateTime, ByVal FechaHasta As DateTime)
            Me.Origen = Origen
            Me.FechaDesde = FechaDesde
            Me.FechaHasta = FechaHasta
        End Sub
    End Class

    <Task()> Public Shared Sub ActualizarTasaCoste(ByVal data As DataActTasaCoste, ByVal services As ServiceProvider)
        AdminData.BeginTx()
        Select Case data.Origen
            Case "Centros"
                Dim VC As New BdgVinoCentro
                Dim FilCentro As New Filter
                FilCentro.Add("Fecha", FilterOperator.GreaterThanOrEqual, data.FechaDesde)
                FilCentro.Add("Fecha", FilterOperator.LessThanOrEqual, data.FechaHasta)
                Dim DtVinoCentro As DataTable = VC.Filter(FilCentro)
                If Not DtVinoCentro Is Nothing AndAlso DtVinoCentro.Rows.Count > 0 Then
                    Dim TasasCentros As List(Of DataRow) = (From c In DtVinoCentro Order By c("IDCentro"), c("Fecha") Select c).ToList
                    For Each drVinoCentro As DataRow In TasasCentros
                        Dim dtVinoCentroAct As DataTable = DtVinoCentro.Clone
                        Dim StGet As New BdgVinoCentro.StGetTasa(drVinoCentro("IDCentro"), drVinoCentro("Fecha"), drVinoCentro("IDIncidencia") & String.Empty)
                        Dim oTasa As TasaInfo = ProcessServer.ExecuteTask(Of BdgVinoCentro.StGetTasa, TasaInfo)(AddressOf BdgVinoCentro.GetTasa, StGet, services)
                        drVinoCentro("Tasa") = oTasa.Tasa
                        drVinoCentro(_VC.UDTiempo) = oTasa.UdTiempo
                        drVinoCentro(_VC.TasaD) = oTasa.TasaD
                        drVinoCentro(_VC.TasaI) = oTasa.TasaI
                        drVinoCentro(_VC.TasaF) = oTasa.TasaF
                        drVinoCentro(_VC.TasaV) = oTasa.TasaV
                        drVinoCentro(_VC.TasaFiscal) = oTasa.TasaFscl
                        dtVinoCentroAct.ImportRow(drVinoCentro)
                        VC.Update(dtVinoCentroAct)

                        Dim StCrear As New BdgVinoCentroTasa.StCrearVinoCentroTasa(drVinoCentro("Fecha"), dtVinoCentroAct)
                        ProcessServer.ExecuteTask(Of BdgVinoCentroTasa.StCrearVinoCentroTasa)(AddressOf BdgVinoCentroTasa.CrearVinoCentroTasa, StCrear, services)
                    Next

                End If
            Case "Operarios"
                Dim ClsVinoMOD As New BdgVinoMod
                Dim FilMOD As New Filter
                FilMOD.Add("Fecha", FilterOperator.GreaterThanOrEqual, data.FechaDesde)
                FilMOD.Add("Fecha", FilterOperator.LessThanOrEqual, data.FechaHasta)
                Dim DtVinoMOD As DataTable = ClsVinoMOD.Filter(FilMOD)
                If Not DtVinoMOD Is Nothing AndAlso DtVinoMOD.Rows.Count > 0 Then
                    For Each DrVinoMOD As DataRow In DtVinoMOD.Select("", "IDOperario, Fecha")
                        Dim StData As New HoraCategoria.DatosPrecioHoraCatOper(DrVinoMOD("IDCategoria"), DrVinoMOD("IDHora"), DrVinoMOD("Fecha"), DrVinoMOD("IDOperario"))
                        Dim DblTasa As Double = ProcessServer.ExecuteTask(Of HoraCategoria.DatosPrecioHoraCatOper, Double)(AddressOf HoraCategoria.ObtenerPrecioHora, StData, services)
                        DrVinoMOD("Tasa") = DblTasa
                    Next
                    ClsVinoMOD.Update(DtVinoMOD)
                End If
            Case "Materiales"
                Dim ClsVinoMat As New BdgVinoMaterial
                Dim FilMat As New Filter
                FilMat.Add("Fecha", FilterOperator.GreaterThanOrEqual, data.FechaDesde)
                FilMat.Add("Fecha", FilterOperator.LessThanOrEqual, data.FechaHasta)
                Dim DtVinoMat As DataTable = New BE.DataEngine().Filter("vfrmBdgVinoMaterialCosteActualizar", FilMat)
                If Not DtVinoMat Is Nothing AndAlso DtVinoMat.Rows.Count > 0 Then
                    For Each DrVinoMat As DataRow In DtVinoMat.Select("", "IDArticulo, Fecha")
                        Dim DtVinoMatAct As DataTable = ClsVinoMat.Filter(New FilterItem("IDVinoMaterial", FilterOperator.Equal, DrVinoMat("IDVinoMaterial")))
                        If Not DtVinoMatAct Is Nothing AndAlso DtVinoMatAct.Rows.Count > 0 Then
                            Select Case DrVinoMat("CriterioValoracion")
                                Case enumtaValoracion.taPrecioEstandar
                                    DtVinoMatAct.Rows(0)("Precio") = Nz(DrVinoMat("PrecioEstandar"), 0)
                                Case enumtaValoracion.taPrecioFIFOFecha
                                    DtVinoMatAct.Rows(0)("Precio") = Nz(DrVinoMat("FifoF"), 0)
                                Case enumtaValoracion.taPrecioFIFOMvto
                                    DtVinoMatAct.Rows(0)("Precio") = Nz(DrVinoMat("FifoFD"), 0)
                                Case enumtaValoracion.taPrecioMedio
                                    DtVinoMatAct.Rows(0)("Precio") = Nz(DrVinoMat("PrecioMedio"), 0)
                                Case enumtaValoracion.taPrecioUltCompra
                                    DtVinoMatAct.Rows(0)("Precio") = Nz(DrVinoMat("PrecioUltimaCompra"), 0)
                            End Select
                            BusinessHelper.UpdateTable(DtVinoMatAct)
                        End If
                    Next
                End If
        End Select
    End Sub

#End Region

End Class

<Serializable()> _
Public Class ArticuloVino
    Public Articulo As DataTable
    Public ArticuloDesglose As DataTable
End Class