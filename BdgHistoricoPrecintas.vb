Public Class BdgHistoricoPrecintas

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgHistoricoPrecintas"

#End Region

#Region "Tasks públicas"
    <Serializable()> _
    Public Class stEstadoPrecintas
        Public Año As String
        Public Periodo As String

        Public Sub New(ByVal año As String, ByVal periodo As String, Optional ByVal sobreescribir As Boolean = True)
            Me.Año = año
            Me.Periodo = periodo
        End Sub
    End Class

    <Task()> _
    Public Shared Function GuardarEstadoPrecintas(ByVal data As stEstadoPrecintas, ByVal services As ServiceProvider) As Boolean
        Try
            Dim bsnHP As New BdgHistoricoPrecintas

            'si hay de la fecha de hoy, los borramos
            Dim f As New Filter
            f.Add(New StringFilterItem("Año", data.Año))
            f.Add(New StringFilterItem("Periodo", data.Periodo))
            Dim dttExistentes As DataTable = bsnHP.Filter(f)
            If (Not dttExistentes Is Nothing AndAlso dttExistentes.Rows.Count > 0) Then
                bsnHP.Delete(dttExistentes)
            End If

            Dim dttUpdate As DataTable = bsnHP.AddNew
            Dim StData As New stObtenerRangosPrecinta()
            StData.ConsejoRegulador = True
            Dim dttF As DataTable = ProcessServer.ExecuteTask(Of stObtenerRangosPrecinta, DataTable)(AddressOf ObtenerRangosLibresPrecinta, StData, services)
            If Not dttF Is Nothing AndAlso dttF.Rows.Count > 0 Then
                For Each dtrEstado As DataRow In dttF.Rows
                    Dim dtrUpdate As DataRow = dttUpdate.NewRow
                    dtrUpdate(_BHP.IDHistoricoPrecintas) = Guid.NewGuid
                    dtrUpdate(_BHP.IDArticulo) = dtrEstado("IDArticulo")
                    dtrUpdate(_BHP.Lote) = dtrEstado("Lote")
                    dtrUpdate(_BHP.SeriePrecinta) = dtrEstado("SeriePrecinta")
                    dtrUpdate(_BHP.NDesdePrecinta) = dtrEstado("NDesdePrecinta")
                    dtrUpdate(_BHP.NHastaPrecinta) = dtrEstado("NHastaPrecinta")
                    dtrUpdate(_BHP.Cantidad) = dtrEstado("StockFisico")
                    dtrUpdate(_BHP.Año) = data.Año
                    dtrUpdate(_BHP.Periodo) = data.Periodo
                    dtrUpdate(_BHP.Fecha) = Today
                    dttUpdate.Rows.Add(dtrUpdate)
                Next
                bsnHP.Update(dttUpdate)
            End If
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Shared Function ExistenPrecintasAñoPeriodo(ByVal data As stEstadoPrecintas, ByVal services As ServiceProvider) As Boolean

        Dim f As New Filter
        f.Add(New StringFilterItem("Año", data.Año))
        f.Add(New StringFilterItem("Periodo", data.Periodo))
        Dim bsnHP As New BdgHistoricoPrecintas
        Dim dttP As DataTable = bsnHP.Filter(f)
        Return (Not dttP Is Nothing AndAlso dttP.Rows.Count > 0)
    End Function
    <Serializable()> _
    Public Class stObtenerRangosPrecinta
        Public Articulo As String
        Public Almacen As String
        Public NSeriePrecinta As String
        Public Tipo As String
        Public Familia As String
        Public Lote As String
        Public Ubicacion As String
        Public MinStock As String
        Public MaxStock As String
        Public EsSeguimiento As Boolean

        Public MinPrecinta As String
        Public MaxPrecinta As String

        Public Bloqueado As enumBoolean

        Public ConsejoRegulador As Boolean?
        Public Sub New()

        End Sub
        Public Sub New(ByVal articulo As String, ByVal almacen As String)
            Me.Articulo = articulo
            Me.Almacen = almacen
        End Sub
        Public Sub New(ByVal articulo As String, ByVal almacen As String, ByVal nserieprecinta As String, ByVal tipo As String, ByVal familia As String, ByVal lote As String, ByVal ubicacion As String, _
                       ByVal minstock As String, ByVal maxstock As String, ByVal bloqueado As enumBoolean)
            Me.Articulo = articulo
            Me.Almacen = almacen
            Me.NSeriePrecinta = nserieprecinta
            Me.Tipo = tipo
            Me.Familia = familia
            Me.Lote = lote
            Me.Ubicacion = ubicacion
            If (IsNumeric(minstock)) Then Me.MinStock = minstock
            If (IsNumeric(maxstock)) Then Me.MaxStock = maxstock
            Me.MaxStock = maxstock
            Me.Bloqueado = bloqueado
        End Sub

        Public Sub New(ByVal articulo As String, ByVal almacen As String, ByVal nserieprecinta As String, ByVal tipo As String, ByVal familia As String, ByVal lote As String, ByVal ubicacion As String, _
                      ByVal minstock As String, ByVal maxstock As String, ByVal bloqueado As enumBoolean, ByVal minprecinta As String, ByVal maxprecinta As String)
            Me.Articulo = articulo
            Me.Almacen = almacen
            Me.NSeriePrecinta = nserieprecinta
            Me.Tipo = tipo
            Me.Familia = familia
            Me.Lote = lote
            Me.Ubicacion = ubicacion
            If (IsNumeric(minstock)) Then Me.MinStock = minstock
            If (IsNumeric(maxstock)) Then Me.MaxStock = maxstock
            Me.MaxStock = maxstock
            Me.Bloqueado = bloqueado
            If (IsNumeric(minprecinta)) Then Me.MinPrecinta = minprecinta
            If (IsNumeric(maxprecinta)) Then Me.Maxprecinta = maxprecinta

        End Sub

    End Class

    <Task()> _
    Public Shared Function ObtenerRangosLibresPrecinta(ByVal data As stObtenerRangosPrecinta, ByVal services As ServiceProvider) As DataTable
        Dim de As New DataEngine
        '1. Obtenemos las precintas existentes del artículo
        Dim f As New Filter
        If ((Length(data.Articulo)) > 0) Then f.Add(New StringFilterItem("IDArticulo", data.Articulo))
        If ((Length(data.Almacen)) > 0) Then f.Add(New StringFilterItem("IDAlmacen", data.Almacen))
        If ((Length(data.Lote)) > 0) Then f.Add(New StringFilterItem("Lote", data.Lote))
        If ((Length(data.Ubicacion)) > 0) Then f.Add(New StringFilterItem("Ubicacion", data.Ubicacion))
        If ((Length(data.NSeriePrecinta)) > 0) Then f.Add(New StringFilterItem("SeriePrecinta", data.NSeriePrecinta))
        If ((Length(data.Tipo)) > 0) Then f.Add(New StringFilterItem("IDTipo", data.Tipo))
        If ((Length(data.Familia)) > 0) Then f.Add(New StringFilterItem("IDFamilia", data.Familia))

        If ((Length(data.MinPrecinta)) > 0) Then f.Add(New StringFilterItem("NDesdePrecintaSerie", data.MinPrecinta))
        If ((Length(data.MaxPrecinta)) > 0) Then f.Add(New StringFilterItem("NHastaPrecintaSerie", data.Familia))
        If Length(data.ConsejoRegulador) > 0 Then f.Add("ConsejoRegulador", FilterOperator.Equal, data.ConsejoRegulador)
        f.Add(New NumberFilterItem("StockFisico", FilterOperator.GreaterThan, 0))
        Dim dttPrecintasExistentes As DataTable = de.Filter("vBdgCIPrecintas", f)

        Dim dttRangosOcupados As DataTable = ProcessServer.ExecuteTask(Of stObtenerRangosPrecinta, DataTable)(AddressOf ObtenerRangosOcupadosPrecinta, data, services)

        'ahora tenemos lo no disponible en dttRangosOcupados. se trata de generar los disponibles a partir de ellos
        Dim dttResult As DataTable = dttPrecintasExistentes.Clone()
        dttPrecintasExistentes.DefaultView.Sort = "IDArticulo, IDAlmacen, Lote, SeriePrecinta"
        For Each dtrPrecinta As DataRow In dttPrecintasExistentes.Rows
            'se buscan los 'ocupados' para este registro
            f.Clear()
            f.Add(New StringFilterItem("IDArticulo", dtrPrecinta("IDArticulo")))
            f.Add(New StringFilterItem("IDAlmacen", dtrPrecinta("IDAlmacen")))
            f.Add(New StringFilterItem("Lote", dtrPrecinta("Lote")))
            f.Add(New StringFilterItem("Ubicacion", dtrPrecinta("Ubicacion")))
            f.Add(New StringFilterItem("SeriePrecinta", dtrPrecinta("SeriePrecinta")))
            Dim dtrOcupados() As DataRow = dttRangosOcupados.Select(f.Compose(New AdoFilterComposer), "NDesdePrecinta")
            If (dtrOcupados.Length = 0) Then
                'si no hay registros, se mete tal cual
                dttResult = BdgHistoricoPrecintas.AñadirLinea(dttResult, dtrPrecinta, dtrPrecinta("NDesdePrecinta"), dtrPrecinta("NHastaPrecinta"))
            Else
                'si los hay, se mete un registro del desde libre hasta el primer ocupado
                dttResult = BdgHistoricoPrecintas.AñadirLinea(dttResult, dtrPrecinta, dtrPrecinta("NDesdePrecinta"), dtrOcupados(0)("NDesdePrecinta") - 1)
                'se meten los registros 'entre registros ocupados' (hasta + 1 (n) - desde - 1 (n+1))
                If dtrOcupados.Length > 1 Then
                    For i As Integer = 0 To dtrOcupados.Length - 2
                        If (dtrOcupados(i + 1)("NDesdePrecinta") - 1) >= (dtrOcupados(i)("NHastaPrecinta") + 1) Then
                            dttResult = BdgHistoricoPrecintas.AñadirLinea(dttResult, dtrPrecinta, dtrOcupados(i)("NHastaPrecinta") + 1, dtrOcupados(i + 1)("NDesdePrecinta") - 1)
                        End If
                    Next
                End If

                'se mete el registro del último hasta + 1 hasta el hasta libre
                dttResult = BdgHistoricoPrecintas.AñadirLinea(dttResult, dtrPrecinta, dtrOcupados(dtrOcupados.Length - 1)("NHastaPrecinta") + 1, dtrPrecinta("NHastaPrecinta"))
            End If

        Next
        'Return dttResult

        If (Length(data.MinStock) = 0 AndAlso Length(data.MaxStock) = 0) Then Return dttResult

        'restringir si hay máximos y mínimos: agrupar por strGroupKey y obtener sólo aquellos registros que cumplan las condiciones
        Dim fGroup As New Filter(FilterUnionOperator.Or)
        If Length(data.MinStock) > 0 Then fGroup.Add("StockFisico", FilterOperator.GreaterThanOrEqual, data.MinStock)
        If Length(data.MaxStock) > 0 Then fGroup.Add("StockFisico", FilterOperator.LessThanOrEqual, data.MaxStock)
        If Not fGroup.Count > 0 Then Return dttResult

        Dim dttAux As DataTable = dttResult.Clone
        dttAux.DefaultView.RowFilter = fGroup.Compose(New AdoFilterComposer)
        For Each dtr As DataRow In dttResult.Select(fGroup.Compose(New AdoFilterComposer))
            dttAux.Rows.Add(dtr.ItemArray)
        Next
        Return dttAux
    End Function
    Public Shared Function AñadirLinea(ByVal dttResult As DataTable, ByVal dtrPrecinta As DataRow, ByVal intDesde As Integer, ByVal intHasta As Integer, Optional ByVal controlarValorCero As Boolean = True) As DataTable
        Dim dblStockFisico As Double = intHasta - intDesde + 1
        If (Not controlarValorCero OrElse dblStockFisico <> 0) Then
            Return AñadirLinea(dttResult, dtrPrecinta, intDesde, intHasta, dblStockFisico)
        End If
        Return dttResult
    End Function
    Public Shared Function AñadirLinea(ByVal dttResult As DataTable, ByVal dtrPrecinta As DataRow, ByVal intDesde As Integer, ByVal intHasta As Integer, ByVal dblStockFisico As Double, Optional ByVal controlarValorCero As Boolean = True) As DataTable
        If Not controlarValorCero OrElse dblStockFisico <> 0 Then
            Dim dtrNewRangoLibre As DataRow = dttResult.NewRow
            dtrNewRangoLibre("IDArticulo") = dtrPrecinta("IDArticulo")
            dtrNewRangoLibre("DescArticulo") = dtrPrecinta("DescArticulo")
            dtrNewRangoLibre("IDTipo") = dtrPrecinta("IDTipo")
            dtrNewRangoLibre("DescTipo") = dtrPrecinta("DescTipo")
            dtrNewRangoLibre("IDFamilia") = dtrPrecinta("IDFamilia")
            dtrNewRangoLibre("DescFamilia") = dtrPrecinta("DescFamilia")
            dtrNewRangoLibre("IDSubfamilia") = dtrPrecinta("IDSubfamilia")
            dtrNewRangoLibre("DescSubfamilia") = dtrPrecinta("DescSubfamilia")
            dtrNewRangoLibre("IDAlmacen") = dtrPrecinta("IDAlmacen")
            dtrNewRangoLibre("Lote") = dtrPrecinta("Lote")
            dtrNewRangoLibre("Ubicacion") = dtrPrecinta("Ubicacion")
            dtrNewRangoLibre("IDCaracteristicaArticulo1") = dtrPrecinta("IDCaracteristicaArticulo1")
            dtrNewRangoLibre("IDCaracteristicaArticulo2") = dtrPrecinta("IDCaracteristicaArticulo2")
            dtrNewRangoLibre("IDCaracteristicaArticulo3") = dtrPrecinta("IDCaracteristicaArticulo3")
            dtrNewRangoLibre("IDCaracteristicaArticulo4") = dtrPrecinta("IDCaracteristicaArticulo4")
            dtrNewRangoLibre("IDCaracteristicaArticulo5") = dtrPrecinta("IDCaracteristicaArticulo5")
            dtrNewRangoLibre("NDesdePrecinta") = intDesde
            dtrNewRangoLibre("NHastaPrecinta") = intHasta
            dtrNewRangoLibre("NDesdePrecintaSerie") = dtrPrecinta("NDesdePrecintaSerie")
            dtrNewRangoLibre("NHastaPrecintaSerie") = dtrPrecinta("NHastaPrecintaSerie")
            dtrNewRangoLibre("StockFisico") = dblStockFisico
            dtrNewRangoLibre("SeriePrecinta") = dtrPrecinta("SeriePrecinta")
            dtrNewRangoLibre("ID") = String.Format("{0}~{1}", dtrPrecinta("ID"), intDesde)

            dttResult.Rows.Add(dtrNewRangoLibre)
            Return dttResult
        End If
        Return dttResult
    End Function

    <Task()> Public Shared Function ObtenerRangosOcupadosPrecinta(ByVal data As stObtenerRangosPrecinta, ByVal services As ServiceProvider) As DataTable

        Dim de As New DataEngine
        '1. Obtenemos las precintas existentes del artículo
        Dim f As New Filter
        If ((Length(data.Articulo)) > 0) Then f.Add(New StringFilterItem("IDArticulo", data.Articulo))
        If ((Length(data.Almacen)) > 0) Then f.Add(New StringFilterItem("IDAlmacen", data.Almacen))
        If ((Length(data.Lote)) > 0) Then f.Add(New StringFilterItem("Lote", data.Lote))
        If ((Length(data.Ubicacion)) > 0) Then f.Add(New StringFilterItem("Ubicacion", data.Ubicacion))
        If ((Length(data.NSeriePrecinta)) > 0) Then f.Add(New StringFilterItem("SeriePrecinta", data.NSeriePrecinta))
        If ((Length(data.Tipo)) > 0) Then f.Add(New StringFilterItem("IDTipo", data.Tipo))
        If ((Length(data.Familia)) > 0) Then f.Add(New StringFilterItem("IDFamilia", data.Familia))

        If ((Length(data.MinPrecinta)) > 0) Then f.Add(New StringFilterItem("NDesdePrecintaSerie", data.MinPrecinta))
        If ((Length(data.MaxPrecinta)) > 0) Then f.Add(New StringFilterItem("NHastaPrecintaSerie", data.MaxPrecinta))

        If Not data.EsSeguimiento Then f.Add(New NumberFilterItem("StockFisico", FilterOperator.GreaterThan, 0))
        Dim dttPrecintasExistentes As DataTable = de.Filter("vBdgCIPrecintas", f)
        Dim dttRangosOcupados As DataTable = dttPrecintasExistentes.Clone

        '2. Para cada uno, miramos sus entradas y salidas en tbHistoricoMovimiento para ver los rangos ocupados (puede que haya AJUSTES, habría que mirar en orden)
        For Each dtrPrecinta As DataRow In dttPrecintasExistentes.Rows
            'TODO - 
            'HAY QUE INTRODUCIR LOS AJUSTES TAMBIÉN AQUÍ. SUPONGAMOS PRECINTA
            '=> PRE    1   100.

            'MOVIMIENTOS, ORDENADOS DESC:

            '=> PRE    91  95 SALIDA   
            '=> PRE    91  100 AJUSTE POSITIVO
            '=> PRE    51  100 SALIDA
            '=> PRE    1   50  SALIDA  
            '=> PRE    1   100 ENTRADA

            'EN CUYO CASO, DEBERÍA MOSTRARNOS COMO RESULTADO EL RANGO
            'SI EL TIPO ES SALIDA, SE AÑADE COMO RANGO NO DISPONIBLE
            'SI EL TIPO ES AJUSTE POSITIVO, SE ELIMINA LA SALIDA A LA QUE AFECTE (SI ES QUE AFECTA) Y SE GENERAN LOS RANGOS PERTINENTES
            'SI EL TIPO ES AJUSTE NEGATIVO, SE AÑADE COMO RANGO NO DISPONIBLE

            'FINALMENTE, CON LOS RANGOS OCUPADOS Y LOS TOTALES, OBTENEMOS LOS LIBRES

            '3. buscamos en la tabla de movimientos los registros relacionados con esta precinta
            f.Clear()
            f.Add(New StringFilterItem("IDArticulo", dtrPrecinta("IDArticulo")))
            f.Add(New StringFilterItem("IDAlmacen", dtrPrecinta("IDAlmacen")))
            f.Add(New StringFilterItem("Lote", dtrPrecinta("Lote")))
            f.Add(New StringFilterItem("Ubicacion", dtrPrecinta("Ubicacion")))
            f.Add(New StringFilterItem("SeriePrecinta", dtrPrecinta("SeriePrecinta")))
            f.Add(New StringFilterItem("NDesdePrecinta", dtrPrecinta("NDesdePrecinta")))
            f.Add(New StringFilterItem("NHastaPrecinta", dtrPrecinta("NHastaPrecinta")))
            f.Add(New NumberFilterItem("IDTipoMovimiento", FilterOperator.NotEqual, enumTipoMovimiento.tmCorreccion))
            'f.Add(New NumberFilterItem("ClaseMovimiento", enumtpmTipoMovimiento.tpmOutput))
            Dim dttPrecintasMovs As DataTable = de.Filter("vBdgCIPrecintasOcupacion", f, , "FechaDocumento, IDLineaMovimiento ASC")
            If (Not dttPrecintasMovs Is Nothing AndAlso dttPrecintasMovs.Rows.Count > 0) Then
                '3.2 Si los hay, hay que mirar qué tipo de movimiento es para si añadir o no un registro y sus cantidades
                Dim intRangoInicial As Integer = dtrPrecinta("NDesdePrecinta")

                For Each dtrMovimiento As DataRow In dttPrecintasMovs.Rows
                    Select Case (dtrMovimiento("IDTipoMovimiento"))
                        Case enumTipoMovimiento.tmSalAjuste, enumTipoMovimiento.tmSalAlbaranVenta, enumTipoMovimiento.tmSalFabrica
                            'SALIDA, SE AÑADE
                            dttRangosOcupados = BdgHistoricoPrecintas.AñadirLinea(dttRangosOcupados, dtrPrecinta, dtrMovimiento("NDesdePrecintaUtilizada"), dtrMovimiento("NHastaPrecintaUtilizada"), _
                                                                                   dtrMovimiento("NHastaPrecintaUtilizada") - dtrMovimiento("NDesdePrecintaUtilizada") + 1)

                        Case enumTipoMovimiento.tmEntAjuste
                            'ENTRADA => SE BUSCA SI HAY ALGÚN REGISTRO AFECTADO, SE ELIMINAN, Y SE GENERAN A PARTIR DE ELLOS
                            Dim faux As New Filter
                            faux.Add(New StringFilterItem("IDArticulo", dtrPrecinta("IDArticulo")))
                            faux.Add(New StringFilterItem("IDAlmacen", dtrPrecinta("IDAlmacen")))
                            faux.Add(New StringFilterItem("Lote", dtrPrecinta("Lote")))
                            faux.Add(New StringFilterItem("Ubicacion", dtrPrecinta("Ubicacion")))
                            faux.Add(New StringFilterItem("SeriePrecinta", dtrPrecinta("SeriePrecinta")))
                            faux.Add("NHastaPrecinta", FilterOperator.GreaterThanOrEqual, dtrMovimiento("NDesdePrecintaUtilizada"))
                            faux.Add("NDesdePrecinta", FilterOperator.LessThanOrEqual, dtrMovimiento("NHastaPrecintaUtilizada"))
                            Dim dvRangosTemp As DataView = New DataView(dttRangosOcupados)
                            'dvRangosTemp.Sort = ""
                            dvRangosTemp.RowFilter = faux.Compose(New AdoFilterComposer)
                            If (dvRangosTemp.Count > 0) Then

                                'metemos las dos líneas nuevas
                                'del desde de la borrada hasta el desde de la nueva
                                dttRangosOcupados = BdgHistoricoPrecintas.AñadirLinea(dttRangosOcupados, dtrPrecinta, dvRangosTemp(0)("NDesdePrecinta"), dtrMovimiento("NDesdePrecintaUtilizada") - 1)

                                'del hasta de la nueva hasta el hasta de la última
                                If dtrMovimiento("NHastaPrecintaUtilizada") < dvRangosTemp(dvRangosTemp.Count - 1)("NHastaPrecinta") Then
                                    dttRangosOcupados = BdgHistoricoPrecintas.AñadirLinea(dttRangosOcupados, dtrPrecinta, dtrMovimiento("NHastaPrecintaUtilizada") + 1, dvRangosTemp(dvRangosTemp.Count - 1)("NHastaPrecinta"))
                                End If
                                'eliminamos las filas y generamos las suyas
                                For Each dtv As DataRowView In dvRangosTemp
                                    'quitamos la vieja
                                    dttRangosOcupados.Rows.Remove(dtv.Row)
                                Next
                            End If

                    End Select
                Next
            End If
        Next
        Return dttRangosOcupados
    End Function

    <Serializable()> _
    Public Class DataProcSegPrecRet
        Public IDEtiqueta As String
        Public Factor As Double
        Public DtLotes As DataTable

        Public Sub New()
        End Sub
        Public Sub New(ByVal IDEtiqueta As String, ByVal Factor As Double)
            Me.IDEtiqueta = IDEtiqueta
            Me.Factor = Factor
        End Sub
    End Class

    <Task()> Public Shared Function ProcesoSegPrecinta(ByVal IDArticulo As String, ByVal services As ServiceProvider) As DataProcSegPrecRet
        Dim ClsArt As New Articulo
        Dim DtArt As DataTable = ClsArt.SelOnPrimaryKey(IDArticulo)
        If Length(DtArt.Rows(0)("IDArticuloPadre")) > 0 Then
            IDArticulo = DtArt.Rows(0)("IDArticuloPadre")
        End If
        Dim StData As New DataProcSegPrecRet
        StData.IDEtiqueta = ProcessServer.ExecuteTask(Of String, String)(AddressOf ObtenerArticuloSegPrecinta, IDArticulo, services)
        If Length(StData.IDEtiqueta) > 0 Then
            'Controlar si Los IDArticulo son iguales
            If DtArt.Rows(0)("IDArticulo") <> IDArticulo Then
                Dim StDataFactor As New UnidadAB.UnidadMedidaInfo
                StDataFactor.IDUdMedidaA = DtArt.Rows(0)("IDUDInterna")
                StDataFactor.IDUdMedidaB = ClsArt.SelOnPrimaryKey(IDArticulo).Rows(0)("IDUDInterna")
                StData.Factor = ProcessServer.ExecuteTask(Of UnidadAB.UnidadMedidaInfo, Double)(AddressOf UnidadAB.FactorDeConversion, StDataFactor, services)
            Else : StData.Factor = 1
            End If
            'Cargar todos los datos de Articulo - Almacen - Lote de la IDEtiqueta y devolver a pantalla
            StData.DtLotes = New ArticuloAlmacenLote().Filter(New FilterItem("IDArticulo", FilterOperator.Equal, StData.IDEtiqueta), "IDArticulo,IDAlmacen,Lote", "IDArticulo,IDAlmacen,Lote,Ubicacion,NDesdePrecinta,NHastaPrecinta")
            StData.DtLotes.Columns.Add("Codigo", GetType(String))
            For Each DrLote As DataRow In StData.DtLotes.Select
                DrLote("Codigo") = DrLote("IDArticulo") & "-" & DrLote("IDAlmacen") & "-" & DrLote("Lote") & "-" & DrLote("Ubicacion")
            Next
        End If
        Return StData
    End Function

    <Task()> Public Shared Function ObtenerArticuloSegPrecinta(ByVal IDArticulo As String, ByVal services As ServiceProvider) As String
        Dim BlnFindSeg As Boolean = False
        Dim IDArtFindSeg As String = String.Empty
        Dim ClsEst As New Estructura
        Dim DtEst As DataTable = ClsEst.Filter(New FilterItem("IDArticulo", FilterOperator.Equal, IDArticulo))
        If Not DtEst Is Nothing AndAlso DtEst.Rows.Count > 0 Then
            For Each DrEst As DataRow In DtEst.Select
                BlnFindSeg = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ObtenerSeguimientoPrecinta, DrEst("IDComponente"), services)
                If BlnFindSeg Then
                    IDArtFindSeg = DrEst("IDComponente")
                    Exit For
                End If
            Next
        End If
        Return IDArtFindSeg
    End Function

    <Task()> Public Shared Function ObtenerSeguimientoPrecinta(ByVal IDArticulo As String, ByVal services As ServiceProvider) As Boolean
        Dim BlnFindSeguimiento As Boolean = False
        Dim DtArt As DataTable = New Articulo().SelOnPrimaryKey(IDArticulo)
        If Length(DtArt.Rows(0)("IDSubFamilia")) > 0 Then
            Dim FilSubFam As New Filter
            FilSubFam.Add("IDTipo", FilterOperator.Equal, DtArt.Rows(0)("IDTipo"))
            FilSubFam.Add("IDFamilia", FilterOperator.Equal, DtArt.Rows(0)("IDFamilia"))
            FilSubFam.Add("IDSubFamilia", FilterOperator.Equal, DtArt.Rows(0)("IDSubFamilia"))
            Dim DtSubFam As DataTable = New Subfamilia().Filter(FilSubFam)
            If Not DtSubFam Is Nothing AndAlso DtSubFam.Rows.Count > 0 Then
                If DtSubFam.Rows(0)("SeguimientoPrecinta") Then
                    BlnFindSeguimiento = True
                End If
            End If
        End If
        If Not BlnFindSeguimiento Then
            Dim FilFam As New Filter
            FilFam.Add("IDTipo", FilterOperator.Equal, DtArt.Rows(0)("IDTipo"))
            FilFam.Add("IDFamilia", FilterOperator.Equal, DtArt.Rows(0)("IDFamilia"))
            Dim DtFam As DataTable = New Familia().Filter(FilFam)
            If Not DtFam Is Nothing AndAlso DtFam.Rows.Count > 0 Then
                If DtFam.Rows(0)("SeguimientoPrecinta") Then
                    BlnFindSeguimiento = True
                End If
            End If
        End If
        Return BlnFindSeguimiento
    End Function

    <Serializable()> _
    Public Class DataExistsEtiSeg
        Public IDArticulo As String
        Public IDAlmacen As String
        Public Lote As String
        Public Ubicacion As String
        Public NDesde As Integer
        Public NHasta As Integer
        Public EsDevolucion As Boolean

        Public Sub New()
        End Sub
        Public Sub New(ByVal IDArticulo As String, ByVal IDAlmacen As String, ByVal Lote As String, ByVal Ubicacion As String, ByVal NDesde As Integer, ByVal NHasta As Integer, ByVal EsDevolucion As Boolean)
            Me.IDArticulo = IDArticulo
            Me.IDAlmacen = IDAlmacen
            Me.Lote = Lote
            Me.Ubicacion = Ubicacion
            Me.NDesde = NDesde
            Me.NHasta = NHasta
            Me.EsDevolucion = EsDevolucion
        End Sub
    End Class

    <Task()> Public Shared Function EtiquetaSeguimientoDisponible(ByVal data As DataExistsEtiSeg, ByVal services As ServiceProvider) As Boolean
        Dim BlnDisp As Boolean
        'Comprobar que existan en Rango de Precintas de su Histórico
        Dim StDataOcup As New stObtenerRangosPrecinta(data.IDArticulo, data.IDAlmacen)
        StDataOcup.Lote = data.Lote
        StDataOcup.Ubicacion = data.Ubicacion
        StDataOcup.EsSeguimiento = True
        Dim dtEtiquetasPegadas As DataTable = ProcessServer.ExecuteTask(Of stObtenerRangosPrecinta, DataTable)(AddressOf ObtenerRangosOcupadosPrecinta, StDataOcup, services)
        If Not dtEtiquetasPegadas Is Nothing AndAlso dtEtiquetasPegadas.Rows.Count > 0 Then
            Dim RangosOcupados As List(Of DataRow) = (From c In dtEtiquetasPegadas _
                                                      Where Not c.IsNull("NDesdePrecinta") AndAlso Not c.IsNull("NHastaPrecinta") AndAlso _
                                                               ((c("NDesdePrecinta") <= data.NDesde AndAlso c("NHastaPrecinta") >= data.NHasta)) _
                                                      Select c).ToList
            If Not RangosOcupados Is Nothing AndAlso RangosOcupados.Count > 0 Then
                '//Comprobar que no existan en seguimiento de albaranes, es decir, comprobar si se han vendido.
                Dim FilExistSeg As New Filter
                FilExistSeg.Add("IDArticulo", FilterOperator.Equal, data.IDArticulo)
                FilExistSeg.Add("IDAlmacen", FilterOperator.Equal, data.IDAlmacen)
                FilExistSeg.Add("Lote", FilterOperator.Equal, data.Lote)
                FilExistSeg.Add("Ubicacion", FilterOperator.Equal, data.Ubicacion)

                Dim FilExistBet As New Filter(FilterUnionOperator.Or)
                FilExistBet.Add(New BetweenFilterItem("NDesde", data.NDesde, data.NHasta, FilterType.Numeric))
                FilExistBet.Add(New BetweenFilterItem("NHasta", data.NDesde, data.NHasta, FilterType.Numeric))

                Dim f As New Filter
                f.Add(New NumberFilterItem("NDesde", FilterOperator.LessThanOrEqual, data.NHasta))
                f.Add(New NumberFilterItem("NHasta", FilterOperator.GreaterThanOrEqual, data.NDesde))
                FilExistBet.Add(f)
                FilExistSeg.Add(FilExistBet)

                Dim dtEtiquetaVendida As DataTable = New BE.DataEngine().Filter("NegBdgAlbaranVentaSeguimiento", FilExistSeg)
                If dtEtiquetaVendida Is Nothing OrElse dtEtiquetaVendida.Rows.Count = 0 Then
                    If data.EsDevolucion Then
                        BlnDisp = False
                        ApplicationService.GenerateError("No se han han realizado ventas de las etiquetas [{0}-{1}]. No se pueden devolver.", data.NDesde, data.NHasta)
                    Else
                        BlnDisp = True
                    End If
                Else
                    If dtEtiquetaVendida.Rows.Count = 1 Then
                        Dim SignoCantidad As Integer = Math.Sign(dtEtiquetaVendida.Rows(0)("Cantidad"))
                        If SignoCantidad < 0 Then '//Devolución
                            ApplicationService.GenerateError("Revise sus datos. Sólo existe una devolución en sus movimientos.{0}Se encuentra en el rango [{1}-{2}]", vbNewLine, dtEtiquetaVendida.Rows(0)("NDesde"), dtEtiquetaVendida.Rows(0)("NHasta"))
                        Else
                            If data.EsDevolucion Then
                                BlnDisp = True
                            Else
                                BlnDisp = False
                            End If
                        End If
                    Else
                        Dim lstSeguimientoOcupadas As New List(Of DataEtiquetaVendidas)
                        '//Si tenemos varias, es por que ha habido devoluciones y se ha vuelto a vender.
                        Dim SeguimientoAV As List(Of DataRow) = (From c In dtEtiquetaVendida Order By c("FechaAlbaran") Ascending Select c).ToList
                        For Each dr As DataRow In SeguimientoAV
                            Dim SignoCantidad As Integer = Math.Sign(dr("Cantidad"))
                            If SignoCantidad < 0 Then     '//Devolución
                                Dim BusquedaEtiVenta As List(Of DataEtiquetaVendidas) = (From e In lstSeguimientoOcupadas _
                                                                                         Where (e.NDesde <= dr("NDesde") AndAlso e.NHasta >= dr("NHasta")) _
                                                                                         Order By e.FechaAlbaran Descending _
                                                                                         Select e).ToList
                                If Not BusquedaEtiVenta Is Nothing AndAlso BusquedaEtiVenta.Count > 0 Then
                                    '//La última venta la desdoblamos en rangos, eliminando el rango devuelto
                                    Dim EtiVentaEliminar As DataEtiquetaVendidas = BusquedaEtiVenta(0)
                                    Dim NDesde As Integer = EtiVentaEliminar.NDesde
                                    Dim NHasta As Integer = EtiVentaEliminar.NHasta

                                    If EtiVentaEliminar.NDesde < dr("NDesde") - 1 Then
                                        Dim eti1 As New DataEtiquetaVendidas(dr("FechaAlbaran"), dr("IDArticulo"), dr("IDAlmacen"), dr("Lote"), dr("Ubicacion"), EtiVentaEliminar.NDesde, dr("NDesde") - 1)
                                        lstSeguimientoOcupadas.Add(eti1)
                                    End If

                                    If dr("NHasta") + 1 < EtiVentaEliminar.NHasta Then
                                        Dim eti2 As New DataEtiquetaVendidas(dr("FechaAlbaran"), dr("IDArticulo"), dr("IDAlmacen"), dr("Lote"), dr("Ubicacion"), dr("NHasta") + 1, EtiVentaEliminar.NHasta)
                                        lstSeguimientoOcupadas.Add(eti2)
                                    End If

                                    lstSeguimientoOcupadas.Remove(EtiVentaEliminar)
                                Else
                                    '//La devolución se encuentra dentro de varios rangos
                                    Dim SalgoRangoPorDesdeyHasta As List(Of DataEtiquetaVendidas) = (From e In lstSeguimientoOcupadas _
                                                                                                     Where (dr("NDesde") <= e.NDesde AndAlso e.NDesde <= dr("NHasta")) AndAlso _
                                                                                                           (dr("NDesde") <= e.NHasta AndAlso e.NHasta <= dr("NHasta")) _
                                                                                                     Order By e.FechaAlbaran Descending _
                                                                                                     Select e).ToList
                                    If Not SalgoRangoPorDesdeyHasta Is Nothing AndAlso SalgoRangoPorDesdeyHasta.Count > 0 Then
                                        Dim et As DataEtiquetaVendidas = SalgoRangoPorDesdeyHasta(0)
                                        lstSeguimientoOcupadas.Remove(et)
                                    End If

                                    Dim SalgoRangoPorDesde As List(Of DataEtiquetaVendidas) = (From e In lstSeguimientoOcupadas _
                                                                                               Where (dr("NDesde") <= e.NDesde AndAlso e.NDesde <= dr("NHasta")) _
                                                                                               Order By e.FechaAlbaran Descending _
                                                                                               Select e).ToList
                                    If Not SalgoRangoPorDesde Is Nothing AndAlso SalgoRangoPorDesde.Count > 0 Then
                                        Dim et As DataEtiquetaVendidas = SalgoRangoPorDesde(0)
                                        If dr("NHasta") + 1 < et.NHasta Then
                                            Dim eti2 As New DataEtiquetaVendidas(dr("FechaAlbaran"), dr("IDArticulo"), dr("IDAlmacen"), dr("Lote"), dr("Ubicacion"), dr("NHasta") + 1, et.NHasta)
                                            lstSeguimientoOcupadas.Add(eti2)
                                        End If

                                        lstSeguimientoOcupadas.Remove(et)
                                    End If

                                    Dim SalgoRangoPorHasta As List(Of DataEtiquetaVendidas) = (From e In lstSeguimientoOcupadas _
                                                                                                 Where (dr("NDesde") <= e.NHasta AndAlso e.NHasta <= dr("NHasta")) _
                                                                                                 Order By e.FechaAlbaran Descending _
                                                                                                 Select e).ToList
                                    If Not SalgoRangoPorHasta Is Nothing AndAlso SalgoRangoPorHasta.Count > 0 Then
                                        Dim et As DataEtiquetaVendidas = SalgoRangoPorHasta(0)

                                        If et.NDesde < dr("NDesde") - 1 Then
                                            Dim eti1 As New DataEtiquetaVendidas(dr("FechaAlbaran"), dr("IDArticulo"), dr("IDAlmacen"), dr("Lote"), dr("Ubicacion"), et.NDesde, dr("NDesde") - 1)
                                            lstSeguimientoOcupadas.Add(eti1)
                                        End If

                                        lstSeguimientoOcupadas.Remove(et)
                                    End If

                                End If
                            Else      '//Venta
                                Dim eti As New DataEtiquetaVendidas(dr("FechaAlbaran"), dr("IDArticulo"), dr("IDAlmacen"), dr("Lote"), dr("Ubicacion"), dr("NDesde"), dr("NHasta"))
                                lstSeguimientoOcupadas.Add(eti)
                            End If
                        Next

                        '//Buscamos si tras varios movimientos seguimos teniendo la etiqueta diponible
                        Dim BusquedaFinal As List(Of DataEtiquetaVendidas) = (From e In lstSeguimientoOcupadas _
                                                                              Where (e.NDesde <= data.NHasta AndAlso e.NHasta >= data.NDesde) _
                                                                              Order By e.FechaAlbaran Descending _
                                                                              Select e).ToList
                        If Not BusquedaFinal Is Nothing AndAlso BusquedaFinal.Count > 0 Then
                            If data.EsDevolucion Then
                                BlnDisp = True
                            Else
                                BlnDisp = False
                            End If
                        Else
                            BlnDisp = True
                        End If

                    End If
                End If

            Else
                ApplicationService.GenerateError("El rango indicado no está preparado para expedir. Alguna etiqueta en el rango [{0} - {1}] no está disponible.", data.NDesde, data.NHasta)
            End If
        End If
        Return BlnDisp
    End Function

  


    <Serializable()> _
    Public Class DataEtiquetaVendidas
        Public FechaAlbaran As Date
        Public IDArticulo As String
        Public IDAlmacen As String
        Public Lote As String
        Public Ubicacion As String
        Public NDesde As Integer
        Public NHasta As Integer

        Public Sub New(ByVal FechaAlbaran As Date, ByVal IDArticulo As String, ByVal IDAlmacen As String, ByVal Lote As String, ByVal Ubicacion As String, ByVal NDesde As Integer, ByVal NHasta As Integer)
            Me.FechaAlbaran = FechaAlbaran
            Me.IDArticulo = IDArticulo
            Me.IDAlmacen = IDAlmacen
            Me.Lote = Lote
            Me.Ubicacion = Ubicacion
            Me.NDesde = NDesde
            Me.NHasta = NHasta
        End Sub

    End Class

#End Region

End Class

Public Class _BHP
    Public Const IDHistoricoPrecintas As String = "IDHistoricoPrecintas"
    Public Const IDArticulo As String = "IDArticulo"
    Public Const Lote As String = "Lote"
    Public Const SeriePrecinta As String = "SeriePrecinta"
    Public Const NDesdePrecinta As String = "NDesdePrecinta"
    Public Const NHastaPrecinta As String = "NHastaPrecinta"
    Public Const Cantidad As String = "Cantidad"
    Public Const Fecha As String = "Fecha"
    Public Const Año As String = "Año"
    Public Const Periodo As String = "Periodo"
End Class
