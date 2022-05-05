Public Class BdgVino

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgVino"

#End Region

#Region "Eventos Entidad"

    Public Overloads Function GetItemRow(ByVal IDVino As Guid) As DataRow
        Dim dt As DataTable = New BdgVino().SelOnPrimaryKey(IDVino)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then
            ApplicationService.GenerateError("No existe el vino")
        Else : Return dt.Rows(0)
        End If
    End Function

#End Region

#Region " RegisterValidateTasks "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarCambioFechaVino)
    End Sub

    <Task()> Public Shared Sub ValidarCambioFechaVino(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Modified AndAlso Nz(data("Fecha"), cnMinDate) <> Nz(data("Fecha", DataRowVersion.Original), cnMinDate) Then
            '//Comprobar que el vino no tenga operaciones posteriores.
            Dim f As New Filter
            f.Add(New GuidFilterItem(_VE.IDVinoComponente, data("IDVino")))
            Dim dtVE As DataTable = New BdgVinoEstructura().Filter(f)
            If dtVE.Rows.Count > 0 Then
                Dim strError As String = String.Format("No se puede modificar la fecha. El vino siguiente tiene operaciones:{0} Depósito: {1} {2} Artículo: {3} {4} Lote: {5} {6} Almacén: {7} {8} Fecha Llenado: {9}", vbNewLine, data(_V.IDDeposito), vbNewLine, data(_V.IDArticulo), vbNewLine, data(_V.Lote), vbNewLine, data(_V.IDAlmacen), vbNewLine, data(_V.Fecha))
                ApplicationService.GenerateError(strError)
            End If
        End If
    End Sub

#End Region

#Region " RegisterUpdateTasks "

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf PrepararCambioFechaVino)
        updateProcess.AddTask(Of DataRow)(AddressOf Comunes.UpdateEntityRow)
        updateProcess.AddTask(Of DataRow)(AddressOf Comunes.MarcarComoActualizado)
        updateProcess.AddTask(Of DataRow)(AddressOf ActualizarCambioFechaVino)
    End Sub

    <Task()> Public Shared Sub PrepararCambioFechaVino(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Modified AndAlso Nz(data("Fecha"), cnMinDate) <> Nz(data("Fecha", DataRowVersion.Original), cnMinDate) Then
            '//Actualizar los días de depósito
            data(_V.DiasDeposito) = 0
            data(_V.DiasBarrica) = 0
            data(_V.DiasBotellero) = 0
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarCambioFechaVino(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Modified AndAlso Nz(data("Fecha"), cnMinDate) <> Nz(data("Fecha", DataRowVersion.Original), cnMinDate) Then
            ProcessServer.ExecuteTask(Of Guid)(AddressOf BdgWorkClass.ActualizarDiasVino, data("IDVino"), services)
        End If
    End Sub

#End Region

#Region " RegisterDeleteTasks "

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf ValidarAnalisisVino)
    End Sub

    <Task()> Public Shared Sub ValidarAnalisisVino(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim dtVinoAnalisis As DataTable = New BdgVinoAnalisis().Filter(New GuidFilterItem("IDVino", data("IDVino")))
        If Not dtVinoAnalisis Is Nothing AndAlso dtVinoAnalisis.Rows.Count > 0 Then
            Dim ListaAnalisis As String = Strings.Join((From c In dtVinoAnalisis Select c("NVinoAnalisis") Distinct).ToArray, ",")
            ApplicationService.GenerateError("No se puede eliminar el Vino, ya que tiene análisis asociados.{0}Análisis: {1}", vbNewLine, ListaAnalisis)
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Sub DeleteVino(ByVal IDVino As Guid, ByVal services As ServiceProvider)
        Dim ClsVino As New BdgVino
        Dim dr As DataRow = ClsVino.GetItemRow(IDVino)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf ValidarAnalisisVino, dr, services)
        ClsVino.Delete(dr)
    End Sub

    <Task()> Public Shared Sub DeleteRw(ByVal oRw As DataRow, ByVal services As ServiceProvider)
        Dim ClsVino As New BdgVino
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf ValidarAnalisisVino, oRw, services)
        ClsVino.Delete(oRw)
    End Sub

    <Serializable()> _
    Public Class StExplosion
        Public IDVino As Guid
        Public NivelMaximo As Integer
        Public VerMaterial As Boolean
        Public VerCosteElaboracion As Boolean
        Public VerCosteEstanciaNave As Boolean
        Public VerCosteInicial As Boolean
        Public VerDetalleOperaciones As Boolean
        Public VerDetalleAnalisis As Boolean
        Public VerDetalleEntradasUva As Boolean
        Public VerDetalleEntradasVino As Boolean

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDVino As Guid, _
        Optional ByVal NivelMaximo As Integer = 100, _
                       Optional ByVal VerMaterial As Boolean = False, _
                       Optional ByVal VerCosteElaboracion As Boolean = False, _
                       Optional ByVal VerCosteEstanciaNave As Boolean = False, _
                       Optional ByVal VerCosteInicial As Boolean = False, _
                       Optional ByVal VerDetalleOperaciones As Boolean = True, _
                       Optional ByVal VerDetalleAnalisis As Boolean = True, _
                       Optional ByVal VerDetalleEntradasUva As Boolean = True, _
                       Optional ByVal VerDetalleEntradasVino As Boolean = True)
            Me.IDVino = IDVino
            Me.NivelMaximo = NivelMaximo
            Me.VerMaterial = VerMaterial
            Me.VerCosteElaboracion = VerCosteElaboracion
            Me.VerCosteEstanciaNave = VerCosteEstanciaNave
            Me.VerCosteInicial = VerCosteInicial
            Me.VerDetalleOperaciones = VerDetalleOperaciones
            Me.VerDetalleAnalisis = VerDetalleAnalisis
            Me.VerDetalleEntradasUva = VerDetalleEntradasUva
            Me.VerDetalleEntradasVino = VerDetalleEntradasVino
        End Sub
    End Class

    <Task()> Public Shared Function ExplosionGrafo(ByVal data As StExplosion, ByVal services As ServiceProvider) As DataExplosionVinoResult
        If Not data.NivelMaximo > 0 Then data.NivelMaximo = 100
        Dim Material As Integer
        If CInt(data.VerMaterial) <> 0 Then
            Material = 1
        Else : Material = 0
        End If
        Dim Elaboracion As Integer
        If CInt(data.VerCosteElaboracion) <> 0 Then
            Elaboracion = 1
        Else
            Elaboracion = 0
        End If
        Dim EstanciaNave As Integer
        If CInt(data.VerCosteEstanciaNave) <> 0 Then
            EstanciaNave = 1
        Else
            EstanciaNave = 0
        End If
        Dim CosteInicial As Integer
        If CInt(data.VerCosteInicial) <> 0 Then
            CosteInicial = 1
        Else
            CosteInicial = 0
        End If

        Dim dat As New DataExplosionVino(data.IDVino, data.VerMaterial, data.VerCosteElaboracion, data.VerCosteEstanciaNave, data.VerCosteInicial, _
                                         data.VerDetalleOperaciones, data.VerDetalleAnalisis, data.VerDetalleEntradasUva, data.VerDetalleEntradasVino)
        Dim rstl As DataExplosionVinoResult = ProcessServer.ExecuteTask(Of DataExplosionVino, DataExplosionVinoResult)(AddressOf BdgExplosionVino.Explosion, dat, services)

        Return rstl

    End Function

    <Task()> Public Shared Function Explosion(ByVal data As StExplosion, ByVal services As ServiceProvider) As DataExplosionVinoResult
        'If Not data.NivelMaximo > 0 Then data.NivelMaximo = 100
        Dim Material As Integer
        If CInt(data.VerMaterial) <> 0 Then
            Material = 1
        Else : Material = 0
        End If
        Dim Elaboracion As Integer
        If CInt(data.VerCosteElaboracion) <> 0 Then
            Elaboracion = 1
        Else
            Elaboracion = 0
        End If
        Dim EstanciaNave As Integer
        If CInt(data.VerCosteEstanciaNave) <> 0 Then
            EstanciaNave = 1
        Else
            EstanciaNave = 0
        End If
        Dim CosteInicial As Integer
        If CInt(data.VerCosteInicial) <> 0 Then
            CosteInicial = 1
        Else
            CosteInicial = 0
        End If

        'Return AdminData.Execute("spBdgVinoExplosion", 2, False, data.IDVino, Material, data.NivelMaximo, Elaboracion, EstanciaNave, CosteInicial)
        Dim dat As New DataExplosionVino(data.IDVino, data.VerMaterial, data.VerCosteElaboracion, data.VerCosteEstanciaNave, data.VerCosteInicial, _
                                         data.VerDetalleOperaciones, data.VerDetalleAnalisis, data.VerDetalleEntradasUva, data.VerDetalleEntradasVino)
        Dim rstl As DataExplosionVinoResult = ProcessServer.ExecuteTask(Of DataExplosionVino, DataExplosionVinoResult)(AddressOf BdgExplosionVino.Explosion, dat, services)

        Return rstl

    End Function

    <Serializable()> _
    Public Class StExplosionDetallada
        Public IDVino As Guid
        Public NivelMaximo As Integer
        Public OpcionesTraza As DataOpcionesTraza

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDVino As Guid, Optional ByVal OpcionesTraza As DataOpcionesTraza = Nothing)
            Me.IDVino = IDVino
            Me.NivelMaximo = 100
            Me.OpcionesTraza = OpcionesTraza
            If Not OpcionesTraza Is Nothing Then
                Me.NivelMaximo = OpcionesTraza.NivelMaximo
            End If
        End Sub
    End Class

    <Task()> Public Shared Function ExplosionDetallada(ByVal data As StExplosionDetallada, ByVal services As ServiceProvider) As DataSet
        If Not data.NivelMaximo > 0 Then data.NivelMaximo = 100
        Return AdminData.Execute("spBdgVinoExplosionDetallada", 2, False, data.IDVino, data.NivelMaximo)
    End Function

    <Serializable()> _
    Public Class DataExplosionDetalladaArbol
        Public IDVino As Guid
        Public NivelMaximo As Integer
        Public OpcionesTraza As DataOpcionesTraza

        Public ExplosionArbol As List(Of VinoExplosionArbol)
        Public ExplosionArbolDetalle As List(Of VinoExplosionArbolDetalle)

        Public Sub New(ByVal IDVino As Guid, Optional ByVal OpcionesTraza As DataOpcionesTraza = Nothing)
            Me.IDVino = IDVino
            Me.NivelMaximo = 100
            Me.OpcionesTraza = OpcionesTraza
            If Not OpcionesTraza Is Nothing Then
                Me.NivelMaximo = OpcionesTraza.NivelMaximo
            End If
        End Sub
    End Class

    <Task()> Public Shared Function ExplosionDetalladaArbol(ByVal data As DataExplosionDetalladaArbol, ByVal services As ServiceProvider) As DataExplosionDetalladaArbol
        If data.NivelMaximo <= 0 Then data.NivelMaximo = 100
        Dim wineStruct As IMultipleResults
        If Not data.OpcionesTraza Is Nothing Then
            Dim dataContext As New WineDataContext(AdminData.GetConnectionString)
            wineStruct = dataContext.GetExplosionDetalladaArbol(data.IDVino, data.NivelMaximo, _
                                                                   data.OpcionesTraza.VerTratamientos, _
                                                                   data.OpcionesTraza.VerCosteElaboracion, _
                                                                   data.OpcionesTraza.VerCosteEstanciaNave, _
                                                                   data.OpcionesTraza.VerCosteInicial, _
                                                                   data.OpcionesTraza.VerDetalleOperaciones, _
                                                                   data.OpcionesTraza.VerDetalleAnalisis, _
                                                                   data.OpcionesTraza.VerDetalleEntradasUva, _
                                                                   data.OpcionesTraza.VerDetalleEntradasVino, _
                                                                   data.OpcionesTraza.VerDetalleSalidasAV, _
                                                                   data.OpcionesTraza.VerDetalleOtrasEntradas, _
                                                                   data.OpcionesTraza.VerDetalleAjustes, _
                                                                   data.OpcionesTraza.VerDetalleOfs, _
                                                                   data.OpcionesTraza.VerDetalleTransferencias, _
                                                                   data.OpcionesTraza.VerDetalleInventarios, _
                                                                   data.OpcionesTraza.VerDetalleOtrasSalidas, _
                                                                   data.OpcionesTraza.VerPorcentajes)
        Else
            Dim dataContext As New WineDataContext(AdminData.GetConnectionString)
            wineStruct = dataContext.GetExplosionDetalladaArbol(data.IDVino, data.NivelMaximo)
        End If

        If Not wineStruct Is Nothing Then
            Dim lstExplosion As IList(Of VinoExplosionArbol) = wineStruct.GetResult(Of VinoExplosionArbol).ToList()
            Dim lstExplosionDetalle As IList(Of VinoExplosionArbolDetalle) = wineStruct.GetResult(Of VinoExplosionArbolDetalle).ToList()
            data.ExplosionArbol = lstExplosion
            data.ExplosionArbolDetalle = lstExplosionDetalle
            Return data
        End If
    End Function

    <Serializable()> _
    Public Class StExplosionTraza
        Public IDVino As Guid
        Public IDTipoOperacion As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDVino As Guid, ByVal IDTipoOperacion As String)
            Me.IDVino = IDVino
            Me.IDTipoOperacion = IDTipoOperacion
        End Sub
    End Class

    <Task()> Public Shared Function ExplosionTrazabilidad(ByVal data As StExplosionTraza, ByVal service As ServiceProvider) As DataTable
        If Length(data.IDTipoOperacion) > 0 Then
            data.IDTipoOperacion = Quoted(data.IDTipoOperacion)
        Else : data.IDTipoOperacion = "null"
        End If
        Return AdminData.Execute("spBdgVinoExplosionTrazabilidad", False, data.IDVino, data.IDTipoOperacion)
    End Function

    <Serializable()> _
    Public Class StExpTrazaAlbVenta
        Public IDLineaLote As Integer
        Public IDTipoOperacion As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDLineaLote As Integer, Optional ByVal IDTipoOperacion As String = "")
            Me.IDLineaLote = IDLineaLote
            Me.IDTipoOperacion = IDTipoOperacion
        End Sub
    End Class

    <Task()> Public Shared Function ExplosionTrazabilidadAlbaranVenta(ByVal data As StExpTrazaAlbVenta, ByVal services As ServiceProvider) As DataTable
        Dim dtTrazaAV As DataTable = AdminData.Execute("spBdgVinoExplosionTrazabilidadAV", False, data.IDLineaLote)
        dtTrazaAV.TableName = "TrazaAV"
        Dim dtResult As DataTable = dtTrazaAV.Clone
        dtResult.TableName = "Result"
        If Not dtTrazaAV Is Nothing AndAlso dtTrazaAV.Rows.Count > 0 Then
            Dim HT As New Hashtable
            Dim idNew As Integer = dtTrazaAV.Rows.Count
            For Each drTrazaAV As DataRow In dtTrazaAV.Select()
                Dim IDVino As Guid = Nz(drTrazaAV("IDVino"), Guid.Empty)
                If IDVino.Equals(Guid.Empty) Then
                    dtResult.Rows.Add(drTrazaAV.ItemArray)
                Else
                    Dim StExp As New StExplosionTraza(IDVino, data.IDTipoOperacion)
                    Dim dtTraza As DataTable = ProcessServer.ExecuteTask(Of StExplosionTraza, DataTable)(AddressOf ExplosionTrazabilidad, StExp, services)

                    Dim drResult As DataRow = dtResult.NewRow
                    Dim i As Integer = 1
                    For Each drTraza As DataRow In dtTraza.Rows
                        For Each column As DataColumn In dtTrazaAV.Columns
                            If dtTraza.Columns.Contains(column.ColumnName) Then
                                If column.ColumnName = "ID" Then
                                    drResult(column.ColumnName) = idNew + i
                                ElseIf column.ColumnName = "IDpdr" Then
                                    If i > 1 Then
                                        If HT.Contains(drTraza("IDpdr")) Then
                                            drResult(column.ColumnName) = HT(drTraza("IDpdr"))
                                        End If
                                    Else : drResult(column.ColumnName) = drTrazaAV("IDpdr")
                                    End If
                                ElseIf column.ColumnName = "IDLineaLote" Then
                                    drResult(column.ColumnName) = drTrazaAV("IDLineaLote")
                                Else : drResult(column.ColumnName) = drTraza(column.ColumnName)
                                End If
                            End If
                        Next
                        HT(drTraza("ID")) = drResult("ID")
                        dtResult.Rows.Add(drResult.ItemArray)
                        i += 1
                    Next
                    idNew += dtTraza.Rows.Count
                End If
            Next
        End If
        Return dtResult
    End Function

    <Task()> Public Shared Function Variables(ByVal IDVino As Guid, ByVal services As ServiceProvider) As DataTable
        'Return AdminData.Execute("spBdgVinoVariable", False, IDVino)
    End Function

    <Task()> Public Shared Function ObtenerFechaVino(ByVal IDVino As Guid, ByVal services As ServiceProvider) As Date
        Dim dtVino As DataTable = New BdgVino().SelOnPrimaryKey(IDVino)
        If Not dtVino Is Nothing And dtVino.Rows.Count > 0 Then
            Return dtVino.Rows(0)(_V.Fecha)
        Else : Return Date.MinValue
        End If
    End Function

#Region " Origenes nuevos  "

    <Task()> Public Shared Function AniadasGrafo(ByVal IDVino As Guid, ByVal services As ServiceProvider) As List(Of OrigenAniadas)
        Dim dat As New DataExplosionVino(IDVino) ', Origenes.Añadas)
        Dim rstl As DataExplosionVinoResult = ProcessServer.ExecuteTask(Of DataExplosionVino, DataExplosionVinoResult)(AddressOf BdgExplosionVino.OrigenesVino, dat, services)
        Return rstl.Añadas
    End Function

    <Task()> Public Shared Function VariedadesGrafo(ByVal IDVino As Guid, ByVal services As ServiceProvider) As List(Of OrigenVariedades)
        Dim dat As New DataExplosionVino(IDVino) ', Origenes.Variedades)
        Dim rstl As DataExplosionVinoResult = ProcessServer.ExecuteTask(Of DataExplosionVino, DataExplosionVinoResult)(AddressOf BdgExplosionVino.OrigenesVino, dat, services)
        Return rstl.Variedades
    End Function

    <Task()> Public Shared Function FincasGrafo(ByVal IDVino As Guid, ByVal services As ServiceProvider) As List(Of OrigenFincas)
        Dim dat As New DataExplosionVino(IDVino) ', Origenes.Fincas)
        Dim rstl As DataExplosionVinoResult = ProcessServer.ExecuteTask(Of DataExplosionVino, DataExplosionVinoResult)(AddressOf BdgExplosionVino.OrigenesVino, dat, services)
        Return rstl.Fincas
    End Function

    <Task()> Public Shared Function ComprasGrafo(ByVal IDVino As Guid, ByVal services As ServiceProvider) As List(Of OrigenCompras)
        Dim dat As New DataExplosionVino(IDVino) ', Origenes.Compras)
        Dim rstl As DataExplosionVinoResult = ProcessServer.ExecuteTask(Of DataExplosionVino, DataExplosionVinoResult)(AddressOf BdgExplosionVino.OrigenesVino, dat, services)
        Return rstl.Compras
    End Function

    <Task()> Public Shared Function EntradasUVAGrafo(ByVal IDVino As Guid, ByVal services As ServiceProvider) As List(Of OrigenEntradasUVA)
        Dim dat As New DataExplosionVino(IDVino) ', Origenes.EntradaUVA)
        Dim rstl As DataExplosionVinoResult = ProcessServer.ExecuteTask(Of DataExplosionVino, DataExplosionVinoResult)(AddressOf BdgExplosionVino.OrigenesVino, dat, services)
        Return rstl.EntradasUVA
    End Function

#End Region

#Region " Estado del Vino"
    <Serializable()> _
    Public Class StModificarEstado
        Public IDVino As Guid
        Public IDEstadoVino As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDVino As Guid, ByVal IDEstadoVino As String)
            Me.IDVino = IDVino
            Me.IDEstadoVino = IDEstadoVino
        End Sub
    End Class

    <Task()> Public Shared Sub ModificarEstado(ByVal data As StModificarEstado, ByVal services As ServiceProvider)
        If Length(data.IDEstadoVino) > 0 Then
            Dim ClsVino As New BdgVino
            Dim dttVino As DataTable = ClsVino.Filter(New GuidFilterItem("IDVino", data.IDVino))
            If Not IsNothing(dttVino) AndAlso dttVino.Rows.Count > 0 Then
                dttVino.Rows(0)("IDEstadoVino") = data.IDEstadoVino
                ClsVino.Update(dttVino)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarUltimoEstadoVino(ByVal IDVino As Guid, ByVal services As ServiceProvider)
        'Se actualiza el Estado del Vino según su Origen:
        ' - Si es Interno tomando el valor de la última operación realizada.
        ' - Si es de Entradas de Vino tomando el valor de la última entrada realizada.
        Dim dtVino As DataTable = New BdgVino().SelOnPrimaryKey(IDVino)
        If dtVino.Rows.Count > 0 Then
            Dim f As New Filter
            f.Add("IDVino", FilterOperator.Equal, IDVino)
            Dim dt As DataTable
            Select Case dtVino.Rows(0)("Origen")
                Case BdgOrigenVino.Interno
                    dt = New BE.DataEngine().Filter("NegBdgUltimoEstadoVino", f, , "Fecha Desc, NOperacion Desc")
                Case BdgOrigenVino.Compra
                    dt = New BE.DataEngine().Filter("NegBdgUltimoEstadoVinoCompra", f, , "Fecha Desc, NEntrada Desc")
            End Select
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                dtVino.Rows(0)(_V.IDEstadoVino) = dt.Rows(0)(_V.IDEstadoVino)
            Else : dtVino.Rows(0)(_V.IDEstadoVino) = System.DBNull.Value
            End If
            BusinessHelper.UpdateTable(dtVino)
        End If
    End Sub
#End Region

#Region " Coste del Vino"
    <Serializable()> _
    Public Class StCoste
        Public IDVino As Guid
        Public FDesde As Date
        Public FHasta As Date

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDVino As Guid, Optional ByVal FDesde As Date = cnMinDate, Optional ByVal FHasta As Date = cnMinDate)
            Me.IDVino = IDVino
            Me.FDesde = FDesde
            Me.FHasta = FHasta
        End Sub
    End Class

    <Task()> Public Shared Function Coste(ByVal data As StCoste, ByVal services As ServiceProvider) As CosteVino
        If Not data.IDVino.Equals(Guid.Empty) Then
            If data.FHasta = cnMinDate Then
                data.FHasta = DateAdd(DateInterval.Day, 1, Date.Today)
            End If
            If data.FDesde = cnMinDate Then
                data.FDesde = "01/01/1900"
            End If
            Dim costeVino As New Bodega.CosteVino

            Dim datGetCoste As New BdgCosteVino.DataGetVinoCoste(data.IDVino, data.FDesde, data.FHasta, True, True)
            Dim detalleCoste As BdgCosteVino.DataVinoCosteDetalle = ProcessServer.ExecuteTask(Of BdgCosteVino.DataGetVinoCoste, BdgCosteVino.DataVinoCosteDetalle)(AddressOf BdgCosteVino.GetVinoCosteDetalle, datGetCoste, services)
            If Not detalleCoste Is Nothing Then
                costeVino.Materiales = detalleCoste.Materiales
                costeVino.ManoDeObra = detalleCoste.ManoObra
                costeVino.Centros = detalleCoste.Centros
                costeVino.Varios = detalleCoste.Varios
                costeVino.ComprasDeVino = detalleCoste.Compras
                costeVino.Tasas = detalleCoste.Tasas
                costeVino.Elaboracion = detalleCoste.VendimiaElaboracion
                costeVino.Uvas = detalleCoste.EntradaUVA
                costeVino.EstanciaNave = detalleCoste.EstanciaNave
                costeVino.CostesIniciales = detalleCoste.CosteInicial
                costeVino.Totales = detalleCoste.Totales


                'Estancia en Nave
                'Costes Separados
                If costeVino.Totales.Count > 0 Then
                    costeVino.Cantidad = costeVino.Totales(0).QTot
                    costeVino.CosteCompraVino = costeVino.Totales(0).CosteComprasUnit
                    costeVino.CosteMateriales = costeVino.Totales(0).CosteMaterialesUnit
                    costeVino.CosteMod = costeVino.Totales(0).CosteManoObraUnit
                    costeVino.CosteCentros = costeVino.Totales(0).CosteCentroUnit
                    costeVino.CosteUvas = costeVino.Totales(0).CosteUvaUnit
                    costeVino.CosteVarios = costeVino.Totales(0).CosteVariosUnit
                    costeVino.CosteElaboracion = costeVino.Totales(0).CosteVendimiaUnit
                    costeVino.CosteUnitario = costeVino.Totales(0).CosteTotalUnit
                    costeVino.CosteMaterialesMerma = costeVino.Totales(0).CosteMaterialesConMermaUnit
                    costeVino.CosteElaboracionMerma = costeVino.Totales(0).CosteVendimiaConMermaUnit
                    costeVino.CosteEstanciaNave = costeVino.Totales(0).CosteEstanciaNaveUnit
                    costeVino.CosteInicial = costeVino.Totales(0).CosteInicialUnit
                End If

                '******   Costes directos - indirectos - fiscal - variable - fijo
                If costeVino.Varios.Count > 0 Then
                    For Each var As VinoCosteVarios In costeVino.Varios
                        'TipoCosteFV = Fijo
                        costeVino.CosteFijo += Nz(var.CosteFijoUnit, 0)
                        'TipoCosteFV = Variable
                        costeVino.CosteVariable += Nz(var.CosteVariableUnit, 0)
                        'TipoCosteDI = Directo
                        costeVino.CosteDirecto += Nz(var.CosteDirectoUnit, 0)
                        'TipoCosteDI = Indirecto
                        costeVino.CosteIndirecto += Nz(var.CosteInDirectoUnit, 0)
                        costeVino.CosteFiscal += Nz(var.CosteFiscalUnit, 0)
                    Next
                End If

                costeVino.CosteVariable += costeVino.CosteCompraVino + costeVino.CosteMateriales + costeVino.CosteMod + costeVino.CosteElaboracion + costeVino.CosteUvas + costeVino.CosteMaterialesMerma + costeVino.CosteInicial
                costeVino.CosteDirecto += costeVino.CosteCompraVino + costeVino.CosteMateriales + costeVino.CosteMod + costeVino.CosteElaboracion + costeVino.CosteUvas + costeVino.CosteMaterialesMerma + costeVino.CosteInicial
                costeVino.CosteFiscal += costeVino.CosteCompraVino + costeVino.CosteMateriales + costeVino.CosteMod + costeVino.CosteElaboracion + costeVino.CosteUvas + costeVino.CosteMaterialesMerma + costeVino.CosteInicial

                If costeVino.Tasas.Count > 0 Then
                    Dim CFijo As Double = Nz((Aggregate t In costeVino.Tasas Into Sum(t.TasaFijaUnit)), 0) 'Nz((from c In costeVino.dtTasas.ToList   Aggregate    Into Sum(CDbl(c("TasaFijaUnit")))), 0) 'Nz(costeVino.dtTasas.Compute("SUM (TasaFijaUnit)", Nothing), 0)
                    Dim CVariable As Double = Nz((Aggregate t In costeVino.Tasas Into Sum(t.TasaVariableUnit)), 0) ' Nz((from c In costeVino.dtTasas Into Sum(CDbl(c("TasaVariableUnit")))), 0) ' Nz(costeVino.dtTasas.Compute("SUM (TasaVariableUnit)", Nothing), 0)
                    Dim CDirecto As Double = Nz((Aggregate t In costeVino.Tasas Into Sum(t.TasaDirectaUnit)), 0) ' Nz((from c In costeVino.dtTasas Into Sum(CDbl(c("TasaDirectaUnit")))), 0) ' Nz(costeVino.dtTasas.Compute("SUM (TasaDirectaUnit)", Nothing), 0)
                    Dim CIndirecto As Double = Nz((Aggregate t In costeVino.Tasas Into Sum(t.TasaInDirectaUnit)), 0) ' Nz((from c In costeVino.dtTasas Into Sum(CDbl(c("TasaIndirectaUnit")))), 0) ' Nz(costeVino.dtTasas.Compute("SUM (TasaIndirectaUnit)", Nothing), 0)
                    Dim CFiscal As Double = Nz((Aggregate t In costeVino.Tasas Into Sum(t.TasaFiscalUnit)), 0) '  Nz((from c In costeVino.dtTasas Into Sum(CDbl(c("TasaFiscalUnit")))), 0) ' Nz(costeVino.dtTasas.Compute("SUM (TasaFiscalUnit)", Nothing), 0)

                    costeVino.CosteVariable += Nz(CVariable, 0)
                    costeVino.CosteFijo += Nz(CFijo, 0)
                    costeVino.CosteDirecto += Nz(CDirecto, 0)
                    costeVino.CosteIndirecto += Nz(CIndirecto, 0)
                    costeVino.CosteFiscal += Nz(CFiscal, 0)
                End If
            End If
            Return costeVino
        End If
    End Function
#End Region

#Region " Coste Inicial"
    <Task()> Public Shared Function BorrarCosteUnitarioInicial(ByVal dtMarcados As DataTable, ByVal services As ServiceProvider)
        Dim bdgV As New BdgVino
        Dim dtV As DataTable
        Dim ff As New Filter

        For Each drMarc As DataRow In dtMarcados.Rows
            ff.Clear()
            ff.Add(New StringFilterItem(_V.IDVino, FilterOperator.Equal, drMarc(_V.IDVino).ToString))
            dtV = bdgV.Filter(ff) ' SelOnIDDeposito(strIdDeposito)
            If Not dtV Is Nothing And dtV.Rows.Count > 0 Then
                dtV.Rows(0)("CosteUnitarioInicialA") = 0
                dtV.Rows(0)("FechaCosteUnitarioInicial") = System.DBNull.Value
                bdgV.Update(dtV)
            End If
        Next
    End Function

    <Serializable()> _
    Public Class StCrearCosteUnitarioInicial
        Public Fecha As Date
        Public Coste As Double
        Public DtMarcados As DataTable

        Public Sub New()
        End Sub

        Public Sub New(ByVal Fecha As Date, ByVal Coste As Double, ByVal DtMarcados As DataTable)
            Me.Fecha = Fecha
            Me.Coste = Coste
            Me.DtMarcados = DtMarcados
        End Sub
    End Class

    <Task()> Public Shared Function CrearCosteUnitarioInicial(ByVal data As StCrearCosteUnitarioInicial, ByVal services As ServiceProvider)
        If data.DtMarcados.Rows.Count > 0 Then
            For Each drMarcado As DataRow In data.DtMarcados.Select
                Dim StCrear As New StCrearCosteUnitarioInicialReg(data.Fecha, data.Coste, drMarcado("IDVino").ToString)
                ProcessServer.ExecuteTask(Of StCrearCosteUnitarioInicialReg)(AddressOf CrearCosteUnitarioInicialRegistro, StCrear, services)
            Next
        End If
    End Function

    <Serializable()> _
    Public Class StCrearCosteUnitarioInicialReg
        Public Fecha As Date
        Public Coste As Double
        Public Vino As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal Fecha As Date, ByVal Coste As Double, ByVal Vino As String)
            Me.Fecha = Fecha
            Me.Coste = Coste
            Me.Vino = Vino
        End Sub
    End Class

    <Task()> Public Shared Function CrearCosteUnitarioInicialRegistro(ByVal data As StCrearCosteUnitarioInicialReg, ByVal services As ServiceProvider)
        Dim bdgV As New BdgVino 'BdgDepositoVino Para poder asignarlos a vinos q no esten vivos.
        Dim dtV As DataTable
        Dim ff As New Filter
        ff.Add(New StringFilterItem(_DV.IDVino, FilterOperator.Equal, data.Vino))
        dtV = bdgV.Filter(ff) ' SelOnIDDeposito(strIdDeposito)
        If Not dtV Is Nothing And dtV.Rows.Count > 0 Then
            dtV.Rows(0)("CosteUnitarioInicialA") = data.Coste
            dtV.Rows(0)("FechaCosteUnitarioInicial") = data.Fecha
            bdgV.Update(dtV)
        End If
    End Function
#End Region

#Region " Consulta Trazabilidad"
    <Task()> Public Shared Function GetSalidasYStock(ByVal IDVinos() As Guid, ByVal services As ServiceProvider) As DataSet
        Dim strB As Text.StringBuilder = New Text.StringBuilder
        strB.Append("create table #v (IDVino uniqueidentifier)")
        strB.Append(vbCrLf)
        For Each IDVino As Guid In IDVinos
            strB.Append("insert into #v exec spBdgVinoImplosion " & Quoted(IDVino.ToString))
            strB.Append(vbCrLf)
        Next

        'strB.Append("select * from frmBdgTrazabilidadSalidas where Traza in (select IDVino from #v) and IDTipoMovimiento = 2")
        Dim strSelect As String = "select ts.* from frmBdgTrazabilidadSalidas ts"
        strSelect &= " inner join tbBdgVino v"
        strSelect &= " on ts.IDAlmacen = v.IDAlmacen and ts.IDArticulo = v.IDArticulo and ts.Lote = v.Lote and ts.Ubicacion = v.IDDeposito"
        strSelect &= " inner join (select distinct IDVino from #v) tv on tv.IDVino = v.IDVino"
        strSelect &= " where IDTipoMovimiento = 2"
        strB.Append(strSelect)
        strB.Append(vbCrLf)
        'strB.Append("select * from frmBdgTrazabilidadStock where Traza in (select IDVino from #v) and StockFisico > 0")
        strSelect = "select * from frmBdgTrazabilidadStock ts"
        strSelect &= " inner join tbBdgVino v"
        strSelect &= " on ts.IDAlmacen = v.IDAlmacen and ts.IDArticulo = v.IDArticulo and ts.Lote = v.Lote and ts.Ubicacion = v.IDDeposito"
        strSelect &= " inner join (select distinct IDVino from #v) tv on tv.IDVino = v.IDVino"
        strSelect &= " where StockFisico > 0"
        strB.Append(strSelect)
        strB.Append(vbCrLf)
        strB.Append("select distinct IDVino from #v")
        strB.Append(vbCrLf)
        strB.Append("select distinct IDVino from #v where IDVino in (select IDVino from tbBdgDepositoVino)")

        Dim oDa As New System.Data.SqlClient.SqlDataAdapter(strB.ToString, AdminData.GetConnectionString)

        Dim ds As DataSet = New DataSet
        oDa.Fill(ds)

        Dim dtVinos As DataTable = ds.Tables(2)
        ds.Tables.Remove(dtVinos)

        'Lotes
        strB = New Text.StringBuilder
        strB.Append("create table #vc1 (IDVino uniqueidentifier, IDVinoComponente uniqueidentifier)")
        strB.Append(vbCrLf)
        For Each oRw As DataRow In dtVinos.Rows
            strB.Append("insert into #vc1 exec spBdgVinoExplosionSmpl " & Quoted(oRw("IDVino").ToString))
            strB.Append(vbCrLf)
        Next
        strB.Append("create table #vc (IDVino uniqueidentifier, IDVinoComponente uniqueidentifier)")
        strB.Append(vbCrLf)
        For Each oRw As DataRow In dtVinos.Rows
            strB.Append("insert into #vc (IDVino,IDVinoComponente) select distinct IDVino,IDVinoComponente from #vc1")
            strB.Append(vbCrLf)
        Next
        strB.Append(vbCrLf)
        '//Para visualizar esta consulta pegar el texto en el diseñador de vistas y reemplazar #vc por tbBdgVinoEstructura
        strB.Append(" SELECT DISTINCT tbMaestroArticulo.IDArticulo, tbMaestroArticulo.DescArticulo, tbBdgVino.Lote, " & vbCrLf _
                    & " tbBdgVinoMaterial.IDArticulo AS Material, tbMaestroArticulo_1.DescArticulo AS DescMaterial, " & vbCrLf _
                    & " tbBdgVinoMaterialLote.Lote AS LoteMaterial, tbBdgVino_1.Fecha, tbBdgVino_1.NOperacion " & vbCrLf _
                    & " FROM            #vc INNER JOIN tbBdgVino AS tbBdgVino_1 ON #vc.IDVinoComponente = tbBdgVino_1.IDVino INNER JOIN  " & vbCrLf _
                    & " tbMaestroArticulo INNER JOIN tbBdgVino ON tbMaestroArticulo.IDArticulo = tbBdgVino.IDArticulo  " & vbCrLf _
                    & " ON #vc.IDVino = tbBdgVino.IDVino INNER JOIN tbBdgVinoMaterial INNER JOIN  " & vbCrLf _
                    & " tbMaestroArticulo AS tbMaestroArticulo_1 ON tbBdgVinoMaterial.IDArticulo = tbMaestroArticulo_1.IDArticulo ON  " & vbCrLf _
                    & " tbBdgVino_1.IDVino = tbBdgVinoMaterial.IDVino LEFT OUTER JOIN  " & vbCrLf _
                    & " tbBdgVinoMaterialLote ON tbBdgVinoMaterial.IDVinoMaterial = tbBdgVinoMaterialLote.IDVinoMaterial  " & vbCrLf _
                    & "        WHERE(Not (tbBdgVinoMaterialLote.Lote Is NULL))")
        oDa.SelectCommand.CommandText = strB.ToString
        Dim dtLotes As New DataTable
        oDa.Fill(dtLotes)
        ds.Tables.Add(dtLotes)
        Return ds
    End Function

    <Serializable()> _
    Public Class DataGetSalidasYStockGrafo
        Public IDVinos() As Guid
        Public IDArticuloPrecinta As String
        Public NumeroPrecinta As String
        Public LotePrecinta As String
        Public SeriePrecinta As String

        Public Sub New(ByVal IDVinos() As Guid)
            Me.IDVinos = IDVinos
        End Sub
    End Class

    <Serializable()> _
    Public Class DataResultGetSalidasYStockGrafo
        ' Public IDVino As Guid

        Public Vinos As List(Of Vino)
        Public VinosEstructura As List(Of VinoEstructura)

        Public TrazabilidadSalidas As List(Of VinoTrazabilidadSalidas)
        Public TrazabilidadStock As List(Of VinoTrazabilidadStock)
        Public TrazabilidadLotes As List(Of VinoTrazabilidadLotes)

        Public Sub New()
            Me.Vinos = New List(Of Vino)
            Me.VinosEstructura = New List(Of VinoEstructura)
            Me.TrazabilidadSalidas = New List(Of VinoTrazabilidadSalidas)
            Me.TrazabilidadStock = New List(Of VinoTrazabilidadStock)
            Me.TrazabilidadLotes = New List(Of VinoTrazabilidadLotes)
        End Sub
    End Class

    <Task()> Public Shared Function GetSalidasYStockGrafo(ByVal data As DataGetSalidasYStockGrafo, ByVal services As ServiceProvider) As DataResultGetSalidasYStockGrafo
        If data Is Nothing OrElse data.IDVinos Is Nothing Then Exit Function

        Dim rslt As New DataResultGetSalidasYStockGrafo

        Dim IDQueryCoste As Integer = New Random().[Next](Integer.MaxValue)
        ProcessServer.ExecuteTask(Of Integer)(AddressOf BdgCosteVino.DeleteTablaQTotalCostes, IDQueryCoste, services)

        For Each IDVino As Guid In data.IDVinos
            Dim datImplVino As New DataImplosionVino(IDVino)
            datImplVino.IDQueryCoste = IDQueryCoste
            Dim rsltImpl As DataImplosionVinoResult = ProcessServer.ExecuteTask(Of DataImplosionVino, DataImplosionVinoResult)(AddressOf BdgExplosionVino.Implosion, datImplVino, services)
            If Not rsltImpl Is Nothing Then
                If Not rsltImpl.Vinos Is Nothing Then
                    For Each v As Vino In rsltImpl.Vinos
                        rslt.Vinos.Add(v)
                    Next
                End If


                If Not rsltImpl.VinosEstructura Is Nothing Then
                    For Each ve As VinoEstructura In rsltImpl.VinosEstructura
                        rslt.VinosEstructura.Add(ve)
                    Next
                End If


                If Not rsltImpl.TrazabilidadSalidas Is Nothing Then
                    For Each ts As VinoTrazabilidadSalidas In rsltImpl.TrazabilidadSalidas
                        If Not rslt.TrazabilidadSalidas.Contains(ts) Then
                            rslt.TrazabilidadSalidas.Add(ts)
                        End If
                    Next
                End If

                If Not rsltImpl.TrazabilidadStock Is Nothing Then
                    For Each tstk As VinoTrazabilidadStock In rsltImpl.TrazabilidadStock
                        If Not rslt.TrazabilidadStock.Contains(tstk) Then
                            rslt.TrazabilidadStock.Add(tstk)
                        End If
                    Next
                End If

                If Not rsltImpl.TrazabilidadLotes Is Nothing Then
                    For Each tLt As VinoTrazabilidadLotes In rsltImpl.TrazabilidadLotes
                        If Not rslt.TrazabilidadLotes.Contains(tLt) Then
                            rslt.TrazabilidadLotes.Add(tLt)
                        End If
                    Next
                End If
            End If
        Next

        Dim TrazabilidadSegPrecinta As Boolean = (Length(data.IDArticuloPrecinta) > 0 AndAlso Length(data.NumeroPrecinta) > 0)
        If TrazabilidadSegPrecinta Then
            Dim fPrecinta As New Filter
            fPrecinta.Add(New FilterItem("IDArticulo", data.IDArticuloPrecinta))
            If Length(data.LotePrecinta) > 0 Then fPrecinta.Add(New FilterItem("Lote", data.LotePrecinta))
            If Length(data.SeriePrecinta) > 0 Then fPrecinta.Add(New FilterItem("SeriePrecinta", data.SeriePrecinta))
            fPrecinta.Add(New NumberFilterItem("NDesde", FilterOperator.LessThanOrEqual, data.NumeroPrecinta))
            fPrecinta.Add(New NumberFilterItem("NHasta", FilterOperator.GreaterThanOrEqual, data.NumeroPrecinta))
            Dim dtSeguimientoPrecinta As DataTable = New BE.DataEngine().Filter("NegBdgAlbaranVentaSeguimiento", fPrecinta)
            If Not rslt.TrazabilidadSalidas Is Nothing Then
                Dim SalidasPrecintas As List(Of VinoTrazabilidadSalidas) = (From TrazaSalida In rslt.TrazabilidadSalidas Join SegPrecinta In dtSeguimientoPrecinta On TrazaSalida.IDLineaMovimiento Equals SegPrecinta("IDMovimientoSalida") _
                                                                                Select TrazaSalida).ToList
                If Not SalidasPrecintas Is Nothing AndAlso SalidasPrecintas.Count > 0 Then
                    For Each tz As VinoTrazabilidadSalidas In SalidasPrecintas
                        tz.EtiquetaContenida = True
                    Next
                End If
            End If
        End If

        ProcessServer.ExecuteTask(Of Integer)(AddressOf BdgCosteVino.DeleteTablaQTotalCostes, IDQueryCoste, services)

        Return rslt
    End Function

    <Serializable()> _
    Public Class SetGetArbolTratamientosImpl
        Public IDNave As String
        Public IDDeposito As String
        Public IDArticuloVino As String
        Public IDTipoVino As String
        Public IDFamiliaVino As String
        Public IDSubFamiliaVino As String
        Public LoteVino As String
        Public IDMaterial As String
        Public LoteMaterial As String
        Public IDEntrada As Integer
        Public TipoDeposito As Integer

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDNave As String, ByVal IDDeposito As String, ByVal IDArticuloVino As String, ByVal IDTipoVino As String, _
                       ByVal IDFamiliaVino As String, ByVal IDSubFamiliaVino As String, ByVal LoteVino As String, ByVal IDMaterial As String, _
                       ByVal LoteMaterial As String, ByVal IDEntrada As Integer, Optional ByVal TipoDeposito As Integer = -1)
            Me.IDNave = IDNave
            Me.IDDeposito = IDDeposito
            Me.IDArticuloVino = IDArticuloVino
            Me.IDTipoVino = IDTipoVino
            Me.IDFamiliaVino = IDFamiliaVino
            Me.IDSubFamiliaVino = IDSubFamiliaVino
            Me.LoteVino = LoteVino
            Me.IDMaterial = IDMaterial
            Me.LoteMaterial = LoteMaterial
            Me.IDEntrada = IDEntrada
            Me.TipoDeposito = TipoDeposito
        End Sub
    End Class

    <Task()> Public Shared Function GetArbolTratamientosImpl(ByVal data As SetGetArbolTratamientosImpl, ByVal services As ServiceProvider) As Guid()
        Dim fltrVinosActuales As New Filter
        If Len(data.IDDeposito) > 0 Then fltrVinosActuales.Add(New StringFilterItem("IDDeposito", data.IDDeposito))
        If data.TipoDeposito <> -1 Then fltrVinosActuales.Add(New NumberFilterItem("TipoDeposito", data.TipoDeposito))
        If Len(data.IDNave) > 0 Then fltrVinosActuales.Add(New StringFilterItem("IDNave", data.IDNave))
        If Len(data.IDArticuloVino) > 0 Then fltrVinosActuales.Add(New StringFilterItem("IDArticulo", data.IDArticuloVino))
        If Len(data.IDTipoVino) > 0 Then fltrVinosActuales.Add(New StringFilterItem("IDTipo", data.IDTipoVino))
        If Len(data.IDFamiliaVino) > 0 Then fltrVinosActuales.Add(New StringFilterItem("IDFamilia", data.IDFamiliaVino))
        If Len(data.IDSubFamiliaVino) > 0 Then fltrVinosActuales.Add(New StringFilterItem("IDSubfamilia", data.IDSubFamiliaVino))
        If Len(data.LoteVino) > 0 Then fltrVinosActuales.Add(New StringFilterItem("Lote", data.LoteVino))

        Dim dtVinos As DataTable
        Dim dtVinosActuales As DataTable
        If fltrVinosActuales.Count > 0 Then
            dtVinosActuales = AdminData.GetData("negBdgArblTratamientosDepositos", fltrVinosActuales)
        End If

        Dim fltrVinosHistoricos As New Filter
        If Length(data.IDMaterial) > 0 Then fltrVinosHistoricos.Add(New StringFilterItem("IDArticulo", data.IDMaterial))
        If Length(data.LoteMaterial) > 0 Then fltrVinosHistoricos.Add(New StringFilterItem("Lote", data.LoteMaterial))

        Dim fltrVinosEntradas As New Filter
        If data.IDEntrada > 0 Then fltrVinosEntradas.Add(New NumberFilterItem("IDEntrada", data.IDEntrada))


        If fltrVinosHistoricos.Count > 0 Or fltrVinosEntradas.Count > 0 Then
            Dim strB As New Text.StringBuilder
            strB.Append("create table #v (IDVino uniqueidentifier)")
            strB.Append(vbCrLf)

            'Vinos procedentes de las entradas.
            If fltrVinosEntradas.Count > 0 Then
                Dim dtVinosEntradas As DataTable = AdminData.GetData("negBdgArblTratamientosEntradas", fltrVinosEntradas)
                Dim ht As New Hashtable
                For Each oRw As DataRow In dtVinosEntradas.Rows
                    Dim IDVino As Guid = oRw("IDVino")
                    If Not ht.Contains(IDVino) Then
                        strB.Append("insert into #v exec spBdgVinoImplosionMaterial " & Quoted(IDVino.ToString))
                        strB.Append(vbCrLf)
                        ht.Add(IDVino, Nothing)
                    End If
                Next
            End If

            'Vinos procedentes de los materiales.
            If fltrVinosHistoricos.Count > 0 Then
                Dim dtVinosHistoricos As DataTable = AdminData.GetData("negBdgArblTratamientosMaterial", fltrVinosHistoricos)
                Dim ht As New Hashtable
                For Each oRw As DataRow In dtVinosHistoricos.Rows
                    Dim IDVino As Guid = oRw("IDVino")
                    If Not ht.Contains(IDVino) Then
                        strB.Append("insert into #v exec spBdgVinoImplosionMaterial " & Quoted(IDVino.ToString))
                        strB.Append(vbCrLf)
                        ht.Add(IDVino, Nothing)
                    End If
                Next
            End If

            'Vinos seleccionados
            strB.Append("select distinct IDVino from #v")
            strB.Append(vbCrLf)

            Dim SqlCmd As Common.DbCommand = AdminData.GetCommand
            SqlCmd.CommandText = strB.ToString
            dtVinos = AdminData.Execute(SqlCmd, ExecuteCommand.ExecuteReader)
        End If
        If dtVinos Is Nothing Then
            dtVinos = dtVinosActuales
        Else
            If dtVinosActuales Is Nothing Then
            Else
                '//hay que conseguir la intersección de ambas tables
                dtVinosActuales.DefaultView.Sort = _V.IDVino
                For Each oRw As DataRow In dtVinos.Select()
                    If dtVinosActuales.DefaultView.Find(oRw(_V.IDVino)) < 0 Then
                        oRw.Delete()
                    End If
                Next
            End If
        End If

        Dim rslt(-1) As Guid
        If Not dtVinos Is Nothing Then
            dtVinos.DefaultView.Sort = _V.IDVino
            Dim IDAct As Guid
            For Each oRw As DataRowView In dtVinos.DefaultView
                If Not IDAct.Equals(oRw(_V.IDVino)) Then
                    IDAct = oRw(_V.IDVino)
                    ReDim Preserve rslt(rslt.Length)
                    rslt(rslt.Length - 1) = IDAct
                End If
            Next
        End If
        Return rslt
    End Function

    <Serializable()> _
    Public Class StGetArbolTratamientoExpl
        Public IDVino As Guid
        Public PrctjMin As Integer
        Public CantMin As Integer
        Public FechaMin As Date

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDVino As Guid, ByVal PrctjMin As Integer, ByVal CantMin As Integer, ByVal FechaMin As Date)
            Me.IDVino = IDVino
            Me.PrctjMin = PrctjMin
            Me.CantMin = CantMin
            Me.FechaMin = FechaMin
        End Sub
    End Class

    <Task()> Public Shared Function GetArbolTratamientosExpl(ByVal data As StGetArbolTratamientoExpl, ByVal services As ServiceProvider) As DataSet
        'Dim strVino As String = Quoted(data.IDVino.ToString)
        Dim strFechaMin As Date = New Date(1900, 1, 1) '"null"
        If Not data.FechaMin.Equals(Date.MinValue) Then strFechaMin = data.FechaMin '.ToString("yyyyMMdd")
        Return AdminData.Execute("spBdgVinoExplosionMaterial", 2, False, data.IDVino, data.PrctjMin, data.CantMin, strFechaMin)
    End Function

    <Task()> Public Shared Function GetVinosEnExistencias(ByVal IDVinos() As Guid, ByVal services As ServiceProvider) As Guid()
        If IDVinos Is Nothing OrElse IDVinos.Length = 0 Then Exit Function

        Dim IDVinosObj(-1) As Object
        For Each IDVino As Guid In IDVinos
            ReDim Preserve IDVinosObj(IDVinosObj.Length)
            IDVinosObj(IDVinosObj.Length - 1) = IDVino
        Next

        Dim f As New Filter
        f.Add(New FilterItem("Cantidad", FilterOperator.GreaterThan, 0))
        f.Add(New InListFilterItem("IDVino", IDVinosObj, FilterType.Guid))
        Dim dtVinosEx As DataTable = New BdgDepositoVino().Filter(f)
        If dtVinosEx.Rows.Count > 0 Then
            Dim IDVinosExistencias As List(Of Object) = (From v In dtVinosEx Select v("IDVino")).ToList
            Dim IDVinosEx(-1) As Guid
            For Each IDVino As Guid In IDVinosExistencias
                ReDim Preserve IDVinosEx(IDVinosEx.Length)
                IDVinosEx(IDVinosEx.Length - 1) = IDVino
            Next
            Return IDVinosEx
        End If
    End Function
#End Region

#Region " Validar Articulo Vino y Articulo Componente"
    <Task()> Public Shared Function ValidarArticuloVino(ByVal IDArticulo As String, ByVal services As ServiceProvider) As Boolean
        Dim dt As DataTable = New BE.DataEngine().Filter("negBdgValidarArticuloVino", New StringFilterItem(_V.IDArticulo, IDArticulo))
        If dt.Rows.Count = 0 Then
            ApplicationService.GenerateError("Compruebe que el artículo | tenga las marcas de:| - Gestión de Stock.| - Gestión de Stock por Lotes.| - Bodega.", IDArticulo, vbNewLine, vbNewLine, vbNewLine)
        End If
        Return True
    End Function

    <Task()> Public Shared Function ValidarArticuloComponente(ByVal IDArticulo As String, ByVal services As ServiceProvider) As Boolean
        Dim dt As DataTable = New BE.DataEngine().Filter("negBdgValidarArticuloVino", New StringFilterItem(_V.IDArticulo, IDArticulo))
        If dt.Rows.Count > 0 Then
            ApplicationService.GenerateError("Compruebe que el artículo | no sea de Bodega.", IDArticulo)
        End If
        Return True
    End Function

    <Task()> Public Shared Function EsArticuloVino(ByVal IDArticulo As String, ByVal services As ServiceProvider) As Boolean
        Dim blRdo As Boolean = True
        Dim dt As DataTable = New BE.DataEngine().Filter("negBdgValidarArticuloVino", New StringFilterItem(_V.IDArticulo, IDArticulo))
        If dt.Rows.Count = 0 Then
            blRdo = False
        End If
        Return blRdo
    End Function

    <Task()> Public Shared Function EsArticuloComponente(ByVal IDArticulo As String, ByVal services As ServiceProvider) As Boolean
        Dim blRdo As Boolean = True
        Dim dt As DataTable = New BE.DataEngine().Filter("negBdgValidarArticuloVino", New StringFilterItem(_V.IDArticulo, IDArticulo))
        If dt.Rows.Count > 0 Then
            blRdo = False
        End If
        Return blRdo
    End Function

    <Task()> Public Shared Function GetFiltroArticuloVino(ByVal data As Object, ByVal services As ServiceProvider) As Filter
        Dim f As New Filter
        f.Add(New BooleanFilterItem("GestionStock", True))
        f.Add(New BooleanFilterItem("GestionStockPorLotes", True))
        f.Add(New IsNullFilterItem("EnsambladoStock", False))
        f.Add(New IsNullFilterItem("GestionStockPorLotes", False))
        Return f
    End Function
#End Region

#Region " Actualización de Lotes con Barrica "

    <Task()> Public Shared Function GetVinosSinLoteActualizado(ByVal data As Object, ByVal services As ServiceProvider) As DataSet
        Dim ds As New DataSet
        ds.Tables.Add(ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf GetVinosHistoricos, New Object, services))
        ds.Tables.Add(ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf GetVinosActualesSinBarrica, New Object, services))
        ds.Tables.Add(ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf GetVinosActualesConBarrica, New Object, services))
        Return ds
    End Function

    <Task()> Public Shared Function GetVinosHistoricos(ByVal data As Object, ByVal services As ServiceProvider) As DataTable
        Return AdminData.Filter("vBdgNegVinosHistoricosConLoteSinBarrica")
    End Function

    <Task()> Public Shared Function GetVinosActualesSinBarrica(ByVal data As Object, ByVal services As ServiceProvider) As DataTable
        Return AdminData.Filter("vBdgNegVinosActualesConLoteSinBarrica")
    End Function

    <Task()> Public Shared Function GetVinosActualesConBarrica(ByVal data As Object, ByVal services As ServiceProvider) As DataTable
        Return AdminData.Filter("vBdgNegVinosActualesConLoteConBarrica")
    End Function

    <Task()> Public Shared Function UpdateVinosHistoricosSinBarrica(ByVal data As Object, ByVal services As ServiceProvider) As DataTable
        Dim strB As Text.StringBuilder = New Text.StringBuilder
        strB.Append("DECLARE @Vinos TABLE (IDVino uniqueidentifier NOT NULL, IDDeposito varchar(25) NULL, IDArticulo varchar(25) NULL, Lote varchar(25) NULL, IDBarrica varchar(25) NULL)")
        strB.Append(vbCrLf)
        strB.Append("INSERT INTO @Vinos (IDVino, IDDeposito, IDArticulo, Lote, IDBarrica) SELECT * FROM vBdgNegVinosHistoricosConLoteSinBarrica")
        strB.Append(vbCrLf)
        strB.Append("UPDATE tbBdgVino SET Lote = IDBarrica WHERE IDVino IN (SELECT IDVino FROM @Vinos)")
        strB.Append(vbCrLf)
        strB.Append("SELECT * FROM @Vinos")

        Dim SqlCmd As Common.DbCommand = AdminData.GetCommand
        SqlCmd.CommandText = strB.ToString
        Return AdminData.Execute(SqlCmd, ExecuteCommand.ExecuteReader)
    End Function

    <Task()> Public Shared Sub UpdateVinosActualesSinBarrica(ByVal dtVinos As DataTable, ByVal services As ServiceProvider)
        If Not dtVinos Is Nothing Then
            Dim ClsVino As New BdgVino
            ClsVino.Update(dtVinos)
        End If
    End Sub

    <Task()> Public Shared Sub UpdateVinosActualesConBarrica(ByVal dtVinosSeleccionados As DataTable, ByVal services As ServiceProvider)
        If Not dtVinosSeleccionados Is Nothing AndAlso dtVinosSeleccionados.Rows.Count > 0 Then
            AdminData.BeginTx()
            dtVinosSeleccionados.TableName = "BdgVino"
            For Each dr As DataRow In dtVinosSeleccionados.Rows
                If dr("Dif") = 0 Then
                    ProcessServer.ExecuteTask(Of DataRow)(AddressOf ActualizarLoteEnAlmacen, dr, services)
                    dr("Lote") = dr("IDBarrica")
                End If
            Next
            Dim ClsVino As New BdgVino
            ClsVino.Update(dtVinosSeleccionados)
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarLoteEnAlmacen(ByVal dr As DataRow, ByVal services As ServiceProvider)
        Dim NumeroMovimiento As Integer = ProcessServer.ExecuteTask(Of Object, Integer)(AddressOf ProcesoStocks.NuevoNumeroMovimiento, Nothing, services)
        Dim data As New StockData

        data.Articulo = dr("IDArticulo")
        data.Almacen = dr("IDAlmacen")
        data.Lote = dr("Lote")
        data.Ubicacion = dr("IDDeposito")
        data.Traza = dr("IDVino")
        data.Cantidad = Nz(dr("Cantidad"), 0)
        data.TipoMovimiento = enumTipoMovimiento.tmSalAjuste
        data.Documento = dr("NOperacion")
        data.FechaDocumento = Today
        data.Texto = "Lote=TipoBarrica"

        Dim dataSalida As New DataNumeroMovimientoSinc(NumeroMovimiento, data, False)
        ProcessServer.ExecuteTask(Of DataNumeroMovimientoSinc)(AddressOf ProcesoStocks.Salida, dataSalida, services)

        data.Lote = dr("IDBarrica")
        data.TipoMovimiento = enumTipoMovimiento.tmEntAjuste

        Dim dataEntrada As New DataNumeroMovimientoSinc(NumeroMovimiento, data, False)
        ProcessServer.ExecuteTask(Of DataNumeroMovimientoSinc)(AddressOf ProcesoStocks.Entrada, dataEntrada, services)
    End Sub

#End Region

#Region " Inventario de Bodega en Fecha "

    <Serializable()> _
    Public Class StDepositosEnFecha
        Public Fecha As Date
        Public FiltroWhere As Filter
        Public Valorado As Boolean = False

        Public Sub New()

        End Sub

        Public Sub New(ByVal Fecha As Date, ByVal FiltroWhere As Filter, ByVal Valorado As Boolean)
            Me.Fecha = Fecha
            Me.FiltroWhere = FiltroWhere
            Me.Valorado = Valorado
        End Sub
    End Class

    <Task()> Public Shared Function DepositosEnFecha(ByVal data As StDepositosEnFecha, ByVal services As ServiceProvider) As DataTable
        Dim dtRdo As DataTable = AdminData.Execute("spBdgDepositosEnFecha", False, data.Fecha, AdminData.ComposeFilter(data.FiltroWhere) & String.Empty)
        If data.Valorado Then
            If Not IsNothing(dtRdo) And dtRdo.Rows.Count > 0 Then
                For Each drVino As DataRow In dtRdo.Rows
                    Dim datVinoCst As New BdgCosteVino.DataGetVinoCoste(drVino("IDVino"), cnMinDate, data.Fecha, True, True)
                    datVinoCst.IDQueryCoste = New Random().[Next](Integer.MaxValue)
                    Dim CosteTot As IList(Of VinoCosteTotal) = ProcessServer.ExecuteTask(Of BdgCosteVino.DataGetVinoCoste, IList(Of VinoCosteTotal))(AddressOf BdgCosteVino.GetVinoCosteTotales, datVinoCst, services)
                    If Not CosteTot Is Nothing AndAlso CosteTot.Count > 0 Then
                        drVino("CosteUnitario") = CosteTot(0).CosteTotalUnit
                        drVino("Importe") = CosteTot(0).CosteTotalUnit * drVino("QFecha")
                    End If
                Next
            End If
        End If
        Return dtRdo
    End Function

#End Region

#Region " Gestión Fuera de Inventario "

    <Task()> Public Shared Sub VinoFueraInventario(ByVal IDVino As Guid, ByVal services As ServiceProvider)
        Dim p As New BdgParametro
        If p.BdgGestionFueraInventario Then
            Dim v As New BdgVino
            Dim dtVino As DataTable = v.SelOnPrimaryKey(IDVino)
            If dtVino.Rows.Count > 0 Then
                dtVino.Rows(0)("FueraInventario") = True
                dtVino.Rows(0)("FechaFueraInventario") = Today
                BusinessHelper.UpdateTable(dtVino)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub VinoDentroInventario(ByVal IDVino As Guid, ByVal services As ServiceProvider)
        Dim p As New BdgParametro
        If p.BdgGestionFueraInventario Then
            Dim v As New BdgVino
            Dim dtVino As DataTable = v.SelOnPrimaryKey(IDVino)
            If dtVino.Rows.Count > 0 Then
                dtVino.Rows(0)("FueraInventario") = False
                BusinessHelper.UpdateTable(dtVino)
            End If
        End If
    End Sub

#End Region
#End Region

End Class

<Serializable()> _
Public Class VinoSubDep
    Public IDSubDep As String
    Public Cantidad As Double
End Class

<Serializable()> _
Public Class VinoComponente
    Public IDVino As Guid
    Public Cantidad As Double
    Public Merma As Double
    Public Factor As Double
    Public SubDeps() As VinoSubDep

    Public Sub New()

    End Sub

    Public Sub New(ByVal pIDVino As Guid, ByVal pCantidad As Double, ByVal pFactor As Double)
        IDVino = pIDVino
        Cantidad = pCantidad
        Factor = pFactor
    End Sub
End Class

<Serializable()> _
Public Class CosteVino
    Friend Sub New()
    End Sub

    Public IDVino As Guid
    Public Cantidad As Double

    Public CosteMateriales As Double
    Public CosteCentros As Double
    Public CosteEstanciaNave As Double
    Public CosteUvas As Double
    Public CosteMod As Double
    Public CosteVarios As Double
    Public CosteCompraVino As Double
    Public CosteElaboracion As Double
    Public CosteInicial As Double

    Public CosteMaterialesMerma As Double
    Public CosteElaboracionMerma As Double

    Public CosteFijo As Double
    Public CosteVariable As Double
    Public CosteDirecto As Double
    Public CosteIndirecto As Double

    Public CosteUnitario As Double
    Public CosteFiscal As Double

    Public Materiales As IList(Of VinoCosteMateriales)
    Public ManoDeObra As IList(Of VinoCosteMOD)
    Public Centros As IList(Of VinoCosteCentros)
    Public Uvas As IList(Of VinoCosteEntradaUVA)
    Public Varios As IList(Of VinoCosteVarios)
    Public ComprasDeVino As IList(Of VinoCosteCompras)
    Public Tasas As IList(Of VinoCosteTasas)
    Public Elaboracion As IList(Of VinoCosteVendimiaElaboracion)
    Public EstanciaNave As IList(Of VinoCosteEstanciaEnNave)
    Public CostesIniciales As IList(Of VinoCosteInicial)
    Public Totales As IList(Of VinoCosteTotal)

    Public Function CosteProduccion() As Double
        Return CosteCompraVino + CosteElaboracion + CosteMateriales + CosteMaterialesMerma + CosteCentros + CosteMod + CosteVarios
    End Function
End Class

<Serializable()> _
Public Class _V
    Public Const IDVino As String = "IDVino"
    Public Const IDDeposito As String = "IDDeposito"
    Public Const TipoDeposito As String = "TipoDeposito"
    Public Const IDArticulo As String = "IDArticulo"
    Public Const Lote As String = "Lote"
    Public Const Fecha As String = "Fecha"
    Public Const IDEstadoVino As String = "IDEstadoVino"
    Public Const Origen As String = "Origen"
    Public Const FechaInicio As String = "FechaInicio"
    Public Const FechaFin As String = "FechaFin"

    'TODO puede sobrar
    Public Const NOperacion As String = "NOperacion"
    Public Const IDUdMedida As String = "IDUdMedida"
    Public Const CosteTotal As String = "CosteTotal"
    Public Const CosteFiscal As String = "CosteFiscal"
    Public Const CosteVariable As String = "CosteVariable"
    Public Const IDBarrica As String = "IDBarrica"
    Public Const IDAlmacen As String = "IDAlmacen"
    Public Const DiasDeposito As String = "DiasDeposito"
    Public Const DiasBarrica As String = "DiasBarrica"
    Public Const DiasBotellero As String = "DiasBotellero"
    Public Const QTotal As String = "QTotal"
End Class

Public Enum BdgOrigenVino
    Interno = 0
    Uva = 1
    Compra = 2
    AlbaranTransferencia = 3
End Enum

Public Enum Origenes
    Variedades = 0
    Fincas = 1
    Compras = 2
    Añadas = 3
    EntradaUVA = 4
End Enum