Public Class BdgVinoMaterial

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgVinoMaterial"

#End Region

#Region "Variables Privadas"

    Private _HM As _HistoricoMovimiento
    Private _AA As _ArticuloAlmacen
    Private _AAL As _ArticuloAlmacenLote

#End Region

#Region "Eventos Entidad"

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf Comunes.BeginTransaction)
        deleteProcess.AddTask(Of DataRow)(AddressOf EliminarMaterialSinLotes)
        deleteProcess.AddTask(Of DataRow)(AddressOf EliminarLotes)
    End Sub

    <Task()> Public Shared Sub EliminarMaterialSinLotes(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Not data Is Nothing Then
            If Length(data("IDLineaMovimiento")) > 0 Then
                Dim dataCorreccion As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Eliminar, data("IDLineaMovimiento"), False)
                ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento)(AddressOf ProcesoStocks.ActualizarMovimiento, dataCorreccion, services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub EliminarLotes(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim filter As New Filter
        filter.Add(New GuidFilterItem("IDVinoMaterial", data("IDVinoMaterial")))
        Dim bsnVML As New BdgVinoMaterialLote
        Dim dttVML As DataTable = bsnVML.Filter(filter)
        bsnVML.Delete(dttVML)
    End Sub

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataTable)(AddressOf DescuentoMateriales)
    End Sub

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim OBrl As New BusinessRules
        OBrl.Add("IDArticulo", AddressOf BdgGeneral.CambioMaterialVinoMaterial)
        OBrl.Add("Cantidad", AddressOf BdgGeneral.CambioCantidadVinoMaterial)
        OBrl.Add("Merma", AddressOf BdgGeneral.CambioMermaVinoMaterial)
        OBrl.Add("IDAlmacen", AddressOf BdgGeneral.CambioAlmacenVinoMaterial)
        Return OBrl
    End Function

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Sub CambioLotes(ByVal dtVinoMaterial As DataTable, ByVal services As ServiceProvider)
        If Not dtVinoMaterial Is Nothing AndAlso dtVinoMaterial.Rows.Count > 0 Then
            AdminData.BeginTx()
            'deshacemos los movimientos de cada uno de los lotes, dejando el idlineamovimiento a nulo para que el proceso de materiales lo recalcule
            ProcessServer.ExecuteTask(Of DataTable)(AddressOf DeshacerDescuentosMateriales, dtVinoMaterial, services)

            AdminData.CommitTx()
        End If
    End Sub

    <Task()> Public Shared Sub CambioLote(ByVal dtVinoMaterial As DataTable, ByVal services As ServiceProvider)
        If Not dtVinoMaterial Is Nothing AndAlso dtVinoMaterial.Rows.Count > 0 Then
            AdminData.BeginTx()
            'Se elimina el Movimiento que se ha modificado
            ProcessServer.ExecuteTask(Of DataTable)(AddressOf DeshacerDescuentosMateriales, dtVinoMaterial, services)

            'Se pone el IdLineaMovimiento a Null, para cuando lo trate en DescuentoMaterial 
            'lo trate como si fuera nuevo y lo añada al movimiento de stock
            For Each dr As DataRow In dtVinoMaterial.Rows
                dr("IDLineaMovimiento") = DBNull.Value
            Next
            'también se ponen a nulos los valores de IDLineaMovimiento
            ProcessServer.ExecuteTask(Of DataTable)(AddressOf DescuentoMateriales, dtVinoMaterial, services)
            Dim ClsVinoMat As New BdgVinoMaterial
            ClsVinoMat.Update(dtVinoMaterial)
            AdminData.CommitTx()
        End If
    End Sub

    <Serializable()> _
    Public Class StCrearVinoMaterial
        Public Fecha As Date
        Public NOperacion As String
        Public Data As DataTable

        Public Sub New()
        End Sub

        Public Sub New(ByVal Fecha As Date, ByVal NOperacion As String, ByVal Data As DataTable)
            Me.Fecha = Fecha
            Me.NOperacion = NOperacion
            Me.Data = Data
        End Sub
    End Class

    <Task()> Public Shared Sub CrearVinoMaterial(ByVal data As StCrearVinoMaterial, ByVal services As ServiceProvider)
        If Not data.Data Is Nothing AndAlso data.Data.Rows.Count > 0 Then
            Dim ClsVinoMat As New BdgVinoMaterial
            Dim dtVM As DataTable = ClsVinoMat.AddNew
            For Each oRw As DataRow In data.Data.Select
                Dim rwVM As DataRow = dtVM.NewRow
                If (oRw.Table.Columns.Contains(_VM.IDVinoMaterial)) Then
                    rwVM(_VM.IDVinoMaterial) = oRw(_VM.IDVinoMaterial)
                End If
                rwVM(_VM.IDVino) = oRw(_VM.IDVino)
                rwVM(_VM.IDArticulo) = oRw(_VM.IDArticulo)
                rwVM(_VM.Fecha) = data.Fecha
                rwVM(_VM.NOperacion) = data.NOperacion
                rwVM(_VM.Cantidad) = oRw(_VM.Cantidad)
                rwVM(_VM.Merma) = oRw(_VM.Merma)
                rwVM(_VM.Precio) = oRw(_VM.Precio)
                If Length(oRw(_VM.IDAlmacen)) = 0 Then
                    ApplicationService.GenerateError("No se ha asignado ningún almacén para el descuento de materiales del artículo |.", Quoted(oRw(_VM.IDArticulo)))
                Else : rwVM(_VM.IDAlmacen) = oRw(_VM.IDAlmacen)
                End If
                rwVM(_VM.RecalcularMaterial) = oRw(_VM.RecalcularMaterial)
                rwVM(_VM.IDOperacionMaterial) = oRw(_VM.IDOperacionMaterial)
                'rwVM("Lote") = oRw("Lote")
                'rwVM("Ubicacion") = oRw("Ubicacion")
                dtVM.Rows.Add(rwVM)
            Next
            ClsVinoMat.Update(dtVM)
        End If
    End Sub

    <Serializable()> _
    Public Class StBorrarCosteVendimiaVinoAgrup
        Public Vendimia As String
        Public IDArticulo As String
        Public DtMarcados As DataTable

        Public Sub New()
        End Sub

        Public Sub New(ByVal Vendimia As String, ByVal IDArticulo As String, ByVal DtMarcados As DataTable)
            Me.Vendimia = Vendimia
            Me.IDArticulo = IDArticulo
            Me.DtMarcados = DtMarcados
        End Sub
    End Class

    <Task()> Public Shared Sub BorrarCosteVendimiaVinoAgrup(ByVal data As StBorrarCosteVendimiaVinoAgrup, ByVal services As ServiceProvider)
        Dim clsHis As New BdgCosteVendimiaHist
        Dim dtCosteVendimiaHist As DataTable
        Dim ff As New Filter
        Dim Values(-1) As Object
        Dim Vinos(-1) As Object
        Dim Depositos(-1) As Object
        Dim IDVinoMaterial(-1) As Object

        ff.Add("Vendimia", FilterOperator.Equal, data.Vendimia)
        'Recopilar los Depositos q se quieren borrar
        For Each drMarc As DataRow In data.DtMarcados.Rows
            ReDim Preserve Depositos(UBound(Depositos) + 1)
            Depositos(UBound(Depositos)) = drMarc("IDDeposito")
            ReDim Preserve Vinos(UBound(Vinos) + 1)
            Vinos(UBound(Vinos)) = drMarc("IDVino")
        Next
        ff.Add(New InListFilterItem("IDDeposito", Depositos, FilterType.String))
        ff.Add(New InListFilterItem("IDVino", Vinos, FilterType.Guid))
        If Length(data.IDArticulo) > 0 AndAlso data.IDArticulo <> String.Empty Then
            ff.Add("IDArticulo", FilterOperator.Equal, data.IDArticulo)
        End If
        Dim dtLista As DataTable = New BE.DataEngine().Filter("vNegBdgBorrarCosteVinoMaterial", ff, , "Vendimia")

        If Not IsNothing(dtLista) AndAlso dtLista.Rows.Count > 0 Then
            'Recoger los IDVinoMaterial
            For Each drCosteVendimiaHist As DataRow In dtLista.Rows
                ReDim Preserve IDVinoMaterial(UBound(IDVinoMaterial) + 1)
                IDVinoMaterial(UBound(IDVinoMaterial)) = drCosteVendimiaHist("IDVinoMaterial")
            Next

            ff.Clear()
            ff.Add("Vendimia", FilterOperator.Equal, data.Vendimia)
            ff.Add(New InListFilterItem("IDVinoMaterial", IDVinoMaterial, FilterType.Guid))
            dtCosteVendimiaHist = clsHis.Filter(ff)
            If Not IsNothing(dtCosteVendimiaHist) AndAlso dtCosteVendimiaHist.Rows.Count > 0 Then 'Borrar el historico
                clsHis.Delete(dtCosteVendimiaHist)
            End If
            'FIN Recoger los Depositos

            'Borrar IDVinoMaterial
            If IDVinoMaterial.Length > 0 Then
                ff.Clear()
                ff.Add(New InListFilterItem("IDVinoMaterial", IDVinoMaterial, FilterType.Guid))
                Dim ClsMater As New BdgVinoMaterial
                Dim dtMater As DataTable = ClsMater.Filter(ff)
                If Not IsNothing(dtMater) AndAlso dtMater.Rows.Count > 0 Then 'Borrar VinoMaterial
                    AdminData.BeginTx()
                    ProcessServer.ExecuteTask(Of DataTable)(AddressOf DeshacerDescuentosMateriales, dtMater, services)
                    'ProcessServer.ExecuteTask(Of DataTable)(AddressOf DeshacerDescuentoMateriales, dtMater, services)
                    ClsMater.Delete(dtMater)
                End If
            End If
            'Fin Borrar IDVinoMaterial
        End If
    End Sub

    <Serializable()> _
    Public Class StCrearCosteVendimiaVinoAgrup
        Public Fecha As Date
        Public IDArticulo As String
        Public Coste As Double
        Public Vendimia As String
        Public DtMarcados As DataTable

        Public Sub New()
        End Sub

        Public Sub New(ByVal Fecha As Date, ByVal IDArticulo As String, ByVal Coste As Double, ByVal Vendimia As String, ByVal DtMarcados As DataTable)
            Me.Fecha = Fecha
            Me.IDArticulo = IDArticulo
            Me.Coste = Coste
            Me.Vendimia = Vendimia
            Me.DtMarcados = DtMarcados
        End Sub
    End Class

    <Task()> Public Shared Function CrearCosteVendimiaVinoAgrup(ByVal data As StCrearCosteVendimiaVinoAgrup, ByVal services As ServiceProvider)
        If Not data.DtMarcados Is Nothing AndAlso data.DtMarcados.Rows.Count > 0 Then
            For Each drMarcado As DataRow In data.DtMarcados.Select
                Dim StCrearCoste As New StCrearCosteVendimiaVino(data.Fecha, drMarcado("IDDeposito"), data.IDArticulo, data.Coste, data.Vendimia, drMarcado("IDVino").ToString)
                ProcessServer.ExecuteTask(Of StCrearCosteVendimiaVino, DataTable)(AddressOf CrearCosteVendimiaVino, StCrearCoste, services)
            Next
        End If
    End Function

    <Serializable()> _
    Public Class StCrearCosteVendimiaVino
        Public Fecha As Date
        Public IDDeposito As String
        Public IDArticulo As String
        Public Coste As Double
        Public Vendimia As String
        Public Vino As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal Fecha As Date, ByVal IDDeposito As String, ByVal IDArticulo As String, ByVal Coste As Double, ByVal Vendimia As String, ByVal Vino As String)
            Me.Fecha = Fecha
            Me.IDDeposito = IDDeposito
            Me.IDArticulo = IDArticulo
            Me.Coste = Coste
            Me.Vendimia = Vendimia
            Me.Vino = Vino
        End Sub
    End Class

    <Task()> Public Shared Function CrearCosteVendimiaVino(ByVal data As StCrearCosteVendimiaVino, ByVal services As ServiceProvider) As DataTable
        Dim bdgDV As New BdgVino 'BdgDepositoVino Para poder asignarlos a vinos q no esten vivos.
        Dim ff As New Filter
        ff.Add(New StringFilterItem(_DV.IDDeposito, FilterOperator.Equal, data.IDDeposito))
        ff.Add(New StringFilterItem(_DV.IDVino, FilterOperator.Equal, data.Vino))
        Dim dtDV As DataTable = bdgDV.Filter(ff) ' SelOnIDDeposito(strIdDeposito)
        If Not dtDV Is Nothing And dtDV.Rows.Count > 0 Then
            Dim QTot As Double = ProcessServer.ExecuteTask(Of Guid, Double)(AddressOf BdgWorkClass.QTotObtener, dtDV.Rows(0)("IdVino"), services)

            Dim aux As Guid = Guid.NewGuid

            Dim ClsVinoMat As New BdgVinoMaterial
            Dim dtVM As DataTable = ClsVinoMat.AddNew
            Dim rwVM As DataRow = dtVM.NewRow

            rwVM(_VM.IDVinoMaterial) = aux
            Dim IdVino As Guid = dtDV.Rows(0)("IdVino")
            rwVM(_VM.IDVino) = data.Vino
            rwVM(_VM.Cantidad) = QTot  'dtDV.Rows(0)("Cantidad")
            rwVM(_VM.Fecha) = data.Fecha
            rwVM(_VM.Precio) = data.Coste
            rwVM(_VM.IDArticulo) = data.IDArticulo
            dtVM.Rows.Add(rwVM)

            dtVM = ClsVinoMat.Update(dtVM)

            If Not dtVM Is Nothing And dtVM.Rows.Count > 0 Then
                Dim StHist As New BdgCosteVendimiaHist.StCrearCosteVendimiaHistInd(data.Vendimia, data.IDDeposito, aux, New Guid(data.Vino), data.IDArticulo)
                ProcessServer.ExecuteTask(Of BdgCosteVendimiaHist.StCrearCosteVendimiaHistInd)(AddressOf BdgCosteVendimiaHist.CrearCosteVendimiaHistInd, StHist, services)
            End If
            Return dtVM
        End If
    End Function

    <Task()> Public Shared Sub DescuentoMateriales(ByVal VinoMaterial As DataTable, ByVal services As ServiceProvider)
        If Not VinoMaterial Is Nothing AndAlso VinoMaterial.Rows.Count > 0 Then
            Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()

            For Each dr As DataRow In VinoMaterial.Select
                'TODO - no estoy seguro si habría que vacíar el campo idmovimiento para DescuentoMaterialesLote
                Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(dr("IDArticulo"))
                If ArtInfo.GestionStockPorLotes Then
                    ProcessServer.ExecuteTask(Of DataRow)(AddressOf DescuentoMaterialesLote, dr, services)
                Else : ProcessServer.ExecuteTask(Of DataRow)(AddressOf DescuentoMaterialesLinea, dr, services)
                End If
            Next
        End If
    End Sub

    'TODO - Refactor?
    <Task()> Public Shared Sub DescuentoMaterialesLote(ByVal VinoMaterialRow As DataRow, ByVal services As ServiceProvider)
        Dim almacenPredeterminado As String = New Parametro().AlmacenPredeterminado
        Dim bsnVML As BusinessHelper = BusinessHelper.CreateBusinessObject("BdgVinoMaterialLote")
        Dim filter As New Filter
        filter.Add(New GuidFilterItem(_BdgVinoMaterialLote.IDVinoMaterial, VinoMaterialRow(_VM.IDVinoMaterial)))
        Dim dttVinoMaterialLote As DataTable = bsnVML.Filter(filter)

        If VinoMaterialRow.HasVersion(DataRowVersion.Original) AndAlso (VinoMaterialRow(_VM.Cantidad) <> VinoMaterialRow(_VM.Cantidad, DataRowVersion.Original) _
                        OrElse Nz(VinoMaterialRow(_VM.Merma), 0) <> Nz(VinoMaterialRow(_VM.Merma, DataRowVersion.Original), 0) _
                        OrElse Nz(VinoMaterialRow(_VM.Precio), 0) <> Nz(VinoMaterialRow(_VM.Precio, DataRowVersion.Original), 0)) Then
            'BORRADO DE LOTES EXISTENTES
            bsnVML.Delete(dttVinoMaterialLote) 'ESTO CORREGIRÁ TAMBIÉN LOS MOVIMIENTOS
        Else
            'HAY QUE HACER NUEVOS MOVIMIENTOS EN CASO DE QUE LOS LOTES NO LO TENGAN ASOCIADO
            'EN CASO CONTRARIO ESTÁ CORREGIDO ARRIBA
            Dim bsnAML As BusinessHelper = BusinessHelper.CreateBusinessObject("ArticuloAlmacenLote")


            For Each VinoMaterialLoteRow As DataRow In dttVinoMaterialLote.Rows
                If (Length(VinoMaterialLoteRow(_BdgVinoMaterialLote.IDLineaMovimiento)) = 0) Then
                    Dim NumeroMovimiento As Integer = ProcessServer.ExecuteTask(Of Object, Integer)(AddressOf ProcesoStocks.NuevoNumeroMovimiento, Nothing, services)
                    Dim Data As New StockData
                    Data.Articulo = VinoMaterialRow(_VM.IDArticulo)
                    If (Not VinoMaterialRow.IsNull(_VM.IDAlmacen)) Then
                        Data.Almacen = VinoMaterialRow(_VM.IDAlmacen)
                    Else
                        Data.Almacen = almacenPredeterminado
                    End If

                    Data.Cantidad = VinoMaterialLoteRow(_BdgVinoMaterialLote.Cantidad)
                    Data.Lote = VinoMaterialLoteRow(_BdgVinoMaterialLote.Lote)
                    Data.Ubicacion = VinoMaterialLoteRow(_BdgVinoMaterialLote.Ubicacion)
                    Data.TipoMovimiento = enumTipoMovimiento.tmSalFabrica
                    Data.PrecioA = Nz(VinoMaterialRow(_VM.Precio))
                    If (Length(VinoMaterialRow(_VM.NOperacion)) > 0) Then
                        Data.Documento = VinoMaterialRow(_VM.NOperacion)
                    End If
                    If (Not IsDate(VinoMaterialRow(_VM.Fecha)) OrElse VinoMaterialRow(_VM.Fecha) = cnMinDate) Then
                        Data.FechaDocumento = Today
                    Else
                        Data.FechaDocumento = New Date(CDate(VinoMaterialRow(_VM.Fecha)).Year, CDate(VinoMaterialRow(_VM.Fecha)).Month, CDate(VinoMaterialRow(_VM.Fecha)).Day)
                    End If

                    If (Length(VinoMaterialLoteRow(_BdgVinoMaterialLote.NDesdePrecinta)) > 0) Then _
                        Data.PrecintaUtilizadaDesde = VinoMaterialLoteRow(_BdgVinoMaterialLote.NDesdePrecinta)
                    If (Length(VinoMaterialLoteRow(_BdgVinoMaterialLote.NHastaPrecinta)) > 0) Then _
                        Data.PrecintaUtilizadaHasta = VinoMaterialLoteRow(_BdgVinoMaterialLote.NHastaPrecinta)
                    If (Length(VinoMaterialLoteRow(_BdgVinoMaterialLote.SeriePrecinta)) > 0) Then _
                        Data.PrecintaNSerie = VinoMaterialLoteRow(_BdgVinoMaterialLote.SeriePrecinta)
                    Dim f As New Filter
                    f.Add("Lote", VinoMaterialLoteRow(_BdgVinoMaterialLote.Lote))
                    f.Add("IDAlmacen", Data.Almacen)
                    f.Add("SeriePrecinta", Data.PrecintaNSerie)
                    f.Add("IDArticulo", Data.Articulo)
                    Dim artalmlote As DataTable = bsnAML.Filter(f)
                    If (Not artalmlote Is Nothing AndAlso artalmlote.Rows.Count > 0) Then
                        Data.PrecintaDesde = artalmlote.Rows(0)("NDesdePrecinta")
                        Data.PrecintaHasta = artalmlote.Rows(0)("NHastaPrecinta")
                    End If


                    Dim dataSalida As New DataNumeroMovimientoSinc(NumeroMovimiento, Data, False)
                    Dim updateData As StockUpdateData = ProcessServer.ExecuteTask(Of DataNumeroMovimientoSinc, StockUpdateData)(AddressOf ProcesoStocks.Salida, dataSalida, services)
                    If Not updateData Is Nothing Then
                        If updateData.Estado = EstadoStock.Actualizado Then
                            VinoMaterialLoteRow(_BdgVinoMaterialLote.IDLineaMovimiento) = updateData.IDLineaMovimiento 'VinoMaterialRow(_VM.IDLineaMovimiento) = updateData.IDLineaMovimiento
                            VinoMaterialRow(_VM.Precio) = updateData.StockData.PrecioA
                        ElseIf updateData.Estado = EstadoStock.NoActualizado Then
                            If Length(Data.Lote) > 0 Then
                                ApplicationService.GenerateError(updateData.Log & vbNewLine & "Lote relacionado: " & Data.Lote)
                            Else
                                ApplicationService.GenerateError(updateData.Log)
                            End If
                        End If
                    End If
                End If
            Next
            bsnVML.Update(dttVinoMaterialLote)
        End If
    End Sub

    <Task()> Public Shared Sub DescuentoMaterialesLinea(ByVal VinoMaterialRow As DataRow, ByVal services As ServiceProvider)
        If Length(VinoMaterialRow(_VM.IDLineaMovimiento)) = 0 Then
            If Nz(VinoMaterialRow(_VM.Cantidad), 0) <> 0 Then
                Dim NumeroMovimiento As Integer = ProcessServer.ExecuteTask(Of Object, Integer)(AddressOf ProcesoStocks.NuevoNumeroMovimiento, Nothing, services)
                Dim Data = New StockData
                Data.Articulo = VinoMaterialRow(_VM.IDArticulo)
                If Not VinoMaterialRow.IsNull(_VM.IDAlmacen) Then
                    Data.Almacen = VinoMaterialRow(_VM.IDAlmacen)
                Else
                    Dim p As New Parametro
                    Data.Almacen = p.AlmacenPredeterminado
                End If
                Data.Cantidad = Nz(VinoMaterialRow(_VM.Cantidad), 0) + Nz(VinoMaterialRow(_VM.Merma), 0)
                'If Length(VinoMaterialRow(_VM.Lote)) > 0 Then
                '    Data.Lote = VinoMaterialRow(_VM.Lote)
                'End If
                'If Length(VinoMaterialRow(_VM.Ubicacion)) > 0 Then
                '    Data.Ubicacion = VinoMaterialRow(_VM.Ubicacion)
                'End If
                Data.TipoMovimiento = enumTipoMovimiento.tmSalFabrica
                Data.PrecioA = Nz(VinoMaterialRow(_VM.Precio), 0)
                If Length(VinoMaterialRow(_VM.NOperacion)) > 0 Then
                    Data.Documento = VinoMaterialRow(_VM.NOperacion)
                End If
                If Not IsDate(VinoMaterialRow(_VM.Fecha)) OrElse VinoMaterialRow(_VM.Fecha) = cnMinDate Then
                    Data.FechaDocumento = Today
                Else
                    Data.FechaDocumento = New Date(CDate(VinoMaterialRow(_VM.Fecha)).Year, CDate(VinoMaterialRow(_VM.Fecha)).Month, CDate(VinoMaterialRow(_VM.Fecha)).Day)
                End If

                Dim dataSalida As New DataNumeroMovimientoSinc(NumeroMovimiento, Data, False)
                Dim updateData As StockUpdateData = ProcessServer.ExecuteTask(Of DataNumeroMovimientoSinc, StockUpdateData)(AddressOf ProcesoStocks.Salida, dataSalida, services)
                If Not updateData Is Nothing Then
                    If updateData.Estado = EstadoStock.Actualizado Then
                        VinoMaterialRow(_VM.IDLineaMovimiento) = updateData.IDLineaMovimiento
                        VinoMaterialRow(_VM.Precio) = updateData.StockData.PrecioA
                    ElseIf updateData.Estado = EstadoStock.NoActualizado Then
                        If Length(Data.Lote) > 0 Then
                            ApplicationService.GenerateError(updateData.Log & vbNewLine & "Lote relacionado: " & Data.Lote)
                        Else
                            ApplicationService.GenerateError(updateData.Log)
                        End If
                    End If
                End If
            End If
        Else
            If VinoMaterialRow(_VM.Cantidad) <> VinoMaterialRow(_VM.Cantidad, DataRowVersion.Original) _
            OrElse VinoMaterialRow(_VM.Merma) <> VinoMaterialRow(_VM.Merma, DataRowVersion.Original) _
            OrElse VinoMaterialRow(_VM.Precio) <> VinoMaterialRow(_VM.Precio, DataRowVersion.Original) Then
                Dim dblNwQ As Double = Nz(VinoMaterialRow(_VM.Cantidad), 0) + Nz(VinoMaterialRow(_VM.Merma), 0)
                Dim dataCorreccion As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Corregir, VinoMaterialRow(_VM.IDLineaMovimiento), _
                                                                                 dblNwQ, VinoMaterialRow(_VM.Precio), 0, False)
                Dim updateData As StockUpdateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento, StockUpdateData)(AddressOf ProcesoStocks.ActualizarMovimiento, dataCorreccion, services)
                If Not updateData Is Nothing Then
                    If updateData.Estado = EstadoStock.Actualizado Then
                        VinoMaterialRow(_VM.IDLineaMovimiento) = updateData.IDLineaMovimiento
                        VinoMaterialRow(_VM.Precio) = updateData.StockData.PrecioA
                    ElseIf updateData.Estado = EstadoStock.NoActualizado Then
                        ApplicationService.GenerateError(updateData.Log)
                    End If
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub DeshacerDescuentosMateriales(ByVal VinoMaterial As DataTable, ByVal services As ServiceProvider)
        If Not VinoMaterial Is Nothing AndAlso VinoMaterial.Rows.Count > 0 Then

            Dim bsnVML As BusinessHelper = BusinessHelper.CreateBusinessObject("BdgVinoMaterialLote")
            Dim filter As New Filter
            For Each dr As DataRow In VinoMaterial.Select
                'deshacemos sus propios movimientos si los hubiere (caso de materiales sin gestión por lotes)
                If Length(dr(_VM.IDLineaMovimiento)) > 0 Then
                    Dim dataCorreccion As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Eliminar, dr(_VM.IDLineaMovimiento), False)
                    ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento)(AddressOf ProcesoStocks.ActualizarMovimiento, dataCorreccion, services)
                End If

                'obtenemos sus líneas
                filter.Clear()
                filter.Add(New GuidFilterItem("IDVinoMaterial", dr("IDVinoMaterial")))
                Dim dttVinoMaterialLote As DataTable = bsnVML.Filter(filter)
                If Not dttVinoMaterialLote Is Nothing AndAlso dttVinoMaterialLote.Rows.Count > 0 Then
                    For Each drLote As DataRow In dttVinoMaterialLote.Rows
                        'deshacemos sus movimientos, si los hubiere
                        If Length(drLote(_BdgVinoMaterialLote.IDLineaMovimiento)) > 0 Then
                            Dim dataCorreccion As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Eliminar, _
                                                                                             drLote(_BdgVinoMaterialLote.IDLineaMovimiento), False)
                            ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento)(AddressOf ProcesoStocks.ActualizarMovimiento, dataCorreccion, services)
                        End If
                    Next
                End If

            Next
        End If
    End Sub

    <Serializable()> _
    Public Class StActualizarCosteVendVinoAgrup
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

    <Task()> Public Shared Sub ActualizarCosteVendimiaVinoAgrup(ByVal data As StActualizarCosteVendVinoAgrup, ByVal services As ServiceProvider)
        If data.DtMarcados.Rows.Count > 0 Then
            For Each drMarcado As DataRow In data.DtMarcados.Rows
                Dim DataActua As New StActualizarCosteVendVino(drMarcado("IDVinoMaterial"), data.Fecha, data.Coste)
                ProcessServer.ExecuteTask(Of StActualizarCosteVendVino)(AddressOf ActualizarCosteVendimiaVino, DataActua, services)
            Next
        End If
    End Sub

    <Serializable()> _
    Public Class StActualizarCosteVendVino
        Public IDVinoMaterial As Guid
        Public Fecha As Date
        Public Coste As Double

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDVinoMaterial As Guid, ByVal Fecha As Date, ByVal Coste As Double)
            Me.IDVinoMaterial = IDVinoMaterial
            Me.Fecha = Fecha
            Me.Coste = Coste
        End Sub
    End Class

    <Task()> Public Shared Sub ActualizarCosteVendimiaVino(ByVal data As StActualizarCosteVendVino, ByVal services As ServiceProvider)
        Dim ClsBdgVinoMat As New BdgVinoMaterial
        Dim ff As New Filter
        'Actualizar VinoMaterial
        If Length(data.IDVinoMaterial.ToString) > 0 Then
            ff.Clear()
            ff.Add("IDVinoMaterial", FilterOperator.Equal, data.IDVinoMaterial)
            Dim dtMater As DataTable = ClsBdgVinoMat.Filter(ff)
            If Not IsNothing(dtMater) AndAlso dtMater.Rows.Count > 0 Then
                If dtMater.Rows(0)("Fecha") <> data.Fecha Then
                    dtMater.Rows(0)("Fecha") = data.Fecha
                End If
                If dtMater.Rows(0)("Precio") <> data.Coste Then
                    dtMater.Rows(0)("Precio") = data.Coste
                End If
                AdminData.BeginTx()
                ClsBdgVinoMat.Update(dtMater)
            End If
        End If
    End Sub

#End Region

End Class

<Serializable()> _
Public Class _VM
    Public Const IDVinoMaterial As String = "IDVinoMaterial"
    Public Const IDVino As String = "IDVino"
    Public Const IDArticulo As String = "IDArticulo"
    Public Const Fecha As String = "Fecha"
    Public Const NOperacion As String = "NOperacion"
    Public Const Cantidad As String = "Cantidad"
    Public Const Merma As String = "Merma"
    Public Const Precio As String = "Precio"
    Public Const IDLineaMovimiento As String = "IDLineaMovimiento"
    Public Const IDAlmacen As String = "IDAlmacen"
    Public Const IDOperacionMaterial As String = "IDOperacionMaterial"
    Public Const RecalcularMaterial As String = "RecalcularMaterial"
End Class