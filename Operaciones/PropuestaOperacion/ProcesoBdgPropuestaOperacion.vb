Public Class ProcesoBdgPropuestaOperacion

#Region " Agrupar datos Origen "

    <Task()> Public Shared Function AgruparOrigenes(ByVal data As DataPrcPropuestaOperacion, ByVal services As ServiceProvider) As List(Of OperCab)
        Select Case data.Origen
            Case OrigenOperacion.OperacionPlanificada
                Return ProcessServer.ExecuteTask(Of DataPrcPropuestaOperacion, List(Of OperCab))(AddressOf AgruparOrigenesOpPlanificada, data, services)
            Case OrigenOperacion.Depositos
                Return ProcessServer.ExecuteTask(Of DataPrcPropuestaOperacion, List(Of OperCab))(AddressOf AgruparOrigenesDepositos, data, services)
            Case OrigenOperacion.OrdenFabricacion
                Return ProcessServer.ExecuteTask(Of DataPrcPropuestaOperacion, List(Of OperCab))(AddressOf AgruparOrigenesOFs, data, services)
            Case OrigenOperacion.Expediciones
                Return ProcessServer.ExecuteTask(Of DataPrcPropuestaOperacion, List(Of OperCab))(AddressOf AgruparOrigenesExpediciones, data, services)
        End Select
    End Function

    <Task()> Public Shared Function AgruparOrigenesOpPlanificada(ByVal data As DataPrcPropuestaOperacion, ByVal services As ServiceProvider) As List(Of OperCab)
        Dim lstOperCab As New List(Of OperCab)

        For Each IDOrigen As String In data.IDOrigen
            Dim op As New OperCabPlanificadas(data.TipoOperacion, data.Origen, IDOrigen, Nz(data.Fecha, cnMinDate))
            lstOperCab.Add(op)
        Next

        Return lstOperCab
    End Function

    <Task()> Public Shared Function AgruparOrigenesDepositos(ByVal data As DataPrcPropuestaOperacion, ByVal services As ServiceProvider) As List(Of OperCab)
        Dim lstOperCab As New List(Of OperCab)
        If data.Multiple Then
            Return ProcessServer.ExecuteTask(Of DataPrcPropuestaOperacion, List(Of OperCab))(AddressOf AgruparOrigenesDepositosOperacionesMultiples, data, services)
        Else
            For Each datOrigen As DataPropuestaDatosDepositos In data.OrigenesDepositos
                Dim op As New OperCabDepositos(data.TipoOperacion, data.Origen, datOrigen.Depositos, datOrigen.IDTipoOperacion, datOrigen.Fecha)
                lstOperCab.Add(op)
            Next
        End If

        Return lstOperCab
    End Function

    <Task()> Public Shared Function AgruparOrigenesDepositosOperacionesMultiples(ByVal data As DataPrcPropuestaOperacion, ByVal services As ServiceProvider) As List(Of OperCab)
        If Not data.Multiple Then Exit Function

        Dim lstOperCab As New List(Of OperCab)
        For Each datOrigen As DataPropuestaDatosDepositos In data.OrigenesDepositos
            Dim TiposOperacion As EntityInfoCache(Of BdgTipoOperacionInfo) = services.GetService(Of EntityInfoCache(Of BdgTipoOperacionInfo))()
            Dim TipoOpInfo As BdgTipoOperacionInfo = TiposOperacion.GetEntity(datOrigen.IDTipoOperacion)

            Dim lstDepositosEnOrigen As List(Of DatosDeposito) = (From c In datOrigen.Depositos Where c.Destino = False Select c).ToList
            Dim lstDepositosEnDestino As List(Of DatosDeposito) = (From c In datOrigen.Depositos Where c.Destino = True Select c).ToList
            If lstDepositosEnOrigen.Count > 0 AndAlso lstDepositosEnDestino.Count > 0 Then
                ApplicationService.GenerateError("Para generar Operaciones múltiples, deben elegir elementos sólo en el Origen o sólo en el Destino.")
            Else

                If Not lstDepositosEnOrigen Is Nothing AndAlso lstDepositosEnOrigen.Count > 0 Then
                    For Each dptoEnOrigen As DatosDeposito In lstDepositosEnOrigen
                        Dim lstDepositosOperacion As New List(Of DatosDeposito)
                        lstDepositosOperacion.Add(dptoEnOrigen)
                        If lstDepositosOperacion.Count > 0 Then
                            Dim op As New OperCabDepositos(data.TipoOperacion, data.Origen, lstDepositosOperacion, datOrigen.IDTipoOperacion, datOrigen.Fecha)
                            lstOperCab.Add(op)
                        End If
                    Next
                End If

                If Not lstDepositosEnDestino Is Nothing AndAlso lstDepositosEnDestino.Count > 0 Then
                    For Each dptoEnDestino As DatosDeposito In lstDepositosEnDestino
                        Dim lstDepositosOperacion As New List(Of DatosDeposito)
                        lstDepositosOperacion.Add(dptoEnDestino)
                        If lstDepositosOperacion.Count > 0 Then
                            Dim op As New OperCabDepositos(data.TipoOperacion, data.Origen, lstDepositosOperacion, datOrigen.IDTipoOperacion, datOrigen.Fecha)
                            lstOperCab.Add(op)
                        End If
                    Next
                End If

            End If

        Next
        Return lstOperCab
    End Function

    <Task()> Public Shared Function AgruparOrigenesOFs(ByVal data As DataPrcPropuestaOperacion, ByVal services As ServiceProvider) As List(Of OperCab)
        Dim lstOperCab As New List(Of OperCab)

        Dim BEDataEngine As New BE.DataEngine

        Dim dataOF As DataPrcPropuestaOperacionOFs = CType(data, DataPrcPropuestaOperacionOFs)
        For Each IDOrden As Integer In dataOF.OrigenesOFs.Keys
            Dim NOrden As String = String.Empty
            Dim dtOF As DataTable = BEDataEngine.Filter("NegBdgOrdenFabricacion", New NumberFilterItem("IDOrden", IDOrden))
            If dtOF.Rows.Count > 0 Then
                NOrden = dtOF.Rows(0)("NOrden")

                Dim blnArticuloVino As Boolean = False
                blnArticuloVino = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf BdgVino.EsArticuloVino, dtOF.Rows(0)("IDArticulo"), services)
                If blnArticuloVino Then
                    Dim IDTipoOperacion As String
                    If Not dataOF Is Nothing AndAlso Not dataOF.TiposOperacionesOrdenes Is Nothing AndAlso dataOF.TiposOperacionesOrdenes.Count > 0 Then
                        IDTipoOperacion = dataOF.TiposOperacionesOrdenes(IDOrden) & String.Empty
                    End If

                    If Length(IDTipoOperacion) = 0 Then
                        If Length(dtOF.Rows(0)("IDTipoOperacion")) > 0 Then
                            IDTipoOperacion = dtOF.Rows(0)("IDTipoOperacion")
                        ElseIf Length(dataOF.IDTipoOperacion) > 0 Then
                            IDTipoOperacion = dataOF.IDTipoOperacion
                        End If
                    End If

                    If Length(IDTipoOperacion) = 0 Then
                        Dim log As LogProcess = services.GetService(Of LogProcess)()
                        Dim Msg As String = Engine.ParseFormatString(AdminData.GetMessageText("No hay un Tipo de Operación Predeterminado para la Orden {0}."), Quoted(NOrden))
                        ReDim Preserve log.Errors(log.Errors.Length)
                        log.Errors(log.Errors.Length - 1) = New ClassErrors(NOrden, Msg)
                    Else
                        Dim op As New OperCabOFs(dataOF.TipoOperacion, dataOF.Origen, IDOrden, NOrden, dataOF.OrigenesOFs(IDOrden), IDTipoOperacion, dataOF.Fecha)
                        lstOperCab.Add(op)
                    End If
                Else
                    Dim log As LogProcess = services.GetService(Of LogProcess)()
                    Dim Msg As String = Engine.ParseFormatString(AdminData.GetMessageText("El Artículo {0} no es un Artículo de Bodega"), Quoted(dtOF.Rows(0)("IDArticulo")))
                    ReDim Preserve log.Errors(log.Errors.Length)
                    log.Errors(log.Errors.Length - 1) = New ClassErrors(NOrden, Msg)
                End If
            End If
        Next

        Return lstOperCab
    End Function

    <Task()> Public Shared Function AgruparOrigenesExpediciones(ByVal data As DataPrcPropuestaOperacion, ByVal services As ServiceProvider) As List(Of OperCab)

        Dim lstOperCab As New List(Of OperCab)

        Dim dataExp As DataPrcPropuestaOperacionExpediciones = CType(data, DataPrcPropuestaOperacionExpediciones)
        For Each datArtComp As DataArtCompatiblesExp In dataExp.OrigenesExp
            Dim op As New OperCabExp(dataExp.TipoOperacion, dataExp.Origen, datArtComp.IDTipoOperacion, datArtComp.Fecha, datArtComp.IDLineaPedido, datArtComp.IDDeposito, datArtComp.dtArtCompatibles)
            lstOperCab.Add(op)
        Next

        Return lstOperCab
    End Function

#End Region

#Region " Crear documento "

    <Task()> Public Shared Function CrearDocumentoBdgOperacion(ByVal oper As OperCab, ByVal services As ServiceProvider) As DocumentoBdgOperacion
        Dim doc As DocumentoBdgOperacion
        Select Case oper.TipoOperacion
            Case enumBdgOrigenOperacion.Planificada
                doc = New DocumentoBdgOperacionPlan(oper, services)
            Case enumBdgOrigenOperacion.Real
                doc = New DocumentoBdgOperacionReal(oper, services)
        End Select
        oper.Doc = doc
        ProcessServer.ExecuteTask(Of DocumentoBdgOperacion)(AddressOf AddCamposComplementarios, doc, services)
        Return doc
    End Function
    <Task()> Public Shared Sub AddCamposComplementarios(ByVal doc As DocumentoBdgOperacion, ByVal services As ServiceProvider)
        If Not doc.dtOperacionVino.Columns.Contains("DescArticulo") Then doc.dtOperacionVino.Columns.Add("DescArticulo", GetType(String))
        If Not doc.dtOperacionVino.Columns.Contains("IDUDMedida") Then doc.dtOperacionVino.Columns.Add("IDUDMedida", GetType(String))
        If Not doc.dtOperacionVino.Columns.Contains("Fecha") Then doc.dtOperacionVino.Columns.Add("Fecha", GetType(Date))


        If Not doc.dtOperacionMaterial.Columns.Contains("DescArticulo") Then doc.dtOperacionMaterial.Columns.Add("DescArticulo", GetType(String))
        If Not doc.dtOperacionMaterial.Columns.Contains("GestionStockPorLotes") Then doc.dtOperacionMaterial.Columns.Add("GestionStockPorLotes", GetType(String))
        If Not doc.dtOperacionMaterial.Columns.Contains("NumLotes") Then doc.dtOperacionMaterial.Columns.Add("NumLotes", GetType(Integer))
        If Not doc.dtOperacionMaterial.Columns.Contains("Lote") Then doc.dtOperacionMaterial.Columns.Add("Lote", GetType(String))
        If Not doc.dtOperacionMaterial.Columns.Contains("Ubicacion") Then doc.dtOperacionMaterial.Columns.Add("Ubicacion", GetType(String))

        If Not doc.dtOperacionMaterialLote.Columns.Contains("IDArticulo") Then doc.dtOperacionMaterialLote.Columns.Add("IDArticulo", GetType(String))
        If Not doc.dtOperacionMaterialLote.Columns.Contains("IDAlmacen") Then doc.dtOperacionMaterialLote.Columns.Add("IDAlmacen", GetType(String))

        If Not doc.dtOperacionMOD.Columns.Contains("DescOperario") Then doc.dtOperacionMOD.Columns.Add("DescOperario", GetType(String))

        If Not doc.dtOperacionCentro.Columns.Contains("DescCentro") Then doc.dtOperacionCentro.Columns.Add("DescCentro", GetType(String))


        If Not doc.dtOperacionVinoMaterial.Columns.Contains("DescArticulo") Then doc.dtOperacionVinoMaterial.Columns.Add("DescArticulo", GetType(String))
        If Not doc.dtOperacionVinoMaterial.Columns.Contains("GestionStockPorLotes") Then doc.dtOperacionVinoMaterial.Columns.Add("GestionStockPorLotes", GetType(String))
        If Not doc.dtOperacionVinoMaterial.Columns.Contains("NumLotes") Then doc.dtOperacionVinoMaterial.Columns.Add("NumLotes", GetType(Integer))
        If Not doc.dtOperacionVinoMaterial.Columns.Contains("Lote") Then doc.dtOperacionVinoMaterial.Columns.Add("Lote", GetType(String))
        If Not doc.dtOperacionVinoMaterial.Columns.Contains("Ubicacion") Then doc.dtOperacionVinoMaterial.Columns.Add("Ubicacion", GetType(String))

        If TypeOf doc Is DocumentoBdgOperacionReal Then
            If Not doc.dtOperacionVinoMaterialLote.Columns.Contains("IDArticulo") Then doc.dtOperacionVinoMaterialLote.Columns.Add("IDArticulo", GetType(String))
            If Not doc.dtOperacionVinoMaterialLote.Columns.Contains("IDAlmacen") Then doc.dtOperacionVinoMaterialLote.Columns.Add("IDAlmacen", GetType(String))
        End If

        If Not doc.dtOperacionVinoMOD.Columns.Contains("DescOperario") Then doc.dtOperacionVinoMOD.Columns.Add("DescOperario", GetType(String))

        If Not doc.dtOperacionVinoCentro.Columns.Contains("DescCentro") Then doc.dtOperacionVinoCentro.Columns.Add("DescCentro", GetType(String))

        If TypeOf doc Is DocumentoBdgOperacionReal Then
            If Not CType(doc, DocumentoBdgOperacionReal).dtOperacionVinoAnalisisVariable.Columns.Contains("Abreviatura") Then
                CType(doc, DocumentoBdgOperacionReal).dtOperacionVinoAnalisisVariable.Columns.Add("Abreviatura", GetType(String))
            End If


            If Not CType(doc, DocumentoBdgOperacionReal).dtOperacionVinoAnalisisVariable.Columns.Contains("IDVino") Then
                CType(doc, DocumentoBdgOperacionReal).dtOperacionVinoAnalisisVariable.Columns.Add("IDVino", GetType(Guid))
            End If
        End If

    End Sub

#End Region

#Region " Retornar elementos creados "


    <Task()> Public Shared Sub AñadirAResultado(ByVal Doc As DocumentoBdgOperacion, ByVal services As ServiceProvider)

        Dim Propuestas As List(Of DataPrcPropuestaOperacionResult) = services.GetService(Of List(Of DataPrcPropuestaOperacionResult))()

        Dim Propuesta As DataPrcPropuestaOperacionResult
        If TypeOf Doc Is DocumentoBdgOperacionPlan Then
            Propuesta = New DataPrcPropuestaOperacionResult(enumBdgOrigenOperacion.Planificada)
        ElseIf TypeOf Doc Is DocumentoBdgOperacionReal Then
            Propuesta = New DataPrcPropuestaOperacionResult(enumBdgOrigenOperacion.Real)
        End If

        Propuesta.OperacionCabecera = Doc.HeaderRow.Table

        If Propuesta.OperacionVinoOrigen Is Nothing Then Propuesta.OperacionVinoOrigen = Doc.dtOperacionVino.Clone
        If Propuesta.OperacionVinoDestino Is Nothing Then Propuesta.OperacionVinoDestino = Doc.dtOperacionVino.Clone

        For Each dr As DataRow In Doc.GetOperacionVinoOrigen
            Propuesta.OperacionVinoOrigen.Rows.Add(dr.ItemArray)
        Next

        For Each dr As DataRow In Doc.GetOperacionVinoDestino
            Propuesta.OperacionVinoDestino.Rows.Add(dr.ItemArray)
        Next

        Propuesta.OperacionMaterial = Doc.dtOperacionMaterial
        Propuesta.OperacionMOD = Doc.dtOperacionMOD
        Propuesta.OperacionCentro = Doc.dtOperacionCentro
        Propuesta.OperacionVarios = Doc.dtOperacionVarios

        Propuesta.OperacionVinoMaterial = Doc.dtOperacionVinoMaterial
        Propuesta.OperacionVinoMOD = Doc.dtOperacionVinoMOD
        Propuesta.OperacionVinoCentro = Doc.dtOperacionVinoCentro
        Propuesta.OperacionVinoVarios = Doc.dtOperacionVinoVarios
        If TypeOf Doc Is DocumentoBdgOperacionReal Then
            Propuesta.OperacionVinoAnalisis = CType(Doc, DocumentoBdgOperacionReal).dtOperacionVinoAnalisis
            Propuesta.OperacionVinoAnalisisVariable = CType(Doc, DocumentoBdgOperacionReal).dtOperacionVinoAnalisisVariable
        End If

        Select Case Doc.Cabecera.Origen
            Case OrigenOperacion.OperacionPlanificada
                Propuesta.OperacionMaterialLotes = Doc.dtOperacionMaterialLote
                Propuesta.OperacionVinoMaterialLotes = Doc.dtOperacionVinoMaterialLote
        End Select

        Propuestas.Add(Propuesta)

        Dim ProcInfo As ProcessInfoOperacion = services.GetService(Of ProcessInfoOperacion)()
        If ProcInfo.GuardarPropuesta Then
            ProcessServer.ExecuteTask(Of DocumentoBdgOperacion)(AddressOf AgregarOperacionAResultado, Doc, services)
        End If
    End Sub

    <Task()> Public Shared Sub AgregarOperacionAResultado(ByVal doc As DocumentoBdgOperacion, ByVal services As ServiceProvider)
        Dim rslt As LogProcess = services.GetService(Of LogProcess)()

        ReDim Preserve rslt.CreatedElements(UBound(rslt.CreatedElements) + 1)
        rslt.CreatedElements(UBound(rslt.CreatedElements)) = New CreateElement
        rslt.CreatedElements(UBound(rslt.CreatedElements)).NElement = doc.HeaderRow(doc.FieldNOperacion)
    End Sub

#End Region


#Region " Propuesta Cabecera de la Operación "

    <Serializable()> _
    Public Class DataImputacionesGlobales
        Public dtOperacionMaterial As DataTable
        Public dtOperacionCentro As DataTable
        Public dtOperacionMOD As DataTable
        Public dtOperacionVarios As DataTable
    End Class

    <Serializable()> _
    Public Class DataOrdenFabricacion
        Public ProductoTerminado As DataTable
        Public MaterialesVino As DataTable
        Public Materiales As DataTable
    End Class
    <Task()> Public Shared Sub PropuestaCabeceraOperacion(ByVal doc As DocumentoBdgOperacion, ByVal services As ServiceProvider)
        If doc.Cabecera Is Nothing Then Exit Sub


        Dim datOrigen As DataGetOrigenPropuesta = ProcessServer.ExecuteTask(Of OperCab, DataGetOrigenPropuesta)(AddressOf GetOrigenPropuestaCabecera, doc.Cabecera, services)
        If Not datOrigen Is Nothing AndAlso Not datOrigen.dtOrigen Is Nothing AndAlso datOrigen.dtOrigen.Rows.Count > 0 Then

            Dim BH As BusinessHelper = BusinessHelper.CreateBusinessObject(doc.EntidadCabecera)
            Dim current As New BusinessData(doc.HeaderRow)

            Dim datFecha As New BdgGeneral.DataGetFechaPropuestaOperacion(doc.EntidadCabecera)
            datFecha.FechaOrigen = Nz(datOrigen.dtOrigen.Rows(0)("Fecha"), cnMinDate)
            Dim FechaOperacion As Date = ProcessServer.ExecuteTask(Of BdgGeneral.DataGetFechaPropuestaOperacion, Date)(AddressOf BdgGeneral.GetFechaPropuestaOperacion, datFecha, services)
            current = BH.ApplyBusinessRule("Fecha", FechaOperacion, current, Nothing)
            current = BH.ApplyBusinessRule("IDTipoOperacion", datOrigen.dtOrigen.Rows(0)("IDTipoOperacion"), current, Nothing)

            Select Case doc.Cabecera.Origen
                Case OrigenOperacion.OperacionPlanificada
                    doc.HeaderRow("ImputacionRealCentro") = datOrigen.dtOrigen.Rows(0)("ImputacionGlobalCentro")
                    doc.HeaderRow("ImputacionRealMaterial") = datOrigen.dtOrigen.Rows(0)("ImputacionGlobalMat")
                    doc.HeaderRow("ImputacionRealMod") = datOrigen.dtOrigen.Rows(0)("ImputacionGlobalMod")
                    doc.HeaderRow("ImputacionRealVarios") = datOrigen.dtOrigen.Rows(0)("ImputacionGlobalVarios")
                Case Else
                    '//Almacenamos en el services
                    Dim datImputGlobales As DataImputacionesGlobales = services.GetService(Of DataImputacionesGlobales)()
                    If current.ContainsKey("DTImputacionMaterialGlobal") Then
                        datImputGlobales.dtOperacionMaterial = current("DTImputacionMaterialGlobal")
                    End If
                    If current.ContainsKey("DTImputacionCentroGlobal") Then
                        datImputGlobales.dtOperacionCentro = current("DTImputacionCentroGlobal")
                    End If
                    If current.ContainsKey("DTImputacionMODGlobal") Then
                        datImputGlobales.dtOperacionMOD = current("DTImputacionMODGlobal")
                    End If
                    If current.ContainsKey("DTImputacionVariosGlobal") Then
                        datImputGlobales.dtOperacionVarios = current("DTImputacionVariosGlobal")
                    End If
            End Select


            If Length(datOrigen.dtOrigen.Rows(0)("IDAnalisis")) > 0 Then
                current = BH.ApplyBusinessRule("IDAnalisis", datOrigen.dtOrigen.Rows(0)("IDAnalisis"), current, Nothing)
            End If
            For Each col As DataColumn In doc.HeaderRow.Table.Columns
                doc.HeaderRow(col.ColumnName) = current(col.ColumnName)
            Next

            doc.HeaderRow("Texto") = datOrigen.dtOrigen.Rows(0)("Texto")

            Select Case doc.Cabecera.Origen
                Case OrigenOperacion.OperacionPlanificada
                    doc.HeaderRow("NOperacionPlan") = datOrigen.dtOrigen.Rows(0)("NOperacionPlan")
                Case OrigenOperacion.OrdenFabricacion
                    doc.HeaderRow("IDOrden") = datOrigen.dtOrigen.Rows(0)("IDOrden")
                    doc.HeaderRow("CantidadOrden") = datOrigen.dtOrigen.Rows(0)("CantidadOrden")
            End Select

            If doc.HeaderRow.Table.Columns.Contains("IDOrden") AndAlso _
               doc.HeaderRow.Table.Columns.Contains("CantidadOrden") Then
                If Length(doc.HeaderRow("IDOrden")) > 0 AndAlso Nz(doc.HeaderRow("CantidadOrden")) > 0 Then
                    Dim datOF As DataOrdenFabricacion = services.GetService(Of DataOrdenFabricacion)()
                    Dim dataDatosOf As New BdgGeneral.DataDatosOfParaOperacion(doc.HeaderRow("IDOrden"), doc.HeaderRow("CantidadOrden"))
                    dataDatosOf = ProcessServer.ExecuteTask(Of BdgGeneral.DataDatosOfParaOperacion, BdgGeneral.DataDatosOfParaOperacion)(AddressOf BdgGeneral.DatosOfParaOperacion, dataDatosOf, services)
                    datOF.ProductoTerminado = dataDatosOf.dtDestino
                    datOF.MaterialesVino = dataDatosOf.dtOrigen
                    datOF.Materiales = dataDatosOf.dtMateriales
                End If
            End If
        End If

    End Sub

    <Serializable()> _
    Public Class DataGetOrigenPropuesta
        Public IDOrigen As String
        Public IDLineaOrigen As Guid
        Public dtOrigen As DataTable

        Public Sub New(ByVal IDOrigen As String, ByVal dtOrigen As DataTable)
            Me.IDOrigen = IDOrigen
            Me.dtOrigen = dtOrigen
        End Sub
    End Class
    <Task()> Public Shared Function GetOrigenPropuestaCabecera(ByVal oper As OperCab, ByVal services As ServiceProvider) As DataGetOrigenPropuesta
        Select Case oper.Origen
            Case OrigenOperacion.OperacionPlanificada
                Return ProcessServer.ExecuteTask(Of OperCab, DataGetOrigenPropuesta)(AddressOf GetOrigenPropuestaCabeceraOpPlanificada, oper, services)
            Case OrigenOperacion.Depositos
                Return ProcessServer.ExecuteTask(Of OperCab, DataGetOrigenPropuesta)(AddressOf GetOrigenPropuestaCabeceraDepositos, oper, services)
            Case OrigenOperacion.OrdenFabricacion
                Return ProcessServer.ExecuteTask(Of OperCab, DataGetOrigenPropuesta)(AddressOf GetOrigenPropuestaCabeceraOFs, oper, services)
            Case OrigenOperacion.Expediciones
                Return ProcessServer.ExecuteTask(Of OperCab, DataGetOrigenPropuesta)(AddressOf GetOrigenPropuestaCabeceraExpedicion, oper, services)
        End Select
    End Function

    <Task()> Public Shared Function CreateEstructuraOrigenCabecera(ByVal data As Object, ByVal services As ServiceProvider) As DataTable
        Dim dt As New DataTable
        dt.Columns.Add("IDTipoOperacion", GetType(String))
        dt.Columns.Add("Fecha", GetType(Date))
        dt.Columns.Add("IDAnalisis", GetType(String))
        dt.Columns.Add("ImputacionGlobalMat", GetType(Boolean))
        dt.Columns.Add("ImputacionGlobalMod", GetType(Boolean))
        dt.Columns.Add("ImputacionGlobalCentro", GetType(Boolean))
        dt.Columns.Add("ImputacionGlobalVarios", GetType(Boolean))
        dt.Columns.Add("Texto", GetType(String))
        dt.Columns.Add("NOperacionPlan", GetType(String))
        dt.Columns.Add("IDOrden", GetType(Integer))
        dt.Columns.Add("CantidadOrden", GetType(Double))
        Return dt
    End Function
    <Task()> Public Shared Function GetOrigenPropuestaCabeceraOpPlanificada(ByVal oper As OperCab, ByVal services As ServiceProvider) As DataGetOrigenPropuesta
        Dim NOperacionPlan As String = CType(oper, OperCabPlanificadas).IDOrigen
        Dim dtOrigenDatosCabecera As DataTable = New BdgOperacionPlan().SelOnPrimaryKey(NOperacionPlan)
        If dtOrigenDatosCabecera.Rows.Count > 0 Then
            Select Case dtOrigenDatosCabecera.Rows(0)("Estado")
                Case BdgEstadoOperacionPlan.Anulado
                    ApplicationService.GenerateError("La Operación planificada está anulada. No se puede generar una Operación de Bodega.")
                Case BdgEstadoOperacionPlan.Confirmado
                    ApplicationService.GenerateError("La Operación planificada está Confirmada. No se puede generar una Operación de Bodega.")
                Case BdgEstadoOperacionPlan.Planificado
                    Dim dtOrigen As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf CreateEstructuraOrigenCabecera, Nothing, services)
                    Dim drNew As DataRow = dtOrigen.NewRow
                    drNew("IDTipoOperacion") = dtOrigenDatosCabecera.Rows(0)("IDTipoOperacion")
                    If oper.Fecha <> cnMinDate Then
                        drNew("Fecha") = oper.Fecha
                    Else
                        drNew("Fecha") = dtOrigenDatosCabecera.Rows(0)("Fecha")
                    End If
                    drNew("IDAnalisis") = dtOrigenDatosCabecera.Rows(0)("IDAnalisis")
                    drNew("ImputacionGlobalMat") = dtOrigenDatosCabecera.Rows(0)("ImputacionGlobalMat")
                    drNew("ImputacionGlobalMod") = dtOrigenDatosCabecera.Rows(0)("ImputacionGlobalMod")
                    drNew("ImputacionGlobalCentro") = dtOrigenDatosCabecera.Rows(0)("ImputacionGlobalCentro")
                    drNew("ImputacionGlobalVarios") = dtOrigenDatosCabecera.Rows(0)("ImputacionGlobalVarios")
                    drNew("Texto") = dtOrigenDatosCabecera.Rows(0)("Texto")
                    drNew("NOperacionPlan") = dtOrigenDatosCabecera.Rows(0)("NOperacionPlan")
                    dtOrigen.Rows.Add(drNew)
                    Dim datOrigen As New DataGetOrigenPropuesta(NOperacionPlan, dtOrigen)
                    Return datOrigen
            End Select
        End If
    End Function
    <Task()> Public Shared Function GetOrigenPropuestaCabeceraDepositos(ByVal oper As OperCab, ByVal services As ServiceProvider) As DataGetOrigenPropuesta
        Dim doc As DocumentoBdgOperacion = oper.Doc
        Dim dtOrigen As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf CreateEstructuraOrigenCabecera, Nothing, services)

        Dim drNew As DataRow = dtOrigen.NewRow
        drNew("IDTipoOperacion") = oper.IDTipoOperacion
        drNew("Fecha") = oper.Fecha
        dtOrigen.Rows.Add(drNew)

        Dim datOrigen As New DataGetOrigenPropuesta(Nothing, dtOrigen)
        Return datOrigen
    End Function
    <Task()> Public Shared Function GetOrigenPropuestaCabeceraOFs(ByVal oper As OperCab, ByVal services As ServiceProvider) As DataGetOrigenPropuesta
        Dim dtOrigen As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf CreateEstructuraOrigenCabecera, Nothing, services)

        Dim drNew As DataRow = dtOrigen.NewRow
        drNew("IDTipoOperacion") = oper.IDTipoOperacion
        drNew("Fecha") = oper.Fecha
        drNew("IDOrden") = CType(oper, OperCabOFs).IDOrden
        drNew("CantidadOrden") = CType(oper, OperCabOFs).CantidadOrden
        dtOrigen.Rows.Add(drNew)

        Dim datOrigen As New DataGetOrigenPropuesta(CType(oper, OperCabOFs).NOrden, dtOrigen)
        Return datOrigen
    End Function

    <Task()> Public Shared Function GetOrigenPropuestaCabeceraExpedicion(ByVal oper As OperCab, ByVal services As ServiceProvider) As DataGetOrigenPropuesta
        Dim dtOrigen As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf CreateEstructuraOrigenCabecera, Nothing, services)

        Dim drNew As DataRow = dtOrigen.NewRow
        drNew("IDTipoOperacion") = oper.IDTipoOperacion
        drNew("Fecha") = oper.Fecha
        dtOrigen.Rows.Add(drNew)

        Dim datOrigen As New DataGetOrigenPropuesta(Nothing, dtOrigen)
        Return datOrigen
    End Function

#End Region

#Region " Propuesta Operacion Vino (Origen y Destino) "

    <Task()> Public Shared Sub PropuestaOrigenOperacion(ByVal doc As DocumentoBdgOperacion, ByVal services As ServiceProvider)
        If Not doc.HeaderRow Is Nothing Then

            If doc.Cabecera Is Nothing Then Exit Sub

            Dim datOrigen As DataGetOrigenPropuesta = ProcessServer.ExecuteTask(Of OperCab, DataGetOrigenPropuesta)(AddressOf GetOrigenOperacionVinoOrigen, doc.Cabecera, services)
            If Not datOrigen Is Nothing AndAlso Not datOrigen.dtOrigen Is Nothing AndAlso datOrigen.dtOrigen.Rows.Count > 0 Then
                For Each drOrigen As DataRow In datOrigen.dtOrigen.Rows
                    Dim datNewLine As New DataNuevaLineaOperacionVino(0, doc, drOrigen)
                    ProcessServer.ExecuteTask(Of DataNuevaLineaOperacionVino)(AddressOf NuevaLineaOperacionVino, datNewLine, services)
                Next
            End If
        End If
    End Sub
    <Task()> Public Shared Function GetOrigenOperacionVinoOrigen(ByVal oper As OperCab, ByVal services As ServiceProvider) As DataGetOrigenPropuesta
        Select Case oper.Origen
            Case OrigenOperacion.OperacionPlanificada
                Return ProcessServer.ExecuteTask(Of OperCab, DataGetOrigenPropuesta)(AddressOf GetOrigenOperacionVinoOrigenOpPlanificada, oper, services)
            Case OrigenOperacion.Depositos
                Return ProcessServer.ExecuteTask(Of OperCab, DataGetOrigenPropuesta)(AddressOf GetOrigenOperacionVinoOrigenDepositos, oper, services)
            Case OrigenOperacion.OrdenFabricacion
                Return ProcessServer.ExecuteTask(Of OperCab, DataGetOrigenPropuesta)(AddressOf GetOrigenOperacionVinoOrigenOFs, oper, services)
            Case OrigenOperacion.Expediciones
                Return ProcessServer.ExecuteTask(Of OperCab, DataGetOrigenPropuesta)(AddressOf GetOrigenOperacionVinoOrigenExpedicion, oper, services)
        End Select
    End Function

    <Task()> Public Shared Function GetOrigenOperacionVinoOrigenOpPlanificada(ByVal oper As OperCab, ByVal services As ServiceProvider) As DataGetOrigenPropuesta
        Dim f As New Filter
        f.Add(New FilterItem("NOperacionPlan", CType(oper, OperCabPlanificadas).IDOrigen))
        f.Add(New FilterItem("Destino", 0))

        Dim dtOperacionVinoOrigen As DataTable = New BE.DataEngine().Filter("frmBdgOperacionVinoOrigenPlan", f)
        Dim datOrigen As New DataGetOrigenPropuesta(CType(oper, OperCabPlanificadas).IDOrigen, dtOperacionVinoOrigen)
        Return datOrigen
    End Function
    <Task()> Public Shared Function GetOrigenOperacionVinoOrigenDepositos(ByVal oper As OperCab, ByVal services As ServiceProvider) As DataGetOrigenPropuesta

        Dim ProcInfo As ProcessInfoOperacion = services.GetService(Of ProcessInfoOperacion)()

        Dim dtOrigen As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf GetEstructuraOrigenDestinoOperacionVino, Nothing, services)
        Dim DepositosOrigen As List(Of DatosDeposito) = (From c In CType(oper, OperCabDepositos).DepositosOp Where c.Destino = False Select c).ToList
        For Each datDpstoOrigen As DatosDeposito In DepositosOrigen
            Dim dr As DataRow = dtOrigen.NewRow
            dr("IDDeposito") = datDpstoOrigen.IDDeposito

            If ProcInfo.MultiplesOperaciones Then
                If Not datDpstoOrigen.dtRegistro Is Nothing AndAlso datDpstoOrigen.dtRegistro.Rows.Count > 0 Then
                    For Each col As DataColumn In dtOrigen.Columns
                        If datDpstoOrigen.dtRegistro.Columns.Contains(col.ColumnName) Then
                            dr(col.ColumnName) = datDpstoOrigen.dtRegistro.Rows(0)(col.ColumnName)
                        End If
                    Next
                End If
            Else
                dr("IDVino") = datDpstoOrigen.IDVino
                Dim dtVino As DataTable = New BdgVino().SelOnPrimaryKey(datDpstoOrigen.IDVino)
                If Not dtVino Is Nothing AndAlso dtVino.Rows.Count > 0 Then
                    dr("IDArticulo") = dtVino.Rows(0)("IDArticulo")
                    dr("IDAlmacen") = dtVino.Rows(0)("IDAlmacen")
                    dr("Lote") = dtVino.Rows(0)("Lote")
                End If

                Dim f As New Filter
                f.Add(New StringFilterItem("IDDeposito", datDpstoOrigen.IDDeposito))
                If (Not datDpstoOrigen.IDVino.Equals(Guid.Empty)) Then
                    f.Add(New GuidFilterItem("IDVino", datDpstoOrigen.IDVino))
                End If
                Dim dtArticulo As DataTable = AdminData.GetData("NegBdgOperacionPlan", f, , "IDArticulo, Lote")
                If dtArticulo.Rows.Count > 0 Then
                    dr("IDUDMedida") = dtArticulo.Rows(0)("IDUDMedida")
                    dr("Cantidad") = dtArticulo.Rows(0)("Cantidad")
                    dr("Merma") = 0
                    dr("IDEstadoVino") = dtArticulo.Rows(0)("IDEstadoVino")
                    dr("TipoDeposito") = dtArticulo.Rows(0)("TipoDeposito")
                    dr("Litros") = dtArticulo.Rows(0)("Ocupacion")
                    dr("QDeposito") = dtArticulo.Rows(0)("Ocupacion")
                    dr("IDTipoMermaVino") = System.DBNull.Value
                    dr("Capacidad") = dtArticulo.Rows(0)("Capacidad")
                    dr("Ocupacion") = dtArticulo.Rows(0)("Ocupacion")
                End If

            End If
            dtOrigen.Rows.Add(dr)
        Next


        Dim datOrigen As New DataGetOrigenPropuesta(Nothing, dtOrigen)
        Return datOrigen
    End Function
    <Task()> Public Shared Function GetOrigenOperacionVinoOrigenOFs(ByVal oper As OperCab, ByVal services As ServiceProvider) As DataGetOrigenPropuesta
        Dim dtOrigen As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf GetEstructuraOrigenDestinoOperacionVino, Nothing, services)
        Dim datOF As DataOrdenFabricacion = services.GetService(Of DataOrdenFabricacion)()
        If Not datOF.MaterialesVino Is Nothing AndAlso datOF.MaterialesVino.Rows.Count > 0 Then
            For Each drMaterial As DataRow In datOF.MaterialesVino.Rows
                Dim drNew As DataRow = dtOrigen.NewRow
                For Each col As DataColumn In dtOrigen.Columns
                    If Not IsDBNull(drMaterial("IDVino")) Then
                        If datOF.Materiales.Columns.Contains(col.ColumnName) Then
                            drNew(col.ColumnName) = drMaterial(col.ColumnName)
                        End If

                        drNew("IDDeposito") = drMaterial("Ubicacion")
                        drNew("IDVino") = drMaterial("IDVino")
                        drNew("Cantidad") = drMaterial("Cantidad")

                    End If
                Next
                If Not IsDBNull(drNew("IDVino")) Then dtOrigen.Rows.Add(drNew)
            Next
        End If

        Dim datOrigen As New DataGetOrigenPropuesta(CType(oper, OperCabOFs).NOrden, dtOrigen)
        Return datOrigen
    End Function
    <Task()> Public Shared Function GetOrigenOperacionVinoOrigenExpedicion(ByVal oper As OperCab, ByVal services As ServiceProvider) As DataGetOrigenPropuesta
        Dim dtOrigen As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf GetEstructuraOrigenDestinoOperacionVino, Nothing, services)
        Dim Cab As OperCabExp = CType(oper, OperCabExp)
        If Not Cab.dtArtCompatibles Is Nothing Then
            For Each dr As DataRow In Cab.dtArtCompatibles.Rows
                Dim drNew As DataRow = dtOrigen.NewRow
                drNew("IDDeposito") = dr("Ubicacion")
                drNew("IDArticulo") = dr("IDArticulo")
                drNew("IDAlmacen") = dr("IDAlmacen")
                drNew("Lote") = dr("Lote")
                drNew("IDVino") = dr("IDVino")
                drNew("Cantidad") = dr("Cantidad")
                dtOrigen.Rows.Add(drNew)
            Next
        End If

        Dim datOrigen As New DataGetOrigenPropuesta(Nothing, dtOrigen)
        Return datOrigen
    End Function



    <Task()> Public Shared Sub PropuestaDestinoOperacion(ByVal doc As DocumentoBdgOperacion, ByVal services As ServiceProvider)
        If Not doc.HeaderRow Is Nothing Then

            If doc.Cabecera Is Nothing Then Exit Sub

            Dim datOrigen As DataGetOrigenPropuesta = ProcessServer.ExecuteTask(Of OperCab, DataGetOrigenPropuesta)(AddressOf GetOrigenOperacionVinoDestino, doc.Cabecera, services)
            If Not datOrigen Is Nothing AndAlso Not datOrigen.dtOrigen Is Nothing AndAlso datOrigen.dtOrigen.Rows.Count > 0 Then
                For Each drOrigen As DataRow In datOrigen.dtOrigen.Rows
                    Dim datNewLine As New DataNuevaLineaOperacionVino(1, doc, drOrigen)
                    ProcessServer.ExecuteTask(Of DataNuevaLineaOperacionVino)(AddressOf NuevaLineaOperacionVino, datNewLine, services)
                Next
            End If
        End If
    End Sub
    <Task()> Public Shared Function GetOrigenOperacionVinoDestino(ByVal oper As OperCab, ByVal services As ServiceProvider) As DataGetOrigenPropuesta
        Select Case oper.Origen
            Case OrigenOperacion.OperacionPlanificada
                Return ProcessServer.ExecuteTask(Of OperCab, DataGetOrigenPropuesta)(AddressOf GetOrigenOperacionVinoDestinoOpPlanificada, oper, services)
            Case OrigenOperacion.Depositos
                Return ProcessServer.ExecuteTask(Of OperCab, DataGetOrigenPropuesta)(AddressOf GetOrigenOperacionVinoDestinoDepositos, oper, services)
            Case OrigenOperacion.OrdenFabricacion
                Return ProcessServer.ExecuteTask(Of OperCab, DataGetOrigenPropuesta)(AddressOf GetOrigenOperacionVinoDestinoOFs, oper, services)
            Case OrigenOperacion.Expediciones
                Return ProcessServer.ExecuteTask(Of OperCab, DataGetOrigenPropuesta)(AddressOf GetOrigenOperacionVinoDestinoExpedicion, oper, services)
        End Select
    End Function

    <Task()> Public Shared Function GetOrigenOperacionVinoDestinoOpPlanificada(ByVal oper As OperCab, ByVal services As ServiceProvider) As DataGetOrigenPropuesta
        Dim f As New Filter
        f.Add(New FilterItem("NOperacionPlan", CType(oper, OperCabPlanificadas).IDOrigen))
        f.Add(New FilterItem("Destino", 1))

        Dim dtOperacionVinoDestino As DataTable = New BE.DataEngine().Filter("frmBdgOperacionVinoDestinoPlan", f)
        Dim datOrigen As New DataGetOrigenPropuesta(CType(oper, OperCabPlanificadas).IDOrigen, dtOperacionVinoDestino)
        Return datOrigen
    End Function
    <Task()> Public Shared Function GetOrigenOperacionVinoDestinoDepositos(ByVal oper As OperCab, ByVal services As ServiceProvider) As DataGetOrigenPropuesta
        Dim ProcInfo As ProcessInfoOperacion = services.GetService(Of ProcessInfoOperacion)()

        Dim dtOrigen As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf GetEstructuraOrigenDestinoOperacionVino, Nothing, services)
        Dim DepositosDestino As List(Of DatosDeposito) = (From c In CType(oper, OperCabDepositos).DepositosOp Where c.Destino = True Select c).ToList
        For Each datDpstoDestino As DatosDeposito In DepositosDestino
            Dim dr As DataRow = dtOrigen.NewRow
            dr("IDDeposito") = datDpstoDestino.IDDeposito

            If ProcInfo.MultiplesOperaciones Then
                If Not datDpstoDestino.dtRegistro Is Nothing AndAlso datDpstoDestino.dtRegistro.Rows.Count > 0 Then
                    For Each col As DataColumn In dtOrigen.Columns
                        If datDpstoDestino.dtRegistro.Columns.Contains(col.ColumnName) Then
                            dr(col.ColumnName) = datDpstoDestino.dtRegistro.Rows(0)(col.ColumnName)
                        End If
                    Next
                End If
            Else

                dr("IDVino") = datDpstoDestino.IDVino

                Dim dtVino As DataTable = New BdgVino().SelOnPrimaryKey(datDpstoDestino.IDVino)
                If Not dtVino Is Nothing AndAlso dtVino.Rows.Count > 0 Then
                    dr("IDArticulo") = dtVino.Rows(0)("IDArticulo")
                    dr("IDAlmacen") = dtVino.Rows(0)("IDAlmacen")
                    dr("Lote") = dtVino.Rows(0)("Lote")
                End If

                Dim f As New Filter
                f.Add(New StringFilterItem("IDDeposito", datDpstoDestino.IDDeposito))
                If (Not datDpstoDestino.IDVino.Equals(Guid.Empty)) Then
                    f.Add(New GuidFilterItem("IDVino", datDpstoDestino.IDVino))
                End If
                Dim dtArticulo As DataTable = AdminData.GetData("NegBdgOperacionPlan", f, , "IDArticulo, Lote")
                If dtArticulo.Rows.Count > 0 Then
                    dr("IDUDMedida") = dtArticulo.Rows(0)("IDUDMedida")
                    dr("Cantidad") = dtArticulo.Rows(0)("Cantidad")
                    dr("Merma") = 0
                    dr("IDEstadoVino") = dtArticulo.Rows(0)("IDEstadoVino")
                    dr("TipoDeposito") = dtArticulo.Rows(0)("TipoDeposito")
                    dr("Litros") = dtArticulo.Rows(0)("Ocupacion")
                    dr("QDeposito") = dtArticulo.Rows(0)("Ocupacion")
                    dr("IDTipoMermaVino") = System.DBNull.Value
                    dr("Capacidad") = dtArticulo.Rows(0)("Capacidad")
                    dr("Ocupacion") = dtArticulo.Rows(0)("Ocupacion")
                End If
            End If
            dtOrigen.Rows.Add(dr)
        Next


        Dim datOrigen As New DataGetOrigenPropuesta(Nothing, dtOrigen)
        Return datOrigen
    End Function
    <Task()> Public Shared Function GetOrigenOperacionVinoDestinoOFs(ByVal oper As OperCab, ByVal services As ServiceProvider) As DataGetOrigenPropuesta
        Dim dtOrigen As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf GetEstructuraOrigenDestinoOperacionVino, Nothing, services)
        Dim datOF As DataOrdenFabricacion = services.GetService(Of DataOrdenFabricacion)()
        If Not datOF.ProductoTerminado Is Nothing AndAlso datOF.ProductoTerminado.Rows.Count > 0 Then
            For Each drPT As DataRow In datOF.ProductoTerminado.Rows
                Dim drNew As DataRow = dtOrigen.NewRow
                For Each col As DataColumn In dtOrigen.Columns
                    If datOF.ProductoTerminado.Columns.Contains(col.ColumnName) Then
                        drNew(col.ColumnName) = drPT(col.ColumnName)
                    End If
                Next
                drNew("IDDeposito") = drPT("IDUbicacion")
                drNew("IDUdMedida") = drPT("IDUdInterna")
                drNew("Cantidad") = oper.Doc.HeaderRow("CantidadOrden")
                drNew("IDOrden") = oper.Doc.HeaderRow("IDOrden")

                dtOrigen.Rows.Add(drNew)
            Next
        End If

        Dim datOrigen As New DataGetOrigenPropuesta(CType(oper, OperCabOFs).NOrden, dtOrigen)
        Return datOrigen
    End Function
    <Task()> Public Shared Function GetOrigenOperacionVinoDestinoExpedicion(ByVal oper As OperCab, ByVal services As ServiceProvider) As DataGetOrigenPropuesta
        Dim dtOrigen As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf GetEstructuraOrigenDestinoOperacionVino, Nothing, services)
        Dim Cab As OperCabExp = CType(oper, OperCabExp)
        Dim dtPVL As DataTable = New PedidoVentaLinea().SelOnPrimaryKey(Cab.IDLineaPedido)
        If dtPVL.Rows.Count > 0 Then
            If Not Cab.dtArtCompatibles Is Nothing AndAlso Cab.dtArtCompatibles.Rows.Count > 0 Then
                Dim drNew As DataRow = dtOrigen.NewRow
                drNew("IDDeposito") = Cab.IDDeposito
                drNew("IDArticulo") = dtPVL.Rows(0)("IDArticulo")
                drNew("IDAlmacen") = dtPVL.Rows(0)("IDAlmacen")

                drNew("Lote") = Cab.dtArtCompatibles.Rows(0)("Lote")
                drNew("Cantidad") = Cab.dtArtCompatibles.Rows(0)("Cantidad") * Cab.dtArtCompatibles.Rows(0)("Factor")
                dtOrigen.Rows.Add(drNew)
            End If
        End If

        Dim datOrigen As New DataGetOrigenPropuesta(Nothing, dtOrigen)
        Return datOrigen
    End Function


    <Task()> Public Shared Function GetEstructuraOrigenDestinoOperacionVino(ByVal oper As Object, ByVal services As ServiceProvider) As DataTable
        Dim dt As New DataTable
        dt.Columns.Add("IDDeposito", GetType(String))
        dt.Columns.Add("IDVino", GetType(Guid))
        dt.Columns.Add("IDArticulo", GetType(String))
        dt.Columns.Add("IDAlmacen", GetType(String))
        dt.Columns.Add("Lote", GetType(String))
        dt.Columns.Add("IDBarrica", GetType(String))
        dt.Columns.Add("IDUDMedida", GetType(String))
        dt.Columns.Add("Capacidad", GetType(Double))
        dt.Columns.Add("Ocupacion", GetType(Double))

        dt.Columns.Add("Cantidad", GetType(Double))
        dt.Columns.Add("Merma", GetType(Double))
        dt.Columns.Add("IDEstadoVino", GetType(String))
        dt.Columns.Add("TipoDeposito", GetType(Integer))
        dt.Columns.Add("QDeposito", GetType(Double))
        dt.Columns.Add("Litros", GetType(Double))
        dt.Columns.Add("IDTipoMermaVino", GetType(String))

        dt.Columns.Add("IdOrden", GetType(Integer))
        dt.Columns.Add("NOrden", GetType(String))
        Return dt
    End Function

    <Serializable()> _
    Public Class DataNuevaLineaOperacionVino
        Public doc As DocumentoBdgOperacion
        Public drOrigen As DataRow
        Public Destino As Integer '0=Origen, 1=Destino 

        Public Sub New(ByVal Destino As Integer, ByVal doc As DocumentoBdgOperacion, ByVal drOrigen As DataRow)
            Me.Destino = Destino
            Me.doc = doc
            Me.drOrigen = drOrigen
        End Sub
    End Class
    <Task()> Public Shared Sub NuevaLineaOperacionVino(ByVal data As DataNuevaLineaOperacionVino, ByVal services As ServiceProvider)

        Dim ProcInfo As ProcessInfoOperacion = services.GetService(Of ProcessInfoOperacion)()
        Dim OpVino As BusinessHelper
        Dim NOperacion As String = data.doc.FieldNOperacion
        Dim FieldIDLineaOrigen As String
        Dim FieldIDLineaDestino As String = data.doc.FieldIDLineaOperacionVino
        Select Case data.doc.Cabecera.TipoOperacion
            Case enumBdgOrigenOperacion.Planificada
                OpVino = New BdgOperacionVinoPlan
                ' FieldIDLineaOrigen = "IDLineaOperacionVinoPlan"
            Case enumBdgOrigenOperacion.Real
                OpVino = New BdgOperacionVino
                FieldIDLineaOrigen = "IDLineaOperacionVinoPlan"
        End Select

        Dim context As New BusinessData(data.doc.HeaderRow)
        Dim drOperacionVino As DataRow = data.doc.dtOperacionVino.NewRow
        drOperacionVino(NOperacion) = data.doc.HeaderRow(NOperacion)
        If data.Destino AndAlso Not data.doc.GetOperacionVinoOrigen Is Nothing AndAlso data.doc.GetOperacionVinoOrigen.Count > 0 Then
            context("OperacionVinoOrigenes") = data.doc.GetOperacionVinoOrigen.CopyToDataTable
        End If
        drOperacionVino = OpVino.ApplyBusinessRule("Destino", data.Destino, drOperacionVino, context)
        drOperacionVino = OpVino.ApplyBusinessRule("IDArticulo", data.drOrigen("IDArticulo"), drOperacionVino, context)
        If data.Destino Then
            Select Case data.doc.Cabecera.Origen
                Case OrigenOperacion.Depositos
                    If Length(drOperacionVino("IDEstadoVino")) > 0 Then data.drOrigen("IDEstadoVino") = System.DBNull.Value
            End Select
        End If


        If Length(data.drOrigen("IDDeposito")) > 0 Then
            drOperacionVino = OpVino.ApplyBusinessRule("IDDeposito", data.drOrigen("IDDeposito"), drOperacionVino, context)

            Dim stvino As New BdgGeneral.DataObtenerVinoEnDeposito(data.drOrigen("IDDeposito"), data.drOrigen("IDArticulo") & String.Empty, data.drOrigen("Lote") & String.Empty, data.drOrigen("IDBarrica") & String.Empty) ',  Nz(drOrigen("IDUDMedida"), String.Empty), drOrigen("Cantidad"))
            stvino = ProcessServer.ExecuteTask(Of BdgGeneral.DataObtenerVinoEnDeposito, BdgGeneral.DataObtenerVinoEnDeposito)(AddressOf BdgGeneral.ObtenerVinoEnDeposito, stvino, services)

            If Not stvino.dtVino Is Nothing AndAlso stvino.dtVino.Rows.Count > 0 Then
                'asignaremos el vino siempre y cuando el artículo que venía indicado sea el mismo qeu el que indica
                'o si la marca de AutoAsignarVino viene marcada => a esto igual hay qeu darle una vuelta TODO
                If Nz(drOperacionVino("IDArticulo"), String.Empty) = stvino.dtVino.Rows(0)("IDArticulo") OrElse ProcInfo.AutoCalcularVino Then
                    'DrDestino("IDVino") = stvino.dtVino.Rows(0)("IDVino")
                    Select Case data.doc.Cabecera.TipoOperacion
                        Case enumBdgOrigenOperacion.Real
                            drOperacionVino(FieldIDLineaDestino) = stvino.IDVino
                    End Select
                    drOperacionVino("Lote") = stvino.dtVino.Rows(0)("Lote")
                    drOperacionVino("IDArticulo") = stvino.dtVino.Rows(0)("IDArticulo")
                    drOperacionVino("DescArticulo") = stvino.dtVino.Rows(0)("DescArticulo")
                    drOperacionVino("IDUdMedida") = stvino.dtVino.Rows(0)("IDUdMedida")
                    If drOperacionVino.Table.Columns.Contains("Fecha") Then drOperacionVino("Fecha") = stvino.dtVino.Rows(0)("Fecha")

                    drOperacionVino = OpVino.ApplyBusinessRule("IDArticulo", drOperacionVino("IDArticulo"), drOperacionVino, context)
                End If
            End If
        End If

        If IsDBNull(drOperacionVino(FieldIDLineaDestino)) OrElse drOperacionVino(FieldIDLineaDestino).Equals(Guid.Empty) Then
            drOperacionVino(FieldIDLineaDestino) = Guid.NewGuid
        End If
        If Length(data.drOrigen("IDUdMedida")) > 0 Then drOperacionVino = OpVino.ApplyBusinessRule("IDUdMedida", data.drOrigen("IDUdMedida"), drOperacionVino, context)
        drOperacionVino = OpVino.ApplyBusinessRule("Cantidad", data.drOrigen("Cantidad"), drOperacionVino, context)
        drOperacionVino = OpVino.ApplyBusinessRule("Merma", data.drOrigen("Merma"), drOperacionVino, context)

        If Length(data.drOrigen("Lote")) > 0 Then drOperacionVino = OpVino.ApplyBusinessRule("Lote", data.drOrigen("Lote"), drOperacionVino, context)
        If Length(data.drOrigen("IDEstadoVino")) > 0 Then drOperacionVino = OpVino.ApplyBusinessRule("IDEstadoVino", data.drOrigen("IDEstadoVino"), drOperacionVino, context)
        If Nz(data.drOrigen("TipoDeposito"), -1) <> -1 Then drOperacionVino = OpVino.ApplyBusinessRule("TipoDeposito", data.drOrigen("TipoDeposito"), drOperacionVino, context)

        If Not IsDBNull(data.drOrigen("QDeposito")) Then drOperacionVino = OpVino.ApplyBusinessRule("QDeposito", data.drOrigen("QDeposito"), drOperacionVino, context)
        If Not IsDBNull(data.drOrigen("Litros")) Then drOperacionVino = OpVino.ApplyBusinessRule("Litros", data.drOrigen("Litros"), drOperacionVino, context)
        If Length(data.drOrigen("IDBarrica")) > 0 Then drOperacionVino = OpVino.ApplyBusinessRule("IDBarrica", data.drOrigen("IDBarrica"), drOperacionVino, context)
        If Length(data.drOrigen("IDTipoMermaVino")) > 0 Then drOperacionVino = OpVino.ApplyBusinessRule("IDTipoMermaVino", Nz(data.drOrigen("IDTipoMermaVino"), String.Empty), drOperacionVino, context)

        If data.drOrigen.Table.Columns.Contains("IDOrden") AndAlso Nz(data.drOrigen("IDOrden"), 0) <> 0 Then drOperacionVino("IDOrden") = data.drOrigen("IDOrden")
        If data.drOrigen.Table.Columns.Contains("JaulonDesde") AndAlso Not IsDBNull(data.drOrigen("JaulonDesde")) Then drOperacionVino("JaulonDesde") = data.drOrigen("JaulonDesde")
        If data.drOrigen.Table.Columns.Contains("JaulonHasta") AndAlso Not IsDBNull(data.drOrigen("JaulonHasta")) Then drOperacionVino("JaulonHasta") = data.drOrigen("JaulonHasta")


        Dim IDVino As Guid
        Select Case data.doc.Cabecera.TipoOperacion
            Case enumBdgOrigenOperacion.Real
                IDVino = Nz(drOperacionVino("IDVino"), Guid.Empty)
        End Select

        Dim datOcupacion As New BdgGeneral.DataCalculoOcupacion(drOperacionVino("IDDeposito") & String.Empty, drOperacionVino("IDArticulo") & String.Empty, drOperacionVino("Lote") & String.Empty, drOperacionVino("IDBarrica") & String.Empty, drOperacionVino("IDUDMedida") & String.Empty, IDVino)
        Dim Ocupacion As Double = ProcessServer.ExecuteTask(Of BdgGeneral.DataCalculoOcupacion, Double)(AddressOf BdgGeneral.CalculoOcupacion, datOcupacion, services)
        drOperacionVino = OpVino.ApplyBusinessRule("Ocupacion", Ocupacion, drOperacionVino, context)
        data.doc.dtOperacionVino.Rows.Add(drOperacionVino)


        If data.Destino = 1 Then
            Dim datImput As New DataGenerarPropuestaImputacionesLinea(data, drOperacionVino, FieldIDLineaOrigen)
            ProcessServer.ExecuteTask(Of DataGenerarPropuestaImputacionesLinea)(AddressOf GenerarPropuestaImputacionesLinea, datImput, services)
        End If
    End Sub

    Public Class DataGenerarPropuestaImputacionesLinea
        Public OrigenNuevaLinea As DataNuevaLineaOperacionVino
        Public drOperacionVino As DataRow
        Public FieldIDLineaOrigen As String

        Public Sub New(ByVal OrigenNuevaLinea As DataNuevaLineaOperacionVino, ByVal drOperacionVino As DataRow, ByVal FieldIDLineaOrigen As String)
            Me.OrigenNuevaLinea = OrigenNuevaLinea
            Me.drOperacionVino = drOperacionVino
            Me.FieldIDLineaOrigen = FieldIDLineaOrigen
        End Sub
    End Class
    <Task()> Public Shared Sub GenerarPropuestaImputacionesLinea(ByVal data As DataGenerarPropuestaImputacionesLinea, ByVal services As ServiceProvider)
        Dim IDLineaOrigen As Guid
        Dim IDLineaDestino As Guid

        Select Case data.OrigenNuevaLinea.doc.Cabecera.TipoOperacion
            Case enumBdgOrigenOperacion.Real
                Select Case data.OrigenNuevaLinea.doc.Cabecera.Origen
                    Case OrigenOperacion.OperacionPlanificada
                        IDLineaOrigen = data.OrigenNuevaLinea.drOrigen(data.FieldIDLineaOrigen)
                        IDLineaDestino = data.drOperacionVino("IDVino")
                    Case OrigenOperacion.Depositos
                        IDLineaOrigen = Guid.Empty
                        IDLineaDestino = data.drOperacionVino(data.OrigenNuevaLinea.doc.FieldIDLineaOperacionVino)
                End Select
                Dim datImputaVino As New DataPropuestaImputacionesVino(data.OrigenNuevaLinea.doc, IDLineaOrigen, IDLineaDestino, data.drOperacionVino)
                ProcessServer.ExecuteTask(Of DataPropuestaImputacionesVino)(AddressOf PropuestaImputacionesVino, datImputaVino, services)

            Case enumBdgOrigenOperacion.Planificada
                Select Case data.OrigenNuevaLinea.doc.Cabecera.Origen
                    Case OrigenOperacion.Depositos
                        IDLineaOrigen = Guid.Empty
                        IDLineaDestino = data.drOperacionVino(data.OrigenNuevaLinea.doc.FieldIDLineaOperacionVino)
                End Select
                Dim datImputaVino As New DataPropuestaImputacionesVino(data.OrigenNuevaLinea.doc, IDLineaOrigen, IDLineaDestino, data.drOperacionVino)
                ProcessServer.ExecuteTask(Of DataPropuestaImputacionesVino)(AddressOf PropuestaImputacionesVino, datImputaVino, services)
        End Select
    End Sub

#End Region

#Region " Propuesta Imputaciones Globales "


    <Task()> Public Shared Sub PropuestaImputacionesGlobales(ByVal doc As DocumentoBdgOperacion, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DocumentoBdgOperacion)(AddressOf PropuestaImputacionesGlobalesMateriales, doc, services)
        ProcessServer.ExecuteTask(Of DocumentoBdgOperacion)(AddressOf PropuestaImputacionesGlobalesMOD, doc, services)
        ProcessServer.ExecuteTask(Of DocumentoBdgOperacion)(AddressOf PropuestaImputacionesGlobalesCentro, doc, services)
        ProcessServer.ExecuteTask(Of DocumentoBdgOperacion)(AddressOf PropuestaImputacionesGlobalesVarios, doc, services)
    End Sub

    <Task()> Public Shared Sub PropuestaImputacionesGlobalesMateriales(ByVal doc As DocumentoBdgOperacion, ByVal services As ServiceProvider)
        If Not doc.HeaderRow Is Nothing Then

            If doc.Cabecera Is Nothing Then Exit Sub

            Dim datOrigen As DataGetOrigenPropuesta = ProcessServer.ExecuteTask(Of OperCab, DataGetOrigenPropuesta)(AddressOf GetOrigenOperacionMaterial, doc.Cabecera, services)
            If Not datOrigen Is Nothing AndAlso Not datOrigen.dtOrigen Is Nothing AndAlso datOrigen.dtOrigen.Rows.Count > 0 Then
                For Each drOrigen As DataRow In datOrigen.dtOrigen.Rows
                    Dim datNewLine As New DataNuevaLineaOperacionImputacion(doc, drOrigen, Guid.Empty, Nothing)
                    ProcessServer.ExecuteTask(Of DataNuevaLineaOperacionImputacion)(AddressOf NuevaLineaOperacionMaterial, datNewLine, services)
                Next
            End If
        End If
    End Sub
    <Task()> Public Shared Function GetOrigenOperacionMaterial(ByVal oper As OperCab, ByVal services As ServiceProvider) As DataGetOrigenPropuesta
        Select Case oper.Origen
            Case OrigenOperacion.OperacionPlanificada
                Return ProcessServer.ExecuteTask(Of OperCab, DataGetOrigenPropuesta)(AddressOf GetOrigenOperacionMaterialOpPlanificada, oper, services)
            Case Else
                Return ProcessServer.ExecuteTask(Of OperCab, DataGetOrigenPropuesta)(AddressOf GetOrigenOperacionMaterialPorTipoOp, oper, services)
        End Select
    End Function
    <Task()> Public Shared Function GetOrigenOperacionMaterialOpPlanificada(ByVal oper As OperCab, ByVal services As ServiceProvider) As DataGetOrigenPropuesta
        Dim f As New Filter
        f.Add(New FilterItem("NOperacionPlan", CType(oper, OperCabPlanificadas).IDOrigen))

        Dim dtOperacionMaterial As DataTable = New BE.DataEngine().Filter("frmBdgOperacionPlanMaterialGlobal", f)
        Dim datOrigen As New DataGetOrigenPropuesta(CType(oper, OperCabPlanificadas).IDOrigen, dtOperacionMaterial)
        Return datOrigen
    End Function
    <Task()> Public Shared Function GetOrigenOperacionMaterialPorTipoOp(ByVal oper As OperCab, ByVal services As ServiceProvider) As DataGetOrigenPropuesta
        Dim datImputGlobales As DataImputacionesGlobales = services.GetService(Of DataImputacionesGlobales)()
        Dim datOrigen As New DataGetOrigenPropuesta(Nothing, datImputGlobales.dtOperacionMaterial)
        Return datOrigen
    End Function

    <Serializable()> _
   Public Class DataNuevaLineaOperacionImputacion
        Public doc As DocumentoBdgOperacion
        Public drOrigen As DataRow
        Public drOperacionVino As DataRow
        Public IDLineaOperacionVino As Guid

        Public Sub New(ByVal doc As DocumentoBdgOperacion, ByVal drOrigen As DataRow, ByVal IDLineaOperacionVino As Guid, ByVal drOperacionVino As DataRow)
            Me.doc = doc
            Me.drOrigen = drOrigen
            Me.IDLineaOperacionVino = IDLineaOperacionVino
            Me.drOperacionVino = drOperacionVino
        End Sub
    End Class
    <Task()> Public Shared Sub NuevaLineaOperacionMaterial(ByVal data As DataNuevaLineaOperacionImputacion, ByVal services As ServiceProvider)
        Dim OpImp As BusinessHelper
        Dim NOperacion As String = data.doc.FieldNOperacion
        Dim FieldIDMaterialDestino As String = data.doc.FieldIDLineaMaterialGlobal
        Dim FieldIDMaterialOrigen As String
        Dim FiledIDLote As String = data.doc.FieldIDLineaMaterialGlobalLote
        Dim View_Lotes_Origen As String
        Select Case data.doc.Cabecera.TipoOperacion
            Case enumBdgOrigenOperacion.Planificada
                OpImp = New BdgOperacionPlanMaterial

                'FieldIDMaterialOrigen = "IDOperacionPlanMaterial"

            Case enumBdgOrigenOperacion.Real
                OpImp = New BdgOperacionMaterial
                Select Case data.doc.Cabecera.Origen
                    Case OrigenOperacion.OperacionPlanificada
                        View_Lotes_Origen = "tbBdgOperacionPlanMaterialLote"
                        FieldIDMaterialOrigen = "IDOperacionPlanMaterial"
                End Select
        End Select

        '//Material
        Dim context As New BusinessData(data.doc.HeaderRow)
        Dim DrDestino As DataRow = data.doc.dtOperacionMaterial.NewRow
        DrDestino(NOperacion) = data.doc.HeaderRow(NOperacion)
        DrDestino(FieldIDMaterialDestino) = Guid.NewGuid

        If Length(data.drOrigen("IDArticulo")) > 0 Then DrDestino = OpImp.ApplyBusinessRule("IDArticulo", data.drOrigen("IDArticulo"), DrDestino, context)
        If Length(data.drOrigen("IDAlmacen")) > 0 Then DrDestino = OpImp.ApplyBusinessRule("IDAlmacen", data.drOrigen("IDAlmacen"), DrDestino, context)
        DrDestino = OpImp.ApplyBusinessRule("RecalcularMaterial", Nz(data.drOrigen("RecalcularMaterial"), False), DrDestino, context)
        DrDestino = OpImp.ApplyBusinessRule("Cantidad", Nz(data.drOrigen("Cantidad"), 0), DrDestino, context)
        DrDestino = OpImp.ApplyBusinessRule("Merma", Nz(data.drOrigen("Merma"), 0), DrDestino, context)
        DrDestino = OpImp.ApplyBusinessRule("Fecha", Nz(data.doc.HeaderRow("Fecha"), Today), DrDestino, context) '????
        data.doc.dtOperacionMaterial.Rows.Add(DrDestino)



        '//Lotes
        If Length(View_Lotes_Origen) > 0 AndAlso Length(FieldIDMaterialOrigen) > 0 Then

            Dim dtLotesOrigen As DataTable = New BE.DataEngine().Filter(View_Lotes_Origen, New GuidFilterItem(FieldIDMaterialOrigen, data.drOrigen(FieldIDMaterialOrigen)))
            If Not dtLotesOrigen Is Nothing AndAlso dtLotesOrigen.Rows.Count > 0 Then
                For Each drLote As DataRow In dtLotesOrigen.Rows
                    Dim drNewLote As DataRow = data.doc.dtOperacionMaterialLote.NewRow
                    For Each col As DataColumn In dtLotesOrigen.Columns
                        If (dtLotesOrigen.Columns.Contains(col.ColumnName) AndAlso _
                            data.doc.dtOperacionMaterialLote.Columns.Contains(col.ColumnName) AndAlso _
                            Length(drLote(col.ColumnName))) Then
                            drNewLote(col.ColumnName) = drLote(col.ColumnName)
                        End If
                    Next

                    drNewLote(FieldIDMaterialDestino) = DrDestino(FieldIDMaterialDestino)
                    drNewLote(FiledIDLote) = Guid.NewGuid
                    If data.doc.dtOperacionMaterialLote.Columns.Contains(NOperacion) Then drNewLote(NOperacion) = data.doc.HeaderRow(NOperacion)
                    drNewLote("IDArticulo") = DrDestino("IDArticulo")
                    drNewLote("IDAlmacen") = DrDestino("IDAlmacen")

                    data.doc.dtOperacionMaterialLote.Rows.Add(drNewLote)
                Next
            End If
        End If



        DrDestino("NumLotes") = 0
        If Not data.doc.dtOperacionMaterialLote Is Nothing AndAlso data.doc.dtOperacionMaterialLote.Rows.Count > 0 Then
            Dim LotesLinea As List(Of DataRow) = (From c In data.doc.dtOperacionMaterialLote Where Not c.IsNull(FieldIDMaterialDestino) AndAlso c(FieldIDMaterialDestino) = DrDestino(FieldIDMaterialDestino) Select c).ToList
            If Not LotesLinea Is Nothing AndAlso LotesLinea.Count > 0 Then
                DrDestino("NumLotes") = LotesLinea.Count
                If LotesLinea.Count = 1 Then
                    DrDestino("Lote") = LotesLinea(0)("Lote")
                    DrDestino("Ubicacion") = LotesLinea(0)("Ubicacion")
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub PropuestaImputacionesGlobalesMOD(ByVal doc As DocumentoBdgOperacion, ByVal services As ServiceProvider)
        If Not doc.HeaderRow Is Nothing Then

            If doc.Cabecera Is Nothing Then Exit Sub

            Dim datOrigen As DataGetOrigenPropuesta = ProcessServer.ExecuteTask(Of OperCab, DataGetOrigenPropuesta)(AddressOf GetOrigenOperacionMOD, doc.Cabecera, services)
            If Not datOrigen Is Nothing AndAlso Not datOrigen.dtOrigen Is Nothing AndAlso datOrigen.dtOrigen.Rows.Count > 0 Then
                For Each drOrigen As DataRow In datOrigen.dtOrigen.Rows
                    Dim datNewLine As New DataNuevaLineaOperacionImputacion(doc, drOrigen, Guid.Empty, Nothing)
                    ProcessServer.ExecuteTask(Of DataNuevaLineaOperacionImputacion)(AddressOf NuevaLineaOperacionMOD, datNewLine, services)
                Next
            End If
        End If
    End Sub
    <Task()> Public Shared Function GetOrigenOperacionMOD(ByVal oper As OperCab, ByVal services As ServiceProvider) As DataGetOrigenPropuesta
        Select Case oper.Origen
            Case OrigenOperacion.OperacionPlanificada
                Return ProcessServer.ExecuteTask(Of OperCab, DataGetOrigenPropuesta)(AddressOf GetOrigenOperacionMODOpPlanificada, oper, services)
            Case Else
                Return ProcessServer.ExecuteTask(Of OperCab, DataGetOrigenPropuesta)(AddressOf GetOrigenOperacionMODPorTipoOp, oper, services)
        End Select
    End Function
    <Task()> Public Shared Function GetOrigenOperacionMODOpPlanificada(ByVal oper As OperCab, ByVal services As ServiceProvider) As DataGetOrigenPropuesta
        Dim NOperacionPlan As String = CType(oper, OperCabPlanificadas).IDOrigen
        Dim f As New Filter
        f.Add(New FilterItem("NOperacionPlan", NOperacionPlan))

        Dim dtOperacionMOD As DataTable = New BE.DataEngine().Filter("frmBdgOperacionPlanMODGlobal", f)
        Dim datOrigen As New DataGetOrigenPropuesta(NOperacionPlan, dtOperacionMOD)
        Return datOrigen
    End Function
    <Task()> Public Shared Function GetOrigenOperacionMODPorTipoOp(ByVal oper As OperCab, ByVal services As ServiceProvider) As DataGetOrigenPropuesta
        Dim datImputGlobales As DataImputacionesGlobales = services.GetService(Of DataImputacionesGlobales)()
        Dim datOrigen As New DataGetOrigenPropuesta(Nothing, datImputGlobales.dtOperacionMOD)
        Return datOrigen
    End Function
    <Task()> Public Shared Sub NuevaLineaOperacionMOD(ByVal data As DataNuevaLineaOperacionImputacion, ByVal services As ServiceProvider)
        Dim OpImp As BusinessHelper
        Dim NOperacion As String = data.doc.FieldNOperacion
        Dim FieldIDMODDestino As String = data.doc.FieldIDLineaMODGlobal
        Dim FieldIDMODOrigen As String

        Select Case data.doc.Cabecera.TipoOperacion
            Case enumBdgOrigenOperacion.Planificada
                OpImp = New BdgOperacionPlanMOD
            Case enumBdgOrigenOperacion.Real
                OpImp = New BdgOperacionMOD

                Select Case data.doc.Cabecera.Origen
                    Case OrigenOperacion.OperacionPlanificada
                        FieldIDMODOrigen = "IDOperacionPlanMOD"
                End Select
        End Select


        Dim context As New BusinessData(data.doc.HeaderRow)
        Dim DrDestino As DataRow = data.doc.dtOperacionMOD.NewRow
        DrDestino(NOperacion) = data.doc.HeaderRow(NOperacion)
        DrDestino(FieldIDMODDestino) = Guid.NewGuid

        If Length(data.drOrigen("IDOperario")) > 0 Then DrDestino = OpImp.ApplyBusinessRule("IDOperario", data.drOrigen("IDOperario"), DrDestino, context)
        DrDestino = OpImp.ApplyBusinessRule("Fecha", Nz(data.doc.HeaderRow("Fecha"), Today), DrDestino, context)
        If Length(data.drOrigen("IDCategoria")) > 0 Then DrDestino = OpImp.ApplyBusinessRule("IDCategoria", data.drOrigen("IDCategoria"), DrDestino, context)
        If Length(data.drOrigen("Tiempo")) > 0 Then DrDestino = OpImp.ApplyBusinessRule("Tiempo", data.drOrigen("Tiempo"), DrDestino, context)
        data.doc.dtOperacionMOD.Rows.Add(DrDestino)
    End Sub


    <Task()> Public Shared Sub PropuestaImputacionesGlobalesCentro(ByVal doc As DocumentoBdgOperacion, ByVal services As ServiceProvider)
        If Not doc.HeaderRow Is Nothing Then

            If doc.Cabecera Is Nothing Then Exit Sub

            Dim datOrigen As DataGetOrigenPropuesta = ProcessServer.ExecuteTask(Of OperCab, DataGetOrigenPropuesta)(AddressOf GetOrigenOperacionCentro, doc.Cabecera, services)
            If Not datOrigen Is Nothing AndAlso Not datOrigen.dtOrigen Is Nothing AndAlso datOrigen.dtOrigen.Rows.Count > 0 Then
                For Each drOrigen As DataRow In datOrigen.dtOrigen.Rows
                    Dim datNewLine As New DataNuevaLineaOperacionImputacion(doc, drOrigen, Guid.Empty, Nothing)
                    ProcessServer.ExecuteTask(Of DataNuevaLineaOperacionImputacion)(AddressOf NuevaLineaOperacionCentro, datNewLine, services)
                Next
            End If
        End If
    End Sub
    <Task()> Public Shared Function GetOrigenOperacionCentro(ByVal oper As OperCab, ByVal services As ServiceProvider) As DataGetOrigenPropuesta
        Select Case oper.Origen
            Case OrigenOperacion.OperacionPlanificada
                Return ProcessServer.ExecuteTask(Of OperCab, DataGetOrigenPropuesta)(AddressOf GetOrigenOperacionCentroOpPlanificada, oper, services)
            Case Else
                Return ProcessServer.ExecuteTask(Of OperCab, DataGetOrigenPropuesta)(AddressOf GetOrigenOperacionCentroPorTipoOp, oper, services)
        End Select
    End Function
    <Task()> Public Shared Function GetOrigenOperacionCentroOpPlanificada(ByVal oper As OperCab, ByVal services As ServiceProvider) As DataGetOrigenPropuesta
        Dim NOperacionPlan As String = CType(oper, OperCabPlanificadas).IDOrigen
        Dim f As New Filter
        f.Add(New FilterItem("NOperacionPlan", NOperacionPlan))

        Dim dtOperacionCentro As DataTable = New BE.DataEngine().Filter("frmBdgOperacionPlanCentroGlobal", f)
        Dim datOrigen As New DataGetOrigenPropuesta(NOperacionPlan, dtOperacionCentro)
        Return datOrigen
    End Function
    <Task()> Public Shared Function GetOrigenOperacionCentroPorTipoOp(ByVal oper As OperCab, ByVal services As ServiceProvider) As DataGetOrigenPropuesta
        Dim datImputGlobales As DataImputacionesGlobales = services.GetService(Of DataImputacionesGlobales)()
        Dim datOrigen As New DataGetOrigenPropuesta(Nothing, datImputGlobales.dtOperacionCentro)
        Return datOrigen
    End Function
    <Task()> Public Shared Sub NuevaLineaOperacionCentro(ByVal data As DataNuevaLineaOperacionImputacion, ByVal services As ServiceProvider)
        Dim OpImp As BusinessHelper
        Dim NOperacion As String = data.doc.FieldNOperacion
        Dim FieldIDCentroDestino As String = data.doc.FieldIDLineaCentroGlobal
        Dim FieldIDCentroOrigen As String

        Select Case data.doc.Cabecera.TipoOperacion
            Case enumBdgOrigenOperacion.Planificada
                OpImp = New BdgOperacionPlanCentro
            Case enumBdgOrigenOperacion.Real
                OpImp = New BdgOperacionCentro

                Select Case data.doc.Cabecera.Origen
                    Case OrigenOperacion.OperacionPlanificada
                        FieldIDCentroOrigen = "IDOperacionPlanCentro"
                End Select
        End Select


        Dim context As New BusinessData(data.doc.HeaderRow)
        Dim DrDestino As DataRow = data.doc.dtOperacionCentro.NewRow
        DrDestino(NOperacion) = data.doc.HeaderRow(NOperacion)
        DrDestino(FieldIDCentroDestino) = Guid.NewGuid

        If Length(data.drOrigen("IDCentro")) > 0 Then DrDestino = OpImp.ApplyBusinessRule("IDCentro", data.drOrigen("IDCentro"), DrDestino, context)
        DrDestino = OpImp.ApplyBusinessRule("Fecha", Nz(data.doc.HeaderRow("Fecha"), Now), DrDestino, context)
        If Length(data.drOrigen("UDTiempo")) > 0 Then DrDestino = OpImp.ApplyBusinessRule("UDTiempo", data.drOrigen("UDTiempo"), DrDestino, context)
        DrDestino = OpImp.ApplyBusinessRule("Tiempo", Nz(data.drOrigen("Tiempo"), 0), DrDestino, context)
        If data.drOrigen("PorCantidad") Then DrDestino = OpImp.ApplyBusinessRule("PorCantidad", Nz(data.drOrigen("PorCantidad"), False), DrDestino, context)
        If Length(data.drOrigen("IDIncidencia")) > 0 Then DrDestino = OpImp.ApplyBusinessRule("IDIncidencia", data.drOrigen("IDIncidencia"), DrDestino, context)
        data.doc.dtOperacionCentro.Rows.Add(DrDestino)
    End Sub


    <Task()> Public Shared Sub PropuestaImputacionesGlobalesVarios(ByVal doc As DocumentoBdgOperacion, ByVal services As ServiceProvider)
        If Not doc.HeaderRow Is Nothing Then

            If doc.Cabecera Is Nothing Then Exit Sub

            Dim datOrigen As DataGetOrigenPropuesta = ProcessServer.ExecuteTask(Of OperCab, DataGetOrigenPropuesta)(AddressOf GetOrigenOperacionVarios, doc.Cabecera, services)
            If Not datOrigen Is Nothing AndAlso Not datOrigen.dtOrigen Is Nothing AndAlso datOrigen.dtOrigen.Rows.Count > 0 Then
                For Each drOrigen As DataRow In datOrigen.dtOrigen.Rows
                    Dim datNewLine As New DataNuevaLineaOperacionImputacion(doc, drOrigen, Guid.Empty, Nothing)
                    ProcessServer.ExecuteTask(Of DataNuevaLineaOperacionImputacion)(AddressOf NuevaLineaOperacionVarios, datNewLine, services)
                Next
            End If
        End If
    End Sub
    <Task()> Public Shared Function GetOrigenOperacionVarios(ByVal oper As OperCab, ByVal services As ServiceProvider) As DataGetOrigenPropuesta
        Select Case oper.Origen
            Case OrigenOperacion.OperacionPlanificada
                Return ProcessServer.ExecuteTask(Of OperCab, DataGetOrigenPropuesta)(AddressOf GetOrigenOperacionVariosOpPlanificada, oper, services)
            Case Else
                Return ProcessServer.ExecuteTask(Of OperCab, DataGetOrigenPropuesta)(AddressOf GetOrigenOperacionVariosPorTipoOp, oper, services)
        End Select
    End Function
    <Task()> Public Shared Function GetOrigenOperacionVariosOpPlanificada(ByVal oper As OperCab, ByVal services As ServiceProvider) As DataGetOrigenPropuesta
        Dim NOperacionPlan As String = CType(oper, OperCabPlanificadas).IDOrigen
        Dim f As New Filter
        f.Add(New FilterItem("NOperacionPlan", NOperacionPlan))

        Dim dtOperacionVarios As DataTable = New BE.DataEngine().Filter("frmBdgOperacionPlanVariosGlobal", f)
        Dim datOrigen As New DataGetOrigenPropuesta(NOperacionPlan, dtOperacionVarios)
        Return datOrigen
    End Function
    <Task()> Public Shared Function GetOrigenOperacionVariosPorTipoOp(ByVal oper As OperCab, ByVal services As ServiceProvider) As DataGetOrigenPropuesta
        Dim datImputGlobales As DataImputacionesGlobales = services.GetService(Of DataImputacionesGlobales)()
        Dim datOrigen As New DataGetOrigenPropuesta(Nothing, datImputGlobales.dtOperacionVarios)
        Return datOrigen
    End Function
    <Task()> Public Shared Sub NuevaLineaOperacionVarios(ByVal data As DataNuevaLineaOperacionImputacion, ByVal services As ServiceProvider)
        Dim OpImp As BusinessHelper
        Dim NOperacion As String = data.doc.FieldNOperacion
        Dim FieldIDVariosDestino As String = data.doc.FieldIDLineaVariosGlobal
        Dim FieldIDVariosOrigen As String

        Select Case data.doc.Cabecera.TipoOperacion
            Case enumBdgOrigenOperacion.Planificada
                OpImp = New BdgOperacionPlanVarios

            Case enumBdgOrigenOperacion.Real
                OpImp = New BdgOperacionVarios

                Select Case data.doc.Cabecera.Origen
                    Case OrigenOperacion.OperacionPlanificada
                        FieldIDVariosOrigen = "IDOperacionPlanVarios"
                End Select
        End Select


        Dim context As New BusinessData(data.doc.HeaderRow)
        Dim DrDestino As DataRow = data.doc.dtOperacionVarios.NewRow
        DrDestino(NOperacion) = data.doc.HeaderRow(NOperacion)
        DrDestino(FieldIDVariosDestino) = Guid.NewGuid

        If Length(data.drOrigen("IDVarios")) > 0 Then DrDestino = OpImp.ApplyBusinessRule("IDVarios", data.drOrigen("IDVarios"), DrDestino, context)
        DrDestino = OpImp.ApplyBusinessRule("Fecha", Nz(data.doc.HeaderRow("Fecha"), Now), DrDestino, context)
        DrDestino = OpImp.ApplyBusinessRule("Cantidad", Nz(data.drOrigen("Cantidad"), 0), DrDestino, context)
        DrDestino = OpImp.ApplyBusinessRule("Tasa", Nz(data.drOrigen("Tasa"), 0), DrDestino, context)
        data.doc.dtOperacionVarios.Rows.Add(DrDestino)
    End Sub

#End Region

#Region " Propuesta Imputaciones Vino "

    <Serializable()> _
    Public Class DataPropuestaImputacionesVino
        Public doc As DocumentoBdgOperacion
        Public IDLineaOperacionVinoOrigen As Guid
        Public IDLineaOperacionVinoDestino As Guid

        Public drOperacionVino As DataRow
        Public IDLineaOperacionVino As Guid

        Public Sub New(ByVal doc As DocumentoBdgOperacion, ByVal IDLineaOrigen As Guid, ByVal IDLineaDestino As Guid, ByVal drOperacionVino As DataRow)
            Me.doc = doc
            Me.IDLineaOperacionVinoOrigen = IDLineaOrigen
            Me.IDLineaOperacionVinoDestino = IDLineaDestino
            Me.drOperacionVino = drOperacionVino
        End Sub
    End Class
    <Task()> Public Shared Sub PropuestaImputacionesVino(ByVal doc As DataPropuestaImputacionesVino, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DataPropuestaImputacionesVino)(AddressOf PropuestaImputacionesVinoMateriales, doc, services)
        ProcessServer.ExecuteTask(Of DataPropuestaImputacionesVino)(AddressOf PropuestaImputacionesVinoMOD, doc, services)
        ProcessServer.ExecuteTask(Of DataPropuestaImputacionesVino)(AddressOf PropuestaImputacionesVinoCentro, doc, services)
        ProcessServer.ExecuteTask(Of DataPropuestaImputacionesVino)(AddressOf PropuestaImputacionesVinoVarios, doc, services)
        ProcessServer.ExecuteTask(Of DataPropuestaImputacionesVino)(AddressOf PropuestaImputacionesVinoAnalisis, doc, services)
    End Sub

    <Task()> Public Shared Sub PropuestaImputacionesVinoMateriales(ByVal data As DataPropuestaImputacionesVino, ByVal services As ServiceProvider)
        If Not data.doc.HeaderRow Is Nothing Then

            If data.doc.Cabecera Is Nothing Then Exit Sub

            Dim datOrigen As DataGetOrigenPropuesta = ProcessServer.ExecuteTask(Of DataPropuestaImputacionesVino, DataGetOrigenPropuesta)(AddressOf GetOrigenOperacionVinoMaterial, data, services)
            If Not datOrigen Is Nothing AndAlso Not datOrigen.dtOrigen Is Nothing AndAlso datOrigen.dtOrigen.Rows.Count > 0 Then
                For Each drOrigen As DataRow In datOrigen.dtOrigen.Rows
                    Dim datNewLine As New DataNuevaLineaOperacionImputacion(data.doc, drOrigen, data.IDLineaOperacionVinoDestino, data.drOperacionVino)
                    ProcessServer.ExecuteTask(Of DataNuevaLineaOperacionImputacion)(AddressOf NuevaLineaOperacionVinoMaterial, datNewLine, services)
                Next
            End If
        End If
    End Sub
    <Task()> Public Shared Function GetOrigenOperacionVinoMaterial(ByVal data As DataPropuestaImputacionesVino, ByVal services As ServiceProvider) As DataGetOrigenPropuesta
        Select Case data.doc.Cabecera.Origen
            Case OrigenOperacion.OperacionPlanificada
                Return ProcessServer.ExecuteTask(Of DataPropuestaImputacionesVino, DataGetOrigenPropuesta)(AddressOf GetOrigenOperacionVinoMaterialOpPlanificada, data, services)
                'Case OrigenOperacion.Depositos
                '    Return ProcessServer.ExecuteTask(Of DataPropuestaImputacionesVino, DataGetOrigenPropuesta)(AddressOf GetOrigenOperacionVinoMaterialPorTipoOp, data, services)
            Case OrigenOperacion.OrdenFabricacion
                Return ProcessServer.ExecuteTask(Of DataPropuestaImputacionesVino, DataGetOrigenPropuesta)(AddressOf GetOrigenOperacionVinoMaterialOFs, data, services)
            Case Else
                Return ProcessServer.ExecuteTask(Of DataPropuestaImputacionesVino, DataGetOrigenPropuesta)(AddressOf GetOrigenOperacionVinoMaterialPorTipoOp, data, services)
        End Select
    End Function
    <Task()> Public Shared Function GetOrigenOperacionVinoMaterialOpPlanificada(ByVal data As DataPropuestaImputacionesVino, ByVal services As ServiceProvider) As DataGetOrigenPropuesta
        Dim NOperacionPlan As String = CType(data.doc.Cabecera, OperCabPlanificadas).IDOrigen
        Dim f As New Filter
        f.Add(New FilterItem("NOperacionPlan", NOperacionPlan))
        f.Add(New FilterItem("IDLineaOperacionVinoPlan", data.IDLineaOperacionVinoOrigen))
        Dim dtOperacionVinoMaterial As DataTable = New BE.DataEngine().Filter("frmBdgOperacionVinoPlanMaterial", f)
        Dim datOrigen As New DataGetOrigenPropuesta(NOperacionPlan, dtOperacionVinoMaterial)
        Return datOrigen
    End Function
    <Task()> Public Shared Function GetOrigenOperacionVinoMaterialPorTipoOp(ByVal data As DataPropuestaImputacionesVino, ByVal services As ServiceProvider) As DataGetOrigenPropuesta
        Dim datImputa As New BdgGeneral.StImputacion(data.doc.HeaderRow("IDTipoOperacion"), data.doc.HeaderRow(data.doc.FieldNOperacion), data.doc.ClaseOperacion, False, data.doc.HeaderRow("Fecha"))
        For Each dr As DataRow In data.doc.GetOperacionVinoDestino
            If dr.RowState <> DataRowState.Deleted Then
                datImputa.LineaDestino.Add(New BdgGeneral.StDestino(dr(data.doc.FieldIDLineaOperacionVino), dr("IDArticulo"), dr("IDEstructura") & String.Empty, Nz(dr("Litros"), 0), Nz(dr("Cantidad"), 0)))
            End If
        Next
        Dim dtOperacionVinoMaterial As DataTable = ProcessServer.ExecuteTask(Of BdgGeneral.StImputacion, DataTable)(AddressOf BdgGeneral.ImputacionLineaMateriales, datImputa, services)
        Dim datOrigen As New DataGetOrigenPropuesta(Nothing, dtOperacionVinoMaterial)
        Return datOrigen
    End Function
    <Task()> Public Shared Function GetOrigenOperacionVinoMaterialOFs(ByVal data As DataPropuestaImputacionesVino, ByVal services As ServiceProvider) As DataGetOrigenPropuesta
        Dim dtOperacionVinoMaterial As DataTable
        Dim datOF As DataOrdenFabricacion = services.GetService(Of DataOrdenFabricacion)()
        If Not datOF.Materiales Is Nothing AndAlso datOF.Materiales.Rows.Count > 0 Then
            dtOperacionVinoMaterial = datOF.Materiales.Clone
            For Each drMaterial As DataRow In datOF.Materiales.Rows
                Dim drNew As DataRow = dtOperacionVinoMaterial.NewRow
                For Each col As DataColumn In dtOperacionVinoMaterial.Columns
                    If IsDBNull(drMaterial("IDVino")) Then
                        If datOF.Materiales.Columns.Contains(col.ColumnName) Then
                            drNew(col.ColumnName) = drMaterial(col.ColumnName)
                        End If
                    End If
                Next
                dtOperacionVinoMaterial.Rows.Add(drNew)
            Next
            Dim datOrigen As New DataGetOrigenPropuesta(CType(data.doc.Cabecera, OperCabOFs).NOrden, dtOperacionVinoMaterial)
            Return datOrigen
        End If
    End Function
    <Task()> Public Shared Sub NuevaLineaOperacionVinoMaterial(ByVal data As DataNuevaLineaOperacionImputacion, ByVal services As ServiceProvider)
        Dim OpImp As BusinessHelper
        Dim NOperacion As String = data.doc.FieldNOperacion
        Dim FieldIDMaterialDestino As String = data.doc.FieldIDLineaMaterialLineas
        Dim FieldIDMaterialOrigen As String
        Dim FiledIDLote As String = data.doc.FieldIDLineaMaterialLineasLote
        Dim FieldIDLineaOperacionVino As String = data.doc.FieldIDLineaOperacionVino
        Dim View_Lotes_Origen As String

        Select Case data.doc.Cabecera.TipoOperacion
            Case enumBdgOrigenOperacion.Planificada
                OpImp = New BdgOperacionVinoPlanMaterial
                'FieldIDMaterialOrigen = "IDOperacionPlanMaterial"

            Case enumBdgOrigenOperacion.Real
                OpImp = New BdgVinoMaterial

                Select Case data.doc.Cabecera.Origen
                    Case OrigenOperacion.OperacionPlanificada
                        View_Lotes_Origen = "tbBdgOperacionVinoPlanMaterialLote"
                        FieldIDMaterialOrigen = "IDOperacionVinoPlanMaterial"
                End Select

        End Select

        '//Material
        Dim context As New BusinessData(data.doc.HeaderRow)
        Dim DrDestino As DataRow = data.doc.dtOperacionVinoMaterial.NewRow
        DrDestino(NOperacion) = data.doc.HeaderRow(NOperacion)
        DrDestino(FieldIDLineaOperacionVino) = data.drOperacionVino(FieldIDLineaOperacionVino)
        DrDestino(FieldIDMaterialDestino) = Guid.NewGuid


        If Length(data.drOrigen("IDArticulo")) > 0 Then DrDestino = OpImp.ApplyBusinessRule("IDArticulo", data.drOrigen("IDArticulo"), DrDestino, context)
        If Length(data.drOrigen("IDAlmacen")) > 0 Then DrDestino = OpImp.ApplyBusinessRule("IDAlmacen", data.drOrigen("IDAlmacen"), DrDestino, context)
        DrDestino = OpImp.ApplyBusinessRule("RecalcularMaterial", Nz(data.drOrigen("RecalcularMaterial"), False), DrDestino, context)
        DrDestino = OpImp.ApplyBusinessRule("Cantidad", Nz(data.drOrigen("Cantidad"), 0), DrDestino, context)
        DrDestino = OpImp.ApplyBusinessRule("Merma", Nz(data.drOrigen("Merma"), 0), DrDestino, context)
        'If Not IsDBNull(data.drOrigen("IDVino")) AndAlso Not data.drOrigen("IDVino").Equals(Guid.Empty) Then DrDestino = OpImp.ApplyBusinessRule("IDVino", data.drOrigen("IDVino"), DrDestino, context)
        DrDestino = OpImp.ApplyBusinessRule("Fecha", Nz(data.doc.HeaderRow("Fecha"), Today), DrDestino, context) '????

        If Length(data.drOrigen("Lote")) > 0 Then DrDestino = OpImp.ApplyBusinessRule("Lote", data.drOrigen("Lote"), DrDestino, context)
        If Length(data.drOrigen("Ubicacion")) > 0 Then DrDestino = OpImp.ApplyBusinessRule("Ubicacion", data.drOrigen("Ubicacion"), DrDestino, context)

        data.doc.dtOperacionVinoMaterial.Rows.Add(DrDestino)



        '//Lotes
        If Length(View_Lotes_Origen) > 0 AndAlso Length(FieldIDMaterialOrigen) > 0 Then
            Dim dtLotesOrigen As DataTable = New BE.DataEngine().Filter(View_Lotes_Origen, New GuidFilterItem(FieldIDMaterialOrigen, data.drOrigen(FieldIDMaterialOrigen)))
            If Not dtLotesOrigen Is Nothing AndAlso dtLotesOrigen.Rows.Count > 0 Then
                For Each drLote As DataRow In dtLotesOrigen.Rows
                    Dim drNewLote As DataRow = data.doc.dtOperacionVinoMaterialLote.NewRow
                    For Each col As DataColumn In dtLotesOrigen.Columns
                        If (dtLotesOrigen.Columns.Contains(col.ColumnName) AndAlso _
                            data.doc.dtOperacionMaterialLote.Columns.Contains(col.ColumnName) AndAlso _
                            Length(drLote(col.ColumnName))) Then
                            drNewLote(col.ColumnName) = drLote(col.ColumnName)
                        End If
                    Next

                    drNewLote(FieldIDMaterialDestino) = DrDestino(FieldIDMaterialDestino)
                    drNewLote(FiledIDLote) = Guid.NewGuid
                    If data.doc.dtOperacionVinoMaterialLote.Columns.Contains(NOperacion) Then drNewLote(NOperacion) = data.doc.HeaderRow(NOperacion)
                    drNewLote("IDArticulo") = DrDestino("IDArticulo")
                    drNewLote("IDAlmacen") = DrDestino("IDAlmacen")

                    data.doc.dtOperacionVinoMaterialLote.Rows.Add(drNewLote)
                Next
            End If
        Else
            If Length(DrDestino("IDArticulo")) > 0 AndAlso Length(DrDestino("IDAlmacen")) > 0 AndAlso Length(DrDestino("Lote")) > 0 AndAlso Length(DrDestino("Ubicacion")) > 0 AndAlso Nz(DrDestino("Cantidad"), 0) > 0 Then
                Dim drNewLote As DataRow = data.doc.dtOperacionVinoMaterialLote.NewRow
                drNewLote(FieldIDMaterialDestino) = DrDestino(FieldIDMaterialDestino)
                drNewLote(FiledIDLote) = Guid.NewGuid
                If data.doc.dtOperacionVinoMaterialLote.Columns.Contains(NOperacion) Then drNewLote(NOperacion) = data.doc.HeaderRow(NOperacion)
                drNewLote("IDArticulo") = DrDestino("IDArticulo")
                drNewLote("IDAlmacen") = DrDestino("IDAlmacen")
                drNewLote("Lote") = DrDestino("Lote")
                drNewLote("Ubicacion") = DrDestino("Ubicacion")
                drNewLote("Cantidad") = DrDestino("Cantidad")
                data.doc.dtOperacionVinoMaterialLote.Rows.Add(drNewLote)
            End If
        End If

        DrDestino("NumLotes") = 0

        If data.doc.ClaseOperacion = BdgClaseOperacion.Real Then
            If Not data.doc.dtOperacionVinoMaterialLote Is Nothing AndAlso data.doc.dtOperacionVinoMaterialLote.Rows.Count > 0 Then
                Dim LotesLinea As List(Of DataRow) = (From c In data.doc.dtOperacionVinoMaterialLote Where Not c.IsNull(FieldIDMaterialDestino) AndAlso c(FieldIDMaterialDestino) = DrDestino(FieldIDMaterialDestino) Select c).ToList
                If Not LotesLinea Is Nothing AndAlso LotesLinea.Count > 0 Then
                    DrDestino("NumLotes") = LotesLinea.Count
                    If LotesLinea.Count = 1 Then
                        DrDestino("Lote") = LotesLinea(0)("Lote")
                        DrDestino("Ubicacion") = LotesLinea(0)("Ubicacion")
                    End If
                End If
            End If
        End If
    End Sub


    <Task()> Public Shared Sub PropuestaImputacionesVinoMOD(ByVal data As DataPropuestaImputacionesVino, ByVal services As ServiceProvider)
        If Not data.doc.HeaderRow Is Nothing Then

            If data.doc.Cabecera Is Nothing Then Exit Sub

            Dim datOrigen As DataGetOrigenPropuesta = ProcessServer.ExecuteTask(Of DataPropuestaImputacionesVino, DataGetOrigenPropuesta)(AddressOf GetOrigenOperacionVinoMOD, data, services)
            If Not datOrigen Is Nothing AndAlso Not datOrigen.dtOrigen Is Nothing AndAlso datOrigen.dtOrigen.Rows.Count > 0 Then
                For Each drOrigen As DataRow In datOrigen.dtOrigen.Rows
                    Dim datNewLine As New DataNuevaLineaOperacionImputacion(data.doc, drOrigen, data.IDLineaOperacionVino, data.drOperacionVino)
                    ProcessServer.ExecuteTask(Of DataNuevaLineaOperacionImputacion)(AddressOf NuevaLineaOperacionVinoMOD, datNewLine, services)
                Next
            End If
        End If
    End Sub
    <Task()> Public Shared Function GetOrigenOperacionVinoMOD(ByVal data As DataPropuestaImputacionesVino, ByVal services As ServiceProvider) As DataGetOrigenPropuesta
        Select Case data.doc.Cabecera.Origen
            Case OrigenOperacion.OperacionPlanificada
                Return ProcessServer.ExecuteTask(Of DataPropuestaImputacionesVino, DataGetOrigenPropuesta)(AddressOf GetOrigenOperacionVinoMODOpPlanificada, data, services)
            Case Else
                Return ProcessServer.ExecuteTask(Of DataPropuestaImputacionesVino, DataGetOrigenPropuesta)(AddressOf GetOrigenOperacionVinoMODPorTipoOp, data, services)
        End Select
    End Function
    <Task()> Public Shared Function GetOrigenOperacionVinoMODOpPlanificada(ByVal data As DataPropuestaImputacionesVino, ByVal services As ServiceProvider) As DataGetOrigenPropuesta
        Dim NOperacionPlan As String = CType(data.doc.Cabecera, OperCabPlanificadas).IDOrigen

        Dim f As New Filter
        f.Add(New FilterItem("NOperacionPlan", NOperacionPlan))
        f.Add(New FilterItem("IDLineaOperacionVinoPlan", data.IDLineaOperacionVinoOrigen))
        Dim dtOperacionVinoMOD As DataTable = New BE.DataEngine().Filter("frmBdgOperacionVinoPlanMOD", f)
        Dim datOrigen As New DataGetOrigenPropuesta(NOperacionPlan, dtOperacionVinoMOD)
        Return datOrigen
    End Function
    <Task()> Public Shared Function GetOrigenOperacionVinoMODPorTipoOp(ByVal data As DataPropuestaImputacionesVino, ByVal services As ServiceProvider) As DataGetOrigenPropuesta
        Dim datImputa As New BdgGeneral.StImputacion(data.doc.HeaderRow("IDTipoOperacion"), data.doc.HeaderRow(data.doc.FieldNOperacion), data.doc.ClaseOperacion, False, data.doc.HeaderRow("Fecha"))
        For Each dr As DataRow In data.doc.GetOperacionVinoDestino
            If dr.RowState <> DataRowState.Deleted Then
                datImputa.LineaDestino.Add(New BdgGeneral.StDestino(dr(data.doc.FieldIDLineaOperacionVino), dr("IDArticulo"), dr("IDEstructura") & String.Empty, Nz(dr("Litros"), 0), Nz(dr("Cantidad"), 0)))
            End If
        Next
        Dim dtOperacionVinoMOD As DataTable = ProcessServer.ExecuteTask(Of BdgGeneral.StImputacion, DataTable)(AddressOf BdgGeneral.ImputacionLineaMOD, datImputa, services)
        Dim datOrigen As New DataGetOrigenPropuesta(Nothing, dtOperacionVinoMOD)
        Return datOrigen
    End Function

    <Task()> Public Shared Sub NuevaLineaOperacionVinoMOD(ByVal data As DataNuevaLineaOperacionImputacion, ByVal services As ServiceProvider)
        Dim OpImp As BusinessHelper
        Dim NOperacion As String = data.doc.FieldNOperacion
        Dim FieldIDMODDestino As String = data.doc.FieldIDLineaMODLineas
        Dim FieldIDMODOrigen As String
        Dim FieldIDLineaOperacionVino As String = data.doc.FieldIDLineaOperacionVino

        Select Case data.doc.Cabecera.TipoOperacion
            Case enumBdgOrigenOperacion.Planificada
                OpImp = New BdgOperacionPlanMOD
            Case enumBdgOrigenOperacion.Real
                OpImp = New BdgOperacionMOD

                Select Case data.doc.Cabecera.Origen
                    Case OrigenOperacion.OperacionPlanificada
                        FieldIDMODOrigen = "IDOperacionVinoPlanMOD"
                End Select
        End Select


        Dim context As New BusinessData(data.doc.HeaderRow)
        Dim DrDestino As DataRow = data.doc.dtOperacionVinoMOD.NewRow
        DrDestino(NOperacion) = data.doc.HeaderRow(NOperacion)
        DrDestino(FieldIDLineaOperacionVino) = data.drOperacionVino(FieldIDLineaOperacionVino)
        DrDestino(FieldIDMODDestino) = Guid.NewGuid

        If Length(data.drOrigen("IDOperario")) > 0 Then DrDestino = OpImp.ApplyBusinessRule("IDOperario", data.drOrigen("IDOperario"), DrDestino, context)
        DrDestino = OpImp.ApplyBusinessRule("Fecha", Nz(data.doc.HeaderRow("Fecha"), Now), DrDestino, context)
        If Length(data.drOrigen("IDCategoria")) > 0 Then DrDestino = OpImp.ApplyBusinessRule("IDCategoria", data.drOrigen("IDCategoria"), DrDestino, context)
        If Length(data.drOrigen("Tiempo")) > 0 Then DrDestino = OpImp.ApplyBusinessRule("Tiempo", data.drOrigen("Tiempo"), DrDestino, context)
        data.doc.dtOperacionVinoMOD.Rows.Add(DrDestino)
    End Sub


    <Task()> Public Shared Sub PropuestaImputacionesVinoCentro(ByVal data As DataPropuestaImputacionesVino, ByVal services As ServiceProvider)
        If Not data.doc.HeaderRow Is Nothing Then

            If data.doc.Cabecera Is Nothing Then Exit Sub
            Dim datOrigen As DataGetOrigenPropuesta = ProcessServer.ExecuteTask(Of DataPropuestaImputacionesVino, DataGetOrigenPropuesta)(AddressOf GetOrigenOperacionVinoCentro, data, services)
            If Not datOrigen Is Nothing AndAlso Not datOrigen.dtOrigen Is Nothing AndAlso datOrigen.dtOrigen.Rows.Count > 0 Then
                For Each drOrigen As DataRow In datOrigen.dtOrigen.Rows
                    Dim datNewLine As New DataNuevaLineaOperacionImputacion(data.doc, drOrigen, data.IDLineaOperacionVino, data.drOperacionVino)
                    ProcessServer.ExecuteTask(Of DataNuevaLineaOperacionImputacion)(AddressOf NuevaLineaOperacionVinoCentro, datNewLine, services)
                Next
            End If
        End If
    End Sub
    <Task()> Public Shared Function GetOrigenOperacionVinoCentro(ByVal data As DataPropuestaImputacionesVino, ByVal services As ServiceProvider) As DataGetOrigenPropuesta
        Select Case data.doc.Cabecera.Origen
            Case OrigenOperacion.OperacionPlanificada
                Return ProcessServer.ExecuteTask(Of DataPropuestaImputacionesVino, DataGetOrigenPropuesta)(AddressOf GetOrigenOperacionVinoCentroOpPlanificada, data, services)
            Case Else
                Return ProcessServer.ExecuteTask(Of DataPropuestaImputacionesVino, DataGetOrigenPropuesta)(AddressOf GetOrigenOperacionVinoCentroPorTipoOp, data, services)
        End Select
    End Function
    <Task()> Public Shared Function GetOrigenOperacionVinoCentroOpPlanificada(ByVal data As DataPropuestaImputacionesVino, ByVal services As ServiceProvider) As DataGetOrigenPropuesta
        Dim NOperacionPlan As String = CType(data.doc.Cabecera, OperCabPlanificadas).IDOrigen
        Dim f As New Filter
        f.Add(New FilterItem("NOperacionPlan", NOperacionPlan))
        f.Add(New FilterItem("IDLineaOperacionVinoPlan", data.IDLineaOperacionVinoOrigen))

        Dim dtOperacionVinoCentro As DataTable = New BE.DataEngine().Filter("frmBdgOperacionVinoPlanCentro", f)
        Dim datOrigen As New DataGetOrigenPropuesta(NOperacionPlan, dtOperacionVinoCentro)
        Return datOrigen
    End Function
    <Task()> Public Shared Function GetOrigenOperacionVinoCentroPorTipoOp(ByVal data As DataPropuestaImputacionesVino, ByVal services As ServiceProvider) As DataGetOrigenPropuesta
        Dim datImputa As New BdgGeneral.StImputacion(data.doc.HeaderRow("IDTipoOperacion"), data.doc.HeaderRow(data.doc.FieldNOperacion), data.doc.ClaseOperacion, False, data.doc.HeaderRow("Fecha"))
        For Each dr As DataRow In data.doc.GetOperacionVinoDestino
            If dr.RowState <> DataRowState.Deleted Then
                datImputa.LineaDestino.Add(New BdgGeneral.StDestino(dr(data.doc.FieldIDLineaOperacionVino), dr("IDArticulo"), dr("IDEstructura") & String.Empty, Nz(dr("Litros"), 0), Nz(dr("Cantidad"), 0)))
            End If
        Next
        Dim dtOperacionVinoCentro As DataTable = ProcessServer.ExecuteTask(Of BdgGeneral.StImputacion, DataTable)(AddressOf BdgGeneral.ImputacionLineaCentros, datImputa, services)
        Dim datOrigen As New DataGetOrigenPropuesta(Nothing, dtOperacionVinoCentro)
        Return datOrigen
    End Function
    <Task()> Public Shared Sub NuevaLineaOperacionVinoCentro(ByVal data As DataNuevaLineaOperacionImputacion, ByVal services As ServiceProvider)
        Dim OpImp As BusinessHelper
        Dim NOperacion As String = data.doc.FieldNOperacion
        Dim FieldIDCentroDestino As String = data.doc.FieldIDLineaCentroLineas
        Dim FieldIDCentroOrigen As String
        Dim FieldIDLineaOperacionVino As String = data.doc.FieldIDLineaOperacionVino

        Select Case data.doc.Cabecera.TipoOperacion
            Case enumBdgOrigenOperacion.Planificada
                OpImp = New BdgOperacionVinoPlanCentro
            Case enumBdgOrigenOperacion.Real
                OpImp = New BdgVinoCentro

                Select Case data.doc.Cabecera.Origen
                    Case OrigenOperacion.OperacionPlanificada
                        FieldIDCentroOrigen = "IDOperacionVinoPlanCentro"
                End Select
        End Select


        Dim context As New BusinessData(data.doc.HeaderRow)
        Dim DrDestino As DataRow = data.doc.dtOperacionVinoCentro.NewRow
        DrDestino(NOperacion) = data.doc.HeaderRow(NOperacion)
        DrDestino(FieldIDLineaOperacionVino) = data.drOperacionVino(FieldIDLineaOperacionVino)
        DrDestino(FieldIDCentroDestino) = Guid.NewGuid

        If Length(data.drOrigen("IDCentro")) > 0 Then DrDestino = OpImp.ApplyBusinessRule("IDCentro", data.drOrigen("IDCentro"), DrDestino, context)
        DrDestino = OpImp.ApplyBusinessRule("Fecha", Nz(data.doc.HeaderRow("Fecha"), Now), DrDestino, context)
        If Length(data.drOrigen("UDTiempo")) > 0 Then DrDestino = OpImp.ApplyBusinessRule("UDTiempo", data.drOrigen("UDTiempo"), DrDestino, context)
        DrDestino = OpImp.ApplyBusinessRule("Tiempo", Nz(data.drOrigen("Tiempo"), 0), DrDestino, context)
        If data.drOrigen("PorCantidad") Then DrDestino = OpImp.ApplyBusinessRule("PorCantidad", Nz(data.drOrigen("PorCantidad"), False), DrDestino, context)
        If Nz(DrDestino("PorCantidad"), False) Then
            DrDestino = OpImp.ApplyBusinessRule("Cantidad", Nz(data.drOperacionVino("Cantidad")), DrDestino, context)
        End If
        If Length(data.drOrigen("IDIncidencia")) > 0 Then DrDestino = OpImp.ApplyBusinessRule("IDIncidencia", data.drOrigen("IDIncidencia"), DrDestino, context)
        data.doc.dtOperacionVinoCentro.Rows.Add(DrDestino)
    End Sub


    <Task()> Public Shared Sub PropuestaImputacionesVinoVarios(ByVal data As DataPropuestaImputacionesVino, ByVal services As ServiceProvider)
        If Not data.doc.HeaderRow Is Nothing Then

            If data.doc.Cabecera Is Nothing Then Exit Sub

            Dim datOrigen As DataGetOrigenPropuesta = ProcessServer.ExecuteTask(Of DataPropuestaImputacionesVino, DataGetOrigenPropuesta)(AddressOf GetOrigenOperacionVinoVarios, data, services)
            If Not datOrigen Is Nothing AndAlso Not datOrigen.dtOrigen Is Nothing AndAlso datOrigen.dtOrigen.Rows.Count > 0 Then
                For Each drOrigen As DataRow In datOrigen.dtOrigen.Rows

                    Dim datNewLine As New DataNuevaLineaOperacionImputacion(data.doc, drOrigen, data.IDLineaOperacionVinoDestino, data.drOperacionVino) 'data.IDLineaDestino = IDvino
                    ProcessServer.ExecuteTask(Of DataNuevaLineaOperacionImputacion)(AddressOf NuevaLineaOperacionVinoVarios, datNewLine, services)
                Next
            End If
        End If
    End Sub
    <Task()> Public Shared Function GetOrigenOperacionVinoVarios(ByVal data As DataPropuestaImputacionesVino, ByVal services As ServiceProvider) As DataGetOrigenPropuesta
        Select Case data.doc.Cabecera.Origen
            Case OrigenOperacion.OperacionPlanificada
                Return ProcessServer.ExecuteTask(Of DataPropuestaImputacionesVino, DataGetOrigenPropuesta)(AddressOf GetOrigenOperacionVinoVariosOpPlanificada, data, services)
            Case Else
                Return ProcessServer.ExecuteTask(Of DataPropuestaImputacionesVino, DataGetOrigenPropuesta)(AddressOf GetOrigenOperacionVinoVariosPorTipoOp, data, services)
        End Select
    End Function
    <Task()> Public Shared Function GetOrigenOperacionVinoVariosOpPlanificada(ByVal data As DataPropuestaImputacionesVino, ByVal services As ServiceProvider) As DataGetOrigenPropuesta
        Dim NOperacionPlan As String = CType(data.doc.Cabecera, OperCabPlanificadas).IDOrigen
        Dim f As New Filter
        f.Add(New FilterItem("NOperacionPlan", NOperacionPlan))
        f.Add(New FilterItem("IDLineaOperacionVinoPlan", data.IDLineaOperacionVinoOrigen))

        Dim dtOperacionVinoVarios As DataTable = New BE.DataEngine().Filter("frmBdgOperacionVinoPlanVarios", f)
        Dim datOrigen As New DataGetOrigenPropuesta(NOperacionPlan, dtOperacionVinoVarios)
        datOrigen.IDLineaOrigen = data.IDLineaOperacionVinoOrigen
        Return datOrigen
    End Function
    <Task()> Public Shared Function GetOrigenOperacionVinoVariosPorTipoOp(ByVal data As DataPropuestaImputacionesVino, ByVal services As ServiceProvider) As DataGetOrigenPropuesta
        Dim datImputa As New BdgGeneral.StImputacion(data.doc.HeaderRow("IDTipoOperacion"), data.doc.HeaderRow(data.doc.FieldNOperacion), data.doc.ClaseOperacion, False, data.doc.HeaderRow("Fecha"))
        For Each dr As DataRow In data.doc.GetOperacionVinoDestino
            If dr.RowState <> DataRowState.Deleted Then
                datImputa.LineaDestino.Add(New BdgGeneral.StDestino(dr(data.doc.FieldIDLineaOperacionVino), dr("IDArticulo"), dr("IDEstructura") & String.Empty, Nz(dr("Litros"), 0), Nz(dr("Cantidad"), 0)))
            End If
        Next
        Dim dtOperacionVinoVarios As DataTable = ProcessServer.ExecuteTask(Of BdgGeneral.StImputacion, DataTable)(AddressOf BdgGeneral.ImputacionLineaVarios, datImputa, services)
        Dim datOrigen As New DataGetOrigenPropuesta(Nothing, dtOperacionVinoVarios)
        Return datOrigen
    End Function
    <Task()> Public Shared Sub NuevaLineaOperacionVinoVarios(ByVal data As DataNuevaLineaOperacionImputacion, ByVal services As ServiceProvider)
        Dim OpImp As BusinessHelper
        Dim NOperacion As String = data.doc.FieldNOperacion
        Dim FieldIDVariosDestino As String = data.doc.FieldIDLineaVariosLineas
        Dim FieldIDVariosOrigen As String
        Dim FieldIDLineaOperacionVino As String = data.doc.FieldIDLineaOperacionVino
        Select Case data.doc.Cabecera.TipoOperacion
            Case enumBdgOrigenOperacion.Planificada
                OpImp = New BdgOperacionVinoPlanVarios
            Case enumBdgOrigenOperacion.Real
                OpImp = New BdgVinoVarios

                Select Case data.doc.Cabecera.Origen
                    Case OrigenOperacion.OperacionPlanificada
                        FieldIDVariosOrigen = "IDOperacionVinoPlanVarios"
                End Select
        End Select

        Dim context As New BusinessData(data.doc.HeaderRow)
        Dim DrDestino As DataRow = data.doc.dtOperacionVinoVarios.NewRow
        DrDestino(NOperacion) = data.doc.HeaderRow(NOperacion)
        DrDestino(FieldIDLineaOperacionVino) = data.drOperacionVino(FieldIDLineaOperacionVino)
        DrDestino(FieldIDVariosDestino) = Guid.NewGuid

        If Length(data.drOrigen("IDVarios")) > 0 Then DrDestino = OpImp.ApplyBusinessRule("IDVarios", data.drOrigen("IDVarios"), DrDestino, context)
        DrDestino = OpImp.ApplyBusinessRule("Fecha", Nz(data.doc.HeaderRow("Fecha"), Now), DrDestino, context)
        DrDestino = OpImp.ApplyBusinessRule("Cantidad", Nz(data.drOrigen("Cantidad"), 0), DrDestino, context)
        DrDestino = OpImp.ApplyBusinessRule("Tasa", Nz(data.drOrigen("Tasa"), 0), DrDestino, context)
        data.doc.dtOperacionVinoVarios.Rows.Add(DrDestino)
    End Sub

    <Task()> Public Shared Sub PropuestaImputacionesVinoAnalisis(ByVal data As DataPropuestaImputacionesVino, ByVal services As ServiceProvider)
        If Not data.doc.HeaderRow Is Nothing AndAlso data.doc.ClaseOperacion = BdgClaseOperacion.Real Then
            If Length(data.doc.HeaderRow("IDAnalisis")) > 0 Then
                Dim datImp As New BdgGeneral.StImputacion
                datImp.IDAnalisis = data.doc.HeaderRow("IDAnalisis")
                Dim dtVariables As DataTable = ProcessServer.ExecuteTask(Of BdgGeneral.StImputacion, DataTable)(AddressOf BdgGeneral.GetImputacionOperacionVinoAnalisis, datImp, services)
                For Each drDestino As DataRow In data.doc.GetOperacionVinoDestino
                    Dim datAnalisis As New BdgGeneral.DataAddNuevaLineaImputacionAnalisis(data.doc.HeaderRow("NOperacion"), data.doc.HeaderRow("IDAnalisis"), data.doc.HeaderRow("Fecha"), drDestino("IDVino"), CType(data.doc, DocumentoBdgOperacionReal).dtOperacionVinoAnalisis, CType(data.doc, DocumentoBdgOperacionReal).dtOperacionVinoAnalisisVariable, dtVariables)
                    ProcessServer.ExecuteTask(Of BdgGeneral.DataAddNuevaLineaImputacionAnalisis)(AddressOf BdgGeneral.AddNuevaLineaImputacionAnalisis, datAnalisis, services)
                Next
            End If
        End If
    End Sub

#End Region

#Region " Validar y Guardar la Propuesta "

    '<Task()> Public Shared Sub ValidarPropuesta(ByVal doc As DocumentoBdgOperacion, ByVal services As ServiceProvider)
    '    Dim ProcInfo As ProcessInfoOperacion = services.GetService(Of ProcessInfoOperacion)()
    '    If ProcInfo.GuardarPropuesta Then
    '        Select Case doc.Cabecera.TipoOperacion
    '            Case enumBdgOrigenOperacion.Planificada

    '            Case enumBdgOrigenOperacion.Real
    '                Dim TiposOperacion As EntityInfoCache(Of BdgTipoOperacionInfo) = services.GetService(Of EntityInfoCache(Of BdgTipoOperacionInfo))()
    '                Dim TipoOpInfo As BdgTipoOperacionInfo = TiposOperacion.GetEntity(doc.HeaderRow("IDTipoOperacion"))

    '                Dim StVal As New BdgOperacion.StValidar(TipoOpInfo.TipoMovimiento, doc.GetOperacionVinoOrigen.CopyToDataTable, doc.GetOperacionVinoDestino.CopyToDataTable, doc.HeaderRow.Table)
    '                Dim oColFrx As Hashtable = ProcessServer.ExecuteTask(Of BdgOperacion.StValidar, Hashtable)(AddressOf BdgOperacion.ValidarOperacion, StVal, services)
    '        End Select
    '    End If
    'End Sub

    <Task()> Public Shared Sub GuardarPropuesta(ByVal doc As DocumentoBdgOperacion, ByVal services As ServiceProvider)
        Dim ProcInfo As ProcessInfoOperacion = services.GetService(Of ProcessInfoOperacion)()
        If ProcInfo.GuardarPropuesta Then
            AdminData.BeginTx()
            Dim Op As Object
            Select Case doc.Cabecera.TipoOperacion
                Case enumBdgOrigenOperacion.Planificada
                    Op = New BdgOperacionPlan
                Case enumBdgOrigenOperacion.Real
                    Op = New BdgOperacion
            End Select

            Dim pck As New UpdatePackage
            For Each dt As DataTable In doc.GetTables
                pck.Add(dt)
            Next
            Op.Update(pck, services)
            AdminData.CommitTx(True)
        End If
    End Sub


#End Region


End Class
