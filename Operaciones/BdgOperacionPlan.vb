Public Class BdgOperacionPlan

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgOperacionPlan"

#End Region

#Region "Eventos Entidad"

#Region "Tareas RegisterAddNewTasks"

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        data("Estado") = BdgEstadoOperacionPlan.Planificado

        ProcessServer.ExecuteTask(Of DataRow)(AddressOf NegocioGeneral.AsignarCentroGestion, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarContadorPredeterminado, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarNumeroOperacionProvisional, data, services)

        'Dim dataCont As New Contador.DatosDefaultCounterValue(data, "BdgOperacionPlan", "NOperacionPlan")
        'ProcessServer.ExecuteTask(Of Contador.DatosDefaultCounterValue)(AddressOf Contador.LoadDefaultCounterValue, dataCont, services)

        'Dim datFecha As New BdgGeneral.DataAsignarFechaPropuestaOperacion(GetType(BdgOperacionPlan).Name, data)
        'ProcessServer.ExecuteTask(Of BdgGeneral.DataAsignarFechaPropuestaOperacion)(AddressOf BdgGeneral.AsignarFechaPropuestaOperacion, datFecha, services)

        Dim datFecha As New BdgGeneral.DataGetFechaPropuestaOperacion(GetType(BdgOperacionPlan).Name)
        data("Fecha") = ProcessServer.ExecuteTask(Of BdgGeneral.DataGetFechaPropuestaOperacion, Date)(AddressOf BdgGeneral.GetFechaPropuestaOperacion, datFecha, services)

        data("ImputacionGlobalMat") = False : data("ImputacionGlobalMod") = False
        data("ImputacionGlobalCentro") = False : data("ImputacionGlobalVarios") = False
    End Sub

    <Task()> Public Shared Sub AsignarContadorPredeterminado(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim CE As New CentroEntidad
        CE.CentroGestion = data("IDCentroGestion") & String.Empty
        CE.ContadorEntidad = CentroGestion.ContadorEntidad.BdgOperacionPlan
        data("IDContador") = ProcessServer.ExecuteTask(Of CentroEntidad, String)(AddressOf CentroGestion.GetContadorPredeterminado, CE, services)
    End Sub


    <Task()> Public Shared Sub AsignarNumeroOperacionProvisional(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDContador")) > 0 Then
            Dim dtContadores As DataTable = ProcessServer.ExecuteTask(Of String, DataTable)(AddressOf Contador.CounterDt, GetType(BdgOperacionPlan).Name, services)
            Dim adr As DataRow() = dtContadores.Select("IDContador = " & Quoted(data("IDContador")))
            If Not IsNothing(adr) AndAlso adr.Length > 0 Then
                data("NOperacionPlan") = adr(0)("ValorProvisional")
            Else
                Dim dtContadorPred As DataTable = ProcessServer.ExecuteTask(Of String, DataTable)(AddressOf Contador.CounterDefault, GetType(BdgOperacionPlan).Name, services)
                If Not dtContadorPred Is Nothing AndAlso dtContadorPred.Rows.Count > 0 Then
                    data("IDContador") = dtContadorPred.Rows(0)("IDContador")
                    adr = dtContadores.Select("IDContador = " & Quoted(data("IDContador")))
                    If Not IsNothing(adr) AndAlso adr.Length > 0 Then
                        data("NOperacionPlan") = adr(0)("ValorProvisional")
                    End If
                End If
            End If
        End If
    End Sub

#End Region

#Region "Tareas RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarCambioEstado)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDTipoOperacion")) = 0 Then ApplicationService.GenerateError("El Tipo de Operación es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarCambioEstado(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Modified Then
            If Nz(data("Estado"), -1) <> Nz(data("Estado", DataRowVersion.Original), -1) AndAlso Nz(data("Estado", DataRowVersion.Original), -1) = BdgEstadoOperacionPlan.Confirmado Then
                ApplicationService.GenerateError("No se puede cambiar el Estado. Para desconfirmar, debe eliminar la Operación de Bodega correspondiente.")
            End If

            If Nz(data("Estado"), -1) <> Nz(data("Estado", DataRowVersion.Original), -1) AndAlso Nz(data("Estado"), -1) = BdgEstadoOperacionPlan.Confirmado Then
                ApplicationService.GenerateError("No se puede cambiar el Estado. Para confirmar, debe generar la Operación de Bodega correspondiente.")
            End If
        End If
    End Sub

#End Region

#Region "Tareas RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of UpdatePackage, DocumentoBdgOperacionPlan)(AddressOf ProcesoBdgOperacionPlan.CrearDocumento)
        updateProcess.AddTask(Of DocumentoBdgOperacionPlan)(AddressOf ProcesoBdgOperacionPlan.AsignarContadorCabecera)
        updateProcess.AddTask(Of DocumentoBdgOperacionPlan)(AddressOf ProcesoBdgOperacionPlan.ImputacionesGlobales)
        updateProcess.AddTask(Of DocumentoBdgOperacionPlan)(AddressOf ProcesoBdgOperacionPlan.VinoPlanOrigenDestino)
        updateProcess.AddTask(Of DocumentoBdgOperacionPlan)(AddressOf Comunes.UpdateDocument)
        updateProcess.AddTask(Of DocumentoBdgOperacionPlan)(AddressOf Comunes.MarcarComoActualizado)
    End Sub

#End Region

#Region "Tareas GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim Obrl As New BusinessRules
        Obrl.Add("IDTipoOperacion", AddressOf CambioIDTipoOperacion)
        Return Obrl
    End Function

    <Task()> Public Shared Sub CambioIDTipoOperacion(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If data.Context Is Nothing Then data.Context = New BusinessData
        data.Context("Origen") = enumBdgOrigenOperacion.Planificada
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf BdgGeneral.CambioIDTipoOperacion, data, services)
    End Sub

#End Region

#Region "Tareas RegisterDeleteTasks"

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf ValidarEstadoOperacion)
    End Sub

    <Task()> Public Shared Sub ValidarEstadoOperacion(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data("Estado") = BdgEstadoOperacionPlan.Confirmado Then
            ApplicationService.GenerateError("No se puede borrar la operación planificada, está confirmada.")
        End If
    End Sub

#End Region

#End Region

#Region "Proceso Generación de Operaciones Planificadas"

#Region "Tareas - Generación de Propuestas"

    <Serializable()> _
    Public Class StPropuestaConfirmacionPlan
        Public NOperacionPlan As String
        Public OperacionCabecera As DataTable
        Public OperacionVinoOrigen As DataTable
        Public OperacionVinoDestino As DataTable
        Public ImputacionGlobalMateriales As DataTable
        Public ImputacionGlobalMaterialesLotes As DataTable
        Public ImputacionGlobalCentros As DataTable
        Public ImputacionGlobalMOD As DataTable
        Public ImputacionGlobalVarios As DataTable
        Public ImputacionLineaMaterial As DataTable
        Public ImputacionLineaMaterialLotes As DataTable
        Public ImputacionLineaCentros As DataTable
        Public ImputacionLineaMOD As DataTable
        Public ImputacionLineaVarios As DataTable
        Public OperacionAnalitica As DataTable

        Public ValidarVinosDepositoDestino As Boolean
        Public AutoAsignarVino As Boolean

        Public Sub New(ByVal nOperacionPlanOrigen As String, Optional ByVal validarVinoDestino As Boolean = True, Optional ByVal autoAsignarVino As Boolean = True)
            Me.NOperacionPlan = nOperacionPlanOrigen
            Me.ValidarVinosDepositoDestino = validarVinoDestino
            Me.AutoAsignarVino = autoAsignarVino
        End Sub

    End Class

    <Task()> Public Shared Function PropuestaConfirmarPlan(ByVal data As StPropuestaConfirmacionPlan, ByVal services As ServiceProvider) As StPropuestaConfirmacionPlan
        Dim BlnPropuesta As Boolean = ProcessServer.ExecuteTask(Of StPropuestaConfirmacionPlan, Boolean)(AddressOf ValidarPropuestaPlan, data, services)
        If BlnPropuesta Then
            data = ProcessServer.ExecuteTask(Of StPropuestaConfirmacionPlan, StPropuestaConfirmacionPlan)(AddressOf PropuestaConfirmacionCabeceraOperacion, data, services)
            Dim DtOperCabecera As DataTable = data.OperacionCabecera
            If Not DtOperCabecera Is Nothing AndAlso DtOperCabecera.Rows.Count > 0 Then
                data = ProcessServer.ExecuteTask(Of StPropuestaConfirmacionPlan, StPropuestaConfirmacionPlan)(AddressOf PropuestaConfirmacionLineasOperacion, data, services)
                data = ProcessServer.ExecuteTask(Of StPropuestaConfirmacionPlan, StPropuestaConfirmacionPlan)(AddressOf PropuestaConfirmacionImputaciones, data, services)
                data = ProcessServer.ExecuteTask(Of StPropuestaConfirmacionPlan, StPropuestaConfirmacionPlan)(AddressOf PropuestaConfirmacionAnalisisOperacion, data, services)
            Else : ApplicationService.GenerateError("Ha sucedido un eror generando la cabecera de la operación.")
            End If
        Else : ApplicationService.GenerateError("No se ha podido confirmar la operación.")
        End If
        Return data
    End Function

    <Task()> Public Shared Function ValidarPropuestaPlan(ByVal data As StPropuestaConfirmacionPlan, ByVal services As ServiceProvider) As Boolean
        'HAY QUE VALIDAR:
        '   - Actualizar  ocupaciones
        '   - Comprobar que no existan cantidades a cero
        '   - Comprobar incoherencias cantidades (en función del tipo mov, no debería haber diferencias => mermas)
        '   - Validar barricas en aquellas operaciones que impliquen mov, y que piden la barrica como lote
        '   - Validar Vino en Depósito
        Return True
    End Function

    <Task()> Public Shared Function PropuestaConfirmacionCabeceraOperacion(ByVal data As StPropuestaConfirmacionPlan, ByVal services As ServiceProvider) As StPropuestaConfirmacionPlan
        Dim DtOperPlan As DataTable = New BdgOperacionPlan().SelOnPrimaryKey(data.NOperacionPlan)
        data.OperacionCabecera = New BdgOperacion().AddNewForm
        data.OperacionCabecera.Rows(0)("NOperacionPlan") = data.NOperacionPlan
        data.OperacionCabecera.Rows(0)("IDAnalisis") = DtOperPlan.Rows(0)("IDAnalisis")
        data.OperacionCabecera.Rows(0)("IDTipoOperacion") = DtOperPlan.Rows(0)("IDTipoOperacion")
        data.OperacionCabecera.Rows(0)("Fecha") = DtOperPlan.Rows(0)("Fecha")
        data.OperacionCabecera.Rows(0)("ImputacionRealCentro") = DtOperPlan.Rows(0)("ImputacionGlobalCentro")
        data.OperacionCabecera.Rows(0)("ImputacionRealMaterial") = DtOperPlan.Rows(0)("ImputacionGlobalMat")
        data.OperacionCabecera.Rows(0)("ImputacionRealMod") = DtOperPlan.Rows(0)("ImputacionGlobalMod")
        data.OperacionCabecera.Rows(0)("ImputacionRealVarios") = DtOperPlan.Rows(0)("ImputacionGlobalVarios")
        data.OperacionCabecera.Rows(0)("Texto") = DtOperPlan.Rows(0)("Texto")
        Return data
    End Function

    <Task()> Public Shared Function PropuestaConfirmacionLineasOperacion(ByVal data As StPropuestaConfirmacionPlan, ByVal services As ServiceProvider) As StPropuestaConfirmacionPlan
        If Not data.OperacionCabecera Is Nothing AndAlso data.OperacionCabecera.Rows.Count > 0 Then
            Dim ClsBE As New DataEngine
            data.OperacionVinoOrigen = ClsBE.Filter("frmBdgOperacionVinoOrigenPlan", New FilterItem("NOperacionPlan", data.OperacionCabecera.Rows(0)("NOperacionPlan")))
            data.OperacionVinoDestino = ClsBE.Filter("frmBdgOperacionVinoDestinoPlan", New FilterItem("NOperacionPlan", data.OperacionCabecera.Rows(0)("NOperacionPlan")))

            Dim dtOV As DataTable = ClsBE.Filter("frmBdgOperacionVinoOrigen", New NoRowsFilterItem)
            Dim Stdata As New StObtenerLineasPropuesta(data.OperacionCabecera, data.OperacionVinoOrigen, dtOV, data.AutoAsignarVino)
            Stdata = ProcessServer.ExecuteTask(Of StObtenerLineasPropuesta, StObtenerLineasPropuesta)(AddressOf ObtenerLineasPropuesta, Stdata, services)
            data.OperacionVinoOrigen = Stdata.Destino.Copy()

            dtOV = ClsBE.Filter("frmBdgOperacionVinoDestino", New NoRowsFilterItem)
            Stdata = New StObtenerLineasPropuesta(data.OperacionCabecera, data.OperacionVinoDestino, dtOV, data.AutoAsignarVino)
            Stdata = ProcessServer.ExecuteTask(Of StObtenerLineasPropuesta, StObtenerLineasPropuesta)(AddressOf ObtenerLineasPropuesta, Stdata, services)
            data.OperacionVinoDestino = Stdata.Destino.Copy()
        End If
        Return data
    End Function

    <Task()> Public Shared Function PropuestaConfirmacionImputaciones(ByVal data As StPropuestaConfirmacionPlan, ByVal services As ServiceProvider) As StPropuestaConfirmacionPlan
        If Not data.OperacionCabecera Is Nothing AndAlso data.OperacionCabecera.Rows.Count > 0 Then
            Dim dtrCabeceraOrigen As DataRow = New BdgOperacionPlan().GetItemRow(data.NOperacionPlan)
            Dim FilNOperPlan As New Filter()
            FilNOperPlan.Add(New StringFilterItem("NOperacionPlan", dtrCabeceraOrigen("NOperacionPlan")))
            Dim dttLineas As DataTable = New BdgOperacionVinoPlan().Filter(FilNOperPlan)
            Dim filterLineas As New Filter(FilterUnionOperator.Or)
            For Each dtrLineaDestinoOrigen As DataRow In dttLineas.Select
                filterLineas.Add(New GuidFilterItem("IDLineaOperacionVinoPlan", dtrLineaDestinoOrigen("IDLineaOperacionVinoPlan")))
            Next

            Dim ClsBE As New DataEngine
            'MATERIALES
            'Dim dttMatGlobalDestinos As DataTable = ClsBE.Filter("frmBdgOperacionMaterialRealGlobal", New NoRowsFilterItem)
            'Dim dttMatLineaDestinos As DataTable = ClsBE.Filter("frmBdgOperacionMaterial", New NoRowsFilterItem)
            'Dim dttMatGlobalOrigenes As DataTable = ClsBE.Filter("frmBdgOperacionPlanMaterialGlobal", FilNOperPlan)
            'Dim dttMatLineaOrigenes As DataTable = ClsBE.Filter("frmBdgOperacionVinoPlanMaterial", FilNOperPlan)
            'Dim stImputacionesMatInfo As New StPropuestaImputacion(data.OperacionCabecera, data.OperacionVinoDestino, dtrCabeceraOrigen("ImputacionGlobalMat"), _
            '                                                            dttMatGlobalOrigenes, dttMatGlobalDestinos, dttMatLineaOrigenes, dttMatLineaDestinos, "IDOperacionMaterial")
            'If Not Nz(dtrCabeceraOrigen("ImputacionGlobalMat"), False) Then
            '    stImputacionesMatInfo.PKField = "IDVinoMaterial"
            'End If

            'stImputacionesMatInfo = ProcessServer.ExecuteTask(Of DataPropuestaImputacionMateriales, DataPropuestaImputacionMateriales)(AddressOf PropuestaImputacionMateriales, stImputacionesMatInfo, services)
            'data.ImputacionGlobalMateriales = stImputacionesMatInfo.TablaImputacionGlobalDestino
            'data.ImputacionLineaMaterial = stImputacionesMatInfo.TablaImputacionLineaDestino
            Dim datMatGlobal As New DataPropuestaImputacionMateriales(data.OperacionCabecera(0)("NOperacion"), dtrCabeceraOrigen("NOperacionPlan"), True)
            datMatGlobal = ProcessServer.ExecuteTask(Of DataPropuestaImputacionMateriales, DataPropuestaImputacionMateriales)(AddressOf PropuestaImputacionMateriales, datMatGlobal, services)
            data.ImputacionGlobalMateriales = datMatGlobal.MaterialesDestino
            data.ImputacionGlobalMaterialesLotes = datMatGlobal.LotesDestino

            Dim datMatLineas As New DataPropuestaImputacionMateriales(data.OperacionCabecera(0)("NOperacion"), dtrCabeceraOrigen("NOperacionPlan"), False)
            datMatLineas = ProcessServer.ExecuteTask(Of DataPropuestaImputacionMateriales, DataPropuestaImputacionMateriales)(AddressOf PropuestaImputacionMateriales, datMatLineas, services)
            data.ImputacionLineaMaterial = datMatLineas.MaterialesDestino
            data.ImputacionLineaMaterialLotes = datMatLineas.LotesDestino

            'data.ImputacionLineaMaterial = stImputacionesMatInfo.TablaImputacionLineaDestino


            'CENTROS
            Dim dttCentroGlobalDestinos As DataTable = ClsBE.Filter("frmBdgOperacionCentroRealGlobal", New NoRowsFilterItem)
            Dim dttCentroLineaDestinos As DataTable = ClsBE.Filter("frmBdgOperacionCentro", New NoRowsFilterItem)
            Dim dttCentroGlobalOrigenes As DataTable = ClsBE.Filter("frmBdgOperacionPlanCentroGlobal", FilNOperPlan)
            Dim dttCentroLineaOrigenes As DataTable = ClsBE.Filter("frmBdgOperacionVinoPlanCentro", FilNOperPlan)

            Dim stImputacionesCentroInfo As New StPropuestaImputacion(data.OperacionCabecera, data.OperacionVinoDestino, dtrCabeceraOrigen("ImputacionGlobalCentro"), _
                                                                        dttCentroGlobalOrigenes, dttCentroGlobalDestinos, dttCentroLineaOrigenes, dttCentroLineaDestinos, "IDOperacionCentro")
            If Not Nz(dtrCabeceraOrigen("ImputacionGlobalCentro"), False) Then
                stImputacionesCentroInfo.PKField = "IDVinoCentro"
            End If
            stImputacionesCentroInfo = ProcessServer.ExecuteTask(Of StPropuestaImputacion, StPropuestaImputacion)(AddressOf PropuestaImputacion, stImputacionesCentroInfo, services)
            data.ImputacionGlobalCentros = stImputacionesCentroInfo.TablaImputacionGlobalDestino
            data.ImputacionLineaCentros = stImputacionesCentroInfo.TablaImputacionLineaDestino
            ProcessServer.ExecuteTask(Of StPropuestaConfirmacionPlan, StPropuestaConfirmacionPlan)(AddressOf AsignarCantidadCentros, data, services)

            'MOD
            Dim dttMODGlobalDestinos As DataTable = ClsBE.Filter("frmBdgOperacionMODRealGlobal", New NoRowsFilterItem)
            Dim dttMODLineaDestinos As DataTable = ClsBE.Filter("frmBdgOperacionMod", New NoRowsFilterItem)

            Dim dttMODGlobalOrigenes As DataTable = ClsBE.Filter("frmBdgOperacionPlanMODGlobal", FilNOperPlan)
            Dim dttMODLineaOrigenes As DataTable = ClsBE.Filter("frmBdgOperacionVinoPlanMOD", FilNOperPlan)

            Dim stImputacionesMODInfo As New StPropuestaImputacion(data.OperacionCabecera, data.OperacionVinoDestino, dtrCabeceraOrigen("ImputacionGlobalMod"), _
                                                                        dttMODGlobalOrigenes, dttMODGlobalDestinos, dttMODLineaOrigenes, dttMODLineaDestinos, "IDOperacionMOD")
            If Not Nz(dtrCabeceraOrigen("ImputacionGlobalMod"), False) Then
                stImputacionesMODInfo.PKField = "IDVinoMOD"
            End If
            stImputacionesMODInfo = ProcessServer.ExecuteTask(Of StPropuestaImputacion, StPropuestaImputacion)(AddressOf PropuestaImputacion, stImputacionesMODInfo, services)
            data.ImputacionGlobalMOD = stImputacionesMODInfo.TablaImputacionGlobalDestino
            data.ImputacionLineaMOD = stImputacionesMODInfo.TablaImputacionLineaDestino
            ProcessServer.ExecuteTask(Of StPropuestaConfirmacionPlan, StPropuestaConfirmacionPlan)(AddressOf AsignarHoraCategoria, data, services)

            'VARIOS
            Dim dttVariosGlobalDestinos As DataTable = ClsBE.Filter("frmBdgOperacionVariosRealGlobal", New NoRowsFilterItem)
            Dim dttVariosLineaDestinos As DataTable = ClsBE.Filter("frmBdgOperacionVarios", New NoRowsFilterItem)
            Dim dttVariosGlobalOrigenes As DataTable = ClsBE.Filter("frmBdgOperacionPlanVariosGlobal", FilNOperPlan)
            Dim dttVariosLineaOrigenes As DataTable = ClsBE.Filter("frmBdgOperacionVinoPlanVarios", FilNOperPlan)

            Dim stImputacionesVariosInfo As New StPropuestaImputacion(data.OperacionCabecera, data.OperacionVinoDestino, dtrCabeceraOrigen("ImputacionGlobalVarios"), _
                                                                        dttVariosGlobalOrigenes, dttVariosGlobalDestinos, dttVariosLineaOrigenes, dttVariosLineaDestinos, "IDOperacionVarios")
            If Not Nz(dtrCabeceraOrigen("ImputacionGlobalVarios"), False) Then
                stImputacionesVariosInfo.PKField = "IDVinoVarios"
            End If
            stImputacionesVariosInfo = ProcessServer.ExecuteTask(Of StPropuestaImputacion, StPropuestaImputacion)(AddressOf PropuestaImputacion, stImputacionesVariosInfo, services)
            data.ImputacionGlobalVarios = stImputacionesVariosInfo.TablaImputacionGlobalDestino
            data.ImputacionLineaVarios = stImputacionesVariosInfo.TablaImputacionLineaDestino
        End If
        Return data
    End Function

    <Task()> Public Shared Function PropuestaConfirmacionAnalisisOperacion(ByVal data As StPropuestaConfirmacionPlan, ByVal services As ServiceProvider) As StPropuestaConfirmacionPlan
        data.OperacionAnalitica = New DataEngine().Filter("frmBdgOperacionAnalisis", New NoRowsFilterItem)
        If Not data.OperacionCabecera Is Nothing AndAlso data.OperacionCabecera.Rows.Count > 0 Then
            Dim drOperacion As DataRow = data.OperacionCabecera.Rows(0)
            If Length(drOperacion("IDAnalisis")) > 0 Then
                Dim dtVars As DataTable = ProcessServer.ExecuteTask(Of String, DataTable)(AddressOf BdgOperacion.GetVars, drOperacion("IDAnalisis"), services)
                For Each drDestino As DataRow In data.OperacionVinoDestino.Select
                    For Each oRw As DataRow In dtVars.Select
                        Dim nwRw As DataRow = data.OperacionAnalitica.NewRow
                        data.OperacionAnalitica.Rows.Add(nwRw)
                        nwRw("IDVino") = drDestino("IDVino")
                        For Each dc As DataColumn In dtVars.Columns
                            If data.OperacionAnalitica.Columns.Contains(dc.ColumnName) Then
                                nwRw(dc.ColumnName) = oRw(dc)
                            End If
                        Next
                    Next
                Next
            End If
        End If
        Return data
    End Function

    <Serializable()> Public Class StObtenerLineasPropuesta
        Public Cabecera As DataTable
        Public Origen As DataTable
        Public Destino As DataTable

        Public AutoCalcularVino As Boolean

        Public Sub New(ByVal cabecera As DataTable, ByVal origen As DataTable, ByVal destino As DataTable, Optional ByVal calcularVino As Boolean = True)
            Me.Cabecera = cabecera
            Me.Origen = origen
            Me.Destino = destino

            Me.AutoCalcularVino = calcularVino
        End Sub

    End Class


    <Task()> Public Shared Function ObtenerLineasPropuesta(ByVal data As StObtenerLineasPropuesta, ByVal services As ServiceProvider) As StObtenerLineasPropuesta
        If (Not data.Destino.Columns.Contains("IDLineaOperacionVinoPlan")) Then data.Destino.Columns.Add("IDLineaOperacionVinoPlan")
        For Each DrOrigen As DataRow In data.Origen.Select
            Dim DrDestino As DataRow = data.Destino.NewRow()
            DrDestino("IDLineaOperacionVinoPlan") = DrOrigen("IDLineaOperacionVinoPlan")
            DrDestino("IDVino") = New Guid
            DrDestino("NOperacion") = data.Cabecera.Rows(0)("NOperacion")
            DrDestino("Cantidad") = DrOrigen("Cantidad")
            DrDestino("IDArticulo") = DrOrigen("IDArticulo")
            DrDestino("DescArticulo") = DrOrigen("DescArticulo")
            DrDestino("IDDeposito") = DrOrigen("IDDeposito")
            DrDestino("Merma") = DrOrigen("Merma")
            DrDestino("IDUdMedida") = DrOrigen("IDUdMedida")
            DrDestino("Ocupacion") = DrOrigen("Ocupacion")
            DrDestino("Lote") = DrOrigen("Lote")
            DrDestino("IDEstadoVino") = DrOrigen("IDEstadoVino")
            DrDestino("TipoDeposito") = DrOrigen("TipoDeposito")
            DrDestino("Destino") = DrOrigen("Destino")
            DrDestino("QDeposito") = DrOrigen("QDeposito")
            DrDestino("Litros") = DrOrigen("Litros")
            DrDestino("IDBarrica") = DrOrigen("IDBarrica")
            DrDestino("IDTipoMermaVino") = Nz(DrOrigen("IDTipoMermaVino"), String.Empty)
            If Length(DrOrigen("IDDeposito")) > 0 Then
                Dim stvino As New BdgGeneral.DataObtenerVinoEnDeposito(DrOrigen("IDDeposito"), Nz(DrOrigen("IDArticulo"), String.Empty), Nz(DrOrigen("Lote"), String.Empty), Nz(DrOrigen("IDBarrica"), String.Empty)) ', Nz(DrOrigen("IDUDMedida"), String.Empty), DrOrigen("Cantidad"))
                Dim vinos As BdgGeneral.DataObtenerVinoEnDeposito = ProcessServer.ExecuteTask(Of BdgGeneral.DataObtenerVinoEnDeposito, BdgGeneral.DataObtenerVinoEnDeposito)(AddressOf BdgGeneral.ObtenerVinoEnDeposito, stvino, services)
                Dim DtVinos As DataTable = vinos.dtVino
                If Not DtVinos Is Nothing AndAlso DtVinos.Rows.Count > 0 Then
                    'asignaremos el vino siempre y cuando el artículo que venía indicado sea el mismo qeu el que indica
                    'o si la marca de AutoAsignarVino viene marcada => a esto igual hay qeu darle una vuelta TODO
                    If Nz(DrDestino("IDArticulo"), String.Empty) = DtVinos.Rows(0)("IDArticulo") OrElse data.AutoCalcularVino Then
                        DrDestino("IDVino") = DtVinos.Rows(0)("IDVino")
                        DrDestino("Lote") = DtVinos.Rows(0)("Lote")
                        DrDestino("IDArticulo") = DtVinos.Rows(0)("IDArticulo")
                        DrDestino("DescArticulo") = DtVinos.Rows(0)("DescArticulo")
                        DrDestino("IDUdMedida") = DtVinos.Rows(0)("IDUdMedida")
                        If DrDestino.Table.Columns.Contains("Fecha") Then DrDestino("Fecha") = DtVinos.Rows(0)("Fecha")
                        DrDestino = New BdgOperacionVinoPlan().ApplyBusinessRule("IDArticulo", DrDestino("IDArticulo"), DrDestino, New BusinessData())
                    End If
                End If
            End If
            data.Destino.Rows.Add(DrDestino)
        Next
        Return data
    End Function

    <Serializable()> _
    Public Class StPropuestaImputacion
        Public CabeceraOperacion As DataTable
        Public LineasOperacion As DataTable

        Public EsImputacionGlobal As Boolean

        Public TablaImputacionGlobalOrigen As DataTable
        Public TablaImputacionLineaOrigen As DataTable

        Public TablaImputacionGlobalDestino As DataTable
        Public TablaImputacionLineaDestino As DataTable

        Public PKField As String

        Public Sub New(ByVal cabeceraOperacion As DataTable, ByVal lineasOperacion As DataTable, ByVal EsImputacionGlobal As Boolean, _
                       ByVal tablaImputacionGlobalOrigen As DataTable, ByVal tablaImputacionGlobalDestino As DataTable, _
                       ByVal tablaImputacionLineaOrigen As DataTable, ByVal tablaImputacionLineaDestino As DataTable, _
                       ByVal PKField As String)
            Me.PKField = PKField
            Me.CabeceraOperacion = cabeceraOperacion
            Me.LineasOperacion = lineasOperacion
            Me.EsImputacionGlobal = EsImputacionGlobal
            Me.TablaImputacionGlobalOrigen = tablaImputacionGlobalOrigen
            Me.TablaImputacionGlobalDestino = tablaImputacionGlobalDestino
            Me.TablaImputacionLineaOrigen = tablaImputacionLineaOrigen
            Me.TablaImputacionLineaDestino = tablaImputacionLineaDestino
        End Sub
    End Class

    <Task()> Protected Shared Function PropuestaImputacion(ByVal data As StPropuestaImputacion, ByVal services As ServiceProvider) As StPropuestaImputacion
        Dim dttGlobalPlan As DataTable = data.TablaImputacionGlobalOrigen
        If data.EsImputacionGlobal Then

            If Not dttGlobalPlan Is Nothing AndAlso dttGlobalPlan.Rows.Count > 0 Then
                Dim stCopiaInfo As New StCopiarRegistros(dttGlobalPlan, data.TablaImputacionGlobalDestino)
                stCopiaInfo = ProcessServer.ExecuteTask(Of StCopiarRegistros, StCopiarRegistros)(AddressOf CopiarRegistros, stCopiaInfo, services)
                For Each dtr As DataRow In stCopiaInfo.TablaDestino.Rows
                    dtr("NOperacion") = data.CabeceraOperacion.Rows(0)("NOperacion")
                    dtr(data.PKField) = Guid.NewGuid
                Next
                data.TablaImputacionGlobalDestino = stCopiaInfo.TablaDestino.Copy
            End If
            For Each dtr As DataRow In data.TablaImputacionGlobalDestino.Select
                If (Length(dtr("Fecha")) = 0) Then dtr("Fecha") = data.CabeceraOperacion.Rows(0)("Fecha")
            Next
        Else
            If (Not data.TablaImputacionLineaOrigen Is Nothing AndAlso data.TablaImputacionLineaOrigen.Rows.Count > 0) Then
                For Each dtrImputacionLineaOrigen As DataRow In data.TablaImputacionLineaOrigen.Select
                    Dim f As New Filter
                    f.Add("IDLineaOperacionVinoPlan", dtrImputacionLineaOrigen("IDLineaOperacionVinoPlan"))
                    Dim dtrResult As DataRow() = data.LineasOperacion.Select(f.Compose(New AdoFilterComposer))
                    If (Not dtrResult Is Nothing AndAlso dtrResult.Length > 0) Then
                        Dim dtrOrigenVino As DataRow = dtrResult(0)
                        Dim drNueva As DataRow = data.TablaImputacionLineaDestino.NewRow
                        For Each dc As DataColumn In data.TablaImputacionLineaDestino.Columns
                            If data.TablaImputacionLineaOrigen.Columns.Contains(dc.ColumnName) AndAlso data.TablaImputacionLineaDestino.Columns.Contains(dc.ColumnName) Then
                                drNueva(dc.ColumnName) = dtrImputacionLineaOrigen(dc.ColumnName)
                            End If
                        Next
                        drNueva("NOperacion") = data.CabeceraOperacion.Rows(0)("NOperacion")
                        drNueva("IDVino") = dtrOrigenVino("IDVino")
                        If IsDBNull(drNueva(data.PKField)) Then drNueva(data.PKField) = Guid.NewGuid
                        data.TablaImputacionLineaDestino.Rows.Add(drNueva.ItemArray)
                    End If
                Next
            End If

            For Each dtr As DataRow In data.TablaImputacionLineaDestino.Select
                If (Length(dtr("Fecha")) = 0) Then dtr("Fecha") = data.CabeceraOperacion.Rows(0)("Fecha")
            Next
        End If
        Return data
    End Function


    <Serializable()> _
    Public Class DataPropuestaImputacionMateriales
        Public NOperacionPlan As String
        Public NOperacion As String

        Public MaterialesOrigen As DataTable
        Public MaterialesDestino As DataTable
        ' Public LotesOrigen As DataTable
        Public LotesDestino As DataTable
        Public EsGlobal As Boolean

        Public Sub New(ByVal NOperacion As String, ByVal NOperacionPlan As String, ByVal EsGlobal As Boolean)
            Me.NOperacion = NOperacion
            Me.NOperacionPlan = NOperacionPlan
            Me.EsGlobal = EsGlobal
        End Sub
    End Class
    <Task()> Protected Shared Function PropuestaImputacionMateriales(ByVal data As DataPropuestaImputacionMateriales, ByVal services As ServiceProvider) As DataPropuestaImputacionMateriales
        'Dim dttMatGlobalDestinos As DataTable = ClsBE.Filter("", New NoRowsFilterItem)
        'Dim dttMatLineaDestinos As DataTable = ClsBE.Filter("frmBdgOperacionMaterial", New NoRowsFilterItem)
        'Dim dttMatGlobalOrigenes As DataTable = ClsBE.Filter("frmBdgOperacionPlanMaterialGlobal", FilNOperPlan)
        'Dim dttMatLineaOrigenes As DataTable = ClsBE.Filter("frmBdgOperacionVinoPlanMaterial", FilNOperPlan)
        Dim f As New Filter
        Dim View_Lotes_Origen As String
        Dim FieldIDMaterialOrigen As String
        Dim FieldIDMaterialDestino As String
        Dim FiledIDLote As String
        Dim Business As BusinessHelper
        f.Add(New FilterItem("NOperacionPlan", data.NOperacionPlan))
        If data.EsGlobal Then
            data.MaterialesOrigen = New BE.DataEngine().Filter("frmBdgOperacionPlanMaterialGlobal", f)

            View_Lotes_Origen = "tbBdgOperacionPlanMaterialLote"
            FieldIDMaterialOrigen = "IDOperacionPlanMaterial"

            FieldIDMaterialDestino = "IDOperacionMaterial"
            FiledIDLote = "IDOperacionMaterialLote"
            data.MaterialesDestino = New BE.DataEngine().Filter("frmBdgOperacionMaterialRealGlobal", New NoRowsFilterItem)
            data.LotesDestino = New BE.DataEngine().Filter("tbBdgOperacionMaterialLote", New NoRowsFilterItem)
            data.LotesDestino.Columns.Add("IDArticulo", GetType(String))
            data.LotesDestino.Columns.Add("IDAlmacen", GetType(String))

        Else
            data.MaterialesOrigen = New BE.DataEngine().Filter("frmBdgOperacionVinoPlanMaterial", f)
            Business = New BdgVinoMaterial
            View_Lotes_Origen = "tbBdgOperacionVinoPlanMaterialLote"
            FieldIDMaterialOrigen = "IDOperacionVinoPlanMaterial"

            FieldIDMaterialDestino = "IDVinoMaterial"
            FiledIDLote = "IDVinoMaterialLote"

            data.MaterialesDestino = New BE.DataEngine().Filter("frmBdgOperacionMaterial", New NoRowsFilterItem)
            data.LotesDestino = New BE.DataEngine().Filter("tbBdgVinoMaterialLote", New NoRowsFilterItem)
            data.LotesDestino.Columns.Add("NOperacion", GetType(String))
            data.LotesDestino.Columns.Add("IDArticulo", GetType(String))
            data.LotesDestino.Columns.Add("IDAlmacen", GetType(String))
        End If


        If Not data.MaterialesOrigen Is Nothing AndAlso data.MaterialesOrigen.Rows.Count > 0 Then
            If Not data.MaterialesOrigen Is Nothing AndAlso data.MaterialesOrigen.Rows.Count > 0 Then
                For Each dtrOrigen As DataRow In data.MaterialesOrigen.Rows


                    Dim dtrNewRow As DataRow = data.MaterialesDestino.NewRow

                    For Each dtcColumn As DataColumn In data.MaterialesOrigen.Columns
                        If (data.MaterialesOrigen.Columns.Contains(dtcColumn.ColumnName) AndAlso _
                            data.MaterialesDestino.Columns.Contains(dtcColumn.ColumnName) AndAlso _
                            Length(dtrOrigen(dtcColumn.ColumnName))) Then

                            'dtrNewRow(dtcColumn.ColumnName) = dtrOrigen(dtcColumn.ColumnName)
                            '    dtrNewRow = Business.ApplyBusinessRule(dtcColumn.ColumnName, dtrOrigen(dtcColumn.ColumnName), dtrNewRow, Nothing)

                        End If
                    Next
                    dtrNewRow(FieldIDMaterialDestino) = Guid.NewGuid
                    dtrNewRow("NOperacion") = data.NOperacion
                    data.MaterialesDestino.Rows.Add(dtrNewRow.ItemArray)



                    '//Lotes
                    Dim dtLotesOrigen As DataTable = New BE.DataEngine().Filter(View_Lotes_Origen, New GuidFilterItem(FieldIDMaterialOrigen, dtrOrigen(FieldIDMaterialOrigen)))
                    If Not dtLotesOrigen Is Nothing AndAlso dtLotesOrigen.Rows.Count > 0 Then
                        For Each drLote As DataRow In dtLotesOrigen.Rows
                            Dim drNewLote As DataRow = data.LotesDestino.NewRow
                            For Each col As DataColumn In dtLotesOrigen.Columns
                                If (dtLotesOrigen.Columns.Contains(col.ColumnName) AndAlso _
                                    data.LotesDestino.Columns.Contains(col.ColumnName) AndAlso _
                                    Length(dtrOrigen(col.ColumnName))) Then
                                    drNewLote(col.ColumnName) = dtrOrigen(col.ColumnName)
                                End If
                            Next
                            drNewLote(FieldIDMaterialDestino) = dtrNewRow(FieldIDMaterialDestino)
                            drNewLote(FiledIDLote) = Guid.NewGuid
                            If data.LotesDestino.Columns.Contains("NOperacion") Then drNewLote("NOperacion") = data.NOperacion
                            drNewLote("IDArticulo") = dtrNewRow("IDArticulo")
                            drNewLote("IDAlmacen") = dtrNewRow("IDAlmacen")

                            data.LotesDestino.Rows.Add(drNewLote)
                        Next
                    End If



                    dtrNewRow("NumLotes") = 0
                    Dim LotesLinea As List(Of DataRow) = (From c In data.LotesDestino Where Not c.IsNull(FieldIDMaterialDestino) AndAlso c(FieldIDMaterialDestino) = dtrNewRow(FieldIDMaterialDestino) Select c).ToList
                    If Not LotesLinea Is Nothing AndAlso LotesLinea.Count > 0 Then
                        dtrNewRow("NumLotes") = LotesLinea.Count
                        If LotesLinea.Count = 1 Then
                            dtrNewRow("Lote") = LotesLinea(0)("Lote")
                            dtrNewRow("Ubicacion") = LotesLinea(0)("Ubicacion")
                        End If
                    End If
                Next
            End If
        End If
        Return data
    End Function



    <Task()> Protected Shared Function AsignarHoraCategoria(ByVal data As StPropuestaConfirmacionPlan, ByVal services As ServiceProvider) As StPropuestaConfirmacionPlan

        Dim SinHoraCategoria As List(Of DataRow) = (From c In data.ImputacionLineaMOD Where c.IsNull("IDHora") AndAlso Not c.IsNull("IDOperario") Select c).ToList
        If Not SinHoraCategoria Is Nothing AndAlso SinHoraCategoria.Count > 0 Then
            Dim context As New BusinessData
            context("Fecha") = Nz(data.OperacionCabecera.Rows(0)("Fecha"), Today)

            For Each dr As DataRow In SinHoraCategoria
                Dim brData As New BusinessRuleData("IDOperario", dr("IDOperario"), New BusinessData(dr), context)
                ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf BdgGeneral.CambioOperario, brData, services)
                For Each col As DataColumn In dr.Table.Columns
                    dr(col.ColumnName) = brData.Current(col.ColumnName)
                Next
            Next
        End If

        Dim SinHoraCategoriaGlobal As List(Of DataRow) = (From c In data.ImputacionGlobalMOD Where c.IsNull("IDHora") AndAlso Not c.IsNull("IDOperario") Select c).ToList
        If Not SinHoraCategoriaGlobal Is Nothing AndAlso SinHoraCategoriaGlobal.Count > 0 Then
            Dim context As New BusinessData
            context("Fecha") = Nz(data.OperacionCabecera.Rows(0)("Fecha"), Today)

            For Each dr As DataRow In SinHoraCategoriaGlobal
                Dim brData As New BusinessRuleData("IDOperario", dr("IDOperario"), New BusinessData(dr), context)
                ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf BdgGeneral.CambioOperario, brData, services)
                For Each col As DataColumn In dr.Table.Columns
                    dr(col.ColumnName) = brData.Current(col.ColumnName)
                Next
            Next
        End If

        Return data
    End Function

    <Task()> Protected Shared Function AsignarCantidadCentros(ByVal data As StPropuestaConfirmacionPlan, ByVal services As ServiceProvider) As StPropuestaConfirmacionPlan
        If Not data.OperacionVinoDestino Is Nothing AndAlso data.OperacionVinoDestino.Rows.Count > 0 Then
            If Not data.ImputacionLineaCentros Is Nothing AndAlso data.ImputacionLineaCentros.Rows.Count > 0 Then

                Dim SinCantidad As List(Of DataRow) = (From c In data.ImputacionLineaCentros _
                                                       Where Not c.IsNull("PorCantidad") AndAlso c("PorCantidad") = True AndAlso _
                                                                (c.IsNull("Cantidad") OrElse c("Cantidad") = 0) _
                                                       Select c).ToList
                If Not SinCantidad Is Nothing AndAlso SinCantidad.Count > 0 Then
                    For Each dr As DataRow In SinCantidad
                        If Not IsDBNull(dr("IDVino")) Then
                            Dim CantidadVino As List(Of DataRow) = (From c In data.OperacionVinoDestino Where c("IDVino") = dr("IDVino") Select c).ToList
                            If Not CantidadVino Is Nothing AndAlso CantidadVino.Count > 0 Then
                                dr("Cantidad") = Nz(CantidadVino(0)("Cantidad"), 0)
                            End If
                        End If

                    Next
                End If
            End If
        End If

        Return data
    End Function

#End Region

#Region "Tareas - Confirmación de Propuestas"

    '<Serializable()> Public Class StErrorConfirmacion
    '    Public Operacion As String
    '    Public Mensaje As String

    '    Public Sub New(ByVal strOper As String, ByVal strMen As String)
    '        Me.Operacion = strOper
    '        Me.Mensaje = strMen
    '    End Sub

    'End Class

    '<Serializable()> Public Class StResultadoConfirmacionMasiva
    '    Public OperacionesPlanificadas() As String ' OPERACIONES PLAN CONFIRMADAS CORRECTAMENTE
    '    Public Operaciones() As String 'OPERACIONES GENERADAS
    '    Public OperacionesFallidas() As StErrorConfirmacion

    '    Public Sub New()
    '        ReDim Operaciones(-1)
    '        ReDim OperacionesFallidas(-1)
    '        ReDim OperacionesPlanificadas(-1)
    '    End Sub
    'End Class

    '<Serializable()> Public Class StConfirmacionMasiva
    '    Public Operaciones() As String
    '    Public ValidarDepositoDestinoVacio As Boolean
    '    Public AutoasignarVino As Boolean

    '    Public Sub New(ByVal operaciones() As String, Optional ByVal validarDepositoDestinoVacio As Boolean = False, Optional ByVal autoAsignarVino As Boolean = True)
    '        Me.Operaciones = operaciones
    '        Me.ValidarDepositoDestinoVacio = validarDepositoDestinoVacio
    '        Me.AutoasignarVino = autoAsignarVino
    '    End Sub

    'End Class

    '<Serializable()> Public Class StResultadoPropuestaConfirmacion
    '    Public NOperacionPlan As String
    '    Public NOperacion As String
    '    Public ErrorMessage As String

    '    Public Sub New(ByVal NOperacionPlan As String)
    '        Me.NOperacionPlan = NOperacionPlan
    '    End Sub
    'End Class

    '<Task()> Public Shared Function ConfirmacionMasiva(ByVal data As StConfirmacionMasiva, ByVal service As ServiceProvider) As StResultadoConfirmacionMasiva
    '    Dim resultData As New StResultadoConfirmacionMasiva
    '    If data Is Nothing Then Return resultData
    '    If data.Operaciones Is Nothing OrElse data.Operaciones.Length = 0 Then Return resultData
    '    For Each operacion As String In data.Operaciones
    '        Try
    '            AdminData.BeginTx()

    '            Dim currentOperationConfirmation As New StPropuestaConfirmacionPlan(operacion, data.ValidarDepositoDestinoVacio, data.AutoasignarVino)
    '            currentOperationConfirmation = ProcessServer.ExecuteTask(Of StPropuestaConfirmacionPlan, StPropuestaConfirmacionPlan)(AddressOf BdgOperacionPlan.PropuestaConfirmarPlan, currentOperationConfirmation, service)
    '            Dim result As StResultadoPropuestaConfirmacion = ProcessServer.ExecuteTask(Of StPropuestaConfirmacionPlan, StResultadoPropuestaConfirmacion)(AddressOf BdgOperacionPlan.ConfirmarPropuestaOperacion, currentOperationConfirmation, service)
    '            If Length(result.NOperacion) = 0 Then
    '                ReDim Preserve resultData.OperacionesFallidas(UBound(resultData.OperacionesFallidas) + 1)
    '                resultData.OperacionesFallidas(resultData.OperacionesFallidas.Length - 1) = New StErrorConfirmacion(operacion, result.ErrorMessage)
    '                AdminData.RollBackTx(True)
    '            Else
    '                ReDim Preserve resultData.Operaciones(UBound(resultData.Operaciones) + 1)
    '                ReDim Preserve resultData.OperacionesPlanificadas(UBound(resultData.OperacionesPlanificadas) + 1)
    '                resultData.OperacionesPlanificadas(resultData.OperacionesPlanificadas.Length - 1) = currentOperationConfirmation.NOperacionPlan
    '                resultData.Operaciones(resultData.Operaciones.Length - 1) = result.NOperacion
    '            End If
    '            AdminData.CommitTx(True)
    '        Catch ex As Exception
    '            AdminData.RollBackTx(True)
    '            ReDim Preserve resultData.OperacionesFallidas(UBound(resultData.OperacionesFallidas) + 1)
    '            resultData.OperacionesFallidas(resultData.OperacionesFallidas.Length - 1) = New StErrorConfirmacion(operacion, ex.Message)
    '        End Try
    '    Next
    '    Return resultData
    'End Function

    '<Task()> Public Shared Function ConfirmarPropuestaOperacion(ByVal data As StPropuestaConfirmacionPlan, ByVal service As ServiceProvider) As StResultadoPropuestaConfirmacion
    '    Dim StResult As New StResultadoPropuestaConfirmacion(data.NOperacionPlan)
    '    If ProcessServer.ExecuteTask(Of StPropuestaConfirmacionPlan, Boolean)(AddressOf ValidarMermas, data, service) Then
    '        If ProcessServer.ExecuteTask(Of StPropuestaConfirmacionPlan, Boolean)(AddressOf ValidarVinoIndicado, data, service) Then
    '            If ProcessServer.ExecuteTask(Of StPropuestaConfirmacionPlan, Boolean)(AddressOf ValidarBarrica, data, service) Then
    '                If (data.ValidarVinosDepositoDestino) Then
    '                    If ProcessServer.ExecuteTask(Of StPropuestaConfirmacionPlan, Boolean)(AddressOf ValidarVinoEnDeposito, data, service) Then
    '                        Dim dblQTotal As Double = 0
    '                        For Each dtrDestino As DataRow In data.OperacionVinoDestino.Select
    '                            dblQTotal += Nz(dtrDestino("Cantidad"), 0)
    '                        Next
    '                        Dim ClsBDgOperPlan As New BdgOperacionPlan
    '                        Dim dttMat As DataTable = ClsBDgOperPlan.RepartoGlobales(data.OperacionCabecera, data.OperacionVinoDestino, data.ImputacionGlobalMateriales, _
    '                                                                  data.ImputacionLineaMaterial, dblQTotal, "ImputacionRealMaterial", "IDOperacionMaterial", AddressOf Business.Bodega.BdgOperacion.MaterialesReparto)
    '                        If (Not dttMat Is Nothing AndAlso dttMat.Rows.Count > 0) Then
    '                            For Each dtr As DataRow In dttMat.Select
    '                                data.ImputacionLineaMaterial.ImportRow(dtr)
    '                            Next
    '                        End If
    '                        Dim dttMod As DataTable = ClsBDgOperPlan.RepartoGlobales(data.OperacionCabecera, data.OperacionVinoDestino, data.ImputacionGlobalMOD, _
    '                                                              data.ImputacionLineaMOD, dblQTotal, "ImputacionRealMod", "IDOperacionMOD", AddressOf Business.Bodega.BdgOperacion.ManoObraReparto)
    '                        If (Not dttMod Is Nothing AndAlso dttMod.Rows.Count > 0) Then
    '                            For Each dtr As DataRow In dttMod.Select
    '                                data.ImputacionLineaMOD.ImportRow(dtr)
    '                            Next
    '                        End If
    '                        Dim dttCentros As DataTable = ClsBDgOperPlan.RepartoGlobales(data.OperacionCabecera, data.OperacionVinoDestino, data.ImputacionGlobalCentros, _
    '                                                      data.ImputacionLineaCentros, dblQTotal, "ImputacionRealCentro", "IDOperacionCentro", AddressOf Business.Bodega.BdgOperacion.CentroReparto)
    '                        If (Not dttCentros Is Nothing AndAlso dttCentros.Rows.Count > 0) Then
    '                            For Each dtr As DataRow In dttCentros.Select
    '                                data.ImputacionLineaCentros.ImportRow(dtr)
    '                            Next
    '                        End If
    '                        Dim dttVarios As DataTable = ClsBDgOperPlan.RepartoGlobales(data.OperacionCabecera, data.OperacionVinoDestino, data.ImputacionGlobalVarios, _
    '                                                      data.ImputacionLineaVarios, dblQTotal, "ImputacionRealVarios", "IDOperacionVarios", AddressOf Business.Bodega.BdgOperacion.VariosReparto)
    '                        If (Not dttVarios Is Nothing AndAlso dttVarios.Rows.Count > 0) Then
    '                            For Each dtr As DataRow In dttVarios.Select
    '                                data.ImputacionLineaVarios.ImportRow(dtr)
    '                            Next
    '                        End If
    '                        If ProcessServer.ExecuteTask(Of StPropuestaConfirmacionPlan, Boolean)(AddressOf ValidarCategoriaImputaciones, data, service) Then
    '                            Dim StCrear As New Business.Bodega.BdgOperacion.StCrearOperacion(data.OperacionCabecera, data.OperacionVinoOrigen, data.OperacionVinoDestino, data.OperacionAnalitica, _
    '                                                                                 data.ImputacionLineaMaterial, data.ImputacionLineaCentros, data.ImputacionLineaMOD, data.ImputacionLineaVarios, _
    '                                                                                 data.ImputacionGlobalMateriales, data.ImputacionGlobalMOD, data.ImputacionGlobalCentros, data.ImputacionGlobalVarios)
    '                            StResult.NOperacion = ProcessServer.ExecuteTask(Of Business.Bodega.BdgOperacion.StCrearOperacion, String)(AddressOf Business.Bodega.BdgOperacion.CrearOperacion, StCrear, service)
    '                        Else : StResult.ErrorMessage = "Existen líneas de imputación sin categoría asignada."
    '                        End If
    '                    Else : StResult.ErrorMessage = "Existe vino en los depósitos de destino."
    '                    End If
    '                End If
    '            Else : StResult.ErrorMessage = "Existen depósitos en los que la barrica es obligatoria."
    '            End If
    '        Else : StResult.ErrorMessage = "No se ha indicado el vino en todas las líneas."
    '        End If
    '    Else : StResult.ErrorMessage = "Las cantidades de origen y de destino no coinciden."
    '    End If
    '    Return StResult
    'End Function


#Region "Tareas Secundarias"

    '<Task()> Public Shared Function ValidarMermas(ByVal data As StPropuestaConfirmacionPlan, ByVal service As ServiceProvider) As Boolean
    '    Dim dtrTipoOper As DataRow = New BdgTipoOperacion().GetItemRow(data.OperacionCabecera.Rows(0)("IDTipoOperacion"))
    '    If dtrTipoOper("TipoMovimiento") = Business.Bodega.TipoMovimiento.SinMovimiento Then Return True
    '    If dtrTipoOper("TipoMovimiento") = Business.Bodega.TipoMovimiento.SinOrigen Then Return True
    '    If dtrTipoOper("TipoMovimiento") = Business.Bodega.TipoMovimiento.Salida Then Return True
    '    Dim dblQTotalOrigen As Double = 0
    '    For Each dtrOrigen As DataRow In data.OperacionVinoOrigen.Rows
    '        dblQTotalOrigen += dtrOrigen("Litros")
    '    Next
    '    For Each dtrDestino As DataRow In data.OperacionVinoDestino.Rows
    '        dblQTotalOrigen -= dtrDestino("Litros")
    '    Next
    '    Return dblQTotalOrigen = 0
    'End Function

    '<Task()> Public Shared Function ValidarBarrica(ByVal data As StPropuestaConfirmacionPlan, ByVal service As ServiceProvider) As Boolean
    '    If Not data.OperacionCabecera Is Nothing AndAlso data.OperacionCabecera.Rows.Count > 0 AndAlso Length(data.OperacionCabecera.Rows(0)("IDTipoOperacion")) > 0 Then
    '        Dim dtrTipoOper As DataRow = New BdgTipoOperacion().GetItemRow(data.OperacionCabecera.Rows(0)("IDTipoOperacion"))
    '        If Not dtrTipoOper Is Nothing Then
    '            If dtrTipoOper("TipoMovimiento") <> Business.Bodega.TipoMovimiento.SinMovimiento Then
    '                Dim strDepositosSinBarrica As String = String.Empty
    '                For Each drDestino As DataRow In data.OperacionVinoDestino.Select
    '                    If drDestino.RowState <> DataRowState.Deleted Then
    '                        Dim blnUsarBarricaComoLote As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf Business.Bodega.BdgDeposito.UsarBarricaComoLote, drDestino("IDDeposito"), service)
    '                        If blnUsarBarricaComoLote AndAlso Length(drDestino("IDBarrica")) = 0 Then
    '                            Return False
    '                        End If
    '                    End If
    '                Next
    '            End If
    '            Return True
    '        Else : Return False
    '        End If
    '    Else : Return False
    '    End If
    'End Function

    '<Task()> Public Shared Function ValidarVinoIndicado(ByVal data As StPropuestaConfirmacionPlan, ByVal service As ServiceProvider) As Boolean
    '    Dim blnIsValid As Boolean = True
    '    If Not data.OperacionVinoOrigen Is Nothing AndAlso data.OperacionVinoOrigen.Rows.Count > 0 Then
    '        Dim DrFind() As DataRow = data.OperacionVinoOrigen.Select("IDVino IS NULL")
    '        If DrFind.Length > 0 Then Return False
    '    End If
    '    If Not data.OperacionVinoDestino Is Nothing AndAlso data.OperacionVinoDestino.Rows.Count > 0 Then
    '        Dim DrFind() As DataRow = data.OperacionVinoDestino.Select("IDVino IS NULL")
    '        If DrFind.Length > 0 Then Return False
    '    End If
    '    Return blnIsValid
    'End Function

    '<Task()> Public Shared Function ValidarVinoEnDeposito(ByVal data As StPropuestaConfirmacionPlan, ByVal service As ServiceProvider) As Boolean
    '    If Not data.OperacionCabecera Is Nothing AndAlso data.OperacionCabecera.Rows.Count > 0 Then
    '        Dim dtrTipoOper As DataRow = New BdgTipoOperacion().GetItemRow(data.OperacionCabecera.Rows(0)("IDTipoOperacion"))
    '        If dtrTipoOper("TipoMovimiento") <> Business.Bodega.TipoMovimiento.SinMovimiento Then
    '            Dim dtDestino As DataTable = data.OperacionVinoDestino
    '            Dim strDepositosConMezclas As String = ProcessServer.ExecuteTask(Of DataTable, String)(AddressOf Business.Bodega.BdgOperacion.ValidaMezclaVino, dtDestino, service)
    '            If Length(strDepositosConMezclas) > 0 Then
    '                Return False
    '            End If
    '        End If
    '        Return True
    '    Else : Return False
    '    End If
    'End Function

    '<Task()> Public Shared Function ValidarCategoriaImputaciones(ByVal data As StPropuestaConfirmacionPlan, ByVal service As ServiceProvider) As Boolean
    '    If (Not data.ImputacionLineaMOD Is Nothing AndAlso data.ImputacionLineaMOD.Rows.Count > 0) Then
    '        Dim DrFind() As DataRow = data.ImputacionLineaMOD.Select("IDCategoria IS NULL")
    '        If DrFind.Length > 0 Then Return False
    '    End If
    '    If (Not data.ImputacionGlobalMOD Is Nothing AndAlso data.ImputacionGlobalMOD.Rows.Count > 0) Then
    '        Dim DrFind() As DataRow = data.ImputacionGlobalMOD.Select("IDCategoria IS NULL")
    '        If DrFind.Length > 0 Then Return False
    '    End If
    '    Return True
    'End Function

    'Private Function RepartoGlobales(ByVal dttCabecera As DataTable, ByVal dttVinoDestino As DataTable, ByVal dttGlobal As DataTable, ByVal dttLinea As DataTable, ByVal totalQ As Double, _
    '                                 ByVal strBoolKey As String, ByVal strKey As String, ByVal metodoReparto As Task(Of Business.Bodega.BdgOperacion.StReparto, DataTable)) As DataTable
    '    If dttCabecera.Rows(0)(strBoolKey) Then
    '        'Eliminar lo que pueda venir si están las marcas de 
    '        Dim strWhere As String = New IsNullFilterItem(strKey, False).Compose(New AdoFilterComposer)
    '        For Each oRw As DataRow In dttLinea.Select(strWhere)
    '            oRw.Delete()
    '        Next
    '        'Regenerarlo
    '        If (dttGlobal.Rows.Count > 0 AndAlso dttVinoDestino.Rows.Count > 0) Then
    '            Dim StMat As New Business.Bodega.BdgOperacion.StReparto(dttVinoDestino, dttGlobal, totalQ)
    '            Return ProcessServer.ExecuteTask(Of Business.Bodega.BdgOperacion.StReparto, DataTable)(metodoReparto, StMat, New ServiceProvider())
    '        End If
    '    End If
    '    Return Nothing
    'End Function

#End Region

#End Region

#Region "Tareas - Generación de Planificadas - Diagramas"

    ' <Serializable()> Public Class StNuevaOperacionPlanificadaMultipleResult
    '     'Public Operaciones() As String
    '     Public NOperacion() As String

    '     Public Sub New()
    '         ReDim Preserve NOperacion(-1)
    '         'ReDim Preserve Operaciones(-1)
    '     End Sub
    ' End Class

    ' <Serializable()> _
    'Public Class StNuevaOperacionPlanificadaMultiple
    '     Public IDTipoOperacion As String
    '     Public FechaOperacion As Date
    '     Public UnaOperacion As Boolean
    '     Public DtOrigen As DataTable
    '     Public DtDestino As DataTable

    '     Public Sub New()
    '     End Sub

    '     Public Sub New(ByVal IDTipoOperacion As String, ByVal FechaOperacion As Date, ByVal UnaOperacion As Boolean, ByVal DtOrigen As DataTable, ByVal DtDestino As DataTable)
    '         Me.IDTipoOperacion = IDTipoOperacion
    '         Me.FechaOperacion = FechaOperacion
    '         Me.UnaOperacion = UnaOperacion
    '         Me.DtOrigen = DtOrigen
    '         Me.DtDestino = DtDestino
    '     End Sub
    ' End Class

    '<Serializable()> _
    'Public Class StNuevaOperacionPlanificadaResult
    '    Public NOperacionPlan As String
    '    Public Errores As New List(Of ClassErrors)

    '    Public Sub New(ByVal noperacionplan As String)
    '        Me.NOperacionPlan = noperacionplan
    '    End Sub
    'End Class

    '<Serializable()> Public Class StNuevaOperacionPlanificada
    '    Public IDDeposito As String
    '    Public IDVino As Guid
    '    Public IDTipoOperacion As String
    '    Public FechaOperacion As Date
    '    Public DtDestino As DataTable
    '    Public Cantidad As Double = 0

    '    Public Sub New(ByVal iddeposito As String, ByVal idvino As Guid, ByVal idtipooperacion As String, ByVal fecha As Date, ByVal dtDestino As DataTable, Optional ByVal Cantidad As Double = 0)
    '        Me.IDDeposito = iddeposito
    '        Me.IDVino = idvino
    '        Me.IDTipoOperacion = idtipooperacion
    '        Me.FechaOperacion = fecha
    '        Me.DtDestino = dtDestino
    '        Me.Cantidad = Cantidad
    '    End Sub
    'End Class

    '<Task()> Public Shared Function NuevaOperacionPlanificada(ByVal data As StNuevaOperacionPlanificada, ByVal service As ServiceProvider) As StNuevaOperacionPlanificadaResult
    '    Try
    '        Dim strNuevaOper As String = String.Empty
    '        If Length(data.IDDeposito) > 0 And Length(data.IDTipoOperacion) > 0 Then
    '            Dim StADD As New StADDCabeceraOperacion(data.IDTipoOperacion, data.FechaOperacion)
    '            Dim dsCabecera As DataSet = ProcessServer.ExecuteTask(Of StADDCabeceraOperacion, DataSet)(AddressOf ADDCabeceraOperacion, StADD, service)
    '            Dim dtOperacion As DataTable = dsCabecera.Tables("BdgOperacionPlan")
    '            Dim dtMaterialGlobal As DataTable = dsCabecera.Tables("BdgOperacionMaterial")
    '            Dim dtMODGlobal As DataTable = dsCabecera.Tables("BdgOperacionMOD")
    '            Dim dtCentroGlobal As DataTable = dsCabecera.Tables("BdgOperacionCentro")
    '            Dim dtVariosGlobal As DataTable = dsCabecera.Tables("BdgOperacionVarios")
    '            Dim drTipoOperacion As DataRow = New BdgTipoOperacion().GetItemRow(data.IDTipoOperacion)
    '            Dim dtOrigen, dtDestino, dtMaterial As DataTable

    '            'origen
    '            Dim nuevaLineaData As New StAddNuevaOperacionVinoPlan(dtOperacion, data.IDDeposito, data.IDVino, , , , , data.Cantidad)
    '            Dim nuevaLineaResultData As StAddNuevaOperacionVinoPlan = ProcessServer.ExecuteTask(Of StAddNuevaOperacionVinoPlan, StAddNuevaOperacionVinoPlan)(AddressOf NuevaOperacionVinoPlan, nuevaLineaData, service)
    '            If Not nuevaLineaResultData.dtOrigen Is Nothing Then dtOrigen = nuevaLineaResultData.dtOrigen.Copy
    '            If Not nuevaLineaResultData.dtDestino Is Nothing Then dtDestino = nuevaLineaResultData.dtDestino.Copy
    '            If Not nuevaLineaResultData.dtMaterial Is Nothing Then dtMaterial = nuevaLineaResultData.dtMaterial.Copy

    '            'destinos
    '            If (Not data.DtDestino Is Nothing AndAlso data.DtDestino.Rows.Count > 0) Then
    '                For Each dtr As DataRow In data.DtDestino.Select
    '                    Dim DtNuevo As DataTable = dtr.Table.Clone
    '                    DtNuevo.Rows.Add(dtr.ItemArray)
    '                    nuevaLineaData = New StAddNuevaOperacionVinoPlan(dtOperacion, dtr("IDDeposito"), Nothing, , , , True, data.Cantidad, IIf(data.Cantidad <> 0, DtNuevo, Nothing))
    '                    nuevaLineaResultData = ProcessServer.ExecuteTask(Of StAddNuevaOperacionVinoPlan, StAddNuevaOperacionVinoPlan)(AddressOf NuevaOperacionVinoPlan, nuevaLineaData, service)
    '                    If Not nuevaLineaResultData.dtDestino Is Nothing Then
    '                        If (dtDestino Is Nothing) Then
    '                            dtDestino = nuevaLineaResultData.dtDestino.Copy
    '                        Else
    '                            For Each dtrDestino As DataRow In nuevaLineaResultData.dtDestino.Select
    '                                dtDestino.ImportRow(dtrDestino)
    '                            Next
    '                        End If
    '                        dtDestino = nuevaLineaResultData.dtDestino.Copy
    '                    End If
    '                    If Not nuevaLineaResultData.dtMaterial Is Nothing Then
    '                        If (dtMaterial Is Nothing) Then
    '                            dtMaterial = nuevaLineaResultData.dtMaterial.Copy
    '                        Else
    '                            For Each dtrMat As DataRow In nuevaLineaResultData.dtMaterial.Select
    '                                dtMaterial.ImportRow(dtrMat)
    '                            Next
    '                        End If
    '                    End If
    '                Next
    '            End If
    '            If (Not dtOperacion Is Nothing AndAlso dtOperacion.Rows.Count > 0) Then strNuevaOper = dtOperacion.Rows(0)("NOperacionPlan")

    '            Dim UpdtPckg As New UpdatePackage
    '            UpdtPckg.Add(dtOperacion)
    '            If Not dtOrigen Is Nothing Then UpdtPckg.Add("BdgOperacionVinoPlan", dtOrigen)
    '            If Not dtDestino Is Nothing Then UpdtPckg.Add("BdgOperacionVinoPlan", dtDestino)
    '            If Not dtMaterial Is Nothing Then UpdtPckg.Add("BdgOperacionVinoPlanMaterial", dtMaterial)
    '            If Not dtMaterialGlobal Is Nothing Then UpdtPckg.Add("BdgOperacionPlanMaterial", dtMaterialGlobal)
    '            If Not dtMODGlobal Is Nothing Then UpdtPckg.Add("BdgOperacionPlanMOD", dtMODGlobal)
    '            If Not dtCentroGlobal Is Nothing Then UpdtPckg.Add("BdgOperacionPlanCentro", dtCentroGlobal)
    '            If Not dtVariosGlobal Is Nothing Then UpdtPckg.Add("BdgOperacionPlanVarios", dtVariosGlobal)

    '            Dim ClsBdgOperPlan As New BdgOperacionPlan
    '            ClsBdgOperPlan.Update(UpdtPckg)
    '        End If

    '        Dim resultData As New StNuevaOperacionPlanificadaResult(strNuevaOper)
    '        Return resultData
    '    Catch ex As Exception
    '        ApplicationService.GenerateError("Se ha producido un error en la generación de la operación planificada. Error: |", ex.Message)
    '    End Try
    'End Function

    '<Task()> Public Shared Function NuevaOperacionPlanificada(ByVal data As StNuevaOperacionPlanificada, ByVal service As ServiceProvider) As StNuevaOperacionPlanificadaResult
    '    Try
    '        Dim resultData As New StNuevaOperacionPlanificadaResult(String.Empty)
    '        Dim lstPropuesta As Dictionary(Of String, DataTable) = ProcessServer.ExecuteTask(Of StNuevaOperacionPlanificada, Dictionary(Of String, DataTable))(AddressOf PropuestaNuevaOperacionPlanificada, data, service)
    '        If Not lstPropuesta Is Nothing AndAlso lstPropuesta.Count > 0 Then
    '            resultData = ProcessServer.ExecuteTask(Of Dictionary(Of String, DataTable), StNuevaOperacionPlanificadaResult)(AddressOf ActualizarNuevaOperacionPlanificada, lstPropuesta, service)
    '        End If
    '        Return resultData
    '    Catch ex As Exception
    '        ApplicationService.GenerateError("Se ha producido un error en la generación de la operación planificada. Error: |", ex.Message)
    '    End Try
    'End Function


    '<Task()> Public Shared Function PropuestaNuevaOperacionPlanificada(ByVal data As StNuevaOperacionPlanificada, ByVal service As ServiceProvider) As Dictionary(Of String, DataTable)
    '    If Length(data.IDDeposito) > 0 And Length(data.IDTipoOperacion) > 0 Then

    '        Dim StADD As New StADDCabeceraOperacion(data.IDTipoOperacion, data.FechaOperacion)
    '        Dim lstPropuesta As Dictionary(Of String, DataTable) = ProcessServer.ExecuteTask(Of StADDCabeceraOperacion, Dictionary(Of String, DataTable))(AddressOf ADDCabeceraOperacion, StADD, service)
    '        If lstPropuesta Is Nothing OrElse lstPropuesta.Count = 0 Then Exit Function

    '        Dim dtOperacion As DataTable
    '        If lstPropuesta.ContainsKey(GetType(BdgOperacionPlan).Name) Then dtOperacion = lstPropuesta(GetType(BdgOperacionPlan).Name)
    '        If dtOperacion Is Nothing OrElse dtOperacion.Rows.Count = 0 Then Exit Function

    '        Dim dtOrigen, dtDestino, dtMaterial As DataTable

    '        Dim IDTipoMovimiento As TipoMovimiento?
    '        If Length(dtOperacion.Rows(0)("IDTipoOperacion")) > 0 Then
    '            Dim dtrOperacion As DataTable = New BdgTipoOperacion().SelOnPrimaryKey(dtOperacion.Rows(0)("IDTipoOperacion"))
    '            If dtrOperacion.Rows.Count > 0 Then
    '                IDTipoMovimiento = CInt(Nz(dtrOperacion.Rows(0)("TipoMovimiento"), -1))
    '            End If
    '        End If

    '        'origen
    '        Dim nuevaLineaData As New StAddNuevaOperacionVinoPlan(dtOperacion, data.IDDeposito, data.IDVino, , , , , data.Cantidad)
    '        nuevaLineaData.TipoMovimiento = IDTipoMovimiento

    '        Dim nuevaLineaResultData As StAddNuevaOperacionVinoPlan = ProcessServer.ExecuteTask(Of StAddNuevaOperacionVinoPlan, StAddNuevaOperacionVinoPlan)(AddressOf NuevaOperacionVinoPlan, nuevaLineaData, service)
    '        If Not nuevaLineaResultData.dtOrigen Is Nothing Then dtOrigen = nuevaLineaResultData.dtOrigen.Copy
    '        If Not nuevaLineaResultData.dtDestino Is Nothing Then dtDestino = nuevaLineaResultData.dtDestino.Copy
    '        If Not nuevaLineaResultData.dtMaterial Is Nothing Then dtMaterial = nuevaLineaResultData.dtMaterial.Copy

    '        'destinos
    '        If (Not data.DtDestino Is Nothing AndAlso data.DtDestino.Rows.Count > 0) Then
    '            For Each dtr As DataRow In data.DtDestino.Select
    '                Dim DtNuevo As DataTable = dtr.Table.Clone
    '                DtNuevo.Rows.Add(dtr.ItemArray)
    '                nuevaLineaData = New StAddNuevaOperacionVinoPlan(dtOperacion, dtr("IDDeposito"), Nothing, , , , True, data.Cantidad, IIf(data.Cantidad <> 0, DtNuevo, Nothing))
    '                nuevaLineaData.TipoMovimiento = IDTipoMovimiento
    '                nuevaLineaResultData = ProcessServer.ExecuteTask(Of StAddNuevaOperacionVinoPlan, StAddNuevaOperacionVinoPlan)(AddressOf NuevaOperacionVinoPlan, nuevaLineaData, service)
    '                If Not nuevaLineaResultData.dtDestino Is Nothing Then
    '                    If (dtDestino Is Nothing) Then
    '                        dtDestino = nuevaLineaResultData.dtDestino.Copy
    '                    Else
    '                        For Each dtrDestino As DataRow In nuevaLineaResultData.dtDestino.Select
    '                            dtDestino.ImportRow(dtrDestino)
    '                        Next
    '                    End If
    '                    dtDestino = nuevaLineaResultData.dtDestino.Copy
    '                End If
    '                If Not nuevaLineaResultData.dtMaterial Is Nothing Then
    '                    If (dtMaterial Is Nothing) Then
    '                        dtMaterial = nuevaLineaResultData.dtMaterial.Copy
    '                    Else
    '                        For Each dtrMat As DataRow In nuevaLineaResultData.dtMaterial.Select
    '                            dtMaterial.ImportRow(dtrMat)
    '                        Next
    '                    End If
    '                End If
    '            Next
    '        End If


    '        'Dim lstPropuesta As New Dictionary(Of String, DataTable)
    '        ' lstPropuesta.Add(GetType(BdgOperacionPlan).Name, dtOperacion)
    '        lstPropuesta.Add(GetType(BdgOperacionVinoPlan).Name & "Origen", dtOrigen)
    '        lstPropuesta.Add(GetType(BdgOperacionVinoPlan).Name & "Destino", dtDestino)
    '        'lstPropuesta.Add(GetType(BdgOperacionPlanMaterial).Name, dtMaterialGlobal)
    '        'lstPropuesta.Add(GetType(BdgOperacionPlanMOD).Name, dtMODGlobal)
    '        'lstPropuesta.Add(GetType(BdgOperacionPlanCentro).Name, dtCentroGlobal)
    '        'lstPropuesta.Add(GetType(BdgOperacionPlanVarios).Name, dtVariosGlobal)
    '        lstPropuesta.Add(GetType(BdgOperacionVinoPlanMaterial).Name, dtMaterial)

    '        Return lstPropuesta
    '    End If

    'End Function

    '<Task()> Public Shared Function ActualizarNuevaOperacionPlanificada(ByVal lstPropuesta As Dictionary(Of String, DataTable), ByVal service As ServiceProvider) As StNuevaOperacionPlanificadaResult
    '    Dim rslt As New StNuevaOperacionPlanificadaResult(String.Empty)
    '    If Not lstPropuesta Is Nothing AndAlso lstPropuesta.Count > 0 Then

    '        Dim dtOperacion As DataTable
    '        Dim dtOperacionVinoPlan, dtOperacionVinoPlanOrigen, dtOperacionVinoPlanDestino As DataTable
    '        Dim dtMaterialGlobal, dtMaterial As DataTable
    '        Dim dtCentrosGlobal, dtCentros As DataTable
    '        Dim dtMODGlobal, dtMOD As DataTable
    '        Dim dtVariosGlobal, dtVarios As DataTable


    '        If lstPropuesta.ContainsKey(GetType(BdgOperacionPlan).Name) Then dtOperacion = lstPropuesta(GetType(BdgOperacionPlan).Name)

    '        If lstPropuesta.ContainsKey(GetType(BdgOperacionVinoPlan).Name) Then dtOperacionVinoPlan = lstPropuesta(GetType(BdgOperacionVinoPlan).Name)
    '        If lstPropuesta.ContainsKey(GetType(BdgOperacionVinoPlan).Name & "Origen") Then dtOperacionVinoPlanOrigen = lstPropuesta(GetType(BdgOperacionVinoPlan).Name & "Origen")
    '        If lstPropuesta.ContainsKey(GetType(BdgOperacionVinoPlan).Name & "Destino") Then dtOperacionVinoPlanDestino = lstPropuesta(GetType(BdgOperacionVinoPlan).Name & "Destino")

    '        If lstPropuesta.ContainsKey(GetType(BdgOperacionPlanMaterial).Name) Then dtMaterialGlobal = lstPropuesta(GetType(BdgOperacionPlanMaterial).Name)
    '        If lstPropuesta.ContainsKey(GetType(BdgOperacionPlanMOD).Name) Then dtMODGlobal = lstPropuesta(GetType(BdgOperacionPlanMOD).Name)
    '        If lstPropuesta.ContainsKey(GetType(BdgOperacionPlanCentro).Name) Then dtCentrosGlobal = lstPropuesta(GetType(BdgOperacionPlanCentro).Name)
    '        If lstPropuesta.ContainsKey(GetType(BdgOperacionPlanVarios).Name) Then dtVariosGlobal = lstPropuesta(GetType(BdgOperacionPlanVarios).Name)


    '        If lstPropuesta.ContainsKey(GetType(BdgOperacionVinoPlanMaterial).Name) Then dtMaterial = lstPropuesta(GetType(BdgOperacionVinoPlanMaterial).Name)
    '        If lstPropuesta.ContainsKey(GetType(BdgOperacionVinoPlanMOD).Name) Then dtMOD = lstPropuesta(GetType(BdgOperacionVinoPlanMOD).Name)
    '        If lstPropuesta.ContainsKey(GetType(BdgOperacionVinoPlanCentro).Name) Then dtCentros = lstPropuesta(GetType(BdgOperacionVinoPlanCentro).Name)
    '        If lstPropuesta.ContainsKey(GetType(BdgOperacionVinoPlanVarios).Name) Then dtVarios = lstPropuesta(GetType(BdgOperacionVinoPlanVarios).Name)

    '        If Not dtOperacion Is Nothing AndAlso dtOperacion.Rows.Count > 0 Then
    '            Try
    '                AdminData.BeginTx()
    '                Dim UpdtPckg As New UpdatePackage
    '                UpdtPckg.Add(dtOperacion)

    '                If dtOperacionVinoPlan Is Nothing OrElse dtOperacionVinoPlan.Rows.Count = 0 Then
    '                    If Not dtOperacionVinoPlanOrigen Is Nothing AndAlso dtOperacionVinoPlanOrigen.Rows.Count > 0 Then
    '                        dtOperacionVinoPlan = dtOperacionVinoPlanOrigen
    '                        If Not dtOperacionVinoPlanDestino Is Nothing AndAlso dtOperacionVinoPlanDestino.Rows.Count > 0 Then
    '                            For Each drDestino As DataRow In dtOperacionVinoPlanDestino.Rows
    '                                dtOperacionVinoPlan.ImportRow(drDestino)
    '                            Next
    '                        End If

    '                    ElseIf Not dtOperacionVinoPlanDestino Is Nothing AndAlso dtOperacionVinoPlanDestino.Rows.Count > 0 Then
    '                        dtOperacionVinoPlan = dtOperacionVinoPlanDestino
    '                    End If
    '                End If

    '                If Not dtOperacionVinoPlan Is Nothing AndAlso dtOperacionVinoPlan.Rows.Count > 0 Then
    '                    UpdtPckg.Add(GetType(BdgOperacionVinoPlan).Name, dtOperacionVinoPlan)
    '                End If


    '                If Not dtMaterialGlobal Is Nothing Then UpdtPckg.Add(GetType(BdgOperacionPlanMaterial).Name, dtMaterialGlobal)
    '                If Not dtMODGlobal Is Nothing Then UpdtPckg.Add(GetType(BdgOperacionPlanMOD).Name, dtMODGlobal)
    '                If Not dtCentrosGlobal Is Nothing Then UpdtPckg.Add(GetType(BdgOperacionPlanCentro).Name, dtCentrosGlobal)
    '                If Not dtVariosGlobal Is Nothing Then UpdtPckg.Add(GetType(BdgOperacionPlanVarios).Name, dtVariosGlobal)

    '                If Not dtMaterial Is Nothing Then UpdtPckg.Add(GetType(BdgOperacionVinoPlanMaterial).Name, dtMaterial)
    '                If Not dtMOD Is Nothing Then UpdtPckg.Add(GetType(BdgOperacionVinoPlanMOD).Name, dtMOD)
    '                If Not dtCentros Is Nothing Then UpdtPckg.Add(GetType(BdgOperacionVinoPlanCentro).Name, dtCentros)
    '                If Not dtVarios Is Nothing Then UpdtPckg.Add(GetType(BdgOperacionVinoPlanVarios).Name, dtVarios)

    '                Dim ClsBdgOperPlan As New BdgOperacionPlan
    '                ClsBdgOperPlan.Update(UpdtPckg)

    '                rslt.NOperacionPlan = dtOperacion.Rows(0)("NOperacionPlan")

    '            Catch ex As Exception
    '                rslt.Errores.Add(New ClassErrors(Nothing, ex.Message & "(ActualizarNuevaOperacionPlanificada)"))
    '                AdminData.RollBackTx()
    '            End Try
    '        End If
    '    Else
    '        rslt.Errores.Add(New ClassErrors(Nothing, "No hay datos para generar Operación planificada."))
    '    End If

    '    Return rslt
    'End Function


    ''//Esta se utilizaba en el diagrama antiguo. En principio ya no se va a utilizar
    '<Task()> Public Shared Function NuevaOperacionPlanificadaMultiple(ByVal data As StNuevaOperacionPlanificadaMultiple, ByVal service As ServiceProvider) As StNuevaOperacionPlanificadaMultipleResult
    '    Dim result As New StNuevaOperacionPlanificadaMultipleResult
    '    If (Length(data.IDTipoOperacion) = 0) Then Return result
    '    Dim dt As DataTable
    '    Dim dtOrigenes As DataTable
    '    Dim dtDestinos As DataTable
    '    If Not data.DtOrigen Is Nothing Then
    '        dt = data.DtOrigen
    '        dtOrigenes = data.DtOrigen.Clone
    '    ElseIf Not data.DtDestino Is Nothing Then
    '        dt = data.DtDestino
    '        dtDestinos = data.DtDestino.Clone
    '    End If

    '    'generaremos una cabecera y N lineas
    '    If data.UnaOperacion Then
    '        Dim StADD As New StADDCabeceraOperacion(data.IDTipoOperacion, data.FechaOperacion)
    '        Dim lstPropuesta As Dictionary(Of String, DataTable) = ProcessServer.ExecuteTask(Of StADDCabeceraOperacion, Dictionary(Of String, DataTable))(AddressOf ADDCabeceraOperacion, StADD, service)
    '        If lstPropuesta Is Nothing OrElse lstPropuesta.Count = 0 Then Exit Function

    '        Dim dtOperacion As DataTable
    '        Dim dtMaterialGlobal As DataTable
    '        Dim dtMODGlobal As DataTable
    '        Dim dtCentroGlobal As DataTable
    '        Dim dtVariosGlobal As DataTable


    '        If lstPropuesta.ContainsKey(GetType(BdgOperacionPlan).Name) Then dtOperacion = lstPropuesta(GetType(BdgOperacionPlan).Name)

    '        If lstPropuesta.ContainsKey(GetType(BdgOperacionMaterial).Name) Then dtMaterialGlobal = lstPropuesta(GetType(BdgOperacionMaterial).Name)
    '        If lstPropuesta.ContainsKey(GetType(BdgOperacionMOD).Name) Then dtMODGlobal = lstPropuesta(GetType(BdgOperacionMOD).Name)
    '        If lstPropuesta.ContainsKey(GetType(BdgOperacionCentro).Name) Then dtCentroGlobal = lstPropuesta(GetType(BdgOperacionCentro).Name)
    '        If lstPropuesta.ContainsKey(GetType(BdgOperacionVarios).Name) Then dtVariosGlobal = lstPropuesta(GetType(BdgOperacionVarios).Name)

    '        Dim dtOrigen, dtDestino, dtMaterial As DataTable
    '        Dim drTipoOperacion As DataRow = New BdgTipoOperacion().GetItemRow(data.IDTipoOperacion)

    '        'orígenes
    '        Dim nuevasLineasData As New StAddNuevasOperacionesVinoPlan(dtOperacion, data.DtOrigen, dtMaterial, dtOrigen, dtDestino, False)
    '        Dim nuevasLineasDataResult As StAddNuevasOperacionesVinoPlan = ProcessServer.ExecuteTask(Of StAddNuevasOperacionesVinoPlan, StAddNuevasOperacionesVinoPlan)(AddressOf AddNuevasOperacionesVinoPlan, nuevasLineasData, service)

    '        'aquí no tenemos que hacer los materiales y demás
    '        If Not nuevasLineasDataResult.DtMaterialResult Is Nothing Then nuevasLineasDataResult.DtMaterialResult.Clear()

    '        'destinos
    '        nuevasLineasData = New StAddNuevasOperacionesVinoPlan(dtOperacion, data.DtDestino, nuevasLineasDataResult.DtMaterialResult, nuevasLineasDataResult.DtOrigenResult, nuevasLineasDataResult.DtDestinoResult, True)
    '        nuevasLineasDataResult = ProcessServer.ExecuteTask(Of StAddNuevasOperacionesVinoPlan, StAddNuevasOperacionesVinoPlan)(AddressOf AddNuevasOperacionesVinoPlan, nuevasLineasData, service)
    '        If Not nuevasLineasData.DtMaterialResult Is Nothing Then dtMaterial = nuevasLineasData.DtMaterialResult.Copy
    '        If Not nuevasLineasData.DtOrigenResult Is Nothing Then dtOrigen = nuevasLineasData.DtOrigenResult.Copy
    '        If Not nuevasLineasData.DtDestinoResult Is Nothing Then dtDestino = nuevasLineasData.DtDestinoResult.Copy

    '        Dim UpdtPckg As New UpdatePackage
    '        UpdtPckg.Add(dtOperacion)
    '        If Not dtOrigen Is Nothing Then UpdtPckg.Add("BdgOperacionVinoPlan", dtOrigen)
    '        If Not dtDestino Is Nothing Then UpdtPckg.Add("BdgOperacionVinoPlan", dtDestino)
    '        If Not dtMaterial Is Nothing Then UpdtPckg.Add("BdgOperacionVinoPlanMaterial", dtMaterial)
    '        If Not dtMaterialGlobal Is Nothing Then UpdtPckg.Add("BdgOperacionPlanMaterial", dtMaterialGlobal)
    '        If Not dtMODGlobal Is Nothing Then UpdtPckg.Add("BdgOperacionPlanMOD", dtMODGlobal)
    '        If Not dtCentroGlobal Is Nothing Then UpdtPckg.Add("BdgOperacionPlanCentro", dtCentroGlobal)
    '        If Not dtVariosGlobal Is Nothing Then UpdtPckg.Add("BdgOperacionPlanVarios", dtVariosGlobal)

    '        Dim ClsBdgOperPlan As New BdgOperacionPlan
    '        ClsBdgOperPlan.Update(UpdtPckg)


    '        ReDim Preserve result.NOperacion(UBound(result.NOperacion) + 1)
    '        result.NOperacion(UBound(result.NOperacion)) = dtOperacion.Rows(0)("NOperacionPlan")
    '    Else
    '        For Each dtrLinea As DataRow In dt.Select
    '            Dim currentParam As StNuevaOperacionPlanificada
    '            If (dt.Columns.Contains("IDVino")) Then
    '                currentParam = New StNuevaOperacionPlanificada(dtrLinea("IDDeposito"), dtrLinea("IDVino"), data.IDTipoOperacion, data.FechaOperacion, data.DtDestino)
    '            Else : currentParam = New StNuevaOperacionPlanificada(dtrLinea("IDDeposito"), Nothing, data.IDTipoOperacion, data.FechaOperacion, data.DtDestino)
    '            End If
    '            Dim partialResult As StNuevaOperacionPlanificadaResult = ProcessServer.ExecuteTask(Of StNuevaOperacionPlanificada, StNuevaOperacionPlanificadaResult)(AddressOf NuevaOperacionPlanificada, currentParam, service)
    '            ReDim Preserve result.NOperacion(UBound(result.NOperacion) + 1)
    '            result.NOperacion(UBound(result.NOperacion)) = partialResult.NOperacionPlan
    '        Next
    '    End If
    '    Return result
    'End Function

#Region "Tareas Secundarias"

    <Serializable()> _
  Public Class StADDCabeceraOperacion
        Public IDTipoOperacion As String
        Public FechaOperacion As Date

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDTipoOperacion As String, Optional ByVal FechaOperacion As Date = cnMinDate)
            Me.IDTipoOperacion = IDTipoOperacion
            Me.FechaOperacion = FechaOperacion
        End Sub
    End Class

    <Task()> Public Shared Function ADDCabeceraOperacion(ByVal data As StADDCabeceraOperacion, ByVal services As ServiceProvider) As Dictionary(Of String, DataTable)
        Dim OpPlan As New BdgOperacionPlan
        Dim dtOperacion As DataTable = OpPlan.AddNewForm
        If data.FechaOperacion <> cnMinDate Then
            dtOperacion.Rows(0)("Fecha") = New Date(data.FechaOperacion.Year, data.FechaOperacion.Month, data.FechaOperacion.Day, Today.Now.Hour, Today.Now.Minute, Today.Now.Second)
        End If
        dtOperacion.Rows(0)("IDTipoOperacion") = data.IDTipoOperacion
        Dim context As New BusinessData
        context("OperacionNew") = True
        dtOperacion.Rows(0).ItemArray = OpPlan.ApplyBusinessRule("IDTipoOperacion", data.IDTipoOperacion, dtOperacion.Rows(0), context).ItemArray

        Dim StDataImp As New BdgGeneral.StImputacion(dtOperacion.Rows(0)("IDTipoOperacion"), dtOperacion.Rows(0)("NOperacionPlan"), enumBdgOrigenOperacion.Planificada, True, dtOperacion.Rows(0)("Fecha"))
        StDataImp = ProcessServer.ExecuteTask(Of BdgGeneral.StImputacion, BdgGeneral.StImputacion)(AddressOf BdgGeneral.Imputaciones, StDataImp, services)

        Dim lstDatosTipoOperacion As New Dictionary(Of String, DataTable)
        lstDatosTipoOperacion.Add(GetType(BdgOperacionPlan).Name, dtOperacion)

        If Not StDataImp.DtMaterialesGlobal Is Nothing AndAlso StDataImp.DtMaterialesGlobal.Rows.Count > 0 Then lstDatosTipoOperacion.Add(GetType(BdgOperacionPlanMaterial).Name, StDataImp.DtMaterialesGlobal)
        If Not StDataImp.DtCentrosGlobal Is Nothing AndAlso StDataImp.DtCentrosGlobal.Rows.Count > 0 Then lstDatosTipoOperacion.Add(GetType(BdgOperacionPlanCentro).Name, StDataImp.DtCentrosGlobal)
        If Not StDataImp.DtMODGlobal Is Nothing AndAlso StDataImp.DtMODGlobal.Rows.Count > 0 Then lstDatosTipoOperacion.Add(GetType(BdgOperacionPlanMOD).Name, StDataImp.DtMODGlobal)
        If Not StDataImp.DtVariosGlobal Is Nothing AndAlso StDataImp.DtVariosGlobal.Rows.Count > 0 Then lstDatosTipoOperacion.Add(GetType(BdgOperacionPlanVarios).Name, StDataImp.DtVariosGlobal)

        'If Not StDataImp.DtMateriales Is Nothing AndAlso StDataImp.DtMateriales.Rows.Count > 0 Then lstDatosTipoOperacion.Add(GetType(BdgOperacionVinoPlanMaterial).Name, StDataImp.DtMateriales)
        'If Not StDataImp.DtCentros Is Nothing AndAlso StDataImp.DtCentros.Rows.Count > 0 Then lstDatosTipoOperacion.Add(GetType(BdgOperacionVinoPlanCentro).Name, StDataImp.DtCentros)
        'If Not StDataImp.DtMOD Is Nothing AndAlso StDataImp.DtMOD.Rows.Count > 0 Then lstDatosTipoOperacion.Add(GetType(BdgOperacionVinoPlanMOD).Name, StDataImp.DtMOD)
        'If Not StDataImp.DtVarios Is Nothing AndAlso StDataImp.DtVarios.Rows.Count > 0 Then lstDatosTipoOperacion.Add(GetType(BdgOperacionVinoPlanVarios).Name, StDataImp.DtVarios)

        Return lstDatosTipoOperacion
    End Function

    '<Serializable()> Public Class StAddNuevasOperacionesVinoPlan
    '    Public DtOperacion As DataTable
    '    Public DtOrigen As DataTable
    '    Public EsDestino As Boolean

    '    Public DtMaterialResult As DataTable
    '    Public DtOrigenResult As DataTable
    '    Public DtDestinoResult As DataTable

    '    Public Sub New(ByVal dtOperacion As DataTable, ByVal dtOrigen As DataTable, Optional ByVal DtMaterialResult As DataTable = Nothing, _
    '                    Optional ByVal dtOrigenResult As DataTable = Nothing, Optional ByVal dtDestinoResult As DataTable = Nothing, _
    '                    Optional ByVal blnEsDestino As Boolean = True)
    '        Me.DtOperacion = dtOperacion
    '        Me.DtOrigen = dtOrigen
    '        Me.DtMaterialResult = DtMaterialResult
    '        Me.DtOrigenResult = dtOrigenResult
    '        Me.DtDestinoResult = dtDestinoResult

    '        Me.EsDestino = blnEsDestino
    '    End Sub

    'End Class

    '<Task()> Public Shared Function AddNuevasOperacionesVinoPlan(ByVal data As StAddNuevasOperacionesVinoPlan, ByVal service As ServiceProvider) As StAddNuevasOperacionesVinoPlan
    '    If Not data.DtOrigen Is Nothing AndAlso data.DtOrigen.Rows.Count > 0 Then
    '        For Each dtrLinea As DataRow In data.DtOrigen.Select
    '            Dim nuevaLineaData As StAddNuevaOperacionVinoPlan
    '            If dtrLinea.Table.Columns.Contains("IDVino") Then
    '                nuevaLineaData = New StAddNuevaOperacionVinoPlan(data.DtOperacion, dtrLinea("IDDeposito"), dtrLinea("IDVino"), Nothing, Nothing, Nothing, data.EsDestino)
    '            Else : nuevaLineaData = New StAddNuevaOperacionVinoPlan(data.DtOperacion, dtrLinea("IDDeposito"), Nothing, Nothing, Nothing, Nothing, data.EsDestino)
    '            End If
    '            Dim nuevaLineaResultData As StAddNuevaOperacionVinoPlan = ProcessServer.ExecuteTask(Of StAddNuevaOperacionVinoPlan, StAddNuevaOperacionVinoPlan)(AddressOf NuevaOperacionVinoPlan, nuevaLineaData, service)
    '            If Not nuevaLineaData.dtOrigen Is Nothing Then
    '                If data.DtOrigenResult Is Nothing Then
    '                    data.DtOrigenResult = nuevaLineaData.dtOrigen.Copy
    '                Else
    '                    For Each dtr As DataRow In nuevaLineaData.dtOrigen.Rows
    '                        dtr("Destino") = False
    '                        data.DtOrigenResult.ImportRow(dtr)
    '                    Next
    '                End If
    '            End If
    '            If Not nuevaLineaData.dtDestino Is Nothing Then
    '                If (data.DtDestinoResult Is Nothing) Then
    '                    data.DtDestinoResult = nuevaLineaData.dtDestino.Copy
    '                Else
    '                    For Each dtr As DataRow In nuevaLineaData.dtDestino.Rows
    '                        dtr("Destino") = True
    '                        data.DtDestinoResult.ImportRow(dtr)
    '                    Next
    '                End If
    '            End If
    '            If Not nuevaLineaData.dtMaterial Is Nothing Then
    '                If data.DtMaterialResult Is Nothing Then
    '                    data.DtMaterialResult = nuevaLineaData.dtMaterial.Copy
    '                Else
    '                    For Each dtr As DataRow In nuevaLineaData.dtMaterial.Rows
    '                        data.DtMaterialResult.ImportRow(dtr)
    '                    Next
    '                End If
    '            End If
    '        Next
    '    End If
    '    Return data
    'End Function

    <Serializable()> _
   Public Class StADDOrigenDestino
        Public IDDeposito As String
        Public NOperacionPlan As String
        Public IDVino As Guid '?? dudo.
        Public Cantidad As Double
        Public DrDestinoExterno As DataTable
        Public dtOperacion As DataTable


        Public Sub New()
        End Sub

        Public Sub New(ByVal IDDeposito As String, ByVal NOperacionPlan As String, ByVal IDVino As Guid, Optional ByVal Cantidad As Double = 0, Optional ByVal DrDestinoExterno As DataTable = Nothing)
            Me.IDDeposito = IDDeposito
            Me.NOperacionPlan = NOperacionPlan
            Me.IDVino = IDVino
            Me.Cantidad = Cantidad
            Me.DrDestinoExterno = DrDestinoExterno
        End Sub
    End Class

    <Task()> Public Shared Function ADDOrigen(ByVal data As StADDOrigenDestino, ByVal services As ServiceProvider) As DataTable
        Dim dtOrigen As DataTable = AdminData.GetData("frmBdgOperacionVinoOrigenPlan", New NoRowsFilterItem)
        Dim drOrigen As DataRow = dtOrigen.NewRow
        drOrigen("IDLineaOperacionVinoPlan") = Guid.NewGuid
        drOrigen("IDDeposito") = data.IDDeposito
        drOrigen("NOperacionPlan") = data.NOperacionPlan


        Dim OpVino As New BdgOperacionVinoPlan
        Dim f As New Filter
        f.Add(New StringFilterItem("IDDeposito", data.IDDeposito))
        If (data.IDVino <> Guid.Empty) Then f.Add(New GuidFilterItem("IDVino", data.IDVino))
        Dim dtArticulo As DataTable = New BE.DataEngine().Filter("NegBdgOperacionPlan", f, , "IDArticulo, Lote")
        If Not dtArticulo Is Nothing AndAlso dtArticulo.Rows.Count = 1 Then

            Dim context As New BusinessData
            If data.dtOperacion Is Nothing OrElse data.dtOperacion.Rows.Count = 0 Then
                data.dtOperacion = New BdgOperacionPlan().SelOnPrimaryKey(data.NOperacionPlan)
            End If
            If Not data.dtOperacion Is Nothing AndAlso data.dtOperacion.Rows.Count > 0 Then
                context = New BusinessData(data.dtOperacion.Rows(0))
            End If

            drOrigen("IDArticulo") = dtArticulo.Rows(0)("IDArticulo")
            drOrigen = OpVino.ApplyBusinessRule("IDArticulo", dtArticulo.Rows(0)("IDArticulo"), drOrigen, context)

            drOrigen("IDUdMedida") = dtArticulo.Rows(0)("IDUdMedida")
            drOrigen("IDBarrica") = dtArticulo.Rows(0)("IDBarrica")
            drOrigen("IDEstadoVino") = dtArticulo.Rows(0)("IDEstadoVino")
            drOrigen("Lote") = dtArticulo.Rows(0)("Lote")
            drOrigen("Ocupacion") = dtArticulo.Rows(0)("Ocupacion")
            drOrigen("TipoDeposito") = dtArticulo.Rows(0)("TipoDeposito")
            drOrigen("Destino") = False


            drOrigen("Cantidad") = IIf(data.Cantidad <> 0, data.Cantidad, xRound(dtArticulo.Rows(0)("Cantidad"), 1))
            drOrigen = OpVino.ApplyBusinessRule("Cantidad", drOrigen("Cantidad"), drOrigen, context)

            'Dim dataUnidad As New UnidadAB.UnidadMedidaInfo
            'dataUnidad.IDUdMedidaA = drOrigen("IDUdMedida")
            'dataUnidad.IDUdMedidaB = New BdgParametro().UnidadesCampoLitros()
            'Dim dblFactor As Double = ProcessServer.ExecuteTask(Of UnidadAB.UnidadMedidaInfo, Double)(AddressOf UnidadAB.FactorDeConversion, dataUnidad, services)
            'drOrigen("Litros") = drOrigen("Cantidad") * dblFactor

            'Dim StGet1 As New BdgOperacion.StGetQDeposito(drOrigen("IDArticulo"), drOrigen("IDDeposito"), drOrigen("IDUdMedida"), drOrigen("Cantidad"))
            'drOrigen("QDeposito") = xRound(ProcessServer.ExecuteTask(Of BdgOperacion.StGetQDeposito, Double)(AddressOf BdgOperacion.GetQDeposito, StGet1, services), 1)
            Dim StGet2 As New BdgGeneral.StGetQDeposito(drOrigen("IDArticulo"), drOrigen("IDDeposito"), drOrigen("IDUdMedida"), drOrigen("Ocupacion"))
            drOrigen("Ocupacion") = xRound(ProcessServer.ExecuteTask(Of BdgGeneral.StGetQDeposito, Double)(AddressOf BdgGeneral.GetQDeposito, StGet2, services), 1)
        End If
        dtOrigen.Rows.Add(drOrigen.ItemArray)
        Return dtOrigen
    End Function

    <Task()> Public Shared Function ADDDestino(ByVal data As StADDOrigenDestino, ByVal services As ServiceProvider) As DataTable
        Dim dtDestino As DataTable = AdminData.GetData("frmBdgOperacionVinoDestinoPlan", New NoRowsFilterItem)
        Dim drDestino As DataRow = dtDestino.NewRow
        drDestino("IDLineaOperacionVinoPlan") = Guid.NewGuid
        drDestino("IDDeposito") = data.IDDeposito
        drDestino("NOperacionPlan") = data.NOperacionPlan
        drDestino("Destino") = True
        Dim f As New Filter
        f.Add(New StringFilterItem("IDDeposito", data.IDDeposito))
        If (data.IDVino <> Guid.Empty) Then
            f.Add(New GuidFilterItem("IDVino", data.IDVino))
        End If

        Dim OpVino As New BdgOperacionVinoPlan
        Dim dtArticulo As DataTable
        If data.Cantidad = 0 Then
            dtArticulo = AdminData.GetData("NegBdgOperacionPlan", f, , "IDArticulo, Lote")
        Else
            dtArticulo = data.DrDestinoExterno.Clone
            dtArticulo.Rows.Add(data.DrDestinoExterno.Rows(0).ItemArray)
        End If
        If Not dtArticulo Is Nothing AndAlso dtArticulo.Rows.Count = 1 Then

            Dim context As New BusinessData
            If data.dtOperacion Is Nothing OrElse data.dtOperacion.Rows.Count = 0 Then
                data.dtOperacion = New BdgOperacionPlan().SelOnPrimaryKey(data.NOperacionPlan)
            End If
            If Not data.dtOperacion Is Nothing AndAlso data.dtOperacion.Rows.Count > 0 Then
                context = New BusinessData(data.dtOperacion.Rows(0))
            End If

            drDestino("IDArticulo") = dtArticulo.Rows(0)("IDArticulo")
            drDestino = OpVino.ApplyBusinessRule("IDArticulo", dtArticulo.Rows(0)("IDArticulo"), drDestino, context)
            ' drDestino("DescArticulo") = dtArticulo.Rows(0)("DescArticulo")
            drDestino("IDUdMedida") = dtArticulo.Rows(0)("IDUdMedida")

            drDestino("Lote") = dtArticulo.Rows(0)("Lote")
            drDestino("Ocupacion") = Nz(dtArticulo.Rows(0)("Ocupacion"), 0)
            drDestino("Capacidad") = Nz(dtArticulo.Rows(0)("Capacidad"), 0)
            drDestino("TipoDeposito") = dtArticulo.Rows(0)("TipoDeposito")


            drDestino("Cantidad") = xRound(dtArticulo.Rows(0)("Cantidad"), 1)
            drDestino = OpVino.ApplyBusinessRule("Cantidad", drDestino("Cantidad"), drDestino, context)

            Dim StGet2 As New BdgGeneral.StGetQDeposito(drDestino("IDArticulo"), drDestino("IDDeposito"), drDestino("IDUdMedida"), drDestino("Ocupacion"))
            drDestino("Ocupacion") = xRound(ProcessServer.ExecuteTask(Of BdgGeneral.StGetQDeposito, Double)(AddressOf BdgGeneral.GetQDeposito, StGet2, services), 1)


            'Dim dataUnidad As New UnidadAB.UnidadMedidaInfo
            'dataUnidad.IDUdMedidaA = drDestino("IDUdMedida")
            'dataUnidad.IDUdMedidaB = New BdgParametro().UnidadesCampoLitros()
            'Dim dblFactor As Double = ProcessServer.ExecuteTask(Of UnidadAB.UnidadMedidaInfo, Double)(AddressOf UnidadAB.FactorDeConversion, dataUnidad, services)

            'drDestino("Litros") = drDestino("Cantidad") * dblFactor
            'Dim StGet1 As New BdgOperacion.StGetQDeposito(drDestino("IDArticulo"), drDestino("IDDeposito"), drDestino("IDUdMedida"), drDestino("Cantidad"))
            'drDestino("QDeposito") = xRound(ProcessServer.ExecuteTask(Of BdgOperacion.StGetQDeposito, Double)(AddressOf BdgOperacion.GetQDeposito, StGet1, services), 1)
            'Dim StGet2 As New BdgOperacion.StGetQDeposito(drDestino("IDArticulo"), drDestino("IDDeposito"), drDestino("IDUdMedida"), drDestino("Ocupacion"))
            'drDestino("Ocupacion") = xRound(ProcessServer.ExecuteTask(Of BdgOperacion.StGetQDeposito, Double)(AddressOf BdgOperacion.GetQDeposito, StGet2, services), 1)
        ElseIf Not dtArticulo Is Nothing AndAlso dtArticulo.Rows.Count > 1 Then
            ApplicationService.GenerateError("El Depósito seleccionado tiene más de un Vino. Proceso cancelado.")
        End If
        dtDestino.Rows.Add(drDestino.ItemArray)
        Return dtDestino
    End Function

    <Serializable()> _
    Public Class StAddNuevaOperacionVinoPlan
        Public dtCabecera As DataTable
        Public TipoMovimiento As TipoMovimiento?
        Public IDDeposito As String
        Public IDVino As Guid

        Public ForzarDestino As Boolean

        'Aquí podrán venir, o no, elementos
        Public dtOrigen As DataTable
        Public dtDestino As DataTable
        Public dtMaterial As DataTable

        Public cantidad As Double

        Public DrDestinoExterno As DataTable

        Public Sub New(ByVal dtOperacionCabecera As DataTable, ByVal IDDeposito As String, ByVal IDVino As Guid, Optional ByVal dtOperacionVinoOrigen As DataTable = Nothing, _
                       Optional ByVal dtOperacionVinoDestino As DataTable = Nothing, Optional ByVal dtOperacionVinoPlanMaterial As DataTable = Nothing, _
                       Optional ByVal forzarDestino As Boolean = False, Optional ByVal Cantidad As Double = 0, Optional ByVal DrDestinoExterno As DataTable = Nothing)

            Me.dtCabecera = dtOperacionCabecera

            Me.IDDeposito = IDDeposito
            Me.IDVino = IDVino

            Me.dtOrigen = dtOperacionVinoOrigen
            Me.dtDestino = dtOperacionVinoDestino
            Me.dtMaterial = dtOperacionVinoPlanMaterial
            Me.ForzarDestino = forzarDestino
            Me.cantidad = Cantidad

            Me.DrDestinoExterno = DrDestinoExterno
        End Sub
    End Class

    <Task()> Public Shared Function NuevaOperacionVinoPlan(ByVal data As StAddNuevaOperacionVinoPlan, ByVal service As ServiceProvider) As StAddNuevaOperacionVinoPlan
        If Not data.dtCabecera Is Nothing AndAlso data.dtCabecera.Rows.Count > 0 Then
            If data.TipoMovimiento Is Nothing Then
                If Length(data.dtCabecera.Rows(0)("IDTipoOperacion")) > 0 Then
                    Dim dtrOperacion As DataTable = New BdgTipoOperacion().SelOnPrimaryKey(data.dtCabecera.Rows(0)("IDTipoOperacion"))
                    If dtrOperacion.Rows.Count > 0 Then
                        data.TipoMovimiento = Nz(dtrOperacion.Rows(0)("TipoMovimiento"), -1)
                    End If
                End If
                If data.TipoMovimiento Is Nothing Then data.TipoMovimiento = -1
            End If
            Dim strNOperacion As String = data.dtCabecera.Rows(0)("NOperacionPlan") & String.Empty
            Dim esDestino As Boolean = data.ForzarDestino OrElse (data.TipoMovimiento = TipoMovimiento.SinMovimiento OrElse data.TipoMovimiento = TipoMovimiento.SinOrigen)

            Dim dtResultados As DataTable
            Dim StDes As New StADDOrigenDestino(data.IDDeposito, strNOperacion, data.IDVino, data.cantidad, data.DrDestinoExterno)
            StDes.dtOperacion = data.dtCabecera
            If esDestino Then
                dtResultados = ProcessServer.ExecuteTask(Of StADDOrigenDestino, DataTable)(AddressOf ADDDestino, StDes, service)
            Else
                dtResultados = ProcessServer.ExecuteTask(Of StADDOrigenDestino, DataTable)(AddressOf ADDOrigen, StDes, service)
            End If

            If Not dtResultados Is Nothing AndAlso dtResultados.Rows.Count > 0 Then
                If (Not esDestino) Then
                    If (data.dtOrigen Is Nothing OrElse data.dtOrigen.Rows.Count = 0) Then
                        data.dtOrigen = dtResultados.Copy
                    ElseIf (dtResultados.Rows.Count = 1) Then
                        data.dtOrigen.ImportRow(dtResultados.Rows(0))
                    End If
                Else
                    If (data.dtDestino Is Nothing OrElse data.dtDestino.Rows.Count = 0) Then
                        data.dtDestino = dtResultados.Copy
                    ElseIf (dtResultados.Rows.Count = 1) Then
                        data.dtDestino.ImportRow(dtResultados.Rows(0))
                    End If
                End If
                Dim stmat As New BdgGeneral.StImputacion(data.dtCabecera.Rows(0)("IDTipoOperacion"), strNOperacion, enumBdgOrigenOperacion.Planificada, False)
                stmat.LineaDestino.Add(New BdgGeneral.StDestino(dtResultados.Rows(0)("IDLineaOperacionVinoPlan"), dtResultados.Rows(0)("IDArticulo") & String.Empty, dtResultados.Rows(0)("IDEstructura") & String.Empty, Nz(dtResultados.Rows(0)("Litros"), 0), Nz(dtResultados.Rows(0)("Cantidad"), 0)))
                Dim dtNuevoMaterial As DataTable = ProcessServer.ExecuteTask(Of BdgGeneral.StImputacion, DataTable)(AddressOf BdgGeneral.ImputacionLineaMateriales, stmat, service)
                If (Not dtNuevoMaterial Is Nothing AndAlso (data.dtMaterial Is Nothing OrElse data.dtMaterial.Rows.Count = 0)) Then
                    data.dtMaterial = dtNuevoMaterial.Copy
                ElseIf (Not dtNuevoMaterial Is Nothing AndAlso dtNuevoMaterial.Rows.Count = 1) Then
                    data.dtMaterial.ImportRow(dtNuevoMaterial.Rows(0))
                End If
            Else : Return data
            End If
        End If
        Return data
    End Function

#End Region

#End Region

#End Region

#Region "Tareas Públicas"

    <Serializable()> _
    Public Class stCambioEstadoPlan
        Public Operaciones As DataTable
        Public NuevoEstado As BdgEstadoOperacionPlan
        Public OperacionesCorrectas(-1) As String
        Public OperacionesIncorrectas(-1) As String
    End Class

    Public Shared Function CambioEstadoPlan(ByVal data As stCambioEstadoPlan, ByVal services As ServiceProvider) As stCambioEstadoPlan
        If data.NuevoEstado <> BdgEstadoOperacionPlan.Confirmado Then
            Try
                For Each dtr As DataRow In data.Operaciones.Select
                    If (dtr("Estado") <> BdgEstadoOperacionPlan.Confirmado) Then
                        dtr("Estado") = data.NuevoEstado
                        ReDim Preserve data.OperacionesCorrectas(UBound(data.OperacionesCorrectas) + 1)
                        data.OperacionesCorrectas(UBound(data.OperacionesCorrectas)) = dtr("NOperacionPlan")
                    Else
                        ReDim Preserve data.OperacionesIncorrectas(UBound(data.OperacionesIncorrectas) + 1)
                        data.OperacionesIncorrectas(UBound(data.OperacionesIncorrectas)) = dtr("NOperacionPlan")
                    End If
                Next
                Dim ClsBdgOperPlan As New BdgOperacionPlan
                ClsBdgOperPlan.Update(data.Operaciones)
            Catch ex As Exception
                For Each dtr As DataRow In data.Operaciones.Select
                    ReDim Preserve data.OperacionesIncorrectas(UBound(data.OperacionesIncorrectas) + 1)
                    data.OperacionesIncorrectas(UBound(data.OperacionesIncorrectas)) = dtr("NOperacionPlan")
                Next
                ApplicationService.GenerateError("Ha sucedido un error en el cambio de estado de las operaciones planificadas: |", ex.Message)
            End Try
        Else : ApplicationService.GenerateError("No es posible cambiar el estado de la operación planificada a Confirmada mediante este proceso.")
        End If
        Return data
    End Function

    <Serializable()> Public Class StCopiarRegistros
        Public TablaOrigen As DataTable
        Public TablaDestino As DataTable
        Public Entidad As BusinessHelper
        Public Clave As String
        Public Sub New(ByVal tablaOrigen As DataTable, ByVal tablaDestino As DataTable)
            Me.TablaDestino = tablaDestino
            Me.TablaOrigen = tablaOrigen
        End Sub
        Public Sub New(ByVal entidad As BusinessHelper, ByVal clave As String, ByVal tablaOrigen As DataTable) ', ByVal tablaDestino As DataTable)
            Me.Entidad = entidad
            Me.Clave = clave
            Me.TablaOrigen = tablaOrigen
            'Me.TablaDestino = tablaDestino
        End Sub

    End Class

    <Task()> Public Shared Function CopiarRegistros(ByVal data As StCopiarRegistros, ByVal services As ServiceProvider) As StCopiarRegistros
        If (Not data.Entidad Is Nothing) Then
            data.TablaDestino = data.Entidad.AddNew()
        End If
        If (data.TablaDestino Is Nothing) Then
            data.TablaDestino = data.TablaOrigen.Clone
        End If

        For Each dtrOrigen As DataRow In data.TablaOrigen.Rows
            Dim dtrNewRow As DataRow = data.TablaDestino.NewRow
            For Each dtcColumn As DataColumn In data.TablaOrigen.Columns
                If (data.TablaOrigen.Columns.Contains(dtcColumn.ColumnName) AndAlso _
                    data.TablaDestino.Columns.Contains(dtcColumn.ColumnName) AndAlso _
                    Length(dtrOrigen(dtcColumn.ColumnName))) Then
                    dtrNewRow(dtcColumn.ColumnName) = dtrOrigen(dtcColumn.ColumnName)
                End If
            Next
            If Length(data.Clave) > 0 Then dtrNewRow(data.Clave) = Guid.NewGuid
            data.TablaDestino.Rows.Add(dtrNewRow.ItemArray)
        Next
        Return data
    End Function


    <Serializable()> Public Class StCambiarEstadoOperacion
        Public NOperacionPlan As String
        Public Estado As BdgEstadoOperacionPlan

        Public Sub New(ByVal strNOperacionPlan As String, ByVal enumEstado As BdgEstadoOperacionPlan)
            NOperacionPlan = strNOperacionPlan
            Estado = enumEstado
        End Sub
    End Class

    <Task()> Public Shared Sub CambiarEstadoOperacion(ByVal data As StCambiarEstadoOperacion, ByVal service As ServiceProvider)
        Dim DtOperacionPlan As DataTable = New BdgOperacionPlan().SelOnPrimaryKey(data.NOperacionPlan)
        If Not DtOperacionPlan Is Nothing AndAlso DtOperacionPlan.Rows.Count > 0 Then
            DtOperacionPlan.Rows(0)("Estado") = data.Estado
            BusinessHelper.UpdateTable(DtOperacionPlan)
        End If
    End Sub

#End Region

End Class

Public Class BdgOperacionPlanInfo
    Inherits ClassEntityInfo

    Public NOperacionPlan As String
    Public IDContador As String
    Public IDTipoOperacion As String
    Public Fecha As String
    Public IDAnalisis As String
    Public Texto As String
    Public ImputacionGlobalMaterial As String
    Public ImputacionGlobalCentro As String
    Public ImputacionGlobalMod As String
    Public ImputacionGlobalVarios As String
    Public Estado As String

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal data As DataRow)
        MyBase.New(data)
    End Sub

    Public Sub New(ByVal NOperacionPlan As String)
        MyBase.New()
        Me.Fill(NOperacionPlan)
    End Sub

    Public Overloads Overrides Sub Fill(ByVal ParamArray PrimaryKey() As Object)
        Dim dttBOP As DataTable = New BdgOperacionPlan().SelOnPrimaryKey(PrimaryKey)
        If dttBOP.Rows.Count > 0 Then
            Me.Fill(dttBOP.Rows(0))
        Else
            ApplicationService.GenerateError("La operación planificada | no existe.", Quoted(PrimaryKey(0)))
        End If
    End Sub
End Class

Public Enum BdgEstadoOperacionPlan
    Planificado
    Confirmado
    Anulado
End Enum