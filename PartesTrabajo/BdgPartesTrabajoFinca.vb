Public Class BdgPartesTrabajoFinca
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgPartesTrabajoFinca"

#Region " RegisterDeleteTasks "

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf BorrarObraModControl)
    End Sub

    <Task()> Public Shared Sub BorrarObraModControl(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim OCC As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraMODControl")
        Dim OC As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraCabecera")
        Dim dt As DataTable = OCC.Filter(New GuidFilterItem("IDPartesTrabajoFinca", data("IDPartesTrabajoFinca")))
        If dt.Rows.Count > 0 Then
            '//AGOSTO: Modificaciones para que recalcule la obra al borrar partes de la finca
            Dim dtObra As DataTable
            For Each dr As DataRow In dt.Rows
                Dim IDObra As Integer = dr("IDObra")
                dtObra = OC.SelOnPrimaryKey(IDObra)
                dr("IDPartesTrabajoFinca") = DBNull.Value
            Next

            OCC.Delete(dt)
            Dim pck As New UpdatePackage(dtObra)
            pck.DeletedData("ObraMODControl") = dt
            OC.Update(pck)
            ProcessServer.ExecuteTask(Of Integer)(AddressOf ObraCabecera.RecalcularObra, dtObra.Rows(0)("IDObra"), services)
        End If
    End Sub

#End Region

#Region " GetBusinessRules "

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBrl As New BusinessRules
        oBrl.Add("IDOperario", AddressOf CambioIDOperario)
        oBrl.Add("IDObjetivoCampo", AddressOf CambioIDObjetivoCampo)
        oBrl.Add("HorasTrabajadas", AddressOf CambioHorasTrabajadas)
        oBrl.Add("HorasExtra", AddressOf CambioHorasExtra)
        oBrl.Add("CantidadTrabajada", AddressOf CambioCantidad)
        oBrl.Add("Gasto", AddressOf CambioGastoBono)
        oBrl.Add("Bono", AddressOf CambioGastoBono)

        Return oBrl
    End Function

    <Task()> Public Shared Sub CambioIDOperario(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim d As dataTasas = ProcessServer.ExecuteTask(Of String, dataTasas)(AddressOf GetTasasOperario, data.Value, services)
            data.Current("TasaHoraNormal") = d.TasaHoraNormal
            data.Current("TasaHoraExtra") = d.TasaHoraExtra
        End If
    End Sub

    <Task()> Public Shared Sub CambioIDObjetivoCampo(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDObjetivoCampo")) > 0 Then
            Dim BdgOC As New BdgObjetivosCampo()
            Dim dtObj As DataTable = BdgOC.SelOnPrimaryKey(data.Current("IDObjetivoCampo"))
            If dtObj.Rows.Count > 0 Then
                data.Current("ObjetivoDia") = Nz(dtObj.Rows(0)("ObjetivoDia"), 0)
                data.Current("PremioUnidad") = Nz(dtObj.Rows(0)("PremioUnidad"), 0)
                data.Current("HorasJornada") = Nz(dtObj.Rows(0)("HorasJornada"), 0)
            End If
        End If
    End Sub

    <Serializable()> _
    Public Class dataTasas
        Public TasaHoraNormal As Double
        Public TasaHoraExtra As Double
    End Class
    <Task()> Public Shared Function GetTasasOperario(ByVal IDOperario As String, ByVal services As ServiceProvider) As dataTasas
        Dim t As New dataTasas
        Dim p As ParametroGeneral = services.GetService(Of ParametroGeneral)()

        Dim f As New Filter
        Dim f_OR As New Filter(FilterUnionOperator.Or)
        f.Add(New StringFilterItem("IDOperario", IDOperario))
        f_OR.Add(New StringFilterItem("IDHora", p.HoraPred))
        f_OR.Add(New StringFilterItem("IDHora", p.HoraExtra))
        f.Add(f_OR)

        Dim dt As DataTable = New OperarioHora().Filter(f)
        If dt.Rows.Count > 0 Then
            For Each dr As DataRow In dt.Rows
                If dr("IDHora") = p.HoraPred Then
                    t.TasaHoraNormal = Nz(dr("TasaHorariaA"), 0)
                Else
                    t.TasaHoraExtra = Nz(dr("TasaHorariaA"), 0)
                End If
            Next
        End If

        Return t
    End Function

    <Task()> Public Shared Sub CambioHorasTrabajadas(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CalcularImporteBaseDia, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CambioCantidad, data, services)
    End Sub

    <Task()> Public Shared Sub CambioHorasExtra(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CalcularImporteHorasExtra, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CambioCantidad, data, services)
    End Sub

    <Task()> Public Shared Sub CambioCantidad(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CalcularCantidadAdicional, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CalcularImporteCantidadDia, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CalcularImporteAcumulado, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CalcularImporteNeto, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf RedondearImportes, data, services)
    End Sub

    <Task()> Public Shared Sub CambioGastoBono(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CalcularImporteAcumulado, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CalcularImporteNeto, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf RedondearImportes, data, services)
    End Sub

    <Task()> Public Shared Sub CalcularImporteHorasExtra(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current("ImporteHorasExtra") = Nz(data.Current("HorasExtra"), 0) * Nz(data.Current("TasaHoraExtra"), 0)
    End Sub

    <Task()> Public Shared Sub CalcularImporteBaseDia(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current("ImporteBaseDia") = Nz(data.Current("HorasTrabajadas"), 0) * Nz(data.Current("TasaHoraNormal"), 0)
    End Sub

    <Task()> Public Shared Sub CalcularCantidadAdicional(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current("CantidadAdicional") = 0
        If Nz(data.Current("HorasTrabajadas"), 0) > 0 AndAlso Nz(data.Current("CantidadTrabajada"), 0) > 0 AndAlso Nz(data.Current("HorasJornada"), 0) > 0 Then
            Dim Objetivo As Double = data.Current("ObjetivoDia") / data.Current("HorasJornada")
            Dim UnidadesDia As Double = data.Current("HorasTrabajadas") * Objetivo
            data.Current("CantidadAdicional") = data.Current("CantidadTrabajada") - UnidadesDia
        End If
    End Sub

    <Task()> Public Shared Sub CalcularImporteCantidadDia(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current("ImporteCantidadDia") = Nz(data.Current("CantidadAdicional"), 0) * Nz(data.Current("PremioUnidad"), 0)
    End Sub

    <Task()> Public Shared Sub CalcularImporteAcumulado(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current("ImporteAcumulado") = Nz(data.Current("ImporteBaseDia"), 0) + Nz(data.Current("ImporteHorasExtra"), 0) + Nz(data.Current("Gasto"), 0) + Nz(data.Current("Bono"), 0) + IIf(Nz(data.Current("ImporteCantidadDia"), 0) > 0, data.Current("ImporteCantidadDia"), 0)
    End Sub

    <Task()> Public Shared Sub CalcularImporteNeto(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current("ImporteNeto") = Nz(data.Current("ImporteBaseDia"), 0) + Nz(data.Current("ImporteHorasExtra"), 0) + Nz(data.Current("Gasto"), 0) + Nz(data.Current("Bono"), 0) + Nz(data.Current("ImporteCantidadDia"), 0)
    End Sub

    <Task()> Public Shared Sub RedondearImportes(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        Dim monedas As MonedaCache = services.GetService(Of MonedaCache)()

        data.Current("ImporteHorasExtra") = xRound(Nz(data.Current("ImporteHorasExtra"), 0), monedas.MonedaA.NDecimalesImporte)
        data.Current("ImporteBaseDia") = xRound(Nz(data.Current("ImporteBaseDia"), 0), monedas.MonedaA.NDecimalesImporte)
        data.Current("ImporteCantidadDia") = xRound(Nz(data.Current("ImporteCantidadDia"), 0), monedas.MonedaA.NDecimalesImporte)
        data.Current("ImporteAcumulado") = xRound(Nz(data.Current("ImporteAcumulado"), 0), monedas.MonedaA.NDecimalesImporte)
        data.Current("ImporteNeto") = xRound(Nz(data.Current("ImporteNeto"), 0), monedas.MonedaA.NDecimalesImporte)
    End Sub

#End Region

#Region " RegisterUpdateTasks "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDObjetivoCampo")) = 0 Then ApplicationService.GenerateError("El Objetivo Campo es un dato obligatorio.")
        If Length(data("IDOperario")) = 0 Then ApplicationService.GenerateError("El Operario es un dato obligatorio.")
        If Length(data("IDTipoObra")) = 0 Then ApplicationService.GenerateError("El Tipo Obra es un dato obligatorio.")
        If Length(data("IDTipoTrabajo")) = 0 Then ApplicationService.GenerateError("El Tipo Trabajo es un dato obligatorio.")
        If Length(data("IDSubTipoTrabajo")) = 0 Then ApplicationService.GenerateError("El Tratamiento es un dato obligatorio.")
        If Length(data("IDFinca")) = 0 Then ApplicationService.GenerateError("La Finca es un dato obligatorio.")
    End Sub

#End Region

#Region " RegisterUpdateTasks "

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarIDPartesTrabajoFinca)
        updateProcess.AddTask(Of DataRow)(AddressOf Business.General.Comunes.UpdateEntityRow)
        updateProcess.AddTask(Of DataRow)(AddressOf Business.General.Comunes.MarcarComoActualizado)
        updateProcess.AddTask(Of DataRow)(AddressOf GenerarParteEnObras)
    End Sub

    <Task()> Public Shared Sub AsignarIDPartesTrabajoFinca(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If IsDBNull(data("IDPartesTrabajoFinca")) OrElse CType(data("IDPartesTrabajoFinca"), Guid).Equals(Guid.Empty) Then
                data("IDPartesTrabajoFinca") = Guid.NewGuid
            End If
        End If
    End Sub

    <Task()> Public Shared Sub GenerarParteEnObras(ByVal data As DataRow, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf BorrarObraModControl, data, services)

        Dim T As dataTrabajoFinca = ProcessServer.ExecuteTask(Of DataRow, dataTrabajoFinca)(AddressOf ObtenerTrabajo, data, services)
        If Not T.IDTrabajo Is Nothing Then
            Dim dataParteMOD As New dataGenerarPartesEnObras(data, T)
            ProcessServer.ExecuteTask(Of dataGenerarPartesEnObras)(AddressOf GenerarPartesEnObras, dataParteMOD, services)
        End If
    End Sub
    <Serializable()> _
    Public Class dataTrabajoFinca
        Public IDTrabajo As Integer?
        Public IDObra As Integer?
    End Class
    <Task()> Public Shared Function ObtenerTrabajo(ByVal data As DataRow, ByVal services As ServiceProvider) As dataTrabajoFinca
        Dim f As New Filter
        f.Add(New GuidFilterItem("IDPartesTrabajoFinca", data("IDPartesTrabajoFinca")))
        f.Add(New StringFilterItem("IDTipoObra", data("IDTipoObra")))
        f.Add(New StringFilterItem("IDTipoTrabajo", data("IDTipoTrabajo")))
        f.Add(New StringFilterItem("IDSubTipoTrabajo", data("IDSubTipoTrabajo")))

        Dim T As New dataTrabajoFinca
        Dim dtParte As DataTable = New BE.DataEngine().Filter("vBdgNegPartesFincaGeneracionTrabajos", f)
        If dtParte.Rows.Count > 0 Then
            Dim drParte() As DataRow = dtParte.Select(New DateFilterItem("FechaInicio", data("Fecha")).Compose(New AdoFilterComposer))
            If drParte.Length > 0 AndAlso Length(drParte(0)("IDTrabajo")) > 0 Then 'Ya existe el Trabajo del Tratamiento para la Fecha del parte.
                T.IDTrabajo = drParte(0)("IDTrabajo")
                T.IDObra = drParte(0)("IDObra")
            ElseIf Length(dtParte.Rows(0)("IDTrabajo")) > 0 AndAlso Length(dtParte.Rows(0)("FechaInicio")) = 0 Then 'Ya existe el Trabajo Padre. Hay que crear el Trabajo del Tratamiento para la Fecha del parte.
                dtParte.Rows(0)("FechaInicio") = data("Fecha")
                T.IDTrabajo = ProcessServer.ExecuteTask(Of DataRow, Integer)(AddressOf ObraParteAgrupado.GenerarTrabajo, dtParte.Rows(0), services)
                T.IDObra = dtParte.Rows(0)("IDObra")
            Else 'No exiten ni el Trabajo Padre ni el Trabajo del Tratamiento para la Fecha del parte.
                dtParte.Rows(0)("FechaInicio") = data("Fecha")
                dtParte.Rows(0)("IDTrabajo") = ProcessServer.ExecuteTask(Of DataRow, Integer)(AddressOf ObraParteAgrupado.GenerarTrabajo, dtParte.Rows(0), services)
                T.IDTrabajo = ProcessServer.ExecuteTask(Of DataRow, Integer)(AddressOf ObraParteAgrupado.GenerarTrabajo, dtParte.Rows(0), services)
                T.IDObra = dtParte.Rows(0)("IDObra")
            End If
        End If

        Return T
    End Function

    Public Class dataGenerarPartesEnObras
        Public IDTrabajo As Integer
        Public IDObra As Integer
        Public IDOperario As String
        Public FechaParte As Date
        Public HorasTrabajadas As Double
        Public HorasExtra As Double
        Public TasaHoraNormal As Double
        Public TasaHoraExtra As Double
        Public ImporteGasto As Double
        Public ImporteBono As Double
        Public ImportePremio As Double
        Public IDPartesTrabajoFinca As Guid
        Public dtMODControl As DataTable
        Public dtMODControlDelete As DataTable
        Public OMC As BusinessHelper

        Public Sub New(ByVal drParte As DataRow, ByVal Trabajo As dataTrabajoFinca)
            Me.IDObra = Trabajo.IDObra
            Me.IDTrabajo = Trabajo.IDTrabajo
            Me.IDOperario = drParte("IDOperario")
            Me.FechaParte = Nz(drParte("Fecha"), Today)
            Me.IDPartesTrabajoFinca = drParte("IDPartesTrabajoFinca")
            Me.HorasTrabajadas = Nz(drParte("HorasTrabajadas"), 0)
            Me.HorasExtra = Nz(drParte("HorasExtra"), 0)
            Me.TasaHoraNormal = Nz(drParte("TasaHoraNormal"), 0)
            Me.TasaHoraExtra = Nz(drParte("TasaHoraExtra"), 0)
            Me.ImporteGasto = Nz(drParte("Gasto"), 0)
            Me.ImporteBono = Nz(drParte("Bono"), 0)
            Me.ImportePremio = Nz(drParte("ImporteCantidadDia"), 0)

            Me.OMC = BusinessHelper.CreateBusinessObject("ObraModControl")
            Me.dtMODControlDelete = OMC.Filter(New GuidFilterItem("IDPartesTrabajoFinca", drParte("IDPartesTrabajoFinca")))
            Me.dtMODControl = dtMODControlDelete.Clone
            If Me.dtMODControlDelete.Rows.Count > 0 Then
                For Each drMODControlDelete As DataRow In Me.dtMODControlDelete.Rows
                    drMODControlDelete("IDPartesTrabajoFinca") = DBNull.Value
                Next
            End If
        End Sub
    End Class
    <Task()> Public Shared Sub GenerarPartesEnObras(ByVal data As dataGenerarPartesEnObras, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of dataGenerarPartesEnObras)(AddressOf GenerarPartesHorasNormales, data, services)
        ProcessServer.ExecuteTask(Of dataGenerarPartesEnObras)(AddressOf GenerarPartesHorasExtras, data, services)
        ProcessServer.ExecuteTask(Of dataGenerarPartesEnObras)(AddressOf GenerarPartesGastos, data, services)
        ProcessServer.ExecuteTask(Of dataGenerarPartesEnObras)(AddressOf GenerarPartesBonos, data, services)
        ProcessServer.ExecuteTask(Of dataGenerarPartesEnObras)(AddressOf GenerarPartesPremios, data, services)

        If data.dtMODControl.Rows.Count > 0 Then
            Dim pck As New UpdatePackage
            If data.dtMODControlDelete.Rows.Count > 0 Then
                pck.DeletedData("ObraMODControl") = data.dtMODControlDelete
            End If

            Dim dtObra As DataTable = AdminData.GetData("tbObraCabecera", New NumberFilterItem("IDObra", data.IDObra))
            dtObra.TableName = "ObraCabecera"
            pck.Add(dtObra)

            pck.Add(data.dtMODControl)

            Dim OC As BusinessHelper = CreateBusinessObject("ObraCabecera")
            OC.Update(pck)
            ProcessServer.ExecuteTask(Of Integer)(AddressOf ObraCabecera.RecalcularObra, data.IDObra, services)
        End If
    End Sub
    <Task()> Public Shared Sub GenerarPartesHorasNormales(ByVal data As dataGenerarPartesEnObras, ByVal services As ServiceProvider)
        If data.HorasTrabajadas <> 0 Then
            Dim drMODControl As DataRow = data.dtMODControl.NewRow

            drMODControl("IDLineaMODControl") = AdminData.GetAutoNumeric
            drMODControl("IDObra") = data.IDObra
            drMODControl("IDTrabajo") = data.IDTrabajo
            drMODControl("DescParte") = AdminData.GetMessageText("HORAS TRABAJADAS")

            drMODControl("IDOperario") = data.IDOperario
            data.OMC.ApplyBusinessRule("IDOperario", drMODControl("IDOperario"), drMODControl, New BusinessData)

            Dim p As ParametroGeneral = services.GetService(Of ParametroGeneral)()
            drMODControl("IDHora") = p.HoraPred
            drMODControl("FechaInicio") = data.FechaParte
            drMODControl("HorasRealMOD") = data.HorasTrabajadas
            drMODControl("TasaRealModA") = data.TasaHoraNormal
            data.OMC.ApplyBusinessRule("TasaRealModA", drMODControl("TasaRealModA"), drMODControl, New BusinessData)

            drMODControl("IDPartesTrabajoFinca") = data.IDPartesTrabajoFinca

            data.dtMODControl.Rows.Add(drMODControl.ItemArray)
        End If
    End Sub
    <Task()> Public Shared Sub GenerarPartesHorasExtras(ByVal data As dataGenerarPartesEnObras, ByVal services As ServiceProvider)
        If data.HorasExtra <> 0 Then
            Dim drMODControl As DataRow = data.dtMODControl.NewRow

            drMODControl("IDLineaMODControl") = AdminData.GetAutoNumeric
            drMODControl("IDObra") = data.IDObra
            drMODControl("IDTrabajo") = data.IDTrabajo
            drMODControl("DescParte") = AdminData.GetMessageText("HORAS EXTRAS")

            drMODControl("IDOperario") = data.IDOperario
            data.OMC.ApplyBusinessRule("IDOperario", drMODControl("IDOperario"), drMODControl, New BusinessData)

            Dim p As ParametroGeneral = services.GetService(Of ParametroGeneral)()
            drMODControl("IDHora") = p.HoraExtra
            drMODControl("FechaInicio") = data.FechaParte
            drMODControl("HorasRealMOD") = data.HorasExtra
            drMODControl("TasaRealModA") = data.TasaHoraExtra
            data.OMC.ApplyBusinessRule("TasaRealModA", drMODControl("TasaRealModA"), drMODControl, New BusinessData)

            drMODControl("IDPartesTrabajoFinca") = data.IDPartesTrabajoFinca

            data.dtMODControl.Rows.Add(drMODControl.ItemArray)
        End If
    End Sub
    <Task()> Public Shared Sub GenerarPartesGastos(ByVal data As dataGenerarPartesEnObras, ByVal services As ServiceProvider)
        If data.ImporteGasto <> 0 Then
            Dim drMODControl As DataRow = data.dtMODControl.NewRow

            drMODControl("IDLineaMODControl") = AdminData.GetAutoNumeric
            drMODControl("IDObra") = data.IDObra
            drMODControl("IDTrabajo") = data.IDTrabajo
            drMODControl("DescParte") = AdminData.GetMessageText("GASTOS GENERADOS")

            drMODControl("IDOperario") = data.IDOperario
            data.OMC.ApplyBusinessRule("IDOperario", drMODControl("IDOperario"), drMODControl, New BusinessData)

            drMODControl("FechaInicio") = data.FechaParte
            drMODControl("HorasRealMOD") = 1
            drMODControl("TasaRealModA") = data.ImporteGasto
            data.OMC.ApplyBusinessRule("TasaRealModA", drMODControl("TasaRealModA"), drMODControl, New BusinessData)

            drMODControl("IDPartesTrabajoFinca") = data.IDPartesTrabajoFinca

            data.dtMODControl.Rows.Add(drMODControl.ItemArray)
        End If
    End Sub
    <Task()> Public Shared Sub GenerarPartesBonos(ByVal data As dataGenerarPartesEnObras, ByVal services As ServiceProvider)
        If data.ImporteBono <> 0 Then
            Dim drMODControl As DataRow = data.dtMODControl.NewRow

            drMODControl("IDLineaMODControl") = AdminData.GetAutoNumeric
            drMODControl("IDObra") = data.IDObra
            drMODControl("IDTrabajo") = data.IDTrabajo
            drMODControl("DescParte") = AdminData.GetMessageText("BONOS")

            drMODControl("IDOperario") = data.IDOperario
            data.OMC.ApplyBusinessRule("IDOperario", drMODControl("IDOperario"), drMODControl, New BusinessData)

            drMODControl("FechaInicio") = data.FechaParte
            drMODControl("HorasRealMOD") = 1
            drMODControl("TasaRealModA") = data.ImporteBono
            data.OMC.ApplyBusinessRule("TasaRealModA", drMODControl("TasaRealModA"), drMODControl, New BusinessData)

            drMODControl("IDPartesTrabajoFinca") = data.IDPartesTrabajoFinca

            data.dtMODControl.Rows.Add(drMODControl.ItemArray)
        End If
    End Sub
    <Task()> Public Shared Sub GenerarPartesPremios(ByVal data As dataGenerarPartesEnObras, ByVal services As ServiceProvider)
        If data.ImportePremio <> 0 Then
            Dim drMODControl As DataRow = data.dtMODControl.NewRow

            drMODControl("IDLineaMODControl") = AdminData.GetAutoNumeric
            drMODControl("IDObra") = data.IDObra
            drMODControl("IDTrabajo") = data.IDTrabajo
            drMODControl("DescParte") = AdminData.GetMessageText("PREMIOS OBTENIDOS")

            drMODControl("IDOperario") = data.IDOperario
            data.OMC.ApplyBusinessRule("IDOperario", drMODControl("IDOperario"), drMODControl, New BusinessData)

            drMODControl("FechaInicio") = data.FechaParte
            drMODControl("HorasRealMOD") = 1
            drMODControl("TasaRealModA") = data.ImportePremio
            data.OMC.ApplyBusinessRule("TasaRealModA", drMODControl("TasaRealModA"), drMODControl, New BusinessData)

            drMODControl("IDPartesTrabajoFinca") = data.IDPartesTrabajoFinca

            data.dtMODControl.Rows.Add(drMODControl.ItemArray)
        End If
    End Sub

#End Region

    <Serializable()> _
    Public Class dataGenerarPartesTrabajoFinca
        Public Fecha As Date
        Public IDGrupo As String
        Public IDFinca As Guid
        Public IDFincaHija As Guid
        Public IDTipoObra As String
        Public IDTipoTrabajo As String
        Public IDSubTipoTrabajo As String
        Public IDObjetivoCampo As String
        Public BorrarPartesExistentes As Boolean = False

        Public Sub New(ByVal Fecha As Date, ByVal IDGrupo As String, ByVal IDTipoObra As String, ByVal IDTipoTrabajo As String, ByVal IDSubTipoTrabajo As String, ByVal IDFinca As Guid, ByVal IDFincaHija As Guid, ByVal IDObjetivoCampo As String)
            Me.Fecha = Fecha
            Me.IDGrupo = IDGrupo
            Me.IDFinca = IDFinca
            Me.IDFincaHija = IDFincaHija
            Me.IDTipoObra = IDTipoObra
            Me.IDTipoTrabajo = IDTipoTrabajo
            Me.IDSubTipoTrabajo = IDSubTipoTrabajo
            Me.IDObjetivoCampo = IDObjetivoCampo
        End Sub
    End Class
    <Task()> Public Shared Sub GenerarPartesTrabajoFinca(ByVal data As dataGenerarPartesTrabajoFinca, ByVal services As ServiceProvider)
        Dim o As New BdgPartesTrabajoFinca
        If data.BorrarPartesExistentes Then
            Dim dt As DataTable = o.Filter(New DateFilterItem("Fecha", data.Fecha))
            If dt.Rows.Count > 0 Then o.Delete(dt)
        End If
        Dim dtNew As DataTable = o.AddNew
        Dim dtOperarios As DataTable = New GrupoOperario().Filter(New StringFilterItem("IDGrupo", data.IDGrupo))
        If dtOperarios.Rows.Count > 0 Then
            For Each drOperario As DataRow In dtOperarios.Rows
                Dim drNew As DataRow = dtNew.NewRow
                drNew("Fecha") = data.Fecha
                drNew("IDObjetivoCampo") = data.IDObjetivoCampo
                drNew("IDOperario") = drOperario("IDOperario")
                drNew("IDTipoObra") = data.IDTipoObra
                drNew("IDTipoTrabajo") = data.IDTipoTrabajo
                drNew("IDSubTipoTrabajo") = data.IDSubTipoTrabajo
                drNew("IDFinca") = data.IDFinca
                If Not data.IDFincaHija.Equals(Guid.Empty) Then drNew("IDFincaHija") = data.IDFincaHija

                Dim d As dataTasas = ProcessServer.ExecuteTask(Of String, dataTasas)(AddressOf GetTasasOperario, drNew("IDOperario"), services)
                drNew("TasaHoraNormal") = d.TasaHoraNormal
                drNew("TasaHoraExtra") = d.TasaHoraExtra

                dtNew.Rows.Add(drNew.ItemArray)
            Next
            BdgPartesTrabajoFinca.UpdateTable(dtNew)
        End If
    End Sub

End Class
