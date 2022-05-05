Public Class BdgOperacion

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgOperacion"

#End Region

#Region "Eventos Entidad"

#Region "Tareas RegisterAddNewTasks"

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf NegocioGeneral.AsignarCentroGestion, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarContadorPredeterminado, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarNumeroOperacionProvisional, data, services)

        Dim datFecha As New BdgGeneral.DataGetFechaPropuestaOperacion(GetType(BdgOperacion).Name)
        data("Fecha") = ProcessServer.ExecuteTask(Of BdgGeneral.DataGetFechaPropuestaOperacion, Date)(AddressOf BdgGeneral.GetFechaPropuestaOperacion, datFecha, services)

        data("ImputacionRealMaterial") = False : data("ImputacionRealMod") = False
        data("ImputacionRealCentro") = False : data("ImputacionRealVarios") = False
    End Sub

    <Task()> Public Shared Sub AsignarContadorPredeterminado(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim CE As New CentroEntidad
        CE.CentroGestion = data("IDCentroGestion") & String.Empty
        CE.ContadorEntidad = CentroGestion.ContadorEntidad.BdgOperacion
        data("IDContador") = ProcessServer.ExecuteTask(Of CentroEntidad, String)(AddressOf CentroGestion.GetContadorPredeterminado, CE, services)
    End Sub

    <Task()> Public Shared Sub AsignarNumeroOperacionProvisional(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDContador")) > 0 Then
            Dim dtContadores As DataTable = ProcessServer.ExecuteTask(Of String, DataTable)(AddressOf Contador.CounterDt, GetType(BdgOperacion).Name, services)
            Dim adr As DataRow() = dtContadores.Select("IDContador = " & Quoted(data("IDContador")))
            If Not IsNothing(adr) AndAlso adr.Length > 0 Then
                data("NOperacion") = adr(0)("ValorProvisional")
            Else
                Dim dtContadorPred As DataTable = ProcessServer.ExecuteTask(Of String, DataTable)(AddressOf Contador.CounterDefault, GetType(BdgOperacion).Name, services)
                If Not dtContadorPred Is Nothing AndAlso dtContadorPred.Rows.Count > 0 Then
                    data("IDContador") = dtContadorPred.Rows(0)("IDContador")
                    adr = dtContadores.Select("IDContador = " & Quoted(data("IDContador")))
                    If Not IsNothing(adr) AndAlso adr.Length > 0 Then
                        data("NOperacion") = adr(0)("ValorProvisional")
                    End If
                End If
            End If
        End If
    End Sub


#End Region

#Region "Tareas RegisterDeleteTasks"

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf Comunes.BeginTransaction)
        deleteProcess.AddTask(Of DataRow)(AddressOf DeleteOperacionesVino)
        deleteProcess.AddTask(Of DataRow)(AddressOf Comunes.DeleteEntityRow)
        deleteProcess.AddTask(Of DataRow)(AddressOf Comunes.MarcarComoEliminado)
        deleteProcess.AddTask(Of DataRow)(AddressOf DeshacerConfirmacion)
        deleteProcess.AddTask(Of DataRow)(AddressOf ActualizarMovimientos)
    End Sub

    <Task()> Public Shared Sub DeleteOperacionesVino(ByVal data As DataRow, ByVal services As ServiceProvider)
        '//NOTA: NO ELIMINAR ESTA TAREA. Hay que eliminar las lineas de OperacionVino en orden. Primero los destinos, despues los origenes...
        '// e IMPORTANTE deshabilitar la marca de borrado en cascada entre BDGOperacion y BDGOperacionVino
        Dim oOV As New BdgOperacionVino
        Dim dtOV As DataTable = ProcessServer.ExecuteTask(Of String, DataTable)(AddressOf BdgOperacionVino.SelOnNOperacion, data("NOperacion"), services)

        Dim DtDelete As DataTable = dtOV.Clone
        For Each Dr As DataRow In dtOV.Select("", "Destino DESC")
            DtDelete.ImportRow(Dr)
        Next
        oOV.Delete(DtDelete)
    End Sub


    <Task()> Public Shared Sub ActualizarMovimientos(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim FechaDoc As New Date(CDate(data("Fecha")).Year, CDate(data("Fecha")).Month, CDate(data("Fecha")).Day)
        If Length(data("IDMovimiento")) <> 0 Then
            Dim StEj As New BdgWorkClass.StEjecutarMovimientosNumero(data("IDMovimiento"), data("NOperacion"), FechaDoc)
            ProcessServer.ExecuteTask(Of BdgWorkClass.StEjecutarMovimientosNumero)(AddressOf BdgWorkClass.EjecutarMovimientosNumero, StEj, services)
        Else
            Dim StEj As New BdgWorkClass.StEjecutarMovimientos(data("NOperacion"), FechaDoc)
            ProcessServer.ExecuteTask(Of BdgWorkClass.StEjecutarMovimientos, Integer)(AddressOf BdgWorkClass.EjecutarMovimientos, StEj, services)
        End If
    End Sub

    <Task()> Public Shared Sub DeshacerConfirmacion(ByVal data As DataRow, ByVal services As ServiceProvider)
        If (Length(data("NOperacionPlan")) > 0) Then
            Dim stData As New BdgOperacionPlan.StCambiarEstadoOperacion(data("NOperacionPlan"), BdgEstadoOperacionPlan.Planificado)
            ProcessServer.ExecuteTask(Of BdgOperacionPlan.StCambiarEstadoOperacion)(AddressOf BdgOperacionPlan.CambiarEstadoOperacion, stData, services)
        End If
    End Sub

#End Region

#Region "Tareas GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim OBrl As New BusinessRules
        OBrl.Add("IDTipoOperacion", AddressOf CambioIDTipoOperacion)
        Return OBrl
    End Function

    <Task()> Public Shared Sub CambioIDTipoOperacion(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If data.Context Is Nothing Then data.Context = New BusinessData
        data.Context("Origen") = enumBdgOrigenOperacion.Real
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf BdgGeneral.CambioIDTipoOperacion, data, services)
    End Sub

#End Region

#Region "Tareas RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDTipoOperacion")) = 0 Then ApplicationService.GenerateError("El Tipo de Operación es un dato obligatorio.")
    End Sub

#End Region

#Region "Tareas RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of UpdatePackage, DocumentoBdgOperacionReal)(AddressOf ProcesoBdgOperacion.CrearDocumento)
        updateProcess.AddTask(Of DocumentoBdgOperacionReal)(AddressOf Comunes.BeginTransaction)
        updateProcess.AddTask(Of DocumentoBdgOperacionReal)(AddressOf ProcesoBdgOperacion.AsignarDatosGlobalesLineas)
        updateProcess.AddTask(Of DocumentoBdgOperacionReal)(AddressOf ProcesoBdgOperacion.ValidacionesGeneralesOperacion)
        updateProcess.AddTask(Of DocumentoBdgOperacionReal)(AddressOf ProcesoBdgOperacion.AsignarNumeroOperacion)
        updateProcess.AddTask(Of DocumentoBdgOperacionReal)(AddressOf ProcesoBdgOperacion.ActualizarEstadoVino)
        updateProcess.AddTask(Of DocumentoBdgOperacionReal)(AddressOf ProcesoBdgOperacion.AsignarDatosOperacionOrigen)
        updateProcess.AddTask(Of DocumentoBdgOperacionReal)(AddressOf ProcesoBdgOperacion.AsignarDatosOperacionDestino)
        updateProcess.AddTask(Of DocumentoBdgOperacionReal)(AddressOf ProcesoBdgOperacion.AsignarDatosVinoCentro)
        updateProcess.AddTask(Of DocumentoBdgOperacionReal)(AddressOf ProcesoBdgOperacion.AsignarDatosVinoAnalisisVariable)
        updateProcess.AddTask(Of DocumentoBdgOperacionReal)(AddressOf ProcesoBdgOperacion.AsignarDatosVinoMOD)
        updateProcess.AddTask(Of DocumentoBdgOperacionReal)(AddressOf ProcesoBdgOperacion.AsignarDatosVinoMaterial)
        updateProcess.AddTask(Of DocumentoBdgOperacionReal)(AddressOf ProcesoBdgOperacion.CambioFechaOperacion)
        updateProcess.AddTask(Of DocumentoBdgOperacionReal)(AddressOf ProcesoBdgOperacion.DeleteImputacionesVino)
        updateProcess.AddTask(Of DocumentoBdgOperacionReal)(AddressOf Comunes.UpdateDocument)
        updateProcess.AddTask(Of DocumentoBdgOperacionReal)(AddressOf Comunes.MarcarComoActualizado)
        updateProcess.AddTask(Of DocumentoBdgOperacionReal)(AddressOf ProcesoBdgOperacion.ActualizarOperacionPlanRelacionada)
        updateProcess.AddTask(Of DocumentoBdgOperacionReal)(AddressOf ProcesoBdgOperacion.ActualizarVinoCentroTasa)
        updateProcess.AddTask(Of DocumentoBdgOperacionReal)(AddressOf ProcesoBdgOperacion.ActualizarCostesVino)
        updateProcess.AddTask(Of DocumentoBdgOperacionReal)(AddressOf ProcesoBdgOperacion.EjecutarMovimientos)
        updateProcess.AddTask(Of DocumentoBdgOperacionReal)(AddressOf ProcesoBdgOperacion.CrearMovimientosMaterialesVino)
        updateProcess.AddTask(Of DocumentoBdgOperacionReal)(AddressOf ProcesoBdgOperacion.ActualizarOF)
    End Sub

#End Region

#End Region

#Region "Funciones Públicas"

#Region " Gestión Operaciones Previstas "

#Region " Generación masiva de Operaciones "

    <Serializable()> _
    Public Class StNuevaOperacion
        Public IDDeposito As String
        Public IDVino As Guid
        Public IDTipoOperacion As String
        Public FechaOperacion As Date

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDDeposito As String, ByVal IDVino As Guid, ByVal IDTipoOperacion As String, Optional ByVal FechaOperacion As Date = cnMinDate)
            Me.IDDeposito = IDDeposito
            Me.IDVino = IDVino
            Me.IDTipoOperacion = IDTipoOperacion
            Me.FechaOperacion = FechaOperacion
        End Sub
    End Class

    <Serializable()> _
    Public Class StNuevaOperacionResult
        Public NOperacion As String
        Public Errores As New List(Of ClassErrors)

        Public Sub New(ByVal NOperacion As String)
            Me.NOperacion = NOperacion
        End Sub
    End Class


    <Task()> Public Shared Function NuevaOperacion(ByVal data As StNuevaOperacion, ByVal services As ServiceProvider) As StNuevaOperacionResult
        If Length(data.IDDeposito) > 0 And Length(data.IDTipoOperacion) > 0 Then

            Dim lstPropuesta As Dictionary(Of String, DataTable) = ProcessServer.ExecuteTask(Of StNuevaOperacion, Dictionary(Of String, DataTable))(AddressOf PropuestaNuevaOperacion, data, services)
            If Not lstPropuesta Is Nothing Then
                Return ProcessServer.ExecuteTask(Of Dictionary(Of String, DataTable), StNuevaOperacionResult)(AddressOf ActualizarNuevaOperacion, lstPropuesta, services)
            End If
        End If

    End Function

    <Task()> Public Shared Function PropuestaNuevaOperacion(ByVal data As StNuevaOperacion, ByVal services As ServiceProvider) As Dictionary(Of String, DataTable)
        If Length(data.IDDeposito) > 0 And Length(data.IDTipoOperacion) > 0 Then

            Dim StADD As New StADDCabeceraOperacion(data.IDTipoOperacion, data.FechaOperacion)
            Dim lstPropuesta As Dictionary(Of String, DataTable) = ProcessServer.ExecuteTask(Of StADDCabeceraOperacion, Dictionary(Of String, DataTable))(AddressOf ADDCabeceraOperacion, StADD, services)
            If lstPropuesta Is Nothing OrElse lstPropuesta.Count = 0 Then Exit Function

            Dim dtOperacion As DataTable
            If lstPropuesta.ContainsKey(GetType(BdgOperacion).Name) Then dtOperacion = lstPropuesta(GetType(BdgOperacion).Name)
            If dtOperacion Is Nothing OrElse dtOperacion.Rows.Count = 0 Then Exit Function

            Dim drTipoOperacion As DataRow = New BdgTipoOperacion().GetItemRow(data.IDTipoOperacion)

            Dim dtOrigen, dtDestino, dtMaterial As DataTable

            Dim strNOperacion As String = dtOperacion.Rows(0)("NOperacion") & String.Empty
            Dim FechaOperacion As Date = Nz(dtOperacion.Rows(0)("Fecha"), Now)

            Select Case drTipoOperacion("TipoMovimiento")
                Case Business.Bodega.TipoMovimiento.Salida, Business.Bodega.TipoMovimiento.DeVariosAVarios, Business.Bodega.TipoMovimiento.DeUnoAUno, Business.Bodega.TipoMovimiento.DeUnoAVarios, Business.Bodega.TipoMovimiento.DeVariosAUno
                    Dim StOr As New StADDOrigenDestino(data.IDDeposito, strNOperacion, data.IDVino)
                    dtOrigen = ProcessServer.ExecuteTask(Of StADDOrigenDestino, DataTable)(AddressOf ADDOrigen, StOr, services)
                    lstPropuesta(GetType(BdgOperacionVino).Name & "Origen") = dtOrigen
                Case Business.Bodega.TipoMovimiento.SinMovimiento, Business.Bodega.TipoMovimiento.SinOrigen
                    Dim StDes As New StADDOrigenDestino(data.IDDeposito, strNOperacion, data.IDVino)
                    dtDestino = ProcessServer.ExecuteTask(Of StADDOrigenDestino, DataTable)(AddressOf ADDDestino, StDes, services)
                    lstPropuesta(GetType(BdgOperacionVino).Name & "Destino") = dtDestino
                    Dim StAddMat As New StADDMateriales(data.IDTipoOperacion, strNOperacion, FechaOperacion, data.IDVino, dtDestino.Rows(0)("IDArticulo") & String.Empty, _
                                               Nz(dtDestino.Rows(0)("Cantidad"), 0), CBool(drTipoOperacion("ImputacionPrevMaterial")))
                    dtMaterial = ProcessServer.ExecuteTask(Of StADDMateriales, DataTable)(AddressOf ADDMateriales, StAddMat, services)
                    lstPropuesta(GetType(BdgVinoMaterial).Name) = dtMaterial
            End Select

            Return lstPropuesta
        End If
    End Function

    <Task()> Public Shared Function ActualizarNuevaOperacion(ByVal lstPropuesta As Dictionary(Of String, DataTable), ByVal service As ServiceProvider) As StNuevaOperacionResult
        Dim rslt As New StNuevaOperacionResult(String.Empty)
        If Not lstPropuesta Is Nothing AndAlso lstPropuesta.Count > 0 Then

            Dim dtOperacion As DataTable
            Dim dtOperacionVino, dtOperacionVinoOrigen, dtOperacionVinoDestino As DataTable
            Dim dtMaterialGlobal, dtMaterial As DataTable
            Dim dtCentrosGlobal, dtCentros As DataTable
            Dim dtMODGlobal, dtMOD As DataTable
            Dim dtVariosGlobal, dtVarios As DataTable
            Dim dtVinoAnalisis, dtVinoAnalisisVariable As DataTable

            If lstPropuesta.ContainsKey(GetType(BdgOperacion).Name) Then dtOperacion = lstPropuesta(GetType(BdgOperacion).Name)

            If lstPropuesta.ContainsKey(GetType(BdgOperacionVino).Name) Then dtOperacionVino = lstPropuesta(GetType(BdgOperacionVino).Name)
            If lstPropuesta.ContainsKey(GetType(BdgOperacionVino).Name & "Origen") Then dtOperacionVinoOrigen = lstPropuesta(GetType(BdgOperacionVino).Name & "Origen")
            If lstPropuesta.ContainsKey(GetType(BdgOperacionVino).Name & "Destino") Then dtOperacionVinoDestino = lstPropuesta(GetType(BdgOperacionVino).Name & "Destino")

            If lstPropuesta.ContainsKey(GetType(BdgOperacionMaterial).Name) Then dtMaterialGlobal = lstPropuesta(GetType(BdgOperacionMaterial).Name)
            If lstPropuesta.ContainsKey(GetType(BdgOperacionMOD).Name) Then dtMODGlobal = lstPropuesta(GetType(BdgOperacionMOD).Name)
            If lstPropuesta.ContainsKey(GetType(BdgOperacionCentro).Name) Then dtCentrosGlobal = lstPropuesta(GetType(BdgOperacionCentro).Name)
            If lstPropuesta.ContainsKey(GetType(BdgOperacionVarios).Name) Then dtVariosGlobal = lstPropuesta(GetType(BdgOperacionVarios).Name)


            If lstPropuesta.ContainsKey(GetType(BdgVinoMaterial).Name) Then dtMaterial = lstPropuesta(GetType(BdgVinoMaterial).Name)
            If lstPropuesta.ContainsKey(GetType(BdgVinoMod).Name) Then dtMOD = lstPropuesta(GetType(BdgVinoMod).Name)
            If lstPropuesta.ContainsKey(GetType(BdgVinoCentro).Name) Then dtCentros = lstPropuesta(GetType(BdgVinoCentro).Name)
            If lstPropuesta.ContainsKey(GetType(BdgVinoVarios).Name) Then dtVarios = lstPropuesta(GetType(BdgVinoVarios).Name)
            If lstPropuesta.ContainsKey(GetType(BdgVinoAnalisis).Name) Then dtVinoAnalisis = lstPropuesta(GetType(BdgVinoAnalisis).Name)
            If lstPropuesta.ContainsKey(GetType(BdgVinoVariable).Name) Then dtVinoAnalisisVariable = lstPropuesta(GetType(BdgVinoVariable).Name)

            If Not dtOperacion Is Nothing AndAlso dtOperacion.Rows.Count > 0 Then
                Try
                    AdminData.BeginTx()
                    Dim UpdtPckg As New UpdatePackage
                    UpdtPckg.Add(dtOperacion)

                    If dtOperacionVino Is Nothing OrElse dtOperacionVino.Rows.Count = 0 Then
                        If Not dtOperacionVinoOrigen Is Nothing AndAlso dtOperacionVinoOrigen.Rows.Count > 0 Then
                            dtOperacionVino = dtOperacionVinoOrigen
                            If Not dtOperacionVinoDestino Is Nothing AndAlso dtOperacionVinoDestino.Rows.Count > 0 Then
                                For Each drDestino As DataRow In dtOperacionVinoDestino.Rows
                                    dtOperacionVino.ImportRow(drDestino)
                                Next
                            End If

                        ElseIf Not dtOperacionVinoDestino Is Nothing AndAlso dtOperacionVinoDestino.Rows.Count > 0 Then
                            dtOperacionVino = dtOperacionVinoDestino
                        End If
                    End If

                    If Not dtOperacionVino Is Nothing AndAlso dtOperacionVino.Rows.Count > 0 Then
                        UpdtPckg.Add(GetType(BdgOperacionVino).Name, dtOperacionVino)
                    End If


                    If Not dtMaterialGlobal Is Nothing Then UpdtPckg.Add(GetType(BdgOperacionMaterial).Name, dtMaterialGlobal)
                    If Not dtMODGlobal Is Nothing Then UpdtPckg.Add(GetType(BdgOperacionMOD).Name, dtMODGlobal)
                    If Not dtCentrosGlobal Is Nothing Then UpdtPckg.Add(GetType(BdgOperacionCentro).Name, dtCentrosGlobal)
                    If Not dtVariosGlobal Is Nothing Then UpdtPckg.Add(GetType(BdgOperacionVarios).Name, dtVariosGlobal)

                    If Not dtMaterial Is Nothing Then UpdtPckg.Add(GetType(BdgVinoMaterial).Name, dtMaterial)
                    If Not dtMOD Is Nothing Then UpdtPckg.Add(GetType(BdgVinoMod).Name, dtMOD)
                    If Not dtCentros Is Nothing Then UpdtPckg.Add(GetType(BdgVinoCentro).Name, dtCentros)
                    If Not dtVarios Is Nothing Then UpdtPckg.Add(GetType(BdgVinoVarios).Name, dtVarios)
                    If Not dtVinoAnalisis Is Nothing Then UpdtPckg.Add(GetType(BdgVinoAnalisis).Name, dtVinoAnalisis)
                    If Not dtVinoAnalisisVariable Is Nothing Then UpdtPckg.Add(GetType(BdgVinoVariable).Name, dtVinoAnalisisVariable)

                    Dim Op As New BdgOperacion
                    Op.Update(UpdtPckg)

                    rslt.NOperacion = dtOperacion.Rows(0)("NOperacion")

                Catch ex As Exception
                    rslt.Errores.Add(New ClassErrors(Nothing, ex.Message & "(ActualizarNuevaOperacion)"))
                    AdminData.RollBackTx()
                End Try
            End If
        Else
            rslt.Errores.Add(New ClassErrors(Nothing, "No hay datos para generar Operación planificada."))
        End If

        Return rslt
    End Function

    '<Serializable()> _
    'Public Class StNuevaOperacionDatatables
    '    Public IDTipoOperacion As String
    '    Public FechaOperacion As Date
    '    Public UnaOperacion As Boolean
    '    Public DtOrigen As DataTable
    '    Public DtDestino As DataTable

    '    Public Sub New()
    '    End Sub

    '    Public Sub New(ByVal IDTipoOperacion As String, ByVal FechaOperacion As Date, ByVal UnaOperacion As Boolean, ByVal DtOrigen As DataTable, ByVal DtDestino As DataTable)
    '        Me.IDTipoOperacion = IDTipoOperacion
    '        Me.FechaOperacion = FechaOperacion
    '        Me.UnaOperacion = UnaOperacion
    '        Me.DtOrigen = DtOrigen
    '        Me.DtDestino = DtDestino
    '    End Sub
    'End Class

    '<Task()> Public Shared Function NuevaOperacionDatatables(ByVal data As StNuevaOperacionDatatables, ByVal services As ServiceProvider) As OperacionUpdateData
    '    Dim OperacionInfo As New OperacionUpdateData
    '    ReDim OperacionInfo.NOperacion(-1)
    '    ReDim OperacionInfo.Operaciones(-1)

    '    If Length(data.IDTipoOperacion) > 0 Then
    '        Dim dt As DataTable
    '        Dim dtOrigenes As DataTable
    '        Dim dtDestinos As DataTable
    '        If Not data.DtOrigen Is Nothing Then
    '            dt = data.DtOrigen
    '            dtOrigenes = data.DtOrigen.Clone
    '        ElseIf Not data.DtDestino Is Nothing Then
    '            dt = data.DtDestino
    '            dtDestinos = data.DtDestino.Clone
    '        End If
    '        If data.UnaOperacion Then
    '            Dim StAddOp As New StADDCabeceraOperacion(data.IDTipoOperacion, data.FechaOperacion)
    '            Dim lstCabecera As Dictionary(Of String, DataTable) = ProcessServer.ExecuteTask(Of StADDCabeceraOperacion, Dictionary(Of String, DataTable))(AddressOf ADDCabeceraOperacion, StAddOp, services)
    '            Dim dtOperacion As DataTable = lstCabecera(GetType(BdgOperacion).Name)
    '            Dim dtMaterialGlobal As DataTable = lstCabecera(GetType(BdgOperacionMaterial).Name)
    '            Dim dtMODGlobal As DataTable = lstCabecera(GetType(BdgOperacionMOD).Name)
    '            Dim dtCentroGlobal As DataTable = lstCabecera(GetType(BdgOperacionCentro).Name)
    '            Dim dtVariosGlobal As DataTable = lstCabecera(GetType(BdgOperacionVarios).Name)

    '            Dim drTipoOperacion As DataRow = New BdgTipoOperacion().GetItemRow(data.IDTipoOperacion)
    '            Dim strNOperacion As String = dtOperacion.Rows(0)("NOperacion") & String.Empty
    '            Dim dtMaterial As DataTable

    '            For Each dr As DataRow In dt.Rows
    '                Select Case drTipoOperacion("TipoMovimiento")
    '                    Case Business.Bodega.TipoMovimiento.Salida, Business.Bodega.TipoMovimiento.DeVariosAVarios, Business.Bodega.TipoMovimiento.DeUnoAUno, Business.Bodega.TipoMovimiento.DeUnoAVarios, Business.Bodega.TipoMovimiento.DeVariosAUno
    '                        Dim StADDOr As New StADDOrigenDestino(dr("IDDeposito"), strNOperacion, dr("IDVino"))
    '                        data.DtOrigen = ProcessServer.ExecuteTask(Of StADDOrigenDestino, DataTable)(AddressOf ADDOrigen, StADDOr, services)
    '                        dtOrigenes.Rows.Add(data.DtOrigen.Rows(0).ItemArray)
    '                    Case Business.Bodega.TipoMovimiento.SinMovimiento, Business.Bodega.TipoMovimiento.SinOrigen
    '                        Dim StAddDes As New StADDOrigenDestino(dr("IDDeposito"), strNOperacion, dr("IDVino"))
    '                        data.DtDestino = ProcessServer.ExecuteTask(Of StADDOrigenDestino, DataTable)(AddressOf ADDDestino, StAddDes, services)
    '                        dtDestinos.Rows.Add(data.DtDestino.Rows(0).ItemArray)

    '                        Dim StAddMat As New StADDMateriales(data.IDTipoOperacion, strNOperacion, dr("IDVino"), data.DtDestino.Rows(0)("IDArticulo") & String.Empty, _
    '                                                                     Nz(data.DtDestino.Rows(0)("Cantidad"), 0))
    '                        Dim dtMaterial2 As DataTable = ProcessServer.ExecuteTask(Of StADDMateriales, DataTable)(AddressOf ADDMateriales, StAddMat, services)

    '                        If Not dtMaterial2 Is Nothing AndAlso dtMaterial2.Rows.Count > 0 Then
    '                            If dtMaterial Is Nothing Then dtMaterial = dtMaterial2.Clone
    '                            For Each drMaterial As DataRow In dtMaterial2.Rows
    '                                dtMaterial.Rows.Add(drMaterial.ItemArray)
    '                            Next
    '                        End If
    '                End Select
    '            Next
    '            Dim StAddOpGen As New StADDOperacionGenerada(OperacionInfo, strNOperacion)
    '            ProcessServer.ExecuteTask(Of StADDOperacionGenerada)(AddressOf ADDOperacionGenerada, StAddOpGen, services)
    '        Else
    '            For Each dr As DataRow In dt.Rows
    '                Dim StNueva As New StNuevaOperacion(dr("IDDeposito"), dr("IDVino"), data.IDTipoOperacion, data.FechaOperacion)
    '                Dim strNOperacion As String = ProcessServer.ExecuteTask(Of StNuevaOperacion, String)(AddressOf NuevaOperacion, StNueva, services)

    '                Dim StAddOpGen As New StADDOperacionGenerada(OperacionInfo, strNOperacion)
    '                ProcessServer.ExecuteTask(Of StADDOperacionGenerada)(AddressOf ADDOperacionGenerada, StAddOpGen, services)
    '            Next
    '        End If

    '        Return OperacionInfo
    '    End If
    'End Function

    <Serializable()> _
    Public Class StADDOperacionGenerada
        Public OperacionInfo As OperacionUpdateData
        Public NOperacion As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal OperacionInfo As OperacionUpdateData, ByVal NOperacion As String)
            Me.OperacionInfo = OperacionInfo
            Me.NOperacion = NOperacion
        End Sub
    End Class

    <Task()> Public Shared Sub ADDOperacionGenerada(ByVal data As StADDOperacionGenerada, ByVal services As ServiceProvider)
        ReDim Preserve data.OperacionInfo.NOperacion(UBound(data.OperacionInfo.NOperacion) + 1)
        data.OperacionInfo.NOperacion(UBound(data.OperacionInfo.NOperacion)) = data.NOperacion
        ReDim Preserve data.OperacionInfo.Operaciones(UBound(data.OperacionInfo.Operaciones) + 1)
        data.OperacionInfo.Operaciones(UBound(data.OperacionInfo.Operaciones)) = data.NOperacion
    End Sub

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

    '<Task()> Public Shared Function ADDCabeceraOperacion(ByVal data As StADDCabeceraOperacion, ByVal services As ServiceProvider) As Dictionary(Of String, DataTable)
    '    Dim ClsBdgOper As New BdgOperacion
    '    Dim dtOperacion As DataTable = ClsBdgOper.AddNewForm

    '    If data.FechaOperacion = cnMinDate Then data.FechaOperacion = Date.Today
    '    dtOperacion.Rows(0)("Fecha") = New Date(data.FechaOperacion.Year, data.FechaOperacion.Month, data.FechaOperacion.Day, Today.Now.Hour, Today.Now.Minute, Today.Now.Second)
    '    dtOperacion.Rows(0)("IDTipoOperacion") = data.IDTipoOperacion

    '    Dim current As New BusinessData(dtOperacion.Rows(0))
    '    Dim context As New BusinessData
    '    context("ADDMaterialGlobalReal") = True
    '    context("ADDMaterialGlobalPrev") = True

    '    ClsBdgOper.ApplyBusinessRule("IDTipoOperacion", data.IDTipoOperacion, current, context)

    '    dtOperacion.Rows(0)("IDAnalisis") = current("IDAnalisis")
    '    dtOperacion.Rows(0)("ImputacionPrevMaterial") = current("ImputacionPrevMaterial")
    '    dtOperacion.Rows(0)("ImputacionPrevMod") = current("ImputacionPrevMod")
    '    dtOperacion.Rows(0)("ImputacionPrevCentro") = current("ImputacionPrevCentro")
    '    dtOperacion.Rows(0)("ImputacionPrevVarios") = current("ImputacionPrevVarios")

    '    dtOperacion.Rows(0)("ImputacionRealMaterial") = current("ImputacionRealMaterial")
    '    dtOperacion.Rows(0)("ImputacionRealMod") = current("ImputacionRealMod")
    '    dtOperacion.Rows(0)("ImputacionRealCentro") = current("ImputacionRealCentro")
    '    dtOperacion.Rows(0)("ImputacionRealVarios") = current("ImputacionRealVarios")

    '    Dim lst As New Dictionary(Of String, DataTable)
    '    lst.Add(GetType(BdgOperacion).Name, dtOperacion)
    '    If Not current("MaterialGlobal") Is Nothing Then
    '        lst.Add(GetType(BdgOperacionMaterial).Name, current("MaterialGlobal"))
    '    Else
    '        lst.Add(GetType(BdgOperacionMaterial).Name, Nothing)
    '    End If
    '    lst.Add(GetType(BdgOperacionMOD).Name, Nothing)
    '    lst.Add(GetType(BdgOperacionCentro).Name, Nothing)
    '    lst.Add(GetType(BdgOperacionVarios).Name, Nothing)
    '    lst.Add(GetType(BdgOperacionVino).Name & "Origen", Nothing)
    '    lst.Add(GetType(BdgOperacionVino).Name & "Destino", Nothing)
    '    lst.Add(GetType(BdgVinoMaterial).Name, Nothing)


    '    Return lst
    'End Function
    <Task()> Public Shared Function ADDCabeceraOperacion(ByVal data As StADDCabeceraOperacion, ByVal services As ServiceProvider) As Dictionary(Of String, DataTable)
        Dim Op As New BdgOperacion
        Dim dtOperacion As DataTable = Op.AddNewForm
        If data.FechaOperacion <> cnMinDate Then
            dtOperacion.Rows(0)("Fecha") = New Date(data.FechaOperacion.Year, data.FechaOperacion.Month, data.FechaOperacion.Day, Today.Now.Hour, Today.Now.Minute, Today.Now.Second)
        End If
        dtOperacion.Rows(0)("IDTipoOperacion") = data.IDTipoOperacion
        Dim context As New BusinessData
        context("OperacionNew") = True
        dtOperacion.Rows(0).ItemArray = Op.ApplyBusinessRule("IDTipoOperacion", data.IDTipoOperacion, dtOperacion.Rows(0), context).ItemArray

        Dim StDataImp As New BdgGeneral.StImputacion(dtOperacion.Rows(0)("IDTipoOperacion"), dtOperacion.Rows(0)("NOperacion"), enumBdgOrigenOperacion.Real, True, dtOperacion.Rows(0)("Fecha"))
        StDataImp = ProcessServer.ExecuteTask(Of BdgGeneral.StImputacion, BdgGeneral.StImputacion)(AddressOf BdgGeneral.Imputaciones, StDataImp, services)

        Dim lstDatosTipoOperacion As New Dictionary(Of String, DataTable)
        lstDatosTipoOperacion.Add(GetType(BdgOperacion).Name, dtOperacion)

        If Not StDataImp.DtMaterialesGlobal Is Nothing AndAlso StDataImp.DtMaterialesGlobal.Rows.Count > 0 Then lstDatosTipoOperacion.Add(GetType(BdgOperacionMaterial).Name, StDataImp.DtMaterialesGlobal)
        If Not StDataImp.DtCentrosGlobal Is Nothing AndAlso StDataImp.DtCentrosGlobal.Rows.Count > 0 Then lstDatosTipoOperacion.Add(GetType(BdgOperacionCentro).Name, StDataImp.DtCentrosGlobal)
        If Not StDataImp.DtMODGlobal Is Nothing AndAlso StDataImp.DtMODGlobal.Rows.Count > 0 Then lstDatosTipoOperacion.Add(GetType(BdgOperacionMOD).Name, StDataImp.DtMODGlobal)
        If Not StDataImp.DtVariosGlobal Is Nothing AndAlso StDataImp.DtVariosGlobal.Rows.Count > 0 Then lstDatosTipoOperacion.Add(GetType(BdgOperacionVarios).Name, StDataImp.DtVariosGlobal)

        If Not StDataImp.DtMateriales Is Nothing AndAlso StDataImp.DtMateriales.Rows.Count > 0 Then lstDatosTipoOperacion.Add(GetType(BdgVinoMaterial).Name, StDataImp.DtMateriales)
        If Not StDataImp.DtCentros Is Nothing AndAlso StDataImp.DtCentros.Rows.Count > 0 Then lstDatosTipoOperacion.Add(GetType(BdgVinoCentro).Name, StDataImp.DtCentros)
        If Not StDataImp.DtMOD Is Nothing AndAlso StDataImp.DtMOD.Rows.Count > 0 Then lstDatosTipoOperacion.Add(GetType(BdgVinoMod).Name, StDataImp.DtMOD)
        If Not StDataImp.DtVarios Is Nothing AndAlso StDataImp.DtVarios.Rows.Count > 0 Then lstDatosTipoOperacion.Add(GetType(BdgVinoVarios).Name, StDataImp.DtVarios)

        Return lstDatosTipoOperacion
    End Function


    <Serializable()> _
    Public Class StADDOrigenDestino
        Public IDDeposito As String
        Public NOperacion As String
        Public IDVino As Guid
        Public dtOperacion As DataTable

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDDeposito As String, ByVal NOperacion As String, ByVal IDVino As Guid)
            Me.IDDeposito = IDDeposito
            Me.NOperacion = NOperacion
            Me.IDVino = IDVino
        End Sub
    End Class

    <Task()> Public Shared Function ADDOrigen(ByVal data As StADDOrigenDestino, ByVal services As ServiceProvider) As DataTable
        Dim dtOrigen As DataTable = AdminData.GetData("frmBdgOperacionVinoOrigen", New NoRowsFilterItem)
        Dim drOrigen As DataRow = dtOrigen.NewRow

        drOrigen("IDDeposito") = data.IDDeposito
        drOrigen("NOperacion") = data.NOperacion

        Dim f As New Filter
        f.Add(New StringFilterItem("IDDeposito", data.IDDeposito))
        f.Add(New GuidFilterItem("IDVino", data.IDVino))

        Dim OpVino As New BdgOperacionVino
        Dim dtArticulo As DataTable = New BE.DataEngine().Filter("NegBdgOperacionPrevista", f, , "IDArticulo, Lote")
        If Not dtArticulo Is Nothing AndAlso dtArticulo.Rows.Count = 1 Then
            Dim context As New BusinessData
            If data.dtOperacion Is Nothing OrElse data.dtOperacion.Rows.Count = 0 Then
                data.dtOperacion = New BdgOperacion().SelOnPrimaryKey(data.NOperacion)
            End If
            If Not data.dtOperacion Is Nothing AndAlso data.dtOperacion.Rows.Count > 0 Then
                context = New BusinessData(data.dtOperacion.Rows(0))
            End If


            drOrigen("IDVino") = dtArticulo.Rows(0)("IDVino")
            drOrigen("IDArticulo") = dtArticulo.Rows(0)("IDArticulo")
            drOrigen = OpVino.ApplyBusinessRule("IDArticulo", dtArticulo.Rows(0)("IDArticulo"), drOrigen, context)

            drOrigen("IDUdMedida") = dtArticulo.Rows(0)("IDUdMedida")

            drOrigen("IDBarrica") = dtArticulo.Rows(0)("IDBarrica")
            drOrigen("IDEstadoVino") = dtArticulo.Rows(0)("IDEstadoVino")
            drOrigen("Lote") = dtArticulo.Rows(0)("Lote")
            drOrigen("Ocupacion") = dtArticulo.Rows(0)("Ocupacion")
            drOrigen("TipoDeposito") = dtArticulo.Rows(0)("TipoDeposito")
            drOrigen("Destino") = False

            drOrigen("Cantidad") = xRound(dtArticulo.Rows(0)("Cantidad"), 1)
            drOrigen = OpVino.ApplyBusinessRule("Cantidad", drOrigen("Cantidad"), drOrigen, context)

            'Dim dataUnidad As New UnidadAB.UnidadMedidaInfo
            'dataUnidad.IDUdMedidaA = drOrigen("IDUdMedida")
            'dataUnidad.IDUdMedidaB = New BdgParametro().UnidadesCampoLitros()
            'Dim dblFactor As Double = ProcessServer.ExecuteTask(Of UnidadAB.UnidadMedidaInfo, Double)(AddressOf UnidadAB.FactorDeConversion, dataUnidad, services)

            'drOrigen("Litros") = drOrigen("Cantidad") * dblFactor
            'Dim StGet1 As New StGetQDeposito(drOrigen("IDArticulo"), drOrigen("IDDeposito"), drOrigen("IDUdMedida"), drOrigen("Cantidad"))
            'drOrigen("QDeposito") = xRound(ProcessServer.ExecuteTask(Of StGetQDeposito, Double)(AddressOf GetQDeposito, StGet1, services), 1)
            'Dim StGet2 As New StGetQDeposito(drOrigen("IDArticulo"), drOrigen("IDDeposito"), drOrigen("IDUdMedida"), drOrigen("Ocupacion"))
            'drOrigen("Ocupacion") = xRound(ProcessServer.ExecuteTask(Of StGetQDeposito, Double)(AddressOf GetQDeposito, StGet2, services), 1)

            Dim StGet2 As New BdgGeneral.StGetQDeposito(drOrigen("IDArticulo"), drOrigen("IDDeposito"), drOrigen("IDUdMedida"), drOrigen("Ocupacion"))
            drOrigen("Ocupacion") = xRound(ProcessServer.ExecuteTask(Of BdgGeneral.StGetQDeposito, Double)(AddressOf BdgGeneral.GetQDeposito, StGet2, services), 1)

        ElseIf Not dtArticulo Is Nothing AndAlso dtArticulo.Rows.Count > 1 Then
            ApplicationService.GenerateError("El Depósito seleccionado tiene más de un Vino. Proceso cancelado.")
        Else : ApplicationService.GenerateError("El Depósito seleccionado no tiene Vino. Proceso cancelado.")
        End If
        dtOrigen.Rows.Add(drOrigen.ItemArray)
        Return dtOrigen
    End Function

    'TODO - Esto por qué mira las PREV?
    <Task()> Public Shared Function ADDDestino(ByVal data As StADDOrigenDestino, ByVal services As ServiceProvider) As DataTable
        Dim dtDestino As DataTable = AdminData.GetData("frmBdgOperacionVinoDestino", New NoRowsFilterItem)
        Dim drDestino As DataRow = dtDestino.NewRow
        drDestino("IDDeposito") = data.IDDeposito
        drDestino("NOperacion") = data.NOperacion

        Dim f As New Filter
        f.Add(New StringFilterItem("IDDeposito", data.IDDeposito))
        f.Add(New GuidFilterItem("IDVino", data.IDVino))

        Dim OpVino As New BdgOperacionVino
        Dim dtArticulo As DataTable = AdminData.GetData("NegBdgOperacionPrevista", f, , "IDArticulo, Lote")
        If Not dtArticulo Is Nothing AndAlso dtArticulo.Rows.Count = 1 Then

            Dim context As New BusinessData
            If data.dtOperacion Is Nothing OrElse data.dtOperacion.Rows.Count = 0 Then
                data.dtOperacion = New BdgOperacion().SelOnPrimaryKey(data.NOperacion)
            End If
            If Not data.dtOperacion Is Nothing AndAlso data.dtOperacion.Rows.Count > 0 Then
                context = New BusinessData(data.dtOperacion.Rows(0))
            End If

            drDestino("IDVino") = dtArticulo.Rows(0)("IDVino")
            drDestino("IDArticulo") = dtArticulo.Rows(0)("IDArticulo")
            drDestino = OpVino.ApplyBusinessRule("IDArticulo", dtArticulo.Rows(0)("IDArticulo"), drDestino, context)

            drDestino("IDUdMedida") = dtArticulo.Rows(0)("IDUdMedida")

            drDestino("Lote") = dtArticulo.Rows(0)("Lote")
            drDestino("Ocupacion") = dtArticulo.Rows(0)("Ocupacion")
            drDestino("Capacidad") = dtArticulo.Rows(0)("Capacidad")
            drDestino("TipoDeposito") = dtArticulo.Rows(0)("TipoDeposito")
            drDestino("Destino") = True

            drDestino("Cantidad") = xRound(dtArticulo.Rows(0)("Cantidad"), 1)
            drDestino = OpVino.ApplyBusinessRule("Cantidad", drDestino("Cantidad"), drDestino, context)

            'Dim dataUnidad As New UnidadAB.UnidadMedidaInfo
            'dataUnidad.IDUdMedidaA = drDestino("IDUdMedida")
            'dataUnidad.IDUdMedidaB = New BdgParametro().UnidadesCampoLitros()
            'Dim dblFactor As Double = ProcessServer.ExecuteTask(Of UnidadAB.UnidadMedidaInfo, Double)(AddressOf UnidadAB.FactorDeConversion, dataUnidad, services)

            'drDestino("Litros") = drDestino("Cantidad") * dblFactor
            'Dim StGet1 As New StGetQDeposito(drDestino("IDArticulo"), drDestino("IDDeposito"), drDestino("IDUdMedida"), drDestino("Cantidad"))
            'drDestino("QDeposito") = xRound(ProcessServer.ExecuteTask(Of StGetQDeposito, Double)(AddressOf GetQDeposito, StGet1, services), 1)
            'Dim StGet2 As New StGetQDeposito(drDestino("IDArticulo"), drDestino("IDDeposito"), drDestino("IDUdMedida"), drDestino("Ocupacion"))
            'drDestino("Ocupacion") = xRound(ProcessServer.ExecuteTask(Of StGetQDeposito, Double)(AddressOf GetQDeposito, StGet2, services), 1)

            Dim StGet2 As New BdgGeneral.StGetQDeposito(drDestino("IDArticulo"), drDestino("IDDeposito"), drDestino("IDUdMedida"), drDestino("Ocupacion"))
            drDestino("Ocupacion") = xRound(ProcessServer.ExecuteTask(Of BdgGeneral.StGetQDeposito, Double)(AddressOf BdgGeneral.GetQDeposito, StGet2, services), 1)
        ElseIf Not dtArticulo Is Nothing AndAlso dtArticulo.Rows.Count > 1 Then
            ApplicationService.GenerateError("El Depósito seleccionado tiene más de un Vino. Proceso cancelado.")
        Else : ApplicationService.GenerateError("El Depósito seleccionado no tiene Vino. Proceso cancelado.")
        End If
        dtDestino.Rows.Add(drDestino.ItemArray)
        Return dtDestino
    End Function

    <Serializable()> _
    Public Class StADDMateriales
        Public IDTipoOperacion As String
        Public NOperacion As String
        Public IDVino As Guid
        Public IDArticulo As String
        Public Cantidad As Double
        Public Imputar As Boolean
        Public FechaOperacion As Date

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDTipoOperacion As String, ByVal NOperacion As String, ByVal FechaOperacion As Date, ByVal IDVino As Guid, ByVal IDArticulo As String, ByVal Cantidad As Double, _
                       Optional ByVal Imputar As Boolean = False)
            Me.IDTipoOperacion = IDTipoOperacion
            Me.NOperacion = NOperacion
            Me.FechaOperacion = FechaOperacion
            Me.IDVino = IDVino
            Me.IDArticulo = IDArticulo
            Me.Cantidad = Cantidad
            Me.Imputar = Imputar
        End Sub
    End Class

    <Task()> Public Shared Function ADDMateriales(ByVal data As StADDMateriales, ByVal services As ServiceProvider) As DataTable
        Dim Stmat As New StMateriales(data.IDTipoOperacion, data.NOperacion, data.FechaOperacion, data.IDArticulo, data.Cantidad, , , , , True, data.Imputar)
        Dim dtMateriales As DataTable = ProcessServer.ExecuteTask(Of StMateriales, DataTable)(AddressOf Materiales, Stmat, services)
        If Not dtMateriales Is Nothing AndAlso dtMateriales.Rows.Count > 0 Then
            For Each drMateriales As DataRow In dtMateriales.Rows
                drMateriales("IDVino") = data.IDVino
                drMateriales("NOperacion") = data.NOperacion
                drMateriales("Fecha") = Date.Today
            Next
        End If
        Return dtMateriales
    End Function

#End Region

    <Task()> Public Shared Function GetVars(ByVal IDAnalisis As String, ByVal services As ServiceProvider) As DataTable
        Static sIDAnalisis As String
        Static sdtVars As DataTable
        If sIDAnalisis Is Nothing OrElse sIDAnalisis <> IDAnalisis Then
            sIDAnalisis = IDAnalisis
            Dim filtro As FilterItem = New FilterItem("IDAnalisis", FilterOperator.Equal, sIDAnalisis, FilterType.String)
            sdtVars = New DataEngine().Filter("frmBdgAnalisisVariable", filtro)
        End If
        Return sdtVars
    End Function

#End Region

#Region " Gestion Operaciones Reales "

#Region " CrearOperacion "

    '<Serializable()> _
    'Public Class StCrearOperacionInt
    '    Public DtOperacion As DataTable
    '    Public Origen As DataTable
    '    Public Destino As DataTable
    '    Public Analisis As DataTable
    '    Public Materiales As DataTable
    '    Public Centros As DataTable
    '    Public Operarios As DataTable
    '    Public Varios As DataTable
    '    Public DtMaterialGlobal As DataTable
    '    Public DtMODGlobal As DataTable
    '    Public DtCentroGlobal As DataTable
    '    Public DtVariosGlobal As DataTable

    '    Public DtLotesMateriales As DataTable
    '    Public DtLotesMaterialesGlobal As DataTable

    '    Public Sub New()
    '    End Sub

    '    Public Sub New(ByVal DtOperacion As DataTable, ByVal Origen As DataTable, ByVal Destino As DataTable, ByVal Analisis As DataTable, _
    '                   ByVal Materiales As DataTable, ByVal Centros As DataTable, ByVal Operarios As DataTable, ByVal Varios As DataTable, _
    '                   ByVal DtMaterialGlobal As DataTable, ByVal DtMODGlobal As DataTable, ByVal DtCentroGlobal As DataTable, ByVal DtVariosGlobal As DataTable, _
    '                   Optional ByVal dtLotesMateriales As DataTable = Nothing, Optional ByVal dtLotesMaterialesGlobal As DataTable = Nothing)
    '        Me.DtOperacion = DtOperacion
    '        Me.Origen = Origen
    '        Me.Destino = Destino
    '        Me.Analisis = Analisis
    '        Me.Materiales = Materiales
    '        Me.Centros = Centros
    '        Me.Operarios = Operarios
    '        Me.Varios = Varios
    '        Me.DtMaterialGlobal = DtMaterialGlobal
    '        Me.DtMODGlobal = DtMODGlobal
    '        Me.DtCentroGlobal = DtCentroGlobal
    '        Me.DtVariosGlobal = DtVariosGlobal

    '        Me.DtLotesMateriales = dtLotesMateriales
    '        Me.DtLotesMaterialesGlobal = dtLotesMaterialesGlobal
    '    End Sub
    'End Class

    '<Task()> Public Shared Sub CrearOperacionInt(ByVal data As StCrearOperacionInt, ByVal services As ServiceProvider)
    '    Dim NOperacion As String = data.DtOperacion.Rows(0)("NOperacion")
    '    Dim IDTipoOperacion As String = data.DtOperacion.Rows(0)("IDTipoOperacion")
    '    Dim Fecha As Date = data.DtOperacion.Rows(0)("Fecha")
    '    Dim FechaDoc As New Date(CDate(data.DtOperacion.Rows(0)("Fecha")).Year, CDate(data.DtOperacion.Rows(0)("Fecha")).Month, CDate(data.DtOperacion.Rows(0)("Fecha")).Day)
    '    Dim IDAnalisis As String = data.DtOperacion.Rows(0)("IDAnalisis") & String.Empty

    '    'Comprobar tipo de operación
    '    Dim oTO As BdgTipoOperacion = New BdgTipoOperacion
    '    Dim rwTO As DataRow = oTO.GetItemRow(IDTipoOperacion)

    '    Dim intTipoMov As Bodega.TipoMovimiento = rwTO(_TO.TipoMovimiento)

    '    Dim StValidar As New StValidar(intTipoMov, data.Origen, data.Destino, data.DtOperacion)
    '    Dim oColFrx As Hashtable = ProcessServer.ExecuteTask(Of StValidar, Hashtable)(AddressOf ValidarOperacion, StValidar, services)

    '    'Materiales Globales
    '    If Not data.DtMaterialGlobal Is Nothing Then
    '        Dim om As New BdgOperacionMaterial
    '        data.DtMaterialGlobal.TableName = om.Entity
    '        om.Update(data.DtMaterialGlobal)
    '    End If
    '    'Lotes Materiales Globales
    '    If Not data.DtLotesMaterialesGlobal Is Nothing Then
    '        Dim olm As New BdgOperacionMaterialLote
    '        data.DtLotesMaterialesGlobal.TableName = olm.Entity
    '        olm.Update(data.DtLotesMaterialesGlobal)
    '    End If
    '    'Mano de Obra Global
    '    If Not data.DtMODGlobal Is Nothing Then
    '        Dim omod As New BdgOperacionMOD
    '        data.DtMODGlobal.TableName = omod.Entity
    '        omod.Update(data.DtMODGlobal)
    '    End If
    '    'Centro Global
    '    If Not data.DtCentroGlobal Is Nothing Then
    '        Dim oCentro As New BdgOperacionCentro
    '        data.DtCentroGlobal.TableName = oCentro.Entity
    '        oCentro.Update(data.DtCentroGlobal)
    '    End If
    '    'Varios Global
    '    If Not data.DtVariosGlobal Is Nothing Then
    '        Dim oVarios As New BdgOperacionVarios
    '        data.DtVariosGlobal.TableName = oVarios.Entity
    '        oVarios.Update(data.DtVariosGlobal)
    '    End If

    '    '//Buscar un Estado de vino
    '    Dim IDEstadoVino As String
    '    If rwTO.IsNull("IDEstadoVino") Then
    '        If Not data.Origen Is Nothing Then
    '            For Each oRwO As DataRow In data.Origen.Rows
    '                If Not oRwO.IsNull(_V.IDEstadoVino) Then
    '                    IDEstadoVino = oRwO(_V.IDEstadoVino)
    '                    Exit For
    '                End If
    '            Next
    '        End If
    '    Else
    '        IDEstadoVino = rwTO("IDEstadoVino")
    '    End If

    '    'Crear cabecera de Operacion
    '    AdminData.BeginTx()

    '    Dim oVinoWC As New BdgWorkClass

    '    'Crear movimiento de vino
    '    If Not data.Origen Is Nothing Then
    '        For Each oRwO As DataRow In data.Origen.Rows
    '            Dim dblQ As Double = 0
    '            If Not oRwO.IsNull(_OV.Cantidad) Then dblQ = oRwO(_OV.Cantidad)
    '            Dim StCambio As New BdgWorkClass.StCambiarOcupacion(oRwO(_OV.IDVino), -dblQ)
    '            ProcessServer.ExecuteTask(Of BdgWorkClass.StCambiarOcupacion)(AddressOf BdgWorkClass.CambiarOcupacion, StCambio, services)
    '        Next
    '    End If

    '    If intTipoMov <> Business.Bodega.TipoMovimiento.SinMovimiento AndAlso Not data.Destino Is Nothing Then
    '        For Each oRwD As DataRow In data.Destino.Rows
    '            Dim oVinos() As VinoComponente
    '            If oColFrx Is Nothing Then
    '                oVinos = Nothing
    '            Else
    '                Dim StCalc As New ProcesoBdgOperacion.StCalcularEstructura(oColFrx(oRwD(_OV.IDVino)), oRwD(_V.IDArticulo), oRwD(_V.IDUdMedida), data.Origen)
    '                oVinos = ProcessServer.ExecuteTask(Of ProcesoBdgOperacion.StCalcularEstructura, VinoComponente())(AddressOf ProcesoBdgOperacion.CalcularEstructura, StCalc, services)
    '            End If
    '            If intTipoMov = Business.Bodega.TipoMovimiento.CrearOrigen Then
    '                Dim StCrear As New BdgWorkClass.StCrearEstructuraOperacion(oRwD(_OV.IDVino), BdgOrigenVino.Interno, NOperacion, oVinos)
    '                ProcessServer.ExecuteTask(Of BdgWorkClass.StCrearEstructuraOperacion)(AddressOf BdgWorkClass.CrearEstructuraOperacion, StCrear, services)
    '            Else
    '                Dim oldID As Guid
    '                If (Length(oRwD(_OV.IDVino)) > 0) Then
    '                    oldID = oRwD(_OV.IDVino)
    '                End If
    '                Dim StCrear As New BdgWorkClass.StCrearVino(oRwD(_V.IDDeposito), oRwD(_V.IDArticulo), oRwD(_V.Lote), Fecha, BdgOrigenVino.Interno, oRwD(_V.IDUdMedida), oRwD(_OV.Cantidad), IDEstadoVino, NOperacion, oVinos, , Nz(oRwD(_OV.IDBarrica)))
    '                oRwD(_OV.IDVino) = ProcessServer.ExecuteTask(Of BdgWorkClass.StCrearVino, Guid)(AddressOf BdgWorkClass.CrearVino, StCrear, services)
    '                'Dim Tables() As DataTable = New DataTable() {data.Analisis, data.Materiales, data.Centros, data.Operarios, data.Varios}
    '                Dim Tables As New List(Of DataTable) '=  New DataTable() {data.Analisis, data.Materiales, data.Centros, data.Operarios, data.Varios}
    '                Tables.AddRange(New DataTable() {data.Analisis, data.Materiales, data.Centros, data.Operarios, data.Varios})

    '                Dim StCambiar As New ProcesoBdgOperacion.DataCambiarIDVino(oldID, oRwD(_OV.IDVino), Tables)
    '                ProcessServer.ExecuteTask(Of ProcesoBdgOperacion.DataCambiarIDVino)(AddressOf ProcesoBdgOperacion.CambiarIDVino, StCambiar, services)
    '            End If
    '        Next
    '    End If

    '    Dim oOV As BdgOperacionVino = New BdgOperacionVino
    '    Dim dtOV As DataTable = oOV.AddNew
    '    If Not data.Origen Is Nothing Then
    '        Dim StADD As New StAddToDtOpVino(dtOV, NOperacion, data.Origen, False)
    '        ProcessServer.ExecuteTask(Of StAddToDtOpVino)(AddressOf AddToDTOpVino, StADD, services)
    '    End If
    '    Dim StADDDes As New StAddToDtOpVino(dtOV, NOperacion, data.Destino, True)
    '    ProcessServer.ExecuteTask(Of StAddToDtOpVino)(AddressOf AddToDTOpVino, StADDDes, services)

    '    '//Actualización de Ordenes de fabricación
    '    Dim StActua As New Business.Bodega.BdgOperacion.StActualizarOF(data.Destino, Fecha)
    '    ProcessServer.ExecuteTask(Of Business.Bodega.BdgOperacion.StActualizarOF)(AddressOf Business.Bodega.BdgOperacion.ActualizarOF, StActua, services)

    '    '//Actualización de dtOV despues de ejecución de movimientos de stock
    '    oOV.Update(dtOV)

    '    Dim StCrearMat As New BdgVinoMaterial.StCrearVinoMaterial(Fecha, NOperacion, data.Materiales)
    '    ProcessServer.ExecuteTask(Of BdgVinoMaterial.StCrearVinoMaterial)(AddressOf BdgVinoMaterial.CrearVinoMaterial, StCrearMat, services)

    '    'GUARDADO DE LOTES
    '    Dim stCrearMov As New BdgVinoMaterialLote.StCrearMovimientos(NOperacion, data.DtLotesMateriales, data.Materiales)
    '    ProcessServer.ExecuteTask(Of BdgVinoMaterialLote.StCrearMovimientos)(AddressOf BdgVinoMaterialLote.CrearMovimientos, stCrearMov, services)

    '    Dim StCrearCentro As New BdgVinoCentro.StCrearVinoCentro(Fecha, NOperacion, data.Centros)
    '    ProcessServer.ExecuteTask(Of BdgVinoCentro.StCrearVinoCentro)(AddressOf BdgVinoCentro.CrearVinoCentro, StCrearCentro, services)

    '    Dim StCrearMod As New BdgVinoMod.StCrearVinoMOD(Fecha, NOperacion, data.Operarios)
    '    ProcessServer.ExecuteTask(Of BdgVinoMod.StCrearVinoMOD)(AddressOf BdgVinoMod.CrearVinoMOD, StCrearMod, services)

    '    Dim StCrearAnalisis As New BdgVinoAnalisis.StCrearVinoAnalisis(IDAnalisis, NOperacion, Fecha, data.Analisis)
    '    ProcessServer.ExecuteTask(Of BdgVinoAnalisis.StCrearVinoAnalisis)(AddressOf BdgVinoAnalisis.CrearVinoAnalisis, StCrearAnalisis, services)

    '    Dim StCrearVarios As New BdgVinoVarios.StCrearVinoVarios(Fecha, NOperacion, data.Varios)
    '    ProcessServer.ExecuteTask(Of BdgVinoVarios.StCrearVinoVarios)(AddressOf BdgVinoVarios.CrearVinoVarios, StCrearVarios, services)

    '    '//Gestion de stock
    '    Dim StEj As New BdgWorkClass.StEjecutarMovimientos(NOperacion, FechaDoc)
    '    Dim NumMov As Integer = ProcessServer.ExecuteTask(Of BdgWorkClass.StEjecutarMovimientos, Integer)(AddressOf BdgWorkClass.EjecutarMovimientos, StEj, services)
    '    If NumMov <> 0 Then
    '        data.DtOperacion.Rows(0)("IDMovimiento") = NumMov
    '        'BusinessHelper.UpdateTable(rwOperacion.Table)
    '    End If
    '    If data.DtOperacion.Rows(0).RowState <> DataRowState.Unchanged Then
    '        BusinessHelper.UpdateTable(data.DtOperacion)
    '    End If

    '    '//Actualización de operacionplan
    '    If (data.DtOperacion.Columns.Contains("NOperacionPlan") AndAlso Length(data.DtOperacion.Rows(0)("NOperacionPlan")) > 0) Then
    '        Dim stData As New BdgOperacionPlan.StCambiarEstadoOperacion(data.DtOperacion.Rows(0)("NOperacionPlan"), BdgEstadoOperacionPlan.Confirmado)
    '        ProcessServer.ExecuteTask(Of BdgOperacionPlan.StCambiarEstadoOperacion)(AddressOf BdgOperacionPlan.CambiarEstadoOperacion, stData, services)
    '    End If

    'End Sub

    '<Serializable()> _
    'Public Class StCrearOperacion
    '    Public DtOperacion As DataTable
    '    Public DtOrigen As DataTable
    '    Public DtDestino As DataTable
    '    Public DtAnalisis As DataTable
    '    Public DtMateriales As DataTable
    '    Public DtLotesMateriales As DataTable
    '    Public DtCentros As DataTable
    '    Public DtOperarios As DataTable
    '    Public DtVarios As DataTable
    '    Public DtMaterialGlobal As DataTable
    '    Public DtLotesMaterialesGlobal As DataTable
    '    Public DtMODGlobal As DataTable
    '    Public DtCentroGlobal As DataTable
    '    Public DtVariosGlobal As DataTable

    '    Public Sub New()
    '    End Sub

    '    Public Sub New(ByVal DtOperacion As DataTable, ByVal DtOrigen As DataTable, ByVal DtDestino As DataTable, ByVal DtAnalisis As DataTable, _
    '                   ByVal DtMateriales As DataTable, ByVal DtCentros As DataTable, ByVal DtOperarios As DataTable, ByVal DtVarios As DataTable, _
    '                   ByVal DtMaterialGlobal As DataTable, ByVal DtMODGlobal As DataTable, ByVal DtCentroGlobal As DataTable, ByVal DtVariosGlobal As DataTable, _
    '                   Optional ByVal DtLotesMateriales As DataTable = Nothing, Optional ByVal DtLotesMaterialesGlobal As DataTable = Nothing)
    '        Me.DtOperacion = DtOperacion
    '        Me.DtOrigen = DtOrigen
    '        Me.DtDestino = DtDestino
    '        Me.DtAnalisis = DtAnalisis
    '        Me.DtMateriales = DtMateriales
    '        Me.DtCentros = DtCentros
    '        Me.DtOperarios = DtOperarios
    '        Me.DtVarios = DtVarios
    '        Me.DtMaterialGlobal = DtMaterialGlobal
    '        Me.DtMODGlobal = DtMODGlobal
    '        Me.DtCentroGlobal = DtCentroGlobal
    '        Me.DtVariosGlobal = DtVariosGlobal

    '        Me.DtLotesMateriales = DtLotesMateriales
    '        Me.DtLotesMaterialesGlobal = DtLotesMaterialesGlobal
    '    End Sub
    'End Class

    '<Task()> Public Shared Function CrearOperacion(ByVal data As StCrearOperacion, ByVal services As ServiceProvider) As String
    '    AdminData.BeginTx()
    '    If Length(data.DtOperacion.Rows(0)("IDContador")) > 0 AndAlso Length(data.DtOperacion.Rows(0)("NOperacion")) = 0 Then 'Si no tiene NOperacion 
    '        Dim strNOperacion As String = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, data.DtOperacion.Rows(0)("IDContador"), services)
    '        data.DtOperacion.Rows(0)("NOperacion") = strNOperacion
    '    ElseIf Length(data.DtOperacion.Rows(0)("IDContador")) > 0 AndAlso Length(data.DtOperacion.Rows(0)("NOperacion")) > 0 AndAlso data.DtOperacion.Rows(0).RowState = DataRowState.Added Then
    '        data.DtOperacion.Rows(0)("NOperacion") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, data.DtOperacion.Rows(0)("IDContador"), services)
    '    End If

    '    'Crear cabecera de Operacion
    '    data.DtOperacion.TableName = New BdgOperacion().Entity
    '    BusinessHelper.UpdateTable(data.DtOperacion)
    '    Dim StOpInt As New StCrearOperacionInt(data.DtOperacion, data.DtOrigen, data.DtDestino, data.DtAnalisis, data.DtMateriales, _
    '                                           data.DtCentros, data.DtOperarios, data.DtVarios, data.DtMaterialGlobal, data.DtMODGlobal, _
    '                                           data.DtCentroGlobal, data.DtVariosGlobal, data.DtLotesMateriales, data.DtLotesMaterialesGlobal)
    '    ProcessServer.ExecuteTask(Of StCrearOperacionInt)(AddressOf CrearOperacionInt, StOpInt, services)

    '    Return data.DtOperacion.Rows(0)("NOperacion") & String.Empty
    'End Function

    '<Serializable()> _
    'Public Class StCrearOperacionInfo
    '    Public Origen As DataTable
    '    Public Destino As DataTable
    '    Public Info As ModificacionesInfo

    '    Public Sub New()
    '    End Sub

    '    Public Sub New(ByVal Origen As DataTable, ByVal Destino As DataTable, ByVal Info As ModificacionesInfo)
    '        Me.Origen = Origen
    '        Me.Destino = Destino
    '        Me.Info = Info
    '    End Sub
    'End Class

    '<Task()> Public Shared Sub CrearOperacionInfo(ByVal data As StCrearOperacionInfo, ByVal services As ServiceProvider)
    '    Dim drOp As DataRow = data.Info.drOperacion
    '    Dim dtMateriales As DataTable = data.Info.dtMateriales
    '    Dim dtMaterialesLotes As DataTable = data.Info.dtMaterialesLotes
    '    Dim dtOperarios As DataTable = data.Info.dtMod
    '    Dim dtCentros As DataTable = data.Info.dtMaquinas
    '    Dim dtVarios As DataTable = data.Info.dtVarios
    '    Dim dtAnalisis As DataTable = data.Info.dtAnalisis
    '    Dim dtElaboracion As DataTable = data.Info.dtElaboracion
    '    Dim dtCosteNave As DataTable = data.Info.dtCosteNave
    '    Dim intTipoMov As Business.Bodega.TipoMovimiento = data.Info.TipoOperacion
    '    Dim strDepositoNew As String = data.Info.DepositoNew
    '    Dim strArticuloNew As String = data.Info.ArticuloNew
    '    Dim strLoteNew As String = data.Info.LoteNew
    '    Dim strBarricaNew As String = data.Info.BarricaNew

    '    Dim oVinoWC As New BdgWorkClass
    '    Dim strUDMedidaLitros As String = New BdgParametro().UnidadesCampoLitros

    '    If Not data.Destino Is Nothing Then
    '        Dim A As New Articulo
    '        For Each oRwD As DataRow In data.Destino.Rows
    '            If data.Info.IDVino.Equals(oRwD(_OV.IDVino)) Then
    '                If Length(strArticuloNew) > 0 Then
    '                    oRwD("IDArticulo") = strArticuloNew
    '                    Dim strUDMedida As String = A.GetItemRow(strArticuloNew)("IDUdInterna")
    '                    If oRwD("IDUdMedida") <> strUDMedida Then
    '                        oRwD("IDUdMedida") = strUDMedida
    '                        Dim dataFactor As New ArticuloUnidadAB.DatosFactorConversion(oRwD(_V.IDArticulo), strUDMedidaLitros, strUDMedida)
    '                        Dim Factor As Double = ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, dataFactor, services)
    '                        oRwD("Cantidad") = oRwD("Litros") * Factor
    '                        Dim StGet As New BdgGeneral.StGetQDeposito(oRwD("IDArticulo"), oRwD("IDDeposito"), oRwD("IDUdMedida"), oRwD("Cantidad"))
    '                        oRwD("QDeposito") = ProcessServer.ExecuteTask(Of BdgGeneral.StGetQDeposito, Double)(AddressOf BdgGeneral.GetQDeposito, StGet, services)
    '                    End If
    '                End If
    '                If Length(strLoteNew) > 0 Then
    '                    oRwD("Lote") = strLoteNew
    '                End If
    '                If Length(strBarricaNew) > 0 Then
    '                    oRwD("IDBarrica") = strBarricaNew
    '                End If
    '            End If
    '        Next
    '    End If

    '    Dim StVal As New StValidar(intTipoMov, data.Origen, data.Destino, drOp.Table)
    '    Dim oColFrx As Hashtable = ProcessServer.ExecuteTask(Of StValidar, Hashtable)(AddressOf ValidarOperacion, StVal, services)

    '    Dim dtDestino As DataTable = data.Destino.Clone
    '    For Each oRwD As DataRow In data.Destino.Rows
    '        If data.Info.IDVino.Equals(oRwD(_OV.IDVino)) Then
    '            Dim oldID As Guid = oRwD(_OV.IDVino)
    '            If Length(strDepositoNew) = 0 Then strDepositoNew = oRwD(_V.IDDeposito) & String.Empty
    '            If Length(strArticuloNew) = 0 Then strArticuloNew = oRwD(_V.IDArticulo) & String.Empty
    '            If intTipoMov <> Business.Bodega.TipoMovimiento.SinMovimiento Then
    '                Dim oVinos() As VinoComponente
    '                If oColFrx Is Nothing Then
    '                    oVinos = Nothing
    '                Else
    '                    Dim StCalc As New ProcesoBdgOperacion.StCalcularEstructura(oColFrx(oRwD(_OV.IDVino)), oRwD(_V.IDArticulo), oRwD(_V.IDUdMedida), data.Origen)
    '                    oVinos = ProcessServer.ExecuteTask(Of ProcesoBdgOperacion.StCalcularEstructura, VinoComponente())(AddressOf ProcesoBdgOperacion.CalcularEstructura, StCalc, services)
    '                End If
    '                If intTipoMov = Business.Bodega.TipoMovimiento.CrearOrigen Then
    '                    Dim StCrearEst As New BdgWorkClass.StCrearEstructuraOperacion(oRwD(_OV.IDVino), BdgOrigenVino.Interno, drOp("NOperacion"), oVinos)
    '                    ProcessServer.ExecuteTask(Of BdgWorkClass.StCrearEstructuraOperacion)(AddressOf BdgWorkClass.CrearEstructuraOperacion, StCrearEst, services)
    '                Else
    '                    Dim StCrearVin As New BdgWorkClass.StCrearVino(strDepositoNew, strArticuloNew, oRwD(_V.Lote), drOp("Fecha"), BdgOrigenVino.Interno, oRwD(_V.IDUdMedida), oRwD(_OV.Cantidad), oRwD("IDEstadoVino") & String.Empty, drOp("NOperacion"), oVinos, , Nz(oRwD(_OV.IDBarrica)))
    '                    oRwD(_OV.IDVino) = ProcessServer.ExecuteTask(Of BdgWorkClass.StCrearVino, Guid)(AddressOf BdgWorkClass.CrearVino, StCrearVin, services)
    '                    Dim DtTables() As DataTable = New DataTable() {dtAnalisis, dtMateriales, dtCentros, dtOperarios, dtVarios, dtElaboracion, dtCosteNave}
    '                    Dim StCambVino As New ProcesoBdgOperacion.StCambiarIDVino(oldID, oRwD(_OV.IDVino), DtTables)
    '                    ProcessServer.ExecuteTask(Of ProcesoBdgOperacion.StCambiarIDVino)(AddressOf ProcesoBdgOperacion.CambiarIDVino, StCambVino, services)
    '                End If
    '            ElseIf intTipoMov = Business.Bodega.TipoMovimiento.SinMovimiento Then
    '                'TODO no parece correcta esta forma de actuar. En el caso de una operación sim Mov, el IDVino debe venir dado. No se crea.
    '                'Si no tiene movimiento hay que buscar un IDVino que haya en ese deposito y cambiar  
    '                'por el IDVino actual
    '                Dim dtDV As DataTable = ProcessServer.ExecuteTask(Of String, DataTable)(AddressOf BdgDepositoVino.SelOnIDDeposito, strDepositoNew, services)
    '                If dtDV.Rows.Count > 0 Then
    '                    oRwD("IDVino") = dtDV.Rows(0)("IDVino")
    '                    Dim drDeposito As DataRow = New BdgDeposito().GetItemRow(strDepositoNew)
    '                    oRwD("Ocupacion") = drDeposito("Ocupacion")

    '                    Dim StTables() As DataTable = New DataTable() {dtAnalisis, dtMateriales, dtCentros, dtOperarios, dtVarios, dtElaboracion, dtCosteNave}
    '                    Dim StCambio As New ProcesoBdgOperacion.StCambiarIDVino(oldID, oRwD(_OV.IDVino), StTables)
    '                    ProcessServer.ExecuteTask(Of ProcesoBdgOperacion.StCambiarIDVino)(AddressOf ProcesoBdgOperacion.CambiarIDVino, StCambio, services)
    '                End If
    '            End If
    '            dtDestino.ImportRow(oRwD)
    '        End If
    '    Next

    '    Dim oOV As New BdgOperacionVino
    '    Dim dtOVNew As DataTable = oOV.AddNew

    '    Dim StADD As New StAddToDtOpVino(dtOVNew, drOp("NOperacion"), dtDestino, True)
    '    ProcessServer.ExecuteTask(Of StAddToDtOpVino)(AddressOf AddToDTOpVino, StADD, services)

    '    '//Actualización de Ordenes de fabricación
    '    Dim StActua As New Business.Bodega.BdgOperacion.StActualizarOF(dtDestino, drOp("Fecha"))
    '    ProcessServer.ExecuteTask(Of Business.Bodega.BdgOperacion.StActualizarOF)(AddressOf Business.Bodega.BdgOperacion.ActualizarOF, StActua, services)

    '    '//Actualización de dtOVNew despues de ejecución de movimientos de stock
    '    oOV.Update(dtOVNew)

    '    Dim StCrear As New BdgVinoMaterial.StCrearVinoMaterial(drOp("Fecha"), drOp("NOperacion"), dtMateriales)
    '    ProcessServer.ExecuteTask(Of BdgVinoMaterial.StCrearVinoMaterial)(AddressOf BdgVinoMaterial.CrearVinoMaterial, StCrear, services)
    '    If Not dtMaterialesLotes Is Nothing AndAlso dtMaterialesLotes.Rows.Count > 0 Then
    '        '  Dim lts As New BdgVinoMaterialLote
    '        For Each drLote As DataRow In dtMaterialesLotes.Rows
    '            drLote.SetAdded()
    '            drLote("IDLineaMovimiento") = System.DBNull.Value
    '        Next
    '        'lts.Update(dtMaterialesLotes)

    '        Dim stCrearMov As New BdgVinoMaterialLote.StCrearMovimientos(drOp("NOperacion"), dtMaterialesLotes, dtMateriales)
    '        ProcessServer.ExecuteTask(Of BdgVinoMaterialLote.StCrearMovimientos)(AddressOf BdgVinoMaterialLote.CrearMovimientos, stCrearMov, services)
    '    End If


    '    If Not IsNothing(dtElaboracion) AndAlso dtElaboracion.Rows.Count > 0 Then
    '        Dim oVEla As New BdgVinoMaterial
    '        For Each drElab As DataRow In dtElaboracion.Rows
    '            Dim Vendimia As String
    '            Dim drEHist() As DataRow = data.Info.dtElaboracionHist.Select("IDVinoMaterial='" & drElab("IDVinoMaterial").ToString & "'")
    '            If drEHist.Length > 0 Then
    '                Vendimia = drEHist(0)("Vendimia")
    '            End If
    '            Dim StCrearVend As New BdgVinoMaterial.StCrearCosteVendimiaVino(drOp("Fecha"), strDepositoNew, drElab("IDArticulo"), drElab("Precio"), Vendimia, drElab("IDVino").ToString)
    '            ProcessServer.ExecuteTask(Of BdgVinoMaterial.StCrearCosteVendimiaVino, DataTable)(AddressOf BdgVinoMaterial.CrearCosteVendimiaVino, StCrearVend, services)
    '        Next
    '    End If

    '    Dim StCrearCen1 As New BdgVinoCentro.StCrearVinoCentro(drOp("Fecha"), drOp("NOperacion"), dtCentros)
    '    ProcessServer.ExecuteTask(Of BdgVinoCentro.StCrearVinoCentro)(AddressOf BdgVinoCentro.CrearVinoCentro, StCrearCen1, services)
    '    Dim StCrearCen2 As New BdgVinoCentro.StCrearVinoCentro(drOp("Fecha"), String.Empty, dtCosteNave)
    '    ProcessServer.ExecuteTask(Of BdgVinoCentro.StCrearVinoCentro)(AddressOf BdgVinoCentro.CrearVinoCentro, StCrearCen2, services)

    '    Dim StCrearMod As New BdgVinoMod.StCrearVinoMOD(drOp("Fecha"), drOp("NOperacion"), dtOperarios)
    '    ProcessServer.ExecuteTask(Of BdgVinoMod.StCrearVinoMOD)(AddressOf BdgVinoMod.CrearVinoMOD, StCrearMod, services)

    '    If Length(drOp("IDAnalisis")) > 0 Then
    '        Dim StCrearAnal As New BdgVinoAnalisis.StCrearVinoAnalisis(drOp("IDAnalisis"), drOp("NOperacion"), drOp("Fecha"), dtAnalisis)
    '        ProcessServer.ExecuteTask(Of BdgVinoAnalisis.StCrearVinoAnalisis)(AddressOf BdgVinoAnalisis.CrearVinoAnalisis, StCrearAnal, services)
    '    End If

    '    Dim StCrearVar As New BdgVinoVarios.StCrearVinoVarios(drOp("Fecha"), drOp("NOperacion"), dtVarios)
    '    ProcessServer.ExecuteTask(Of BdgVinoVarios.StCrearVinoVarios)(AddressOf BdgVinoVarios.CrearVinoVarios, StCrearVar, services)

    '    '//Gestion de stock
    '    If data.Info.TipoOperacion <> Business.Bodega.TipoMovimiento.SinMovimiento Then
    '        Dim StEj As New BdgWorkClass.StEjecutarMovimientosNumero(drOp("IDMovimiento"), drOp("NOperacion"), drOp("Fecha"))
    '        ProcessServer.ExecuteTask(Of BdgWorkClass.StEjecutarMovimientosNumero)(AddressOf BdgWorkClass.EjecutarMovimientosNumero, StEj, services)
    '    End If
    'End Sub

    '<Serializable()> _
    'Public Class StCrearOperacionNOperacion
    '    Public NOperacion As String
    '    Public DtOrigen As DataTable
    '    Public DtDestino As DataTable
    '    Public DtAnalisis As DataTable
    '    Public DtMateriales As DataTable
    '    Public DtCentros As DataTable
    '    Public DtOperarios As DataTable
    '    Public DtVarios As DataTable
    '    Public DtMaterialGlobal As DataTable
    '    Public DtMODGlobal As DataTable
    '    Public DtCentroGlobal As DataTable
    '    Public DtVariosGlobal As DataTable

    '    Public Sub New()
    '    End Sub

    '    Public Sub New(ByVal NOperacion As String, ByVal DtOrigen As DataTable, ByVal DtDestino As DataTable, ByVal DtAnalisis As DataTable, _
    '                   ByVal DtMateriales As DataTable, ByVal DtCentros As DataTable, ByVal DtOperarios As DataTable, ByVal DtVarios As DataTable, _
    '                   ByVal DtMaterialGlobal As DataTable, ByVal DtMODGlobal As DataTable, ByVal DtCentroGlobal As DataTable, ByVal DtVariosGlobal As DataTable)
    '        Me.NOperacion = NOperacion
    '        Me.DtOrigen = DtOrigen
    '        Me.DtDestino = DtDestino
    '        Me.DtAnalisis = DtAnalisis
    '        Me.DtMateriales = DtMateriales
    '        Me.DtCentros = DtCentros
    '        Me.DtOperarios = DtOperarios
    '        Me.DtVarios = DtVarios
    '        Me.DtMaterialGlobal = DtMaterialGlobal
    '        Me.DtMODGlobal = DtMODGlobal
    '        Me.DtCentroGlobal = DtCentroGlobal
    '        Me.DtVariosGlobal = DtVariosGlobal
    '    End Sub
    'End Class

    '<Task()> Public Shared Sub CrearOperacionNOperacion(ByVal data As StCrearOperacionNOperacion, ByVal services As ServiceProvider)
    '    Dim StCrearOpInt As New StCrearOperacionInt(New BdgOperacion().GetItemRow(data.NOperacion).Table, data.DtOrigen, data.DtDestino, data.DtAnalisis, data.DtMateriales, data.DtCentros, data.DtOperarios, data.DtVarios, data.DtMaterialGlobal, data.DtMODGlobal, data.DtCentroGlobal, data.DtVariosGlobal)
    '    ProcessServer.ExecuteTask(Of StCrearOperacionInt)(AddressOf CrearOperacionInt, StCrearOpInt, services)
    'End Sub


#End Region

    <Serializable()> _
    Public Class StValidar
        Public TipoMov As Business.Bodega.TipoMovimiento
        Public Origen As DataTable
        Public Destino As DataTable
        Public DtOperacion As DataTable

        Public Sub New()
        End Sub

        Public Sub New(ByVal TipoMov As Business.Bodega.TipoMovimiento, ByVal Origen As DataTable, ByVal Destino As DataTable, ByVal DtOperacion As DataTable)
            Me.TipoMov = TipoMov
            Me.Origen = Origen
            Me.Destino = Destino
            Me.DtOperacion = DtOperacion
        End Sub
    End Class

    <Task()> Public Shared Function ValidarOperacion(ByVal data As StValidar, ByVal services As ServiceProvider) As Hashtable
        Dim cOperacionNoPermitida As String = "Operación no permitida por el Tipo de Operación."
        Dim oColFrx As Hashtable

        Dim rwsOrigenAux() As DataRow
        Dim rwsDestinoAux() As DataRow

        Dim CurrentRows As DataRowState = DataRowState.Added Or DataRowState.Modified Or DataRowState.Unchanged
        If Not data.Origen Is Nothing Then rwsOrigenAux = data.Origen.Select(Nothing, Nothing, DataViewRowState.CurrentRows)
        If Not data.Destino Is Nothing Then rwsDestinoAux = data.Destino.Select(Nothing, Nothing, DataViewRowState.CurrentRows)

        Select Case data.TipoMov
            Case Business.Bodega.TipoMovimiento.SinMovimiento, Business.Bodega.TipoMovimiento.SinOrigen
                If Not rwsOrigenAux Is Nothing AndAlso rwsOrigenAux.Length <> 0 Then ApplicationService.GenerateError(cOperacionNoPermitida)
                If rwsDestinoAux Is Nothing OrElse rwsDestinoAux.Length = 0 Then ApplicationService.GenerateError(cOperacionNoPermitida)
            Case Business.Bodega.TipoMovimiento.DeUnoAUno
                If rwsOrigenAux Is Nothing OrElse rwsOrigenAux.Length <> 1 Then ApplicationService.GenerateError(cOperacionNoPermitida)
                If rwsDestinoAux Is Nothing OrElse rwsDestinoAux.Length <> 1 Then ApplicationService.GenerateError(cOperacionNoPermitida)
            Case Business.Bodega.TipoMovimiento.DeUnoAVarios
                If rwsOrigenAux Is Nothing OrElse rwsOrigenAux.Length <> 1 Then ApplicationService.GenerateError(cOperacionNoPermitida)
                If rwsDestinoAux Is Nothing OrElse rwsDestinoAux.Length = 0 Then ApplicationService.GenerateError(cOperacionNoPermitida)
            Case Business.Bodega.TipoMovimiento.DeVariosAUno
                If rwsOrigenAux Is Nothing OrElse rwsOrigenAux.Length = 0 Then ApplicationService.GenerateError(cOperacionNoPermitida)
                If rwsDestinoAux Is Nothing OrElse rwsDestinoAux.Length <> 1 Then ApplicationService.GenerateError(cOperacionNoPermitida)
            Case Business.Bodega.TipoMovimiento.Salida
                If rwsOrigenAux Is Nothing OrElse rwsOrigenAux.Length = 0 Then ApplicationService.GenerateError(cOperacionNoPermitida)
                If Not rwsDestinoAux Is Nothing AndAlso rwsDestinoAux.Length > 0 Then ApplicationService.GenerateError(cOperacionNoPermitida)
            Case Business.Bodega.TipoMovimiento.DeVariosAVarios, Business.Bodega.TipoMovimiento.CrearOrigen
                If rwsOrigenAux Is Nothing OrElse rwsOrigenAux.Length <= 0 Then ApplicationService.GenerateError(cOperacionNoPermitida)
                If rwsDestinoAux Is Nothing OrElse rwsDestinoAux.Length <= 0 Then ApplicationService.GenerateError(cOperacionNoPermitida)
        End Select

        Dim oArt As New Articulo
        Dim oDep As New BdgDeposito
        Dim oParam As New BdgParametro
        Dim LoteDef As String = oParam.LotePorDefecto
        Dim LoteExplicitoEnBotellero As Boolean = oParam.LoteExplicitoEnBotellero


        'Comprobar que los artículos están bien configurados
        Dim V As New BdgVino
        For Each rwOri As DataRow In rwsOrigenAux
            ProcessServer.ExecuteTask(Of String)(AddressOf BdgVino.ValidarArticuloVino, rwOri(_V.IDArticulo), services)
        Next
        For Each rwDst As DataRow In rwsDestinoAux
            ProcessServer.ExecuteTask(Of String)(AddressOf BdgVino.ValidarArticuloVino, rwDst(_V.IDArticulo), services)
        Next
        'Comprobar que la Fecha de la Operación sea igual o posterior a la Fecha de los vinos de Origen.
        If data.Origen.Columns.Contains("_V.Fecha") Then
            For Each rwOri As DataRow In rwsOrigenAux
                Dim StVal As New ProcesoBdgOperacion.StValidarFechaOperacion(rwOri(_V.Fecha), data.DtOperacion.Rows(0)("Fecha"))
                ProcessServer.ExecuteTask(Of ProcesoBdgOperacion.StValidarFechaOperacion)(AddressOf ProcesoBdgOperacion.ValidarFechaOperacion, StVal, services)
            Next
        Else
            For Each rwOri As DataRow In rwsOrigenAux
                Dim StVal As New ProcesoBdgOperacion.StValidarFechaOperacion(ProcessServer.ExecuteTask(Of Guid, Date)(AddressOf BdgVino.ObtenerFechaVino, rwOri(_V.IDVino), services), data.DtOperacion.Rows(0)("Fecha"))
                ProcessServer.ExecuteTask(Of ProcesoBdgOperacion.StValidarFechaOperacion)(AddressOf ProcesoBdgOperacion.ValidarFechaOperacion, StVal, services)
            Next
        End If

        'Comprobar que un Vino no esté en Origen y Destino
        For Each rwOri As DataRow In rwsOrigenAux
            Dim StVal As New StValidarVinoRepetidoOperacion(rwOri, rwsDestinoAux)
            ProcessServer.ExecuteTask(Of StValidarVinoRepetidoOperacion)(AddressOf ValidarVinoRepetidoOperacion, StVal, services)
        Next
        For Each rwDst As DataRow In rwsDestinoAux
            Dim StVal As New StValidarVinoRepetidoOperacion(rwDst, rwsOrigenAux)
            ProcessServer.ExecuteTask(Of StValidarVinoRepetidoOperacion)(AddressOf ValidarVinoRepetidoOperacion, StVal, services)
        Next

        '//comprobar la asignación lotes en función del tipo de deposito 
        For Each rwDst As DataRow In rwsDestinoAux
            If rwDst.IsNull(_D.IDDeposito) Then ApplicationService.GenerateError("No se ha especificado un depósito destino")
            Dim rwDep As DataRow = oDep.GetItemRow(rwDst(_D.IDDeposito))
            Dim td As TipoDeposito = rwDep(_D.TipoDeposito)
            If td = TipoDeposito.Almacen OrElse td = TipoDeposito.Botellero Then
                If LoteExplicitoEnBotellero AndAlso rwDst.IsNull(_V.Lote) Then
                    ApplicationService.GenerateError("No se ha especificado un lote para el artículo {0}", rwDst(_V.IDArticulo))
                End If
            ElseIf td = TipoDeposito.Barricas Then
                If rwDep(_D.UsarBarricaComoLote) And data.TipoMov <> Business.Bodega.TipoMovimiento.SinMovimiento Then
                    If rwDst.IsNull(_OV.IDBarrica) Then
                        ApplicationService.GenerateError("No se ha especificado un lote de barricas para el artículo {0}", rwDst(_V.IDArticulo))
                    Else
                        rwDst(_V.Lote) = rwDst(_OV.IDBarrica)
                    End If
                End If
            End If
            If rwDst.IsNull(_V.Lote) Then rwDst(_V.Lote) = LoteDef
        Next

        'TODO//Verificar los lotes de Origen
        'For Each oRw As DataRow In rwsOrigenAux
        '    Dim oVL As BdgVinoLote
        '    Dim rwD As DataRow = oDep.GetItemRow(oRw(_D.IDDeposito))
        '    Dim td As TipoDeposito = rwD(_D.TipoDeposito)
        '    If td = TipoDeposito.Almacen OrElse td = TipoDeposito.Botellero Then
        '        If oRw.IsNull(_OV.Lote) Then
        '            ApplicationService.GenerateError("No se ha especificado un lote")
        '        Else
        '            If oVL Is Nothing Then oVL = New BdgVinoLote
        '            If oVL.SelOnPrimaryKey(oRw(_OV.IDVino), oRw(_OV.Lote)).Rows.Count = 0 Then
        '                ApplicationService.GenerateError("No se encontró el lote |", oRw(_OV.Lote))
        '            End If
        '        End If
        '    Else
        '        oRw(_OV.Lote) = LoteDef
        '    End If
        'Next

        'comprobar los tipos de depositos permitidos por el tipo de operacion

        'Comprobar cantidad Origen = cantidad Destino
        If data.TipoMov <> Business.Bodega.TipoMovimiento.SinMovimiento And data.TipoMov <> Business.Bodega.TipoMovimiento.SinOrigen And data.TipoMov <> Business.Bodega.TipoMovimiento.Salida Then
            Dim IDUdMedida As String
            Dim dblQ As Double

            For Each oRw As DataRow In rwsOrigenAux
                If oRw.IsNull(_V.IDUdMedida) Then ApplicationService.GenerateError("El artículo no tiene una unidad de medida")
                If Len(IDUdMedida) = 0 Then IDUdMedida = oRw(_V.IDUdMedida)
                Dim dblQO As Double = 0
                If Not oRw.IsNull(_OV.Cantidad) Then dblQO = oRw(_OV.Cantidad)
                'If Not oRw.IsNull(_OV.Merma) AndAlso oRw(_OV.Merma) < 0 Then dblQO -= oRw(_OV.Merma)
                If Not oRw.IsNull(_OV.Merma) Then dblQO -= oRw(_OV.Merma)
                If IDUdMedida <> oRw(_V.IDUdMedida) Then
                    Dim dataFactor As New ArticuloUnidadAB.DatosFactorConversion(oRw(_V.IDArticulo), oRw(_V.IDUdMedida), IDUdMedida)
                    Dim f As Double = ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, dataFactor, services)
                    dblQ += dblQO * f
                Else
                    dblQ += dblQO
                End If
            Next

            Dim dblQNetaTot As Double = dblQ

            oColFrx = New Hashtable

            For Each oRw As DataRow In rwsDestinoAux
                Dim dblQD As Double
                If oRw.IsNull(_V.IDUdMedida) Then ApplicationService.GenerateError("El artículo no tiene una unidad de medida")
                If IDUdMedida <> oRw(_V.IDUdMedida) Then
                    Dim dataFactor As New ArticuloUnidadAB.DatosFactorConversion(oRw(_V.IDArticulo), oRw(_V.IDUdMedida), IDUdMedida)
                    Dim f As Double = ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, dataFactor, services)
                    dblQD = oRw(_OV.Cantidad) * f
                Else
                    dblQD = oRw(_OV.Cantidad)
                End If
                dblQ -= dblQD
                If dblQNetaTot <> 0 Then oColFrx.Add(oRw(_OV.IDVino), dblQD / dblQNetaTot)
            Next

            If dblQ * Math.Sign(dblQ) >= 0.01 Then ApplicationService.GenerateError("No coinciden las cantidades Origen y Destino")
        End If

        Return oColFrx
    End Function

    '<Serializable()> _
    'Public Class StCambiarIDVino
    '    Public OldID As Guid
    '    Public NewID As Guid
    '    Public Tables() As DataTable

    '    Public Sub New()
    '    End Sub

    '    Public Sub New(ByVal OldID As Guid, ByVal NewID As Guid, ByVal Tables() As DataTable)
    '        Me.OldID = OldID
    '        Me.NewID = NewID
    '        Me.Tables = Tables
    '    End Sub
    'End Class

    ''TODO se puede evitar esta función?
    '<Task()> Public Shared Sub CambiarIDVino(ByVal data As StCambiarIDVino, ByVal services As ServiceProvider)
    '    For Each dt As DataTable In data.Tables
    '        If Not dt Is Nothing Then
    '            Dim dc As DataColumn = dt.Columns(_V.IDVino)
    '            For Each oRw As DataRow In dt.Select(AdoFilterComposer.ComposeGuidFilter(_V.IDVino, FilterOperator.Equal, data.OldID))
    '                oRw(dc) = data.NewID
    '            Next
    '        End If
    '    Next
    'End Sub

    '<Serializable()> _
    'Public Class StCalcularEstructura
    '    Public FactorComp As Double
    '    Public IDArtDst As String
    '    Public IDUDMedidaDst As String
    '    Public Origen As DataTable

    '    Public Sub New()
    '    End Sub

    '    Public Sub New(ByVal FactorComp As Double, ByVal IDArtDst As String, ByVal IDUDMedidaDst As String, ByVal Origen As DataTable)
    '        Me.FactorComp = FactorComp
    '        Me.IDArtDst = IDArtDst
    '        Me.IDUDMedidaDst = IDUDMedidaDst
    '        Me.Origen = Origen
    '    End Sub
    'End Class

    '<Task()> Public Shared Function CalcularEstructura(ByVal data As StCalcularEstructura, ByVal services As ServiceProvider) As VinoComponente()
    '    Dim aVC(-1) As VinoComponente
    '    If Not data.Origen Is Nothing Then
    '        ReDim aVC(data.Origen.Rows.Count - 1)
    '        Dim oArt As Articulo = New Articulo
    '        For i As Integer = 0 To data.Origen.Rows.Count - 1
    '            Dim oRw As DataRow = data.Origen.Rows(i)

    '            Dim Merma As Double
    '            Dim Cantidad As Double

    '            If oRw.IsNull(_OV.Cantidad) Then
    '                Cantidad = 0
    '            Else
    '                Cantidad = oRw(_OV.Cantidad) * data.FactorComp
    '            End If
    '            If oRw.IsNull(_OV.Merma) Then
    '                Merma = 0
    '            Else
    '                Merma = oRw(_OV.Merma) * data.FactorComp
    '            End If
    '            If Merma > 0 Then Cantidad = Cantidad - Merma

    '            aVC(i) = New VinoComponente(oRw(_OV.IDVino), Cantidad, 1)
    '            Dim dataFactor As New ArticuloUnidadAB.DatosFactorConversion(data.IDArtDst, oRw(_V.IDUdMedida), data.IDUDMedidaDst)
    '            aVC(i).Factor = ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, dataFactor, services)
    '            aVC(i).Merma = Merma
    '        Next
    '    End If

    '    Return aVC
    'End Function

#Region " ModificarOperacion "

    '<Serializable()> _
    'Public Class StModificarOperacion
    '    Public Operacion As DataTable
    '    Public Origen As DataTable
    '    Public Destino As DataTable
    '    Public Analisis As DataTable
    '    Public Materiales As DataTable
    '    Public Centros As DataTable
    '    Public Operarios As DataTable
    '    Public Varios As DataTable
    '    Public DtMaterialGlobal As DataTable
    '    Public DtMODGlobal As DataTable
    '    Public DtCentrosGlobal As DataTable
    '    Public DtVariosGlobal As DataTable

    '    Public DtMaterialLote As DataTable
    '    Public DtMaterialLoteGlobal As DataTable

    '    Public Sub New()
    '    End Sub

    '    Public Sub New(ByVal Operacion As DataTable, ByVal Origen As DataTable, ByVal Destino As DataTable, ByVal Analisis As DataTable, _
    '                   ByVal Materiales As DataTable, ByVal Centros As DataTable, ByVal Operarios As DataTable, ByVal Varios As DataTable, _
    '                   ByVal DtMaterialGlobal As DataTable, ByVal DtMODGlobal As DataTable, ByVal DtCentrosGlobal As DataTable, ByVal DtVariosGlobal As DataTable, _
    '                   Optional ByVal dtMaterialLote As DataTable = Nothing, Optional ByVal dtMaterialLoteGlobal As DataTable = Nothing)
    '        Me.Operacion = Operacion
    '        Me.Origen = Origen
    '        Me.Destino = Destino
    '        Me.Analisis = Analisis
    '        Me.Materiales = Materiales
    '        Me.Centros = Centros
    '        Me.Operarios = Operarios
    '        Me.Varios = Varios
    '        Me.DtMaterialGlobal = DtMaterialGlobal
    '        Me.DtMODGlobal = DtMODGlobal
    '        Me.DtCentrosGlobal = DtCentrosGlobal
    '        Me.DtVariosGlobal = DtVariosGlobal

    '        Me.DtMaterialLote = dtMaterialLote
    '        Me.DtMaterialLoteGlobal = dtMaterialLoteGlobal
    '    End Sub
    'End Class

    '<Task()> Public Shared Sub ModificarOperacion(ByVal data As StModificarOperacion, ByVal services As ServiceProvider)
    '    '//Comprobar tipo de operación
    '    Dim rwOp As DataRow = data.Operacion.Rows(0)
    '    If rwOp.RowState = DataRowState.Modified Then
    '        If rwOp("IDTipoOperacion") <> rwOp("IDTipoOperacion", DataRowVersion.Original) Then
    '            ApplicationService.GenerateError("No se puede cambiar el tipo de operación de una operación.")
    '        End If
    '    End If

    '    'Materiales Globales
    '    ProcessServer.ExecuteTask(Of DataTable)(AddressOf guardarMaterialesGlobales, data.DtMaterialGlobal, services)
    '    'Mano de Obra
    '    ProcessServer.ExecuteTask(Of DataTable)(AddressOf guardarMODGlobales, data.DtMODGlobal, services)
    '    'Centro
    '    ProcessServer.ExecuteTask(Of DataTable)(AddressOf guardarCentroGlobales, data.DtCentrosGlobal, services)
    '    'Varios
    '    ProcessServer.ExecuteTask(Of DataTable)(AddressOf guardarVariosGlobales, data.DtVariosGlobal, services)

    '    ''-------------------------------------<Murrieta>-----------------------------------------------------------------''
    '    ''Problema: Cuando se Modifica un deposito un articulo o incluso Lote, 
    '    'abajo tenemos las siguiente sentencia:
    '    '---------------------If Not dtUpdtDst Is Nothing Then oOV.Update(dtUpdtDst)-------------------------
    '    'El update no esta sobreescrito en BdgOperacionVino y no siempre actuliza los datos modificados correctamente
    '    'A veces en el update no se modifica el IDVino y otra veces si. Para solucionar este problema se optado por:
    '    'Poner la linea que se a modificado en estado borrado y añadir la misma linea que se a puesto en estado borrada,
    '    'de modo que la nueva linea estara en estado added y con los mismos datos que la que hemos puesto en estado borrado.
    '    'El update no da problemas al tener que añadir registros.
    '    'El update solo funciona mal cuando las modificaciones se han dado en Articulo, Almacen o Fecha. Porque habria que 
    '    'modificar el IdVino, mientras que cuando los cambios solo se dan en Cantidades el update funciona Correctamente

    '    Dim DestinoMod As DataTable = data.Destino.GetChanges(DataRowState.Modified)
    '    If Not DestinoMod Is Nothing AndAlso DestinoMod.Rows.Count > 0 Then
    '        For Each drDesMod As DataRow In DestinoMod.Rows
    '            If drDesMod("IdDeposito") <> drDesMod("IdDeposito", DataRowVersion.Original) _
    '            OrElse drDesMod("IdArticulo") <> drDesMod("IdArticulo", DataRowVersion.Original) Then
    '                data.Destino.Rows.Add(drDesMod.ItemArray)
    '            End If
    '        Next
    '    End If
    '    If Not data.Destino Is Nothing AndAlso data.Destino.Rows.Count > 0 Then
    '        For Each drDestino As DataRow In data.Destino.Rows
    '            If drDestino.RowState = DataRowState.Modified Then
    '                If drDestino("IdDeposito") <> drDestino("IdDeposito", DataRowVersion.Original) _
    '                OrElse drDestino("IdArticulo") <> drDestino("IdArticulo", DataRowVersion.Original) Then
    '                    drDestino.Delete()
    '                End If
    '            End If
    '        Next
    '    End If
    '    ''--------------------------------------------------------------------------------------------------------------''
    '    Dim NOperacion As String = rwOp("NOperacion")
    '    Dim Fecha As Date = rwOp("Fecha")
    '    Dim FechaDoc As New Date(CDate(rwOp("Fecha")).Year, CDate(rwOp("Fecha")).Month, CDate(rwOp("Fecha")).Day)

    '    Dim oTO As BdgTipoOperacion = New BdgTipoOperacion
    '    Dim rwTO As DataRow = oTO.GetItemRow(rwOp("IDTipoOperacion"))
    '    Dim IDEstadoVino As String
    '    If Not rwTO.IsNull(_TO.IDEstadoVino) Then IDEstadoVino = rwTO(_TO.IDEstadoVino)

    '    Dim intTipoMov As Bodega.TipoMovimiento = rwTO(_TO.TipoMovimiento)

    '    Dim StVal As New StValidar(intTipoMov, data.Origen, data.Destino, rwOp.Table)
    '    Dim oColFrx As Hashtable = ProcessServer.ExecuteTask(Of StValidar, Hashtable)(AddressOf ValidarOperacion, StVal, services)

    '    'Crear cabecera de Operacion
    '    AdminData.BeginTx()

    '    ''''''''''''''''''''''''''si cambian la fecha cambia la consulta movimientos
    '    If Fecha <> rwOp("Fecha", DataRowVersion.Original) Then
    '        'Hay que comprobar que el tipo de operacion tenga puesto que controle los movimientos
    '        Dim TipoO As New BdgTipoOperacion
    '        Dim dtTipoOperacion As DataTable = TipoO.SelOnPrimaryKey(rwOp("IDTipoOperacion"))

    '        'Cambiamos la fecha a los movimientos asociados a Origen y Destino si la operación genera mvtos.
    '        If Not dtTipoOperacion Is Nothing AndAlso dtTipoOperacion.Rows.Count > 0 AndAlso _
    '        dtTipoOperacion.Rows(0)("TipoMovimiento") <> Business.Bodega.TipoMovimiento.SinMovimiento Then
    '            Dim StModif As New StModificarFechaOperacion(NOperacion, Fecha, rwOp("IDMovimiento"), rwTO.Table)
    '            ProcessServer.ExecuteTask(Of StModificarFechaOperacion)(AddressOf ModificarFechaOperacion, StModif, services)
    '        End If

    '        'Cambiamos la fecha a los movimientos asociados a los materiales.
    '        Dim VinoMatLote As New BdgVinoMaterialLote
    '        Dim VinoMat As New BdgVinoMaterial
    '        Dim fil As New Filter
    '        fil.Add(New StringFilterItem("NOperacion", FilterOperator.Equal, rwOp("NOperacion")))
    '        Dim dtVinoMaterial As DataTable = VinoMat.Filter(fil)
    '        If Not dtVinoMaterial Is Nothing AndAlso dtVinoMaterial.Rows.Count > 0 Then
    '            For Each drVinoMaterial As DataRow In dtVinoMaterial.Rows
    '                If Length(drVinoMaterial("IDLineaMovimiento")) > 0 Then
    '                    Dim dataCorreccion As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Corregir, drVinoMaterial("IDLineaMovimiento"), FechaDoc, False)
    '                    ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento)(AddressOf ProcesoStocks.ActualizarMovimiento, dataCorreccion, services)
    '                Else
    '                    Dim filAux As New Filter
    '                    filAux.Add("IDVinoMaterial", drVinoMaterial("IDVinoMaterial"))
    '                    Dim dtVinoMaterialLote As DataTable = VinoMatLote.Filter(filAux)
    '                    If Not dtVinoMaterialLote Is Nothing AndAlso dtVinoMaterialLote.Rows.Count > 0 Then
    '                        For Each drVinoMaterialLote As DataRow In dtVinoMaterialLote.Rows
    '                            If Length(drVinoMaterialLote("IDLineaMovimiento")) > 0 Then
    '                                Dim dataCorreccion As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Corregir, drVinoMaterialLote("IDLineaMovimiento"), FechaDoc, False)
    '                                ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento)(AddressOf ProcesoStocks.ActualizarMovimiento, dataCorreccion, services)
    '                            End If
    '                        Next
    '                    End If
    '                End If
    '            Next
    '        End If
    '    End If

    '    Dim ClsOp As New BdgOperacion
    '    ClsOp.Update(data.Operacion)

    '    'Crear movimiento de vino
    '    Dim oVwc As New BdgWorkClass
    '    Dim blnOrgChngd As Boolean

    '    Dim OrigenCurrents, OrigenDeleteds As DataTable

    '    If Not data.Origen Is Nothing Then
    '        '//Datos a usar para calcular las nuevas estructuras
    '        OrigenCurrents = data.Origen.GetChanges(DataRowState.Added Or DataRowState.Modified Or DataRowState.Unchanged)
    '        OrigenDeleteds = data.Origen.GetChanges(DataRowState.Deleted)

    '        If Not OrigenCurrents Is Nothing Then
    '            For Each rwO As DataRow In OrigenCurrents.Rows
    '                Dim dblQAnt As Double = 0
    '                Dim dblQAct As Double = 0
    '                If rwO.HasVersion(DataRowVersion.Original) Then
    '                    dblQAnt = rwO(_OV.Cantidad, DataRowVersion.Original)
    '                    'If rwO(_OV.Merma, DataRowVersion.Original) > 0 Then dblQAnt += rwO(_OV.Merma, DataRowVersion.Original)
    '                End If
    '                If rwO.HasVersion(DataRowVersion.Current) Then
    '                    If Not rwO.IsNull(_OV.Cantidad) Then dblQAct = rwO(_OV.Cantidad)
    '                    'If Not rwO.IsNull(_OV.Merma) AndAlso rwO(_OV.Merma) > 0 Then dblQAct += rwO(_OV.Merma)
    '                End If

    '                If rwO.RowState = DataRowState.Added Or rwO.RowState = DataRowState.Modified Or rwO.RowState = DataRowState.Deleted Then
    '                    If rwO.RowState = DataRowState.Added Then
    '                        rwO(_OV.NOperacion) = NOperacion
    '                        rwO(_OV.Destino) = False
    '                        blnOrgChngd = True

    '                    ElseIf rwO.RowState = DataRowState.Modified Then
    '                        If CType(rwO(_OV.IDVino), Guid).Equals(rwO(_OV.IDVino, DataRowVersion.Original)) Then
    '                            dblQAct = dblQAct - dblQAnt
    '                            dblQAnt = 0
    '                        End If
    '                        blnOrgChngd = True
    '                    End If
    '                    If dblQAnt <> 0 Then
    '                        Dim StCambio As New BdgWorkClass.StCambiarOcupacion(rwO(_OV.IDVino, DataRowVersion.Original), dblQAnt)
    '                        ProcessServer.ExecuteTask(Of BdgWorkClass.StCambiarOcupacion)(AddressOf BdgWorkClass.CambiarOcupacion, StCambio, services)
    '                    End If
    '                    If dblQAct <> 0 Then
    '                        Dim StCambio As New BdgWorkClass.StCambiarOcupacion(rwO(_OV.IDVino), -dblQAct)
    '                        ProcessServer.ExecuteTask(Of BdgWorkClass.StCambiarOcupacion)(AddressOf BdgWorkClass.CambiarOcupacion, StCambio, services)
    '                    End If
    '                End If
    '            Next
    '        End If

    '        If Not blnOrgChngd Then blnOrgChngd = Not OrigenDeleteds Is Nothing
    '    End If

    '    'Se ha cambiado el orden, porque si se da el caso de que se meta una nueva linea de destino, a la cual se le impute
    '    'un material, un coste,... Daba un error porque se intentaba meter un material de un IdVino que no estaba dado de alta.
    '    Dim oOV As New BdgOperacionVino
    '    Dim v As New BdgVino
    '    If Not data.Destino Is Nothing Then
    '        For Each rwD As DataRow In data.Destino.Rows
    '            If rwD.RowState = DataRowState.Added Then
    '                Dim oVinos() As VinoComponente
    '                If oColFrx Is Nothing Then
    '                    oVinos = Nothing
    '                Else
    '                    Dim StCalc As New StCalcularEstructura(oColFrx(rwD(_OV.IDVino)), rwD(_V.IDArticulo), rwD(_V.IDUdMedida), OrigenCurrents)
    '                    oVinos = ProcessServer.ExecuteTask(Of StCalcularEstructura, VinoComponente())(AddressOf CalcularEstructura, StCalc, services)
    '                End If
    '                If intTipoMov = Business.Bodega.TipoMovimiento.SinMovimiento Then
    '                ElseIf intTipoMov = Business.Bodega.TipoMovimiento.CrearOrigen Then
    '                    Dim StCrear As New BdgWorkClass.StCrearEstructuraOperacion(rwD(_OV.IDVino), BdgOrigenVino.Interno, NOperacion, oVinos)
    '                    ProcessServer.ExecuteTask(Of BdgWorkClass.StCrearEstructuraOperacion)(AddressOf BdgWorkClass.CrearEstructuraOperacion, StCrear, services)
    '                Else
    '                    Dim oldID As Guid = rwD(_OV.IDVino)
    '                    Dim StCrear As New BdgWorkClass.StCrearVino(rwD(_V.IDDeposito), rwD(_V.IDArticulo), rwD(_V.Lote), Fecha, BdgOrigenVino.Interno, rwD(_V.IDUdMedida), rwD(_OV.Cantidad), IDEstadoVino, NOperacion, oVinos, , rwD(_OV.IDBarrica) & String.Empty)
    '                    rwD(_OV.IDVino) = ProcessServer.ExecuteTask(Of BdgWorkClass.StCrearVino, Guid)(AddressOf BdgWorkClass.CrearVino, StCrear, services)
    '                    Dim Tables() As DataTable = New DataTable() {data.Analisis, data.Materiales, data.Centros, data.Operarios, data.Varios}
    '                    Dim StCambiar As New StCambiarIDVino(oldID, rwD(_OV.IDVino), Tables)
    '                    ProcessServer.ExecuteTask(Of StCambiarIDVino)(AddressOf CambiarIDVino, StCambiar, services)
    '                End If

    '                rwD(_OV.NOperacion) = NOperacion
    '                rwD(_OV.Destino) = True

    '            ElseIf rwD.RowState = DataRowState.Deleted Then
    '                rwD.RejectChanges()
    '                Dim DtDel As DataTable = rwD.Table.Clone
    '                DtDel.ImportRow(rwD)
    '                DtDel.AcceptChanges()
    '                oOV.Delete(DtDel)
    '            ElseIf rwD.RowState = DataRowState.Modified Then
    '                If Not oColFrx Is Nothing Then
    '                    Dim StRec As New StReconstruirEstructura(oColFrx(rwD(_OV.IDVino)), rwD(_OV.IDVino), rwD(_V.IDArticulo), rwD(_V.IDUdMedida), OrigenCurrents, _
    '                                                             OrigenDeleteds, oVwc, BdgOrigenVino.Interno, NOperacion)
    '                    ProcessServer.ExecuteTask(Of StReconstruirEstructura)(AddressOf ReconstruirEstructura, StRec, services)
    '                End If

    '                Dim dblQN As Double = rwD(_OV.Cantidad)
    '                Dim dblQ As Double = rwD(_OV.Cantidad, DataRowVersion.Original)
    '                If dblQN - dblQ <> 0 Then
    '                    Dim StInc As New BdgWorkClass.StIncrementarIDVino(rwD(_OV.IDVino), dblQN - dblQ)
    '                    ProcessServer.ExecuteTask(Of BdgWorkClass.StIncrementarIDVino)(AddressOf BdgWorkClass.IncrementarCantidadIDVino, StInc, services)
    '                Else
    '                    '//La linea puede venir como modificada aunque no se haya modificado voluntariamente
    '                End If
    '                'TODO CUIDADO ¿se debe modificar siempre el IDBarrica?
    '                '//diferentes lineas de OperacionVino pueden tener como destino el mismo Vino, y eb cada linea se puede asignar un IDBarrica diferente
    '                Dim IDBArricaOld As String = rwD(_OV.IDBarrica, DataRowVersion.Original) & String.Empty
    '                If Len(IDBArricaOld) > 0 AndAlso IDBArricaOld <> rwD(_OV.IDBarrica) & String.Empty Then
    '                    Dim drV As DataRow = v.GetItemRow(rwD(_OV.IDVino))
    '                    drV(_V.IDBarrica) = rwD(_OV.IDBarrica)
    '                    v.Update(drV.Table)
    '                End If
    '            ElseIf rwD.RowState = DataRowState.Unchanged Then
    '                If blnOrgChngd Then
    '                    If Not oColFrx Is Nothing Then
    '                        Dim StRec As New StReconstruirEstructura(oColFrx(rwD(_OV.IDVino)), rwD(_OV.IDVino), rwD(_V.IDArticulo), rwD(_V.IDUdMedida), OrigenCurrents, _
    '                                                                 OrigenDeleteds, oVwc, BdgOrigenVino.Interno, NOperacion)
    '                        ProcessServer.ExecuteTask(Of StReconstruirEstructura)(AddressOf ReconstruirEstructura, StRec, services)
    '                    End If
    '                End If
    '            End If
    '        Next
    '    End If

    '    If Not OrigenDeleteds Is Nothing Then
    '        OrigenDeleteds.RejectChanges()
    '        oOV.Delete(OrigenDeleteds)
    '    End If

    '    If Not OrigenCurrents Is Nothing Then oOV.Update(OrigenCurrents)

    '    If Not data.Destino Is Nothing Then
    '        Dim dtUpdtDst As DataTable = data.Destino.GetChanges(DataRowState.Added Or DataRowState.Modified)
    '        If Not dtUpdtDst Is Nothing Then
    '            dtUpdtDst.TableName = "BdgOperacionVino"
    '            oOV.Update(dtUpdtDst)
    '        End If
    '    End If

    '    Dim f As New Filter(FilterUnionOperator.Or)

    '    If Not data.Materiales Is Nothing Then
    '        Dim oVMat As New BdgVinoMaterial
    '        Dim oVinoLote As New BdgVinoMaterialLote
    '        Dim oMatLote As New BdgOperacionMaterialLote
    '        Dim dtDel As DataTable = data.Materiales.GetChanges(DataRowState.Deleted)
    '        If Not dtDel Is Nothing AndAlso dtDel.Rows.Count > 0 Then
    '            f.Clear()
    '            For Each drDel As DataRow In dtDel.Rows
    '                f.Add(New GuidFilterItem("IDVinoMaterial", drDel("IDVinoMaterial", DataRowVersion.Original)))
    '            Next
    '            dtDel = oVMat.Filter(f)
    '            oVMat.Delete(dtDel)
    '        End If
    '        data.Materiales = data.Materiales.GetChanges(DataRowState.Added Or DataRowState.Modified)

    '        If Not data.Materiales Is Nothing Then
    '            Dim dttVinoLote As DataTable = oVinoLote.AddNew() 'obtenemos la estructura
    '            For Each oRw As DataRow In data.Materiales.Rows
    '                If oRw.RowState = DataRowState.Added Then
    '                    If Length(oRw(_VM.IDVinoMaterial)) = 0 Then
    '                        oRw(_VM.IDVinoMaterial) = Guid.NewGuid
    '                    End If
    '                    oRw(_VM.NOperacion) = NOperacion
    '                    oRw(_VM.Fecha) = Fecha
    '                    ''si tiene IDOperacionMaterial hay que buscar los lotes asociados si los hay, para guardarlos tmb
    '                    'If (Length(oRw("IDOperacionMaterial")) > 0) Then
    '                    '    Dim dttMatLote As DataTable = oMatLote.Filter(New StringFilterItem("IDOperacionMaterial", oRw("IDOperacionMaterial")))
    '                    '    For Each dtrMatLote As DataRow In dttMatLote.Rows
    '                    '        Dim dtrVinoLote As DataRow = dttVinoLote.NewRow
    '                    '        dtrVinoLote(_BdgVinoMaterialLote.IDVinoMaterialLote) = Guid.NewGuid
    '                    '        dtrVinoLote(_BdgVinoMaterialLote.IDVinoMaterial) = oRw(_VM.IDVinoMaterial)
    '                    '        dtrVinoLote(_BdgVinoMaterialLote.Cantidad) = dtrMatLote("Cantidad")
    '                    '        dtrVinoLote(_BdgVinoMaterialLote.Lote) = dtrMatLote("Lote")
    '                    '        dtrVinoLote(_BdgVinoMaterialLote.Ubicacion) = dtrMatLote("Ubicacion")

    '                    '        dttVinoLote.Rows.Add(dtrVinoLote)
    '                    '    Next
    '                    'End If
    '                End If
    '            Next
    '            data.Materiales.TableName = oVMat.Entity
    '            oVMat.Update(data.Materiales)

    '            '
    '        End If
    '    End If

    '    'TODO - GUARDAR LOTES/GENERAR MOVIMIENTOS
    '    If Not data.DtMaterialLoteGlobal Is Nothing Then
    '        Dim oMatLote As New BdgOperacionMaterialLote
    '        oMatLote.Update(data.DtMaterialLoteGlobal)
    '    End If

    '    If Not data.DtMaterialLote Is Nothing Then
    '        Dim oMatLote As New BdgVinoMaterialLote
    '        oMatLote.Update(data.DtMaterialLote)
    '        '
    '        ProcessServer.ExecuteTask(Of DataTable)(AddressOf BdgVinoMaterial.DescuentoMateriales, data.Materiales, services)

    '    End If


    '    If Not data.Operarios Is Nothing Then
    '        Dim oVMod As BdgVinoMod = New BdgVinoMod
    '        Dim dtDel As DataTable = data.Operarios.GetChanges(DataRowState.Deleted)
    '        If Not dtDel Is Nothing AndAlso dtDel.Rows.Count > 0 Then
    '            f.Clear()
    '            For Each drDel As DataRow In dtDel.Rows
    '                f.Add(New GuidFilterItem("IdVinoMOD", drDel("IdVinoMOD", DataRowVersion.Original)))
    '            Next
    '            dtDel = oVMod.Filter(f)
    '            oVMod.Delete(dtDel)
    '        End If
    '        data.Operarios = data.Operarios.GetChanges(DataRowState.Added Or DataRowState.Modified)
    '        If Not data.Operarios Is Nothing Then
    '            For Each oRw As DataRow In data.Operarios.Rows
    '                If oRw.RowState = DataRowState.Added Then
    '                    oRw(_VMD.NOperacion) = NOperacion
    '                    oRw(_VMD.Fecha) = Fecha
    '                End If
    '            Next
    '            data.Operarios.TableName = oVMod.Entity
    '            oVMod.Update(data.Operarios)
    '        End If
    '    End If

    '    If Not data.Centros Is Nothing Then
    '        Dim oVCtr As BdgVinoCentro = New BdgVinoCentro
    '        Dim dtDel As DataTable = data.Centros.GetChanges(DataRowState.Deleted)
    '        If Not dtDel Is Nothing AndAlso dtDel.Rows.Count > 0 Then
    '            f.Clear()
    '            For Each drDel As DataRow In dtDel.Rows
    '                f.Add(New GuidFilterItem("IdVinoCentro", drDel("IdVinoCentro", DataRowVersion.Original)))
    '            Next
    '            dtDel = oVCtr.Filter(f)
    '            oVCtr.Delete(dtDel)
    '        End If
    '        data.Centros = data.Centros.GetChanges(DataRowState.Added Or DataRowState.Modified)

    '        If Not data.Centros Is Nothing Then
    '            For Each oRw As DataRow In data.Centros.Rows
    '                If oRw.RowState = DataRowState.Added Then
    '                    oRw(_VC.NOperacion) = NOperacion
    '                    oRw(_VC.Fecha) = Fecha
    '                End If
    '            Next
    '            data.Centros.TableName = oVCtr.Entity
    '            oVCtr.Update(data.Centros)
    '        End If
    '    End If

    '    If Not data.Analisis Is Nothing Then
    '        Dim IDVino As String = String.Empty

    '        For Each drAnalisis As DataRow In data.Analisis.Select("", "IDVino")
    '            If drAnalisis("IDVino").ToString <> IDVino Then
    '                Dim f_A As New Filter
    '                f_A.Add(New GuidFilterItem("IDVino", drAnalisis("IDVino")))
    '                f_A.Add(New StringFilterItem("NOperacion", rwOp("NOperacion")))
    '                Dim dtVA As DataTable = New BdgVinoAnalisis().Filter(f_A)
    '                If Not IsNothing(dtVA) AndAlso dtVA.Rows.Count > 0 Then
    '                    Dim dtVV As DataTable = New BdgVinoVariable().Filter(New GuidFilterItem("IDVinoAnalisis", dtVA.Rows(0)("IDVinoAnalisis")))
    '                    If Not IsNothing(dtVV) AndAlso dtVV.Rows.Count > 0 Then
    '                        For Each drVV As DataRow In dtVV.Rows
    '                            f_A.Clear()
    '                            f_A.Add(New GuidFilterItem("IDVino", drAnalisis("IDVino")))
    '                            f_A.Add(New StringFilterItem("IDVariable", drVV("IDVariable")))
    '                            Dim where As String = f_A.Compose(New AdoFilterComposer)
    '                            Dim dDatos() As DataRow = data.Analisis.Select(where)
    '                            If dDatos.Length > 0 Then
    '                                drVV("IDVariable") = dDatos(0)("IDVariable")
    '                                drVV("Valor") = dDatos(0)("Valor")
    '                                drVV("ValorNumerico") = dDatos(0)("ValorNumerico")
    '                            End If
    '                        Next
    '                        BdgVinoVariable.UpdateTable(dtVV)
    '                    End If
    '                End If
    '            End If
    '            IDVino = drAnalisis("IDVino").ToString
    '        Next
    '    End If

    '    If Not data.Varios Is Nothing Then
    '        Dim oVV As BdgVinoVarios = New BdgVinoVarios
    '        Dim dtDel As DataTable = data.Varios.GetChanges(DataRowState.Deleted)
    '        If Not dtDel Is Nothing AndAlso dtDel.Rows.Count > 0 Then
    '            f.Clear()
    '            For Each drDel As DataRow In dtDel.Rows
    '                f.Add(New GuidFilterItem("IdVinoVarios", drDel("IdVinoVarios", DataRowVersion.Original)))
    '            Next
    '            dtDel = oVV.Filter(f)
    '            oVV.Delete(dtDel)
    '        End If
    '        data.Varios = data.Varios.GetChanges(DataRowState.Added Or DataRowState.Modified)
    '        If Not data.Varios Is Nothing Then
    '            For Each oRw As DataRow In data.Varios.Rows
    '                If oRw.RowState = DataRowState.Added Then
    '                    oRw(_VMD.NOperacion) = NOperacion
    '                    oRw(_VMD.Fecha) = Fecha
    '                End If
    '            Next
    '            data.Varios.TableName = oVV.Entity
    '            oVV.Update(data.Varios)
    '        End If
    '    End If

    '    '//Actualización de Ordenes de fabricación
    '    Dim StActuaOF As New Business.Bodega.BdgOperacion.StActualizarOF(data.Destino, Fecha)
    '    ProcessServer.ExecuteTask(Of Business.Bodega.BdgOperacion.StActualizarOF)(AddressOf Business.Bodega.BdgOperacion.ActualizarOF, StActuaOF, services)

    '    '//Gestion de stock
    '    If rwOp.IsNull("IDMovimiento") Then
    '        Dim StEjec As New BdgWorkClass.StEjecutarMovimientos(rwOp("NOperacion"), FechaDoc)
    '        rwOp("IDMovimiento") = ProcessServer.ExecuteTask(Of BdgWorkClass.StEjecutarMovimientos, Integer)(AddressOf BdgWorkClass.EjecutarMovimientos, StEjec, services)
    '        BusinessHelper.UpdateTable(rwOp.Table)
    '    Else
    '        Dim StEj As New BdgWorkClass.StEjecutarMovimientosNumero(rwOp("IDMovimiento"), rwOp("NOperacion"), FechaDoc)
    '        ProcessServer.ExecuteTask(Of BdgWorkClass.StEjecutarMovimientosNumero)(AddressOf BdgWorkClass.EjecutarMovimientosNumero, StEj, services)
    '    End If
    'End Sub

    '<Serializable()> _
    'Public Class StValidarFechaOperacion
    '    Public FechaVino As Date
    '    Public FechaOperacion As Date

    '    Public Sub New()
    '    End Sub

    '    Public Sub New(ByVal FechaVino As Date, ByVal FechaOperacion As Date)
    '        Me.FechaVino = FechaVino
    '        Me.FechaOperacion = FechaOperacion
    '    End Sub
    'End Class

    '<Task()> Public Shared Function ValidarFechaOperacion(ByVal data As StValidarFechaOperacion, ByVal services As ServiceProvider) As Boolean
    '    'Comprobar si la Fecha del Vino es posterior a la Fecha de la Operación.

    '    If New BdgParametro().BdgComprobarFechaOperacion() Then
    '        If data.FechaVino > data.FechaOperacion Then
    '            ApplicationService.GenerateError("La Fecha de la Operación {0} es anterior a la Fecha del Vino Origen {1}.", Format(data.FechaOperacion, "dd/MM/yyyy"), Format(data.FechaVino, "dd/MM/yyyy"))
    '            Return False
    '        Else : Return True
    '        End If
    '    Else : Return True
    '    End If
    'End Function

    '<Serializable()> _
    'Public Class StModificarFechaOperacion
    '    Public NOperacion As String
    '    Public Fecha As Date
    '    Public IDMovimiento As Integer
    '    Public DtTO As DataTable

    '    Public Sub New()
    '    End Sub

    '    Public Sub New(ByVal NOperacion As String, ByVal Fecha As Date, ByVal IDMovimiento As Integer, ByVal DtTO As DataTable)
    '        Me.NOperacion = NOperacion
    '        Me.Fecha = Fecha
    '        Me.IDMovimiento = IDMovimiento
    '        Me.DtTO = DtTO
    '    End Sub
    'End Class

    '<Task()> Public Shared Sub ModificarFechaOperacion(ByVal data As StModificarFechaOperacion, ByVal services As ServiceProvider)
    '    Dim dtOpHis As DataTable = AdminData.GetData("tbHistoricoMovimiento", New NumberFilterItem("IDMovimiento", data.IDMovimiento))
    '    For Each rwHM As DataRow In dtOpHis.Rows
    '        If Length(rwHM("IDLineaMovimiento")) > 0 Then
    '            Dim dataCorreccion As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Corregir, rwHM("IDLineaMovimiento"), data.Fecha, False)
    '            ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento)(AddressOf ProcesoStocks.ActualizarMovimiento, dataCorreccion, services)
    '        End If
    '    Next

    '    If data.DtTO.Rows(0)(_TO.TipoMovimiento) <> Business.Bodega.TipoMovimiento.SinMovimiento _
    '        And data.DtTO.Rows(0)(_TO.TipoMovimiento) <> Business.Bodega.TipoMovimiento.CrearOrigen Then

    '        Dim oOVaux As New BdgOperacionVino
    '        Dim oVaux As New BdgVino
    '        Dim fltr As New Filter
    '        fltr.Add(New StringFilterItem(_OV.NOperacion, data.NOperacion))
    '        fltr.Add(New BooleanFilterItem(_OV.Destino, True))
    '        Dim dtOVaux As DataTable = oOVaux.Filter(fltr)
    '        For Each rwOVaux As DataRow In dtOVaux.Rows
    '            Dim rwVaux As DataRow = oVaux.GetItemRow(rwOVaux(_OV.IDVino))
    '            rwVaux(_V.Fecha) = data.Fecha
    '            rwVaux(_V.DiasDeposito) = 0
    '            rwVaux(_V.DiasBarrica) = 0
    '            rwVaux(_V.DiasBotellero) = 0
    '            BusinessHelper.UpdateTable(rwVaux.Table)
    '            ProcessServer.ExecuteTask(Of Guid)(AddressOf BdgWorkClass.ActualizarDiasVino, rwOVaux(_OV.IDVino), services)
    '        Next
    '    End If
    'End Sub

    '<Serializable()> _
    'Public Class StReconstruirEstructura
    '    Public Factor As Double
    '    Public IDVIno As Guid
    '    Public IDArticulo As String
    '    Public IDUDMedida As String
    '    Public CurOrigen As DataTable
    '    Public DeslOrigen As DataTable
    '    Public oVWC As BdgWorkClass
    '    Public Origen As BdgOrigenVino
    '    Public NOperacion As String

    '    Public Sub New()
    '    End Sub

    '    Public Sub New(ByVal Factor As Double, ByVal IDVino As Guid, ByVal IDArticulo As String, ByVal IDUDMedida As String, _
    '                   ByVal CurOrigen As DataTable, ByVal DeslOrigen As DataTable, ByVal oVWC As BdgWorkClass, ByVal Origen As BdgOrigenVino, _
    '                   ByVal NOperacion As String)
    '        Me.Factor = Factor
    '        Me.IDVIno = IDVino
    '        Me.IDArticulo = IDArticulo
    '        Me.IDUDMedida = IDUDMedida
    '        Me.CurOrigen = CurOrigen
    '        Me.DeslOrigen = DeslOrigen
    '        Me.oVWC = oVWC
    '        Me.Origen = Origen
    '        Me.NOperacion = NOperacion
    '    End Sub
    'End Class

    '<Task()> Public Shared Sub ReconstruirEstructura(ByVal data As StReconstruirEstructura, ByVal services As ServiceProvider)
    '    Dim StCalc As New StCalcularEstructura(data.Factor, data.IDArticulo, data.IDUDMedida, data.CurOrigen)
    '    Dim oVinos() As VinoComponente = ProcessServer.ExecuteTask(Of StCalcularEstructura, VinoComponente())(AddressOf CalcularEstructura, StCalc, services)

    '    Dim Dels() As Guid
    '    If Not data.DeslOrigen Is Nothing Then
    '        ReDim Dels(data.DeslOrigen.Rows.Count - 1)
    '        For i As Integer = 0 To data.DeslOrigen.Rows.Count - 1
    '            Dels(i) = data.DeslOrigen.Rows(i)(_OV.IDVino, DataRowVersion.Original)
    '        Next
    '    End If
    '    Dim StCrear As New BdgWorkClass.StCrearEstructuraOperacion(data.IDVIno, data.Origen, data.NOperacion, oVinos, Dels)
    '    ProcessServer.ExecuteTask(Of BdgWorkClass.StCrearEstructuraOperacion)(AddressOf BdgWorkClass.CrearEstructuraOperacion, StCrear, services)
    'End Sub

#End Region

    <Serializable()> _
    Public Class StAddToDtOpVino
        Public DtOV As DataTable
        Public NOperacion As String
        Public Dt As DataTable
        Public Destino As Boolean

        Public Sub New()
        End Sub

        Public Sub New(ByVal DtOV As DataTable, ByVal NOperacion As String, ByVal Dt As DataTable, ByVal Destino As Boolean)
            Me.DtOV = DtOV
            Me.NOperacion = NOperacion
            Me.Dt = Dt
            Me.Destino = Destino
        End Sub
    End Class

    <Task()> Public Shared Sub AddToDTOpVino(ByVal data As StAddToDtOpVino, ByVal services As ServiceProvider)
        For Each oRw As DataRow In data.Dt.Rows
            Dim rwOV As DataRow = data.DtOV.NewRow
            rwOV(_OV.NOperacion) = data.NOperacion
            rwOV(_OV.IDVino) = oRw(_OV.IDVino)
            rwOV(_OV.Ocupacion) = oRw(_OV.Ocupacion)
            rwOV(_OV.Cantidad) = oRw(_OV.Cantidad)
            rwOV(_OV.QDeposito) = oRw(_OV.QDeposito)
            rwOV(_OV.Merma) = oRw(_OV.Merma)
            'rwOV(_OV.Lote) = oRw(_OV.Lote)
            rwOV(_OV.Litros) = oRw(_OV.Litros)
            rwOV(_OV.Destino) = data.Destino
            If data.Destino Then
                rwOV(_OV.IDOrden) = oRw(_OV.IDOrden)
                rwOV(_OV.IDBarrica) = oRw(_OV.IDBarrica)
                rwOV(_OV.IDEstadoVino) = oRw(_OV.IDEstadoVino)
            Else
                rwOV("IDTipoMermaVino") = oRw("IDTipoMermaVino")
            End If
            data.DtOV.Rows.Add(rwOV)
        Next
    End Sub

#End Region

#Region "Mano de Obra asociados a la operacion"

    '<Serializable()> _
    'Public Class StModTipoOperacionGlobal
    '    Public IDTipoOperacion As String
    '    Public MODS As DataTable
    '    Public Fecha As Date
    '    Public NOperacion As String
    '    Public Imputar As Boolean
    '    Public DtModGlobal As DataTable

    '    Public Sub New()
    '    End Sub

    '    Public Sub New(ByVal IDTipoOperacion As String, ByVal MODS As DataTable, Optional ByVal Fecha As Date = cnMinDate, _
    '                   Optional ByVal NOperacion As String = "", Optional ByVal Imputar As Boolean = False, _
    '                   Optional ByVal DtModGlobal As DataTable = Nothing)
    '        Me.IDTipoOperacion = IDTipoOperacion
    '        Me.MODS = MODS
    '        Me.Fecha = Fecha
    '        Me.NOperacion = NOperacion
    '        Me.Imputar = Imputar
    '        Me.DtModGlobal = DtModGlobal
    '    End Sub
    'End Class

    ''Si es nueva la Operacion
    '<Task()> Public Shared Sub ModTipoOperacionGlobal(ByVal data As StModTipoOperacionGlobal, ByVal services As ServiceProvider)
    '    If Len(data.IDTipoOperacion) > 0 Then
    '        Dim componentes As DataTable
    '        If data.Imputar Then
    '            componentes = data.DtModGlobal.Copy
    '        Else
    '            componentes = AdminData.GetData("frmBdgTipoOperacionMod", New StringFilterItem("IDTipoOperacion", data.IDTipoOperacion))
    '        End If
    '        For Each componente As DataRow In componentes.Rows
    '            Dim drmanoObra As DataRow = data.MODS.NewRow
    '            data.MODS.Rows.Add(drmanoObra)
    '            drmanoObra("NOperacion") = data.NOperacion
    '            drmanoObra("IDOperario") = componente("IDOperario")
    '            drmanoObra("DescOperario") = componente("DescOperario")
    '            drmanoObra("Tiempo") = componente("Tiempo")
    '            Dim mstrHoraPred As String = New Parametro().HoraPred() 'Recuperar la Hora por defecto

    '            Dim dataHora As New General.HoraCategoria.DatosPrecioHoraCatOper(componente("IDCategoria"), mstrHoraPred, data.Fecha, componente("IDOperario"))
    '            Dim dtOperarioHora As DataTable = ProcessServer.ExecuteTask(Of General.HoraCategoria.DatosPrecioHoraCatOper, DataTable)(AddressOf General.HoraCategoria.ObtenerPrecioHoraCategoriaOperario, dataHora, services)
    '            'Como los datos pueden venir de TbOperarioHora o TbMaestroHoraCategoria, y las columnas aunque muestren el 
    '            'mismo dato, no se llaman igual se utiliza la funcion de abajo para saber de que tabla provienen los datos
    '            If Not dtOperarioHora Is Nothing AndAlso dtOperarioHora.Rows.Count > 0 Then
    '                If dtOperarioHora.TableName = "OperarioHora" Then
    '                    drmanoObra("Tasa") = dtOperarioHora.Rows(0)("TasaHorariaA")
    '                Else
    '                    drmanoObra("Tasa") = dtOperarioHora.Rows(0)("PrecioHoraA")
    '                End If
    '                drmanoObra("IDHora") = dtOperarioHora.Rows(0)("IDHora")
    '            Else
    '                ApplicationService.GenerateError("Debe estar Relleno la hora por categoria")
    '            End If

    '            'drmanoObra("Importe") = Nz(componente("Tiempo"), 0) * Nz(drmanoObra("Tasa"), 0)
    '            drmanoObra("Fecha") = data.Fecha
    '            drmanoObra("IDCategoria") = componente("IDCategoria")

    '            'Asignar los ID´s
    '            If data.MODS.Columns.Contains("IDOperacionMod") Then
    '                If data.MODS.TableName = "BdgOperacionMod" AndAlso Length(drmanoObra("IDOperacionMod")) = 0 Then
    '                    drmanoObra("IDOperacionMod") = Guid.NewGuid()
    '                ElseIf componentes.Columns.Contains("IDOperacionMod") Then
    '                    drmanoObra("IDOperacionMod") = componente("IDOperacionMod")
    '                End If
    '            End If
    '        Next
    '    End If
    'End Sub

    '<Serializable()> _
    'Public Class StModTipoOperacion
    '    Public IDTipoOperacion As String
    '    Public Imputar As Boolean
    '    Public Fecha As Date

    '    Public Sub New()
    '    End Sub

    '    Public Sub New(ByVal IDTipoOperacion As String, ByVal Imputar As Boolean, Optional ByVal Fecha As Date = cnMinDate)
    '        Me.IDTipoOperacion = IDTipoOperacion
    '        Me.Imputar = Imputar
    '        Me.Fecha = Fecha
    '    End Sub
    'End Class

    ''Si no estan chequeados los del Tipo de Operaciones
    '<Task()> Public Shared Function ModTipoOperacion(ByVal data As StModTipoOperacion, ByVal services As ServiceProvider) As DataTable
    '    If Len(data.IDTipoOperacion) > 0 AndAlso Not data.Imputar Then
    '        Dim dtMod As DataTable = AdminData.GetData("frmBdgOperacionModPrev", New NoRowsFilterItem)
    '        Dim dtImputacionGlobal As DataTable = AdminData.GetData("frmBdgTipoOperacionMod", New StringFilterItem("IDTipoOperacion", data.IDTipoOperacion))

    '        For Each drGlobal As DataRow In dtImputacionGlobal.Rows
    '            Dim drFilaMOD As DataRow = dtMod.NewRow

    '            drFilaMOD("IDOperario") = drGlobal("IDOperario")
    '            drFilaMOD("DescOperario") = drGlobal("DescOperario")
    '            drFilaMOD("Tiempo") = drGlobal("Tiempo")
    '            drFilaMOD("IDCategoria") = drGlobal("IDCategoria")

    '            Dim mstrHoraPred As String = New Parametro().HoraPred() 'Recuperar la Hora por defecto
    '            Dim dataHora As New General.HoraCategoria.DatosPrecioHoraCatOper(drGlobal("IDCategoria"), mstrHoraPred, data.Fecha, drGlobal("IDOperario"))
    '            Dim dtOperarioHora As DataTable = ProcessServer.ExecuteTask(Of General.HoraCategoria.DatosPrecioHoraCatOper, DataTable)(AddressOf General.HoraCategoria.ObtenerPrecioHoraCategoriaOperario, dataHora, services)
    '            'Como los datos pueden venir de TbOperarioHora o TbMaestroHoraCategoria, y las columnas aunque muestren el 
    '            'mismo dato, no se llaman igual se utiliza la funcion de abajo para saber de que tabla provienen los datos
    '            If Not dtOperarioHora Is Nothing AndAlso dtOperarioHora.Rows.Count > 0 Then
    '                If dtOperarioHora.TableName = "OperarioHora" Then
    '                    drFilaMOD("Tasa") = dtOperarioHora.Rows(0)("TasaHorariaA")
    '                Else
    '                    drFilaMOD("Tasa") = dtOperarioHora.Rows(0)("PrecioHoraA")
    '                End If
    '                drFilaMOD("IDHora") = dtOperarioHora.Rows(0)("IDHora")
    '            Else
    '                ApplicationService.GenerateError("Debe estar Relleno la hora por categoria")
    '            End If
    '            'drFilaMOD("Importe") = Nz(drGlobal("Tiempo"), 0) * Nz(drFilaMOD("Tasa"), 0)
    '            dtMod.Rows.Add(drFilaMOD)
    '        Next
    '        Return dtMod
    '    End If
    'End Function

    '<Serializable()> _
    'Public Class StMODS
    '    Public TipoOperacion As String
    '    Public Fecha As Date
    '    Public DtModGlobal As DataTable
    '    Public ImputacionGlobalReal As Boolean
    '    Public ImputacionGlobalPrev As Boolean
    '    Public Imputar As Boolean

    '    Public Sub New()
    '    End Sub

    '    Public Sub New(ByVal TipoOperacion As String, Optional ByVal Fecha As Date = cnMinDate, Optional ByVal DtModGlobal As DataTable = Nothing, _
    '                   Optional ByVal ImputacionGlobalReal As Boolean = False, Optional ByVal ImputacionGlobalPrev As Boolean = False, _
    '                   Optional ByVal Imputar As Boolean = False)
    '        Me.TipoOperacion = TipoOperacion
    '        Me.Fecha = Fecha
    '        Me.DtModGlobal = DtModGlobal
    '        Me.ImputacionGlobalReal = ImputacionGlobalReal
    '        Me.ImputacionGlobalPrev = ImputacionGlobalPrev
    '        Me.Imputar = Imputar
    '    End Sub
    'End Class

    '<Task()> Public Shared Function MODS(ByVal data As StMODS, ByVal services As ServiceProvider) As DataTable
    '    Dim dtMods As DataTable
    '    If data.ImputacionGlobalPrev Then
    '        Dim StModTipo As New StModTipoOperacion(data.TipoOperacion, data.Imputar, data.Fecha)
    '        dtMods = ProcessServer.ExecuteTask(Of StModTipoOperacion, DataTable)(AddressOf ModTipoOperacion, StModTipo, services)
    '    ElseIf data.ImputacionGlobalReal Then
    '        dtMods = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf GetDTMods, New Object, services)
    '        Dim StModTipo As New StModTipoOperacionGlobal(data.TipoOperacion, dtMods, data.Fecha, , data.Imputar, data.DtModGlobal)
    '        ProcessServer.ExecuteTask(Of StModTipoOperacionGlobal)(AddressOf ModTipoOperacionGlobal, StModTipo, services)
    '    End If
    '    Return dtMods
    'End Function

    '<Task()> Public Shared Function GetDTMods(ByVal data As Object, ByVal services As ServiceProvider) As DataTable
    '    Dim oVC As New BdgVinoMod
    '    Dim mods As DataTable = oVC.AddNew()
    '    mods.Columns.Add("DescOperario", GetType(String))
    '    Return mods
    'End Function

#End Region

#Region "Centro asociados a la operacion"

    '<Serializable()> _
    'Public Class StCentroTipoOperacionGlobal
    '    Public IDTipoOperacion As String
    '    Public Centros As DataTable
    '    Public Cantidad As Double
    '    Public Fecha As Date
    '    Public NOperacion As String
    '    Public Imputar As Boolean
    '    Public DtCentroGlobal As DataTable

    '    Public Sub New()
    '    End Sub

    '    Public Sub New(ByVal IDTipoOperacion As String, ByVal Centros As DataTable, Optional ByVal Cantidad As Double = 1, Optional ByVal Fecha As Date = cnMinDate, _
    '                   Optional ByVal NOperacion As String = "", Optional ByVal Imputar As Boolean = False, Optional ByVal DtCentroGlobal As DataTable = Nothing)
    '        Me.IDTipoOperacion = IDTipoOperacion
    '        Me.Centros = Centros
    '        Me.Cantidad = Cantidad
    '        Me.Fecha = Fecha
    '        Me.NOperacion = NOperacion
    '        Me.Imputar = Imputar
    '        Me.DtCentroGlobal = DtCentroGlobal
    '    End Sub
    'End Class

    ''Si es nueva la Operacion Global
    '<Task()> Public Shared Sub CentroTipoOperacionGlobal(ByVal data As StCentroTipoOperacionGlobal, ByVal services As ServiceProvider)
    '    If Len(data.IDTipoOperacion) > 0 Then
    '        Dim componentes As DataTable
    '        If data.Imputar Then
    '            componentes = data.DtCentroGlobal.Copy
    '        Else
    '            componentes = AdminData.GetData("frmBdgTipoOperacionCentro", New StringFilterItem("IDTipoOperacion", data.IDTipoOperacion))
    '        End If
    '        For Each componente As DataRow In componentes.Rows
    '            Dim drFilaCentro As DataRow = data.Centros.NewRow
    '            data.Centros.Rows.Add(drFilaCentro)
    '            drFilaCentro("NOperacion") = data.NOperacion
    '            drFilaCentro("IDCentro") = componente("IDCentro")
    '            'drFilaCentro("DescCentro") = componente("DescCentro")
    '            drFilaCentro("Tiempo") = componente("Tiempo")


    '            Dim oDe As DataEngine = New DataEngine
    '            Dim dtCentro As DataTable = oDe.RetrieveData("frmMntoCentro", , New FilterItem("IDCentro", FilterOperator.Equal, componente("IDCentro")))
    '            If dtCentro.Rows.Count = 1 Then
    '                Dim strIdUdMedidaCentro As String = dtCentro.Rows(0)("IDUdMedida") & String.Empty
    '                'Dim strIdUdMedidaArticulo As String = dtCentroGlobal.Rows(0)("IDUdMedida") & String.Empty
    '                'Dim dblFactor As Double = New UnidadAB().FactorDeConversion(strIdUdMedidaArticulo, strIdUdMedidaCentro)
    '                'Dim dblQ As Double = componente("Cantidad") '* dblFactor 'GetCurrentQ(GridDestino) * dblFactor
    '                drFilaCentro("Tasa") = dtCentro.Rows(0)("TasaEjecucionA")
    '                'drFilaCentro("Cantidad") = 0
    '                drFilaCentro("IdUdMedidaCentro") = strIdUdMedidaCentro
    '                drFilaCentro("IdUdMedidaArticulo") = strIdUdMedidaCentro
    '                Dim UDTiempo As Integer
    '                'Dim IDVino As Guid = GetCurrentVino(grdDestino)
    '                If dtCentro.Rows(0)("UDTiempo") Is System.DBNull.Value Then
    '                    UDTiempo = enumstdUdTiempo.Horas 'Por defecto
    '                Else
    '                    UDTiempo = dtCentro.Rows(0)("UDTiempo")
    '                End If
    '                drFilaCentro("UDTiempo") = UDTiempo
    '                drFilaCentro("PorCantidad") = componente("PorCantidad")
    '            End If

    '            'drFilaCentro("Importe") = Nz(componente("Tiempo"), 0) * Nz(drFilaCentro("Tasa"), 0)
    '            drFilaCentro("Fecha") = data.Fecha


    '            'Asignar los ID´s
    '            If data.Centros.Columns.Contains("IDOperacionCentro") Then
    '                If data.Centros.TableName = "BdgOperacionCentro" AndAlso Length(drFilaCentro("IDOperacionCentro")) = 0 Then
    '                    drFilaCentro("IDOperacionCentro") = Guid.NewGuid()
    '                ElseIf componentes.Columns.Contains("IDOperacionCentro") Then
    '                    drFilaCentro("IDOperacionCentro") = componente("IDOperacionCentro")
    '                End If
    '            End If
    '        Next
    '    End If
    'End Sub

    '<Serializable()> _
    'Public Class StCentroTipoOperacion
    '    Public IDTipoOperacion As String
    '    Public Imputar As Boolean

    '    Public Sub New()
    '    End Sub

    '    Public Sub New(ByVal IDTipoOperacion As String, ByVal Imputar As Boolean)
    '        Me.IDTipoOperacion = IDTipoOperacion
    '        Me.Imputar = Imputar
    '    End Sub
    'End Class

    ''Si no estan chequeados los del Tipo de Operaciones
    '<Task()> Public Shared Function CentroTipoOperacion(ByVal data As StCentroTipoOperacion, ByVal services As ServiceProvider) As DataTable
    '    If Len(data.IDTipoOperacion) > 0 AndAlso Not data.Imputar Then
    '        Dim oArt As New Articulo
    '        Dim dtCentro As DataTable = AdminData.GetData("frmBdgOperacionCentroPrev", New NoRowsFilterItem)
    '        Dim dtImputacionGlobal As DataTable = AdminData.GetData("frmBdgTipoOperacionCentro", New StringFilterItem("IDTipoOperacion", data.IDTipoOperacion))

    '        For Each drGlobal As DataRow In dtImputacionGlobal.Rows
    '            Dim drFilaCentro As DataRow = dtCentro.NewRow
    '            Dim strIdUdMedidaCentro As String '= strIdUdMedidaArticulo
    '            If Not drGlobal.IsNull("IDUdMedida") Then strIdUdMedidaCentro = drGlobal("IDUdMedida")
    '            drFilaCentro("UdTiempo") = drGlobal("UdTiempo")
    '            drFilaCentro("Tiempo") = drGlobal("Tiempo")
    '            drFilaCentro("IDCentro") = drGlobal("IDCentro")
    '            drFilaCentro("DescCentro") = drGlobal("DescCentro")
    '            drFilaCentro("Tasa") = drGlobal("TasaEjecucionA")
    '            drFilaCentro("PorCantidad") = drGlobal("PorCantidad")
    '            drFilaCentro("IdUdMedidaCentro") = strIdUdMedidaCentro
    '            dtCentro.Rows.Add(drFilaCentro)
    '        Next
    '        Return dtCentro
    '    End If
    'End Function

    '<Task()> Public Shared Function GetDTCentros(ByVal data As Object, ByVal services As ServiceProvider) As DataTable
    '    Dim oVC As New BdgVinoCentro
    '    Dim centros As DataTable = oVC.AddNew()
    '    centros.Columns.Add("DescCentro", GetType(String))
    '    Return centros
    'End Function

    '<Serializable()> _
    'Public Class StCentros
    '    Public TipoOperacion As String
    '    Public IDArticulo As String
    '    Public Cantidad As Double
    '    Public IDOrden As Integer
    '    Public DtCentroGlobal As DataTable
    '    Public ImputacionGlobalReal As Boolean
    '    Public ImputacionGlobaLPrev As Boolean
    '    Public Imputar As Boolean

    '    Public Sub New()
    '    End Sub

    '    Public Sub New(ByVal TipoOperacion As String, ByVal IDArticulo As String, ByVal Cantidad As Double, Optional ByVal IDOrden As Integer = 0, _
    '                   Optional ByVal DtCentroGlobal As DataTable = Nothing, Optional ByVal ImputacionGlobalReal As Boolean = False, _
    '                   Optional ByVal ImputacionGlobalPrev As Boolean = False, Optional ByVal Imputar As Boolean = False)
    '        Me.TipoOperacion = TipoOperacion
    '        Me.IDArticulo = IDArticulo
    '        Me.Cantidad = Cantidad
    '        Me.IDOrden = IDOrden
    '        Me.DtCentroGlobal = DtCentroGlobal
    '        Me.ImputacionGlobalReal = ImputacionGlobalReal
    '        Me.ImputacionGlobaLPrev = ImputacionGlobalPrev
    '        Me.Imputar = Imputar
    '    End Sub
    'End Class

    '<Task()> Public Shared Function Centros(ByVal data As StCentros, ByVal services As ServiceProvider) As DataTable
    '    Dim dtCentros As DataTable
    '    If data.ImputacionGlobaLPrev Then
    '        Dim StCentroOp As New StCentroTipoOperacion(data.TipoOperacion, data.Imputar)
    '        dtCentros = ProcessServer.ExecuteTask(Of StCentroTipoOperacion, DataTable)(AddressOf CentroTipoOperacion, StCentroOp, services)
    '    ElseIf data.ImputacionGlobalReal Then
    '        dtCentros = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf GetDTCentros, New Object, services)
    '        Dim StCentroTipo As New StCentroTipoOperacionGlobal(data.TipoOperacion, dtCentros, data.Cantidad, , , data.Imputar, data.DtCentroGlobal)
    '        ProcessServer.ExecuteTask(Of StCentroTipoOperacionGlobal)(AddressOf CentroTipoOperacionGlobal, StCentroTipo, services)
    '    End If
    '    If data.IDOrden > 0 Then
    '        '    Dim dtOFCentros As DataTable = AdminData.GetData("NegBdgOrdenCentro", New NumberFilterItem("IDOrden", IDOrden))
    '        '    For Each rwC As DataRow In dtOFCentros.Rows
    '        '        Dim nwRw As DataRow = dtCentros.NewRow
    '        '        dtCentros.Rows.Add(nwRw)
    '        '        For Each dc As DataColumn In dtCentros.Columns
    '        '            If dtOFCentros.Columns.Contains(dc.ColumnName) Then
    '        '                nwRw(dc) = rwC(dc.ColumnName)
    '        '            End If
    '        '        Next

    '        '        nwRw(_VC.PorCantidad) = True
    '        '        nwRw(_VC.Tiempo) = rwC("TiempoEjecUnit")
    '        '        nwRw(_VC.Cantidad) = Cantidad
    '        '        nwRw(_VC.UDTiempo) = rwC("UdTiempoEjec")
    '        '        nwRw("IdUdMedidaCentro") = strIdUdMedidaArticulo
    '        '        nwRw("IdUdMedidaArticulo") = strIdUdMedidaArticulo            
    '    End If
    '    Return dtCentros
    'End Function

#End Region

#Region "Varios asociados a la operacion"

    '<Serializable()> _
    'Public Class StVariosTipoOperacionGlobal
    '    Public IDTipoOperacion As String
    '    Public Varios As DataTable
    '    Public Fecha As Date
    '    Public NOperacion As String
    '    Public Imputar As Boolean
    '    Public DtVariosGlobal As DataTable

    '    Public Sub New()
    '    End Sub

    '    Public Sub New(ByVal IDTipoOperacion As String, ByVal Varios As DataTable, Optional ByVal Fecha As Date = cnMinDate, _
    '                   Optional ByVal NOperacion As String = "", Optional ByVal Imputar As Boolean = False, _
    '                   Optional ByVal DtVariosGlobal As DataTable = Nothing)
    '        Me.IDTipoOperacion = IDTipoOperacion
    '        Me.Varios = Varios
    '        Me.Fecha = Fecha
    '        Me.NOperacion = NOperacion
    '        Me.Imputar = Imputar
    '        Me.DtVariosGlobal = DtVariosGlobal
    '    End Sub
    'End Class

    ''Si es nueva la Operacion
    '<Task()> Public Shared Sub VariosTipoOperacionGlobal(ByVal data As StVariosTipoOperacionGlobal, ByVal services As ServiceProvider)
    '    If Length(data.IDTipoOperacion) > 0 Then
    '        Dim componentes As DataTable
    '        If data.Imputar Then
    '            componentes = data.DtVariosGlobal.Copy
    '        Else
    '            componentes = AdminData.GetData("frmBdgTipoOperacionVarios", New StringFilterItem("IDTipoOperacion", data.IDTipoOperacion))
    '        End If
    '        For Each componente As DataRow In componentes.Rows
    '            Dim drFilaVarios As DataRow = data.Varios.NewRow
    '            data.Varios.Rows.Add(drFilaVarios)
    '            drFilaVarios("NOperacion") = data.NOperacion
    '            drFilaVarios("IDVarios") = componente("IDVarios")
    '            drFilaVarios("DescVarios") = componente("DescVarios")
    '            drFilaVarios("Cantidad") = componente("Cantidad")
    '            drFilaVarios("TipoCosteFV") = componente("TipoCosteFV")
    '            drFilaVarios("TipoCosteDI") = componente("TipoCosteDI")
    '            drFilaVarios("Fiscal") = componente("Fiscal")
    '            drFilaVarios("Tasa") = componente("Tasa")

    '            'Asignar los ID´s
    '            If data.Varios.Columns.Contains("IDOperacionVarios") Then
    '                If data.Varios.TableName = "BdgOperacionVarios" AndAlso Length(drFilaVarios("IDOperacionVarios")) = 0 Then
    '                    drFilaVarios("IDOperacionVarios") = Guid.NewGuid()
    '                ElseIf componentes.Columns.Contains("IDOperacionVarios") Then
    '                    drFilaVarios("IDOperacionVarios") = componente("IDOperacionVarios")
    '                End If
    '            ElseIf data.Varios.Columns.Contains("IDOperacionVariosPrev") Then
    '                If data.Varios.TableName = "BdgOperacionVariosPrev" AndAlso Length(drFilaVarios("IDOperacionVariosPrev")) = 0 Then
    '                    drFilaVarios("IDOperacionVariosPrev") = Guid.NewGuid()
    '                End If
    '            End If
    '        Next
    '    End If
    'End Sub

    '<Serializable()> _
    'Public Class StVariosTipoOperacion
    '    Public IDTipoOperacion As String
    '    Public Imputar As Boolean

    '    Public Sub New()
    '    End Sub

    '    Public Sub New(ByVal IDTIpoOperacion As String, ByVal Imputar As Boolean)
    '        Me.IDTipoOperacion = IDTIpoOperacion
    '        Me.Imputar = Imputar
    '    End Sub
    'End Class

    ''Si no estan chequeados los del Tipo de Operaciones
    '<Task()> Public Shared Function VariosTipoOperacion(ByVal data As StVariosTipoOperacion, ByVal services As ServiceProvider) As DataTable
    '    If Len(data.IDTipoOperacion) > 0 AndAlso Not data.Imputar Then
    '        Dim dtVarios As DataTable = New BE.DataEngine().Filter("frmBdgOperacionVariosPrev", New NoRowsFilterItem)
    '        Dim dtImputacionGlobal As DataTable = New BE.DataEngine().Filter("frmBdgTipoOperacionVarios", New StringFilterItem("IDTipoOperacion", data.IDTipoOperacion))

    '        For Each drGlobal As DataRow In dtImputacionGlobal.Rows
    '            Dim drFilaVarios As DataRow = dtVarios.NewRow
    '            drFilaVarios("IDVarios") = drGlobal("IDVarios")
    '            drFilaVarios("DescVarios") = drGlobal("DescVarios")
    '            drFilaVarios("Cantidad") = Nz(drGlobal("Cantidad"), 0)
    '            drFilaVarios("TipoCosteFV") = drGlobal("TipoCosteFV")
    '            drFilaVarios("TipoCosteDI") = drGlobal("TipoCosteDI")
    '            drFilaVarios("Fiscal") = drGlobal("Fiscal")
    '            drFilaVarios("Tasa") = drGlobal("Tasa")
    '            dtVarios.Rows.Add(drFilaVarios)
    '        Next

    '        Return dtVarios
    '    End If
    'End Function

    '<Serializable()> _
    'Public Class StVarios
    '    Public TipoOperacion As String
    '    Public DtVariosGlobal As DataTable
    '    Public ImputacionGlobalReal As Boolean
    '    Public ImputacionGlobalPrev As Boolean
    '    Public Imputar As Boolean

    '    Public Sub New()
    '    End Sub

    '    Public Sub New(ByVal TipoOperacion As String, Optional ByVal DtVariosGlobal As DataTable = Nothing, Optional ByVal ImputacionGlobalReal As Boolean = False, _
    '                   Optional ByVal ImputacionGlobalPrev As Boolean = False, Optional ByVal Imputar As Boolean = False)
    '        Me.TipoOperacion = TipoOperacion
    '        Me.DtVariosGlobal = DtVariosGlobal
    '        Me.ImputacionGlobalReal = ImputacionGlobalReal
    '        Me.ImputacionGlobalPrev = ImputacionGlobalPrev
    '        Me.Imputar = Imputar
    '    End Sub
    'End Class

    '<Task()> Public Shared Function Varios(ByVal data As StVarios, ByVal services As ServiceProvider) As DataTable
    '    Dim dtVarios As DataTable
    '    If data.ImputacionGlobalPrev Then
    '        Dim StVariosTipo As New StVariosTipoOperacion(data.TipoOperacion, data.Imputar)
    '        dtVarios = ProcessServer.ExecuteTask(Of StVariosTipoOperacion, DataTable)(AddressOf VariosTipoOperacion, StVariosTipo, services)
    '    ElseIf data.ImputacionGlobalReal Then
    '        Dim oVV As New BdgVinoVarios
    '        dtVarios = oVV.AddNew()
    '        Dim StVariosTipo As New StVariosTipoOperacionGlobal(data.TipoOperacion, dtVarios, , , data.Imputar, data.DtVariosGlobal)
    '        ProcessServer.ExecuteTask(Of StVariosTipoOperacionGlobal)(AddressOf VariosTipoOperacionGlobal, StVariosTipo, services)
    '    End If
    '    Return dtVarios
    'End Function

#End Region

#Region " Materiales asociados a la operacion "

    '<Serializable()> _
    'Public Class StMaterialesTipoOperacion
    '    Public IDTipoOperacion As String
    '    Public Imputar As Boolean

    '    Public Sub New()
    '    End Sub

    '    Public Sub New(ByVal IDTipoOperacion As String, ByVal Imputar As Boolean)
    '        Me.IDTipoOperacion = IDTipoOperacion
    '        Me.Imputar = Imputar
    '    End Sub
    'End Class

    '<Task()> Public Shared Function MaterialesTipoOperacion(ByVal data As StMaterialesTipoOperacion, ByVal services As ServiceProvider) As DataTable
    '    If Len(data.IDTipoOperacion) > 0 AndAlso Not data.Imputar Then
    '        Dim dtMateriales As DataTable = AdminData.GetData("frmBdgOperacionMaterial", New NoRowsFilterItem)
    '        Dim dtImputacionGlobal As DataTable = AdminData.GetData("NegBdgTipoOperacionMaterial", New StringFilterItem("IDTipoOperacion", data.IDTipoOperacion))

    '        For Each drGlobal As DataRow In dtImputacionGlobal.Rows
    '            Dim material As DataRow = dtMateriales.NewRow
    '            material("IDArticulo") = drGlobal("IDArticulo")
    '            material("IDAlmacen") = ProcessServer.ExecuteTask(Of String, String)(AddressOf BdgGeneral.AlmacenMaterial, material("IDArticulo"), services)
    '            material("DescArticulo") = drGlobal("DescArticulo")
    '            material("Cantidad") = drGlobal("Cantidad")
    '            material("GestionStockPorLotes") = drGlobal("GestionStockPorLotes")
    '            material("RecalcularMaterial") = drGlobal("RecalcularMaterial")
    '            Dim StPrecio As New BdgGeneral.StObtenerPrecio(material("IDArticulo"), material("IDAlmacen"))
    '            material("Precio") = ProcessServer.ExecuteTask(Of BdgGeneral.StObtenerPrecio, Double)(AddressOf BdgGeneral.ObtenerPrecio, StPrecio, services)
    '            material("Merma") = 0
    '            'material("Importe") = material("Precio") * (Nz(material("Cantidad"), 0) + Nz(material("Merma"), 0))

    '            dtMateriales.Rows.Add(material)
    '        Next

    '        Return dtMateriales
    '    End If
    'End Function

    '<Serializable()> _
    'Public Class StMaterialesTipoOperacionGlobal
    '    Public IDTipoOperacion As String
    '    Public Materiales As DataTable
    '    Public Cantidad As Double
    '    Public Fecha As Date
    '    Public NOperacion As String
    '    Public Imputar As Boolean
    '    Public DtMaterialGlobal As DataTable

    '    Public Sub New()
    '    End Sub

    '    Public Sub New(ByVal IDTipoOperacion As String, ByVal Materiales As DataTable, Optional ByVal cantidad As Double = 1, _
    '                   Optional ByVal Fecha As Date = cnMinDate, Optional ByVal NOperacion As String = "", _
    '                   Optional ByVal Imputar As Boolean = False, Optional ByVal DtMaterialGlobal As DataTable = Nothing)
    '        Me.IDTipoOperacion = IDTipoOperacion
    '        Me.Materiales = Materiales
    '        Me.Cantidad = cantidad
    '        Me.Fecha = Fecha
    '        Me.NOperacion = NOperacion
    '        Me.Imputar = Imputar
    '        Me.DtMaterialGlobal = DtMaterialGlobal
    '    End Sub
    'End Class

    '<Task()> Public Shared Sub MaterialesTipoOperacionGlobal(ByVal data As StMaterialesTipoOperacionGlobal, ByVal services As ServiceProvider)
    '    If Length(data.IDTipoOperacion) > 0 Then
    '        Dim componentes As DataTable
    '        If data.Imputar Then
    '            componentes = data.DtMaterialGlobal.Copy
    '        Else
    '            componentes = AdminData.GetData("NegBdgTipoOperacionMaterial", New StringFilterItem("IDTipoOperacion", data.IDTipoOperacion))
    '        End If
    '        For Each componente As DataRow In componentes.Rows
    '            Dim material As DataRow = data.Materiales.NewRow
    '            data.Materiales.Rows.Add(material)
    '            material("IDArticulo") = componente("IDArticulo")
    '            material("IDAlmacen") = ProcessServer.ExecuteTask(Of String, String)(AddressOf BdgGeneral.AlmacenMaterial, material("IDArticulo"), services)
    '            material("DescArticulo") = componente("DescArticulo")
    '            If componente("RecalcularMaterial") Then
    '                material("Cantidad") = componente("Cantidad") * data.Cantidad
    '            Else
    '                material("Cantidad") = componente("Cantidad")
    '            End If
    '            material("GestionStockPorLotes") = componente("GestionStockPorLotes")
    '            material("Merma") = 0

    '            material("RecalcularMaterial") = componente("RecalcularMaterial")
    '            If componentes.Columns.Contains("Merma") Then
    '                material("Merma") = componente("Merma")
    '            End If
    '            If componentes.Columns.Contains("Precio") Then
    '                material("Precio") = componente("Precio")
    '                'material("Importe") = componente("Precio") * (Nz(material("Cantidad"), 0) + Nz(material("Merma"), 0))
    '            Else
    '                Dim StPrecio As New BdgGeneral.StObtenerPrecio(material("IDArticulo"), material("IDAlmacen"))
    '                material("Precio") = ProcessServer.ExecuteTask(Of BdgGeneral.StObtenerPrecio, Double)(AddressOf BdgGeneral.ObtenerPrecio, StPrecio, services)
    '                'If material.Table.Columns.Contains("Importe") Then
    '                '    material("Importe") = material("Precio") * (Nz(material("Cantidad"), 0) + Nz(material("Merma"), 0))
    '                'End If
    '            End If

    '            'revisión pdte.
    '            'If componentes.Columns.Contains("Lote") Then
    '            '    material("Lote") = componente("Lote")
    '            'End If
    '            'If componentes.Columns.Contains("Ubicacion") Then
    '            '    material("Ubicacion") = componente("Ubicacion")
    '            'End If
    '            If data.Materiales.Columns.Contains("Fecha") Then
    '                If data.Fecha <> cnMinDate Then material("Fecha") = data.Fecha
    '                If Length(data.NOperacion) > 0 Then material("NOperacion") = data.NOperacion
    '            End If
    '            If data.Materiales.Columns.Contains("IDOperacionMaterial") Then
    '                If data.Materiales.TableName = "BdgOperacionMaterial" AndAlso Length(material("IDOperacionMaterial")) = 0 Then
    '                    material("IDOperacionMaterial") = Guid.NewGuid()
    '                ElseIf componentes.Columns.Contains("IDOperacionMaterial") Then
    '                    material("IDOperacionMaterial") = componente("IDOperacionMaterial")
    '                End If
    '            ElseIf data.Materiales.Columns.Contains("IDOperacionMaterialPrev") Then
    '                If data.Materiales.TableName = "BdgOperacionMaterialPrev" AndAlso Length(material("IDOperacionMaterialPrev")) = 0 Then
    '                    material("IDOperacionMaterialPrev") = Guid.NewGuid()
    '                End If
    '            End If
    '        Next
    '    End If
    'End Sub

    'Public Class StMaterialesArticulo
    '    Public IDArticulo As String
    '    Public Materiales As DataTable
    '    Public IDEstructura As String
    '    Public Cantidad As Double

    '    Public Sub New()
    '    End Sub

    '    Public Sub New(ByVal IDArticulo As String, ByVal materiales As DataTable, Optional ByVal IDEstructura As String = "", Optional ByVal Cantidad As Double = 1)
    '        Me.IDArticulo = IDArticulo
    '        Me.Materiales = materiales
    '        Me.IDEstructura = IDEstructura
    '        Me.Cantidad = Cantidad
    '    End Sub
    'End Class

    '<Task()> Public Shared Sub MaterialesArticulo(ByVal data As StMaterialesArticulo, ByVal services As ServiceProvider)
    '    If Length(data.IDArticulo) > 0 Then
    '        Dim componentes As DataTable

    '        Dim A As New Negocio.Articulo
    '        Dim rwArticulo As DataRow = A.GetItemRow(data.IDArticulo)

    '        If rwArticulo("TipoEstructura") Then
    '            Dim IDTipoEstructura As String = rwArticulo("IDTipoEstructura") & String.Empty
    '            If Len(IDTipoEstructura) > 0 Then
    '                componentes = AdminData.GetData("NegBdgArticuloMaterial", New StringFilterItem("IDTipoEstructura", IDTipoEstructura))
    '            End If
    '        Else
    '            Dim AE As New Negocio.ArticuloEstructura

    '            If Len(data.IDEstructura) = 0 Then
    '                Dim fltr As New Filter
    '                fltr.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
    '                fltr.Add(New BooleanFilterItem("Principal", True))
    '                Dim estructura As DataTable = AE.Filter(fltr)
    '                If estructura.Rows.Count > 0 Then
    '                    data.IDEstructura = estructura.Rows(0)("IDEstructura")
    '                End If
    '            End If
    '            If Len(data.IDEstructura) <> 0 Then
    '                Dim fltr As New Filter
    '                fltr.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
    '                fltr.Add(New StringFilterItem("IDEstructura", data.IDEstructura))
    '                componentes = AdminData.GetData("NegBdgArticuloMaterial", fltr)
    '            End If
    '        End If

    '        If Not componentes Is Nothing And Not data.Materiales Is Nothing Then
    '            For Each componente As DataRow In componentes.Rows
    '                Dim material As DataRow = data.Materiales.NewRow
    '                data.Materiales.Rows.Add(material)
    '                material("IDArticulo") = componente("IDComponente")
    '                material("DescArticulo") = componente("DescComponente")
    '                material("Cantidad") = Format(componente("Cantidad") * data.Cantidad * (1 + (componente("Merma") / 100)), "##,##0.00")
    '                material("Merma") = componente("Merma")
    '                material("GestionStockPorLotes") = componente("GestionStockPorLotes")
    '                material("IDAlmacen") = ProcessServer.ExecuteTask(Of String, String)(AddressOf BdgGeneral.AlmacenMaterial, material("IDArticulo"), services)
    '                material("RecalcularMaterial") = True
    '                Dim StPrecio As New BdgGeneral.StObtenerPrecio(material("IDArticulo"), material("IDAlmacen"))
    '                material("Precio") = Format(ProcessServer.ExecuteTask(Of BdgGeneral.StObtenerPrecio, Double)(AddressOf BdgGeneral.ObtenerPrecio, StPrecio, services), "##,##0.000")
    '                'material("Importe") = material("Precio") * (Nz(material("Cantidad"), 0) + Nz(material("Merma"), 0))
    '            Next
    '        End If
    '    End If
    'End Sub



   
    <Serializable()> _
    Public Class StMateriales
        Public TipoOperacion As String
        Public NOperacion As String
        Public FechaOperacion As Date
        Public Articulo As String
        Public Cantidad As Double
        Public IDEstructura As String
        Public IDOrden As Integer
        Public DtMaterialGlobal As DataTable
        Public ImputacionGlobalReal As Boolean
        Public ImputacionGlobalPrev As Boolean
        Public Imputar As Boolean

        Public Sub New()
        End Sub

        Public Sub New(ByVal TipoOperacion As String, ByVal NOperacion As String, ByVal FechaOperacion As Date, ByVal Articulo As String, Optional ByVal Cantidad As Double = 1, _
                       Optional ByVal IDEstructura As String = "", Optional ByVal IDOrden As Integer = 0, _
                       Optional ByVal DtMaterialGlobal As DataTable = Nothing, Optional ByVal ImputacionGlobalReal As Boolean = False, _
                       Optional ByVal ImputacionGlobalPrev As Boolean = False, Optional ByVal Imputar As Boolean = False)
            Me.TipoOperacion = TipoOperacion
            Me.Articulo = Articulo
            Me.Cantidad = Cantidad
            Me.IDEstructura = IDEstructura
            Me.IDOrden = IDOrden
            Me.DtMaterialGlobal = DtMaterialGlobal
            Me.ImputacionGlobalReal = ImputacionGlobalReal
            Me.ImputacionGlobalPrev = ImputacionGlobalPrev
            Me.Imputar = Imputar
            Me.NOperacion = NOperacion
            Me.FechaOperacion = FechaOperacion
        End Sub
    End Class

    <Task()> Public Shared Function Materiales(ByVal data As StMateriales, ByVal services As ServiceProvider) As DataTable
        Dim datImputacion As New BdgGeneral.StImputacion(data.TipoOperacion, data.NOperacion, enumBdgOrigenOperacion.Planificada, False, data.FechaOperacion)
        Return ProcessServer.ExecuteTask(Of BdgGeneral.StImputacion, DataTable)(AddressOf BdgGeneral.ImputacionLineaMateriales, datImputacion, services)
    End Function


#End Region

    <Task()> Public Shared Function ValidaMezclaVino(ByVal dtDestino As DataTable, ByVal services As ServiceProvider) As String
        If Not dtDestino Is Nothing AndAlso dtDestino.Rows.Count > 0 Then
            Dim strMezclas As String = String.Empty
            Dim VWC As New BdgWorkClass

            For Each drDestino As DataRow In dtDestino.Rows
                If drDestino.RowState <> DataRowState.Deleted Then
                    Dim drDeposito As DataRow = New BdgDeposito().GetItemRow(drDestino("IDDeposito"))
                    If Not drDeposito("MultiplesVinos") Then
                        Dim dtDepV As DataTable = ProcessServer.ExecuteTask(Of String, DataTable)(AddressOf BdgDepositoVino.SelOnIDDeposito, drDestino("IDDeposito"), services)
                        If Not dtDepV Is Nothing AndAlso dtDepV.Rows.Count > 0 Then
                            If Length(strMezclas) > 0 Then strMezclas = strMezclas & ","
                            strMezclas = strMezclas & drDestino("IDDeposito")
                        End If
                    End If
                End If
            Next
            Return strMezclas
        End If
    End Function

    <Task()> Public Shared Function DevolverOcupacion(ByVal IDDeposito As String, ByVal services As ServiceProvider) As Double
        Dim suma As Double
        If Len(IDDeposito) > 0 Then
            Dim art As New Articulo
            Dim objVino As New BdgVino

            Dim strUDLitros As String = New BdgParametro().UnidadesCampoLitros
            Dim dtDepositoVino As DataTable = ProcessServer.ExecuteTask(Of String, DataTable)(AddressOf BdgDepositoVino.SelOnIDDeposito, IDDeposito, services)
            For Each drDepositoVino As DataRow In dtDepositoVino.Rows
                Dim rwVino As DataRow = objVino.GetItemRow(drDepositoVino(_DV.IDVino))
                Dim rwArt As DataRow = art.GetItemRow(rwVino(_V.IDArticulo))
                If Not rwArt.IsNull("IDUDInterna") Then
                    Dim dataFactor As New ArticuloUnidadAB.DatosFactorConversion(rwVino(_V.IDArticulo), rwArt("IDUDInterna"), strUDLitros)
                    Dim Factor As Double = ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, dataFactor, services)

                    suma += drDepositoVino(_DV.Cantidad) * Factor
                End If
            Next
        End If

        Return suma
    End Function

#Region "Guardar Globales"

    <Task()> Public Shared Sub guardarMaterialesGlobales(ByVal dtMaterialGlobal As DataTable, ByVal services As ServiceProvider)
        If Not dtMaterialGlobal Is Nothing Then
            Dim om As New BdgOperacionMaterial
            If Not dtMaterialGlobal.GetChanges(DataRowState.Deleted) Is Nothing AndAlso dtMaterialGlobal.GetChanges(DataRowState.Deleted).Rows.Count > 0 Then
                Dim dtMaterialGlobalDEL As DataTable = dtMaterialGlobal.Clone
                For Each Dr As DataRow In dtMaterialGlobal.GetChanges(DataRowState.Deleted).Rows
                    Dr.RejectChanges()
                    dtMaterialGlobalDEL.Rows.Add(Dr.ItemArray)
                Next
                dtMaterialGlobalDEL.AcceptChanges()
                dtMaterialGlobalDEL.TableName = om.Entity
                om.Delete(dtMaterialGlobalDEL)
            End If

            Dim dtMaterialGlobalMod As DataTable = dtMaterialGlobal.GetChanges(DataRowState.Modified + DataRowState.Added)
            If Not dtMaterialGlobalMod Is Nothing AndAlso dtMaterialGlobalMod.Rows.Count > 0 Then
                dtMaterialGlobalMod.TableName = om.Entity
                om.Update(dtMaterialGlobalMod)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub guardarMODGlobales(ByVal dtMODGlobal As DataTable, ByVal services As ServiceProvider)
        If Not dtMODGlobal Is Nothing Then
            Dim oMOD As New BdgOperacionMOD
            If Not dtMODGlobal.GetChanges(DataRowState.Deleted) Is Nothing AndAlso dtMODGlobal.GetChanges(DataRowState.Deleted).Rows.Count > 0 Then
                Dim dtMODGlobalDEL As DataTable = dtMODGlobal.Clone
                For Each Dr As DataRow In dtMODGlobal.GetChanges(DataRowState.Deleted).Rows
                    Dr.RejectChanges()
                    dtMODGlobalDEL.Rows.Add(Dr.ItemArray)
                Next
                dtMODGlobalDEL.AcceptChanges()
                dtMODGlobalDEL.TableName = oMOD.Entity
                oMOD.Delete(dtMODGlobalDEL)
            End If

            Dim dtMODGlobalMod As DataTable = dtMODGlobal.GetChanges(DataRowState.Modified + DataRowState.Added)
            If Not dtMODGlobalMod Is Nothing AndAlso dtMODGlobalMod.Rows.Count > 0 Then
                dtMODGlobalMod.TableName = oMOD.Entity
                oMOD.Update(dtMODGlobalMod)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub guardarCentroGlobales(ByVal dtCentroGlobal As DataTable, ByVal services As ServiceProvider)
        If Not dtCentroGlobal Is Nothing Then
            Dim oCentro As New BdgOperacionCentro
            If Not dtCentroGlobal.GetChanges(DataRowState.Deleted) Is Nothing AndAlso dtCentroGlobal.GetChanges(DataRowState.Deleted).Rows.Count > 0 Then
                Dim dtCentroGlobalDEL As DataTable = dtCentroGlobal.Clone
                For Each Dr As DataRow In dtCentroGlobal.GetChanges(DataRowState.Deleted).Rows
                    Dr.RejectChanges()
                    dtCentroGlobalDEL.Rows.Add(Dr.ItemArray)
                Next
                dtCentroGlobalDEL.AcceptChanges()
                dtCentroGlobalDEL.TableName = oCentro.Entity
                oCentro.Delete(dtCentroGlobalDEL)
            End If

            Dim dtCentroGlobalCentro As DataTable = dtCentroGlobal.GetChanges(DataRowState.Modified + DataRowState.Added)
            If Not dtCentroGlobalCentro Is Nothing AndAlso dtCentroGlobalCentro.Rows.Count > 0 Then
                dtCentroGlobalCentro.TableName = oCentro.Entity
                oCentro.Update(dtCentroGlobalCentro)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub guardarVariosGlobales(ByVal dtVariosGlobal As DataTable, ByVal services As ServiceProvider)
        If Not dtVariosGlobal Is Nothing Then
            Dim oVarios As New BdgOperacionVarios
            If Not dtVariosGlobal.GetChanges(DataRowState.Deleted) Is Nothing AndAlso dtVariosGlobal.GetChanges(DataRowState.Deleted).Rows.Count > 0 Then
                Dim dtVariosGlobalDEL As DataTable = dtVariosGlobal.Clone
                For Each Dr As DataRow In dtVariosGlobal.GetChanges(DataRowState.Deleted).Rows
                    Dr.RejectChanges()
                    dtVariosGlobalDEL.Rows.Add(Dr.ItemArray)
                Next
                dtVariosGlobalDEL.AcceptChanges()
                dtVariosGlobalDEL.TableName = oVarios.Entity
                oVarios.Delete(dtVariosGlobalDEL)
            End If

            Dim dtVariosGlobalVarios As DataTable = dtVariosGlobal.GetChanges(DataRowState.Modified + DataRowState.Added)
            If Not dtVariosGlobalVarios Is Nothing AndAlso dtVariosGlobalVarios.Rows.Count > 0 Then
                dtVariosGlobalVarios.TableName = oVarios.Entity
                oVarios.Update(dtVariosGlobalVarios)
            End If
        End If
    End Sub

#End Region

#Region "Reparto"

    <Serializable()> _
    Public Class StReparto
        Public DtDestino As DataTable
        Public DtImputacionGlobal As DataTable
        Public QTotal As Double

        Public Sub New()
        End Sub

        Public Sub New(ByVal DtDestino As DataTable, ByVal DtImputacionGlobal As DataTable, ByVal QTotal As Double)
            Me.DtDestino = DtDestino
            Me.DtImputacionGlobal = DtImputacionGlobal
            Me.QTotal = QTotal
        End Sub
    End Class

    <Task()> Public Shared Function ManoObraReparto(ByVal data As StReparto, ByVal services As ServiceProvider) As DataTable
        If Not data.DtDestino Is Nothing AndAlso Not data.DtImputacionGlobal Is Nothing Then
            Dim dtManoObra As DataTable = New BE.DataEngine().Filter("frmBdgOperacionMod", New NoRowsFilterItem)
            For Each drGlobal As DataRow In data.DtImputacionGlobal.Select
                For Each drDestino As DataRow In data.DtDestino.Select
                    If drDestino.RowState <> DataRowState.Deleted And drGlobal.RowState <> DataRowState.Deleted Then
                        Dim manoObra As DataRow = dtManoObra.NewRow
                        manoObra("IDVino") = drDestino("IDVino")
                        manoObra("IDOperario") = drGlobal("IDOperario")
                        manoObra("Fecha") = drGlobal("Fecha")
                        If drGlobal.Table.Columns.Contains("DescOperario") Then
                            manoObra("DescOperario") = drGlobal("DescOperario")
                        End If
                        If drGlobal.Table.Columns.Contains("DescOperacion") Then
                            manoObra("DescOperacion") = drGlobal("DescOperacion")
                        End If

                        If data.QTotal <> 0 Then
                            manoObra("Tiempo") = (Nz(drDestino("Cantidad"), 0) / data.QTotal) * Nz(drGlobal("Tiempo"), 0) ', "##,##0.00")
                        Else : manoObra("Tiempo") = 0
                        End If

                        manoObra("IDCategoria") = drGlobal("IDCategoria") & String.Empty
                        manoObra("Tasa") = Nz(drGlobal("Tasa"), 0)
                        manoObra("Importe") = drGlobal("Tiempo") * manoObra("Tasa")
                        manoObra("IDHora") = drGlobal("IDHora")
                        manoObra("NOperacion") = drGlobal("NOperacion")
                        manoObra("IDOperacionMod") = drGlobal("IDOperacionMod")
                        dtManoObra.Rows.Add(manoObra)
                    End If
                Next
            Next
            Return dtManoObra
        End If
    End Function

    <Task()> Public Shared Function MaterialesReparto(ByVal data As StReparto, ByVal services As ServiceProvider) As DataTable
        If Not data.DtDestino Is Nothing AndAlso Not data.DtImputacionGlobal Is Nothing Then
            Dim dtMateriales As DataTable = New BE.DataEngine().Filter("frmBdgOperacionMaterial", New NoRowsFilterItem)

            For Each drGlobal As DataRow In data.DtImputacionGlobal.Select
                For Each drDestino As DataRow In data.DtDestino.Select
                    If drDestino.RowState <> DataRowState.Deleted And drGlobal.RowState <> DataRowState.Deleted Then
                        Dim material As DataRow = dtMateriales.NewRow
                        material("IDVino") = drDestino("IDVino")
                        material("IDArticulo") = drGlobal("IDArticulo")
                        material("IDAlmacen") = drGlobal("IDAlmacen")
                        If drGlobal.Table.Columns.Contains("DescArticulo") Then
                            material("DescArticulo") = drGlobal("DescArticulo")
                        End If
                        If data.QTotal <> 0 Then
                            material("Cantidad") = xRound((Nz(drDestino("Cantidad"), 0) / data.QTotal) * Nz(drGlobal("Cantidad"), 0), BdgGeneral.NUM_DECIMALES_CANTIDADES)
                            material("Merma") = xRound((Nz(drDestino("Cantidad"), 0) / data.QTotal) * Nz(drGlobal("Merma"), 0), BdgGeneral.NUM_DECIMALES_CANTIDADES)
                        Else
                            material("Cantidad") = 0
                            material("Merma") = 0
                        End If

                        If drGlobal.Table.Columns.Contains("GestionStockPorLotes") Then
                            material("GestionStockPorLotes") = drGlobal("GestionStockPorLotes")
                        End If
                        material("RecalcularMaterial") = drGlobal("RecalcularMaterial")
                        material("Precio") = Nz(drGlobal("Precio"), 0)
                        material("Importe") = material("Precio") * (material("Cantidad") + Nz(material("Merma"), 0))
                        'material("Lote") = drGlobal("Lote")
                        'material("Ubicacion") = drGlobal("Ubicacion")
                        material("Fecha") = drGlobal("Fecha")
                        material("NOperacion") = drGlobal("NOperacion")
                        material("IDOperacionMaterial") = drGlobal("IDOperacionMaterial")
                        dtMateriales.Rows.Add(material)
                    End If
                Next
            Next
            Return dtMateriales
        End If
    End Function

    <Task()> Public Shared Function CentroReparto(ByVal data As StReparto, ByVal services As ServiceProvider) As DataTable
        Dim clsOperacionCentro As New BdgOperacionCentro
        If Not data.DtDestino Is Nothing AndAlso Not data.DtImputacionGlobal Is Nothing Then
            Dim dtCentro As DataTable = New BE.DataEngine().Filter("frmBdgOperacionCentro", New NoRowsFilterItem)

            For Each drGlobal As DataRow In data.DtImputacionGlobal.Select
                For Each drDestino As DataRow In data.DtDestino.Select
                    If drDestino.RowState <> DataRowState.Deleted And drGlobal.RowState <> DataRowState.Deleted Then
                        Dim filaCentro As DataRow = dtCentro.NewRow
                        filaCentro("IDVino") = drDestino("IDVino")
                        filaCentro("IDCentro") = drGlobal("IDCentro")
                        filaCentro("Fecha") = drGlobal("Fecha")
                        If drGlobal.Table.Columns.Contains("DescCentro") Then
                            filaCentro("DescCentro") = drGlobal("DescCentro")
                        End If
                        If drGlobal.Table.Columns.Contains("DescOperacion") Then
                            filaCentro("DescOperacion") = drGlobal("DescOperacion")
                        End If

                        If data.QTotal <> 0 Then
                            filaCentro("Tiempo") = (Nz(drDestino("Cantidad"), 0) / data.QTotal) * Nz(drGlobal("Tiempo"), 0) ', "##,##0.00")
                        Else
                            filaCentro("Tiempo") = Nz(drGlobal("Tiempo"), 0)
                        End If
                        filaCentro("Cantidad") = drDestino("Cantidad") 'dblQTotal

                        filaCentro("Tasa") = Nz(drGlobal("Tasa"), 0)
                        filaCentro("NOperacion") = drGlobal("NOperacion")
                        filaCentro("IDOperacionCentro") = drGlobal("IDOperacionCentro")

                        filaCentro("IdUdMedidaCentro") = drGlobal("IdUdMedidaCentro")
                        filaCentro("IdUdMedidaArticulo") = drDestino("IdUdMedida")
                        Dim UDTiempo As Integer
                        'Dim IDVino As Guid = GetCurrentVino(grdDestino)
                        If drGlobal("UDTiempo") Is System.DBNull.Value Then
                            UDTiempo = enumstdUdTiempo.Horas 'Por defecto
                        Else
                            UDTiempo = drGlobal("UDTiempo")
                        End If
                        filaCentro("UDTiempo") = UDTiempo
                        filaCentro("PorCantidad") = drGlobal("PorCantidad")

                        dtCentro.Rows.Add(filaCentro)
                    End If
                Next
            Next
            Return dtCentro
        End If
    End Function

    <Task()> Public Shared Function VariosReparto(ByVal data As StReparto, ByVal services As ServiceProvider) As DataTable
        If Not data.DtDestino Is Nothing AndAlso Not data.DtImputacionGlobal Is Nothing Then
            Dim dtVarios As DataTable = AdminData.GetData("frmBdgOperacionVarios", New NoRowsFilterItem)

            For Each drGlobal As DataRow In data.DtImputacionGlobal.Select
                For Each drDestino As DataRow In data.DtDestino.Select
                    If drDestino.RowState <> DataRowState.Deleted And drGlobal.RowState <> DataRowState.Deleted Then
                        Dim filaCentro As DataRow = dtVarios.NewRow
                        filaCentro("IDVino") = drDestino("IDVino")
                        filaCentro("IDVarios") = drGlobal("IDVarios")
                        filaCentro("Fecha") = drGlobal("Fecha")
                        If drGlobal.Table.Columns.Contains("DescVarios") Then
                            filaCentro("DescVarios") = drGlobal("DescVarios")
                        End If
                        If drGlobal.Table.Columns.Contains("DescOperacion") Then
                            filaCentro("DescOperacion") = drGlobal("DescOperacion")
                        End If

                        If data.QTotal <> 0 Then
                            filaCentro("Cantidad") = (Nz(drDestino("Cantidad"), 0) / data.QTotal) * Nz(drGlobal("Cantidad"), 0)  ' * Nz(drGlobal("Tiempo"), 0) ', "##,##0.00")
                        Else
                            filaCentro("Cantidad") = 0
                        End If
                        'filaCentro("Cantidad") = drDestino("Cantidad")
                        'filaCentro("Tiempo") = Nz(drGlobal("Tiempo"), 0)
                        filaCentro("Tasa") = Nz(drGlobal("Tasa"), 0)
                        'filaCentro("IDHora") = drGlobal("IDHora")
                        filaCentro("NOperacion") = drGlobal("NOperacion")
                        filaCentro("IDOperacionVarios") = drGlobal("IDOperacionVarios")

                        'filaCentro("IdUdMedidaCentro") = drGlobal("IdUdMedidaCentro")
                        'filaCentro("IdUdMedidaArticulo") = drGlobal("IdUdMedidaCentro")
                        Dim UDTiempo As Integer
                        'Dim IDVino As Guid = GetCurrentVino(grdDestino)

                        filaCentro("Importe") = Nz(drGlobal("Cantidad"), 0) * Nz(drGlobal("Tasa"), 0)
                        dtVarios.Rows.Add(filaCentro)
                    End If
                Next
            Next
            Return dtVarios
        End If
    End Function

    <Serializable()> _
    Public Class StRepartoLotes
        Public DtMateriales As DataTable
        Public DtImputacionGlobal As DataTable
        Public DtImputacionGlobalLote As DataTable

        Public Sub New(ByVal dtMateriales As DataTable, ByVal dtImputacionGlobal As DataTable, ByVal dtImputacionGlobalLote As DataTable)
            Me.DtImputacionGlobal = dtImputacionGlobal
            Me.DtImputacionGlobalLote = dtImputacionGlobalLote
            Me.DtMateriales = dtMateriales
        End Sub
    End Class

    <Task()> Public Shared Function RepartoLotes(ByVal data As StRepartoLotes, ByVal services As ServiceProvider) As DataTable
        'TODO - Validar
        'se puede hacer una task que lo valide ya que desde presentación tmb se tiene q validar
        Dim dttResult As DataTable = New BdgVinoMaterialLote().AddNew

        Dim bsnOperacionMaterialLote As New BdgOperacionMaterialLote
        Dim ClsArt As New Articulo
        For Each dtrImputacionGlobal As DataRow In data.DtImputacionGlobal.Rows
            Dim DtArt As DataTable = ClsArt.SelOnPrimaryKey(dtrImputacionGlobal("IDArticulo"))
            If (DtArt.Rows(0)("GestionStockPorLotes")) Then 'If (dtrImputacionGlobal("GestionStockPorLotes")) Then
                Dim filter As New Filter
                filter.Add("IDOperacionMaterial", dtrImputacionGlobal("IDOperacionMaterial"))


                If data.DtImputacionGlobalLote Is Nothing _
                            OrElse data.DtImputacionGlobalLote.Rows.Count = 0 Then
                    data.DtImputacionGlobalLote = New BdgOperacionMaterialLote().Filter(filter)
                Else
                    If data.DtImputacionGlobalLote.Select(filter.Compose(New AdoFilterComposer)).Length = 0 Then
                        Dim dtExtra As DataTable = New BdgOperacionMaterialLote().Filter(filter)
                        For Each dtrExtra As DataRow In dtExtra.Rows
                            data.DtImputacionGlobalLote.ImportRow(dtrExtra)
                        Next
                    End If
                End If


                If Not data.DtImputacionGlobalLote Is Nothing _
                    AndAlso data.DtImputacionGlobalLote.Rows.Count > 0 Then
                    'reparto
                    Dim dblCantidadAcumulada As Double = 0
                    Dim dtrImputacionesGlobalesLote() As DataRow = data.DtImputacionGlobalLote.Select(filter.Compose(New AdoFilterComposer))
                    For Each dtrImputacionGlobalLote As DataRow In dtrImputacionesGlobalesLote
                        Dim dtrMat() As DataRow = data.DtMateriales.Select(filter.Compose(New AdoFilterComposer))

                        For Each dtrMaterial As DataRow In dtrMat
                            If (dtrMaterial.RowState <> DataRowState.Deleted) Then
                                Dim dtrLoteMaterial As DataRow = dttResult.NewRow
                                dtrLoteMaterial(_BdgVinoMaterialLote.IDVinoMaterialLote) = Guid.NewGuid
                                dtrLoteMaterial(_BdgVinoMaterialLote.IDVinoMaterial) = dtrMaterial("IDVinoMaterial")
                                dtrLoteMaterial(_BdgVinoMaterialLote.Lote) = dtrImputacionGlobalLote("Lote")
                                dtrLoteMaterial(_BdgVinoMaterialLote.Ubicacion) = dtrImputacionGlobalLote("Ubicacion")

                                Dim dblTotalMaterial As Double = dtrMaterial("Cantidad") + Nz(dtrMaterial("Merma"), 0)
                                Dim dblTotalImpu As Double = dtrImputacionGlobal(_VM.Cantidad) + Nz(dtrImputacionGlobal(_VM.Merma), 0)
                                dtrLoteMaterial(_BdgVinoMaterialLote.Cantidad) = dtrImputacionGlobalLote("Cantidad") * dblTotalMaterial / dblTotalImpu
                                If (Length(dtrImputacionGlobalLote(_BdgVinoMaterialLote.SeriePrecinta)) > 0) Then
                                    dtrLoteMaterial(_BdgVinoMaterialLote.SeriePrecinta) = dtrImputacionGlobalLote("SeriePrecinta")
                                    dtrLoteMaterial(_BdgVinoMaterialLote.NDesdePrecinta) = dtrImputacionGlobalLote("NDesde") + dblCantidadAcumulada
                                    dtrLoteMaterial(_BdgVinoMaterialLote.NHastaPrecinta) = dtrLoteMaterial(_BdgVinoMaterialLote.NDesdePrecinta) + dtrLoteMaterial(_BdgVinoMaterialLote.Cantidad) - 1
                                End If
                                dblCantidadAcumulada += dtrLoteMaterial(_BdgVinoMaterialLote.Cantidad)

                                'redondeo?? OJO
                                dttResult.Rows.Add(dtrLoteMaterial)
                            End If
                        Next
                    Next
                End If

            End If
        Next
        Return dttResult
    End Function

#End Region

    Public Class StValidarVinoRepetidoOperacion
        Public DrOri As DataRow
        Public DrsDestinoAux() As DataRow

        Public Sub New(ByVal DrOri As DataRow, ByVal DrsDestinoAux() As DataRow)
            Me.DrOri = DrOri
            Me.DrsDestinoAux = DrsDestinoAux
        End Sub
    End Class

    <Task()> Public Shared Function ValidarVinoRepetidoOperacion(ByVal data As BdgOperacion.StValidarVinoRepetidoOperacion, ByVal services As ServiceProvider) As Boolean
        'Comprobar si el Vino ya está en la Operación
        If Length(data.DrOri(_V.IDArticulo)) > 0 AndAlso Length(data.DrOri(_V.IDDeposito)) > 0 AndAlso Length(data.DrOri(_V.Lote)) > 0 Then
            Dim VinoExiste As List(Of DataRow) = (From c In data.DrsDestinoAux Where c.RowState <> DataRowState.Deleted AndAlso _
                                                                                     Not c.IsNull(_V.IDArticulo) AndAlso _
                                                                                     Not c.IsNull(_V.IDDeposito) AndAlso _
                                                                                     Not c.IsNull(_V.Lote) AndAlso _
                                                                                    c(_V.IDArticulo) = data.DrOri(_V.IDArticulo) AndAlso _
                                                                                    c(_V.IDDeposito) = data.DrOri(_V.IDDeposito) AndAlso _
                                                                                    c(_V.Lote) = data.DrOri(_V.Lote) AndAlso _
                                                                                    Not c(_V.IDVino).Equals(data.DrOri(_V.IDVino))).ToList
            If Not VinoExiste Is Nothing AndAlso VinoExiste.Count > 0 Then
                ApplicationService.GenerateError("El vino siguiente no puede estar en Origen y Destino a la vez:|- Depósito: |- Artículo: |- Lote: |", vbNewLine, data.DrOri(_V.IDDeposito) & vbNewLine, data.DrOri(_V.IDArticulo) & vbNewLine, data.DrOri(_V.Lote))
            End If
        End If
    End Function

    <Serializable()> _
    Public Class dataAplicarRulesMerma
        Public IDArticulo As String
        Public Cantidad As Double
        Public Merma As Double

        Public Sub New(ByVal IDArticulo As String, ByVal Cantidad As Double, ByVal Merma As Double)
            Me.IDArticulo = IDArticulo
            Me.Cantidad = Cantidad
            Me.Merma = Merma
        End Sub
    End Class

    <Task()> Public Shared Function ValidarMermaMaterial(ByVal data As dataAplicarRulesMerma, ByVal services As ServiceProvider) As Boolean
        If Length(data.IDArticulo) > 0 AndAlso data.Cantidad > 0 AndAlso data.Merma > 0 Then
            Dim dt As DataTable = New BE.DataEngine().Filter("vNegBdgGetPorcMermaMaxima", New StringFilterItem("IDArticulo", data.IDArticulo))
            If dt.Rows.Count > 0 Then
                Dim PorcMerma As Double = (100 * data.Merma) / data.Cantidad
                Dim SuperaMerma As Boolean = False

                If Length(dt.Rows(0)("PorcMermaMaximaSubfamilia")) > 0 Then
                    SuperaMerma = dt.Rows(0)("PorcMermaMaximaSubfamilia") < PorcMerma
                ElseIf Length(dt.Rows(0)("PorcMermaMaxima")) > 0 Then
                    SuperaMerma = dt.Rows(0)("PorcMermaMaxima") < PorcMerma
                End If
                If SuperaMerma Then
                    Return ProcessServer.ExecuteTask(Of Integer, Boolean)(AddressOf Business.Rules.Rules.TienePermisos, Business.Rules.Rules.enumRules.MermaMateriales, services)
                End If
            End If
        End If
        Return True
    End Function

    <Serializable()> _
    Public Class dataAplicarRulesMermaOperacion
        Public PermitirMerma As Boolean
        Public PorcMermaMaxima As String
        Public TotalPendiente As Double
        Public TotalOrigen As Double
        Public TotalDestino As Double

        Public Sub New(ByVal PermitirMerma As Boolean, ByVal PorcMermaMaxima As String, ByVal TotalPendiente As Double, ByVal TotalOrigen As Double, ByVal TotalDestino As Double)
            Me.PermitirMerma = PermitirMerma
            Me.PorcMermaMaxima = PorcMermaMaxima
            Me.TotalPendiente = TotalPendiente
            Me.TotalOrigen = TotalOrigen
            Me.TotalDestino = TotalDestino
        End Sub
    End Class

    <Task()> Public Shared Function ValidarMermaOperacion(ByVal data As dataAplicarRulesMermaOperacion, ByVal services As ServiceProvider) As Boolean
        Dim Merma As Double = Math.Abs(data.TotalOrigen - data.TotalDestino)
        If Not data.PermitirMerma And Merma <> 0 Then
            Return False
        ElseIf Merma > 0 AndAlso Length(data.PorcMermaMaxima) > 0 Then
            If data.TotalOrigen * CDbl(data.PorcMermaMaxima) / 100 < Merma Then
                Return ProcessServer.ExecuteTask(Of Integer, Boolean)(AddressOf Business.Rules.Rules.TienePermisos, Business.Rules.Rules.enumRules.MermaOperaciones, services)
            End If
        End If
        Return True
    End Function

#Region " Gestion con OFS"
    <Serializable()> _
Public Class StActualizarOF
        Public DtDst As DataTable
        Public Fecha As Date

        Public Sub New()
        End Sub

        Public Sub New(ByVal DtDst As DataTable, ByVal Fecha As Date)
            Me.DtDst = DtDst
            Me.Fecha = Fecha
        End Sub
    End Class

    'TODO este código debiera estar en el negocio de producción
    <Task()> Public Shared Sub ActualizarOF(ByVal data As StActualizarOF, ByVal services As ServiceProvider)
        Dim oOF As BusinessHelper = BusinessHelper.CreateBusinessObject("OrdenFabricacion")
        For Each rwDst As DataRow In data.DtDst.Rows
            Dim IDOrden As Integer = 0
            Dim QOld As Double = 0
            Dim QNew As Double = 0
            If rwDst.HasVersion(DataRowVersion.Current) Then
                QNew = rwDst(_OV.Cantidad)
                If Not rwDst.IsNull(_OV.IDOrden) Then IDOrden = rwDst(_OV.IDOrden)
            End If
            If rwDst.HasVersion(DataRowVersion.Original) Then
                QOld = rwDst(_OV.Cantidad, DataRowVersion.Original)
                If IDOrden = 0 Then
                    If Not IsDBNull(rwDst(_OV.IDOrden, DataRowVersion.Original)) Then IDOrden = rwDst(_OV.IDOrden, DataRowVersion.Original)
                End If
            End If

            If IDOrden <> 0 Then
                Dim rwOF As DataRow = oOF.GetItemRow(IDOrden)

                rwOF("QFabricada") += QNew - QOld
                rwOF("QIniciada") += QNew - QOld

                If rwOF("Estado") = enumofEstado.ofePlanificada And rwOF("QIniciada") <> 0 Then
                    rwOF("Estado") = enumofEstado.ofeIniciada
                Else
                    rwOF("Estado") = enumofEstado.ofePlanificada
                End If

                If rwOF("QFabricada") >= rwOF("QFabricar") Then
                    Dim p As New Parametro
                    If p.CierreAutomaticoOF() Then
                        rwOF("Estado") = enumofEstado.ofeTerminada
                    End If
                End If
                If rwOF("Estado") = enumofEstado.ofeTerminada Then
                    rwOF("FechaFinReal") = data.Fecha
                End If
                oOF.Update(rwOF.Table)
            End If
        Next
    End Sub

  
#End Region

#Region " AjusteNegativoOF "

    <Serializable()> _
    Public Class dataAjusteNegativoOF
        Public IDOrden As Integer
        Public NOrden As String
        Public IDArticulo As String
        Public IDAlmacen As String
        Public Lote As String
        Public IDUbicacion As String
        Public Fecha As Date
        Public Cantidad As Double
        Public QFabricada As Double
        Public TipoMovimientoSalida As enumTipoMovimiento
        Public dtLineas As DataTable
        Public TipoMovimientoEntrada As enumTipoMovimiento

        Public Sub New(ByVal IDOrden As Integer, ByVal NOrden As String, ByVal QFabricada As Double, ByVal IDArticulo As String, ByVal IDAlmacen As String, ByVal Lote As String, ByVal IDUbicacion As String, ByVal Fecha As Date, ByVal Cantidad As Double, ByVal TipoMovimientoSalida As enumTipoMovimiento, ByVal dtLineas As DataTable, ByVal TipoMovimientoEntrada As enumTipoMovimiento)
            Me.IDOrden = IDOrden
            Me.NOrden = NOrden
            Me.QFabricada = QFabricada
            Me.IDArticulo = IDArticulo
            Me.IDAlmacen = IDAlmacen
            Me.Lote = Lote
            Me.IDUbicacion = IDUbicacion
            Me.Fecha = Fecha
            Me.Cantidad = Cantidad
            Me.TipoMovimientoSalida = TipoMovimientoSalida
            Me.dtLineas = dtLineas
            Me.TipoMovimientoEntrada = TipoMovimientoEntrada
        End Sub
    End Class
    <Task()> Public Shared Function AjusteNegativoOF(ByVal data As dataAjusteNegativoOF, ByVal services As ServiceProvider) As StockUpdateData()
        AdminData.BeginTx()
        Dim u(-1) As StockUpdateData

        Dim MovSinActualizar As Boolean = False
        'Movimiento salida
        Dim NumeroMovimiento As Integer = ProcessServer.ExecuteTask(Of Object, Integer)(AddressOf ProcesoStocks.NuevoNumeroMovimiento, Nothing, services)

        Dim sdSalida As New StockData(data.IDArticulo, data.IDAlmacen, data.Lote, data.IDUbicacion, data.Cantidad, 0, 0, data.Fecha, data.TipoMovimientoSalida)
        sdSalida.IDDocumento = data.IDOrden
        sdSalida.Documento = data.NOrden

        Dim dataMovimiento As New DataNumeroMovimientoSinc(NumeroMovimiento, sdSalida)
        Dim s As StockUpdateData = ProcessServer.ExecuteTask(Of DataNumeroMovimientoSinc, StockUpdateData)(AddressOf ProcesoStocks.Salida, dataMovimiento, services)
        If s.Estado = EstadoStock.NoActualizado Then
            MovSinActualizar = True
        End If
        ArrayManager.Copy(s, u)

        'Movimiento entrada
        If Not data.dtLineas Is Nothing AndAlso data.dtLineas.Rows.Count > 0 Then
            For Each drLinea As DataRow In data.dtLineas.Select
                Dim sdEntradas(-1) As StockData
                Dim sdEntrada As StockData = New StockData(drLinea("IDComponente"), drLinea("IDAlmacen"), drLinea("Lote") & String.Empty, drLinea("Ubicacion") & String.Empty, drLinea("QDescontar"), 0, 0, data.Fecha, data.TipoMovimientoEntrada)
                sdEntrada.IDDocumento = data.IDOrden
                sdEntrada.Documento = data.NOrden

                dataMovimiento = New DataNumeroMovimientoSinc(NumeroMovimiento, sdEntrada)
                Dim e As StockUpdateData = ProcessServer.ExecuteTask(Of DataNumeroMovimientoSinc, StockUpdateData)(AddressOf ProcesoStocks.Entrada, dataMovimiento, services)
                If s.Estado = EstadoStock.NoActualizado Then
                    MovSinActualizar = True
                End If
                ArrayManager.Copy(e, u)
            Next
        End If

        If Not MovSinActualizar Then
            Dim dd As New ProcesoBdgOperacion.dataActualizarEstadoOF(data.IDOrden, data.Cantidad, data.QFabricada)
            ProcessServer.ExecuteTask(Of ProcesoBdgOperacion.dataActualizarEstadoOF)(AddressOf ProcesoBdgOperacion.ActualizarEstadoOF, dd, services)
        Else
            AdminData.RollBackTx()
        End If

        Return u
    End Function

#End Region

#End Region

End Class

<Serializable()> _
Public Class OperacionUpdateData
    Public Operaciones() As String
    Public NOperacion() As String
End Class