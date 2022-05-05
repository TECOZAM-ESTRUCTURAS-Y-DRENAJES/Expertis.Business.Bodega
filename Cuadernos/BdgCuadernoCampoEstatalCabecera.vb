Public Class BdgCuadernoCampoEstatalCabecera

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    
    Private Const cnEntidad As String = "tbBdgCuadernoCampoEstatalCabecera"

#End Region

#Region "Eventos Entidad"
    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf CargarContador)
        addnewProcess.AddTask(Of DataRow)(AddressOf AsignarClavePrimaria)
    End Sub

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarContador)
    End Sub
    
    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
    End Sub
#End Region

#Region "Funciones Públicas"
    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("DescCuaderno")) = 0 Then ApplicationService.GenerateError("La Descripción es obligatoria.")
        If Length(data("FechaDesde")) = 0 Then ApplicationService.GenerateError("La Fecha Desde es obligatoria.")
        If Length(data("FechaHasta")) = 0 Then ApplicationService.GenerateError("La Fecha Hasta es obligatoria.")
        If Length(data("IDCuadernoEstado")) = 0 Then ApplicationService.GenerateError("El Estado es obligatorio.")
    End Sub

    <Task()> Public Shared Sub CargarContador(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim StDatos As New Contador.DatosDefaultCounterValue
        StDatos.row = data
        StDatos.EntityName = "BdgCuadernoCampoEstatalCabecera"
        StDatos.FieldName = "NCuaderno"
        ProcessServer.ExecuteTask(Of Contador.DatosDefaultCounterValue)(AddressOf Contador.LoadDefaultCounterValue, StDatos, services)
    End Sub

    <Task()> Public Shared Sub AsignarContador(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDContador")) > 0 Then
                data("NCuaderno") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, data("IDContador"), services)
            Else
                Dim StDatos As New Contador.DatosDefaultCounterValue
                StDatos.row = data
                StDatos.EntityName = "BdgCuadernoCampoEstatalCabecera"
                StDatos.FieldName = "NCuaderno"
                ProcessServer.ExecuteTask(Of Contador.DatosDefaultCounterValue)(AddressOf Contador.LoadDefaultCounterValue, StDatos, services)
                If Length(data("IDContador")) > 0 Then
                    data("NCuaderno") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, data("IDContador"), services)
                Else
                    ApplicationService.GenerateError("No se ha configurado contador predeterminado para el Cuaderno de Explotación.")
                End If

            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then data("IDCuaderno") = Guid.NewGuid
    End Sub
#End Region

#Region "Copia Cuaderno "
    <Serializable()> _
      Public Class DatosCuadernoCopia
        Public IDCuaderno As String
        Public IDContador As String
        Public NCuaderno As String
        Public FechaDesde As String
        Public FechaHasta As String
        Public BlnCopyDatosGenerales As Boolean
        Public BlnCopyInformacionGeneral As Boolean
        Public BlnCopyPersonas As Boolean
        Public BlnCopyEquipos As Boolean
        Public BlnCopyAsesores As Boolean
        Public BlnCopyFincas As Boolean
        Public BlnCopyZonas As Boolean
        Public BlnCopyAnalisisFito As Boolean
        Public BlnCopyTratamientosFito As Boolean
        Public BlnCopyParcelas As Boolean
        Public BlnCopyParcelasAsesoradas As Boolean
        Public BlnCopySemillas As Boolean
        Public BlnCopyPostcosecha As Boolean
        Public BlnCopyLocales As Boolean
        Public BlnCopyTransporte As Boolean
        Public BlnCopyVentas As Boolean
        Public BlnCopyFertilizacion As Boolean
        Public IDCuadernoNew As String
        Public IDCuadernoPadre As String

        Public dtCabecera As DataTable
        Public dtCuadernoNew As DataTable
        Public dtMODNew As DataTable
        Public dtCentrosNew As DataTable
        Public dtAsesoresNew As DataTable
        Public dtFincasNew As DataTable
        Public dtZonasNew As DataTable
        Public dtAnalisisFitoNew As DataTable
        Public dtTratamientosFitoNew As DataTable
        Public dtParcelasNew As DataTable
        Public dtParcelasAsesoradasNew As DataTable
        Public dtSemillasNew As DataTable
        Public dtPostcosechaNew As DataTable
        Public dtLocalesNew As DataTable
        Public dtTransporteNew As DataTable
        Public dtVentasNew As DataTable
        Public dtFertilizacionNew As DataTable
    End Class

    <Task()> Public Shared Function CopiaCuaderno(ByVal data As DatosCuadernoCopia, ByVal services As ServiceProvider) As String
        If Length(data.IDCuaderno) > 0 Then
            ProcessServer.ExecuteTask(Of DatosCuadernoCopia)(AddressOf GeneraNuevoCuaderno, data, services)
            If Length(data.IDCuadernoNew) > 0 Then
                If data.BlnCopyInformacionGeneral And data.BlnCopyPersonas Then
                    ProcessServer.ExecuteTask(Of DatosCuadernoCopia)(AddressOf CopiarCuadernoMOD, data, services)
                End If
                If data.BlnCopyInformacionGeneral And data.BlnCopyEquipos Then
                    ProcessServer.ExecuteTask(Of DatosCuadernoCopia)(AddressOf CopiarCuadernoEquipos, data, services)
                End If
                If data.BlnCopyInformacionGeneral And data.BlnCopyAsesores Then
                    ProcessServer.ExecuteTask(Of DatosCuadernoCopia)(AddressOf CopiarCuadernoAsesores, data, services)
                End If
                If data.BlnCopyFincas Then
                    ProcessServer.ExecuteTask(Of DatosCuadernoCopia)(AddressOf CopiarCuadernoFincas, data, services)
                End If
                If data.BlnCopyZonas Then
                    ProcessServer.ExecuteTask(Of DatosCuadernoCopia)(AddressOf CopiarCuadernoZonas, data, services)
                End If
                If data.BlnCopyTratamientosFito And data.BlnCopyParcelas Then
                    ProcessServer.ExecuteTask(Of DatosCuadernoCopia)(AddressOf CopiarCuadernoParcelas, data, services)
                End If
                If data.BlnCopyTratamientosFito And data.BlnCopyParcelasAsesoradas Then
                    ProcessServer.ExecuteTask(Of DatosCuadernoCopia)(AddressOf CopiarCuadernoParcelasAsesoradas, data, services)
                End If
                If data.BlnCopyTratamientosFito And data.BlnCopySemillas Then
                    ProcessServer.ExecuteTask(Of DatosCuadernoCopia)(AddressOf CopiarCuadernoSemillas, data, services)
                End If
                If data.BlnCopyTratamientosFito And data.BlnCopyPostcosecha Then
                    ProcessServer.ExecuteTask(Of DatosCuadernoCopia)(AddressOf CopiarCuadernoPostcosecha, data, services)
                End If
                If data.BlnCopyTratamientosFito And data.BlnCopyLocales Then
                    ProcessServer.ExecuteTask(Of DatosCuadernoCopia)(AddressOf CopiarCuadernoLocales, data, services)
                End If
                If data.BlnCopyTratamientosFito And data.BlnCopyTransporte Then
                    ProcessServer.ExecuteTask(Of DatosCuadernoCopia)(AddressOf CopiarCuadernoTransporte, data, services)
                End If
                If data.BlnCopyAnalisisFito Then
                    ProcessServer.ExecuteTask(Of DatosCuadernoCopia)(AddressOf CopiarCuadernoAnalisisFitosanitarios, data, services)
                End If
                If data.BlnCopyVentas Then
                    ProcessServer.ExecuteTask(Of DatosCuadernoCopia)(AddressOf CopiarCuadernoVentas, data, services)
                End If
                If data.BlnCopyFertilizacion Then
                    ProcessServer.ExecuteTask(Of DatosCuadernoCopia)(AddressOf CopiarCuadernoFertilizacion, data, services)
                End If

                Return data.IDCuadernoNew
            Else
                Return String.Empty
            End If
        Else
            Return String.Empty
        End If
    End Function

    <Task()> Public Shared Function GeneraNuevoCuaderno(ByVal data As DatosCuadernoCopia, ByVal services As ServiceProvider) As String
        Dim Cuaderno As New BdgCuadernoCampoEstatalCabecera
        Dim dt As DataTable = Cuaderno.SelOnPrimaryKey(data.IDCuaderno)
        If dt.Rows.Count > 0 Then
            data.dtCuadernoNew = Cuaderno.AddNew()
            data.dtCuadernoNew.Rows.Add(dt.Rows(0).ItemArray)
            If Len(data.IDContador) > 0 Then
                data.dtCuadernoNew.Rows(0)("IDCuaderno") = Guid.NewGuid
                data.dtCuadernoNew.Rows(0)("NCuaderno") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, data.IDContador, services)
                data.dtCuadernoNew.Rows(0)("IDContador") = data.IDContador
            Else
                data.dtCuadernoNew.Rows(0)("IDCuaderno") = DBNull.Value
                data.dtCuadernoNew.Rows(0)("IDContador") = DBNull.Value
            End If
        End If
        Cuaderno.Update(data.dtCuadernoNew)
        data.IDCuadernoNew = CType(data.dtCuadernoNew.Rows(0)("IDCuaderno"), Guid).ToString
        Return data.IDCuadernoNew

    End Function

    <Task()> Public Shared Sub CopiarCuadernoMOD(ByVal data As DatosCuadernoCopia, ByVal services As ServiceProvider)
        Dim A As New BdgCuadernoCampoEstatalMOD
        Dim dt As DataTable = A.Filter(New StringFilterItem("IDCuaderno", data.IDCuaderno))
        If dt.Rows.Count > 0 Then
            data.dtMODNew = A.AddNew()
            For Each dr As DataRow In dt.Select
                Dim drNew As DataRow = data.dtMODNew.NewRow
                drNew.ItemArray = dr.ItemArray
                drNew("IDCuaderno") = data.IDCuadernoNew

                data.dtMODNew.Rows.Add(drNew)
            Next
        End If
        A.Update(data.dtMODNew)
    End Sub

    <Task()> Public Shared Sub CopiarCuadernoEquipos(ByVal data As DatosCuadernoCopia, ByVal services As ServiceProvider)
        Dim A As New BdgCuadernoCampoEstatalCentro
        Dim dt As DataTable = A.Filter(New StringFilterItem("IDCuaderno", data.IDCuaderno))
        If dt.Rows.Count > 0 Then
            data.dtCentrosNew = A.AddNew()
            For Each dr As DataRow In dt.Select
                Dim drNew As DataRow = data.dtCentrosNew.NewRow
                drNew.ItemArray = dr.ItemArray
                drNew("IDCuaderno") = data.IDCuadernoNew

                data.dtCentrosNew.Rows.Add(drNew)
            Next
        End If
        A.Update(data.dtCentrosNew)
    End Sub

    <Task()> Public Shared Sub CopiarCuadernoAsesores(ByVal data As DatosCuadernoCopia, ByVal services As ServiceProvider)
        Dim A As New BdgCuadernoCampoEstatalAsesor
        Dim dt As DataTable = A.Filter(New StringFilterItem("IDCuaderno", data.IDCuaderno))
        If dt.Rows.Count > 0 Then
            data.dtAsesoresNew = A.AddNew()
            For Each dr As DataRow In dt.Select
                Dim drNew As DataRow = data.dtAsesoresNew.NewRow
                drNew.ItemArray = dr.ItemArray
                drNew("IDCuaderno") = data.IDCuadernoNew

                data.dtAsesoresNew.Rows.Add(drNew)
            Next
        End If
        A.Update(data.dtAsesoresNew)
    End Sub

    <Task()> Public Shared Sub CopiarCuadernoFincas(ByVal data As DatosCuadernoCopia, ByVal services As ServiceProvider)
        Dim A As New BdgCuadernoCampoEstatalParcelasAgro
        Dim dt As DataTable = A.Filter(New StringFilterItem("IDCuaderno", data.IDCuaderno))
        If dt.Rows.Count > 0 Then
            data.dtFincasNew = A.AddNew()
            For Each dr As DataRow In dt.Select
                Dim drNew As DataRow = data.dtFincasNew.NewRow
                drNew.ItemArray = dr.ItemArray
                drNew("IDCuaderno") = data.IDCuadernoNew

                data.dtFincasNew.Rows.Add(drNew)
            Next
        End If
        A.Update(data.dtFincasNew)
    End Sub

    <Task()> Public Shared Sub CopiarCuadernoZonas(ByVal data As DatosCuadernoCopia, ByVal services As ServiceProvider)
        Dim A As New BdgCuadernoCampoEstatalParcelasMedioAmb
        Dim dt As DataTable = A.Filter(New StringFilterItem("IDCuaderno", data.IDCuaderno))
        If dt.Rows.Count > 0 Then
            data.dtZonasNew = A.AddNew()
            For Each dr As DataRow In dt.Select
                Dim drNew As DataRow = data.dtZonasNew.NewRow
                drNew.ItemArray = dr.ItemArray
                drNew("IDCuaderno") = data.IDCuadernoNew

                data.dtZonasNew.Rows.Add(drNew)
            Next
        End If
        A.Update(data.dtZonasNew)
    End Sub

    <Task()> Public Shared Sub CopiarCuadernoParcelas(ByVal data As DatosCuadernoCopia, ByVal services As ServiceProvider)
        Dim A As New BdgCuadernoCampoEstatalFitosanitario
        Dim filtro As New Filter
        filtro.Add("IDCuaderno", FilterOperator.Equal, data.IDCuaderno, FilterType.String)
        filtro.Add("IDTipoRegistro", FilterOperator.Equal, 0, FilterType.String)
        Dim dt As DataTable = A.Filter(filtro)
        If dt.Rows.Count > 0 Then
            data.dtParcelasNew = A.AddNew()
            For Each dr As DataRow In dt.Select
                Dim drNew As DataRow = data.dtParcelasNew.NewRow
                drNew.ItemArray = dr.ItemArray
                drNew("IDCuaderno") = data.IDCuadernoNew

                data.dtParcelasNew.Rows.Add(drNew)
            Next
        End If
        A.Update(data.dtParcelasNew)
    End Sub

    <Task()> Public Shared Sub CopiarCuadernoParcelasAsesoradas(ByVal data As DatosCuadernoCopia, ByVal services As ServiceProvider)
        Dim A As New BdgCuadernoCampoEstatalFitosanitario
        Dim filtro As New Filter
        filtro.Add("IDCuaderno", FilterOperator.Equal, data.IDCuaderno, FilterType.String)
        filtro.Add("IDTipoRegistro", FilterOperator.Equal, 1, FilterType.String)
        Dim dt As DataTable = A.Filter(filtro)
        If dt.Rows.Count > 0 Then
            data.dtParcelasAsesoradasNew = A.AddNew()
            For Each dr As DataRow In dt.Select
                Dim drNew As DataRow = data.dtParcelasAsesoradasNew.NewRow
                drNew.ItemArray = dr.ItemArray
                drNew("IDCuaderno") = data.IDCuadernoNew

                data.dtParcelasAsesoradasNew.Rows.Add(drNew)
            Next
        End If
        A.Update(data.dtParcelasAsesoradasNew)
    End Sub

    <Task()> Public Shared Sub CopiarCuadernoSemillas(ByVal data As DatosCuadernoCopia, ByVal services As ServiceProvider)
        Dim A As New BdgCuadernoCampoEstatalFitosanitario
        Dim filtro As New Filter
        filtro.Add("IDCuaderno", FilterOperator.Equal, data.IDCuaderno, FilterType.String)
        filtro.Add("IDTipoRegistro", FilterOperator.Equal, 2, FilterType.String)
        Dim dt As DataTable = A.Filter(filtro)
        If dt.Rows.Count > 0 Then
            data.dtSemillasNew = A.AddNew()
            For Each dr As DataRow In dt.Select
                Dim drNew As DataRow = data.dtSemillasNew.NewRow
                drNew.ItemArray = dr.ItemArray
                drNew("IDCuaderno") = data.IDCuadernoNew

                data.dtSemillasNew.Rows.Add(drNew)
            Next
        End If
        A.Update(data.dtSemillasNew)
    End Sub

    <Task()> Public Shared Sub CopiarCuadernoPostcosecha(ByVal data As DatosCuadernoCopia, ByVal services As ServiceProvider)
        Dim A As New BdgCuadernoCampoEstatalFitosanitario
        Dim filtro As New Filter
        filtro.Add("IDCuaderno", FilterOperator.Equal, data.IDCuaderno, FilterType.String)
        filtro.Add("IDTipoRegistro", FilterOperator.Equal, 3, FilterType.String)
        Dim dt As DataTable = A.Filter(filtro)
        If dt.Rows.Count > 0 Then
            data.dtPostcosechaNew = A.AddNew()
            For Each dr As DataRow In dt.Select
                Dim drNew As DataRow = data.dtPostcosechaNew.NewRow
                drNew.ItemArray = dr.ItemArray
                drNew("IDCuaderno") = data.IDCuadernoNew

                data.dtPostcosechaNew.Rows.Add(drNew)
            Next
        End If
        A.Update(data.dtPostcosechaNew)
    End Sub

    <Task()> Public Shared Sub CopiarCuadernoLocales(ByVal data As DatosCuadernoCopia, ByVal services As ServiceProvider)
        Dim A As New BdgCuadernoCampoEstatalFitosanitario
        Dim filtro As New Filter
        filtro.Add("IDCuaderno", FilterOperator.Equal, data.IDCuaderno, FilterType.String)
        filtro.Add("IDTipoRegistro", FilterOperator.Equal, 4, FilterType.String)
        Dim dt As DataTable = A.Filter(filtro)
        If dt.Rows.Count > 0 Then
            data.dtLocalesNew = A.AddNew()
            For Each dr As DataRow In dt.Select
                Dim drNew As DataRow = data.dtLocalesNew.NewRow
                drNew.ItemArray = dr.ItemArray
                drNew("IDCuaderno") = data.IDCuadernoNew

                data.dtLocalesNew.Rows.Add(drNew)
            Next
        End If
        A.Update(data.dtLocalesNew)
    End Sub

    <Task()> Public Shared Sub CopiarCuadernoTransporte(ByVal data As DatosCuadernoCopia, ByVal services As ServiceProvider)
        Dim A As New BdgCuadernoCampoEstatalFitosanitario
        Dim filtro As New Filter
        filtro.Add("IDCuaderno", FilterOperator.Equal, data.IDCuaderno, FilterType.String)
        filtro.Add("IDTipoRegistro", FilterOperator.Equal, 5, FilterType.String)
        Dim dt As DataTable = A.Filter(filtro)
        If dt.Rows.Count > 0 Then
            data.dtTransporteNew = A.AddNew()
            For Each dr As DataRow In dt.Select
                Dim drNew As DataRow = data.dtTransporteNew.NewRow
                drNew.ItemArray = dr.ItemArray
                drNew("IDCuaderno") = data.IDCuadernoNew

                data.dtTransporteNew.Rows.Add(drNew)
            Next
        End If
        A.Update(data.dtTransporteNew)
    End Sub

    <Task()> Public Shared Sub CopiarCuadernoAnalisisFitosanitarios(ByVal data As DatosCuadernoCopia, ByVal services As ServiceProvider)
        Dim A As New BdgCuadernoCampoEstatalAnalisisFitosanitario
        Dim dt As DataTable = A.Filter(New StringFilterItem("IDCuaderno", data.IDCuaderno))
        If dt.Rows.Count > 0 Then
            data.dtAnalisisFitoNew = A.AddNew()
            For Each dr As DataRow In dt.Select
                Dim drNew As DataRow = data.dtAnalisisFitoNew.NewRow
                drNew.ItemArray = dr.ItemArray
                drNew("IDCuaderno") = data.IDCuadernoNew

                data.dtAnalisisFitoNew.Rows.Add(drNew)
            Next
        End If
        A.Update(data.dtAnalisisFitoNew)
    End Sub

    <Task()> Public Shared Sub CopiarCuadernoVentas(ByVal data As DatosCuadernoCopia, ByVal services As ServiceProvider)
        Dim A As New BdgCuadernoCampoEstatalVenta
        Dim dt As DataTable = A.Filter(New StringFilterItem("IDCuaderno", data.IDCuaderno))
        If dt.Rows.Count > 0 Then
            data.dtVentasNew = A.AddNew()
            For Each dr As DataRow In dt.Select
                Dim drNew As DataRow = data.dtVentasNew.NewRow
                drNew.ItemArray = dr.ItemArray
                drNew("IDCuaderno") = data.IDCuadernoNew

                data.dtVentasNew.Rows.Add(drNew)
            Next
        End If
        A.Update(data.dtVentasNew)
    End Sub

    <Task()> Public Shared Sub CopiarCuadernoFertilizacion(ByVal data As DatosCuadernoCopia, ByVal services As ServiceProvider)
        Dim A As New BdgCuadernoCampoEstatalFertilizacion
        Dim dt As DataTable = A.Filter(New StringFilterItem("IDCuaderno", data.IDCuaderno))
        If dt.Rows.Count > 0 Then
            data.dtFertilizacionNew = A.AddNew()
            For Each dr As DataRow In dt.Select
                Dim drNew As DataRow = data.dtFertilizacionNew.NewRow
                drNew.ItemArray = dr.ItemArray
                drNew("IDCuaderno") = data.IDCuadernoNew

                data.dtFertilizacionNew.Rows.Add(drNew)
            Next
        End If
        A.Update(data.dtFertilizacionNew)
    End Sub
#End Region

#Region "Volcado Cuaderno"
    <Serializable()> _
      Public Class VolcarDatosCuaderno
        Public IDCuaderno As Guid
        Public blnInsertar As Boolean
        Public blnReemplazar As Boolean

        Public blnFincas As Boolean
        Public blnOperarios As Boolean
        Public blnCentros As Boolean

        Public blnFitoParcela As Boolean
        Public blnFitoParcelaAsesorada As Boolean
        Public blnFitoSemilla As Boolean
        Public blnFitoLocales As Boolean
        Public blnFitoPostcosecha As Boolean
        Public blnFitoTransporte As Boolean

        Public blnAnalisis As Boolean
        Public blnVentas As Boolean
        Public blnFertilizacion As Boolean

        Public dtFincasNew As DataTable
        Public dtFincasSelect As DataTable
        Public dtZonasNew As DataTable

        Public dtOperariosNew As DataTable
        Public lstOperariosSelect As List(Of String)

        Public dtCentrosNew As DataTable
        Public lstCentrosSelect As List(Of String)

        Public dtMaterialesNew As DataTable
        Public dtMaterialesSelect As DataTable

        Public dtAnalisisNew As DataTable
        Public dtAnalisisSelect As DataTable

        Public dtVentasNew As DataTable
        Public dtVentasSelect As DataTable
    End Class

    <Task()> Public Shared Function VolcadoDatosCuaderno(ByVal data As VolcarDatosCuaderno, ByVal services As ServiceProvider) As Guid

        If Length(data.IDCuaderno) > 0 Then
            'Fincas
            If data.blnFincas Then
                ProcessServer.ExecuteTask(Of VolcarDatosCuaderno)(AddressOf TratarCuadernoFincas, data, services)
            End If
            'Operarios
            If data.blnOperarios Then
                ProcessServer.ExecuteTask(Of VolcarDatosCuaderno)(AddressOf TratarCuadernoOperarios, data, services)
            End If
            'Centros
            If data.blnCentros Then
                ProcessServer.ExecuteTask(Of VolcarDatosCuaderno)(AddressOf TratarCuadernoCentros, data, services)
            End If
            'Materiales
            If data.blnFitoParcela OrElse data.blnFitoParcelaAsesorada OrElse data.blnFitoSemilla OrElse data.blnFitoPostcosecha Then
                ProcessServer.ExecuteTask(Of VolcarDatosCuaderno)(AddressOf TratarCuadernoFitoParcela, data, services)
            End If
            'Analisis
            If data.blnAnalisis Then
                ProcessServer.ExecuteTask(Of VolcarDatosCuaderno)(AddressOf TratarCuadernoAnalisisFitosanitario, data, services)
            End If
            'Ventas
            If data.blnVentas Then
                ProcessServer.ExecuteTask(Of VolcarDatosCuaderno)(AddressOf TratarCuadernoVentas, data, services)
            End If
            'Fertilizacion
            If data.blnFertilizacion Then
                ProcessServer.ExecuteTask(Of VolcarDatosCuaderno)(AddressOf TratarCuadernoFertilizacion, data, services)
            End If

            Return data.IDCuaderno
        Else
            Return Guid.Empty
        End If

    End Function

    <Task()> Public Shared Sub TratarCuadernoFincas(ByVal data As VolcarDatosCuaderno, ByVal services As ServiceProvider)
        Dim A As New BdgCuadernoCampoEstatalParcelasAgro
        Dim Z As New BdgCuadernoCampoEstatalParcelasMedioAmb
        Dim dtFincaCuaderno As DataTable = A.Filter(New StringFilterItem("IDCuaderno", data.IDCuaderno))
        Dim dtZonaCuaderno As DataTable = Z.Filter(New StringFilterItem("IDCuaderno", data.IDCuaderno))

        If data.dtFincasSelect.Rows.Count > 0 Then
            data.dtFincasNew = A.AddNew()
            data.dtZonasNew = Z.AddNew()
            Dim ExisteFinca As List(Of DataRow)
            For Each dr As DataRow In data.dtFincasSelect.Select
                ExisteFinca = Nothing
                If Not data.blnReemplazar Then
                    ExisteFinca = (From op In dtFincaCuaderno _
                                        Where Not op.IsNull("IDCuaderno") AndAlso op("IDCuaderno").Equals(data.IDCuaderno) AndAlso _
                                             Not op.IsNull("IDFinca") AndAlso op("IDFinca") = dr("IDFinca") AndAlso _
                                             Not op.IsNull("IDFincaSigpac") AndAlso op("IDFincaSigpac") = dr("IDFincaSigPac") _
                                        Select op).ToList
                End If
                If ExisteFinca Is Nothing OrElse ExisteFinca.Count = 0 Then
                    ExisteFinca = (From op In data.dtFincasNew _
                                        Where Not op.IsNull("IDCuaderno") AndAlso op("IDCuaderno").Equals(data.IDCuaderno) AndAlso _
                                             Not op.IsNull("IDFinca") AndAlso op("IDFinca") = dr("IDFinca") AndAlso _
                                            Not op.IsNull("IDFincaSigpac") AndAlso op("IDFincaSigpac") = dr("IDFincaSigPac") _
                                        Select op).ToList
                    If ExisteFinca Is Nothing OrElse ExisteFinca.Count = 0 Then
                        'Agregar Finca
                        Dim drNew As DataRow = data.dtFincasNew.NewRow
                        drNew("IDCuadernoParcelasAgro") = Guid.NewGuid
                        drNew("IDCuaderno") = data.IDCuaderno
                        drNew("IDFinca") = dr("IDFinca")
                        A.ApplyBusinessRule("IDFinca", drNew("IDFinca"), drNew)
                        drNew("IDFincaSigpac") = dr("IDFincaSigPac")
                        A.ApplyBusinessRule("IDFincaSigpac", drNew("IDFincaSigpac"), drNew)
                        data.dtFincasNew.Rows.Add(drNew)

                        'Agregar Zonas medioambientales 2.2
                        Dim ClsZona As New BdgFincaZonaEspecifica
                        Dim dtZona As DataTable = ClsZona.Filter(New StringFilterItem("IDFinca", FilterOperator.Equal, dr("IDFinca")))
                        Dim ExisteZona As List(Of DataRow)
                        For Each drZona As DataRow In dtZona.Select
                            ExisteZona = Nothing
                            If Not data.blnReemplazar Then
                                ExisteZona = (From op In dtZonaCuaderno _
                                                    Where Not op.IsNull("IDCuaderno") AndAlso op("IDCuaderno").Equals(data.IDCuaderno) AndAlso _
                                                         Not op.IsNull("IDFinca") AndAlso op("IDFinca") = dr("IDFinca") AndAlso _
                                                         Not op.IsNull("IDZonaEspecifica") AndAlso op("IDZonaEspecifica") = drZona("IDZonaEspecifica") _
                                                    Select op).ToList
                            End If
                            If ExisteZona Is Nothing OrElse ExisteZona.Count = 0 Then
                                ExisteZona = (From op In data.dtZonasNew _
                                                    Where Not op.IsNull("IDCuaderno") AndAlso op("IDCuaderno").Equals(data.IDCuaderno) AndAlso _
                                                         Not op.IsNull("IDFinca") AndAlso op("IDFinca") = dr("IDFinca") AndAlso _
                                                         Not op.IsNull("IDZonaEspecifica") AndAlso op("IDZonaEspecifica") = drZona("IDZonaEspecifica") _
                                                    Select op).ToList
                                If ExisteZona Is Nothing OrElse ExisteZona.Count = 0 Then
                                    'Agregar Finca
                                    Dim drNewZona As DataRow = data.dtZonasNew.NewRow
                                    drNewZona("IDCuadernoParcelasMedioAmb") = Guid.NewGuid
                                    drNewZona("IDCuaderno") = data.IDCuaderno
                                    drNewZona("IDFinca") = dr("IDFinca")
                                    drNewZona("IDZonaEspecifica") = drZona("IDZonaEspecifica")
                                    Z.ApplyBusinessRule("IDFinca", drNewZona("IDFinca"), drNewZona)
                                    Z.ApplyBusinessRule("IDZonaEspecifica", drNewZona("IDZonaEspecifica"), drNewZona)
                                    data.dtZonasNew.Rows.Add(drNewZona)
                                End If
                            End If
                        Next
                    End If
                End If
            Next
        End If
        If data.blnReemplazar Then
            Z.Delete(dtZonaCuaderno)
            A.Delete(dtFincaCuaderno)
        End If
        A.Update(data.dtFincasNew)
        Z.Update(data.dtZonasNew)
    End Sub

    <Task()> Public Shared Sub TratarCuadernoOperarios(ByVal data As VolcarDatosCuaderno, ByVal services As ServiceProvider)
        Dim OperariosAsesores As List(Of String) = ProcessServer.ExecuteTask(Of List(Of String), List(Of String))(AddressOf GetOperariosAsesores, data.lstOperariosSelect, services)
        Dim OperariosNoAsesores As List(Of String) = (From opInsert In data.lstOperariosSelect Select opInsert).Except(OperariosAsesores).ToList

        'Asesores
        If Not OperariosAsesores Is Nothing AndAlso OperariosAsesores.Count > 0 Then
            Dim Ase As New BdgCuadernoCampoEstatalAsesor
            Dim dt As DataTable = Ase.Filter(New StringFilterItem("IDCuaderno", data.IDCuaderno))
            data.dtOperariosNew = Ase.AddNew()
            Dim ExisteOperario As List(Of DataRow)
            For Each strIDOperario As String In OperariosAsesores
                ExisteOperario = Nothing
                If Not data.blnReemplazar Then
                    ExisteOperario = (From op In dt _
                                        Where Not op.IsNull("IDCuaderno") AndAlso op("IDCuaderno").Equals(data.IDCuaderno) AndAlso _
                                                Not op.IsNull("IDAsesor") AndAlso op("IDAsesor") = strIDOperario _
                                        Select op).ToList
                End If
                If ExisteOperario Is Nothing OrElse ExisteOperario.Count = 0 Then
                    ExisteOperario = (From op In data.dtOperariosNew _
                                        Where Not op.IsNull("IDCuaderno") AndAlso op("IDCuaderno").Equals(data.IDCuaderno) AndAlso _
                                                Not op.IsNull("IDAsesor") AndAlso op("IDAsesor") = strIDOperario _
                                        Select op).ToList
                    If ExisteOperario Is Nothing OrElse ExisteOperario.Count = 0 Then
                        Dim drNew As DataRow = data.dtOperariosNew.NewRow
                        drNew("IDCuadernoAsesor") = Guid.NewGuid
                        drNew("IDCuaderno") = data.IDCuaderno
                        drNew("IDAsesor") = strIDOperario
                        drNew = Ase.ApplyBusinessRule("IDAsesor", drNew("IDAsesor"), drNew)
                        data.dtOperariosNew.Rows.Add(drNew)
                    End If
                End If
            Next
            If data.blnReemplazar Then
                Ase.Delete(dt)
            End If
            Ase.Update(data.dtOperariosNew)
        End If

        'Resto Operarios
        If Not OperariosNoAsesores Is Nothing AndAlso OperariosNoAsesores.Count > 0 Then
            Dim A As New BdgCuadernoCampoEstatalMOD
            Dim dt As DataTable = A.Filter(New StringFilterItem("IDCuaderno", data.IDCuaderno))
            data.dtOperariosNew = A.AddNew()
            Dim ExisteOperario As List(Of DataRow)
            For Each strIDOperario As String In OperariosNoAsesores
                ExisteOperario = Nothing
                If Not data.blnReemplazar Then
                    ExisteOperario = (From op In dt _
                                        Where Not op.IsNull("IDCuaderno") AndAlso op("IDCuaderno").Equals(data.IDCuaderno) AndAlso _
                                                Not op.IsNull("IDOperario") AndAlso op("IDOperario") = strIDOperario _
                                        Select op).ToList
                End If
                If ExisteOperario Is Nothing OrElse ExisteOperario.Count = 0 Then
                    ExisteOperario = (From op In data.dtOperariosNew _
                                        Where Not op.IsNull("IDCuaderno") AndAlso op("IDCuaderno").Equals(data.IDCuaderno) AndAlso _
                                                Not op.IsNull("IDOperario") AndAlso op("IDOperario") = strIDOperario _
                                        Select op).ToList
                    If ExisteOperario Is Nothing OrElse ExisteOperario.Count = 0 Then
                        Dim drNew As DataRow = data.dtOperariosNew.NewRow
                        drNew("IDCuadernoMOD") = Guid.NewGuid
                        drNew("IDCuaderno") = data.IDCuaderno
                        drNew("IDOperario") = strIDOperario
                        drNew = A.ApplyBusinessRule("IDOperario", drNew("IDOperario"), drNew)
                        data.dtOperariosNew.Rows.Add(drNew)
                    End If
                End If
            Next

            If data.blnReemplazar Then
                A.Delete(dt)
            End If
            A.Update(data.dtOperariosNew)
        End If

    End Sub

    <Task()> Public Shared Function GetOperariosAsesores(ByVal lstOperarios As List(Of String), ByVal services As ServiceProvider) As List(Of String)
        Dim f As New Filter
        f.Add(New BooleanFilterItem("Asesor", True))
        If Not lstOperarios Is Nothing AndAlso lstOperarios.Count > 0 Then
            f.Add(New InListFilterItem("IDOperario", lstOperarios.ToArray, FilterType.String))
        End If
        Return (From c In New Operario().Filter(f, , "IDOperario") Select CStr(c("IDOperario")) Distinct).ToList
    End Function

    <Task()> Public Shared Sub TratarCuadernoCentros(ByVal data As VolcarDatosCuaderno, ByVal services As ServiceProvider)
        Dim A As New BdgCuadernoCampoEstatalCentro
        Dim dt As DataTable = A.Filter(New StringFilterItem("IDCuaderno", data.IDCuaderno))
        If data.lstCentrosSelect.Count > 0 Then
            data.dtCentrosNew = A.AddNew()
            Dim ExisteCentro As List(Of DataRow)
            For Each strIDCentro As String In data.lstCentrosSelect
                ExisteCentro = Nothing
                If Not data.blnReemplazar Then
                    ExisteCentro = (From op In dt _
                                         Where Not op.IsNull("IDCuaderno") AndAlso op("IDCuaderno").Equals(data.IDCuaderno) AndAlso _
                                               Not op.IsNull("IDCentro") AndAlso op("IDCentro") = strIDCentro _
                                         Select op).ToList
                End If
                If ExisteCentro Is Nothing OrElse ExisteCentro.Count = 0 Then
                    ExisteCentro = (From op In data.dtCentrosNew _
                                                      Where Not op.IsNull("IDCuaderno") AndAlso op("IDCuaderno").Equals(data.IDCuaderno) AndAlso _
                                                            Not op.IsNull("IDCentro") AndAlso op("IDCentro") = strIDCentro _
                                                      Select op).ToList
                    If ExisteCentro Is Nothing OrElse ExisteCentro.Count = 0 Then
                        Dim drNew As DataRow = data.dtCentrosNew.NewRow
                        drNew("IDCuadernoCentro") = Guid.NewGuid
                        drNew("IDCuaderno") = data.IDCuaderno
                        drNew("IDCentro") = strIDCentro
                        drNew = A.ApplyBusinessRule("IDCentro", drNew("IDCentro"), drNew)

                        data.dtCentrosNew.Rows.Add(drNew)
                    End If
                End If
            Next
        End If
        If data.blnReemplazar Then
            A.Delete(dt)
        End If
        A.Update(data.dtCentrosNew)
    End Sub

    <Task()> Public Shared Sub TratarCuadernoFitoParcela(ByVal data As VolcarDatosCuaderno, ByVal services As ServiceProvider)
        Dim A As New BdgCuadernoCampoEstatalFitosanitario
        If data.blnReemplazar Then
            Dim f As New Filter
            f.Add(New StringFilterItem("IDCuaderno", data.IDCuaderno))
            If data.blnFitoParcela Then
                f.Add(New NumberFilterItem("IDTipoRegistro", BdgTipoFitosanitario.Parcela))
            End If
            If data.blnFitoParcelaAsesorada Then
                f.Add(New NumberFilterItem("IDTipoRegistro", BdgTipoFitosanitario.ParcelaAsesorada))
            End If
            If data.blnFitoSemilla Then
                f.Add(New NumberFilterItem("IDTipoRegistro", BdgTipoFitosanitario.Semilla))
            End If
            If data.blnFitoPostcosecha Then
                f.Add(New NumberFilterItem("IDTipoRegistro", BdgTipoFitosanitario.Postcosecha))
            End If
            Dim dt As DataTable = A.Filter(f)
            A.Delete(dt)
        End If
        If data.dtMaterialesSelect.Rows.Count > 0 Then
            data.dtMaterialesNew = A.AddNew()
            For Each dr As DataRow In data.dtMaterialesSelect.Select
                Dim drNew As DataRow = data.dtMaterialesNew.NewRow
                drNew("IDCuadernoFitosanitario") = Guid.NewGuid
                drNew("IDCuaderno") = data.IDCuaderno
                If data.blnFitoParcela Then
                    drNew("IDTipoRegistro") = BdgTipoFitosanitario.Parcela
                End If
                If data.blnFitoParcelaAsesorada Then
                    drNew("IDTipoRegistro") = BdgTipoFitosanitario.ParcelaAsesorada
                End If
                If data.blnFitoSemilla Then
                    drNew("IDTipoRegistro") = BdgTipoFitosanitario.Semilla
                End If
                If data.blnFitoPostcosecha Then
                    drNew("IDTipoRegistro") = BdgTipoFitosanitario.Postcosecha
                End If

                drNew("IDFinca") = dr("IDFinca")
                A.ApplyBusinessRule("IDFinca", drNew("IDFinca"), drNew)

                drNew("CantidadTratada") = Nz(dr("CantidadTratada"), 0)
                A.ApplyBusinessRule("CantidadTratada", drNew("CantidadTratada"), drNew)
                drNew("Fecha") = dr("Fecha")

                drNew("IDOperario") = dr("IDOperario")
                drNew("IDCentro") = dr("IDCentro")
                drNew("IDArticulo") = dr("IDMaterial")
                A.ApplyBusinessRule("IDArticulo", drNew("IDArticulo"), drNew)
                drNew("CantidadArticulo") = dr("QReal")
                A.ApplyBusinessRule("CantidadArticulo", drNew("CantidadArticulo"), drNew)
                '[IDProblematicaFitosanitaria]
                '[JustificacionProblematicaFitosanitaria]
                '[IDMedidaIntervencion]
                '[IntensidadMedidaIntervencion]
                '[IDEficacia]
                '[Texto]

                drNew("IDLineaMaterialControl") = dr("IDLineaMaterialControl")
                drNew("IDParteAgrupado") = dr("IDParteAgrupado")

                data.dtMaterialesNew.Rows.Add(drNew)
            Next
        End If
        A.Update(data.dtMaterialesNew)
    End Sub

    <Task()> Public Shared Sub TratarCuadernoAnalisisFitosanitario(ByVal data As VolcarDatosCuaderno, ByVal services As ServiceProvider)
        Dim A As New BdgCuadernoCampoEstatalAnalisisFitosanitario
        Dim dt As DataTable = A.Filter(New StringFilterItem("IDCuaderno", data.IDCuaderno))
        If data.dtAnalisisSelect.Rows.Count > 0 Then
            data.dtAnalisisNew = A.AddNew()
            Dim ExisteAnalisis As List(Of DataRow)
            For Each dr As DataRow In data.dtAnalisisSelect.Select
                ExisteAnalisis = Nothing
                If Not data.blnReemplazar Then
                    ExisteAnalisis = (From op In dt _
                                         Where Not op.IsNull("IDCuaderno") AndAlso op("IDCuaderno").Equals(data.IDCuaderno) AndAlso _
                                                Not op.IsNull("IDFinca") AndAlso op("IDFinca") = dr("IDFinca") AndAlso _
                                                Not op.IsNull("IDAnalisisCabecera") AndAlso op("IDAnalisisCabecera") = dr("IDAnalisisCabecera") _
                                         Select op).ToList
                End If
                If ExisteAnalisis Is Nothing OrElse ExisteAnalisis.Count = 0 Then
                    ExisteAnalisis = (From op In data.dtAnalisisNew _
                                          Where Not op.IsNull("IDCuaderno") AndAlso op("IDCuaderno").Equals(data.IDCuaderno) AndAlso _
                                                Not op.IsNull("IDFinca") AndAlso op("IDFinca") = dr("IDFinca") AndAlso _
                                                Not op.IsNull("IDAnalisisCabecera") AndAlso op("IDAnalisisCabecera") = dr("IDAnalisisCabecera") _
                                                      Select op).ToList
                    If ExisteAnalisis Is Nothing OrElse ExisteAnalisis.Count = 0 Then
                        Dim drNew As DataRow = data.dtAnalisisNew.NewRow
                        drNew("IDCuadernoAnalisisFitosanitario") = Guid.NewGuid
                        drNew("IDCuaderno") = data.IDCuaderno
                        drNew("IDAnalisisCabecera") = dr("IDAnalisisCabecera")
                        A.ApplyBusinessRule("IDAnalisisCabecera", drNew("IDAnalisisCabecera"), drNew)
                        'La finca se coge a través de IDAnalisisCabecera, pero también puede venir de un Observatorio,
                        'por eso ponemos la que viene de la pantalla de generación del cuaderno.
                        drNew("IDFinca") = dr("IDFinca")
                        A.ApplyBusinessRule("IDFinca", drNew("IDFinca"), drNew)
                        drNew("IDParteAgrupado") = dr("IDParteAgrupado")
                        data.dtAnalisisNew.Rows.Add(drNew)
                    End If
                End If
            Next
        End If
        If data.blnReemplazar Then
            A.Delete(dt)
        End If
        A.Update(data.dtAnalisisNew)
    End Sub

    <Task()> Public Shared Sub TratarCuadernoVentas(ByVal data As VolcarDatosCuaderno, ByVal services As ServiceProvider)
        Dim A As New BdgCuadernoCampoEstatalVenta
        If data.blnReemplazar Then
            Dim dt As DataTable = A.Filter(New StringFilterItem("IDCuaderno", data.IDCuaderno))
            A.Delete(dt)
        End If
        If data.dtVentasSelect.Rows.Count > 0 Then
            data.dtVentasNew = A.AddNew()
            For Each dr As DataRow In data.dtVentasSelect.Select
                Dim drNew As DataRow = data.dtVentasNew.NewRow
                drNew("IDCuadernoVenta") = Guid.NewGuid
                drNew("IDCuaderno") = data.IDCuaderno
                drNew("IDLineaAlbaran") = dr("IDLineaAlbaran")
                A.ApplyBusinessRule("IDLineaAlbaran", drNew("IDLineaAlbaran"), drNew)
                data.dtVentasNew.Rows.Add(drNew)
            Next
        End If
        A.Update(data.dtVentasNew)
    End Sub

    <Task()> Public Shared Sub TratarCuadernoFertilizacion(ByVal data As VolcarDatosCuaderno, ByVal services As ServiceProvider)
        Dim A As New BdgCuadernoCampoEstatalFertilizacion
        If data.blnReemplazar Then
            Dim dt As DataTable = A.Filter(New StringFilterItem("IDCuaderno", data.IDCuaderno))
            A.Delete(dt)
        End If
        If data.dtMaterialesSelect.Rows.Count > 0 Then
            data.dtMaterialesNew = A.AddNew()
            For Each dr As DataRow In data.dtMaterialesSelect.Select
                Dim drNew As DataRow = data.dtMaterialesNew.NewRow
                drNew("IDCuadernoFertilizacion") = Guid.NewGuid
                drNew("IDCuaderno") = data.IDCuaderno
                drNew("IDFinca") = dr("IDFinca")
                A.ApplyBusinessRule("IDFinca", drNew("IDFinca"), drNew)
                drNew("Fecha") = dr("Fecha")
                drNew("IDArticulo") = dr("IDMaterial")
                A.ApplyBusinessRule("IDArticulo", drNew("IDArticulo"), drNew)
                drNew("CantidadArticulo") = dr("QReal")
                drNew("IDLineaMaterialControl") = dr("IDLineaMaterialControl")
                drNew("IDParteAgrupado") = dr("IDParteAgrupado")
                data.dtMaterialesNew.Rows.Add(drNew)
            Next
        End If
        A.Update(data.dtMaterialesNew)
    End Sub
#End Region

    <Task()> Public Shared Function EsCuadernoEditable(ByVal IDCuaderno As Guid, ByVal services As ServiceProvider) As Boolean
        Dim blnRdo As Boolean = False
        If Length(IDCuaderno) > 0 Then
            Dim Cuaderno As New BdgCuadernoCampoEstatalCabecera
            Dim dtCuadernoCampo As DataTable = Cuaderno.SelOnPrimaryKey(IDCuaderno)
            If Not dtCuadernoCampo Is Nothing AndAlso dtCuadernoCampo.Rows.Count > 0 Then
                If Length(dtCuadernoCampo.Rows(0)("IDCuadernoEstado")) > 0 Then
                    Dim f As New Filter
                    f.Add("IDDeclaracionEstado", FilterOperator.Equal, dtCuadernoCampo.Rows(0)("IDCuadernoEstado"), FilterType.String)
                    Dim dtEstado As DataTable = New BE.DataEngine().Filter("tbBdgMaestroDeclaracionEstado", f, , "SoloLectura")
                    If Not dtEstado Is Nothing AndAlso dtEstado.Rows.Count > 0 Then
                        blnRdo = Not dtEstado.Rows(0)("SoloLectura")
                    End If
                End If
            End If
        End If
        Return blnRdo
    End Function

End Class