Public Class BdgCuadernoCampoEstatalParcelasMedioAmb

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgCuadernoCampoEstatalParcelasMedioAmb"

#End Region

#Region "Eventos Entidad"
    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub
  
    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClavePrimaria)
    End Sub

    Public Overrides Function GetBusinessRules() As Solmicro.Expertis.Engine.BE.BusinessRules
        Dim oBrl As New BusinessRules
        oBrl.Add("CFinca", AddressOf CambioCFinca)
        oBrl.Add("IDVariedad", AddressOf CambioVariedad)
        oBrl.Add("IDFinca", AddressOf CambioIDFinca)
        oBrl.Add("NZonaEspecifica", AddressOf CambioNZonaEspecifica)
        oBrl.Add("IDZonaEspecifica", AddressOf CambioIDZonaEspecifica)
        oBrl.Add("IDEspecie", AddressOf CambioEspecie)
        Return oBrl
    End Function

#End Region

#Region "Funciones Públicas"
    <Task()> Public Shared Sub AsignarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then data("IDCuadernoParcelasMedioAmb") = Guid.NewGuid
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDFinca")) = 0 Then ApplicationService.GenerateError("La Finca es un dato obligatorio.")
        Dim dt As DataTable = New BdgFinca().SelOnPrimaryKey(data("IDFinca"))
        If Length(dt) = 0 Then ApplicationService.GenerateError("La Finca no existe.")
        If Length(data("IDZonaEspecifica")) = 0 Then ApplicationService.GenerateError("La Zona Específica es un dato obligatorio.")
        If data("ZonaIncluidaEnFinca") Is DBNull.Value Then
            ApplicationService.GenerateError("La Zona Incluida en Finca es un dato obligatorio.")
        End If
        If data("FincaIncluidaEnZona") Is DBNull.Value Then
            ApplicationService.GenerateError("La Finca Incluida en Zona es un dato obligatorio.")
        End If
    End Sub

    <Task()> Public Shared Sub CambioCFinca(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim ClsBdg As New BdgCuadernoCampoEstatalParcelasMedioAmb
            Dim f As New StringFilterItem("CFinca", FilterOperator.Equal, data.Value)
            Dim dtFinca As DataTable = New BdgFinca().Filter(f)
            If Not dtFinca Is Nothing AndAlso dtFinca.Rows.Count > 0 Then
                data.Current("IDFinca") = dtFinca.Rows(0)("IDFinca")
                data.Current = ClsBdg.ApplyBusinessRule("IDFinca", data.Current("IDFinca"), data.Current)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioIDFinca(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim dtFinca As DataTable = New BdgFinca().SelOnPrimaryKey(data.Value)
            If Not dtFinca Is Nothing AndAlso dtFinca.Rows.Count > 0 Then
                Dim ClsBdg As New BdgCuadernoCampoEstatalParcelasMedioAmb
                data.Current("CFinca") = dtFinca.Rows(0)("CFinca")
                data.Current("IDVariedad") = dtFinca.Rows(0)("IDVariedad")
                data.Current = ClsBdg.ApplyBusinessRule("IDVariedad", data.Current("IDVariedad"), data.Current)

                If Length(data.Current("IDZonaEspecifica")) = 0 Then
                    Dim f As New StringFilterItem("IDFinca", FilterOperator.Equal, data.Value)
                    Dim dtZona As DataTable = New BdgFincaZonaEspecifica().Filter(f)
                    If Not dtZona Is Nothing AndAlso dtZona.Rows.Count = 1 Then
                        'La finca sólo tiene una zona
                        data.Current("IDZonaEspecifica") = dtZona.Rows(0)("IDZonaEspecifica")
                        data.Current = ClsBdg.ApplyBusinessRule("IDZonaEspecifica", data.Current("IDZonaEspecifica"), data.Current)
                    End If
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioVariedad(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        Dim ClsBdg As New BdgCuadernoCampoEstatalParcelasMedioAmb
        If Length(data.Value) > 0 Then
            Dim dt As DataTable = New BdgVariedad().SelOnPrimaryKey(data.Value)
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                data.Current("DescVariedad") = dt.Rows(0)("DescVariedad")
                data.Current("IDEspecie") = dt.Rows(0)("IDEspecie")
                data.Current = ClsBdg.ApplyBusinessRule("IDEspecie", data.Current("IDEspecie"), data.Current)
            End If
        Else
            data.Current("DescVariedad") = System.DBNull.Value
            data.Current("IDEspecie") = System.DBNull.Value
            data.Current = ClsBdg.ApplyBusinessRule("IDEspecie", data.Current("IDEspecie"), data.Current)
        End If
    End Sub

    <Task()> Public Shared Sub CambioEspecie(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim dt As DataTable = New BdgEspecie().SelOnPrimaryKey(data.Value)
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                data.Current("DescEspecie") = dt.Rows(0)("DescEspecie")
            End If
        Else
            data.Current("DescEspecie") = System.DBNull.Value
        End If
    End Sub

    <Task()> Public Shared Sub CambioNZonaEspecifica(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim f As New StringFilterItem("NZonaEspecifica", FilterOperator.Equal, data.Value)
            Dim dt As DataTable = New BdgZonaEspecifica().Filter(f)
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                data.Current("IDZonaEspecifica") = dt.Rows(0)("IDZonaEspecifica")
                Dim ClsBdg As New BdgCuadernoCampoEstatalParcelasMedioAmb
                data.Current = ClsBdg.ApplyBusinessRule("IDZonaEspecifica", data.Current("IDZonaEspecifica"), data.Current)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioIDZonaEspecifica(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim dtZona As DataTable = New BdgZonaEspecifica().SelOnPrimaryKey(data.Value)
            If Not dtZona Is Nothing AndAlso dtZona.Rows.Count > 0 Then
                data.Current("NZonaEspecifica") = dtZona.Rows(0)("NZonaEspecifica")
                data.Current("DescZonaEspecifica") = dtZona.Rows(0)("DescZonaEspecifica")
                data.Current("CoordenadasUTM") = dtZona.Rows(0)("CoordenadasUTM")

                Dim f As New Filter
                f.Add("IDZonaEspecifica", FilterOperator.Equal, data.Value)
                If Length(data.Current("IDFinca")) > 0 Then
                    f.Add("IDFinca", FilterOperator.Equal, data.Current("IDFinca"))
                End If
                Dim dtZonaFinca As DataTable = New BdgFincaZonaEspecifica().Filter(f)
                If Not dtZonaFinca Is Nothing AndAlso dtZonaFinca.Rows.Count = 1 Then
                    'La zona sólo tiene una finca
                    data.Current("IDZonaEspecifica") = dtZonaFinca.Rows(0)("IDZonaEspecifica")
                    data.Current("ZonaIncluidaEnFinca") = dtZonaFinca.Rows(0)("ZonaIncluidaEnFinca")
                    data.Current("FincaIncluidaEnZona") = dtZonaFinca.Rows(0)("FincaIncluidaEnZona")
                    data.Current("DistanciaMetros") = dtZonaFinca.Rows(0)("DistanciaMetros")
                    data.Current("SuperficieAfectadaHA") = dtZonaFinca.Rows(0)("SuperficieAfectadaHA")
                    If Length(data.Current("IDFinca")) = 0 Then
                        data.Current("IDFinca") = dtZonaFinca.Rows(0)("IDFinca")
                        Dim ClsBdg As New BdgCuadernoCampoEstatalParcelasMedioAmb
                        data.Current = ClsBdg.ApplyBusinessRule("IDFinca", data.Current("IDFinca"), data.Current)
                    End If
                End If
            End If
        End If
   End Sub
#End Region

End Class