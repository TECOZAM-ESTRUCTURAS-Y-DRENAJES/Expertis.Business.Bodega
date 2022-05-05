Public Class BdgCuadernoCampoEstatalParcelasAgro

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgCuadernoCampoEstatalParcelasAgro"

#End Region

#Region "Eventos Entidad"
    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
    End Sub

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClavePrimaria)
    End Sub
    
    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
    End Sub

    Public Overrides Function GetBusinessRules() As Solmicro.Expertis.Engine.BE.BusinessRules
        Dim oBrl As New BusinessRules
        oBrl.Add("IDFinca", AddressOf CambioIDFinca)
        oBrl.Add("CFinca", AddressOf CambioCFinca)
        oBrl.Add("IDFincaSigpac", AddressOf CambioIDFincaSigpac)
        oBrl.Add("CFincaSigpac", AddressOf CambioCFincaSigpac)
        oBrl.Add("IDVariedad", AddressOf CambioVariedad)
        Return oBrl
    End Function
#End Region

#Region "Funciones Públicas"
    <Task()> Public Shared Sub AsignarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then data("IDCuadernoParcelasAgro") = Guid.NewGuid
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDFinca")) = 0 Then ApplicationService.GenerateError("La Finca de Trabajo es un dato obligatorio.")
        Dim dt As DataTable = New BdgFinca().SelOnPrimaryKey(data("IDFinca"))
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then ApplicationService.GenerateError("La Finca de Trabajo no existe.")
    End Sub

    <Task()> Public Shared Sub CambioIDFinca(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDFinca")) > 0 Then
            Dim ClsBdg As New BdgCuadernoCampoEstatalParcelasAgro
            Dim dttFinca As DataTable = New BdgFinca().SelOnPrimaryKey(data.Current("IDFinca"))
            If (Not dttFinca Is Nothing AndAlso dttFinca.Rows.Count > 0) Then
                With dttFinca
                    data.Current("IDVariedad") = .Rows(0)("IDVariedad")
                    data.Current = ClsBdg.ApplyBusinessRule("IDVariedad", data.Current("IDVariedad"), data.Current)
                    data.Current("IDRiego") = .Rows(0)("IDRiego")
                End With
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioCFinca(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim ClsBdg As New BdgCuadernoCampoEstatalParcelasAgro
            Dim filtro As New StringFilterItem("CFinca", FilterOperator.Equal, data.Value)
            Dim dttFinca As DataTable = New BdgFinca().Filter(filtro)
            If (Not dttFinca Is Nothing AndAlso dttFinca.Rows.Count > 0) Then
                data.Current("IDFinca") = dttFinca.Rows(0)("IDFinca")
                data.Current = ClsBdg.ApplyBusinessRule("IDFinca", data.Current("IDFinca"), data.Current)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioIDFincaSigpac(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDFincaSigpac")) > 0 Then
            Dim ClsBdg As New BdgCuadernoCampoEstatalParcelasAgro
            Dim dttFinca As DataTable = New BdgFinca().SelOnPrimaryKey(data.Current("IDFincaSigpac"))
            If (Not dttFinca Is Nothing AndAlso dttFinca.Rows.Count > 0) Then
                With dttFinca
                    data.Current("IDProvinciaSigpac") = .Rows(0)("IDProvinciaSigpac")
                    data.Current("IDMunicipioSigpac") = .Rows(0)("IDMunicipioSigpac")
                    data.Current("AgregadoSigpac") = .Rows(0)("AgregadoSigpac")
                    data.Current("ZonaSigPac") = .Rows(0)("ZonaSigPac")
                    data.Current("Poligono") = .Rows(0)("Poligono")
                    data.Current("Parcela") = .Rows(0)("Parcela")
                    data.Current("RecintoSigpac") = .Rows(0)("RecintoSigpac")
                    data.Current("IDUsoSigpac") = .Rows(0)("IDUsoSigpac")
                    data.Current("IDAireLibre") = .Rows(0)("IDAireLibre")
                    data.Current("IDGestionPlagas") = .Rows(0)("IDGestionPlagas")
                    data.Current("Superficie") = .Rows(0)("Superficie")
                    data.Current("SuperficieViñedo") = .Rows(0)("SuperficieViñedo")
                End With
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioCFincaSigpac(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim ClsBdg As New BdgCuadernoCampoEstatalParcelasAgro
            Dim filtro As New StringFilterItem("CFinca", FilterOperator.Equal, data.Value)
            Dim dttFinca As DataTable = New BdgFinca().Filter(filtro)
            If (Not dttFinca Is Nothing AndAlso dttFinca.Rows.Count > 0) Then
                data.Current("IDFincaSigpac") = dttFinca.Rows(0)("IDFinca")
                data.Current = ClsBdg.ApplyBusinessRule("IDFincaSigpac", data.Current("IDFincaSigpac"), data.Current)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioVariedad(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim ClsBdg As New BdgCuadernoCampoEstatalParcelasAgro
            Dim filtro As New StringFilterItem("IDVariedad", FilterOperator.Equal, data.Value)
            Dim dttVariedad As DataTable = New BdgVariedad().Filter(filtro)
            If (Not dttVariedad Is Nothing AndAlso dttVariedad.Rows.Count > 0) Then
                data.Current("IDEspecie") = dttVariedad.Rows(0)("IDEspecie")
                data.Current = ClsBdg.ApplyBusinessRule("IDEspecie", data.Current("IDEspecie"), data.Current)
            End If
        End If
    End Sub
#End Region

End Class