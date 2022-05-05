Public Class BdgFincaInversion

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgFincaInversion"

#End Region

#Region "Eventos"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClavePrimaria)
    End Sub

    <Task()> Public Shared Sub AsignarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then data("IDFincaInversion") = Guid.NewGuid
    End Sub

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("FechaDesde")) = 0 Then ApplicationService.GenerateError("La Fecha Desde es obligatoria.")
        If Length(data("FechaHasta")) = 0 Then ApplicationService.GenerateError("La Fecha Hasta es obligatoria.")
        If (data("FechaHasta") < data("FechaDesde")) Then ApplicationService.GenerateError("La Fecha Hasta no puede ser inferior a la Fecha Desde.")
    End Sub

#End Region

End Class
