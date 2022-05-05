Public Class BdgDiagramaConfiguracionVariable

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgDiagramaConfiguracionVariable"

#End Region

#Region "Eventos entidad"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarRegistroDuplicado)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.IsNull("IDConfiguracionVariable") Then
            data("IDConfiguracionVariable") = Guid.NewGuid
        End If
    End Sub

#End Region

#Region "Task privadas"

    <Task()> _
    Public Shared Sub ValidarRegistroDuplicado(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim f As New Filter
        f.Add(New GuidFilterItem("IDConfiguracion", data("IDConfiguracion")))
        f.Add(New StringFilterItem("IDVariable", data("IDVariable")))
        f.Add(New StringFilterItem("IDConfiguracionVariable", FilterOperator.NotEqual, data("IDConfiguracionVariable")))
        Dim dttReg As DataTable = New BdgDiagramaConfiguracionVariable().Filter(f)
        If Not dttReg Is Nothing AndAlso dttReg.Rows.Count > 0 Then
            ApplicationService.GenerateError("No se pueden duplicar las variables para la misma configuración.")
        End If
    End Sub

#End Region

End Class
