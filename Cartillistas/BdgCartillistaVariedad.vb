Public Class BdgCartillistaVariedad

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgCartillistaVariedad"

#End Region

#Region "Eventos Entidad"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDCartillista")) = 0 Then ApplicationService.GenerateError("El código de cartillista es obligatorio.")
        If Length(data("IDVariedad")) = 0 Then ApplicationService.GenerateError("La variedad es obligatoria.")
    End Sub

#End Region

End Class