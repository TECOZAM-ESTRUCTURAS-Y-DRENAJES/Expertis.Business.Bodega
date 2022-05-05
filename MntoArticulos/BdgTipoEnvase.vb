Public Class BdgTipoEnvase

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgMaestroTipoEnvase"

#End Region

#Region "Eventos ValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDTipoEnvase")) = 0 Then ApplicationService.GenerateError("El Tipo de envase es un dato obligatorio.")
        If Length(data("DescTipoEnvase")) = 0 Then ApplicationService.GenerateError("La Descripción del tipo de envase es un dato obligatorio.")
    End Sub

#End Region

End Class