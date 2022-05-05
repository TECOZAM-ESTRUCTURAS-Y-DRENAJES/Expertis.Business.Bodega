Public Class BdgFormatoEnvase

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

  
    Private Const cnEntidad As String = "tbBdgMaestroFormatoEnvase"

#End Region

#Region "Eventos Entidad"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub
	
#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDFormatoEnvase")) = 0 Then ApplicationService.GenerateError("El Formato Envase es un dato obligatorio.")
    End Sub

#End Region

End Class