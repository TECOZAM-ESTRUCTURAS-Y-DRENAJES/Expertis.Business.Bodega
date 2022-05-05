Public Class BdgOperacionVinoPlanVarios

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgOperacionVinoPlanVarios"

#End Region

#Region "Tareas GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim OBrl As New BusinessRules
        OBrl.Add("IDVarios", AddressOf BdgGeneral.CambioVarios)
        Return OBrl
    End Function

#End Region

#Region " RegisterValidateTasks "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDVarios")) = 0 Then ApplicationService.GenerateError("El Varios es un dato obligatorio.")
    End Sub

#End Region

End Class