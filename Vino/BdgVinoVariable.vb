Public Class BdgVinoVariable

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgVinoVariable"

#End Region

#Region "Eventos Entidad"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarVariable)
    End Sub

    <Task()> Public Shared Sub ValidarVariable(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Not Length(data("IDVariable")) > 0 Then ApplicationService.GenerateError("El código de variable es obligatorio.")
    End Sub

#End Region

End Class

<Serializable()> _
Public Class _VV
    Public Const IDVinoAnalisis As String = "IDVinoAnalisis"
    Public Const IDVariable As String = "IDVariable"
    Public Const Valor As String = "Valor"
    Public Const ValorNumerico As String = "ValorNumerico"
    Public Const Orden As String = "Orden"
End Class