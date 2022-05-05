Public Class BdgObservatorioFinca

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgObservatorioFinca"

#End Region

#Region "Eventos Entidad"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClave)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveExterna)
    End Sub

    <Task()> Public Shared Sub ValidarClave(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDObservatorio")) = 0 Then ApplicationService.GenerateError("No se ha indicado un Observatorio.")
        If Length(data("IDFinca")) = 0 Then ApplicationService.GenerateError("No se ha indicado una Finca.")
    End Sub

    <Task()> Public Shared Sub ValidarClaveExterna(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim dt As DataTable = New BdgObservatorio().SelOnPrimaryKey(data("IDObservatorio"))
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then
            ApplicationService.GenerateError("No se ha encontrado el Observatorio '|' en la Base de Datos.", data("IDObservatorio"))
        End If

        dt = New BdgFinca().SelOnPrimaryKey(data("IDFinca"))
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then
            ApplicationService.GenerateError("No se ha encontrado la Finca '|' en la Base de Datos.", data("IDFinca"))
        End If
    End Sub

#End Region

End Class