Public Class BdgObjetivosCampo
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgMaestroObjetivosCampo"

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarIDObjetivoCampo)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDescObjetivoCampo)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
    End Sub

    <Task()> Public Shared Sub ValidarIDObjetivoCampo(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDObjetivoCampo")) = 0 Then ApplicationService.GenerateError("El Objetivo Campo es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarDescObjetivoCampo(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("DescObjetivoCampo")) = 0 Then ApplicationService.GenerateError("La Descripción es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dt As DataTable = New BdgObjetivosCampo().SelOnPrimaryKey(data("IDObjetivoCampo"))
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("El Objetivo Campo ya existe en la base de datos.")
            End If
        End If
    End Sub

#End Region

End Class
