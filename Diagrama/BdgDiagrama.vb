Public Class BdgDiagrama

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgDiagrama"

#End Region

#Region "Eventos Entidad"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDDiagrama")) = 0 Then ApplicationService.GenerateError("El identificador de diagrama no es válido.")
        If Length(data("DescDiagrama")) = 0 Then ApplicationService.GenerateError("La descripción es obligatoria.")
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dt As DataTable = New BdgDiagrama().SelOnPrimaryKey(data("IDDiagrama"))
            If dt.Rows.Count > 0 Then ApplicationService.GenerateError("El registro introducido ya existe en la base de datos")
        End If
    End Sub

#End Region

End Class