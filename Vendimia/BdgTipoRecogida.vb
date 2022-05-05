Public Class BdgTipoRecogida

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgMaestroTipoRecogida"

#End Region

#Region "Update tasks"

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        'data(_TR.IDTipoRecogida) = AdminData.GetAutoNumeric
        data(_TR.IncrementoTipoRecogida) = 0
    End Sub

#End Region

#Region "Validate task"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data(_TR.IDTipoRecogida)) = 0 Then ApplicationService.GenerateError("El valor asignado al IDTipoRecogida no es válido.")
        If Length(data(_TR.DescTipoRecogida)) = 0 Then ApplicationService.GenerateError("El valor asignado al Tipo de Recogida no es válido.")
        If Length(data(_TR.IncrementoTipoRecogida)) = 0 Then ApplicationService.GenerateError("El valor asignado al Incremento no es válido.")
    End Sub

#End Region

End Class
<Serializable()> _
Public Class _TR
    Public Const IDTipoRecogida As String = "IDTipoRecogida"
    Public Const DescTipoRecogida As String = "DescTipoRecogida"
    Public Const IncrementoTipoRecogida As String = "IncrementoTipoRecogida"
    Public Const FechaCreacionAudi As String = "FechaCreacionAudi"
    Public Const FechaModificacionAudi As String = "FechaModificacionAudi"
    Public Const UsuarioAudi As String = "UsuarioAudi"
End Class