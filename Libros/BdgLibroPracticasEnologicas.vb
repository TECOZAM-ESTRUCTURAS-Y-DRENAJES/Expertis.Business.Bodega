Public Class BdgLibroPracticasEnologicas
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgLibroPracticasEnologicas"

#Region " RegisterDeleteTasks "

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf ComprobarEstado)
    End Sub

    <Task()> Friend Shared Sub ComprobarEstado(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data("Declarado") Then ApplicationService.GenerateError("Sólo se permite eliminar movimientos sin declarar.")
    End Sub

#End Region

#Region " RegisterValidateTasks "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("FechaMovimiento")) = 0 Then ApplicationService.GenerateError("La Fecha de Movimiento es obligatoria.")
    End Sub

#End Region

    <Serializable()> _
     Public Class DatosEntradasSalidas
        Public FechaDesde As Date
        Public FechaHasta As Date

        Public Sub New(ByVal FechaDesde As Date, ByVal FechaHasta As Date)
            Me.FechaDesde = FechaDesde
            Me.FechaHasta = FechaHasta
        End Sub
    End Class
    <Task()> Public Shared Sub Declarar(ByVal data As DatosEntradasSalidas, ByVal services As ServiceProvider)
        Dim f As New Filter
        f.Add(New DateFilterItem("FechaMovimiento", FilterOperator.GreaterThanOrEqual, data.FechaDesde))
        f.Add(New DateFilterItem("FechaMovimiento", FilterOperator.LessThanOrEqual, data.FechaHasta))
        f.Add(New BooleanFilterItem("Declarado", False))
        Dim dtLibro As DataTable = New BdgLibroPracticasEnologicas().Filter(f)
        If dtLibro.Rows.Count > 0 Then
            For Each drLibro As DataRow In dtLibro.Rows
                drLibro("Declarado") = True
            Next
            BdgLibroEmbotellados.UpdateTable(dtLibro)
        End If
    End Sub

End Class
