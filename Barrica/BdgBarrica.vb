Public Class BdgBarrica

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgBarrica"

#End Region

#Region "Eventos Entidad"

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf AsignarContador)
        addnewProcess.AddTask(Of DataRow)(AddressOf AsignarValoresPredeterminados)
    End Sub

    <Task()> Public Shared Sub AsignarContador(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim Dt As New Contador.DatosDefaultCounterValue(data, "BdgBarrica", "IDBarrica")
        ProcessServer.ExecuteTask(Of Contador.DatosDefaultCounterValue)(AddressOf Contador.LoadDefaultCounterValue, Dt, services)
    End Sub

    <Task()> Public Shared Sub AsignarValoresPredeterminados(ByVal data As DataRow, ByVal services As ServiceProvider)
        data("A�oCompra") = Date.Today.Year
        data("Capacidad") = 0
        data("Utilizadas") = 0
    End Sub

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDBarrica")) = 0 Then ApplicationService.GenerateError("El c�digo de barrica es obligatorio.")
        If Length(data("DescBarrica")) = 0 Then ApplicationService.GenerateError("La descripci�n es obligatoria.")
        If Length(data("Capacidad")) > 0 Then
            If data("Capacidad") < 0 Then ApplicationService.GenerateError("La capacidad no puede ser negativa.")
        End If
        If Length(data("A�oCompra")) = 0 OrElse Length(data("A�oCompra")) > 0 AndAlso Length(CType(data("A�oCompra"), String)) <> 4 Then
            ApplicationService.GenerateError("El a�o de compra no es v�lido.")
        End If
    End Sub

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClavePrimaria)
    End Sub

    <Task()> Public Shared Sub AsignarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDContador")) > 0 Then data("IDBarrica") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, data("IDContador"), services)
        End If
    End Sub

#End Region

End Class