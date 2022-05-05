Public Class BdgZonaEspecifica

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub


    Private Const cnEntidad As String = "tbBdgMaestroZonaEspecifica"

#End Region

#Region "Eventos Entidad"

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf AsignarContador)
        addnewProcess.AddTask(Of DataRow)(AddressOf AsignarClavePrimaria)
    End Sub

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
    End Sub

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDZonaEspecifica")) = 0 Then ApplicationService.GenerateError("Es necesario un identificador para la zona específica .")
    End Sub

    '<Task()> Public Shared Sub CargarContador(ByVal data As DataRow, ByVal services As ServiceProvider)
    '    Dim StDatos As New Contador.DatosDefaultCounterValue
    '    StDatos.row = data
    '    StDatos.EntityName = "BdgMaestroZonaEspecifica"
    '    StDatos.FieldName = "NZonaEspecifica"
    '    ProcessServer.ExecuteTask(Of Contador.DatosDefaultCounterValue)(AddressOf Contador.LoadDefaultCounterValue, StDatos, services)
    'End Sub

    <Task()> Public Shared Sub AsignarContador(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim StDatos As New Contador.DatosDefaultCounterValue
        StDatos.row = data
        StDatos.EntityName = "BdgMaestroZonaEspecifica"
        StDatos.FieldName = "NZonaEspecifica"
        ProcessServer.ExecuteTask(Of Contador.DatosDefaultCounterValue)(AddressOf Contador.LoadDefaultCounterValue, StDatos, services)
        'If data.RowState = DataRowState.Added Then
        '    If Length(data("IDContador")) > 0 Then
        '        data("NZonaEspecifica") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, data("IDContador"), services)
        '    Else
        '        Dim StDatos As New Contador.DatosDefaultCounterValue
        '        StDatos.row = data
        '        StDatos.EntityName = "BdgMaestroZonaEspecifica"
        '        StDatos.FieldName = "NZonaEspecifica"
        '        ProcessServer.ExecuteTask(Of Contador.DatosDefaultCounterValue)(AddressOf Contador.LoadDefaultCounterValue, StDatos, services)
        '        If Length(data("IDContador")) > 0 Then
        '            data("NZonaEspecifica") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, data("IDContador"), services)
        '        Else
        '            ApplicationService.GenerateError("No se ha configurado contador predeterminado para la Zona Específica.")
        '        End If

        '    End If
        'End If
    End Sub

    <Task()> Public Shared Sub AsignarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then data("IDZonaEspecifica") = Guid.NewGuid
    End Sub

#End Region

End Class