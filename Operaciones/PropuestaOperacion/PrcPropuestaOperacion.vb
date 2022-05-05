Imports Solmicro.Expertis.Engine.BE.BusinessProcesses

Public Class PrcPropuestaOperacion
    Inherits Process(Of DataPrcPropuestaOperacion, DataPrcPropuestaOperacionResultLog)

    Public Overrides Sub RegisterTasks()
        ' Me.AddTask(Of DataPrcPropuestaOperacion)(AddressOf Comunes.BeginTransaction)
        Me.AddTask(Of DataPrcPropuestaOperacion)(AddressOf ValidacionesPrevias)
        Me.AddTask(Of DataPrcPropuestaOperacion)(AddressOf DatosIniciales)
        Me.AddTask(Of DataPrcPropuestaOperacion, List(Of OperCab))(AddressOf ProcesoBdgPropuestaOperacion.AgruparOrigenes)
        Me.AddTask(Of List(Of OperCab))(AddressOf ValidarExistenDatosGenerarOperacion)
        Me.AddForEachTask(Of PrcCrearPropuestaOperacion)(OnExceptionBehaviour.NextLoop)
        Me.AddTask(Of Object, DataPrcPropuestaOperacionResultLog)(AddressOf GetResultadoPropuesta)
    End Sub

    <Task()> Public Shared Sub ValidacionesPrevias(ByVal data As DataPrcPropuestaOperacion, ByVal services As ServiceProvider)

    End Sub

    <Task()> Public Shared Sub DatosIniciales(ByVal data As DataPrcPropuestaOperacion, ByVal services As ServiceProvider)
        services.RegisterService(New ProcessInfoOperacion(data.AutoCalcularVino, data.GuardarPropuesta, data.Multiple))

        '//Prepara en el service lo que va a retornar el proceso, para tenerlo disponible en todo el proceso.
        services.RegisterService(New LogProcess)

        '//Preparamos el retorno de los resultados.
        services.RegisterService(New List(Of DataPrcPropuestaOperacionResult))

    End Sub

    <Task()> Public Shared Sub ValidarExistenDatosGenerarOperacion(ByVal data As List(Of OperCab), ByVal services As ServiceProvider)
        If data Is Nothing OrElse data.Count = 0 Then ApplicationService.GenerateError("No hay datos a para generar una Operación.")
    End Sub

    <Task()> Public Shared Function GetResultadoPropuesta(ByVal data As Object, ByVal services As ServiceProvider) As DataPrcPropuestaOperacionResultLog
        Dim lstPropuesta As List(Of DataPrcPropuestaOperacionResult) = services.GetService(Of List(Of DataPrcPropuestaOperacionResult))()
        Dim logPropuesta As LogProcess = services.GetService(Of LogProcess)()

        Dim rslt As New DataPrcPropuestaOperacionResultLog(lstPropuesta, logPropuesta)

        Return rslt
    End Function


    Protected Overrides Function OnException(ByVal exceptionArgs As Engine.BE.BusinessProcesses.ProcessExceptionArgs) As Engine.BE.BusinessProcesses.OnExceptionBehaviour
        AdminData.RollBackTx()

        Dim log As LogProcess = exceptionArgs.Services.GetService(Of LogProcess)()
        ReDim Preserve log.Errors(log.Errors.Length)
        log.Errors(log.Errors.Length - 1) = New ClassErrors(Nothing, exceptionArgs.Exception.Message)
        Return BusinessProcesses.OnExceptionBehaviour.NextTask
    End Function



   

End Class
