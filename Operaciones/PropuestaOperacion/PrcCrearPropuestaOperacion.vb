Imports Solmicro.Expertis.Engine.BE.BusinessProcesses
Public Class PrcCrearPropuestaOperacion
    Inherits Process

    Public Overrides Sub RegisterTasks()
        Me.AddTask(Of OperCab, DocumentoBdgOperacion)(AddressOf ProcesoBdgPropuestaOperacion.CrearDocumentoBdgOperacion)
        Me.AddTask(Of DocumentoBdgOperacion)(AddressOf ProcesoBdgPropuestaOperacion.PropuestaCabeceraOperacion)
        Me.AddTask(Of DocumentoBdgOperacion)(AddressOf ProcesoBdgPropuestaOperacion.PropuestaImputacionesGlobales)
        Me.AddTask(Of DocumentoBdgOperacion)(AddressOf ProcesoBdgPropuestaOperacion.PropuestaOrigenOperacion)
        Me.AddTask(Of DocumentoBdgOperacion)(AddressOf ProcesoBdgPropuestaOperacion.PropuestaDestinoOperacion)
        'Me.AddTask(Of DocumentoBdgOperacion)(AddressOf ProcesoBdgPropuestaOperacion.ValidarPropuesta)
        Me.AddTask(Of DocumentoBdgOperacion)(AddressOf ProcesoBdgPropuestaOperacion.GuardarPropuesta)
        Me.AddTask(Of DocumentoBdgOperacion)(AddressOf ProcesoBdgPropuestaOperacion.AñadirAResultado)
    End Sub


    Protected Overrides Function OnException(ByVal exceptionArgs As Engine.BE.BusinessProcesses.ProcessExceptionArgs) As Engine.BE.BusinessProcesses.OnExceptionBehaviour
        AdminData.RollBackTx()

        Dim log As LogProcess = exceptionArgs.Services.GetService(Of LogProcess)()
        ReDim Preserve log.Errors(log.Errors.Length)

        If TypeOf exceptionArgs.TaskData Is DocumentoBdgOperacion Then
            Dim oper As OperCab = CType(exceptionArgs.TaskData, DocumentoBdgOperacion).Cabecera
            Select Case oper.Origen
                Case OrigenOperacion.OperacionPlanificada
                    log.Errors(log.Errors.Length - 1) = New ClassErrors(CType(oper, OperCabPlanificadas).IDOrigen, exceptionArgs.Exception.Message)
                Case OrigenOperacion.Depositos
                    log.Errors(log.Errors.Length - 1) = New ClassErrors(Nothing, exceptionArgs.Exception.Message)
                Case OrigenOperacion.OrdenFabricacion
                    log.Errors(log.Errors.Length - 1) = New ClassErrors(CType(oper, OperCabOFs).NOrden, exceptionArgs.Exception.Message)
            End Select
        ElseIf TypeOf exceptionArgs.TaskData Is OperCab Then
            Dim oper As OperCab = CType(exceptionArgs.TaskData, OperCab)
            Select Case oper.Origen
                Case OrigenOperacion.OperacionPlanificada
                    log.Errors(log.Errors.Length - 1) = New ClassErrors(CType(oper, OperCabPlanificadas).IDOrigen, exceptionArgs.Exception.Message)
                Case OrigenOperacion.Depositos
                    log.Errors(log.Errors.Length - 1) = New ClassErrors(Nothing, exceptionArgs.Exception.Message)
                Case OrigenOperacion.OrdenFabricacion
                    log.Errors(log.Errors.Length - 1) = New ClassErrors(CType(oper, OperCabOFs).NOrden, exceptionArgs.Exception.Message)
            End Select
        Else
            log.Errors(log.Errors.Length - 1) = New ClassErrors(Nothing, exceptionArgs.Exception.Message)
        End If
        Return MyBase.OnException(exceptionArgs)
    End Function

 


End Class
