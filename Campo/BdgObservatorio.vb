Public Class BdgObservatorio

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgObservatorio"

#End Region

#Region "Eventos Entidad"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClavePrimaria)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveExterna)
    End Sub

    <Task()> Public Shared Sub ValidarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDObservatorio")) = 0 Then
            ApplicationService.GenerateError("No se ha indicado un identificador para el Observatorio.")
        ElseIf Length(data("IDObservatorio")) > 25 Then
            ApplicationService.GenerateError("El identificador del Observatorio no puede ser mayor de 25 caracteres.")
        Else
            If data.RowState = DataRowState.Added AndAlso ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ExisteIDObservatorio, data("IDObservatorio"), services) Then
                ApplicationService.GenerateError("El identificador del Observatorio '|' ya se encuentra registrado en la Base de Datos.", data("IDObservatorio"))
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("DescObservatorio")) = 0 Then
            ApplicationService.GenerateError("No se ha indicado una descripción para el Observatorio.")
        ElseIf Length(data("DescObservatorio")) > 100 Then
            ApplicationService.GenerateError("La descripción para el Observatorio no puede superar los 100 caracteres.")
        End If
    End Sub

    <Task()> Public Shared Sub ValidarClaveExterna(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDTipoObservatorio")) > 0 Then
            If Not ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf BdgTipoObservatorio.ExisteIDTipoObservatorio, data("IDTipoObservatorio"), services) Then
                ApplicationService.GenerateError("No se ha encontrado el Tipo de Observatorio '|' en la Base de Datos.", data("IDTipoObservatorio"))
            End If
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Function ExisteIDObservatorio(ByVal data As String, ByVal services As ServiceProvider) As Boolean
        Dim dt As DataTable = New BdgObservatorio().SelOnPrimaryKey(data)
        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
            Return True
        Else : Return False
        End If
    End Function

#End Region

End Class