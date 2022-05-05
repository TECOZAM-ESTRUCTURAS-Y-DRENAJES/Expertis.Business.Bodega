Public Class BdgTipoObservatorio

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgTipoObservatorio"

#End Region

#Region "Eventos Entidad"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClavePrimaria)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDTipoObservatorio")) = 0 Then
            ApplicationService.GenerateError("No se ha indicado un Tipo de Observatorio.")
        ElseIf Length(data("IDTipoObservatorio")) > 2 Then
            ApplicationService.GenerateError("El Tipo de Observatorio no puede ser mayor de 2 caracteres.")
        Else
            If data.RowState = DataRowState.Added AndAlso ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ExisteIDTipoObservatorio, data("IDTipoObservatorio"), services) Then
                ApplicationService.GenerateError("El Tipo de Observatorio '|' ya se encuentra registrado en la Base de Datos.", data("IDTipoObservatorio"))
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDTipoObservatorio")) = 0 Then
            ApplicationService.GenerateError("No se ha indicado una descripción para el Tipo de Observatorio.")
        ElseIf Length(data("IDTipoObservatorio")) > 100 Then
            ApplicationService.GenerateError("La descripción para el Tipo de Observatorio no puede superar los 100 caracteres.")
        End If
    End Sub

#End Region


#Region "Funciones Públicas"

    <Task()> Public Shared Function ExisteIDTipoObservatorio(ByVal data As String, ByVal services As ServiceProvider) As Boolean
        Dim dt As DataTable = New BdgTipoObservatorio().SelOnPrimaryKey(data)
        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
            Return True
        Else : Return False
        End If
    End Function

#End Region


End Class