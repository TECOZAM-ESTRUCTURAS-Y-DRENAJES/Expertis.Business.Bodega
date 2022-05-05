Public Class BdgCuadernoCampoEstatalCentro

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgCuadernoCampoEstatalCentro"

#End Region

#Region "Eventos Entidad"

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
    End Sub

    
    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
    End Sub

    
    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClavePrimaria)
    End Sub

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
    End Sub
	

#End Region

#Region " GetBusinessRules "

    Public Overrides Function GetBusinessRules() As Solmicro.Expertis.Engine.BE.BusinessRules
        Dim oBrl As New BusinessRules
        oBrl.Add("IDCentro", AddressOf CambioIDCentro)
        Return oBrl
    End Function

    <Task()> Public Shared Sub CambioIDCentro(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDCentro")) > 0 Then
            Dim Centros As EntityInfoCache(Of CentroInfo) = services.GetService(Of EntityInfoCache(Of CentroInfo))()
            Dim c As CentroInfo = Centros.GetEntity(data.Current("IDCentro"))
            data.Current("DescCentro") = c.DescCentro
            data.Current("NumeroIncripcionROMA") = c.NumeroIncripcionROMA
            If c.FechaAdquisicion <> cnMinDate Then data.Current("FechaAdquisicion") = c.FechaAdquisicion
            If c.FechaUltimaInspeccion <> cnMinDate Then data.Current("FechaUltimaInspeccion") = c.FechaUltimaInspeccion
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Sub AsignarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then data("IDCuadernoCentro") = Guid.NewGuid
    End Sub

#End Region

End Class