Public Class BdgCuadernoCampoEstatalAsesor

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    
    Private Const cnEntidad As String = "tbBdgCuadernoCampoEstatalAsesor"

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
        oBrl.Add("IDAsesor", AddressOf CambioIDAsesor)
        Return oBrl
    End Function

    <Task()> Public Shared Sub CambioIDAsesor(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDAsesor")) > 0 Then
            Dim Operarios As EntityInfoCache(Of OperarioInfo) = services.GetService(Of EntityInfoCache(Of OperarioInfo))()
            Dim op As OperarioInfo = Operarios.GetEntity(data.Current("IDAsesor"))
            data.Current("DescAsesor") = op.DescOperario
            data.Current("DNI") = op.DNI
            data.Current("NumeroIdentificacion") = op.NumeroIdentificacionAsesor
            data.Current("IDGestionPlagas") = op.IDGestionPlagas
        End If
    End Sub

#End Region


#Region "Funciones Públicas"
    <Task()> Public Shared Sub AsignarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then data("IDCuadernoAsesor") = Guid.NewGuid
    End Sub
#End Region

End Class