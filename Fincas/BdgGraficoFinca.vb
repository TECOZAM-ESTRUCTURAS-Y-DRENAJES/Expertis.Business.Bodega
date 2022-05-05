Imports System.Data.OleDb

Public Class BdgGraficoFinca

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgGraficoFinca"

#End Region

#Region "Eventos Entidad"

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        data(_GF.IDGrafico) = Guid.NewGuid
    End Sub

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDescripcion)
    End Sub

    <Task()> Public Shared Sub ValidarDescripcion(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data(_GF.DescGrafico)) = 0 Then ApplicationService.GenerateError("La descripción no puede estar vacía")
    End Sub

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClavePrimaria)
    End Sub

    <Task()> Public Shared Sub AsignarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data(_GF.IDGrafico)) = 0 Then data(_GF.IDGrafico) = Guid.NewGuid
        End If
    End Sub

#End Region

End Class

<Serializable()> _
Public Class _GF
    Public Const IDGrafico As String = "IDGrafico"
    Public Const DescGrafico As String = "DescGrafico"
    Public Const GraficoXml As String = "GraficoXml"
End Class