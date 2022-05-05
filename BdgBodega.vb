Public Class BdgBodega

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgBodega"

#End Region

#Region "Eventos Entidad"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClavePrimaria)
    End Sub

    <Task()> Public Shared Sub AsignarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then data("IDBodega") = Guid.NewGuid
    End Sub

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Sub ValidatePrimaryKey(ByVal IdBodega As String, ByVal Services As ServiceProvider)
        Dim DtAux As DataTable = New BdgBodega().SelOnPrimaryKey(IdBodega)
        If DtAux.Rows.Count = 0 Then ApplicationService.GenerateError("Actualización en conflicto con el valor de la clave | de la tabla |", IdBodega, cnEntidad)
    End Sub

#End Region

End Class

<Serializable()> _
Public Class _B
    Public Const IdBodega As String = "IdBodega"
    Public Const DescBodega As String = "DescBodega"
End Class