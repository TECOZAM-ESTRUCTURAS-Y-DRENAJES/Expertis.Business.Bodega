Public Class BdgAnalisisVariable

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgAnalisisVariable"

#End Region

#Region "Eventos Entidad"

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf Comunes.DeleteEntityRow)
        deleteProcess.AddTask(Of DataRow)(AddressOf Comunes.MarcarComoEliminado)
        deleteProcess.AddTask(Of DataRow)(AddressOf ActualizarOrden)
    End Sub

    <Task()> Public Shared Sub ActualizarOrden(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim DtAnalisis As DataTable = New BdgAnalisisVariable().Filter(New StringFilterItem("IDAnalisis", data("IDAnalisis")), "Orden")
        If Not DtAnalisis Is Nothing AndAlso DtAnalisis.Rows.Count > 0 Then
            Dim i As Integer = 1
            For Each drAnaVar As DataRow In DtAnalisis.Select
                drAnaVar("Orden") = i
                i += 1
            Next
            BusinessHelper.UpdateTable(DtAnalisis)
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Function PlantillaAnalisisParametro(ByVal Data As Object, ByVal services As ServiceProvider) As DataTable
        Dim pid As String = New BdgParametro().AnalisisEntrada()
        If Length(pid) > 0 Then Return New BdgAnalisisVariable().Filter(New StringFilterItem("IDAnalisis", pid))
    End Function

    <Task()> Public Shared Function PlantillaAnalisisParametroEntVino(ByVal Data As Object, ByVal services As ServiceProvider) As DataTable
        Dim pid As String = New BdgParametro().AnalisisEntradaVino()
        If Length(pid) > 0 Then Return New BdgAnalisisVariable().Filter(New StringFilterItem("IDAnalisis", pid))
    End Function

    <Task()> Public Shared Function SelOnIDAnalisis(ByVal IDAnalisis As String, ByVal services As ServiceProvider) As DataTable
        Return New BdgAnalisisVariable().Filter(New StringFilterItem(_AV.IDAnalisis, FilterOperator.Equal, IDAnalisis))
    End Function

#End Region

End Class

<Serializable()> _
Public Class _AV
    Public Const IDAnalisis As String = "IDAnalisis"
    Public Const IDVariable As String = "IDVariable"
End Class